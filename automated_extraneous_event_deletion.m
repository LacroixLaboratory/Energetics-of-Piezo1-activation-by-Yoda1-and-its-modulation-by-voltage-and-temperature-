% This script accepts Excel files containing ClampFit event extraction
% data. Each tab of the Excel file corresponds to a single recording. 
[status,sheets] = xlsfinfo('MasterFile-Yoda1-Reanalysis');
finalsheetaverages = zeros(100,length(sheets));
finalsheetnumbers = zeros(100,length(sheets));
% The ClampFit-assigned event levels are typically stored in column 3,
% event dwell times in column 9, and the true amplitude of each event in
% column 7. Column numbers can be adjusted in case the input data does not
% fit this pattern. 
for primary = 1:length(sheets);
    try
    F = readtable('MasterFile-Yoda1-Reanalysis','Sheet',primary);
        event_amplitudes = table2array(F(:,3));
        event_durations = table2array(F(:,9));
        event_amptrue = table2array(F(:,7));
        event_array = [event_amplitudes event_durations event_amptrue];
% If a closing event has an amplitude above 75% of the preceding opening
% event, assume that the 'closure' is due to thermal noise rather than a
% true return to zero, and retain an open-channel state. 
for hm2 = length(event_array):-1:2;
    if event_array(hm2,1) == 0 & event_array(hm2,3) > 0.75*event_array(hm2-1,3); 
        event_array(hm2,1) = event_array(hm2-1,1); 
    end
end
% If an opening event has an amplitude below 125% of the preceding closing
% event, assume that the 'opening' is due to thermal noise rather than a
% true return to zero, and retain a closed-channel state. 
for hm2 = length(event_array):-1:2;
    if event_array(hm2,1) == 1 & event_array(hm2,3) < 1.25*event_array(hm2-1,3); 
        event_array(hm2,1) = event_array(hm2-1,1); 
    end
end
% If an event lasts less than 0.05 ms, combine it to the preceding event as
% dead noise. 
for hm2 = length(event_array):-1:1;
    if event_array(hm2,2)<0.05; 
        event_array(hm2,1)=event_array(hm2-1,1);
    end
end
% Delete all open-to-open and close-to-close transitions, merging the
% multi-open and multi-close states created by the above conditions. 
for hm = length(event_array):-1:2;
    if event_array(hm,1) == event_array(hm-1,1);
        event_array(hm-1,2) = event_array(hm-1,2) + event_array(hm,2);  
        event_array(hm,:) = [];
    end;
end;
event_array2 = event_array
    end
% Save the resulting documents as text file, with each file corresponding
% to a tab. 
output = ['tab_' num2str(primary) '.txt']
csvwrite(output,event_array)
clearvars -except primary status sheets
fprintf("Completed sheet %d",primary)
end