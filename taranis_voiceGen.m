%% Taranis OpenTX SAPI TTS Voice Generator
% Author: L.Nyhof
% Description: This script uses the Win32 version of SAPI so you will need
% to run a 32bit version of Matlab to get this to work, sorry, I have no
% control over this.
% This Script takes a list of Spoken text, directory structure and file
% names from a Matlab array, the TTS voices are then spoken, stored and
% written as .WAV files with the following parameters:
%  Fs = 32000
%  Bit Depth = 16 bit
%  Format = PCM uLaw
%  
% 
% Examples: 
%       tts('Welcome to OpenTX.');      % Speak the text;
%       tts('','List');                 % List availble voices;
%       w = tts('Welcome to OpenTX.',      [],      0,  32000);%       % Do not speak out, store the speech in a variable;
% handle^ func^    text to be read^    voice^  speed^     Fs^
%       wavplay(w,44100);


%%
clear all;close all;
load('TaranisVoiceData2.mat');
Fs        = 32000;
N         = 228;
R         = 2;
Voice     = 3;
FileName  = 2;
Directory = 1;
Speed     = 0;   %range goes from -5 to 5
Nbits     = 16;

% This section extracts the data from the Excel file and converts it into
% text to be converted to speech and saved using SAPI5 TTS voices, then the
% wav file is saved into a directory structure suitable for the FrSky
% Taranis OpenTX firmware.

for R = 2:N
    % the first one is a string
    if R == 1
        VoiceList{R-1,Voice} = cellstr(textdata(R,Voice));  % use this for cells that contain 'strings'
        VoiceList{R-1,FileName} = cellstr(textdata(R,FileName));  % use this for cells that contain 'strings'
        VoiceList{R-1,Directory} = cellstr(textdata(R,Directory));  % use this for cells that contain 'strings'
    else
    % the next 2 to 110 are numbers
        if R >= 2
            if R <= 110
                VoiceList{R-1,Voice} = num2str(textdata{R,Voice});    % use this for cells that contain numbers
                VoiceList{R-1,FileName} = num2str(textdata{R,FileName});    % use this for cells that contain numbers
                VoiceList{R-1,Directory} = num2str(textdata{R,Directory});    % use this for cells that contain numbers
            elseif R > 110
                VoiceList{R-1,Voice} = num2str(textdata{R,Voice});    % use this for cells that contain numbers
                VoiceList{R-1,FileName} = num2str(textdata{R,FileName});    % use this for cells that contain numbers
                VoiceList{R-1,Directory} = num2str(textdata{R,Directory});    % use this for cells that contain numbers      
            end
        end 
    end
end

%% Now to convert the text into speech and save
Y=1;
for Y=1:N
    if Y==227
        break
    end
     VoiceTemp = VoiceList(Y,3);
     w = tts(char(VoiceTemp),'VW Kate',Speed,Fs);

    currentPath = pwd;     % Get the current working directory
    [s,~,~] = mkdir(strcat(currentPath,'\SOUNDS'),'EN');     % make a new directory to store the files in ../EN & ../EN/SYSTEM
    [s,~,~] = mkdir(strcat(currentPath,'\SOUNDS\EN'),'SYSTEM');     % make a new directory to store the files in ../EN & ../EN/SYSTEM

    currentPath  = strcat(currentPath,'\',VoiceList{Y,Directory},'\',VoiceList{Y,FileName}); % get the full file path
    SaveToHere   = currentPath;             % store full file path in temp variable
    % Save the file
    wavwrite(w,Fs,Nbits,currentPath);
    Y                                       %counter to monitor progress
 end








































