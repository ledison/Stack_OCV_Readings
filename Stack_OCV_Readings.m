function varargout = Stack_OCV_Readings(varargin)
% STACK_OCV_READINGS MATLAB code for Stack_OCV_Readings.fig
%      STACK_OCV_READINGS, by itself, creates a new STACK_OCV_READINGS or raises the existing
%      singleton*.
%
%      H = STACK_OCV_READINGS returns the handle to a new STACK_OCV_READINGS or the handle to
%      the existing singleton*.
%
%      STACK_OCV_READINGS('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in STACK_OCV_READINGS.M with the given input arguments.
%
%      STACK_OCV_READINGS('Property','Value',...) creates a new STACK_OCV_READINGS or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Stack_OCV_Readings_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Stack_OCV_Readings_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Stack_OCV_Readings

% Last Modified by GUIDE v2.5 15-Aug-2016 14:54:01

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Stack_OCV_Readings_OpeningFcn, ...
                   'gui_OutputFcn',  @Stack_OCV_Readings_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Stack_OCV_Readings is made visible.
function Stack_OCV_Readings_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Stack_OCV_Readings (see VARARGIN)

% Choose default command line output for Stack_OCV_Readings
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

assignin('base','GUI_Handles',handles);
assignin('base','COM_Port_Open',0);

% UIWAIT makes Stack_OCV_Readings wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Stack_OCV_Readings_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

Instruction1 = '1. Connect the DMM Remote Control Cable to the DMM and the computer';
Instruction2 = '2. Turn on the DMM and set it to VDC and the range to +/-50';
Instruction3 = '3. Remove the dongle from the presenter device and connect it to the computer';
Instruction4 = '4. Connect the barcode scanner to the computer';
uiwait(msgbox({Instruction1 Instruction2 Instruction3 Instruction4},...
              'Test Setup Instructions'));

%Get the first stack serial number
uicontrol(handles.Current_Serial_Number);


% --- Executes on button press in Save_Readings.
function Save_Readings_Callback(hObject, eventdata, handles)
% hObject    handle to Save_Readings (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
Save_Readings(handles);

% --- Executes on button press in Close_Program.
function Close_Program_Callback(hObject, eventdata, handles)
% hObject    handle to Close_Program (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
dmm = evalin('base','dmm_hdl');
fclose(dmm);
evalin('base','clear all');
closereq;

% Handle inputs from the presenter device
function getKeyPress(~, event)
handles = evalin('base','GUI_Handles');
Battery = evalin('base','NextBatt');
if strcmp(event.Key,'f5') % get a reading from the DMM and display it
    LastF5Time = evalin('base','LastF5Time');
    CurrentF5Time = clock;
    ElapsedTime = etime(CurrentF5Time,LastF5Time);
    if ElapsedTime < 1
        return;
    end
    assignin('base','LastF5Time',CurrentF5Time);
    % get the DMM handle
    dmm = evalin('base','dmm_hdl');
    % send the command
    fprintf(dmm,'QM');
    % get and discard the ack
    ack = fscanf(dmm);
    % get the response
    Response = fscanf(dmm);
    % find the commas in the response
    CommaLoc = strfind(Response,',');
    % get the string version of the value read
    ValueStr = Response(1:CommaLoc(1)-1);
    % round it to the nearest hundreths
    NewValue = (round(100*str2num(ValueStr)))/100;
    % display the value
    eval(['set(handles.B' num2str(Battery) '_Voltage,''String'',' num2str(NewValue,'%4.2f') ')']);
    % check to see if the value is valid
    if NewValue <= 2 || NewValue >= 6.5
        % if the value is invalid, try to make the computer beep
        beep on;
        beep;
        beep;
    end
    % retrieve the array of readings
    value = evalin('base','Readings');
    % add the new value to the array
    value(Battery) = NewValue;
    % store the updated array in base
    assignin('base','Readings',value);
    % if this is not Battery 8, move to the next battery
    if Battery<8
        BatteryStates = evalin('base','BatteryState');
        if BatteryStates(Battery) == 1
            % set the background back to white
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 1.0])']);
        else
            AnalyzeReadings(handles,value,Battery);
        end
        Battery = Battery+1;
        assignin('base','NextBatt',Battery);
        % set the background for the next reading to light blue
        eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[.678 .922 1.0])']);
        InstructStr = sprintf('Probe Battery # %d',Battery);
        set(handles.Instructions,'String',InstructStr);
    else % if it is Battery 8 always analyze the results
        AnalyzeReadings(handles,value,Battery);
    end
elseif strcmp(event.Key,'pageup')
    if Battery<8
        BatteryStates = evalin('base','BatteryState');
        if BatteryStates(Battery) == 0
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 0.0])']);
        else
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 1.0])']);
        end    
        Battery = Battery+1;
        assignin('base','NextBatt',Battery);
        if BatteryStates(Battery) == 0
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[0.6 0.6 0.0])']);
        else
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[.678 .922 1.0])']);
        end
        InstructStr = sprintf('Probe Battery # %d',Battery);
        set(handles.Instructions,'String',InstructStr);
    end
elseif strcmp(event.Key,'pagedown')
    if Battery>1
        BatteryStates = evalin('base','BatteryState');
        if BatteryStates(Battery) == 0
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 0.0])']);
        else
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 1.0])']);
        end    
        Battery = Battery-1;
        assignin('base','NextBatt',Battery);
        if BatteryStates(Battery) == 0
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[0.6 0.6 0.0])']);
        else
            eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[.678 .922 1.0])']);
        end
        InstructStr = sprintf('Probe Battery # %d',Battery);
        set(handles.Instructions,'String',InstructStr);
    end
elseif strcmp(event.Key,'period')
    Save_Readings(handles);
end

function AnalyzeReadings(handles,value,CurrentBattery)
MaxReading = max(value);
MinReading = min(value);
Battery = 1;
BatteryState = evalin('base','BatteryState');
for Battery = 1:8
    if MaxReading - value(Battery) > 0.05
        % if the delta is high change the the cell background color
        eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 0.0])']);
        BatteryState(Battery) = 0;
        assignin('base','BatteryState',BatteryState);
    else
        eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 1.0])']);
        BatteryState(Battery) = 1;
        assignin('base','BatteryState',BatteryState);
	end
end

function Close_Program(~, event)
dmm = evalin('base','dmm_hdl');
fclose(dmm);

% backup the data to the T drive
TDrivePath = 'T:\End of Line Testing Data\Stack Summary Data\';
CDrivePath = evalin('base','LocalDrivePath');
XLSFileName = evalin('base','DataFileName');
copyfile([CDrivePath XLSFileName],[TDrivePath XLSFileName]);

evalin('base','clear all');
closereq;


% --- Executes on button press in ProductionRadioButton.
function ProductionRadioButton_Callback(hObject, eventdata, handles)
% hObject    handle to ProductionRadioButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%State = get(hObject,'Value');
%if State ~= 1
    %set(hObject,'Value',1.0);
    set(handles.LabRadioButton,'Value',0.0);
    assignin('base','TestType','Production SOQA');
    uicontrol(handles.B1_Voltage);
%end

% --- Executes on button press in LabRadioButton.
function LabRadioButton_Callback(hObject, eventdata, handles)
% hObject    handle to LabRadioButton (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%State = get(hObject,'Value');
%if State ~= 1
    %set(hObject,'Value',1.0);
    set(handles.ProductionRadioButton,'Value',0.0);
    assignin('base','TestType','Lab SOQA');
    uicontrol(handles.B1_Voltage);
%end



function Current_Serial_Number_Callback(hObject, eventdata, handles)
% hObject    handle to Current_Serial_Number (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
str = get(hObject,'String');
set(handles.Current_Serial_Number,'String',str);
assignin('base','StackSN',str);
set(handles.Instructions,'String','Probe Battery # 1');

warning off;
if evalin('base','COM_Port_Open') == 0
    
    % Set the COM port address for the DMM
    COM_Port_Str = FindDMMPort;
    %COM_Port_Str = 'COM4';
    %COM_Port = FirstLine(COM_Port_Str);
    
    % Initialize the DMM COM port
    dmm = serial(COM_Port_Str,'BaudRate',115200,'Terminator','CR');
    assignin('base','dmm_hdl',dmm);
    fopen(dmm);
    assignin('base','COM_Port_Open',1);
end
warning on;
% change the background color to green and set the Production Radio Button
% to on and make the Close button visible
set(handles.figure1,'Color',[0.2 0.6 0]);
set(handles.ProductionRadioButton,'BackgroundColor',[0.2 0.6 0]);
set(handles.LabRadioButton,'BackgroundColor',[0.2 0.6 0]);

% check the states of the radio buttons, set to default of Production ON if
% both are 0
ProdRadButtonState = get(handles.ProductionRadioButton,'Value');
LabRadButtonState = get(handles.LabRadioButton,'Value');
if ProdRadButtonState == 0 && LabRadButtonState == 0
    set(handles.ProductionRadioButton,'Value',1.0);
    set(handles.LabRadioButton,'Value',0.0);
    assignin('base','TestType','Production SOQA');
    set(handles.Close_Program,'Visible','on');
end

% Initialize the counter and values variables
NextBatt = 1;
eval(['set(handles.B' num2str(NextBatt) '_Voltage,''BackgroundColor'',[.678 .922 1.0])']);
assignin('base','NextBatt',NextBatt);
values(1:8) = 0;
assignin('base','Readings',values);
BattState(1:8) = 1;
assignin('base','BatteryState',BattState);
LastF5Time = clock;
assignin('base','LastF5Time',LastF5Time);

% Assign the CloseRequestFcn property to Close_Program
set(handles.figure1,'CloseRequestFcn',@Close_Program);

% Assign Key Press processing to the getKeyPress function
set(handles.B1_Voltage,'KeyPressFcn',@getKeyPress);

uicontrol(handles.B1_Voltage);

function Save_Readings(handles)
% check to make sure there is actually something to be saved
SNStr = get(handles.Current_Serial_Number,'String');
SNEmpty = strcmp(SNStr,'');
if SNEmpty == 1
    uicontrol(handles.Current_Serial_Number);
    return;
end

% Put the data into an array
OCVData{1,1} = evalin('base','StackSN');
OCVData{1,2} = datestr(now,'mm/dd/yy');
Values = evalin('base','Readings');
for Battery = 1:8
    OCVData{1,Battery+2} = Values(Battery);
end
OCVData{1,11} = max(Values)-min(Values);
OCVData{1,12} = evalin('base','TestType');

% Save the data for the stack in its own .txt file for importing into PLEX
% Make sure the directory exists
Plex_Data_Path = 'C:\OCV_Data\Plex_Data\';
if ~exist(Plex_Data_Path,'dir') 
    mkdir(Plex_Data_Path)
end
% put the Plex data into a separate array
Plex_Data_Filename = [OCVData{1,1} '.txt'];
Plex_Data{1} = OCVData{1,1};
Plex_Data{2} = OCVData{1,2};
for i = 3:11
    Plex_Data{i} = num2str(OCVData{1,i});
end
%Plex_Data{11} = num2str(OCVData{1,11});
if OCVData{1,11} > 0.2
    Plex_Data{12} = 'FAIL';
else
    Plex_Data{12} = 'PASS';
end
fid = fopen([Plex_Data_Path Plex_Data_Filename],'w');
for i = 1:12
    fprintf(fid,'%s\t',Plex_Data{i});
end
fclose(fid);
%dlmwrite([Plex_Data_Path Plex_Data_Filename], Plex_Data,'delimiter','\t');

% Initialize file info variables
DrivePath = 'C:\OCV_Data\';
assignin('base','LocalDrivePath',DrivePath);
XLSFileName = 'Stack OCV Data.xlsx';
assignin('base','DataFileName',XLSFileName);
XLSFileName = [DrivePath XLSFileName];
    
% Get xls summary file info
FileExist = exist(XLSFileName,'file');
if FileExist > 0
    [type, Sheet] = xlsfinfo(XLSFileName);
    XLSExists = 1;
else
    XLSExists = 0;
end
    
% open excel server
Excel = actxserver ('Excel.Application');
assignin('base','Excel',Excel);
if ~exist(XLSFileName,'file') 
    ExcelWorkbook = Excel.workbooks.Add; 
    ExcelWorkbook.SaveAs(XLSFileName,1); 
    ExcelWorkbook.Close(false); 
end 
invoke(Excel.Workbooks,'Open',XLSFileName);

if XLSExists == 0 % New sheet - should only happen once
    tdata = {'Serial Number','Date','Battery # 1 OCV','Battery # 2 OCV',...
             'Battery # 3 OCV','Battery # 4 OCV','Battery # 5 OCV',...
             'Battery # 6 OCV','Battery # 7 OCV','Battery # 8 OCV',...
             'OCV Delta','Test Type','Comments',};
    xlswrite1(XLSFileName, tdata ,'Sheet1','A1');
    lastline_summary  = 0;
    nextline  = 2;
else % Sheet already exists
    % Get Previous data
    Previous_data = xlsread1(XLSFileName,'Sheet1');
    % Find the last line of data
    %if isempty(discharge_info) == 0
        nextline  = size(Previous_data,1)+2;
    %end
end

% write the new data to the next row of the spreadsheet
xlswrite1(XLSFileName, OCVData, 'Sheet1',['A' num2str(nextline)]);
% close excel server
invoke(Excel.ActiveWorkbook,'Save'); 
Excel.Quit; 
Excel.delete;
clear Excel;

% reset the variables and clear the fields
value = evalin('base','Readings');
value(1:8) = 0;
assignin('base','Readings',value);
for Battery = 1:8
    eval(['set(handles.B' num2str(Battery) '_Voltage,''String'','''')']);
    eval(['set(handles.B' num2str(Battery) '_Voltage,''BackgroundColor'',[1.0 1.0 1.0])']);
end
NextBatt = 1;
assignin('base','NextBatt',NextBatt);

SerNum = evalin('base','StackSN');
PrevSerNum = SerNum;
SerNum = '';
assignin('base','StackSN',SerNum);
set(handles.Instructions,'String','Scan Stack Serial Number');
set(handles.Current_Serial_Number,'String',SerNum);
set(handles.Previous_Serial_Number,'String',PrevSerNum);

% Get the next serial number
uicontrol(handles.Current_Serial_Number);


% --- Executes during object creation, after setting all properties.
function Previous_Serial_Number_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Previous_Serial_Number (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function Current_Serial_Number_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Current_Serial_Number (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
