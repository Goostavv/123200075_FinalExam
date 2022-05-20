function varargout = FinalExam_123200075(varargin)
% FINALEXAM_123200075 MATLAB code for FinalExam_123200075.fig
%      FINALEXAM_123200075, by itself, creates a new FINALEXAM_123200075 or raises the existing
%      singleton*.
%
%      H = FINALEXAM_123200075 returns the handle to a new FINALEXAM_123200075 or the handle to
%      the existing singleton*.
%
%      FINALEXAM_123200075('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in FINALEXAM_123200075.M with the given input arguments.
%
%      FINALEXAM_123200075('Property','Value',...) creates a new FINALEXAM_123200075 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before FinalExam_123200075_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to FinalExam_123200075_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help FinalExam_123200075

% Last Modified by GUIDE v2.5 20-May-2022 09:03:56

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @FinalExam_123200075_OpeningFcn, ...
                   'gui_OutputFcn',  @FinalExam_123200075_OutputFcn, ...
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


% --- Executes just before FinalExam_123200075 is made visible.
function FinalExam_123200075_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to FinalExam_123200075 (see VARARGIN)

% Choose default command line output for FinalExam_123200075
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes FinalExam_123200075 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = FinalExam_123200075_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in calculate_075.
function calculate_075_Callback(hObject, eventdata, handles)
% hObject    handle to calculate_075 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
showdata = readcell('laptop_price.csv', 'Range', 'A2:M51'); 
header = readcell('laptop_price.csv', 'Range', 'A1:M1');
set(handles.uitable1_075, 'Data', showdata, 'ColumnName',header);

opts = detectImportOptions('laptop_price.csv');
opts.SelectedVariableNames = {'Inches','Ram','Weight','Price_euros'};
criterion= readtable('laptop_price.csv', opts);

k=[1,1,1,0]; 
w=[1,3,2,4]; 


[m n]=size (showdata); 
w=w./sum(w); 

for j=1:n,
    if k(j)==0, w(j)=-1*w(j);
    end;
end;
for i=1:m,
    S(i)=prod(criterion(i,:).^w);
end;

opts = detectImportOptions('laptop_price.csv');
opts.SelectedVariableNames = (1);
new = readmatrix('laptop_price.csv', opts);
xlswrite('Result_WP.xlsx', new, 'Sheet1', 'A1'); 
S=S'; 
xlswrite('Result_WP.xlsx', S, 'Sheet1', 'B1'); 

opts = detectImportOptions('Result_WP.xlsx');
opts.SelectedVariableNames = (1:2);
data = readmatrix('Result_WP.xlsx', opts); 

X=sortrows(data,2,'descend'); 
set(handles.tabel2,'data',X,'visible','on'); 


% --- Executes during object creation, after setting all properties.
function result_075_CreateFcn(hObject, eventdata, handles)
% hObject    handle to result_075 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
