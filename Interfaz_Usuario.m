function varargout = Interfaz_Usuario(varargin)
% INTERFAZ_USUARIO MATLAB code for Interfaz_Usuario.fig
%      INTERFAZ_USUARIO, by itself, creates a new INTERFAZ_USUARIO or raises the existing
%      singleton*.
%
%      H = INTERFAZ_USUARIO returns the handle to a new INTERFAZ_USUARIO or the handle to
%      the existing singleton*.
%
%      INTERFAZ_USUARIO('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in INTERFAZ_USUARIO.M with the given input arguments.
%
%      INTERFAZ_USUARIO('Property','Value',...) creates a new INTERFAZ_USUARIO or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Interfaz_Usuario_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Interfaz_Usuario_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Interfaz_Usuario

% Last Modified by GUIDE v2.5 01-Feb-2021 13:18:04

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Interfaz_Usuario_OpeningFcn, ...
                   'gui_OutputFcn',  @Interfaz_Usuario_OutputFcn, ...
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

%*************************************************************************%
%**********************CONEXION MATLAB CON HYSYS**************************%
%*************************************************************************%
function hyapp=hyconnect(Ruta_FileNameString,VisibleBoolean)
        hy=feval('actxserver','Hysys.application');
        
        if nargin<=1
              hy.Visible=1;
        else
              hy.Visible=VisibleBoolean;
        end
           try
              if nargin>0
                 
hyy=invoke(hy.SimulationCases,'Open',Ruta_FileNameString);
                 invoke(hyy,'Activate');
               end
           catch
                disp('No se encuentra el archivo');
           end
             hyapp=hy;
             
             
% --- Executes just before Interfaz_Usuario is made visible.
function Interfaz_Usuario_OpeningFcn(hObject, eventdata, handles, varargin)
clc; clear global;
global booleano opc1 opc2 actpre actmol     
movegui('center');
opc1 = 1;
opc2 = 0;
booleano = true;
set(handles.Panel1,'visible','on')
set(handles.Panel2,'visible','off')
set(handles.Panel10,'visible','off')

axes(handles.Separator);
[x1,map1]=imread('SeparatorGrey.jpg');
image(x1);
colormap(map1);
axis off
hold on

axes(handles.Config);
[x10,map10]=imread('SinReduccion.jpg');
image(x10);
colormap(map10);
axis off

g = imread('Run1.jpg');
set(handles.Go_To_Plant,'CData',g);

gg = imread('Reset1.jpg');
set(handles.clean,'CData',gg);

filaname = {'Methane','Ethane','Propane','i-Butane','n-Butane',...
            'i-Pentane','n-Pentane','n-Hexane','n-Heptane','n-Octane',...
            'n-Nonane','n-Decane','n-C11','CO2','Nitrogen','H2S','H2O'};
columnaname = {'Mole Fraction'};
L = num2cell(95*ones(1,17));
set(handles.uitable6,'Data',cell(17,1),'ColumnEditable',true,...
    'RowName',filaname,'ColumnName',columnaname,'ColumnWidth',L);

set(handles.activepreent,'value',0);
set(handles.activemolarent,'value',0);
set(handles.edit10,'String',' ');
set(handles.edit10,'Enable','on');
set(handles.edit12,'String',' ');
set(handles.edit12,'Enable','on');
set(handles.edit11,'String',' ');
set(handles.presal,'value',0);
set(handles.molarsal,'value',0);
set(handles.edit13,'String','');
set(handles.edit13,'visible','off');
set(handles.edit16,'String','');
set(handles.edit16,'visible','off');
set(handles.DiametroSeparador,'String',' ');
set(handles.LongitudSeparador,'String',' ');
set(handles.compl1,'visible','off');
set(handles.nocompl1,'visible','on');

set(handles.sinreduccion,'Value',1);
set(handles.panelsr,'visible','on');
set(handles.serie,'String',' ');
set(handles.serie,'Enable','on');
set(handles.confirmarserie,'Enable','on');
set(handles.text2,'Visible','off')
set(handles.text3,'Visible','off')
set(handles.etapa,'Visible','off')
set(handles.confirmaretapa,'Visible','off')
set(handles.confirmaretapa,'Enable','on'); 
set(handles.text3,'String','0');
set(handles.etapa,'String',' ');
set(handles.etapa,'Enable','on'); 
set(handles.text101,'Visible','off');
set(handles.paralelo,'String',' ');
set(handles.paralelo,'Enable','on')
set(handles.paralelo,'Visible','off');
set(handles.confirmarparalelo,'Visible','off');
set(handles.confirmarparalelo,'Enable','on');
set(handles.conreduccion,'Value',0);
set(handles.panelampred,'visible','off');
set(handles.iniparatrains,'String','');
set(handles.iniparatrains,'Enable','on');
set(handles.confirmarinitrain,'Enable','on');
set(handles.text114,'Visible','off');
set(handles.confirmarfintrain,'Visible','off');
set(handles.confirmarfintrain,'Enable','on');
set(handles.finparatrains,'Visible','off');
set(handles.finparatrains,'Enable','on');
set(handles.finparatrains,'String','');
set(handles.text120,'Visible','off');
set(handles.etapainiparatrain,'String','');
set(handles.etapainiparatrain,'visible','off');
set(handles.etapainiparatrain,'Enable','on');
set(handles.confirmaretapini,'Visible','off');
set(handles.confirmaretapini,'Enable','on');
set(handles.text214,'Visible','off');
set(handles.etapafinparatrain,'String','');
set(handles.etapafinparatrain,'Visible','off');
set(handles.etapafinparatrain,'Enable','on');
set(handles.confirmaretapfin,'Visible','off');
set(handles.confirmaretapfin,'Enable','on');
set(handles.uipanel26,'visible','on');
set(handles.Panel31,'visible','off');
set(handles.text33,'string',' ');
set(handles.text34,'string',' ');
set(handles.text35,'string',' ');
set(handles.text43,'string',' ');
set(handles.uipanel40,'visible','off');
set(handles.uipanel26,'title','Your choice');
set(handles.text134,'String','');
set(handles.text136,'String','');
set(handles.text143,'String','');
set(handles.text157,'String','');
set(handles.text175,'String','');
set(handles.text145,'String','');
set(handles.aborts,'visible','on');
set(handles.compl2,'visible','off');
set(handles.nocompl2,'visible','on');

set(handles.file,'String','');
set(handles.resucondin,'visible','off');
set(handles.text184,'String','');
set(handles.text186,'String','');
set(handles.text200,'String','');
set(handles.resucondout,'visible','off');
set(handles.text190,'String','');
set(handles.text202,'String','');
set(handles.resusepar,'visible','off')
set(handles.text203,'string','')
set(handles.text206,'string','')
set(handles.resutrasta,'visible','off');
set(handles.resusinreduc,'visible','off');
set(handles.text59,'String','');
set(handles.text61,'String','');
set(handles.text62,'String','');
set(handles.text152,'String','');
set(handles.resuconreduc,'visible','off');
set(handles.text155,'String','');
set(handles.text157,'String','');
set(handles.text173,'String','');
set(handles.text175,'String','');
set(handles.compl3,'visible','off');
set(handles.nocompl3,'visible','on');
% set(handles.Go_To_Plant,'Enable','off');
% actmol = 0;
% actpre = 0;
guidata(hObject,handles);

% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Interfaz_Usuario (see VARARGIN)

% Choose default command line output for Interfaz_Usuario
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Interfaz_Usuario wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Interfaz_Usuario_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;
handles.output = hObject;
guidata(hObject, handles);


% --- Executes on button press in Tab1.
function Tab1_Callback(hObject, eventdata, handles)
set(handles.Panel1,'visible','on')
set(handles.Panel2,'visible','off')
set(handles.Panel10,'visible','off')
guidata(hObject, handles);

% --- Executes on button press in Tab2.
function Tab2_Callback(hObject, eventdata, handles)
set(handles.Panel2,'visible','on')
set(handles.Panel1,'visible','off')
set(handles.Panel10,'visible','off')
guidata(hObject, handles);

% --- Executes on button press in Tab10.
function Tab10_Callback(hObject, eventdata, handles)
set(handles.Panel10,'visible','on')
set(handles.Panel1,'visible','off')
set(handles.Panel2,'visible','off')
% handles.Panel10 = Panel10;
% handles.Panel1 = Panel1;
% handles.Panel2 = Panel2;
guidata(hObject, handles);

% --- Executes on button press in activepreent.
function activepreent_Callback(hObject, eventdata, handles)
global R11 actpre actmol T1 T2
R11 = get(hObject,'Value');
if R11 == 1
set(handles.activemolarent,'Value',0);
set(handles.edit12,'Enable','off');
set(handles.edit10,'Enable','on');
actpre = 1;
actmol = 0;
% T1 = 1;
% T2 = 0;
else
set(handles.activemolarent,'Value',1);
set(handles.activemolarent,'visible','on');
set(handles.activepreent,'Value',0);
set(handles.edit12,'Enable','on');
set(handles.edit10,'Enable','off');
actpre = 0;
actmol = 1;
% T1 = 0;
% T2 = 1;
end
guidata(hObject,handles);

% --- Executes on button press in activemolarent.
function activemolarent_Callback(hObject, eventdata, handles)
global R22 actpre actmol T1 T2
R22 = get(hObject,'Value');
if R22 == 1
set(handles.activepreent,'Value',0);
set(handles.edit10,'Enable','off');
set(handles.edit12,'Enable','on');
actmol = 1;
actpre = 0;
% T2 = 1;
% T1 = 0;
else
set(handles.activepreent,'Value',1);
set(handles.activepreent,'visible','on');
set(handles.activemolarent,'Value',0);
set(handles.edit10,'Enable','on');
set(handles.edit12,'Enable','off');
actmol = 0;
actpre = 1;
% T2 = 0;
% T1 = 1;
end
guidata(hObject,handles);

% --- Executes on button press in presal.
function presal_Callback(hObject, eventdata, handles)
global R1
R1 = get(hObject,'Value');
if R1 == 1
set(handles.edit13,'visible','on');
set(handles.edit13,'string','');
set(handles.edit16,'visible','off');
set(handles.molarsal,'Value',0);
set(handles.edit16,'String',0);
else
set(handles.edit13,'visible','off')
set(handles.edit13,'string',0);
set(handles.edit16,'visible','on')
set(handles.molarsal,'Value',1);
set(handles.edit16,'String','');
end
guidata(hObject,handles);

% --- Executes on button press in molarsal.
function molarsal_Callback(hObject, eventdata, handles)
global R2
R2 = get(hObject,'Value');
if R2 == 1
set(handles.edit16,'visible','on')
set(handles.edit16,'String','')
set(handles.edit13,'visible','off')
set(handles.presal,'Value',0);
set(handles.edit13,'String',0);
else
set(handles.edit16,'visible','off')
set(handles.edit16,'String',0)
set(handles.edit13,'visible','on')
set(handles.presal,'Value',1);
set(handles.edit13,'String','');
end
guidata(hObject,handles);

 % --- Executes on button press in confirmarcondi.
function confirmarcondi_Callback(hObject, eventdata, handles)
global comp1 comp2 comp3 comp4 comp5 comp6 comp7 comp8 comp9 comp10...
    comp11 comp12 comp13 comp14 comp15 comp16 comp17 presionent tempent...
    molfloent presionsal molflosal fraccion diametro longitud R11 R22 go1...
    go2 go3 actmol actpre R1 R2

%*************************************************************************%
%**************************PRESION DE ENTRADA*****************************%
%*************************************************************************%
presionent = str2double(get(handles.edit10,'string'));
if isnan(presionent)
    errordlg('The value entered must be numeric','ERROR','modal')
    set(handles.edit10,'string','0')
    return;
end
if isempty(presionent) || presionent <= 0
    errordlg('The pressure inlet is incorrect','ERROR','modal')
    set(handles.edit10,'string','0')
    return;
end
handles.presionent = str2double(get(hObject,'string'));
guidata(hObject,handles);

%*************************************************************************%
%************************TEMPERATURA DE ENTRADA***************************%
%*************************************************************************%
tempent = str2double(get(handles.edit11,'string'));
if isnan(tempent)
    errordlg('The value entered must be numeric','ERROR','modal')
    set(handles.edit11,'string','0')
    return;
end
if isempty(tempent)
    errordlg('The initial temperature is incorrect','ERROR','modal')
    set(handles.edit11,'string','0')
    return;
end
handles.tempent = str2double(get(hObject,'string'));
guidata(hObject,handles);


%*************************************************************************%
%**************************FLUJO DE ENTRADA*******************************%
%*************************************************************************%
molfloent = str2double(get(handles.edit12,'string'));
if isnan(molfloent)
   errordlg('The value entered must be numeric','ERROR','modal')
   set(handles.edit12,'string','0')
   return;
end
if isempty(molfloent) || molfloent <=0
    errordlg('The molar flow inlet is incorrect','ERROR','modal')
   set(handles.edit12,'string','0')
   return;
end
handles.molfloent = str2double(get(hObject,'string'));
guidata(hObject,handles);

%*************************************************************************%
%****************************DIÁMETRO VESSEL******************************%
%*************************************************************************%
diametro = str2double(get(handles.DiametroSeparador,'string'));
longitud = str2double(get(handles.LongitudSeparador,'string'));

if isnan(diametro)
errordlg('The diameter must be numeric','ERROR','modal')
set(handles.DiametroSeparador,'string','0')
return;
end
if isempty(diametro) || diametro <= 0
errordlg('The diameter is incorrect','ERROR','modal')
set(handles.DiametroSeparador,'string','0')
return;
end
if isnan(longitud)
errordlg('The length must be numeric','ERROR','modal')
set(handles.LongitudSeparador,'string','0')
return;
end
if isempty(longitud) || longitud <= 0
errordlg('The length is incorrect','ERROR','modal')
set(handles.LongitudSeparador,'string','0')
return;
end

handles.molflosal = str2double(get(hObject,'string'));
guidata(hObject,handles);

set(handles.text199,'visible','on')
set(handles.text200,'visible','on')
set(handles.text183,'visible','on')
set(handles.text184,'visible','on')
set(handles.text184,'string',molfloent)
set(handles.text200,'string',presionent)
set(handles.text186,'string',tempent)

if presionsal == 0
    set(handles.text201,'visible','off')
    set(handles.text202,'visible','off')
    set(handles.text189,'visible','on')
    set(handles.text190,'visible','on')
    set(handles.text190,'string',molflosal)
else
    set(handles.text201,'visible','on')
    set(handles.text202,'visible','on')
    set(handles.text189,'visible','off')
    set(handles.text190,'visible','off')
    set(handles.text202,'string',presionsal)
end

set(handles.resusepar,'visible','on')
set(handles.text203,'string',diametro)
set(handles.text206,'string',longitud)

%*************************************************************************%
%***************************FRACCION MOLAR********************************%
%*************************************************************************%
% fraccion = get(handles.uitable6,'Data');
fraccion = cell2mat(get(handles.uitable6,'Data'));
if isempty(fraccion)
    errordlg('The molar fractions are empty','ERROR','modal')
    return;
end
comp1 = fraccion(1,1);
% if isempty(fraccion(1,1))
%     errordlg('The molar fraction of the Methane is empty','ERROR','modal')
%     return;
% else
%     comp1 = fraccion(1,1);
% end
comp2 = fraccion(2,1);
% if isempty(fraccion(1,2))
%     errordlg('The molar fraction of the Ethane is empty','ERROR','modal')
%     return;
% else
%     comp2 = fraccion(2,1);
% end
comp3 = fraccion(3,1);
comp4 = fraccion(4,1);
comp5 = fraccion(5,1);
comp6 = fraccion(6,1);
comp7 = fraccion(7,1);
comp8 = fraccion(8,1);
comp9 = fraccion(9,1);
comp10 = fraccion(10,1);
comp11 = fraccion(11,1);
comp12 = fraccion(12,1);
comp13 = fraccion(13,1);
comp14 = fraccion(14,1);
comp15 = fraccion(15,1);
comp16 = fraccion(16,1);
comp17 = fraccion(17,1);

if isnan(comp1)
    errordlg('The molar fraction of the Methane must be numeric','ERROR','modal')
    return;
end
if isnan(comp2)
    errordlg('The molar fraction of the Ethane must be numeric','ERROR','modal')
    return;
end
if isnan(comp3)
    errordlg('The molar fraction of the Propane must be numeric','ERROR','modal')
    return;
end
if isnan(comp4)
    errordlg('The molar fraction of the i-Butane must be numeric','ERROR','modal')
    return;
end
if isnan(comp5)
    errordlg('The molar fraction of the n-Butane must be numeric','ERROR','modal')
    return;
end
if isnan(comp6)
    errordlg('The molar fraction of the i-Pentane must be numeric','ERROR','modal')
    return;
end
if isnan(comp7)
    errordlg('The molar fraction of the n-Pentane must be numeric','ERROR','modal')
    return;
end
if isnan(comp8)
    errordlg('The molar fraction of the n-Hexane must be numeric','ERROR','modal')
    return;
end
if isnan(comp9)
    errordlg('The molar fraction of the n-Heptane must be numeric','ERROR','modal')
    return;
end
if isnan(comp10)
    errordlg('The molar fraction of the n-Octane must be numeric','ERROR','modal')
    return;
end
if isnan(comp11)
    errordlg('The molar fraction of the n-Nonane must be numeric','ERROR','modal')
    return;
end
if isnan(comp12)
    errordlg('The molar fraction of the n-Decane must be numeric','ERROR','modal')
    return;
end
if isnan(comp13)
    errordlg('The molar fraction of the n-C11 must be numeric','ERROR','modal')
    return;
end
if isnan(comp14)
    errordlg('The molar fraction of the CO2 must be numeric','ERROR','modal')
    return;
end
if isnan(comp15)
    errordlg('The molar fraction of the Nitrogen must be numeric','ERROR','modal')
    return;
end
if isnan(comp16)
    errordlg('The molar fraction of the H2S must be numeric','ERROR','modal')
    return;
end
if isnan(comp17)
    errordlg('The molar fraction of the H2O must be numeric','ERROR','modal')
    return;
end


if isempty(R1) && isempty(R2)
    errordlg('You should to indicate the outlet','ERROR','modal')
    return;
end
%*************************************************************************%
%**************************PRESION DE SALIDA******************************%
%*************************************************************************%
presionsal = str2double(get(handles.edit13,'string'));
if isnan(presionsal)
    errordlg('The value entered must be numeric','ERROR','modal')
    set(handles.edit13,'string','0')
    return;
end
if isempty(presionsal) %|| presionsal <= 0
    errordlg('The pressure outlet is incorrect','ERROR','modal')
    set(handles.edit13,'string','0')
    return;
end
handles.presionsal = str2double(get(hObject,'string'));

%*************************************************************************%
%****************************FLUJO DE SALIDA******************************%
%*************************************************************************%
molflosal = str2double(get(handles.edit16,'string'));
if isnan(molflosal)
    errordlg('The value entered must be numeric','ERROR','modal')
    set(handles.edit16,'string','0')
    return;
end
if isempty(molflosal) %|| molflosal <= 0
    errordlg('The molar flow outlet is incorrect','ERROR','modal')
    set(handles.edit16,'string','0')
    return;
end

%*************************************************************************%
%********************ESPECIFICACION DINAMICA ENTRADA**********************%
%*************************************************************************%
if isempty(actpre) && isempty(actmol)
    errordlg('You should to active one of those dynamics specifications','ERROR','modal')
    return;
end

% go1 = 1;
% if go1 == 1 
%     if go2 ==1 || go3 == 1
%    set(handles.Go_To_Plant,'Enable','on'); 
%     end
% end
warndlg({'The conditions have been completed'},'NOTIFICATION','modal');
set(handles.resucondin,'visible','on');
set(handles.resucondout,'visible','on');
set(handles.compl1,'visible','on');
set(handles.nocompl1,'visible','off');
guidata(hObject,handles); 


% --- Executes on button press in sinreduccion.
function sinreduccion_Callback(hObject, eventdata, handles)
global R10 opc1 opc2
R10 = get(hObject,'Value');
if R10 == 1
    opc1 = 1;
    opc2 = 0;
%     set(handles.nocompl2,'visible','on');
%     set(handles.compl2,'visible','off');
    set(handles.panelsr,'visible','on');
    set(handles.conreduccion,'Value',0);
    set(handles.panelampred,'visible','off');
    axes(handles.Config);
    [x10,map10]=imread('SinReduccion.jpg');
    image(x10);
    colormap(map10);
    axis off
else
    opc1 = 0;
    opc2 = 1;
    set(handles.panelsr,'visible','off');
    set(handles.conreduccion,'Value',1);
    set(handles.panelampred,'visible','on');
    axes(handles.Config);
    [x100,map100]=imread('ConReduccion.jpg');
    image(x100);
    colormap(map100);
    axis off  
end
guidata(hObject, handles);


% --- Executes on button press in conreduccion.
function conreduccion_Callback(hObject, eventdata, handles)
global R100 opc1 opc2
R100 = get(hObject,'Value');
if R100 == 1
    opc2 = 1;
    opc1 = 0;
%     set(handles.nocompl2,'visible','on');
%     set(handles.compl2,'visible','off');
    set(handles.panelampred,'visible','on');
    set(handles.sinreduccion,'Value',0);
    set(handles.panelsr,'visible','off');
    axes(handles.Config);
    [x100,map100]=imread('ConReduccion.jpg');
    image(x100);
    colormap(map100);
    axis off
else
    opc2 = 0;
    opc1 = 1;
    set(handles.panelampred,'visible','off');
    set(handles.sinreduccion,'Value',1);
    set(handles.panelsr,'visible','on');
    axes(handles.Config);
    [x10,map10]=imread('SinReduccion.jpg');
    image(x10);
    colormap(map10);
    axis off
end
guidata(hObject, handles);


% --- Executes on button press in confirmarserie.
function confirmarserie_Callback(hObject, eventdata, handles)
global etapaserie se sumaes anterior booleano
clear inipara finpara etafin etaini 
se = str2double(get(handles.serie,'string'));

if isnan(se)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.serie,'string','0')
set(handles.text3,'string','0')
return;
end

if se > 2
errordlg('The number of series trains must be less than 2','ERROR','modal')
set(handles.serie,'string','0')
return;
end

if isempty(se) || se <= 0
errordlg('The minimum value of series trains must be 1','ERROR','modal')
set(handles.serie,'string','0')
return;
end

handles.se = str2double(get(hObject,'string'));
set(handles.serie,'Enable','off')
set(handles.confirmarserie,'Enable','off');
guidata(hObject,handles);

set(handles.text3,'String',1)
set(handles.text33,'String',se)
set(handles.text59,'String',se)
set(handles.text2,'Visible','on')
set(handles.text3,'Visible','on')
set(handles.etapa,'Visible','on')
set(handles.confirmaretapa,'Visible','on')
booleano = true;
etapaserie = 1:handles.se;
sumaes = 0;
anterior = 4; %%Numero maximo de etapas
handles.se = se;
set(handles.Panel31,'visible','off');
set(handles.uipanel40,'visible','off');
set(handles.resuconreduc,'visible','off');
set(handles.resusinreduc,'visible','off');
set(handles.resutrasta,'visible','off');
set(handles.uipanel26,'title','Your choice');
guidata(hObject,handles);


% --- Executes on button press in confirmaretapa.
function confirmaretapa_Callback(hObject, eventdata, handles)
global etapaserie se et a sumaes anterior booleano
    
a = str2double(get(handles.text3,'String'));
et = str2double(get(handles.etapa,'string'));

if isnan(et)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.etapa,'string','0')
return;
end

if isempty(et) || et <= 0
errordlg('The minimum number for stages is 1','ERROR','modal')
set(handles.etapa,'string','0')
return;
end

if et > 4
errordlg({'The maximum number for stages is 4'},'ERROR','modal');
set(handles.etapa,'string','0')
return;
end

if et > anterior || et>(5-a)
errordlg({'The number should be lower than the last one you wrote'},'ERROR','modal');
set(handles.etapa,'string','0')
return;
end

if a <= se
etapaserie(a) = et; 
sumaes = sumaes+et;
end

if  a < se
set(handles.text3,'String',a+1);
else
set(handles.etapa,'Enable','off');
set(handles.confirmaretapa,'Enable','off'); 
set(handles.text101,'Visible','on');
set(handles.paralelo,'Visible','on');
set(handles.confirmarparalelo,'Visible','on');
end

set(handles.etapa,'String','');

if a == 1
set(handles.text34,'string',etapaserie(1));
set(handles.text61,'string',etapaserie(1));
end

if a == 2
set(handles.text34,'string',etapaserie(1));
set(handles.text35,'string',etapaserie(2));
set(handles.text61,'string',etapaserie(1));
set(handles.text62,'string',etapaserie(2));
end

booleano = true;
anterior = et;
handles.et = et;
handles.a = a;
guidata(hObject, handles);


% --- Executes on button press in confirmarparalelo.
function confirmarparalelo_Callback(hObject, eventdata, handles)
global para etapasetren1 etapasetren2 booleano etafin etaini

para = str2double(get(handles.paralelo,'string'));

if isnan(para)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.paralelo,'String','0')
return;
end

if isempty(para) || para < 0
errordlg('The minimum number for parallel trains is 2','ERROR','modal')
set(handles.paralelo,'string','0')
return;
end

if para > 5
errordlg('The maximum number for parallel trains is 5','ERROR','modal')
set(handles.paralelo,'String','0')
return;
else
set(handles.paralelo,'Enable','off')
set(handles.confirmarparalelo,'Enable','off');
set(handles.text43,'string',para);
set(handles.Panel31,'visible','on');
set(handles.uipanel26,'visible','on');
set(handles.uipanel26,'title','No change in the number of parallel trains');
set(handles.aborts,'visible','on');
warndlg({'The configuration is correct'},'NOTIFICATION','modal');
end
% go2 = 1;
% if go1 == 1 && go2 ==1
%    set(handles.Go_To_Plant,'Enable','on'); 
% end
% handles.para = str2double(get(hObject,'string'));

etapasetren1 = str2double(get(handles.text34,'string'));
etapasetren2 = str2double(get(handles.text35,'string'));
booleano = true;

set(handles.uipanel40,'visible','off');
set(handles.resuconreduc,'visible','off');
set(handles.text152,'string',para);
set(handles.resutrasta,'visible','on');
set(handles.resusinreduc,'visible','on');
set(handles.compl2,'visible','on');
set(handles.nocompl2,'visible','off');
clear inipara finpara etafin etaini
guidata(hObject,handles);


% --- Executes on button press in confirmarinitrain.
function confirmarinitrain_Callback(hObject, eventdata, handles)
global inipara booleano
clear se para etapaserie et
inipara = str2double(get(handles.iniparatrains,'string'));

if isnan(inipara)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.iniparatrains,'string','0')
return;
end

if inipara > 5
errordlg('The maximum number for parallel trains initials is 5','ERROR','modal')
set(handles.iniparatrains,'string','0')
return;
end

if isempty(inipara) || inipara <= 0
errordlg('The minimum number for parallel trains initials is 1','ERROR','modal')
set(handles.iniparatrains,'string','0')
return;
end
booleano = true;
set(handles.text134,'String',inipara)
set(handles.text155,'String',inipara)
set(handles.iniparatrains,'Enable','off')
set(handles.confirmarinitrain,'Enable','off')
set(handles.text120,'Visible','on');
set(handles.text120,'Visible','on');
set(handles.etapainiparatrain,'Visible','on');
set(handles.confirmaretapini,'Visible','on');

set(handles.Panel31,'visible','off');
set(handles.uipanel40,'visible','off');
set(handles.resuconreduc,'visible','off');
set(handles.resusinreduc,'visible','off');
set(handles.resutrasta,'visible','off');
set(handles.uipanel26,'title','Your choice');

handles.inipara = inipara;
guidata(hObject,handles);


% --- Executes on button press in confirmaretapini.
function confirmaretapini_Callback(hObject, eventdata, handles)
global etaini

% a1 = str2double(get(handles.contini,'String'));
etaini = str2double(get(handles.etapainiparatrain,'string'));

if isnan(etaini)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.etapainiparatrain,'string','0')
return;
end

if etaini > 4 
errordlg({'The maximum number for stages is 4'},'ERROR','modal');
set(handles.etapainiparatrain,'string','0')
return;
end

if isempty(etaini) || etaini <= 0 
errordlg('The minimum value for stages is 1','ERROR','modal')
set(handles.etapainiparatrain,'string','0')
return;
end


set(handles.etapainiparatrain,'Enable','off');
set(handles.confirmaretapini,'Enable','off');
set(handles.text114,'visible','on');
set(handles.finparatrains,'visible','on');
set(handles.confirmarfintrain,'visible','on');
set(handles.confirmarfintrain,'enable','on');
set(handles.finparatrains,'Enable','on');

set(handles.text136,'string',etaini);
set(handles.text157,'string',etaini);

handles.etaini = etaini;
guidata(hObject, handles);


% --- Executes on button press in confirmarfintrain.
function confirmarfintrain_Callback(hObject, eventdata, handles)
global finpara inipara
    
finpara = str2double(get(handles.finparatrains,'string'));

if isnan(finpara)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.finparatrains,'string','0')
return;
end

if finpara == inipara
errordlg({'The parallel trains number finals can not be same as the parallel trains initials'},'ERROR','modal');
set(handles.finparatrains,'string','0')
return;
end

if finpara > 5 
errordlg({'The maximum number for parallel trains finals is 5'},'ERROR','modal');
set(handles.finparatrains,'string','0')
return;
end

if isempty(finpara) || finpara <= 0
errordlg('The minimum number for parallel trains finals is 1','ERROR','modal')
set(handles.finparatrains,'string','0')
return;
end

set(handles.text145,'String',finpara);
set(handles.text173,'String',finpara);
set(handles.text214,'Visible','on');
set(handles.finparatrains,'Enable','off')
set(handles.confirmarfintrain,'Enable','off')
set(handles.etapafinparatrain,'Visible','on');
set(handles.etapafinparatrain,'string',1);
set(handles.confirmaretapfin,'Visible','on');


handles.finpara = finpara;
guidata(hObject, handles);
    

% --- Executes on button press in confirmaretapfin.
function confirmaretapfin_Callback(hObject, eventdata, handles)
global etafin go1 go3 etaini se para etapaserie et

etafin = str2double(get(handles.etapafinparatrain,'string'));

if isnan(etafin)
errordlg('The value entered must be numeric','ERROR','modal')
set(handles.etapafinparatrain,'string','1')
return;
end

if etafin > 1 
errordlg({'The maximum number for stages is 1'},'ERROR','modal');
set(handles.etapafinparatrain,'string','1')
return;
end

if isempty(etafin) || etafin <= 0 
errordlg('The minimum value for stages is 1','ERROR','modal')
set(handles.etapafinparatrain,'string','1')
return;
end

set(handles.etapafinparatrain,'Enable','off');
set(handles.confirmaretapfin,'Enable','off');
set(handles.uipanel26,'title','With change in the number of parallel trains');
set(handles.uipanel40,'visible','on');

set(handles.text143,'string',etafin);
set(handles.text175,'string',etafin);

warndlg({'The configuration is correct'},'NOTIFICATION','modal');
% go3 == 1;
% if go1 == 1 && go3 ==1
%    set(handles.Go_To_Plant,'Enable','on'); 
% end

set(handles.Panel31,'visible','off');
set(handles.resusinreduc,'visible','off');
set(handles.resutrasta,'visible','on');
set(handles.resuconreduc,'visible','on');
set(handles.compl2,'visible','on');
set(handles.nocompl2,'visible','off');
clear se para etapaserie et
handles.etafin = etafin;
guidata(hObject, handles);

% --- Executes on button press in Go_To_Plant.
function Go_To_Plant_Callback(hObject, eventdata, handles)
global  hyapp sp_condiciones entrada SP_DELETE ruta presionent tempent...
        molfloent presionsal molflosal comp1 comp2 comp3 comp4 comp5...
        comp6 comp7 comp8 comp9 comp10 comp11 comp12 comp13 comp14 comp15...
        comp16 comp17 se etapasetren1 etapasetren2 para...
        exit_4_0 exit_104_1_1 exit_69_1_1 exit_38_1_1 exit_6_0 exit_68_2_1...
        exit_37_2_1 exit_4_0_2 exit_104_1_2 exit_69_1_2 exit_38_1_2...
        exit_6_0_2 exit_68_2_2 exit_37_2_2 exit_4_0_3 exit_104_1_3... 
        exit_69_1_3 exit_38_1_3 exit_6_0_3 exit_68_2_3 exit_37_2_3...
        exit_4_0_4 exit_104_1_4 exit_69_1_4 exit_38_1_4 exit_6_0_4...
        exit_68_2_4 exit_37_2_4 exit_4_0_5 exit_104_1_5 exit_69_1_5...
        exit_38_1_5 exit_6_0_5 exit_68_2_5 exit_37_2_5 diametro longitud...
        actpre actmol UnitOps inipara etaini finpara etafin booleano...
        opc1 opc2...
           
%*************************************************************************%
%***********************CONECTIVIDAD MATLAB-HYSYS*************************%
%*************************************************************************%
% hyapp = hyconnect('C:\Users\Marcos Cueto\Dropbox\TESIS\Simulacion\simulacion_dinamica_02_11_2020.hsc',1);
% hyapp = hyconnect('C:\Users\Gusta\Dropbox\TESIS\Simulacion\simulacion_dinamica_28_11_2020.hsc',1);
ruta = get(handles.file,'string');
if isempty(ruta)
errordlg('File path is not available','ERROR','modal')
set(handles.file,'string','')
return;
else
pause (0.001);       
set(handles.nocompl3,'visible','off');
pause (0.001);
set(handles.compl3,'visible','on');
set(handles.compl3,'string','Sending data in process');
pause (0.001);    
hyapp = hyconnect(ruta,1);
booleano = true;
end

%*************************************************************************%
%**********************DECLARACION DE LAS CORRIENTES**********************%
%*************************************************************************%
entrada = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('Alimentacion');

exit_4_0 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('4-0');
exit_104_1_1 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('104-1-1');
exit_69_1_1 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('69-1-1');
exit_38_1_1 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('38-1-1');
exit_6_0 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('6-0');
exit_68_2_1 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('68-2-1');
exit_37_2_1 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('37-2-1');

exit_4_0_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('4-0-2');
exit_104_1_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('104-1-2');
exit_69_1_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('69-1-2');
exit_38_1_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('38-1-2');
exit_6_0_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('6-0-2');
exit_68_2_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('68-2-2');
exit_37_2_2 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('37-2-2');

exit_4_0_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('4-0-3');
exit_104_1_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('104-1-3');
exit_69_1_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('69-1-3');
exit_38_1_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('38-1-3');
exit_6_0_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('6-0-3');
exit_68_2_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('68-2-3');
exit_37_2_3 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('37-2-3');

exit_4_0_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('4-0-4');
exit_104_1_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('104-1-4');
exit_69_1_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('69-1-4');
exit_38_1_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('38-1-4');
exit_6_0_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('6-0-4');
exit_68_2_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('68-2-4');
exit_37_2_4 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('37-2-4');

exit_4_0_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('4-0-5');
exit_104_1_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('104-1-5');
exit_69_1_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('69-1-5');
exit_38_1_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('38-1-5');
exit_6_0_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('6-0-5');
exit_68_2_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('68-2-5');
exit_37_2_5 = hyapp.ActiveDocument.Flowsheet.MaterialStreams.Item('37-2-5');


%*************************************************************************%
%************DECLARACION DEL SPREADSHEET PARA CONDICIONES*****************% 
%*************************************************************************%
sp_condiciones = hyapp.ActiveDocument.Flowsheet.Operations.Item('condiciones');

%*************************************************************************%
%**********DECLARACION DE SPREADSHEET PARA ELIMINAR TRENES****************% 
%*************************************************************************%
SP_DELETE = hyapp.ActiveDocument.Flowsheet.Operations.Item('DELETE');

%*************************************************************************%
%**********DECLARACION DE SPREADSHEET PARA DIMENSIONAMIENTO***************% 
%*************************************************************************%
UnitOps = hyapp.ActiveDocument.Flowsheet.Operations.Item('UnitOps');

%*************************************************************************%
%****************************TRENES EN SERIE******************************%
%*************************************************************************%
etapasetren1 = str2double(get(handles.text34,'string'));
etapasetren2 = str2double(get(handles.text35,'string'));

if opc1 == 1
if para == 0 || para == 1
if se == 1
    %*********************ESTO ELIMINA TRENE 2-1**************************%
    for E = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    %**************ESTO ELIMINA TRENES PARALELO 1-2 y 2-2*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-3*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Q = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    for QI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for Z = 6:184
         if booleano == false
            break;
         end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
     end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    switch etapasetren1
        case 4
            SP_DELETE.Cell('B256').CellValue = 1;
            sp_condiciones.Cell('C8').CellValue = 0;
        case 3
            for B = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C10').CellValue = 0;
        case 2
            for B = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C12').CellValue = 0;
        case 1            
            for B = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C14').CellValue = 0;      
    end
end


if se == 2   
    %**************ESTO ELIMINA TRENES PARALELO 1-2 y 2-2*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-3*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for Q = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    for QI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for Z = 6:184
         if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
     end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end    
    if etapasetren1 == 4
        switch etapasetren2
            case 3
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C16').CellValue = 0;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
        end
    end
    if etapasetren1 == 3
        switch etapasetren2
            case 3
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                sp_condiciones.Cell('F21').CellValue = 100;
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
        end
    end
    if etapasetren1 == 2
        switch etapasetren2
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
        end 
    end
    if etapasetren1 == 1
        switch etapasetren2
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B248').CellValue = 0;
                else
                SP_DELETE.Cell('B248').CellValue = 1;
                end
%                 SP_DELETE.Cell('B248').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F23').CellValue = 100;
        end
    end
end
end


%*************************************************************************%
%************************TRENES EN PARALELO*******************************%
%*************************************************************************%
if para == 2
    
if se == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-1 y 2-2*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-3*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for E = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for Z = 6:184
         if booleano == false
            break;
         end
         pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
     end
    for AF = 2:180
         if booleano == false
            break;
         end
         pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AC = 2:247
         if booleano == false
            break;
         end
         pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
         if booleano == false
            break;
         end
         pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
         if booleano == false
            break;
         end
         pause (1E-08);
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    switch etapasetren1
        case 4
            SP_DELETE.Cell('B256').CellValue = 1;
            sp_condiciones.Cell('C8').CellValue = 0;
            SP_DELETE.Cell('Q256').CellValue = 1;
            sp_condiciones.Cell('C20').CellValue = 0;
        case 3
            for B = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C10').CellValue = 0;
            sp_condiciones.Cell('C22').CellValue = 0;
        case 2
            for B = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C12').CellValue = 0;
            sp_condiciones.Cell('C24').CellValue = 0;
        case 1
            for B = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08);
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C14').CellValue = 0;
            sp_condiciones.Cell('C26').CellValue = 0;
    end
end


if se == 2
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-3*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    if etapasetren1 == 4
        switch etapasetren2
            case 3  
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('C28').CellValue = 0;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('C29').CellValue = 0;
        end
    end
    if etapasetren1 == 3
        switch etapasetren2
            case 3
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('F24').CellValue = 100;
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('C28').CellValue = 0;
                sp_condiciones.Cell('F24').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F24').CellValue = 100;
        end
    end
    if etapasetren1 == 2
        switch etapasetren2
            case 2 
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                SP_DELETE.Cell('Q249').CellValue = 0;    
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                SP_DELETE.Cell('Q249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
%                 SP_DELETE.Cell('Q249').CellValue = 1;
                sp_condiciones.Cell('C28').CellValue = 0;
                sp_condiciones.Cell('F25').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                SP_DELETE.Cell('Q249').CellValue = 0;    
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                SP_DELETE.Cell('Q249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
%                 SP_DELETE.Cell('Q249').CellValue = 1;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F25').CellValue = 100;
        end 
    end
    if etapasetren1 == 1
        switch etapasetren2
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08);
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B248').CellValue = 0;
                SP_DELETE.Cell('Q248').CellValue = 0;    
                else
                SP_DELETE.Cell('B248').CellValue = 1;
                SP_DELETE.Cell('Q248').CellValue = 1;
                end
%                 SP_DELETE.Cell('B248').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F23').CellValue = 100;
%                 SP_DELETE.Cell('Q248').CellValue = 1;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F26').CellValue = 100;
        end
    end
end
end


if para == 3
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    
if se == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-1 y 2-2*****************%
    %******************ESTO ELIMINA TRENES PARALELO 2-3 ******************%
    for E = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    switch etapasetren1
        case 4
            SP_DELETE.Cell('B256').CellValue = 1;
            sp_condiciones.Cell('C8').CellValue = 0;
            SP_DELETE.Cell('Q256').CellValue = 1;
            sp_condiciones.Cell('C20').CellValue = 0;
            SP_DELETE.Cell('H256').CellValue = 1;
            sp_condiciones.Cell('C40').CellValue = 0;
        case 3
            for B = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            for H = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C10').CellValue = 0;
            sp_condiciones.Cell('C22').CellValue = 0;
            sp_condiciones.Cell('C41').CellValue = 0;
        case 2
            for B = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end         
            for H = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C12').CellValue = 0;
            sp_condiciones.Cell('C24').CellValue = 0;
            sp_condiciones.Cell('C42').CellValue = 0;
        case 1
            for B = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end            
            for H = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C14').CellValue = 0;
            sp_condiciones.Cell('C26').CellValue = 0;
            sp_condiciones.Cell('C43').CellValue = 0;
    end
end


if se == 2
    if etapasetren1 == 4
        switch etapasetren2
            case 3
            case 2
               for E = 124:175
                   if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end                
                for K = 124:175
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 180:181
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 118:120
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end 
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('C28').CellValue = 0;
                sp_condiciones.Cell('C44').CellValue = 0;
            case 1
                for E = 70:175
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                       break;
                   end
                   pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end                
                for K = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('C45').CellValue = 0;
        end
    end
    if etapasetren1 == 3
        switch etapasetren2
            case 3
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('F24').CellValue = 100;
                sp_condiciones.Cell('F27').CellValue = 100;
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for K = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end                
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('C28').CellValue = 0;
                sp_condiciones.Cell('F24').CellValue = 100;
                sp_condiciones.Cell('C44').CellValue = 0;
                sp_condiciones.Cell('F27').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for K = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F21').CellValue = 100;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F24').CellValue = 100;
                sp_condiciones.Cell('C45').CellValue = 0;
                sp_condiciones.Cell('F27').CellValue = 100;
        end
    end
    if etapasetren1 == 2
        switch etapasetren2
            case 2
                for E = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for K = 124:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 180:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 118:120
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end                
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                SP_DELETE.Cell('Q249').CellValue = 0;
                SP_DELETE.Cell('H249').CellValue = 0;
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                SP_DELETE.Cell('Q249').CellValue = 1;
                SP_DELETE.Cell('H249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C16').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
%                 SP_DELETE.Cell('Q249').CellValue = 1;
                sp_condiciones.Cell('C28').CellValue = 0;
                sp_condiciones.Cell('F25').CellValue = 100;
%                 SP_DELETE.Cell('H249').CellValue = 1;
                sp_condiciones.Cell('C44').CellValue = 0;
                sp_condiciones.Cell('F28').CellValue = 100;
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for K = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end                
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B249').CellValue = 0;
                SP_DELETE.Cell('Q249').CellValue = 0;
                SP_DELETE.Cell('H249').CellValue = 0;
                else
                SP_DELETE.Cell('B249').CellValue = 1;
                SP_DELETE.Cell('Q249').CellValue = 1;
                SP_DELETE.Cell('H249').CellValue = 1;
                end
%                 SP_DELETE.Cell('B249').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F22').CellValue = 100;
%                 SP_DELETE.Cell('Q249').CellValue = 1;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F25').CellValue = 100;
%                 SP_DELETE.Cell('H249').CellValue = 1;
                sp_condiciones.Cell('C45').CellValue = 0;
                sp_condiciones.Cell('F28').CellValue = 100;
        end
    end
    if etapasetren1 == 1
        switch etapasetren2
            case 1
                for E = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                end
                for EI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                end
                for EII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                end
                for EIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                end
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                for BIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                end                
                for T = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                end
                for TI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                end
                for TII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                end
                for TIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                end
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for QIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                end                
                for K = 70:175
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                end
                for KI = 178:181
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                end
                for KII = 64:66
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                end
                for KIII = 182:184
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 20:22
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for HIII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B248').CellValue = 0;
                SP_DELETE.Cell('Q248').CellValue = 0;
                SP_DELETE.Cell('H248').CellValue = 0;
                else
                SP_DELETE.Cell('B248').CellValue = 1;
                SP_DELETE.Cell('Q248').CellValue = 1;
                SP_DELETE.Cell('H248').CellValue = 1;
                end
%                 SP_DELETE.Cell('B248').CellValue = 1;
                sp_condiciones.Cell('C17').CellValue = 0;
                sp_condiciones.Cell('F23').CellValue = 100;
%                 SP_DELETE.Cell('Q248').CellValue = 1;
                sp_condiciones.Cell('C29').CellValue = 0;
                sp_condiciones.Cell('F26').CellValue = 100;
%                 SP_DELETE.Cell('H248').CellValue = 1;
                sp_condiciones.Cell('C45').CellValue = 0;
                sp_condiciones.Cell('F29').CellValue = 100;
        end
    end
end
end


if para == 4
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    
    if se == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-1 y 2-2*****************%
    %**************ESTO ELIMINA TRENES PARALELO 2-3 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for E = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z= 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    switch etapasetren1
        case 4
            SP_DELETE.Cell('B256').CellValue = 1;
            sp_condiciones.Cell('C8').CellValue = 0;
            SP_DELETE.Cell('Q256').CellValue = 1;
            sp_condiciones.Cell('C20').CellValue = 0;
            SP_DELETE.Cell('H256').CellValue = 1;
            sp_condiciones.Cell('C40').CellValue = 0;
            SP_DELETE.Cell('W256').CellValue = 1;
            sp_condiciones.Cell('C53').CellValue = 0;
        case 3
            for B = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            for H = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C10').CellValue = 0;
            sp_condiciones.Cell('C22').CellValue = 0;
            sp_condiciones.Cell('C41').CellValue = 0;
            sp_condiciones.Cell('C55').CellValue = 0;
        case 2
            for B = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end            
            for H = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C12').CellValue = 0;
            sp_condiciones.Cell('C24').CellValue = 0;
            sp_condiciones.Cell('C42').CellValue = 0;
            sp_condiciones.Cell('C57').CellValue = 0;
        case 1
            for B = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end            
            for H = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C14').CellValue = 0;
            sp_condiciones.Cell('C26').CellValue = 0;
            sp_condiciones.Cell('C43').CellValue = 0;
            sp_condiciones.Cell('C59').CellValue = 0; 
    end    
    end
    
    if se == 2
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                case 2
                    for E = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('C63').CellValue = 0;
                case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('C65').CellValue = 0;
            end
        end
        if etapasetren1 == 3
            switch etapasetren2
                case 3
                    for B = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for Q = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for H = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                
                    for W = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('F30').CellValue = 100;
                case 2
                    for E = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('C63').CellValue = 0;
                    sp_condiciones.Cell('F30').CellValue = 100;
                case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end
                    for H = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F30').CellValue = 100;
            end
        end
        if etapasetren1 == 2
            switch etapasetren2
                case 2
                    for E = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    if booleano == false
                    SP_DELETE.Cell('B249').CellValue = 0;
                    SP_DELETE.Cell('Q249').CellValue = 0;
                    SP_DELETE.Cell('H249').CellValue = 0;
                    SP_DELETE.Cell('W249').CellValue = 0;
                    else
                    SP_DELETE.Cell('B249').CellValue = 1;
                    SP_DELETE.Cell('Q249').CellValue = 1;
                    SP_DELETE.Cell('H249').CellValue = 1;
                    SP_DELETE.Cell('W249').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B249').CellValue = 1;
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('F22').CellValue = 100;
%                     SP_DELETE.Cell('Q249').CellValue = 1;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('F25').CellValue = 100;
%                     SP_DELETE.Cell('H249').CellValue = 1;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('F28').CellValue = 100;
%                     SP_DELETE.Cell('W249').CellValue = 1;
                    sp_condiciones.Cell('C63').CellValue = 0;
                    sp_condiciones.Cell('F31').CellValue = 100;
                case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    if booleano == false
                    SP_DELETE.Cell('B249').CellValue = 0;
                    SP_DELETE.Cell('Q249').CellValue = 0;
                    SP_DELETE.Cell('H249').CellValue = 0;
                    SP_DELETE.Cell('W249').CellValue = 0;
                    else
                    SP_DELETE.Cell('B249').CellValue = 1;
                    SP_DELETE.Cell('Q249').CellValue = 1;
                    SP_DELETE.Cell('H249').CellValue = 1;
                    SP_DELETE.Cell('W249').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B249').CellValue = 1;
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F22').CellValue = 100;
%                     SP_DELETE.Cell('Q249').CellValue = 1;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F25').CellValue = 100;
%                     SP_DELETE.Cell('H249').CellValue = 1;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F28').CellValue = 100;
%                     SP_DELETE.Cell('W249').CellValue = 1;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F31').CellValue = 100;
            end
        end
        if etapasetren1 == 1
            switch etapasetren2
                case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end
                    for H = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                     if booleano == false
                    SP_DELETE.Cell('B248').CellValue = 0;
                    SP_DELETE.Cell('Q248').CellValue = 0;
                    SP_DELETE.Cell('H248').CellValue = 0;
                    SP_DELETE.Cell('W248').CellValue = 0;
                    else
                    SP_DELETE.Cell('B248').CellValue = 1;
                    SP_DELETE.Cell('Q248').CellValue = 1;
                    SP_DELETE.Cell('H248').CellValue = 1;
                    SP_DELETE.Cell('W248').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B248').CellValue = 1;
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F23').CellValue = 100;
%                     SP_DELETE.Cell('Q248').CellValue = 1;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F26').CellValue = 100;
%                     SP_DELETE.Cell('H248').CellValue = 1;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F29').CellValue = 100;
%                     SP_DELETE.Cell('W248').CellValue = 1;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F32').CellValue = 100;  
            end
        end
    end
end

if para == 5
    %**********************ESTO ELIMINA DISTRIBUIDOR**********************%
    for N = 2:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end

    if se == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-1 y 2-2*****************%
    %**************ESTO ELIMINA TRENES PARALELO 2-3 y 2-4*****************%
    %*****************ESTO ELIMINA TRENES PARALELO 2-5********************%
    for E = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z= 6:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    switch etapasetren1
        case 4
            SP_DELETE.Cell('B256').CellValue = 1;
            sp_condiciones.Cell('C8').CellValue = 0;
            SP_DELETE.Cell('Q256').CellValue = 1;
            sp_condiciones.Cell('C20').CellValue = 0;
            SP_DELETE.Cell('H256').CellValue = 1;
            sp_condiciones.Cell('C40').CellValue = 0;
            SP_DELETE.Cell('W256').CellValue = 1;
            sp_condiciones.Cell('C53').CellValue = 0;
            SP_DELETE.Cell('AC256').CellValue = 1;
            sp_condiciones.Cell('C69').CellValue = 0;
        case 3
            for B = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end
            for H = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            for AC = 182:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
            end
            for ACI = 246:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
            end
            for ACII = 176:178
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
            end
            for ACIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C10').CellValue = 0;
            sp_condiciones.Cell('C22').CellValue = 0;
            sp_condiciones.Cell('C41').CellValue = 0;
            sp_condiciones.Cell('C55').CellValue = 0;
            sp_condiciones.Cell('C71').CellValue = 0;
        case 2
           for B = 122:239
               if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end            
            for H = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            for AC = 122:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
            end
            for ACI = 244:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
            end
            for ACII = 116:118
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
            end
            for ACIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C12').CellValue = 0;
            sp_condiciones.Cell('C24').CellValue = 0;
            sp_condiciones.Cell('C42').CellValue = 0;
            sp_condiciones.Cell('C57').CellValue = 0;
            sp_condiciones.Cell('C73').CellValue = 0; 
        case 1
            for B = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
            end
            for BI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
            end
            for BII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
            end
            for BIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
            end            
            for Q = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
            end
            for QI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
            end
            for QII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
            end
            for QIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
            end            
            for H = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
            end
            for HI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
            end
            for HII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
            end
            for HIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
            end            
            for W = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
            end
            for WI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
            end
            for WII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
            end
            for WIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
            end
            for AC = 68:239
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
            end
            for ACI = 62:64
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
            end
            for ACII = 242:247
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
            end
            for ACIII = 253:256
                if booleano == false
                    break;
                end
                pause (1E-08)
                SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
            end
            sp_condiciones.Cell('C14').CellValue = 0;
            sp_condiciones.Cell('C26').CellValue = 0;
            sp_condiciones.Cell('C43').CellValue = 0;
            sp_condiciones.Cell('C59').CellValue = 0;
            sp_condiciones.Cell('C75').CellValue = 0;
    end
    end
    
    if se == 2
       if etapasetren1 == 4
           switch etapasetren2
               case 3
               case 2
                    for E = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for AF = 120:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 176:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 114:116
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('C63').CellValue = 0;
                    sp_condiciones.Cell('C80').CellValue = 0;
               case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for AF = 66:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 174:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 60:62
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end 
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('C81').CellValue = 0;                   
           end                   
       end
       if etapasetren1 == 3
           switch etapasetren2
               case 3
                   for B = 246:247
                       if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for Q = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for H = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                
                    for W = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AC = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('F30').CellValue = 100;
                    sp_condiciones.Cell('F33').CellValue = 100;
               case 2
                    for E = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AF = 120:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 176:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 114:116
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    for AC = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end  
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('C63').CellValue = 0;
                    sp_condiciones.Cell('F30').CellValue = 100;
                    sp_condiciones.Cell('C80').CellValue = 0;
                    sp_condiciones.Cell('F33').CellValue = 100;
               case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end
                    for H = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AF = 66:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 174:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 60:62
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    for AC = 179:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 246:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 116:118
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F21').CellValue = 100;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F24').CellValue = 100;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F27').CellValue = 100;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F30').CellValue = 100;
                    sp_condiciones.Cell('C81').CellValue = 0;
                    sp_condiciones.Cell('F34').CellValue = 100;
           end
       end
       if etapasetren1 == 2
           switch etapasetren2
               case 2
                   for E = 124:175
                       if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 124:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 180:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 118:120
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AF = 120:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 176:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 114:116
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    for AC = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end
                    if booleano == false
                    SP_DELETE.Cell('B249').CellValue = 0;
                    SP_DELETE.Cell('Q249').CellValue = 0;
                    SP_DELETE.Cell('H249').CellValue = 0;
                    SP_DELETE.Cell('W249').CellValue = 0;
                    SP_DELETE.Cell('AC249').CellValue = 0;
                    else
                    SP_DELETE.Cell('B249').CellValue = 1;
                    SP_DELETE.Cell('Q249').CellValue = 1;
                    SP_DELETE.Cell('H249').CellValue = 1;
                    SP_DELETE.Cell('W249').CellValue = 1;
                    SP_DELETE.Cell('AC249').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B249').CellValue = 1;
                    sp_condiciones.Cell('C16').CellValue = 0;
                    sp_condiciones.Cell('F22').CellValue = 100;
%                     SP_DELETE.Cell('Q249').CellValue = 1;
                    sp_condiciones.Cell('C28').CellValue = 0;
                    sp_condiciones.Cell('F25').CellValue = 100;
%                     SP_DELETE.Cell('H249').CellValue = 1;
                    sp_condiciones.Cell('C44').CellValue = 0;
                    sp_condiciones.Cell('F28').CellValue = 100;
%                     SP_DELETE.Cell('W249').CellValue = 1;
                    sp_condiciones.Cell('C63').CellValue = 0;
                    sp_condiciones.Cell('F31').CellValue = 100;
%                     SP_DELETE.Cell('AC249').CellValue = 1;
                    sp_condiciones.Cell('C80').CellValue = 0;
                    sp_condiciones.Cell('F34').CellValue = 100;
               case 1
                   for E = 70:175
                       if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end                
                    for H = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AF = 66:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 174:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 60:62
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    for AC = 119:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 244:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 62:64
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end
                    if booleano == false
                    SP_DELETE.Cell('B249').CellValue = 0;
                    SP_DELETE.Cell('Q249').CellValue = 0;
                    SP_DELETE.Cell('H249').CellValue = 0;
                    SP_DELETE.Cell('W249').CellValue = 0;
                    SP_DELETE.Cell('AC249').CellValue = 0;
                    else
                    SP_DELETE.Cell('B249').CellValue = 1;
                    SP_DELETE.Cell('Q249').CellValue = 1;
                    SP_DELETE.Cell('H249').CellValue = 1;
                    SP_DELETE.Cell('W249').CellValue = 1;
                    SP_DELETE.Cell('AC249').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B249').CellValue = 1;
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F22').CellValue = 100;
%                     SP_DELETE.Cell('Q249').CellValue = 1;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F25').CellValue = 100;
%                     SP_DELETE.Cell('H249').CellValue = 1;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F28').CellValue = 100;
%                     SP_DELETE.Cell('W249').CellValue = 1;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F31').CellValue = 100;
%                     SP_DELETE.Cell('AC249').CellValue = 1;
                    sp_condiciones.Cell('C81').CellValue = 0;
                    sp_condiciones.Cell('F34').CellValue = 100;                   
           end
       end
       if etapasetren1 == 1
           switch etapasetren2
               case 1
                    for E = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
                    end
                    for EI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
                    end
                    for EII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
                    end
                    for EIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('E',num2str(EIII))).CellValue = 1;
                    end
                    for B = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    for BI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                    end
                    for BII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                    end
                    for BIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('B',num2str(BIII))).CellValue = 1;
                    end                
                    for T = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
                    end
                    for TI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
                    end
                    for TII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
                    end
                    for TIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('T',num2str(TIII))).CellValue = 1;
                    end
                    for Q = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                    end
                    for QI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                    end
                    for QII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                    end
                    for QIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Q',num2str(QIII))).CellValue = 1;
                    end                
                    for K = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
                    end
                    for KI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
                    end
                    for KII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
                    end
                    for KIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('K',num2str(KIII))).CellValue = 1;
                    end
                    for H = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                    end
                    for HI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                    end
                    for HII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                    end
                    for HIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('H',num2str(HIII))).CellValue = 1;
                    end                    
                    for Z = 70:175
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
                    end
                    for ZI = 178:181
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
                    end
                    for ZII = 64:66
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
                    end
                    for ZIII = 182:184
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('Z',num2str(ZIII))).CellValue = 1;
                    end
                    for W = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                    end
                    for WI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                    end
                    for WII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                    end
                    for WIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('W',num2str(WIII))).CellValue = 1;
                    end
                    for AF = 66:171
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
                    end
                    for AFI = 174:177
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
                    end
                    for AFII = 60:62
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
                    end
                    for AFIII = 178:180
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
                    end
                    for AC = 65:239
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                    end
                    for ACI = 20:22
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                    end
                    for ACII = 242:247
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                    end
                    for ACIII = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08)
                        SP_DELETE.Cell(strcat('AC',num2str(ACIII))).CellValue = 1;
                    end
                    if booleano == false
                    SP_DELETE.Cell('B248').CellValue = 0;
                    SP_DELETE.Cell('Q248').CellValue = 0;
                    SP_DELETE.Cell('H248').CellValue = 0;
                    SP_DELETE.Cell('W248').CellValue = 0;
                    SP_DELETE.Cell('AC248').CellValue = 0;
                    else
                    SP_DELETE.Cell('B248').CellValue = 1;
                    SP_DELETE.Cell('Q248').CellValue = 1;
                    SP_DELETE.Cell('H248').CellValue = 1;
                    SP_DELETE.Cell('W248').CellValue = 1;
                    SP_DELETE.Cell('AC248').CellValue = 1;
                    end
%                     SP_DELETE.Cell('B248').CellValue = 1;
                    sp_condiciones.Cell('C17').CellValue = 0;
                    sp_condiciones.Cell('F23').CellValue = 100;
%                     SP_DELETE.Cell('Q248').CellValue = 1;
                    sp_condiciones.Cell('C29').CellValue = 0;
                    sp_condiciones.Cell('F26').CellValue = 100;
%                     SP_DELETE.Cell('H248').CellValue = 1;
                    sp_condiciones.Cell('C45').CellValue = 0;
                    sp_condiciones.Cell('F29').CellValue = 100;
%                     SP_DELETE.Cell('W248').CellValue = 1;
                    sp_condiciones.Cell('C65').CellValue = 0;
                    sp_condiciones.Cell('F32').CellValue = 100;
%                     SP_DELETE.Cell('AC248').CellValue = 1;
                    sp_condiciones.Cell('C81').CellValue = 0;
                    sp_condiciones.Cell('F35').CellValue = 100;
           end
       end
    end
end
end


%*************************************************************************%
%****************************AMPLIACION TRENES****************************%
%*************************************************************************%
if opc2 == 1
    if inipara == 1 && finpara == 2
    %**************ESTO ELIMINA TRENES PARALELO 1-2 y 2-3*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-5*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for Q = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    for QI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
    end
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
         if booleano == false
            break;
         end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for N = 9:11
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 13:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
                for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;                
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;    
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;    
                end
        end
    end
    end
    if inipara == 1 && finpara == 3
        %************ESTO ELIMINA TRENES PARALELO 1-2*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-3 y 2-4***************%
        %************ESTO ELIMINA TRENES PARALELO 1-4 y 2-5***************%
        %************ESTO ELIMINA TRENES PARALELO 1-5*********************%
        for Q = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
        end
        for QI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
        end
        for H = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
        end
        for HI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
        end
        for W = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
        end
        for WI = 253:256
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
        end
        for AC = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
        end
        for ACI = 253:256
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
        end        
        for Z = 6:184
             if booleano == false
                break;
             end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
        end
        for AF = 2:180
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
        end
        %*****************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end        
     for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end 
        %*****************************************************************%
        for N = 9:11
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
        end
        for NI = 13:15
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
        end
        for NII = 17:24
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NII))).CellValue = 1;
        end
        %*****************************************************************%
        sp_condiciones.Cell('C17').CellValue = 0;
        sp_condiciones.Cell('C29').CellValue = 0;
        sp_condiciones.Cell('C45').CellValue = 0;
        if etafin == 1
            switch etaini
                case 4
                    for B = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08);
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    %*****************************************************%
                    sp_condiciones.Cell('F37').CellValue = 100;
                case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;    
                end
                case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;    
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;    
                end
            end
        end
    end
    if inipara == 1 && finpara == 4
        %************ESTO ELIMINA TRENES PARALELO 1-2*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-3*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-4 y 2-5***************%
        %************ESTO ELIMINA TRENES PARALELO 1-5*********************%
        for Q = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
        end
        for QI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
        end
        for H = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
        end
        for HI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
        end
        for W = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
        end
        for WI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
        end
        for AC = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
        end
        for ACI = 253:256
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
        end        
        for AF = 2:180
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
        end
        %*****************************************************************%
     for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end        
     for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end 
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end 
        %*****************************************************************%
        for N = 9:11
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
        end
        for NI = 13:15
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
        end
        for NII = 17:19
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NII))).CellValue = 1;
        end
        for NIII = 21:24
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NIII))).CellValue = 1;
        end
        %*****************************************************************%
        sp_condiciones.Cell('C17').CellValue = 0;
        sp_condiciones.Cell('C29').CellValue = 0;
        sp_condiciones.Cell('C45').CellValue = 0;
        sp_condiciones.Cell('C65').CellValue = 0;
        if etafin == 1
            switch etaini
                case 4
                    for B = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08);
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    %*****************************************************%
                    sp_condiciones.Cell('F37').CellValue = 100;
                case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;    
                end
                case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;    
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;    
                end
            end
        end
    end
    if inipara == 1 && finpara == 5
        %************ESTO ELIMINA TRENES PARALELO 1-2*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-3*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-4*********************%
        %************ESTO ELIMINA TRENES PARALELO 1-5*********************%
        for Q = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
        end
        for QI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
        end
        for H = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
        end
        for HI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
        end
        for W = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
        end
        for WI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
        end
        for AC = 2:247
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
        end
        for ACI = 253:255
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
        end
        %*****************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end        
     for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end 
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for AF = 66:166
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AFI = 60:62
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
    end
    for AFII = 174:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
    end
    for AFIII = 167:171
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
    end
        %*****************************************************************%
        for N = 9:11
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
        end
        for NI = 13:15
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
        end
        for NII = 17:19
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NII))).CellValue = 1;
        end
        for NIII = 21:23
            if booleano == false
                break;
            end
            pause (1E-08);
            SP_DELETE.Cell(strcat('N',num2str(NIII))).CellValue = 1;
        end
        %*****************************************************************%
        sp_condiciones.Cell('C17').CellValue = 0;
        sp_condiciones.Cell('C29').CellValue = 0;
        sp_condiciones.Cell('C45').CellValue = 0;
        sp_condiciones.Cell('C65').CellValue = 0;
        sp_condiciones.Cell('C81').CellValue = 0;
        if etafin == 1
            switch etaini
                case 4
                    for B = 253:255
                        if booleano == false
                            break;
                        end
                        pause (1E-08);
                        SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                    end
                    %*****************************************************%
                    sp_condiciones.Cell('F37').CellValue = 100;
                case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;    
                end
                case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;    
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;    
                end
            end
        end
    end
    if inipara == 2 && finpara == 3
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-5*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for Z = 6:184
         if booleano == false
            break;
         end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for N = 13:15
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 17:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
                for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end     
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                end
        end
    end
    end
    if inipara == 2 && finpara == 4
        %**************ESTO ELIMINA TRENES PARALELO 1-3***********************%
        %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-5*****************%
        %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for N = 13:15
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 17:19
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end
    for NII = 21:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NII))).CellValue = 1;
    end
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end     
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                end
        end
    end
    end
    if inipara == 2 && finpara == 5
    %**************ESTO ELIMINA TRENES PARALELO 1-3***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for AF = 66:166
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AFI = 60:62
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
    end
    for AFII = 174:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
    end
    for AFIII = 167:171
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
    end
    for N = 13:15
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 17:19
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end
    for NII = 21:23
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NII))).CellValue = 1;
    end
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    sp_condiciones.Cell('C81').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
                for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end     
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                end
        end
    end
    end
    if inipara == 3 && finpara == 4
    %**************ESTO ELIMINA TRENES PARALELO 2-5***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for N = 17:19
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 21:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
                for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end 
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                end
        end
    end
    end
    if inipara == 3 && finpara == 5
    %**************ESTO ELIMINA TRENES PARALELO 1-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for AF = 66:166
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AFI = 60:62
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
    end
    for AFII = 174:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
    end
    for AFIII = 167:171
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
    end
    for N = 17:19
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    for NI = 21:23
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(NI))).CellValue = 1;
    end
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    sp_condiciones.Cell('C81').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
            case 3
                 for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end 
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                end
        end
    end
    end
    if inipara == 4 && finpara == 5
    %**************ESTO ELIMINA TRENES PARALELO 1-5***********************%
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:255
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end
    for AF = 66:166
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for AFI = 60:62
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFI))).CellValue = 1;
    end
    for AFII = 174:180
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFII))).CellValue = 1;
    end
    for AFIII = 167:171
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AF',num2str(AFIII))).CellValue = 1;
    end
    for N = 21:23
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    sp_condiciones.Cell('C81').CellValue = 0; 
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                end
        end
    end
    end
    
%*************************************************************************%
%****************************REDUCCION TRENES*****************************%
%*************************************************************************%
if inipara == 5 && finpara == 4
    %**************ESTO ELIMINA TRENES PARALELO 2-5***********************%
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end   
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    SP_DELETE.Cell('N24').CellValue = 1;
    SP_DELETE.Cell('AC256').CellValue = 1;
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for AC = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
                sp_condiciones.Cell('F41').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                SP_DELETE.Cell('AC252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                SP_DELETE.Cell('AC252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                SP_DELETE.Cell('AC251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                SP_DELETE.Cell('AC251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                SP_DELETE.Cell('AC250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                SP_DELETE.Cell('AC250').CellValue = 1;
                end
        end
    end
end

if inipara == 5 && finpara == 3
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-5***********************%
    for Z = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end   
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    SP_DELETE.Cell('N24').CellValue = 1;
    SP_DELETE.Cell('AC256').CellValue = 1;
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    sp_condiciones.Cell('C65').CellValue = 0;
    %*********************************************************************%
    SP_DELETE.Cell('N20').CellValue = 1;
    SP_DELETE.Cell('N24').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    SP_DELETE.Cell('AC256').CellValue = 1;
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for AC = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
                sp_condiciones.Cell('F41').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                SP_DELETE.Cell('AC252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                SP_DELETE.Cell('AC252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                SP_DELETE.Cell('AC251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                SP_DELETE.Cell('AC251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                SP_DELETE.Cell('AC250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                SP_DELETE.Cell('AC250').CellValue = 1;
                end
        end
    end
end

   if inipara == 5 && finpara == 2
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-5***********************%
    for K = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end   
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    %*********************************************************************%
    SP_DELETE.Cell('N16').CellValue = 1;
    SP_DELETE.Cell('N20').CellValue = 1;
    SP_DELETE.Cell('N24').CellValue = 1;
    SP_DELETE.Cell('H256').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    SP_DELETE.Cell('AC256').CellValue = 1;
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for AC = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
                sp_condiciones.Cell('F41').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                SP_DELETE.Cell('AC252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                SP_DELETE.Cell('AC252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                SP_DELETE.Cell('AC251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                SP_DELETE.Cell('AC251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                SP_DELETE.Cell('AC250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                SP_DELETE.Cell('AC250').CellValue = 1;
                end
        end
    end
   end   
   
if inipara == 5 && finpara == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-2***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-5***********************%
    for T = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end
    for Z = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for ZI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZI))).CellValue = 1;
    end
    for ZII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Z',num2str(ZII))).CellValue = 1;
    end         
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    %*********************************************************************%
    SP_DELETE.Cell('N12').CellValue = 1;
    SP_DELETE.Cell('N16').CellValue = 1;
    SP_DELETE.Cell('N20').CellValue = 1;
    SP_DELETE.Cell('N24').CellValue = 1;
    SP_DELETE.Cell('Q256').CellValue = 1;
    SP_DELETE.Cell('H256').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    SP_DELETE.Cell('AC256').CellValue = 1;
    for AC = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for Q = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for AC = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
                sp_condiciones.Cell('F41').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                SP_DELETE.Cell('AC252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                SP_DELETE.Cell('AC252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                SP_DELETE.Cell('AC251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                SP_DELETE.Cell('AC251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                for AC = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for AC = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
                end
                for ACI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
                end
                for ACII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('AC',num2str(ACII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                SP_DELETE.Cell('AC250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                SP_DELETE.Cell('AC250').CellValue = 1;
                end
        end
    end
end

if inipara == 4 && finpara == 3
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end
    for K = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for KI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KI))).CellValue = 1;
    end
    for KII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('K',num2str(KII))).CellValue = 1;
    end    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    sp_condiciones.Cell('C45').CellValue = 0;
    %*********************************************************************%
    for N = 21:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    SP_DELETE.Cell('N20').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
                sp_condiciones.Cell('F41').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;                
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end                
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                end
        end
    end
end

if inipara == 4 && finpara == 2
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************% 
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for K = 2:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end   
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    %*********************************************************************%
    for N = 20:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    SP_DELETE.Cell('N16').CellValue = 1;
    SP_DELETE.Cell('N12').CellValue = 1;
    SP_DELETE.Cell('Q256').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    SP_DELETE.Cell('H256').CellValue = 1;
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for Q = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end                
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                end
        end
    end
end

if inipara == 4 && finpara == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-2***********************% 
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************% 
    %**************ESTO ELIMINA TRENES PARALELO 2-4***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
     for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end   
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    %*********************************************************************%
    for N = 20:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    SP_DELETE.Cell('N12').CellValue = 1;
    SP_DELETE.Cell('N16').CellValue = 1;
    SP_DELETE.Cell('Q256').CellValue = 1;
    SP_DELETE.Cell('W256').CellValue = 1;
    SP_DELETE.Cell('H256').CellValue = 1;
    for Q = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for W = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for W = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
                sp_condiciones.Cell('F40').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                SP_DELETE.Cell('W252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                SP_DELETE.Cell('W252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                SP_DELETE.Cell('W251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                SP_DELETE.Cell('W251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                for W = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for W = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
                end
                for WI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
                end
                for WII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('W',num2str(WII))).CellValue = 1;
                end                
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                SP_DELETE.Cell('W250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                SP_DELETE.Cell('W250').CellValue = 1;
                end
        end
    end
end

if inipara == 3 && finpara == 2
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************% 
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end
    for T = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for TI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TI))).CellValue = 1;
    end
    for TII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('T',num2str(TII))).CellValue = 1;
    end  
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    sp_condiciones.Cell('C29').CellValue = 0;
    %*********************************************************************%
    for N = 16:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end   
    SP_DELETE.Cell('H256').CellValue = 1;
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
            case 3
               for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                end
            case 1
                 for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                end
        end
    end
end

if inipara == 3 && finpara == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-2***********************%
    %**************ESTO ELIMINA TRENES PARALELO 2-3***********************% 
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    %*********************************************************************%
    for N = 16:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end 
    SP_DELETE.Cell('N12').CellValue = 1;
    SP_DELETE.Cell('Q256').CellValue = 1;
    SP_DELETE.Cell('H256').CellValue = 1;
    for Q = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end
    for H = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for H = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
                sp_condiciones.Cell('F39').CellValue = 100;
            case 3
               for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end 
                for H = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                SP_DELETE.Cell('H252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                SP_DELETE.Cell('H252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                SP_DELETE.Cell('H251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                SP_DELETE.Cell('H251').CellValue = 1;
                end
            case 1
                 for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                for H = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for H = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
                end
                for HI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
                end
                for HII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('H',num2str(HII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                SP_DELETE.Cell('H250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                SP_DELETE.Cell('H250').CellValue = 1;
                end
        end
    end
end

if inipara == 2 && finpara == 1
    %**************ESTO ELIMINA TRENES PARALELO 2-2***********************%
    %**************ESTO ELIMINA TRENES PARALELO 1-3 y 2-3*****************% 
    %**************ESTO ELIMINA TRENES PARALELO 1-4 y 2-4*****************%
    %**************ESTO ELIMINA TRENES PARALELO 1-5 y 2-5*****************%    
    for H = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(H))).CellValue = 1;
    end
    for HI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('H',num2str(HI))).CellValue = 1;
    end
    for T = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('T',num2str(T))).CellValue = 1;
    end
    for K = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('K',num2str(K))).CellValue = 1;
    end
    for Z = 6:184
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('Z',num2str(Z))).CellValue = 1;
    end
    for AF = 2:180
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AF',num2str(AF))).CellValue = 1;
    end
    for W = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(W))).CellValue = 1;
    end
    for WI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('W',num2str(WI))).CellValue = 1;
    end
     for AC = 2:247
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(AC))).CellValue = 1;
    end
    for ACI = 253:256
        if booleano == false
            break;
        end
        pause (1E-08);
        SP_DELETE.Cell(strcat('AC',num2str(ACI))).CellValue = 1;
    end    
    %*********************************************************************%
    for E = 70:175
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(E))).CellValue = 1;
    end
    for EI = 64:66
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EI))).CellValue = 1;
    end
    for EII = 178:184
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('E',num2str(EII))).CellValue = 1;
    end    
    %*********************************************************************%
    sp_condiciones.Cell('C17').CellValue = 0;
    %*********************************************************************%
    for N = 12:24
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('N',num2str(N))).CellValue = 1;
    end
    SP_DELETE.Cell('Q256').CellValue = 1;    
    for Q = 176:178
        if booleano == false
            break;
        end
        pause (1E-08)
        SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
    end    
    if etafin == 1
        switch etaini
            case 4
               for B = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for Q = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                %*********************************************************%
                sp_condiciones.Cell('F37').CellValue = 100;
                sp_condiciones.Cell('F38').CellValue = 100;
            case 3
                for B = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 179:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 246:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 176:178
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B252').CellValue = 0;
                SP_DELETE.Cell('Q252').CellValue = 0;
                else
                SP_DELETE.Cell('B252').CellValue = 1;
                SP_DELETE.Cell('Q252').CellValue = 1;
                end
            case 2
                for B = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 119:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 244:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 116:118
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                %*********************************************************%
                if booleano == false
                SP_DELETE.Cell('B251').CellValue = 0;
                SP_DELETE.Cell('Q251').CellValue = 0;
                else
                SP_DELETE.Cell('B251').CellValue = 1;
                SP_DELETE.Cell('Q251').CellValue = 1;
                end
            case 1
                for B = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for B = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(B))).CellValue = 1;
                end
                for BI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BI))).CellValue = 1;
                end
                for BII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('B',num2str(BII))).CellValue = 1;
                end                
                for Q = 65:239
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for Q = 242:247
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(Q))).CellValue = 1;
                end
                for QI = 62:64
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QI))).CellValue = 1;
                end
                for QII = 253:255
                    if booleano == false
                        break;
                    end
                    pause (1E-08)
                    SP_DELETE.Cell(strcat('Q',num2str(QII))).CellValue = 1;
                end
                if booleano == false
                SP_DELETE.Cell('B250').CellValue = 0;
                SP_DELETE.Cell('Q250').CellValue = 0;
                else
                SP_DELETE.Cell('B250').CellValue = 1;
                SP_DELETE.Cell('Q250').CellValue = 1;
                end
        end
    end
end
end

    

%*************************************************************************%
%*****************************CONDICIONES*********************************%
%*************************************************************************%
%************************CONDICIONES DE SALIDA****************************%
%*************************************************************************%
if booleano == false
%     disp('Envío abortado');
    pause(1E-8);
else
if opc1 == 1
if para == 0 || para == 1
if se == 1
        switch etapasetren1
            case 4
                if presionsal == 0
                    sp_condiciones.Cell('B7').CellValue = 0;
                    sp_condiciones.Cell('C7').CellValue = 1;
                    exit_4_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B7').CellValue = 1;
                    sp_condiciones.Cell('C7').CellValue = 0;
                    exit_4_0.Pressure.SetValue(presionsal,['kPa']);
                end
            case 3
                if presionsal == 0
                    sp_condiciones.Cell('B9').CellValue = 0;
                    sp_condiciones.Cell('C9').CellValue = 1;
                    exit_104_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B9').CellValue = 1;
                    sp_condiciones.Cell('C9').CellValue = 0;
                    exit_104_1_1.Pressure.SetValue(presionsal,['kPa']);
                end
            case 2
                if presionsal == 0
                    sp_condiciones.Cell('B11').CellValue = 0;
                    sp_condiciones.Cell('C11').CellValue = 1;
                    exit_69_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B11').CellValue = 1;
                    sp_condiciones.Cell('C11').CellValue = 0;
                    exit_69_1_1.Pressure.SetValue(presionsal,['kPa']);
                end
            case 1
                if presionsal == 0
                    sp_condiciones.Cell('B13').CellValue = 0;
                    sp_condiciones.Cell('C13').CellValue = 1;
                    exit_38_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B13').CellValue = 1;
                    sp_condiciones.Cell('C13').CellValue = 0;
                    exit_38_1_1.Pressure.SetValue(presionsal,['kPa']);
                end
        end
end
end

if se == 2
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                    end 
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
    if etapasetren1 == 3
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                    end 
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
    if etapasetren1 == 2
            switch etapasetren2
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
    if etapasetren1 == 1
            switch etapasetren2
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
end
end
end

%*************************************************************************%
%************************CONDICIONES DE SALIDA****************************%
%*****************************EN PARALELO*********************************%
%*************************************************************************%
if opc1 == 1
if se == 1
    if para == 2
        switch etapasetren1
            case 4
                if presionsal == 0
                    sp_condiciones.Cell('B7').CellValue = 0;
                    sp_condiciones.Cell('C7').CellValue = 1;
                    exit_4_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B19').CellValue = 0;
                    sp_condiciones.Cell('C19').CellValue = 1;
                    exit_4_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B7').CellValue = 1;
                    sp_condiciones.Cell('C7').CellValue = 0;
                    exit_4_0.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B19').CellValue = 1;
                    sp_condiciones.Cell('C19').CellValue = 0;
                    exit_4_0_2.Pressure.SetValue(presionsal,['kPa']);
                end
            case 3
                if presionsal == 0
                    sp_condiciones.Cell('B9').CellValue = 0;
                    sp_condiciones.Cell('C9').CellValue = 1;
                    exit_104_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B21').CellValue = 0;
                    sp_condiciones.Cell('C21').CellValue = 1;
                    exit_104_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B9').CellValue = 1;
                    sp_condiciones.Cell('C9').CellValue = 0;
                    exit_104_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B21').CellValue = 1;
                    sp_condiciones.Cell('C21').CellValue = 0;
                    exit_104_1_2.Pressure.SetValue(presionsal,['kPa']);
                end
            case 2
                if presionsal == 0
                    sp_condiciones.Cell('B11').CellValue = 0;
                    sp_condiciones.Cell('C11').CellValue = 1;
                    exit_69_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B23').CellValue = 0;
                    sp_condiciones.Cell('C23').CellValue = 1;
                    exit_69_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B11').CellValue = 1;
                    sp_condiciones.Cell('C11').CellValue = 0;
                    exit_69_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B23').CellValue = 1;
                    sp_condiciones.Cell('C23').CellValue = 0;
                    exit_69_1_2.Pressure.SetValue(presionsal,['kPa']);
                end
            case 1
                if presionsal == 0
                    sp_condiciones.Cell('B13').CellValue = 0;
                    sp_condiciones.Cell('C13').CellValue = 1;
                    exit_38_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B25').CellValue = 0;
                    sp_condiciones.Cell('C25').CellValue = 1;
                    exit_38_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B13').CellValue = 1;
                    sp_condiciones.Cell('C13').CellValue = 0;
                    exit_38_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B25').CellValue = 1;
                    sp_condiciones.Cell('C25').CellValue = 0;
                    exit_38_1_2.Pressure.SetValue(presionsal,['kPa']);
                end
        end
    end
end

if se == 2
    if para == 2
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                    end 
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
    if etapasetren1 == 3
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                    end 
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
    if etapasetren1 == 2
            switch etapasetren2
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
    if etapasetren1 == 1
            switch etapasetren2
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
    end
    end
end

if para == 3
    if se == 1
        switch etapasetren1
            case 4
                if presionsal == 0
                    sp_condiciones.Cell('B7').CellValue = 0;
                    sp_condiciones.Cell('C7').CellValue = 1;
                    exit_4_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B19').CellValue = 0;
                    sp_condiciones.Cell('C19').CellValue = 1;
                    exit_4_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B46').CellValue = 0;
                    sp_condiciones.Cell('C46').CellValue = 1;
                    exit_4_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B7').CellValue = 1;
                    sp_condiciones.Cell('C7').CellValue = 0;
                    exit_4_0.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B19').CellValue = 1;
                    sp_condiciones.Cell('C19').CellValue = 0;
                    exit_4_0_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B46').CellValue = 1;
                    sp_condiciones.Cell('C46').CellValue = 0;
                    exit_4_0_3.Pressure.SetValue(presionsal,['kPa']);
                end
            case 3
                if presionsal == 0
                    sp_condiciones.Cell('B9').CellValue = 0;
                    sp_condiciones.Cell('C9').CellValue = 1;
                    exit_104_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B21').CellValue = 0;
                    sp_condiciones.Cell('C21').CellValue = 1;
                    exit_104_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B47').CellValue = 0;
                    sp_condiciones.Cell('C47').CellValue = 1;
                    exit_104_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B9').CellValue = 1;
                    sp_condiciones.Cell('C9').CellValue = 0;
                    exit_104_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B21').CellValue = 1;
                    sp_condiciones.Cell('C21').CellValue = 0;
                    exit_104_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B47').CellValue = 1;
                    sp_condiciones.Cell('C47').CellValue = 0;
                    exit_104_1_3.Pressure.SetValue(presionsal,['kPa']);
                end
            case 2
                if presionsal == 0
                    sp_condiciones.Cell('B11').CellValue = 0;
                    sp_condiciones.Cell('C11').CellValue = 1;
                    exit_69_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B23').CellValue = 0;
                    sp_condiciones.Cell('C23').CellValue = 1;
                    exit_69_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B48').CellValue = 0;
                    sp_condiciones.Cell('C48').CellValue = 1;
                    exit_69_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B11').CellValue = 1;
                    sp_condiciones.Cell('C11').CellValue = 0;
                    exit_69_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B23').CellValue = 1;
                    sp_condiciones.Cell('C23').CellValue = 0;
                    exit_69_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B48').CellValue = 1;
                    sp_condiciones.Cell('C48').CellValue = 0;
                    exit_69_1_3.Pressure.SetValue(presionsal,['kPa']);
                end
            case 1
                if presionsal == 0
                    sp_condiciones.Cell('B13').CellValue = 0;
                    sp_condiciones.Cell('C13').CellValue = 1;
                    exit_38_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B25').CellValue = 0;
                    sp_condiciones.Cell('C25').CellValue = 1;
                    exit_38_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B49').CellValue = 0;
                    sp_condiciones.Cell('C49').CellValue = 1;
                    exit_38_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B13').CellValue = 1;
                    sp_condiciones.Cell('C13').CellValue = 0;
                    exit_38_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B25').CellValue = 1;
                    sp_condiciones.Cell('C25').CellValue = 0;
                    exit_38_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B49').CellValue = 1;
                    sp_condiciones.Cell('C49').CellValue = 0;
                    exit_38_1_3.Pressure.SetValue(presionsal,['kPa']);
                end
        end
    end
    
    if se == 2
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 3
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 2
            switch etapasetren2
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 1
            switch etapasetren2
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                    end  
            end
        end
    end
end

if para == 4
    if se == 1
        switch etapasetren1
            case 4
               if presionsal == 0
                    sp_condiciones.Cell('B7').CellValue = 0;
                    sp_condiciones.Cell('C7').CellValue = 1;
                    exit_4_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B19').CellValue = 0;
                    sp_condiciones.Cell('C19').CellValue = 1;
                    exit_4_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B46').CellValue = 0;
                    sp_condiciones.Cell('C46').CellValue = 1;
                    exit_4_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B52').CellValue = 0;
                    sp_condiciones.Cell('C52').CellValue = 1;
                    exit_4_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B7').CellValue = 1;
                    sp_condiciones.Cell('C7').CellValue = 0;
                    exit_4_0.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B19').CellValue = 1;
                    sp_condiciones.Cell('C19').CellValue = 0;
                    exit_4_0_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B46').CellValue = 1;
                    sp_condiciones.Cell('C46').CellValue = 0;
                    exit_4_0_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B52').CellValue = 1;
                    sp_condiciones.Cell('C52').CellValue = 0;
                    exit_4_0_4.Pressure.SetValue(presionsal,['kPa']);
               end
            case 3
                if presionsal == 0
                    sp_condiciones.Cell('B9').CellValue = 0;
                    sp_condiciones.Cell('C9').CellValue = 1;
                    exit_104_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B21').CellValue = 0;
                    sp_condiciones.Cell('C21').CellValue = 1;
                    exit_104_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B47').CellValue = 0;
                    sp_condiciones.Cell('C47').CellValue = 1;
                    exit_104_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B54').CellValue = 0;
                    sp_condiciones.Cell('C54').CellValue = 1;
                    exit_104_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B9').CellValue = 1;
                    sp_condiciones.Cell('C9').CellValue = 0;
                    exit_104_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B21').CellValue = 1;
                    sp_condiciones.Cell('C21').CellValue = 0;
                    exit_104_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B47').CellValue = 1;
                    sp_condiciones.Cell('C47').CellValue = 0;
                    exit_104_1_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B54').CellValue = 1;
                    sp_condiciones.Cell('C54').CellValue = 0;
                    exit_104_1_4.Pressure.SetValue(presionsal,['kPa']);
                end
            case 2
                if presionsal == 0
                    sp_condiciones.Cell('B11').CellValue = 0;
                    sp_condiciones.Cell('C11').CellValue = 1;
                    exit_69_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B23').CellValue = 0;
                    sp_condiciones.Cell('C23').CellValue = 1;
                    exit_69_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B48').CellValue = 0;
                    sp_condiciones.Cell('C48').CellValue = 1;
                    exit_69_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B56').CellValue = 0;
                    sp_condiciones.Cell('C56').CellValue = 1;
                    exit_69_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B11').CellValue = 1;
                    sp_condiciones.Cell('C11').CellValue = 0;
                    exit_69_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B23').CellValue = 1;
                    sp_condiciones.Cell('C23').CellValue = 0;
                    exit_69_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B56').CellValue = 1;
                    sp_condiciones.Cell('C56').CellValue = 0;
                    exit_69_1_4.Pressure.SetValue(presionsal,['kPa']);
                end
            case 1
                if presionsal == 0
                    sp_condiciones.Cell('B13').CellValue = 0;
                    sp_condiciones.Cell('C13').CellValue = 1;
                    exit_38_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B25').CellValue = 0;
                    sp_condiciones.Cell('C25').CellValue = 1;
                    exit_38_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B49').CellValue = 0;
                    sp_condiciones.Cell('C49').CellValue = 1;
                    exit_38_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B58').CellValue = 0;
                    sp_condiciones.Cell('C58').CellValue = 1;
                    exit_38_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B13').CellValue = 1;
                    sp_condiciones.Cell('C13').CellValue = 0;
                    exit_38_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B25').CellValue = 1;
                    sp_condiciones.Cell('C25').CellValue = 0;
                    exit_38_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B49').CellValue = 1;
                    sp_condiciones.Cell('C49').CellValue = 0;
                    exit_38_1_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B58').CellValue = 1;
                    sp_condiciones.Cell('C58').CellValue = 0;
                    exit_38_1_4.Pressure.SetValue(presionsal,['kPa']);
                end
        end
    end
    
    if se == 2
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B61').CellValue = 0;
                        sp_condiciones.Cell('C61').CellValue = 1;
                        exit_6_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B61').CellValue = 1;
                        sp_condiciones.Cell('C61').CellValue = 0;
                        exit_6_0_4.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 3
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B61').CellValue = 0;
                        sp_condiciones.Cell('C61').CellValue = 1;
                        exit_6_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B61').CellValue = 1;
                        sp_condiciones.Cell('C61').CellValue = 0;
                        exit_6_0_4.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 2
            switch etapasetren2
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 1
            switch etapasetren2
                case 1
                   if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                   end
            end
        end
    end
end

if para == 5
    if se == 1
        switch etapasetren1
            case 4
                if presionsal == 0
                    sp_condiciones.Cell('B7').CellValue = 0;
                    sp_condiciones.Cell('C7').CellValue = 1;
                    exit_4_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B19').CellValue = 0;
                    sp_condiciones.Cell('C19').CellValue = 1;
                    exit_4_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B46').CellValue = 0;
                    sp_condiciones.Cell('C46').CellValue = 1;
                    exit_4_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B52').CellValue = 0;
                    sp_condiciones.Cell('C52').CellValue = 1;
                    exit_4_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B68').CellValue = 0;
                    sp_condiciones.Cell('C68').CellValue = 1;
                    exit_4_0_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B7').CellValue = 1;
                    sp_condiciones.Cell('C7').CellValue = 0;
                    exit_4_0.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B19').CellValue = 1;
                    sp_condiciones.Cell('C19').CellValue = 0;
                    exit_4_0_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B46').CellValue = 1;
                    sp_condiciones.Cell('C46').CellValue = 0;
                    exit_4_0_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B52').CellValue = 1;
                    sp_condiciones.Cell('C52').CellValue = 0;
                    exit_4_0_4.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B68').CellValue = 1;
                    sp_condiciones.Cell('C68').CellValue = 0;
                    exit_4_0_5.Pressure.SetValue(presionsal,['kPa']);
                end
            case 3
                if presionsal == 0
                    sp_condiciones.Cell('B9').CellValue = 0;
                    sp_condiciones.Cell('C9').CellValue = 1;
                    exit_104_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B21').CellValue = 0;
                    sp_condiciones.Cell('C21').CellValue = 1;
                    exit_104_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B47').CellValue = 0;
                    sp_condiciones.Cell('C47').CellValue = 1;
                    exit_104_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B54').CellValue = 0;
                    sp_condiciones.Cell('C54').CellValue = 1;
                    exit_104_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B70').CellValue = 0;
                    sp_condiciones.Cell('C70').CellValue = 1;
                    exit_104_1_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B9').CellValue = 1;
                    sp_condiciones.Cell('C9').CellValue = 0;
                    exit_104_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B21').CellValue = 1;
                    sp_condiciones.Cell('C21').CellValue = 0;
                    exit_104_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B47').CellValue = 1;
                    sp_condiciones.Cell('C47').CellValue = 0;
                    exit_104_1_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B54').CellValue = 1;
                    sp_condiciones.Cell('C54').CellValue = 0;
                    exit_104_1_4.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B70').CellValue = 1;
                    sp_condiciones.Cell('C70').CellValue = 0;
                    exit_104_1_5.Pressure.SetValue(presionsal,['kPa']);
                end
            case 2
                if presionsal == 0
                    sp_condiciones.Cell('B11').CellValue = 0;
                    sp_condiciones.Cell('C11').CellValue = 1;
                    exit_69_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B23').CellValue = 0;
                    sp_condiciones.Cell('C23').CellValue = 1;
                    exit_69_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B48').CellValue = 0;
                    sp_condiciones.Cell('C48').CellValue = 1;
                    exit_69_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B56').CellValue = 0;
                    sp_condiciones.Cell('C56').CellValue = 1;
                    exit_69_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B72').CellValue = 0;
                    sp_condiciones.Cell('C72').CellValue = 1;
                    exit_69_1_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B11').CellValue = 1;
                    sp_condiciones.Cell('C11').CellValue = 0;
                    exit_69_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B23').CellValue = 1;
                    sp_condiciones.Cell('C23').CellValue = 0;
                    exit_69_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B56').CellValue = 1;
                    sp_condiciones.Cell('C56').CellValue = 0;
                    exit_69_1_4.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B72').CellValue = 1;
                    sp_condiciones.Cell('C72').CellValue = 0;
                    exit_69_1_5.Pressure.SetValue(presionsal,['kPa']);
                end
            case 1
                if presionsal == 0
                    sp_condiciones.Cell('B13').CellValue = 0;
                    sp_condiciones.Cell('C13').CellValue = 1;
                    exit_38_1_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B25').CellValue = 0;
                    sp_condiciones.Cell('C25').CellValue = 1;
                    exit_38_1_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B49').CellValue = 0;
                    sp_condiciones.Cell('C49').CellValue = 1;
                    exit_38_1_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B58').CellValue = 0;
                    sp_condiciones.Cell('C58').CellValue = 1;
                    exit_38_1_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    sp_condiciones.Cell('B74').CellValue = 0;
                    sp_condiciones.Cell('C74').CellValue = 1;
                    exit_38_1_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                else
                    sp_condiciones.Cell('B13').CellValue = 1;
                    sp_condiciones.Cell('C13').CellValue = 0;
                    exit_38_1_1.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B25').CellValue = 1;
                    sp_condiciones.Cell('C25').CellValue = 0;
                    exit_38_1_2.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B49').CellValue = 1;
                    sp_condiciones.Cell('C49').CellValue = 0;
                    exit_38_1_3.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B58').CellValue = 1;
                    sp_condiciones.Cell('C58').CellValue = 0;
                    exit_38_1_4.Pressure.SetValue(presionsal,['kPa']);
                    sp_condiciones.Cell('B74').CellValue = 1;
                    sp_condiciones.Cell('C74').CellValue = 0;
                    exit_38_1_5.Pressure.SetValue(presionsal,['kPa']);
                end
        end
    end
    if se == 2
        if etapasetren1 == 4
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B61').CellValue = 0;
                        sp_condiciones.Cell('C61').CellValue = 1;
                        exit_6_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B77').CellValue = 0;
                        sp_condiciones.Cell('C77').CellValue = 1;
                        exit_6_0_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B61').CellValue = 1;
                        sp_condiciones.Cell('C61').CellValue = 0;
                        exit_6_0_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B77').CellValue = 1;
                        sp_condiciones.Cell('C77').CellValue = 0;
                        exit_6_0_5.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B78').CellValue = 0;
                        sp_condiciones.Cell('C78').CellValue = 1;
                        exit_68_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B78').CellValue = 1;
                        sp_condiciones.Cell('C78').CellValue = 0;
                        exit_68_2_5.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B79').CellValue = 0;
                        sp_condiciones.Cell('C79').CellValue = 1;
                        exit_37_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B79').CellValue = 1;
                        sp_condiciones.Cell('C79').CellValue = 0;
                        exit_37_2_5.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
        if etapasetren1 == 3
            switch etapasetren2
                case 3
                    if presionsal == 0
                        sp_condiciones.Cell('B31').CellValue = 0;
                        sp_condiciones.Cell('C31').CellValue = 1;
                        exit_6_0.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B32').CellValue = 0;
                        sp_condiciones.Cell('C32').CellValue = 1;
                        exit_6_0_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B37').CellValue = 0;
                        sp_condiciones.Cell('C37').CellValue = 1;
                        exit_6_0_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B61').CellValue = 0;
                        sp_condiciones.Cell('C61').CellValue = 1;
                        exit_6_0_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B77').CellValue = 0;
                        sp_condiciones.Cell('C77').CellValue = 1;
                        exit_6_0_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B31').CellValue = 1;
                        sp_condiciones.Cell('C31').CellValue = 0;
                        exit_6_0.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B32').CellValue = 1;
                        sp_condiciones.Cell('C32').CellValue = 0;
                        exit_6_0_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B37').CellValue = 1;
                        sp_condiciones.Cell('C37').CellValue = 0;
                        exit_6_0_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B61').CellValue = 1;
                        sp_condiciones.Cell('C61').CellValue = 0;
                        exit_6_0_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B77').CellValue = 1;
                        sp_condiciones.Cell('C77').CellValue = 0;
                        exit_6_0_5.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B78').CellValue = 0;
                        sp_condiciones.Cell('C78').CellValue = 1;
                        exit_68_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B78').CellValue = 1;
                        sp_condiciones.Cell('C78').CellValue = 0;
                        exit_68_2_5.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                   if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B79').CellValue = 0;
                        sp_condiciones.Cell('C79').CellValue = 1;
                        exit_37_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B79').CellValue = 1;
                        sp_condiciones.Cell('C79').CellValue = 0;
                        exit_37_2_5.Pressure.SetValue(presionsal,['kPa']);
                   end 
            end
        end
        if etapasetren1 == 2
            switch etapasetren2
                case 2
                    if presionsal == 0
                        sp_condiciones.Cell('B33').CellValue = 0;
                        sp_condiciones.Cell('C33').CellValue = 1;
                        exit_68_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B34').CellValue = 0;
                        sp_condiciones.Cell('C34').CellValue = 1;
                        exit_68_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B38').CellValue = 0;
                        sp_condiciones.Cell('C38').CellValue = 1;
                        exit_68_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B62').CellValue = 0;
                        sp_condiciones.Cell('C62').CellValue = 1;
                        exit_68_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B78').CellValue = 0;
                        sp_condiciones.Cell('C78').CellValue = 1;
                        exit_68_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B33').CellValue = 1;
                        sp_condiciones.Cell('C33').CellValue = 0;
                        exit_68_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B34').CellValue = 1;
                        sp_condiciones.Cell('C34').CellValue = 0;
                        exit_68_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B38').CellValue = 1;
                        sp_condiciones.Cell('C38').CellValue = 0;
                        exit_68_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B62').CellValue = 1;
                        sp_condiciones.Cell('C62').CellValue = 0;
                        exit_68_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B78').CellValue = 1;
                        sp_condiciones.Cell('C78').CellValue = 0;
                        exit_68_2_5.Pressure.SetValue(presionsal,['kPa']);
                    end
                case 1
                    if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                    end
            end
        end
       if etapasetren1 == 1
           switch etapasetren2
               case 1
                 if presionsal == 0
                        sp_condiciones.Cell('B35').CellValue = 0;
                        sp_condiciones.Cell('C35').CellValue = 1;
                        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B36').CellValue = 0;
                        sp_condiciones.Cell('C36').CellValue = 1;
                        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B39').CellValue = 0;
                        sp_condiciones.Cell('C39').CellValue = 1;
                        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B64').CellValue = 0;
                        sp_condiciones.Cell('C64').CellValue = 1;
                        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
                        sp_condiciones.Cell('B79').CellValue = 0;
                        sp_condiciones.Cell('C79').CellValue = 1;
                        exit_37_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
                    else
                        sp_condiciones.Cell('B35').CellValue = 1;
                        sp_condiciones.Cell('C35').CellValue = 0;
                        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B36').CellValue = 1;
                        sp_condiciones.Cell('C36').CellValue = 0;
                        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B39').CellValue = 1;
                        sp_condiciones.Cell('C39').CellValue = 0;
                        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B64').CellValue = 1;
                        sp_condiciones.Cell('C64').CellValue = 0;
                        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
                        sp_condiciones.Cell('B79').CellValue = 1;
                        sp_condiciones.Cell('C79').CellValue = 0;
                        exit_37_2_5.Pressure.SetValue(presionsal,['kPa']);
                 end  
           end
       end
    end    
end
end


%*************************************************************************%
%****************************AMPLIACION TRENES****************************%
%*************************************************************************%
if booleano == false
%     disp('Envío abortado')
pause(1E-8);
else
if opc2 == 1
if inipara == 1 && finpara == 2
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
    end
end
if inipara == 1 || inipara == 2
    if finpara == 3
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B39').CellValue = 0;
        sp_condiciones.Cell('C39').CellValue = 1;
        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B39').CellValue = 1;
        sp_condiciones.Cell('C39').CellValue = 0;
        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
    end
    end
end
if inipara == 1 || inipara == 2 || inipara == 3
    if finpara == 4
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B39').CellValue = 0;
        sp_condiciones.Cell('C39').CellValue = 1;
        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B64').CellValue = 0;
        sp_condiciones.Cell('C64').CellValue = 1;
        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B39').CellValue = 1;
        sp_condiciones.Cell('C39').CellValue = 0;
        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B64').CellValue = 1;
        sp_condiciones.Cell('C64').CellValue = 0;
        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
    end
    end
end
if inipara == 1 || inipara == 2 || inipara == 3 || inipara == 4 
    if finpara == 5
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B39').CellValue = 0;
        sp_condiciones.Cell('C39').CellValue = 1;
        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B64').CellValue = 0;
        sp_condiciones.Cell('C64').CellValue = 1;
        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B79').CellValue = 0;
        sp_condiciones.Cell('C79').CellValue = 1;
        exit_37_2_5.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B39').CellValue = 1;
        sp_condiciones.Cell('C39').CellValue = 0;
        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B64').CellValue = 1;
        sp_condiciones.Cell('C64').CellValue = 0;
        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B79').CellValue = 1;
        sp_condiciones.Cell('C79').CellValue = 0;
        exit_37_2_5.Pressure.SetValue(presionsal,['kPa']);
    end
    end  
end

%*************************************************************************%
%****************************REDUCCION TRENES*****************************%
%*************************************************************************%
if inipara == 5 && finpara == 4
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B39').CellValue = 0;
        sp_condiciones.Cell('C39').CellValue = 1;
        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B64').CellValue = 0;
        sp_condiciones.Cell('C64').CellValue = 1;
        exit_37_2_4.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B39').CellValue = 1;
        sp_condiciones.Cell('C39').CellValue = 0;
        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B64').CellValue = 1;
        sp_condiciones.Cell('C64').CellValue = 0;
        exit_37_2_4.Pressure.SetValue(presionsal,['kPa']);
    end
end
if inipara == 5 || inipara == 4
    if finpara == 3
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B39').CellValue = 0;
        sp_condiciones.Cell('C39').CellValue = 1;
        exit_37_2_3.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B39').CellValue = 1;
        sp_condiciones.Cell('C39').CellValue = 0;
        exit_37_2_3.Pressure.SetValue(presionsal,['kPa']);
    end
    end
end
if inipara == 5 || inipara == 4 || inipara == 3
    if finpara == 2
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
        sp_condiciones.Cell('B36').CellValue = 0;
        sp_condiciones.Cell('C36').CellValue = 1;
        exit_37_2_2.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
        sp_condiciones.Cell('B36').CellValue = 1;
        sp_condiciones.Cell('C36').CellValue = 0;
        exit_37_2_2.Pressure.SetValue(presionsal,['kPa']);
    end
    end
end
if inipara == 5 || inipara == 4 || inipara == 3 || inipara == 2 
    if finpara == 1
    if presionsal == 0
        sp_condiciones.Cell('B35').CellValue = 0;
        sp_condiciones.Cell('C35').CellValue = 1;
        exit_37_2_1.MolarFlow.SetValue(molflosal,['kgmole/h']);
    else
        sp_condiciones.Cell('B35').CellValue = 1;
        sp_condiciones.Cell('C35').CellValue = 0;
        exit_37_2_1.Pressure.SetValue(presionsal,['kPa']);
    end
    end
end
end
end
%*************************************************************************%
%***********************CONDICIONES DE ENTRADA****************************%
%*************************************************************************%
entrada.Temperature.SetValue(tempent,['C']);
sp_condiciones.Cell('H2').CellValue = molfloent;
% sp_condiciones.Cell('H2').CellValue = molfloent;
sp_condiciones.Cell('C1').CellValue = 1;
sp_condiciones.Cell('C2').CellValue = 1;
% entrada.MolarFlow.SetValue(molfloent,['kgmole/h']);

entrada.Pressure.SetValue(presionent,['kPa']);

if actpre == 1
sp_condiciones.Cell('C1').CellValue = 1;
sp_condiciones.Cell('C2').CellValue = 0;
end

if actmol == 1
sp_condiciones.Cell('C2').CellValue = 1;
sp_condiciones.Cell('C1').CellValue = 0;
end

%*************************************************************************%
%***************************FLUJO MOLAR***********************************%
%*************************************************************************%
sp_condiciones.Cell('F2').CellValue = comp1;
sp_condiciones.Cell('F3').CellValue = comp2;
sp_condiciones.Cell('F4').CellValue = comp3;
sp_condiciones.Cell('F5').CellValue = comp4;
sp_condiciones.Cell('F6').CellValue = comp5;
sp_condiciones.Cell('F7').CellValue = comp6;
sp_condiciones.Cell('F8').CellValue = comp7;
sp_condiciones.Cell('F9').CellValue = comp8;
sp_condiciones.Cell('F10').CellValue = comp9;
sp_condiciones.Cell('F11').CellValue = comp10;
sp_condiciones.Cell('F12').CellValue = comp11;
sp_condiciones.Cell('F13').CellValue = comp12;
sp_condiciones.Cell('F14').CellValue = comp13;
sp_condiciones.Cell('F15').CellValue = comp14;
sp_condiciones.Cell('F16').CellValue = comp15;
sp_condiciones.Cell('F17').CellValue = comp16;
sp_condiciones.Cell('F18').CellValue = comp17;

%*************************************************************************%
%********************DIMENSIONAMIENTO DE EQUIPOS**************************%
%*************************************************************************%
UnitOps.Cell('B2').CellValue = diametro;
UnitOps.Cell('B3').CellValue = longitud;

if booleano == true
pause(0.001);
set(handles.compl3,'string','Sending data complete');
warndlg('The gas compression plant is ready','NOTIFICATION','modal');
end
guidata(hObject, handles);

% --- Executes on button press in abortarparalelo.
function abortarparalelo_Callback(hObject, eventdata, handles)
set(handles.paralelo,'String','');
set(handles.text43,'String','');
set(handles.reduction,'String','');
set(handles.paralelo,'String','');
set(handles.sinreduccion,'Value',0);
set(handles.conreduccion,'Value',0);
set(handles.reduction,'visible','off');
set(handles.paralelo,'visible','off');
set(handles.text40,'visible','off');
set(handles.text48,'visible','off');
set(handles.text50,'visible','on');

set(handles.text52,'visible','off');
set(handles.text51,'String','');
set(handles.text53,'string','');

guidata(hObject, handles);


% --- Executes on button press in aborts.
function aborts_Callback(hObject, eventdata, handles)
global booleano opc1 opc2

booleano = true;
opc1 = 1;
opc2 = 0;
axes(handles.Config);
[x10,map10]=imread('SinReduccion.jpg');
image(x10);
colormap(map10);
axis off
set(handles.sinreduccion,'Value',1);
set(handles.serie,'String','');
set(handles.serie,'Enable','on')
set(handles.confirmarserie,'Enable','on');
set(handles.text2,'Visible','off')
set(handles.text3,'Visible','off')
set(handles.etapa,'Visible','off')
set(handles.confirmaretapa,'Visible','off')
set(handles.confirmaretapa,'Enable','on'); 
set(handles.etapa,'String','');
set(handles.etapa,'Enable','on')
set(handles.text101,'Visible','off');
set(handles.paralelo,'String','');
set(handles.paralelo,'Enable','on')
set(handles.paralelo,'Visible','off');
set(handles.confirmarparalelo,'Visible','off');
set(handles.confirmarparalelo,'Enable','on');
set(handles.text33,'String','');
set(handles.text34,'String','');
set(handles.text35,'String','');
set(handles.text43,'String','');
set(handles.text3,'String','0');
set(handles.Panel31,'visible','off');
set(handles.panelsr,'visible','on');

set(handles.conreduccion,'Value',0);
set(handles.iniparatrains,'String','');
set(handles.finparatrains,'Enable','on');
set(handles.text114,'Visible','off');
set(handles.finparatrains,'Visible','off');
set(handles.finparatrains,'String','');
set(handles.confirmarfintrain,'Visible','off');
set(handles.confirmarfintrain,'Enable','on');
set(handles.iniparatrains,'Enable','on')
set(handles.confirmarfintrain,'Enable','off')
set(handles.etapainiparatrain,'String','');
set(handles.etapainiparatrain,'Enable','on');
set(handles.etapainiparatrain,'Visible','off');
set(handles.confirmaretapini,'Enable','on');
set(handles.confirmarinitrain,'Enable','on')
set(handles.text120,'Visible','off');

set(handles.confirmaretapini,'visible','off');
set(handles.text214,'Visible','off');

set(handles.etapafinparatrain,'String','');
set(handles.etapafinparatrain,'Enable','on');
set(handles.etapafinparatrain,'Visible','off');
set(handles.confirmaretapfin,'visible','off');
set(handles.confirmaretapfin,'Enable','on');
set(handles.text134,'String','');
set(handles.text136,'String','');
set(handles.text145,'String','');
set(handles.text143,'String','');

set(handles.uipanel40,'visible','off');
set(handles.panelampred,'visible','off');
set(handles.panelampred,'visible','off');

set(handles.uipanel26,'title','Your choice');

set(handles.resusinreduc,'visible','off');
set(handles.text59,'string','');
set(handles.text61,'string','');
set(handles.text62,'string','');
set(handles.text152,'string','');
set(handles.resuconreduc,'visible','off');
set(handles.text155,'String','');
set(handles.text157,'string','');
set(handles.text173,'string','');
set(handles.text175,'string','');

set(handles.resutrasta,'visible','off');
set(handles.compl2,'visible','off');
set(handles.nocompl2,'visible','on');
set(handles.nocompl3,'visible','on');
set(handles.compl3,'visible','off');
clear inipara etaini finpara etafin se etapaserie se et a sumaes anterior...
      para opc1 opc2
warndlg('The trains and stages has been canceled','NOTIFICATION','modal');
guidata(hObject, handles);

% --- Executes on button press in abortarcondi.
function abortarcondi_Callback(hObject, eventdata, handles)
global actpre actmol  
clear presionent tempent molfloent presionsal molflosal...
      fraccion diametro longitud...
      actpre actmol R1 R2
set(handles.activepreent,'Value',0);
set(handles.edit10,'Enable','on');
set(handles.edit10,'String','');
set(handles.activemolarent,'Value',0);
set(handles.edit12,'Enable','on');
set(handles.edit11,'String','');
set(handles.edit12,'String','');
set(handles.edit13,'String','');
set(handles.edit16,'String','');
set(handles.presal,'Value',0);
set(handles.molarsal,'Value',0);
set(handles.edit13,'visible','off');
set(handles.edit16,'visible','off');
set(handles.uitable6,'Data',cell(17,1));
set(handles.DiametroSeparador,'string','');
set(handles.LongitudSeparador,'string','');
set(handles.resucondin,'visible','off');
set(handles.text184,'String','');
set(handles.text186,'string','');
set(handles.text200,'string','');
set(handles.resucondout,'visible','off');
set(handles.text190,'String','');
set(handles.text202,'String','');
set(handles.resusepar,'visible','off');
set(handles.text203,'String','');
set(handles.text206,'String','');
set(handles.compl1,'visible','off');
set(handles.nocompl1,'visible','on');
set(handles.nocompl3,'visible','on');
set(handles.compl3,'visible','off');
% actpre = 0; 
% actmol = 0;  
warndlg('The conditions have been canceled','WARNING','modal');
guidata(hObject, handles);


% --- Executes on button press in clean.
function clean_Callback(hObject, eventdata, handles)
global  opc1 opc2 booleano
opc1 = 1;
opc2 = 0;

set(handles.activepreent,'Value',0);
set(handles.activemolarent,'Value',0);
set(handles.edit10,'String','');
set(handles.edit10,'Enable','on');
set(handles.edit12,'String','');
set(handles.edit12,'Enable','on');
set(handles.edit11,'String','');
set(handles.uitable6,'Data',cell(17,1));
set(handles.presal,'Value',0);
set(handles.molarsal,'Value',0);
set(handles.edit13,'String','');
set(handles.edit13,'visible','off');
set(handles.edit16,'String','');
set(handles.edit16,'visible','off');
set(handles.text33,'String','');
set(handles.text34,'String','');
set(handles.text35,'String','');
set(handles.compl1,'visible','off');
set(handles.nocompl1,'visible','on');

set(handles.text43,'string','');

axes(handles.Config);
[x10,map10]=imread('SinReduccion.jpg');
image(x10);
colormap(map10);
axis off
set(handles.sinreduccion,'Value',1);
set(handles.panelsr,'visible','on');
set(handles.conreduccion,'Value',0);
set(handles.panelampred,'visible','off');
set(handles.serie,'String','');
set(handles.serie,'Enable','on')
set(handles.confirmarserie,'Enable','on');
set(handles.text2,'Visible','off')
set(handles.text3,'Visible','off')
set(handles.etapa,'Visible','off')
set(handles.confirmaretapa,'Visible','off')
set(handles.confirmaretapa,'Enable','on'); 
set(handles.text3,'String','0');
set(handles.etapa,'String','');
set(handles.etapa,'Enable','on')
set(handles.text101,'Visible','off');
set(handles.paralelo,'String','');
set(handles.paralelo,'Enable','on')
set(handles.paralelo,'Visible','off');
set(handles.confirmarparalelo,'Visible','off');
set(handles.confirmarparalelo,'Enable','on');
set(handles.iniparatrains,'string','');
set(handles.iniparatrains,'Enable','on');
set(handles.confirmarinitrain,'Enable','on');
set(handles.text114,'Visible','off');
set(handles.finparatrains,'Visible','off');
set(handles.finparatrains,'Enable','on');
set(handles.confirmarfintrain,'Enable','on');
set(handles.finparatrains,'string','');
set(handles.confirmarfintrain,'Visible','off');
set(handles.text120,'Visible','off');
set(handles.etapainiparatrain,'string','');
set(handles.etapainiparatrain,'Visible','off');
set(handles.etapainiparatrain,'Enable','on');
set(handles.confirmaretapini,'Enable','on');
set(handles.compl2,'visible','off');
set(handles.nocompl2,'visible','on');

set(handles.confirmaretapini,'Visible','off');
set(handles.text214,'Visible','off');

set(handles.etapafinparatrain,'string','');
set(handles.etapafinparatrain,'Visible','off');
set(handles.etapafinparatrain,'Enable','on');
set(handles.confirmaretapfin,'Visible','off');
set(handles.confirmaretapfin,'Enable','on');

set(handles.text33,'String','');
set(handles.text34,'String','');
set(handles.text35,'String','');
set(handles.text43,'String','');
set(handles.text3,'String','0');
set(handles.Panel31,'visible','off');
set(handles.text134,'String','');
set(handles.text145,'String','');

set(handles.text136,'String','');

set(handles.text143,'String','');

set(handles.text157,'String','');

set(handles.text175,'String','');

set(handles.uipanel40,'visible','off');
set(handles.uipanel26,'title','Your choice');
set(handles.aborts,'visible','on');

set(handles.DiametroSeparador,'String','');
set(handles.LongitudSeparador,'String','');

set(handles.file,'String','');
set(handles.resucondin,'visible','off');
set(handles.text184,'String','');
set(handles.text186,'String','');
set(handles.text200,'String','');
set(handles.resucondout,'visible','off');
set(handles.text190,'String','');
set(handles.text202,'String','');
set(handles.resusepar,'visible','off')
set(handles.text203,'string','')
set(handles.text206,'string','')
set(handles.resutrasta,'visible','off');
set(handles.resusinreduc,'visible','off');
set(handles.text59,'string','');
set(handles.text61,'string','');
set(handles.text62,'string','');
set(handles.text152,'string','');
set(handles.resuconreduc,'visible','off');
set(handles.text155,'String','');
set(handles.text157,'string','');
set(handles.text173,'string','');
set(handles.text175,'string','');

set(handles.resucondin,'visible','off');
set(handles.text184,'String','');
set(handles.text186,'string','');
set(handles.resucondout,'visible','off');
set(handles.text190,'String','');

set(handles.Panel1,'visible','on');
set(handles.Panel2,'visible','off');
set(handles.Panel10,'visible','off');

set(handles.nocompl3,'visible','on');
set(handles.compl3,'visible','off');

clear se para inipara finpara etafin etaini etapasetren1 etapasetren2...
      actpre actmol R1 R2
booleano = false;      
clc;
warndlg('The configuration has been aborted','NOTIFICATION','modal');
guidata(hObject, handles);

% --- Executes when user attempts to close figure1.
function figure1_CloseRequestFcn(hObject, eventdata, handles)
opc = questdlg('Do you want to exit the program?','EXIT','Yes','No','No');
if strcmp(opc,'No')
return
else
clc;delete(hObject);
end
