unit URDP;

interface

uses
  Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.OleCtrls, MSTSCLib_TLB, Vcl.ComCtrls, Vcl.ImgList,IdIcmpClient,StrUtils,Math,
  Vcl.Themes,Vcl.Styles,Winapi.Windows, JvNetscapeSplitter, System.Actions,Vcl.Samples.Spin,
  Inifiles,Vcl.ButtonGroup, Vcl.Menus,ActiveX,
  Registry, JvExExtCtrls, System.ImageList, Vcl.Mask;
/////////////////////////////////////////

 type /////описание для кнопки закрыть
  TTabControlStyleHookBtnClose = class(TTabControlStyleHook)
  private
    FHotIndex       : Integer;
    FWidthModified  : Boolean;
    procedure WMMouseMove(var Message: TMessage); message WM_MOUSEMOVE;
    procedure WMLButtonUp(var Message: TWMMouse); message WM_LBUTTONUP;
    function GetButtonCloseRect(Index: Integer):TRect;
  strict protected
    procedure DrawTab(Canvas: TCanvas; Index: Integer); override;
    procedure MouseEnter; override;
    procedure MouseLeave; override;
  public
    constructor Create(AControl: TWinControl); override;    
  end;
///////////////////////////////////////  
type
  TFormRDP = class(TForm)
    Panel1: TPanel;
    Memo1: TMemo;
    ComboKomp: TComboBoxEx;
    PageControl1: TPageControl;
    SpeedButton4: TButton;
    SpeedButton3: TButton;
    SpeedButton2: TButton;
    SpeedButton1: TButton;
    ImageButton: TImageList;
    Button1: TButton;
    Button2: TButton;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    CheckBoxRecAudio: TCheckBox;
    CheckBoxClipboard: TCheckBox;
    CheckBoxPrint: TCheckBox;
    CheckBoxDisk: TCheckBox;
    CheckBoxPort: TCheckBox;
    CheckBoxSmartCard: TCheckBox;
    Label2: TLabel;
    ComboAudioRedirectionMode: TComboBox;
    Label3: TLabel;
    ComboKey: TComboBox;
    CheckRedirectDevices: TCheckBox;
    CheckBoxCredSsp: TCheckBox;
    ComboBoxNetworkConnectionType: TComboBox;
    Label5: TLabel;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox9: TCheckBox;
    CheckBox10: TCheckBox;
    ComboBoxAuthLevel: TComboBox;
    Label4: TLabel;
    Label12: TLabel;
    SpinMaxReconnectAttempts: TSpinEdit;
    CheckEnableAutoReconnect: TCheckBox;
    Button3: TButton;
    ComboBoxColorDepth: TComboBox;
    Label13: TLabel;
    CheckBoxBitmapPeristence: TCheckBox;
    CheckBoxCachePersistenceActive: TCheckBox;
    GroupBox2: TGroupBox;
    CheckBox1: TCheckBox;
    GroupBox3: TGroupBox;
    ImageList1: TImageList;
    CheckBoxSmartSizing: TCheckBox;
    PageControl2: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Label9: TLabel;
    Label8: TLabel;
    SpinBitmapVirtualCacheSize: TSpinEdit;
    SpinEditCache32BppSize: TSpinEdit;
    SpinEditCache16BppSize: TSpinEdit;
    Label7: TLabel;
    Label6: TLabel;
    SpinBitmapCacheSize: TSpinEdit;
    CheckBoxEnSuperPan: TCheckBox;
    CheckBoxMouseMode: TCheckBox;
    CheckBoxConToAdmSrv: TCheckBox;
    CheckBoxDisableCtrlAltDel: TCheckBox;
    CheckBoxDoubleClickDetect: TCheckBox;
    CheckBoxEnableWindowsKey: TCheckBox;
    Button7: TButton;
    Button6: TButton;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    CheckBoxGrabFocusOnConnect: TCheckBox;
    LabeledEditRdpPort: TLabeledEdit;
    SpinOverallConnectionTimeout: TSpinEdit;
    SpinMinutesToIdleTimeout: TSpinEdit;
    Label10: TLabel;
    Label11: TLabel;
    GroupBox1: TGroupBox;
    GroupBox7: TGroupBox;
    Button4: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    JvNetscapeSplitter2: TJvNetscapeSplitter;
    N5: TMenuItem;
    GroupBox8: TGroupBox;
    BigImage1: TImageList;
    N6: TMenuItem;
    PopupLog: TPopupMenu;
    N7: TMenuItem;
    PopupOtherWin: TPopupMenu;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    PopupOsherWinPanel: TPopupMenu;
    N12: TMenuItem;
    LabeledEdit3: TEdit;
    LabeledEdit1: TEdit;
    LabeledEdit2: TEdit;
    BigImage: TImageList;
    procedure SpeedButton1Click(Sender: TObject);
    procedure NewRdpAuthenticationWarningDismissed(Sender: TObject);
    procedure NewRdpAuthenticationWarningDisplayed(Sender: TObject);
    procedure NewRdpAutoReconnected(Sender: TObject);
    procedure NewRdpAutoReconnecting(ASender: TObject; disconnectReason,
      attemptCount: Integer);
    procedure NewRdpAutoReconnecting2(ASender: TObject;
      disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
      maxAttemptCount: Integer);
    procedure NewRdpChannelReceivedData(ASender: TObject; const chanName,
      data: WideString);
    procedure NewRdpConfirmClose(Sender: TObject);
    procedure NewRdpConnected(Sender: TObject);
    procedure NewRdpConnecting(Sender: TObject);
    procedure NewRdpDisconnected(ASender: TObject; discReason: Integer);
    procedure NewRdpExit(Sender: TObject);
    procedure NewRdpFatalError(ASender: TObject; errorCode: Integer);
    procedure NewRdpLoginComplete(Sender: TObject);
    procedure NewRdpLogonError(ASender: TObject; lError: Integer);
    procedure NewRdpNetworkStatusChanged(ASender: TObject;
      qualityLevel: Cardinal; bandwidth, rtt: Integer);
    procedure NewRdpWarning(ASender: TObject; warningCode: Integer);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure TabSheet1Resize(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure ComboKompKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SpeedButton1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SpeedButton4MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure MsRdpClient91DevicesButtonPressed(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Button2Click(Sender: TObject);
    procedure ComboBoxNetworkConnectionTypeChange(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ButtonUpdateClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure ComboKompSelect(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    function createRDP(NamePC,Domain,UserName,Passwd:string;AutoConnect:boolean;
      ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
      VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
      DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
      OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
      MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
      SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
      RedirectDrives,RedirectPrinters,RelativeMouseMode
      ,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
      AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):bool;
    procedure PageControl1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure PageControl1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure PageControl1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PageControl1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CheckBoxSmartSizingClick(Sender: TObject);
    procedure MsRdpClient91AuthenticationWarningDismissed(Sender: TObject);
    procedure NewAuthenticationWarningDisplayed(Sender: TObject);
    procedure MsRdpClient91AutoReconnected(Sender: TObject);
    procedure MsRdpClient91AutoReconnecting(ASender: TObject; disconnectReason,
      attemptCount: Integer);
    procedure MsRdpClient91AutoReconnecting2(ASender: TObject;
      disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
      maxAttemptCount: Integer);


   procedure MsRdpClient91ServiceMessageReceived(ASender: TObject;
      const serviceMessage: WideString);
    procedure Button4Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure OtherWinForRDPClientClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonOnForForm(Sender: TObject);
    procedure ButtonOffForForm(Sender: TObject);
    procedure ButtonResizeForForm(Sender: TObject);
    procedure ButtonCloseForForm(Sender: TObject);
    procedure PanelForButMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
    procedure PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
    procedure PanelForButMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
    procedure ButonMoveWinMouseDown(Sender: TObject;
                           Button: TMouseButton;
                            Shift: TShiftState;
                             X, Y: Integer);
    procedure ButtonMinimizeForForm(Sender: TObject);
    procedure ButtonMaxNormalForForm(Sender: TObject);
    procedure NewRDPFormResize(Sender: TObject);
    //procedure WMSysCommand(var Msg: TWMSysCommand);  message WM_SYSCOMMAND;
    procedure NewRDPFormShow(Sender: TObject);
    procedure ShowButtonOnClick(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure PopupLogPopup(Sender: TObject);
    procedure ShowBMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
    procedure ShowBMouseLeave(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure PopupOsherWinPanelPopup(Sender: TObject);
   private


    procedure WMCopyData(var MessageData: TWMCopyData); message WM_COPYDATA;
    Function RequestStr(nameStr:string;Receiv:integer):bool;
    function ping(s:string):boolean;
    function sizeRegion(i,z:integer):bool;
    function loadlistpc(sFile:string):Tstringlist; // загрузка сохраненного списка комьютеров

    function FlagsChecked (x:bool):bool;
    function FlagsEnabled(x:bool):bool;
    function readSettingsInDB(s:string):bool; /// чтение настроек из бд
    //function writesettingToDB(s:string):bool; // запись настроек в бд
    function ConnectToCurentSetting(NamePC:string;IndexTab:integer):boolean; // подклбчение текущего уже созданного RDP клиента с указанными настройками
    function createnewconnectInTab(namepc:string):bool;
  public
    function  CalcPerformanceFlags(s:string):integer;
    function  OtherWinForRDPClient(indTab:integer):bool;
    procedure findComponent;



  end;



var
  FormRDP: TFormRDP;
  //MyRDPWin: array [0..1000] of  TMsRdpClient9;
  //Ntab:     array [0..1000] of  TTabSheet;//TCloseTabSheet;//TTabSheet;

  HM,HML: THandle; //// мьютекс для копии приложения второй для поиска родительского приложения
  RootPatch: HKEY;
  TabShetCount,CountNewForm:integer;// количество вкладок и форм в динамике
  CountTab:integer;
  PMouse:Tpoint;
  movingPan:boolean;
  x0,y0:integer;
  hwn: THandle;
 const
 stringpatch:string ='software\MRPC\MRPC';

implementation

  uses MyDMRDP,LoadListPC;

{$R *.dfm}
{procedure TFormRDP.WMSysCommand(var Msg: TWMSysCommand);
begin
Memo1.Lines.Add('WMSysCommand');
  if (fsCreating in FormState) then   //fsCreating, fsVisible, fsShowing, fsModal,    fsCreatedMDIChild, fsActivated
  begin
   case Msg.CmdType {and $FFF0}{ of
    SC_MINIMIZE:
      begin
      Memo1.Lines.Add('Свернули');
      BorderStyle:=bsSingle;
      end;
    SC_RESTORE:
      begin
      Memo1.Lines.Add('Развернули');
      BorderStyle:=bsNone;
      end;
     end;
  end;
  inherited;
end; }

function TFormRDP.FlagsChecked (x:bool):bool;
begin
CheckBox1.checked:=x;
CheckBox2.checked:=x;
CheckBox3.checked:=x;
CheckBox5.checked:=x;
CheckBox6.checked:=x;
CheckBox7.checked:=x;
CheckBox8.checked:=x;
CheckBox9.checked:=x;
CheckBox10.checked:=x;
end;
 function TFormRDP.FlagsEnabled(x:bool):bool;
begin
CheckBox1.Enabled:=x;
CheckBox2.Enabled:=x;
CheckBox3.Enabled:=x;
CheckBox5.Enabled:=x;
CheckBox6.Enabled:=x;
CheckBox7.Enabled:=x;
CheckBox8.Enabled:=x;
CheckBox9.Enabled:=x;
CheckBox10.Enabled:=x;
end;

function booltoint(z:boolean):integer;
begin
  if z then result:=1
  else result:=0;
end;

function inttobool(z:integer):bool;
begin
  if z=1 then result:=true
  else result:=false;
end;

function bitcolorDepth(index:integer):integer;
begin
case index of
0:result:=8;
1:result:=15;
2:result:=16;
3:result:=24;
4:result:=32;
end;
end;


function TFormRDP.createnewconnectInTab(namepc:string):bool; // перед вызовом этой функции надо прочитать настройки для компьютера из базы данных
begin
try
  if createRDP(namepc, //- имя компа
 labelededit3.Text,               // - домен
 labelededit1.Text,               //пользователь
 labelededit2.Text,               // пароль
 true,                            /// подключаться сразу при созаднии
 bitcolorDepth(ComboBoxColorDepth.ItemIndex), // количество бит на пиксель
 booltoint(CheckBoxBitmapPeristence.Checked),//BitmapPeristence
 booltoint(CheckboxCachePersistenceActive.Checked), //CachePersistenceActive
 SpinBitmapCacheSize.Value,                //BitmapCacheSize
 SpinEditCache16BppSize.Value,           //VirtualCache16BppSize
 SpinEditCache32BppSize.Value,          //VirtualCache32BppSize
 SpinBitmapVirtualCacheSize.Value,             //VirtualCacheSize
 booltoint(CheckBoxDisableCtrlAltDel.Checked),   //DisableCtrlAltDel
 booltoint(CheckBoxDoubleClickDetect.Checked),   // DoubleClickDetect
 booltoint(CheckBoxEnableWindowsKey.Checked),   // EnableWindowsKey
 SpinMinutesToIdleTimeout.Value,            //MinutesToIdleTimeout
 SpinOverallConnectionTimeout.Value,       // OverallConnectionTimeout
 strtoint(LabeledEditRdpPort.Text),      //RdpPort
 CalcPerformanceFlags(''),               //PerformanceFlags
 ComboBoxNetworkConnectionType.ItemIndex, // NetworkConnectionType
 SpinMaxReconnectAttempts.Value,           // MaxReconnectAttempts
 ComboAudioRedirectionMode.ItemIndex,      // AudioRedirectionMode
 ComboBoxAuthLevel.ItemIndex,             //  AuthenticationLevel
 ComboKey.ItemIndex,                        // KeyboardHookMode
 CheckBoxSmartSizing.Checked,         //SmartSizing
 checkboxGrabFocusOnConnect.Checked,            //GrabFocusOnConnect
 true,                                          //BandwidthDetection
 CheckEnableAutoReconnect.Checked,                  //EnableAutoReconnect
 CheckBoxDisk.Checked,                       //RedirectDrives
 CheckBoxPrint.Checked,                       //RedirectPrinters
 CheckBoxMouseMode.Checked,               //RelativeMouseMode
 CheckBoxClipboard.Checked,                   //RedirectClipboard
 CheckRedirectDevices.Checked,              //RedirectDevices
 CheckBoxPort.Checked,                       //RedirectPorts
 CheckBoxConToAdmSrv.Checked,                  //ConnectToAdministerServer
 CheckBoxRecAudio.Checked,                 //AudioCaptureRedirectionMode
 CheckBoxEnSuperPan.Checked,               //EnableSuperPan
 CheckBoxCredSsp.Checked                    //EnableCredSspSupport
 )then  result:=true
 else result:=false;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка создания подключения '+namePC+' - '+e.Message)
  end;
end;
 end;



procedure TFormRDP.findComponent; ///frmDomainInfo.findElement
var
i:integer;
CurentNamePC:string;
begin
CountTab:=3;
FormRDP.caption:=FormRDP.caption+' ВНИМАНИЕ!!! НЕ ЗАРЕГИСТРИРОВАННАЯ ВЕРСИЯ, ДОСТУПНО 3 ПОДКЛЮЧЕНИЯ.';
end;

function regeditread(patch,Section:string):string;
var      /// чтение из реестра
regFile:TregIniFile;
begin
result:='';
RegFile:=TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then
begin
result:=(regFile.ReadString(patch,Section,''));
end;
if Assigned(RegFile) then freeandnil(regFile);
end;

function existreg(patch:string):bool;  //// существование пути в реестре
var
regFile:TregInifile;
begin
regFile:=Treginifile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then result:=true
else result :=false;
end;

procedure TFormRDP.ButtonOnForForm(Sender: TObject); //// подключение по RDP на отдельной форму
var
i:integer;
namePC:string;
LevelAut:integer;
begin
try
LevelAut:=2; /// предупреждать
if not( TComponent(sender).Owner is TPanel) then exit;
for I := 0 to ( TComponent(sender).Owner as TPanel).ComponentCount-1 do
if Components[i] is Tcombobox then LevelAut:=(Components[i] as Tcombobox).ItemIndex;

if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do
 begin
     with ((TComponent(sender).Owner as TPanel).Owner as Tform) do
     if Components[i] is TMsRdpClient9 then
     begin
       namePC:=(Components[i] as TMsRdpClient9).Server;
       if (Components[i] as TMsRdpClient9).Connected<>1 then
       begin
        (Components[i] as TMsRdpClient9).AdvancedSettings9.AuthenticationLevel:=LevelAut;
        (Components[i] as TMsRdpClient9).Connect;
       end;
      break;
     end;
 end;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при подключении к компьютеру '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonOffForForm(Sender: TObject); //// отключение от RDP на отдельной форму
var
i,z:integer;
namePC:string;
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do
 begin
  with ((TComponent(sender).Owner as TPanel).Owner as Tform) do
     if Components[i] is TMsRdpClient9 then
     begin
       namePC:=(Components[i] as TMsRdpClient9).Server;
       if (Components[i] as TMsRdpClient9).Connected<>0 then
       begin
        (Components[i] as TMsRdpClient9).Disconnect;
       end;
      break;
     end;
 end;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при отключении от компьютера '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonResizeForForm(Sender: TObject); //// на весь экран  RDP на отдельной форму
var
i,z:integer;
namePC:string;
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do
 begin
 with ((TComponent(sender).Owner as TPanel).Owner as Tform) do
     if Components[i] is TMsRdpClient9 then
     begin
       namePC:=(Components[i] as TMsRdpClient9).Server;
       if (Components[i] as TMsRdpClient9).Connected<>0 then
       begin
        (Components[i] as TMsRdpClient9).Width:= Width; /// клиент по ширине и высоте формы
        (Components[i] as TMsRdpClient9).Height:=Height;
        (Components[i] as TMsRdpClient9).Reconnect(Width-1,Height-1);
       end;
      break;
     end;
 end;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при переподключении '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonCloseForForm(Sender: TObject); //// закрыть отдельную форму
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
 ((TComponent(sender).Owner as TPanel).Owner as Tform).close;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при закрытии подключения  - '+e.Message)
  end;
end;
end;


procedure TFormRDP.ShowButtonOnClick(Sender: TObject); /// кнопка на панели GroupBox8 для разворачивания и сворачивания окон
begin
try
if sender is Tbutton then
begin
if  (sender as Tbutton).Owner is Tform then
begin
if ((sender as Tbutton).Owner as Tform).WindowState=wsMinimized then
begin
 ((sender as Tbutton).Owner as Tform).WindowState:=wsMaximized;
 ((sender as Tbutton).Owner as Tform).Visible:=true;
end
 else
if ((sender as Tbutton).Owner as Tform).WindowState=wsMaximized then
begin
 ((sender as Tbutton).Owner as Tform).WindowState:=wsMinimized;
 ((sender as Tbutton).Owner as Tform).Visible:=false;
end;
end;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при сворачивании/разворачивании окна - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonMinimizeForForm(Sender: TObject); //// свернуть отдельную форму
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 //((TComponent(sender).Owner as TPanel).Owner as Tform).BorderStyle:=bsSingle;
 ((TComponent(sender).Owner as TPanel).Owner as Tform).WindowState:=wsMinimized;
 ((TComponent(sender).Owner as TPanel).Owner as Tform).Visible:=false;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при сворачивании окна - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonMaxNormalForForm(Sender: TObject); //// развернуть или свернуть отдельную форму
var
i:integer;
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
  if ((TComponent(sender).Owner as TPanel).Owner as Tform).WindowState=wsNormal then
  begin
  ((TComponent(sender).Owner as TPanel).Owner as Tform).WindowState:=wsMaximized;
  ((TComponent(sender).Owner as TPanel).Owner as Tform).BorderStyle:=bsNone;
   for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do
    if ((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] is TMsRdpClient9 then
      begin
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing:=false;
      break;
      end;
  end
  else
  if ((TComponent(sender).Owner as TPanel).Owner as Tform).WindowState=wsMaximized then
  begin
  ((TComponent(sender).Owner as TPanel).Owner as Tform).WindowState:=wsNormal;
  ((TComponent(sender).Owner as TPanel).Owner as Tform).BorderStyle:=bsSizeToolWin;
  ((TComponent(sender).Owner as TPanel).Owner as Tform).BorderWidth:=1;
      for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do
    if ((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] is TMsRdpClient9 then
      begin
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing:=true;
      break;
      end;
  end;
end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при изменении размера окна - '+e.Message)
  end;
end;
end;

procedure TFormRDP.OtherWinForRDPClientClose(Sender: TObject; var Action: TCloseAction); /// уничтожение формы при закрытии
begin
if Sender is Tform then
(Sender as TForm).Release;
end;

procedure TFormRDP.PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
begin
     movingPan:=true;
     x0:=x;
     y0:=y;
end;

procedure TFormRDP.PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// перемещение панели по  форме
  Y: Integer);
begin
     if movingPan then
   begin
     (sender as tpanel).Left:=(sender as tpanel).Left+x-x0;
     (sender as tpanel).Top:=(sender as tpanel).Top+y-y0;
   end;
end;

procedure TFormRDP.PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
begin
  movingPan := false;
end;

procedure TFormRDP.ButonMoveWinMouseDown(Sender: TObject;  /// перемещение формы за кнопку
                           Button: TMouseButton;
                            Shift: TShiftState;
                             X, Y: Integer);
var
FrmHND:Thandle;
begin
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 FrmHND:=((TComponent(sender).Owner as TPanel).Owner as Tform).Handle; // получаем хендл формы для перемешения
 ReleaseCapture;
 SendMessage(FrmHND, WM_SYSCOMMAND, 61458, 0) ;
end;
end;

procedure TFormRDP.NewRDPFormResize(Sender: TObject);
begin
if not(sender is Tform) then exit;
if (sender as Tform).WindowState=wsMaximized then
  begin
   (sender as Tform).BorderStyle:=bsNone;
  end
else
 if (sender as Tform).WindowState=wsMinimized then
  begin
   (sender as Tform).BorderStyle:=bsSingle;
  end;
end;

procedure TFormRDP.NewRDPFormShow(Sender: TObject);
begin
if not (sender is Tform) then   exit;
if (sender as Tform).BorderStyle=bsSingle then
  (sender as Tform).BorderStyle:=bsNone;
end;


procedure TFormRDP.ShowBMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if not (sender is Tbutton) then  exit;
groupbox8.Caption:='Подключение для '+(sender as Tbutton).Hint;
end;


procedure TFormRDP.ShowBMouseLeave(Sender: TObject);
begin
if not (sender is Tbutton) then exit;
groupbox8.Caption:='Подключения в отдельном окне';
end;

procedure TFormRDP.N8Click(Sender: TObject); /// попуп, закрывать подключение в отдельном окне
var
namerdp:string;
begin
namerdp:=(Tbutton(PopupOtherWin.PopupComponent).Hint);
if (Tbutton(PopupOtherWin.PopupComponent).owner is TForm) then
 (Tbutton(PopupOtherWin.PopupComponent).owner as TForm).Release;
end;


procedure TFormRDP.N9Click(Sender: TObject);
var
namerdp:string;
begin
namerdp:=(Tbutton(PopupOtherWin.PopupComponent).Hint);
if (Tbutton(PopupOtherWin.PopupComponent).owner is TForm) then
begin
 (Tbutton(PopupOtherWin.PopupComponent).owner as TForm).Release;
 if not readSettingsInDB(namerdp)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
 readSettingsInDB('SETDEFAULT');
 createnewconnectInTab(namerdp); // создание подключения
end;
end;

function TFormRDP.OtherWinForRDPClient(indTab:integer):bool;
var
NameCapt:string;
i,z:integer;
CountMonitor:LongWord;
UseRemotMon :wordBool;
UseSecondMon:boolean;
MyRDPWin:TMsRdpClient9;
FormForRDP:TForm;
PanelForBut:Tpanel;
ButFP:Tbutton;
AutLevel:TcomboBox;
step:string;
//MyRDPWinNoScr:IMsRdpClientNonScriptable5;

begin
try
inc(CountNewForm);
CoInitialize(nil);
NameCapt:=PageControl1.Pages[indTab].Caption;
FormForRDP:=Tform.Create(FormRDP);
FormForRDP.Name:='FormForRDP'+inttostr(CountNewForm);
FormForRDP.Position:=poMainFormCenter;
FormForRDP.OnClose:=OtherWinForRDPClientClose;
FormForRDP.Width:=Screen.Width;
FormForRDP.Height:=Screen.Height;
//FormForRDP.WindowState:=wsMaximized;
FormForRDP.Tag:= indTab;
FormForRDP.FormStyle:=fsStayOnTop;//fsMDIChild;//fsNormal;//fsMDIForm;//fsStayOnTop;
FormForRDP.Caption:=NameCapt;
FormForRDP.BorderStyle:=bsnone;
//FormForRDP.OnShow:=NewRDPFormShow;
//FormForRDP.OnResize:=NewRDPFormResize;


PanelForBut:=Tpanel.Create(FormForRDP);
PanelForBut.Parent:=FormForRDP;
PanelForBut.Name:='PanelForBut'+ inttostr(CountNewForm);
PanelForBut.Caption:='';//NameCapt;
//PanelForBut.Alignment:=taRightJustify;
PanelForBut.ShowHint:=true;
PanelForBut.Hint:='Подключение к '+NameCapt+'. Нажмите для перемещения';
PanelForBut.Top:=0;
PanelForBut.Left:=0;
PanelForBut.Width:=380;
//PanelForBut.Align:=alTop;
PanelForBut.Height:=30;

PanelForBut.PopupMenu:= PopupOsherWinPanel;
PanelForBut.OnMouseDown:=PanelForButMouseDown;
PanelForBut.OnMouseMove:=PanelForButMouseMove;
PanelForBut.OnMouseUp:=PanelForButMouseUp;

ButFP:=TButton.Create(PanelForBut);
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPOn'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Подключится к '+NameCapt;
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=0;
ButFP.Top:=4;
ButFP.Left:=5;
ButFP.OnClick:=ButtonOnForForm;

ButFP:=TButton.Create(PanelForBut);
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPOFF'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Отключится от '+NameCapt;
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=1;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width)+10;
ButFP.OnClick:=ButtonOffForForm;

ButFP:=TButton.Create(PanelForBut);
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPResize'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Картинка по размеру экрана';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=7;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*2)+15;
ButFP.OnClick:=ButtonResizeForForm;

AutLevel:=TCombobox.Create(PanelForBut);
AutLevel.Parent:=PanelForBut;
AutLevel.Name:='Level'+inttostr(CountNewForm);
AutLevel.Width:=150;
AutLevel.Top:=4;
AutLevel.Hint:='Проверка подлинности (CredSSP) для '+NameCapt;
AutLevel.ShowHint:=true;
AutLevel.Items.Add('Подключаться без предупреждения');
AutLevel.Items.Add('Не соединять');
AutLevel.Items.Add('Предупреждать');
AutLevel.ItemIndex:=2;
AutLevel.Left:=(ButFP.Width*3)+20;
AutLevel.TabOrder:=4;

ButFP:=TButton.Create(PanelForBut); /// перетаскивание формы
ButFP.Parent:=PanelForBut;
ButFP.Name:='MoveWin'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Переместить экран';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=12;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*3)+25+AutLevel.Width;
ButFP.OnMouseDown:=ButonMoveWinMouseDown;

ButFP:=TButton.Create(PanelForBut); /// сворачивание формы
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPCollapse'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Свернуть';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=11;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*4)+30+AutLevel.Width;
ButFP.OnClick:= ButtonMinimizeForForm;  //

ButFP:=TButton.Create(PanelForBut); /// развернуть или свернуть
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPCollapseWin'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Свернуть к окну';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=10;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*5)+35+AutLevel.Width;
ButFP.OnClick:=ButtonMaxNormalForForm;

ButFP:=TButton.Create(PanelForBut);
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPClose'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='Закрыть подключение с '+NameCapt;
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=9;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*6)+40+AutLevel.Width;
ButFP.OnClick:=ButtonCloseForForm;

if GroupBox8.Caption='' then GroupBox8.Caption:='Подключения в отдельном окне';
ButFP:=TButton.Create(FormForRDP);
ButFP.Parent:=GroupBox8;
ButFP.Name:='ShowB'+inttostr(CountNewForm);
ButFP.Caption:='';
ButFP.Hint:=NameCapt;
ButFP.ShowHint:=true;
ButFP.Images:=BigImage;
ButFP.ImageAlignment:=iaCenter;
ButFP.ImageIndex:=0;
ButFP.Tag:=CountNewForm;
ButFP.Width:=20;
ButFP.Height:=20;
ButFP.Align:=alLeft;
ButFP.OnClick:=ShowButtonOnClick;
ButFP.OnMouseMove:=ShowBMouseMove;
ButFP.OnMouseLeave:=ShowBMouseLeave;
ButFP.PopupMenu:=PopupOtherWin;

MyRDPWin:=TMsRdpClient9.Create(FormForRDP);
MyRDPWin.Parent:= FormForRDP;
MyRDPWin.Name:='NewRDPWin'+inttostr(CountNewForm);
MyRDPWin.SendToBack;/// Для того чтобы было видно панель

if not readSettingsInDB(NameCapt)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
readSettingsInDB('SETDEFAULT');

for I := 0 to PageControl1.Pages[indTab].ComponentCount-1 do
if PageControl1.Pages[indTab].Components[i] is TMsRdpClient9 then
 begin
 with PageControl1.Pages[indTab].Components[i] as TMsRdpClient9 do
 begin
 MyRDPWin.Name:='MyRDPWin'+inttostr(CountNewForm);
 MyRDPWin.OnAuthenticationWarningDismissed:=NewRdpAuthenticationWarningDismissed;
 MyRDPWin.OnAuthenticationWarningDisplayed:=NewRdpAuthenticationWarningDisplayed;
 MyRDPWin.OnAutoReconnected:=NewRdpAutoReconnected;
 MyRDPWin.OnAutoReconnecting:=MsRdpClient91AutoReconnecting;
 MyRDPWin.OnAutoReconnecting2:=MsRdpClient91AutoReconnecting2;
 MyRDPWin.OnChannelReceivedData:=NewRdpChannelReceivedData;
 MyRDPWin.OnConfirmClose:=NewRdpConfirmClose;
 MyRDPWin.OnConnected:=NewRdpConnected;
 MyRDPWin.OnConnecting:=NewRdpConnecting;
 MyRDPWin.OnDisconnected:=NewRdpDisconnected;
 MyRDPWin.OnExit:=NewRdpExit;
 MyRDPWin.OnFatalError:=NewRdpFatalError;
 MyRDPWin.OnLoginComplete:=NewRdpLoginComplete;
 MyRDPWin.OnLogonError:=NewRdpLogonError;
 MyRDPWin.OnNetworkStatusChanged:=NewRdpNetworkStatusChanged;
 MyRDPWin.OnWarning:=NewRdpWarning;
 MyRDPWin.OnServiceMessageReceived:=MsRdpClient91ServiceMessageReceived;
 MyRDPWin.ConnectingText:='Установка соединения c '+NameCapt+'...';
 MyRDPWin.DisconnectedText:='Клиент '+NameCapt+' отключен';
 MyRDPWin.Server:=Widestring(NameCapt);
 if Domain<>'' then  MyRDPWin.Domain:=WideString(Domain)
 else MyRDPWin.Domain:=LabeledEdit3.Text;
 if UserName<>'' then MyRDPWin.UserName:=UserName
 else MyRDPWin.UserName:='';



step:='';
MyRDPWin.ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex);
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=booltoint(CheckBoxBitmapPeristence.Checked); // 0 выключить,1 включить.  Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='CachePersistenceActive';
MyRDPWin.AdvancedSettings9.CachePersistenceActive:=booltoint(CheckboxCachePersistenceActive.Checked); // 0 - отключить, 1 - включить. Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='BitmapCacheSize';
MyRDPWin.AdvancedSettings9.BitmapCacheSize:=SpinBitmapCacheSize.Value;; // от 1 до 32 . Размер файла кэша растровых изображений в килобайтах, используемого для растровых изображений 8 бит на пиксель
step:='VirtualCache16BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=SpinEditCache16BppSize.Value; // Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCache32BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=SpinEditCache32BppSize.Value;// Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCacheSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=SpinBitmapVirtualCacheSize.Value;  //Задает размер файла постоянного кэша растрового изображения в мегабайтах, который используется для цвета 8 бит на пиксель. Допустимые числовые значения этого свойства - от 1 до 32 включительно. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ. Связанные свойства включают свойства BitmapVirtualCache16BppSize и BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=booltoint(CheckBoxDisableCtrlAltDel.Checked); // 0 - отключить, 1 - включить.  Указывает, должен ли отображаться начальный пояснительный экран в Winlogon.
step:='SmartSizing';
MyRDPWin.AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked); //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=booltoint(CheckBoxEnableWindowsKey.Checked);   //0 - отключить, 1 - включить. Указывает, можно ли использовать ключ Windows в удаленном сеансе.
step:='GrabFocusOnConnect';
MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked; //0 - отключить, 1 - включить. Указывает, должен ли элемент управления клиента иметь фокус при подключении. Элемент управления не будет пытаться захватить фокус из окна, запущенного в другом процессе.
step:='MinutesToIdleTimeout';
MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= SpinMinutesToIdleTimeout.Value;  //По умолчанию 0 только т.е. не отслеживать время. Задает максимальный промежуток времени в минутах, в течение которого клиент должен оставаться подключенным без участия пользователя. Если указанное время истекло, элемент управления вызывает метод IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= SpinOverallConnectionTimeout.Value; //Определяет общую продолжительность времени в секундах, в течение которого клиентский элемент управления ожидает завершения соединения. Максимальное допустимое значение этого свойства - 600, что соответствует 10 минутам. Если указанное время истекает до завершения соединения, элемент управления отключается и вызывает метод IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
MyRDPWin.AdvancedSettings9.RdpPort:=strtoint(LabeledEditRdpPort.Text);


{MyRDPWinNoScr:=(MyRDPWin.ControlInterface as IMsRdpClientNonScriptable5);//.Set_AllowPromptingForCredentials(false);
//MyRDPWinNoScr.Set_AllowPromptingForCredentials(false); ////// если это применить то окно авторизации не выходит  //////// и если проверка на уровне сети то хуй подключишься
MyRDPWinNoScr.get_RemoteMonitorCount(CountMonitor); /// возвращает количество удаленных мониторов
MyRDPWinNoScr.get_RemoteMonitorLayoutMatchesLocal(UseRemotMon); //возвращает, идентична ли схема удаленного монитора локальной компоновке монитора. Если это свойство содержит VARIANT_TRUE , компоновка удаленного монитора идентична компоновке локального монитора. Если это свойство содержит VARIANT_FALSE , удаленный и локальный мониторы имеют разные макеты.
 //MyRDPWinNoScr.Get_UseMultimon(UseSecondMon); /// запрашивает, используется ли несколько мониторов
 //MyRDPWinNoScr.Set_UseMultimon(True); /// устанавливает использование нескольких мониторов
   showmessage('Количество удаленных мониторов - '+inttostr(CountMonitor));

  z:=MessageBox(Self.Handle, PChar('Использовать все мониторы для подключения?')
        , PChar('Компьютер - '+server ) ,MB_YESNO+MB_ICONQUESTION);
   if z=IDYES then
   begin
    if UseRemotMon then ShowMessage('Схема мониторов идентична') else ShowMessage('Схема мониторов не идентична');
    MyRDPWinNoScr.Set_UseMultimon(True); /// устанавливаем несколько мониторов
   end; }

  if ComboBoxNetworkConnectionType.ItemIndex=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=true;          /// если NetworkConnectionType = 0 то....яет, будут ли автоматически обнаружены изменения пропускной способности.
//MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end
else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // установка типа используемой сети 1- модем, 2-Низкоскоростная широкополосная связь (от 256 Кбит / с до 2 Мбит / с), 3Спутник (от 2 Мбит / с до 16 Мбит / с, с высокой задержкой), 4 -Высокоскоростной широкополосный доступ (от 2 Мбит / с до 10 Мбит / с), 5- Глобальная сеть (WAN) (10 Мбит / с или выше, с высокой задержкой), 6 -Локальная сеть (LAN) (10 Мбит / с или выше)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  Задает набор функций, которые можно установить на сервере для повышения производительности.
end;
//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:= SpinMaxReconnectAttempts.Value;   //Задает количество попыток переподключения во время автоматического переподключения. Допустимые значения этого свойства - от 0 до 200 включительно.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:= CheckEnableAutoReconnect.Checked;       //Указывает, следует ли разрешить клиентскому элементу управления автоматически подключаться к сеансу в случае отключения сети.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=ComboAudioRedirectionMode.ItemIndex;  // 0-Перенаправить звуки на клиента. 1-Воспроизведение звуков на удаленном компьютере. 2-Отключить перенаправление звука; не воспроизводить звуки на сервере.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=CheckBoxDisk.Checked;     ///разрешить перенаправлять диски
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=CheckBoxPrint.Checked;    /// разрешить перенаправлять принтеры
if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// если необходимо перенаправить принтеры и буфер обмена то
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// Установите для этого параметра значение 1, чтобы отключить перенаправление, или 0, чтобы включить перенаправление.
step:='RelativeMouseMode';
MyRDPWin.AdvancedSettings9.RelativeMouseMode:=CheckBoxMouseMode.Checked;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=CheckBoxClipboard.Checked;  /// Указывает, следует ли включить или отключить перенаправление буфера обмена.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=CheckRedirectDevices.Checked;   /// Указывает, должны ли перенаправленные устройства быть включены или отключены.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=CheckBoxPort.Checked;       /// перенаправление портов
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=CheckBoxConToAdmSrv.Checked; /// указывает должен ли элемент управления ActiveX пытаться подключиться к серверу в административных целях.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=CheckBoxRecAudio.Checked;  ///  указывает, перенаправляется ли устройство аудиовхода по умолчанию от клиента к удаленному сеансу
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=CheckBoxEnSuperPan.Checked;              ///позволяет пользователю легко перемещаться по удаленному рабочему столу в полноэкранном режиме, когда размеры удаленного рабочего стола больше размеров текущего окна клиента. Вместо того, чтобы использовать полосы прокрутки для навигации по рабочему столу, пользователь может указать на границу окна, и удаленный рабочий стол будет автоматически прокручиваться в этом направлении. SuperPan не поддерживает более одного монитора.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=ComboBoxAuthLevel.ItemIndex;           ///Задает уровень проверки подлинности для подключения.  (0.1.2)0- Нет аутентификации на сервере. 1- Проверка подлинности сервера требуется и должна успешно завершиться для продолжения подключения. 2- Попытка аутентификации сервера. Если аутентификация не удалась, пользователю будет предложено отменить соединение или продолжить без аутентификации на сервере.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=CheckBoxCredSsp.Checked;       ///Указывает или извлекает, включен ли для этого подключения  Поставщик службы безопасности учетных данных (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=ComboKey.ItemIndex; // 0-Применяйте комбинации клавиш только локально на клиентском компьютере. 1- Применяйте комбинации клавиш на удаленном сервере. 2- Применяйте комбинации клавиш к удаленному серверу только тогда, когда клиент работает в полноэкранном режиме. Это значение по умолчанию.

end;
break;
end;
//FormForRDP.WindowState:=wsNormal;
FormForRDP.Show;
MyRDPWin.Align:=alClient;
//MyRDPWin.Width:=FormForRDP.Width;
//MyRDPWin.Height:=FormForRDP.Height;
MyRDPWin.Connect;
PageControl1.Pages[indTab].Destroy;
CoUninitialize;
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка создания отдельного окна  :'+E.Message);
        if assigned(FormForRDP) then
         if FormForRDP.visible=false then  FormForRDP.release;
      end;
   end;
end;


 function Check(s:string): boolean;
begin
  HM:= OpenMutex(MUTEX_ALL_ACCESS, false, pchar(s));  //MForMRPC
  Result:= (HM <> 0);
  if HM = 0 then HM:= CreateMutex(nil, false, pchar(s)); //MForMRPC
end;

function ActProc(TipForm,NameForm:string):bool; /// функция поиска приложения и вывод на передний план если оно запущено
var
hwn: THandle;
begin           //'TForm1', 'Form1'
  try
  hwn := FindWindow(pchar(TipForm),nil);
  if hwn<>0 then
  begin
  //OwnerHwn:=GetWindow(hwn,GW_OWNER); /// получаем родителя
  ShowWindow(GetWindow(hwn,GW_OWNER), SW_SHOWNORMAL); //SW_RESTORE   // SW_SHOWMAXIMIZED  //SW_SHOWNORMAL
  SetForegroundWindow(GetWindow(hwn,GW_OWNER));
  end
  else result:=false;
  except
   result:=false;
   end;
end;


function TFormRDP.ConnectToCurentSetting(NamePC:string;IndexTab:integer):boolean;
var     //// применение всех настроек для созданнго но отключенного подключения
i:integer;
begin
for I := 0 to PageControl1.pages[IndexTab].ComponentCount-1 do
begin
 if PageControl1.pages[IndexTab].Components[i] is TMsRdpClient9 then
 with PageControl1.pages[IndexTab].Components[i] as TMsRdpClient9 do
  begin
  ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex); // количество бит на пиксель
  AdvancedSettings9.BitmapPeristence:=booltoint(CheckBoxBitmapPeristence.Checked);//BitmapPeristence
  AdvancedSettings9.CachePersistenceActive:=booltoint(CheckboxCachePersistenceActive.Checked); //CachePersistenceActive
  AdvancedSettings9.BitmapCacheSize:=SpinBitmapCacheSize.Value;                //BitmapCacheSize
  AdvancedSettings9.BitmapVirtualCache16BppSize:=SpinEditCache16BppSize.Value;           //VirtualCache16BppSize
  AdvancedSettings9.BitmapVirtualCache32BppSize:= SpinEditCache32BppSize.Value;          //VirtualCache32BppSize
  AdvancedSettings9.BitmapVirtualCacheSize:=SpinBitmapVirtualCacheSize.Value;             //VirtualCacheSize
  AdvancedSettings9.DisableCtrlAltDel:=booltoint(CheckBoxDisableCtrlAltDel.Checked);   //DisableCtrlAltDel
  AdvancedSettings9.DoubleClickDetect:=booltoint(CheckBoxDoubleClickDetect.Checked);   // DoubleClickDetect
  AdvancedSettings9.EnableWindowsKey:= booltoint(CheckBoxEnableWindowsKey.Checked);   // EnableWindowsKey
  AdvancedSettings9.MinutesToIdleTimeout:= SpinMinutesToIdleTimeout.Value;            //MinutesToIdleTimeout
  AdvancedSettings9.OverallConnectionTimeout:= SpinOverallConnectionTimeout.Value;       // OverallConnectionTimeout
  AdvancedSettings9.RdpPort:= strtoint(LabeledEditRdpPort.Text);      //RdpPort
  AdvancedSettings9.PerformanceFlags:= CalcPerformanceFlags('');               //PerformanceFlags
  ConnectingText:='Установка соединения c '+NamePC+'...';
  DisconnectedText:='Клиент '+NamePC+' отключен';
  if ComboBoxNetworkConnectionType.ItemIndex=0 then
  begin
  AdvancedSettings9.BandwidthDetection:=true;          /// если NetworkConnectionType = 0 то....яет, будут ли автоматически обнаружены изменения пропускной способности.
  AdvancedSettings9.PerformanceFlags:=0;
  end
  else
  begin
  AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // установка типа используемой сети 1- модем, 2-Низкоскоростная широкополосная связь (от 256 Кбит / с до 2 Мбит / с), 3Спутник (от 2 Мбит / с до 16 Мбит / с, с высокой задержкой), 4 -Высокоскоростной широкополосный доступ (от 2 Мбит / с до 10 Мбит / с), 5- Глобальная сеть (WAN) (10 Мбит / с или выше, с высокой задержкой), 6 -Локальная сеть (LAN) (10 Мбит / с или выше)
  AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  Задает набор функций, которые можно установить на сервере для повышения производительности.
  end;
  AdvancedSettings9.MaxReconnectAttempts:= SpinMaxReconnectAttempts.Value;           // MaxReconnectAttempts
  AdvancedSettings9.AudioRedirectionMode:= ComboAudioRedirectionMode.ItemIndex;      // AudioRedirectionMode
  AdvancedSettings9.AuthenticationLevel:= ComboBoxAuthLevel.ItemIndex;             //  AuthenticationLevel
  SecuredSettings3.KeyboardHookMode := ComboKey.ItemIndex;                        // KeyboardHookMode
  AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked);         //SmartSizing
  AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked;            //GrabFocusOnConnect
  AdvancedSettings9.BandwidthDetection:=true;                                          //BandwidthDetection
  AdvancedSettings9.EnableAutoReconnect:=CheckEnableAutoReconnect.Checked;                  //EnableAutoReconnect
  if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// если необходимо перенаправить принтеры и буфер обмена то
  AdvancedSettings.DisableRdpdr:=0; //// Установите для этого параметра значение 1, чтобы отключить перенаправление, или 0, чтобы включить перенаправление.
  AdvancedSettings9.RedirectDrives:=CheckBoxDisk.Checked;                       //RedirectDrives
  AdvancedSettings9.RedirectPrinters:=CheckBoxPrint.Checked;                       //RedirectPrinters
  AdvancedSettings9.RelativeMouseMode:=CheckBoxMouseMode.Checked;               //RelativeMouseMode
  AdvancedSettings9.RedirectClipboard:=CheckBoxClipboard.Checked;                   //RedirectClipboard
  AdvancedSettings9.RedirectDevices:=CheckRedirectDevices.Checked;              //RedirectDevices
  AdvancedSettings9.RedirectPorts:=CheckBoxPort.Checked;                       //RedirectPorts
  AdvancedSettings9.ConnectToAdministerServer:=CheckBoxConToAdmSrv.Checked;                  //ConnectToAdministerServer
  AdvancedSettings9.AudioCaptureRedirectionMode:=CheckBoxRecAudio.Checked;                 //AudioCaptureRedirectionMode
  AdvancedSettings9.EnableSuperPan:= CheckBoxEnSuperPan.Checked;               //EnableSuperPan
  AdvancedSettings9.EnableCredSspSupport:= CheckBoxCredSsp.Checked;                   //EnableCredSspSupport
  end;
end;
end;

function  TFormRDP.CalcPerformanceFlags(s:string):integer;
var
flgs:integer;
begin    //// если нет отмеченых галок то flgs=0 ничего не отключено
flgs:=0;

if CheckBox2.checked then  flgs:= flgs + 1;  ///откл фон рабочего стоа
if CheckBox3.checked then  flgs:= flgs + 2;  //Перетаскивание в полном окне отключено; при перемещении окна отображается только контур окна.
if CheckBox5.checked then  flgs:= flgs + 4; //Анимации меню отключены.
if CheckBox6.checked then  flgs:= flgs + 8; //Темы отключены.
if CheckBox7.checked then  flgs:= flgs + 16; //Включить визуальные эффекты при отображении меню и окон (+10)
if CheckBox8.checked then  flgs:= flgs + 32; // Отключить тень для курсора
if CheckBox9.checked then  flgs:= flgs + 64; // Отключить мигание курсора
if CheckBox1.checked then  flgs:= flgs + 128; //Включить сглаживание шрифтов
if CheckBox10.checked then  flgs:= flgs + 256; //Включить композицию на рабочем столе
result:=flgs;
end;



procedure TFormRDP.CheckBoxSmartSizingClick(Sender: TObject);
var
i:integer;
begin
try
if PageControl1.PageCount=0 then exit;
for I := 0 to PageControl1.ActivePage.ComponentCount-1 do
 begin
   if PageControl1.ActivePage.Components[i] is TMsRdpClient9 then
    begin
     (PageControl1.ActivePage.Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing:=
     not (PageControl1.ActivePage.Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing;
    break;
    end;
 end;
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка SmartSizing  :'+E.Message);
      end;
   end;
end;


///////////////////////////////////////////////////////////////////////////////////////
//////////////////Начало описания и действий кнопки закрыть на каждой странице//////////////////////////////////
constructor TTabControlStyleHookBtnClose.Create(AControl: TWinControl);
begin
  inherited;
  FHotIndex:=-1;
  FWidthModified:=False;
end;

procedure TTabControlStyleHookBtnClose.DrawTab(Canvas: TCanvas; Index: Integer);
var
  Details : TThemedElementDetails;
  ButtonR : TRect;
  FButtonState: TThemedWindow;
begin
  inherited;
  if Control is TPageControl then
  if (Control as TPageControl).Name='PageControl1' then
  begin
  if (FHotIndex>=0) and (Index=FHotIndex) then
   FButtonState :=twCloseButtonHot //twSmallCloseButtonHot
  else
  if Index = TabIndex then
   FButtonState := twCloseButtonNormal//twSmallCloseButtonNormal
  else
   FButtonState := twCloseButtonDisabled;//twSmallCloseButtonDisabled;

  Details := StyleServices.GetElementDetails(FButtonState);

  ButtonR:= GetButtonCloseRect(Index);

  if ButtonR.Bottom - ButtonR.Top > 0 then
   StyleServices.DrawElement(Canvas.Handle, Details, ButtonR);
  end;
end;

procedure TTabControlStyleHookBtnClose.WMLButtonUp(var Message: TWMMouse);
Var
  LPoint : TPoint;
  LIndex : Integer;
begin
  if Control is TPageControl then
  if (Control as TPageControl).Name='PageControl1' then
  begin
  LPoint:=Message.Pos;
  for LIndex := 0 to TabCount-1 do
   if PtInRect(GetButtonCloseRect(LIndex), LPoint) then
   begin
      if Control is TPageControl then
      begin
        TPageControl(Control).Pages[LIndex].Parent:=nil;
        TPageControl(Control).Pages[LIndex].Free;
      end;
      break;
   end;
  end;
end;

procedure TTabControlStyleHookBtnClose.WMMouseMove(var Message: TMessage);
Var
  LPoint : TPoint;
  LIndex : Integer;
  LHotIndex : Integer;
begin
  if Control is TPageControl then
  if (Control as TPageControl).Name='PageControl1' then
  begin
  inherited;
  LHotIndex:=-1;
  LPoint:=TWMMouseMove(Message).Pos;
  for LIndex := 0 to TabCount-1 do
   if PtInRect(GetButtonCloseRect(LIndex), LPoint) then
   begin
      LHotIndex:=LIndex;
      break;
   end;

   if (FHotIndex<>LHotIndex) then
   begin
     FHotIndex:=LHotIndex;
     Invalidate;
   end;
  end;
end;

function TTabControlStyleHookBtnClose.GetButtonCloseRect(Index: Integer): TRect;
var
  FButtonState: TThemedWindow;
  Details : TThemedElementDetails;
  R, ButtonR : TRect;
begin
if Control is TPageControl then
  if (Control as TPageControl).Name='PageControl1' then
  begin
  R := TabRect[Index];
  if R.Left < 0 then Exit;

  if TabPosition in [tpTop, tpBottom] then
  begin
    if Index = TabIndex then
      InflateRect(R, 0, 2);
  end
  else
  if Index = TabIndex then
    Dec(R.Left, 2)
  else
    Dec(R.Right, 2);

  Result := R;
  FButtonState :=twCloseButtonNormal;//twSmallCloseButtonNormal;

  Details := StyleServices.GetElementDetails(FButtonState);
  if not StyleServices.GetElementContentRect(0, Details, Result, ButtonR) then
    ButtonR := Rect(0, 0, 0, 0);

  //Result.Left :=Result.Right - (ButtonR.Width) - 5; ///расположение  кнопки на вкладке
  Result.Left := Result.Right - (ButtonR.Width);
  Result.Width:=ButtonR.Width;
  end;
end;

procedure TTabControlStyleHookBtnClose.MouseEnter;
begin
if Control is TPageControl then
  if (Control as TPageControl).Name='PageControl1' then
  begin
  inherited;
  FHotIndex := -1;
  end;
end;

procedure TTabControlStyleHookBtnClose.MouseLeave;
begin
  inherited;
  if FHotIndex >= 0 then
  begin
    FHotIndex := -1;
    Invalidate;
  end;
end;

///////////////// конец описания и действий кнопки для закрытия страницы///////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////
function TFormRDP.readSettingsInDB(s:string):bool;
begin
{PC_NAME,ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,'
+'VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,'
+'DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,'
+'OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,'
+'MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode,'
+'SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,'
+'RedirectDrives,RedirectPrinters,RelativeMouseMode'
+',RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,'
+'AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport}
try
if not DataMod.ConnectionDB.Connected then
begin
result:=false;
memo1.Lines.add('Не подключена база данных');
exit;
end;
DataMod.FDQueryRead.SQL.Clear;
DataMod.FDQueryRead.SQL.Text:='SELECT * FROM RDP_SET WHERE PC_NAME='''+s+'''';
DataMod.FDQueryRead.Open;
if DataMod.FDQueryRead.FieldByName('PC_NAME').Value<>null then
begin
ComboBoxColorDepth.ItemIndex:=DataMod.FDQueryRead.FieldByName('ColorDepth').AsInteger;
CheckBoxBitmapPeristence.Checked:=inttobool(DataMod.FDQueryRead.FieldByName('BitmapPeristence').AsInteger);
CheckboxCachePersistenceActive.Checked:=inttobool(DataMod.FDQueryRead.FieldByName('CachePersistenceActive').Asinteger);
SpinBitmapCacheSize.Value:=DataMod.FDQueryRead.FieldByName('BitmapCacheSize').AsInteger;
SpinEditCache16BppSize.Value:= DataMod.FDQueryRead.FieldByName('VirtualCache16BppSize').AsInteger;
SpinEditCache32BppSize.Value:=DataMod.FDQueryRead.FieldByName('VirtualCache32BppSize').AsInteger;
SpinBitmapVirtualCacheSize.Value:=DataMod.FDQueryRead.FieldByName('VirtualCacheSize').AsInteger;
CheckBoxDisableCtrlAltDel.Checked:=inttobool(DataMod.FDQueryRead.FieldByName('DisableCtrlAltDel').AsInteger);
CheckBoxDoubleClickDetect.Checked:=inttobool(DataMod.FDQueryRead.FieldByName('DoubleClickDetect').AsInteger);
CheckBoxEnableWindowsKey.Checked:=inttobool(DataMod.FDQueryRead.FieldByName('EnableWindowsKey').AsInteger);
SpinMinutesToIdleTimeout.Value:=DataMod.FDQueryRead.FieldByName('MinutesToIdleTimeout').AsInteger;
SpinOverallConnectionTimeout.Value := DataMod.FDQueryRead.FieldByName('OverallConnectionTimeout').AsInteger;
LabeledEditRdpPort.Text:= inttostr(DataMod.FDQueryRead.FieldByName('RdpPort').AsInteger);
ComboBoxNetworkConnectionType.ItemIndex:= DataMod.FDQueryRead.FieldByName('NetworkConnectionType').AsInteger;
SpinMaxReconnectAttempts.Value :=  DataMod.FDQueryRead.FieldByName('MaxReconnectAttempts').AsInteger;
ComboAudioRedirectionMode.ItemIndex:=DataMod.FDQueryRead.FieldByName('AudioRedirectionMode').AsInteger;
ComboBoxAuthLevel.ItemIndex:=DataMod.FDQueryRead.FieldByName('AuthenticationLevel').AsInteger;
ComboKey.ItemIndex:=DataMod.FDQueryRead.FieldByName('KeyboardHookMode').AsInteger;
CheckBoxSmartSizing.Checked:=DataMod.FDQueryRead.FieldByName('SmartSizing').AsBoolean;
checkboxGrabFocusOnConnect.Checked:=DataMod.FDQueryRead.FieldByName('GrabFocusOnConnect').AsBoolean;
CheckEnableAutoReconnect.Checked:=DataMod.FDQueryRead.FieldByName('EnableAutoReconnect').AsBoolean;
CheckBoxDisk.Checked:=DataMod.FDQueryRead.FieldByName('RedirectDrives').AsBoolean;
CheckBoxPrint.Checked:=DataMod.FDQueryRead.FieldByName('RedirectPrinters').AsBoolean;
CheckBoxMouseMode.Checked:=DataMod.FDQueryRead.FieldByName('RelativeMouseMode').AsBoolean;
CheckBoxClipboard.Checked:=DataMod.FDQueryRead.FieldByName('RedirectClipboard').AsBoolean;
CheckRedirectDevices.Checked:=DataMod.FDQueryRead.FieldByName('RedirectDevices').AsBoolean;
CheckBoxPort.Checked:=DataMod.FDQueryRead.FieldByName('RedirectPorts').AsBoolean;
CheckBoxConToAdmSrv.Checked:=DataMod.FDQueryRead.FieldByName('ConnectToAdministerServer').AsBoolean;
CheckBoxRecAudio.Checked:=DataMod.FDQueryRead.FieldByName('AudioCaptureRedirectionMode').AsBoolean;
CheckBoxEnSuperPan.Checked:=DataMod.FDQueryRead.FieldByName('EnableSuperPan').AsBoolean;
CheckBoxCredSsp.Checked:=DataMod.FDQueryRead.FieldByName('EnableCredSspSupport').AsBoolean;
result:=true;
end else result:=false; /// если настройки для компа не нашли то false
DataMod.FDQueryRead.Close;
Except
  on E: Exception do
  begin
  result:=false;
  memo1.Lines.add('При чтении настроек для '+s+' произошла ошибка '+E.Message);
  end;
end;
end;

function TFormRDP.sizeRegion(i,z:integer):bool; /// проверка размеров окон просмотра
var
res:integer;
begin   /// если размер окна просмотра больше на 15 пикселей вклаки на которой он находится то меняем его размер
if i>z then  res:=i-z
else res:=z-i;
if res>=15 then result:=true
else result:=false;
end;

function TFormRDP.loadlistpc(sFile:string):Tstringlist; /// 
var
i:integer;
begin
try
result:=TstringList.Create;
if FileExists(ExtractFileDir(Application.ExeName)+'\'+sFile) then  // s- 'RDPlist.dat'
 begin
 result.LoadFromFile(ExtractFileDir(Application.ExeName)+'\'+sFile);
 end;
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка загрузки списка подключений :'+E.Message);
  end;
end;
end;

 function TFormRDP.ping(s:string):boolean;
var
z:integer;
Myidicmpclient:TIdIcmpClient;
begin
try
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=3000;
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  Memo1.Lines.Add(s+': Превышен интервал ожидания запроса');
  end
else
  begin
  result:=true; ///доступен
 Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;
   except
    begin
    result:=false;
    Memo1.Lines.Add(s+': Узел не доступен');
    end;
   end;
Myidicmpclient.Free;
end;





procedure TFormRDP.PopupMenu1Popup(Sender: TObject);
begin
if PageControl1.MultiLine then PopupMenu1.Items[7].Caption:='Вкладки в один ряд'
else PopupMenu1.Items[7].Caption:='Вкладки Multiline'

end;



procedure TFormRDP.Button1Click(Sender: TObject);
var
i,z:integer;
begin
try
//Memo1.Lines.Add('Активная страница - '+inttostr(PageControl1.ActivePageIndex));
//PageControl1.ActivePage
if PageControl1.PageCount=0 then  exit;

for I := 0 to PageControl1.ActivePage.ComponentCount-1 do //// смотрим компоненты в цикле
begin
 if PageControl1.ActivePage.Components[i] is TMsRdpClient9 then /// если нашли TMsRdpClient9
 begin
 (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Width:=PageControl1.ActivePage.Width;
 (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Height:=PageControl1.ActivePage.Height;
 (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Reconnect(PageControl1.ActivePage.Width,PageControl1.ActivePage.Height);
 break;
 end;
end;

Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка изменения размера окна :'+E.Message);
  end;
end;
end;



procedure TFormRDP.ComboBoxNetworkConnectionTypeChange(Sender: TObject);
var
i:integer;
begin   ///// изменение типа сети по которой подключаемся по RDP
case (sender as Tcombobox).ItemIndex of
0: // Определить автоматически , Глобальная сеть (10 Мбит/с или выще с большой задержкой),Локальная сеть (10 Мбит/с и выше)
  begin
   FlagsEnabled(false);
  end;
1:  //Модем (56 кбит/с)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // снимаем флаги
  checkbox2.Checked:=true;
  checkbox3.Checked:=true;
  checkbox5.Checked:=true;
  checkbox6.Checked:=true;
  checkbox8.Checked:=true;
  checkbox9.Checked:=true;
end;
2: //Низкоскоростное широкополосное подключение (256 Кбит/с - 2 Мбит/с)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // снимаем флаги
  checkbox2.Checked:=true;
  checkbox3.Checked:=true;
  checkbox5.Checked:=true;
  checkbox8.Checked:=true;
  checkbox9.Checked:=true;
end;
3,4: // Спутник (2-16 Мбит/с с большой задержкой), Высокоскоростное широкополосное подключение (2-10 Мбит/с)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // снимаем флаги
  checkbox2.Checked:=true;
  checkbox3.Checked:=true;
  checkbox5.Checked:=true;
  checkbox8.Checked:=true;
  checkbox9.Checked:=true;
  checkbox1.Checked:=true;
  checkbox10.Checked:=true;
end;
5,6:
begin
  FlagsEnabled(true);
  FlagsChecked(false); // снимаем флаги
  checkbox2.Checked:=false;
  checkbox3.Checked:=false;
  checkbox5.Checked:=false;
  checkbox6.Checked:=false;
  checkbox7.Checked:=true;
  checkbox8.Checked:=false;
  checkbox9.Checked:=false;
  checkbox1.Checked:=true;
  checkbox10.Checked:=true;
end;

end;
end;

procedure TFormRDP.ComboKompKeyPress(Sender: TObject; var Key: Char);
begin
if key=#13 then 
begin
if not readSettingsInDB(ComboKomp.Text)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
readSettingsInDB('SETDEFAULT');
SpeedButton3.Click; /// если нажали enter
end;
end;


procedure TFormRDP.ComboKompSelect(Sender: TObject);
begin /// при выборе компьютера считываем настройки из бд
if not readSettingsInDB(ComboKomp.Text)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
readSettingsInDB('SETDEFAULT');
end;

function TFormRDP.createRDP(NamePC,Domain,UserName,Passwd:string;AutoConnect:boolean;
ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):bool;
var
Ntab: TTabSheet;
MyRDPWin:  TMsRdpClient9;
MyRDPWinNoScr:IMsRdpClientNonScriptable5;//TMsRdpClient9;//TMsRdpClient9NotSafeForScripting;
step:string;
 hToken : THandle;
 errLogUs:integer;

begin
try
 PageControl1.TabWidth:=150;
 PageControl1.OwnerDraw:=true;
 Ntab:=TTabSheet.Create(PageControl1);
 with Ntab do
  begin
    PageControl := PageControl1;
    Name:='Ntab'+inttostr(TabShetCount);
    Caption := NamePC;
    ImageIndex:=0;
    tag:=TabShetCount;
    DoubleBuffered:=true;
    Highlighted:=true;
    Align:=alClient;
    //Width:=PageControl1.Width;
    //Height:=PageControl.Height;
  end;
  ///////////////////////////////////////////////
  {try
  if AutoConnect then
  if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (domain),
                               PAnsiChar (passwd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
                               LOGON32_PROVIDER_WINNT50, hToken))
    then GetLastError();
  //CloseHandle (hToken);
  except
        on E: Exception do Memo1.Lines.Add('Ошибка при LogonUser '+NamePC+' - '+e.Message)
   end; }
  ////////////////////////////////////////////
 MyRDPWin:=TMsRdpClient9.Create(Ntab);
 //MyRDPWin.ParentWindow:=Ntab.Handle;
 // MyRDPWin.Parent:=nil;
 MyRDPWin.Parent:=Ntab;
 MyRDPWin.Height:=Ntab.Height;
 MyRDPWin.Width:=Ntab.Width;
 MyRDPWin.DesktopWidth:=Ntab.Width;
 MyRDPWin.DesktopHeight:=Ntab.Height;
 MyRDPWin.Align:=alClient;


 MyRDPWin.Name:='MyRDPWin'+inttostr(TabShetCount);
 MyRDPWin.OnAuthenticationWarningDismissed:=NewRdpAuthenticationWarningDismissed;
 MyRDPWin.OnAuthenticationWarningDisplayed:=NewRdpAuthenticationWarningDisplayed;
 MyRDPWin.OnAutoReconnected:=NewRdpAutoReconnected;
 MyRDPWin.OnAutoReconnecting:=MsRdpClient91AutoReconnecting;
 MyRDPWin.OnAutoReconnecting2:=MsRdpClient91AutoReconnecting2;
 MyRDPWin.OnChannelReceivedData:=NewRdpChannelReceivedData;
 MyRDPWin.OnConfirmClose:=NewRdpConfirmClose;
 MyRDPWin.OnConnected:=NewRdpConnected;
 MyRDPWin.OnConnecting:=NewRdpConnecting;
 MyRDPWin.OnDisconnected:=NewRdpDisconnected;
 MyRDPWin.OnExit:=NewRdpExit;
 MyRDPWin.OnFatalError:=NewRdpFatalError;
 MyRDPWin.OnLoginComplete:=NewRdpLoginComplete;
 MyRDPWin.OnLogonError:=NewRdpLogonError;
 MyRDPWin.OnNetworkStatusChanged:=NewRdpNetworkStatusChanged;
 MyRDPWin.OnWarning:=NewRdpWarning;
 MyRDPWin.OnServiceMessageReceived:=MsRdpClient91ServiceMessageReceived;
 MyRDPWin.ConnectingText:='Установка соединения с '+NamePC+'...';
 MyRDPWin.DisconnectedText:='Клиент '+NamePC+' отключен';
 MyRDPWin.Server:=Widestring(NamePC);
 MyRDPWin.Domain:=WideString(Domain);
 if UserName<>'' then MyRDPWin.UserName:=Widestring(UserName);
 if Passwd<>'' then MyRDPWin.AdvancedSettings9.ClearTextPassword:=Widestring(Passwd);

//MyRDPWin.SecuredSettings.StartProgram:='notepad.exe';
step:='';
//MyRDPWin.AdvancedSettings9.RedirectPOSDevices:=true; // Устанавливает или извлекает конфигурацию для перенаправления устройства Point of Service.
MyRDPWin.ColorDepth:=ColorDepth; //// глубина цвета 32-24-16-8 бит на пиксель
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=BitmapPeristence; // 0 выключить,1 включить.  Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='CachePersistenceActive';
MyRDPWin.AdvancedSettings9.CachePersistenceActive:=CachePersistenceActive; // 0 - отключить, 1 - включить. Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='BitmapCacheSize';
MyRDPWin.AdvancedSettings9.BitmapCacheSize:=BitmapCacheSize; // от 1 до 32 . Размер файла кэша растровых изображений в килобайтах, используемого для растровых изображений 8 бит на пиксель
step:='VirtualCache16BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=VirtualCache16BppSize;// Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCache32BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=VirtualCache32BppSize;// Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCacheSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=VirtualCacheSize; //Задает размер файла постоянного кэша растрового изображения в мегабайтах, который используется для цвета 8 бит на пиксель. Допустимые числовые значения этого свойства - от 1 до 32 включительно. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ. Связанные свойства включают свойства BitmapVirtualCache16BppSize и BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=DisableCtrlAltDel; // 0 - отключить, 1 - включить.  Указывает, должен ли отображаться начальный пояснительный экран в Winlogon.
step:='SmartSizing';
MyRDPWin.AdvancedSettings9.SmartSizing:= not(SmartSizing); //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=EnableWindowsKey;   //0 - отключить, 1 - включить. Указывает, можно ли использовать ключ Windows в удаленном сеансе.
step:='GrabFocusOnConnect';
MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=GrabFocusOnConnect; //0 - отключить, 1 - включить. Указывает, должен ли элемент управления клиента иметь фокус при подключении. Элемент управления не будет пытаться захватить фокус из окна, запущенного в другом процессе.
step:='MinutesToIdleTimeout';
MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= MinutesToIdleTimeout; //По умолчанию 0 только т.е. не отслеживать время. Задает максимальный промежуток времени в минутах, в течение которого клиент должен оставаться подключенным без участия пользователя. Если указанное время истекло, элемент управления вызывает метод IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= OverallConnectionTimeout; //Определяет общую продолжительность времени в секундах, в течение которого клиентский элемент управления ожидает завершения соединения. Максимальное допустимое значение этого свойства - 600, что соответствует 10 минутам. Если указанное время истекает до завершения соединения, элемент управления отключается и вызывает метод IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
MyRDPWin.AdvancedSettings9.RdpPort:=RdpPort;


if NetworkConnectionType=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=BandwidthDetection;          /// если NetworkConnectionType = 0 то....яет, будут ли автоматически обнаружены изменения пропускной способности.
MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end
else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= NetworkConnectionType; // установка типа используемой сети 1- модем, 2-Низкоскоростная широкополосная связь (от 256 Кбит / с до 2 Мбит / с), 3Спутник (от 2 Мбит / с до 16 Мбит / с, с высокой задержкой), 4 -Высокоскоростной широкополосный доступ (от 2 Мбит / с до 10 Мбит / с), 5- Глобальная сеть (WAN) (10 Мбит / с или выше, с высокой задержкой), 6 -Локальная сеть (LAN) (10 Мбит / с или выше)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=PerformanceFlags; //  Задает набор функций, которые можно установить на сервере для повышения производительности.
//Memo1.Lines.add('PerformanceFlags - '+inttostr(PerformanceFlags));
//Memo1.Lines.add('NetworkConnectionType - '+inttostr(NetworkConnectionType));
end;
//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:=MaxReconnectAttempts; //Задает количество попыток переподключения во время автоматического переподключения. Допустимые значения этого свойства - от 0 до 200 включительно.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:= EnableAutoReconnect;       //Указывает, следует ли разрешить клиентскому элементу управления автоматически подключаться к сеансу в случае отключения сети.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=AudioRedirectionMode;  // 0-Перенаправить звуки на клиента. 1-Воспроизведение звуков на удаленном компьютере. 2-Отключить перенаправление звука; не воспроизводить звуки на сервере.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=RedirectDrives;     ///разрешить перенаправлять диски
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=RedirectPrinters;   /// разрешить перенаправлять принтеры

if RedirectClipboard or RedirectPrinters then /// если необходимо перенаправить принтеры и буфер обмена то
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// Установите для этого параметра значение 1, чтобы отключить перенаправление, или 0, чтобы включить перенаправление.

step:='RelativeMouseMode';
MyRDPWin.AdvancedSettings9.RelativeMouseMode:=RelativeMouseMode;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=RedirectClipboard;  /// Указывает, следует ли включить или отключить перенаправление буфера обмена.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=RedirectDevices;    /// Указывает, должны ли перенаправленные устройства быть включены или отключены.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=RedirectPorts;      /// перенаправление портов
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=ConnectToAdministerServer; /// указывает должен ли элемент управления ActiveX пытаться подключиться к серверу в административных целях.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=AudioCaptureRedirectionMode;  ///  указывает, перенаправляется ли устройство аудиовхода по умолчанию от клиента к удаленному сеансу
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=EnableSuperPan;              ///позволяет пользователю легко перемещаться по удаленному рабочему столу в полноэкранном режиме, когда размеры удаленного рабочего стола больше размеров текущего окна клиента. Вместо того, чтобы использовать полосы прокрутки для навигации по рабочему столу, пользователь может указать на границу окна, и удаленный рабочий стол будет автоматически прокручиваться в этом направлении. SuperPan не поддерживает более одного монитора.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=AuthenticationLevel;            ///Задает уровень проверки подлинности для подключения.  (0.1.2)0- Нет аутентификации на сервере. 1- Проверка подлинности сервера требуется и должна успешно завершиться для продолжения подключения. 2- Попытка аутентификации сервера. Если аутентификация не удалась, пользователю будет предложено отменить соединение или продолжить без аутентификации на сервере.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=EnableCredSspSupport;      ///Указывает или извлекает, включен ли для этого подключения  Поставщик службы безопасности учетных данных (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=KeyboardHookMode; // 0-Применяйте комбинации клавиш только локально на клиентском компьютере. 1- Применяйте комбинации клавиш на удаленном сервере. 2- Применяйте комбинации клавиш к удаленному серверу только тогда, когда клиент работает в полноэкранном режиме. Это значение по умолчанию.

PageControl1.ActivePageIndex:=PageControl1.PageCount-1; //// активация последней созданной страницы
//MyRDPWin.StartConnected:=1;   //Установите для этого параметра значение TRUE, если элемент управления должен подключиться сразу при запуске, или FALSE в противном случае.

if AutoConnect then
begin
 ////// если это применить то окно авторизации не выходит  //////// и если проверка на уровне сети то хуй подключишься
//MyRDPWinNoScr:=(MyRDPWin.ControlInterface as IMsRdpClientNonScriptable5);//.Set_AllowPromptingForCredentials(false);
//MyRDPWinNoScr.Set_AllowPromptingForCredentials(false);
/////////////////////////////////////////////////////////////
 MyRDPWin.Connect;
end;
 inc(TabShetCount);
 result:=true;
 except
 on E:Exception do
  begin
   result:=false;
   Memo1.Lines.Add('Ошибка при создании подключения к компьютеру '+NamePC+' - '+e.Message) ;
   Memo1.Lines.Add('Ошибка на этапе настройки - '+step) ;
  end;
end;
end;



procedure TFormRDP.SpeedButton1Click(Sender: TObject); //// подключение по RDP
var
i,z:integer;
TabTrue:Boolean;
begin
try
  ////////////////////// если текщая вкладка и клиент на ней отключен
if PageControl1.TabIndex<>-1 then    /// если есть открытые вкладки
begin
  if PageControl1.ActivePage.ImageIndex<>4 then /// если статус подклчения не равно ok
   begin
   for I :=0 to PageControl1.ActivePage.ComponentCount-1 do
     begin
       if PageControl1.ActivePage.Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Connected<>1 then /// если соединение не установлено
        begin
        ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// применение настроек на открытом клиенте
        (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Connect;
        end;
        break;
       end;
     end;
    exit;
   end;
end;

if (ComboKomp.Text='')  then
begin
ShowMessage('Не указано имя компьютера');
ComboKomp.SetFocus; /// переводим фокус на комбобокс
exit;
end;

if not ping(ComboKomp.Text) then  /// если компьютер не доступен
begin
i:=MessageBox(Self.Handle, PChar('Продолжить подключение?')
        , PChar('Компьютер - '+ComboKomp.Text+' не доступен' ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
end;




///////////////////// ищем вкладку с подключение
TabTrue:=false; /// признак для создания нового подключения
if PageControl1.TabIndex<>-1 then    /// если есть открытые вкладки
begin
for I := 0 to PageControl1.PageCount-1 do   /// смотрим изх в цикле
Begin
 if PageControl1.Pages[i].Caption=ComboKomp.Text then /// сначала ищем, если уже подключен
 begin
  PageControl1.ActivePageIndex:=i; /// если нашли, переходим на вкладку
  for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ищем клиент рдп на вкладке
   begin
     if PageControl1.ActivePage.Components[z] is TMsRdpClient9  then  ///если нашли
     begin
     if (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connected<>1 then /// если соединение не установлено
      begin
      ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// применение настроек на открытом клиенте
      (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connect;
      end;
      TabTrue:=true;
      break; // если нашли и подключили то выходим из цикла
     end;
   end;
 break; /// если нашли вкладку то выходим из цикла.
 end;
End;
end;

if not TabTrue then  /// если подключение в отдельном окне
begin
  for I := 0 to GroupBox8.ControlCount-1 do
    begin
    if GroupBox8.Controls[i] is Tbutton then
    if (GroupBox8.Controls[i]as Tbutton).Hint=ComboKomp.Text then
     begin
     (GroupBox8.Controls[i]as Tbutton).OnClick((GroupBox8.Controls[i]as Tbutton));
      TabTrue:=true;
      break;
     end;
    end;
end;


if  (PageControl1.PageCount<CountTab) and (not TabTrue) then /// сли не нашли подклбчений то создаем новое
 begin
 i:=MessageBox(Self.Handle, PChar('Создать новое подключение?')
 , PChar('Подключение для '+ComboKomp.Text ) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then exit;
 createnewconnectInTab(ComboKomp.Text);
 end
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при подключении к компьютеру '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;

end;



procedure TFormRDP.SpeedButton2Click(Sender: TObject); //// отключение RDp на вкладке
var
i:integer;
begin
try
if PageControl1.TabIndex<>-1 then  /// если есть вкладки
for I :=0 to PageControl1.ActivePage.ComponentCount-1 do
     begin
       if PageControl1.ActivePage.Components[i] is TMsRdpClient9  then
       begin
        if (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Connected<>0 then 
        begin
        (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Disconnect;
        end;
        break;
       end;           
     end;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('Ошибка при отключении от компьютера '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;
end;



procedure TFormRDP.SpeedButton3Click(Sender: TObject); /// добавить вкладку
var
i,z:integer;
TabTrue:Boolean;
begin
if ComboKomp.Text='' then
  begin
  ShowMessage('Не указано имя компьютера');
  ComboKomp.SetFocus; /// переводим фокус на комбобокс
  exit;
  end;
TabTrue:=false;
for I := 0 to PageControl1.PageCount-1 do /// смотрим есть ли уже подключеный клиент на вкладках
begin
  if PageControl1.Pages[i].Caption=ComboKomp.Text then  /// если нашли 
  begin
  PageControl1.ActivePageIndex:=i;  /// перехрдим на вкладку
  for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ищем клиент рдп
    begin
    if PageControl1.ActivePage.Components[z] is TMsRdpClient9  then  /// если нашли
      begin
       if (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connected<>1 then /// если соединение не установлено
       begin
       ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// применение настроек на открытом клиенте
      (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connect;
       end;
      TabTrue:=true;
      break; // если нашли и подключили то выходим из цикла
      end;           
    end;
   break;
   end;
end;

if not TabTrue then  ///смотрим если подключение в отдельном окне
begin
  for I := 0 to GroupBox8.ControlCount-1 do
    begin
    if GroupBox8.Controls[i] is Tbutton then
    if (GroupBox8.Controls[i]as Tbutton).Hint=ComboKomp.Text then
     begin
     (GroupBox8.Controls[i]as Tbutton).OnClick((GroupBox8.Controls[i]as Tbutton));
      TabTrue:=true;
      break;
     end;
    end;
end;


if not ping(ComboKomp.Text) then  /// если компьютер не доступен
begin
i:=MessageBox(Self.Handle, PChar('Продолжить подключение?')
        , PChar('Компьютер - '+ComboKomp.Text+' не доступен' ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
end;
/// так как элементов в массиве только 1000 то больше 1000 вкладок не создавать
if (PageControl1.PageCount<CountTab) and (not TabTrue) then
 begin  //// создаем подключение
 createnewconnectInTab(ComboKomp.Text);
 end;

end;



procedure TFormRDP.SpeedButton4Click(Sender: TObject); /// удалить вкладку
var
z:integer;
begin
try
if PageControl1.PageCount>0 then
 begin
    for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ищем клиент рдп
    begin
    if PageControl1.ActivePage.Components[z] is TMsRdpClient9  then
      begin
      (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Destroy;
      end;
    end;
    PageControl1.ActivePage.Destroy;
 end;
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка удаления вкладки:'+E.Message);
      end;
   end;
end;







procedure TFormRDP.TabSheet1Resize(Sender: TObject); /// изменение размера области просмотра при изменении размера вкладки
var
i:integer;
begin
try

 if (PageControl1.PageCount<1)or(PageControl1.TabIndex=-1) then exit; /// если страниц нет то и менять размер не у чего
 for I := 0 to PageControl1.ActivePage.ComponentCount-1 do //// смотрим компоненты в цикле
begin
 if PageControl1.ActivePage.Components[i] is TMsRdpClient9 then /// если нашли TMsRdpClient9
 begin 
 if(sizeRegion((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopWidth,PageControl1.ActivePage.Width)) or  /// проверяем ширину и высоту окна
 (sizeRegion((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopHeight,PageControl1.ActivePage.Height)) then   /// если отличается делаем reconect
   begin
    (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Reconnect(PageControl1.ActivePage.Width,PageControl1.ActivePage.Height);
   break;
   end;
 {Memo1.Lines.add('Ширина TMsRdpClient9 - '+inttostr((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopWidth));  
 Memo1.Lines.add('Ширина ActivePage - '+inttostr(PageControl1.ActivePage.Width)); 
 Memo1.Lines.add('Высота TMsRdpClient9 - '+inttostr((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopHeight)); 
 Memo1.Lines.add('Высота ActivePage - '+inttostr(PageControl1.ActivePage.Height)); 
 Memo1.Lines.Add('Компьютер - '+(PageControl1.ActivePage.Components[i] as TMsRdpClient9).Server)};
 end;

 
end;
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка при изменении размера страницы :'+E.Message);
  end;
end;
end;

procedure TFormRDP.WMCopyData(var MessageData: TWMCopyData); /// принимаем сообщение
var
ListStr:string;
ListPC:TstringList;
i:integer;

const
 SEND_SETLISTTEXT = 1;  /// Receiv отправить
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv запрос на получение списка
 RECEIVE_REQ_USPASS=3;    ///Receiv запрос пароля и логина
begin
try
  //если пришел список заданная команда совпадает
  if MessageData.CopyDataStruct.dwData = SEND_SETLISTTEXT then
  Begin
    //Устанавливаем текст из полученных данных
    ListStr := PAnsiChar((MessageData.CopyDataStruct.lpData));
    MessageData.Result := 1;
    if ListStr<>'' then
      begin
      ComboKomp.Clear;
      ListPC:=TstringList.create;
      ListPC.CommaText:=ListStr;
      for I := 0 to ListPC.Count-1 do
        begin
        ComboKomp.Items.Add(ListPC.Names[i]);
        ComboKomp.ItemsEx[i].ImageIndex:=strtoint(ListPC.ValueFromIndex[i]);
        end;
      if ComboKomp.Items.Count<>-1 then ComboKomp.ItemIndex:=0;
      ListPC.Free;
      end;
   End;

   if MessageData.CopyDataStruct.dwData = RECEIVE_REQ_USPASS then
  Begin
    //Устанавливаем текст из полученных данных
    ListStr := PAnsiChar((MessageData.CopyDataStruct.lpData));
    MessageData.Result := 1;
    if ListStr<>'' then
      begin
      ListPC:=TstringList.create;
      ListPC.CommaText:=ListStr;
      if ListPC.Count=3 then
        begin
        LabeledEdit1.Text:=ListPC.ValueFromIndex[0];/// user
        LabeledEdit2.Text:=ListPC.ValueFromIndex[1];/// pas
        LabeledEdit3.Text:=ListPC.ValueFromIndex[2]; // dom
        end;
      ListPC.Free;
      end;
   End                                  
    else MessageData.Result := 0;

 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка получения данных из главной программы :'+E.Message);
      end;
   end;
//Memo1.Lines.Add('Строка - '+ListStr);
end;

Function TFormRDP.RequestStr(nameStr:string;Receiv:integer):bool; /// функция запроса списка компов или логина и пароля сообщения из главного приложения
var
  CDS: TCopyDataStruct;
  AllStr:TstringList;
  i:integer;
const
 SEND_SETLISTTEXT = 1;  /// Receiv отправить
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv запрос на получение списка
 RECEIVE_REQ_USPASS=3;    ///Receiv запрос пароля и логина
 begin
 try
//Устанавливаем тип команды
  CDS.dwData := Receiv;
  //Устанавливаем длину передаваемых данных
  CDS.cbData := Length('') + 1;
  //Выделяем память буфера для передачи данных
  GetMem(CDS.lpData, CDS.cbData);
  try
    //Копируем данные в буфер
    StrPCopy(CDS.lpData, AnsiString(''));
    //Отсылаем сообщение в окно с заголовком StringReceiver
    SendMessage(FindWindow(Pchar(nameStr), nil),
                  WM_COPYDATA, Handle, Integer(@CDS));
  finally
    //Высвобождаем буфер
    FreeMem(CDS.lpData, CDS.cbData);
  end;
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка запроса данных из главной программы :'+E.Message);
      end;
   end;
end;




procedure TFormRDP.FormCreate(Sender: TObject);
var
SetIni:Tinifile;
begin
// FormRDP.Width:=Screen.Width;
// FormRDP.Height:=Screen.Height;
 FormRDP.WindowState:=wsMaximized;

 TabShetCount:=0; //// количество созданных элементов. Для добавления индекса создаваемого компонента
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    PageControl1.MultiLine:=SetInI.readBool('MRRDP','multiline',false);
    JvNetscapeSplitter2.Maximized:=SetInI.ReadBool('MRRDP','ViewLog',false);
    try
    TstyleManager.TrySetStyle(SetIni.ReadString('Style','ST','Silver'),True);
    except
    TstyleManager.TrySetStyle('Windows',True);
    end;
    SetInI.Free;
    end;
CountTab:=1000;
CountNewForm:=0;

{RootPatch:=HKEY_CURRENT_USER;
if existreg(stringpatch) then
if regeditread(stringpatch,'isOK')='' then
begin
CountTab:=3;
FormRDP.caption:=FormRDP.caption+' ВНИМАНИЕ!!! НЕ ЗАРЕГИСТРИРОВАННАЯ ВЕРСИЯ, ДОСТУПНО 3 ПОДКЛЮЧЕНИЯ.';
end;}

end;

procedure TFormRDP.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (ssCtrl in Shift) AND (key=38) then    //38- стрелка вверх 40- стрелка вниз
  FormRDP.Align:=alClient;
 if (ssCtrl in Shift) AND (key=40) then    //38- стрелка вверх 40- стрелка вниз
  FormRDP.Align:=alNone;
end;



procedure TFormRDP.FormShow(Sender: TObject);
var
listpc:TstringList;
i:integer;
const
 SEND_SETLISTTEXT = 1;  /// Receiv отправить
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv запрос на получение списка
 RECEIVE_REQ_USPASS=3;    ///Receiv запрос пароля и логина
begin
try
RequestStr('TfrmDomainInfo',2); /// запрос списка компов из главного приложения
RequestStr('TfrmDomainInfo',3); /// запрос логина, пароля и домена из главного приложения
//////////////////////////////////////////////////////////////////////////////
readSettingsInDB('SETDEFAULT'); /// запрос для применения настроек сохраненных по умолчанию
//////////////////////////////////////////////////////////////////////////

 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка загрузки приложения : '+E.Message);
      end;
   end;
end;



procedure TFormRDP.MsRdpClient91AuthenticationWarningDismissed(Sender: TObject);
begin
   ///Вызывается после того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).
end;

procedure TFormRDP.PopupOsherWinPanelPopup(Sender: TObject);  ////при отображении попуп меню на парели в отдельном окне
var                                                           ///
i:integer;
begin
{if (Tpanel(PopupOsherWinPanel.PopupComponent).owner is TForm) then
begin
 with (Tpanel(PopupOsherWinPanel.PopupComponent).owner as TForm) do
 for I := 0 to ComponentCount-1 do
  begin
  if (Components[i] is TMsRdpClient9) then
    begin
    PopupOsherWinPanel.Items[0].Checked:=
    (Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing;
    break;
    end;
  end;
end; }
end;

procedure TFormRDP.N12Click(Sender: TObject); //// попуп меню на парели в отдельном окне, полосы прокрутки
var
namerdp:string;
i:integer;
begin
if (Tpanel(PopupOsherWinPanel.PopupComponent).owner is TForm) then
begin
 with (Tpanel(PopupOsherWinPanel.PopupComponent).owner as TForm) do
 for I := 0 to ComponentCount-1 do
  begin
  if (Components[i] is TMsRdpClient9) then
    begin
    (Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing:=
    not (Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing;
   // PopupOsherWinPanel.Items[0].Checked:= not
   // (Components[i] as TMsRdpClient9).AdvancedSettings9.SmartSizing;
    break;
    end;
  end;
end;
end;

procedure TFormRDP.N1Click(Sender: TObject); /// попуп меню, подключить на вкладке
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;

//if PageControl1.TabIndex<>-1 then    /// если есть открытые вкладки
begin
  if PageControl1.Pages[tabInd].ImageIndex<>4 then /// если статус подклчения не равно ok
   begin
   for I :=0 to PageControl1.Pages[tabInd].ComponentCount-1 do
     begin
       if PageControl1.Pages[tabInd].Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Connected<>1 then /// если соединение не установлено
        begin
         //PageControl1.TabIndex:=tabInd;
         ConnectToCurentSetting(PageControl1.Pages[tabInd].caption,tabInd); /// применение настроек на открытом клиенте
        (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Connect;
        end;
        break;
       end;
     end;
    exit;
   end;
end;
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка подключения :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N2Click(Sender: TObject); //// попуп меню отключение клиента
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;
//if PageControl1.TabIndex<>-1 then    /// если есть открытые вкладки
begin
  if PageControl1.Pages[tabInd].ImageIndex<>5 then /// если статус подклчения не равно отключено
   begin
   for I :=0 to PageControl1.Pages[tabInd].ComponentCount-1 do
     begin
       if PageControl1.Pages[tabInd].Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Connected<>0 then /// если соединение  установлено (1) или еще подключается (2)
        //if PageControl1.Pages[tabInd].ImageIndex<>5 then /// если соединение не установлено, или находится в стадии установки
        begin
         //PageControl1.TabIndex:=tabInd;
        (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Disconnect;
        end;
        break;
       end;
     end;
    exit;
   end;
end;
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка отключения :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N3Click(Sender: TObject); /// попуп меню, закрыть другие вкладки
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;
PageControl1.TabIndex:=tabInd;
PageControl1.ActivePage.PageIndex:=0;
for i := PageControl1.PageCount - 1 downto 1 do
 begin
 if PageControl1.Pages[i].Caption<>tabCapt then
 PageControl1.Pages[i].Destroy;
 end;
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка закрытия подключений :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N4Click(Sender: TObject);  ///попуп меню, закрыть текушее подключение
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;
PageControl1.Pages[tabInd].Destroy;

Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка закрытия подключения :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N5Click(Sender: TObject);  /// попуп  в новом окне
var
tabInd,i:integer;
tabCapt:string;
begin
if (CountTab=3) and  (GroupBox8.ControlCount>=1) then
begin
ShowMessage('Больше подключений в зарегистрированной версии программы.');
exit;
end;

tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
OtherWinForRDPClient(tabInd);
end;

procedure TFormRDP.N6Click(Sender: TObject);
var
SetInI:Tinifile;
begin
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    SetInI.WriteBool('MRRDP','multiline',not PageControl1.MultiLine);
    SetInI.Free;
    showmessage('Перезапустите клиента для применения настроек');
    end;
end;

procedure TFormRDP.N7Click(Sender: TObject); // скрыть /отобразить лог
var
SetInI:Tinifile;
begin
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    if SetInI.readBool('MRRDP','ViewLog',true) then // если лог скрытый
    begin
    SetInI.WriteBool('MRRDP','ViewLog',false);
    SetInI.Free;
    exit;
    end;
    if SetInI.readBool('MRRDP','ViewLog',true)=false then /// если лог отображается
    begin
    SetInI.WriteBool('MRRDP','ViewLog',true);
    JvNetscapeSplitter2.Maximized:=true; // скрыть лог
    end;
    SetInI.Free;
    end;
end;



procedure TFormRDP.PopupLogPopup(Sender: TObject);
var
SetInI:Tinifile;
begin
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    if SetInI.readBool('MRRDP','ViewLog',true) then
    PopupLog.Items[0].Caption:='Отображать лог';

    if SetInI.readBool('MRRDP','ViewLog',true)=false then
    PopupLog.Items[0].Caption:='Скрывыть лог';

    SetInI.Free;
    end;
end;

procedure TFormRDP.NewAuthenticationWarningDisplayed(Sender: TObject);
begin
 //Вызывается до того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).

end;

procedure TFormRDP.MsRdpClient91AutoReconnected(Sender: TObject);
begin
 //Вызывается, когда клиентский элемент управления автоматически повторно подключается к удаленному сеансу.
end;

procedure TFormRDP.MsRdpClient91AutoReconnecting(ASender: TObject;
  disconnectReason, attemptCount: Integer);  //Вызывается, когда клиент находится в процессе автоматического повторного подключения сеанса к серверу узла сеансов удаленных рабочих столов (RD Session Host).
begin
//Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Переподключение, попытка - '+inttostr(disconnectReason)+'/ Причина - '+ SysErrorMessage(attemptCount));
end;

procedure TFormRDP.MsRdpClient91AutoReconnecting2(ASender: TObject;   //Вызывается, когда клиент находится в процессе автоматического повторного подключения сеанса к серверу узла сеансов удаленных рабочих столов (RD Session Host).
  disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
  maxAttemptCount: Integer);
var
errorstr:string;
id2:integer;
 function NetW(s:bool):string;
 begin
   if s then result:='подключена'
   else result:='отключена';
 end;
begin
errorstr:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Переподключение:  '+errorstr+
'/ Состояние сети :'+netW(networkAvailable)+' / Попытка - '+inttostr(attemptCount)+' из '+inttostr(maxAttemptCount));
end;


procedure TFormRDP.MsRdpClient91DevicesButtonPressed(Sender: TObject);
begin
////вроде как нажати кнопок на клаве
end;



procedure TFormRDP.MsRdpClient91ServiceMessageReceived(ASender: TObject;
  const serviceMessage: WideString);
begin
Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP : '+serviceMessage);
end;


procedure TFormRDP.NewRdpAuthenticationWarningDismissed(Sender: TObject);
begin  //Вызывается после того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).
try
 if not (sender is TMsRdpClient9) then exit;

{if  ((sender as TMsRdpClient9).Connected=0)and   /// tckb соединение не установлено
 ((sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport=false) then  // и отключена  EnableCredSspSupport
 begin
   (sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport:=true;
   (sender as TMsRdpClient9).Connect;
 end;}


Memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - Предупреждение Аутентификации');
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - Предупреждение Аутентификации." :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpAuthenticationWarningDisplayed(Sender: TObject);
 const  //Вызывается до того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).
  AutType: array [0..4] of string =('No authentication is used.',
  'Используется проверка подлинности сертификата','Используется проверка подлинности Kerberos',
  'Используется оба типа проверки подлинности, сертификата и по протоколу','неизвестное состояние');
begin
try
  if not (sender is TMsRdpClient9) then  exit;
 if ((sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType<5) then
 memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - Проверка подлинности. - '
 + AutType[(sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType]);

 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - Проверка подлинности." :'+E.Message);
      end;
   end;
end;





procedure TFormRDP.NewRdpAutoReconnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
Memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - Переподключение к клиентсому компьютеру.- ');
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - Переподключение к клиентсому компьютеру" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpAutoReconnecting(ASender: TObject; disconnectReason,
  attemptCount: Integer);
  var
strerror:string;
id2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
strerror:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - Автоматическое переподключение сеанса удаленных рабочих столов: '+strerror);
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - Автоматическое переподключение сеанса удаленных рабочих столов" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpAutoReconnecting2(ASender: TObject;
  disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
  maxAttemptCount: Integer);
var
strerror:string;
id2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
strerror:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - клиентский элемент управления переподключается к удаленному сеансу :'+strerror);
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - клиентский элемент управления переподключается к удаленному сеансу" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpChannelReceivedData(ASender: TObject; const chanName,
  data: WideString);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - клиент получает данные на виртуальном канале с возможностью сценария');
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - клиент получает данные на виртуальном канале с возможностью сценария" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConfirmClose(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - Завершение сеанса служб удаленного рабочего стола');
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - Завершение сеанса служб удаленного рабочего стола" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// после ввода пароля
memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - установка соединения...');
Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - установка соединения" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConnecting(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// сдесьперед вводом пароля
memo1.Lines.Add((sender as TMsRdpClient9).Server+
' : RDP - инициализация соединения...');
for I := 0 to Groupbox8.ControlCount-1 do
begin
  if GroupBox8.Controls[i] is TButton then
  if GroupBox8.Controls[i].Owner.Name=(sender as TMsRdpClient9).Owner.Name then
   begin
    (GroupBox8.Controls[i] as Tbutton).ImageIndex:=7;
   end;
end;

Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка "RDP - инициализация соединения" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpDisconnected(ASender: TObject; discReason: Integer);
var
errorstr:string;
disconnectReason:integer;
i:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
errorstr:=(Asender as TMsRdpClient9).GetErrorDescription(disconnectReason,discReason);
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Завершение сеанса - '+errorstr);
if TComponent(Asender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
TTabSheet(TComponent(Asender).Owner).ImageIndex:=5;
for I := 0 to Groupbox8.ControlCount-1 do
begin
  if GroupBox8.Controls[i] is TButton then
  if GroupBox8.Controls[i].Owner.Name=(Asender as TMsRdpClient9).Owner.Name then
   begin
    (GroupBox8.Controls[i] as Tbutton).ImageIndex:=5;
   end;
end;

Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка обновления статуса :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpExit(Sender: TObject);
begin
try
{if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - фокус перемещен'); }

Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка "RDP - фокус перемещен" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpFatalError(ASender: TObject; errorCode: Integer);
const
StrError : array [0..7] of string=('Произошла неизвестная ошибка.','Код внутренней ошибки 1.'
,'Произошла ошибка нехватки памяти.','Произошла ошибка создания окна.','Код внутренней ошибки 2.'
,'Внутренний код ошибки 3','Код внутренней ошибки 4.','Во время подключения к клиенту произошла неисправимая ошибка.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if errorCode=100 then memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - произошщла неустранимая ошибка - Ошибка инициализации Winsock.')
else
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP - произошщла неустранимая ошибка - '+StrError[errorCode]);
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка "RDP - произошщла неустранимая ошибка" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpLoginComplete(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - успешный вход в систему');

if TComponent(sender).Owner is TTabSheet then   /// если родитель рдп клиента TabSheet а не отдельное окно
begin
TTabSheet(TComponent(sender).Owner).ImageIndex:=4;
DataMod.WriteItemDB((sender as TMsRdpClient9).Server); /// если удачно подключились записываем настройки подключения в базу
end;

for I := 0 to Groupbox8.ControlCount-1 do
begin
  if GroupBox8.Controls[i] is TButton then
  if GroupBox8.Controls[i].Owner.Name=(sender as TMsRdpClient9).Owner.Name then
   begin
    (GroupBox8.Controls[i] as Tbutton).ImageIndex:=4;
   end;
end;

Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка обновления статуса :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpLogonError(ASender: TObject; lError: Integer);
var
StatusText,ErrorDesc:string;
ID2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
StatusText:=(Asender as TMsRdpClient9).GetStatusText(lError);
ErrorDesc:=(Asender as TMsRdpClient9).GetErrorDescription(ID2,lError);
memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - StatusText - '+ StatusText);
memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - ErrorDesc - '+ ErrorDesc);

memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - ошибка входа в систему - '+ SysErrorMessage(lError));
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка "ошибка входа в систему" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpNetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
const
qLevel: array [0..5] of string =('Скорость не известна...','Скорость менее 512 Кбит/с',
'Скорость от 512 Кбит/с до 2 Мбит/с.','Скорость от 2 до 10 Мбит/с.'
,'Скорость больше 10 Мбит/с.','Скорость больше 100 Мбит/с.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - произошло изменение состояния сети');
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : '+qLevel[qualityLevel]
+' / Пропускная способность (Bandwidth) - '+inttostr(bandwidth)+' Кбит/с / Задержка соединения (RTT) - '+inttostr(rtt)+' мс.')
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка при чтении характеристик изменения состояния сети :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpWarning(ASender: TObject; warningCode: Integer);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if warningCode=1 then
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP -  Кэш растрового изображения поврежден.');
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка при оповещении OnWarning :'+E.Message);
  end;
end;
end;





procedure TFormRDP.PageControl1Change(Sender: TObject);
var
i:integer;
begin
try
if PageControl1.PageCount>CountTab then
begin
 PageControl1.Pages[PageControl1.PageCount-1].Destroy;
 exit;
end;

if PageControl1.ActivePage.ImageIndex=5 then  /// если клиент отключен, то читаем настройки из бд
begin
if not readSettingsInDB(PageControl1.ActivePage.Caption)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
readSettingsInDB('SETDEFAULT');
end;
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка при переключении :'+E.Message);
  end;
end;
end;





procedure TFormRDP.PageControl1DragDrop(Sender, Source: TObject; X, Y: Integer); /// для перемещения вкладок
const
  TCM_GETITEMRECT = $130A;
var
  i: Integer;
  r: TRect;
begin
try
if PageControl1.PageCount=0 then exit;
/// можно так
 {if not (Sender is TPageControl) then Exit;
  with sender as TPageControl do
  begin
    for i := 0 to PageCount - 1 do
    begin
      Perform(TCM_GETITEMRECT, i, lParam(@r));
      if PtInRect(r, Point(X, Y)) then
      begin
        if i <> ActivePage.PageIndex then
          ActivePage.PageIndex := i;
        Exit;
      end;
    end;
  end;}
  // или так


if (Source is TPageControl) and ((PageControl1.ActivePageIndex<>PageControl1.IndexOfTabAt(x,y))) then
   begin
    PageControl1.ActivePage.PageIndex:=PageControl1.IndexOfTabAt(x,y);
    PageControl1.EndDrag(true);
   // PageControl1.tag:=0;
   end;
  PageControl1Change(nil);
  Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка PageControlDragDrop :'+E.Message);
  end;
end;
end;

procedure TFormRDP.PageControl1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
try
if (Button=mbRight) then
begin
  PMouse.X:=x;
  PMouse.Y:=y;
end;

if PageControl1.PageCount=1 then exit;
if (Button=mbLeft) and (PageControl1.tag=0) then
begin
PageControl1.BeginDrag(false);
PageControl1.tag:=1;
end;
  Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка PageControlMouseDown :'+E.Message);
  end;
end;
end;

procedure TFormRDP.PageControl1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
try
  PageControl1.tag:=0;
  Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка PageControlMouseUp :'+E.Message);
  end;
end;
end;

procedure TFormRDP.PageControl1DragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
try
Accept:=(Source is TPageControl);
  Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка PageControl1DragOver :'+E.Message);
  end;
end;
end;



procedure TFormRDP.Button2Click(Sender: TObject); //// сохранение списка подключений
var
i:integer;
ListIni:Tinifile;
NewDescription:string;
begin
try
if PageControl1.PageCount<=0 then
begin  
  showmessage('Нет подключений');
  exit;
end;

if not InputQuery('Описание списка компьютеров', 'Описание', NewDescription)
    then exit;
if NewDescription='' then NewDescription:=DateTimeToStr(now);

ListIni:=TIniFile.Create(ExtractFileDir(Application.ExeName)+'\RDPlist.dat');
  for I :=0 to PageControl1.PageCount-1 do
  begin
   ListIni.WriteString(NewDescription,PageControl1.Pages[i].Caption,inttostr(i));
  end;

for i:=0 to GroupBox8.ControlCount-1 do
 begin
 if  GroupBox8.Controls[i] is Tbutton then
  begin
  ListIni.WriteString(NewDescription,(GroupBox8.Controls[i] as Tbutton).Hint,inttostr(i));
  end;
 end;
ListIni.Free
Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка сохранения списка подключений :'+E.Message);
  end;
end;
end;


procedure TFormRDP.Button3Click(Sender: TObject);
begin
FormLoadList.Show;
end;

procedure TFormRDP.Button4Click(Sender: TObject);
var   //// обновление статуса всех компьютеров при раскрытии списка
  CDS: TCopyDataStruct;
  AllStr:TstringList;
  i:integer;
  const
 SEND_SETLISTTEXT = 1;  /// отправить
  RECEIVE_REQ_LISTTEXT=2;  /// запрос на получение списка
begin
try
//Устанавливаем тип команды
  CDS.dwData := RECEIVE_REQ_LISTTEXT;
  //Устанавливаем длину передаваемых данных
  CDS.cbData := Length('') + 1;
  //Выделяем память буфера для передачи данных
  GetMem(CDS.lpData, CDS.cbData);
  try
    //Копируем данные в буфер
    StrPCopy(CDS.lpData, AnsiString(''));
    //Отсылаем сообщение в окно с заголовком StringReceiver
   // SendMessage(GetWindow(FindWindow('TfrmDomainInfo', nil),GW_OWNER),
   // WM_COPYDATA, Handle, Integer(@CDS));
   SendMessage(FindWindow('TfrmDomainInfo', nil),
    WM_COPYDATA, Handle, Integer(@CDS));
  finally
    //Высвобождаем буфер
    FreeMem(CDS.lpData, CDS.cbData);
  end;

Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка обновления статуса компьютеров:'+E.Message);
      end;
   end;
end;


procedure TFormRDP.Button5Click(Sender: TObject);
begin
Memo1.Lines.Add('Groupbox8 - '+inttostr(Groupbox8.controlcount));
Memo1.Lines.Add('CountTab - '+inttostr(CountTab));
//if not readSettingsInDB(ComboKomp.Text)then /// если нет настроек для этого компа, то читаем настройки по умолчанию
//readSettingsInDB('SETDEFAULT');
end;


procedure TFormRDP.Button6Click(Sender: TObject);
var
i:integer;
begin
i:=MessageBox(Self.Handle, PChar('Установить текущие настройки по умолчанию '+#10#13+' для всех новых подключений?')
        , PChar('Настройки для RDP подключений' ) ,MB_YESNO+MB_ICONQUESTION);
if i=IDYES then
 begin
 if not DataMod.writesetDB('SETDEFAULT', //- имя компа
 bitcolorDepth(ComboBoxColorDepth.ItemIndex), // количество бит на пиксель
 booltoint(CheckBoxBitmapPeristence.Checked),//BitmapPeristence
 booltoint(CheckboxCachePersistenceActive.Checked), //CachePersistenceActive
 SpinBitmapCacheSize.Value,                //BitmapCacheSize
 SpinEditCache16BppSize.Value,           //VirtualCache16BppSize
 SpinEditCache32BppSize.Value,          //VirtualCache32BppSize
 SpinBitmapVirtualCacheSize.Value,             //VirtualCacheSize
 booltoint(CheckBoxDisableCtrlAltDel.Checked),   //DisableCtrlAltDel
 booltoint(CheckBoxDoubleClickDetect.Checked),   // DoubleClickDetect
 booltoint(CheckBoxEnableWindowsKey.Checked),   // EnableWindowsKey
 SpinMinutesToIdleTimeout.Value,            //MinutesToIdleTimeout
 SpinOverallConnectionTimeout.Value,       // OverallConnectionTimeout
 strtoint(LabeledEditRdpPort.Text),      //RdpPort
 CalcPerformanceFlags(''),               //PerformanceFlags
 ComboBoxNetworkConnectionType.ItemIndex, // NetworkConnectionType
 SpinMaxReconnectAttempts.Value,           // MaxReconnectAttempts
 ComboAudioRedirectionMode.ItemIndex,      // AudioRedirectionMode
 ComboBoxAuthLevel.ItemIndex,             //  AuthenticationLevel
 ComboKey.ItemIndex,                        // KeyboardHookMode
 CheckBoxSmartSizing.Checked,         //BoxSmartSizing
 checkboxGrabFocusOnConnect.Checked,            //GrabFocusOnConnect
 true,                                          //BandwidthDetection
 CheckEnableAutoReconnect.Checked,                  //EnableAutoReconnect
 CheckBoxDisk.Checked,                       //RedirectDrives
 CheckBoxPrint.Checked,                       //RedirectPrinters
 CheckBoxMouseMode.Checked,               //RelativeMouseMode
 CheckBoxClipboard.Checked,                   //RedirectClipboard
 CheckRedirectDevices.Checked,              //RedirectDevices
 CheckBoxPort.Checked,                       //RedirectPorts
 CheckBoxConToAdmSrv.Checked,                  //ConnectToAdministerServer
 CheckBoxRecAudio.Checked,                 //AudioCaptureRedirectionMode
 CheckBoxEnSuperPan.Checked,               //EnableSuperPan
 CheckBoxCredSsp.Checked                    //EnableCredSspSupport
 ) then
 ShowMessage('Ошибка записи настроек по умолчанию');
 end;
end;


procedure TFormRDP.Button7Click(Sender: TObject);
begin
DataMod.clearRdpTabl;
end;



procedure TFormRDP.ButtonUpdateClick(Sender: TObject);
var
i:integer;
begin
if PageControl1.PageCount=0 then exit;
for I := 0 to PageControl1.ActivePage.ComponentCount-1 do
begin
  if (PageControl1.ActivePage.Components[i] is TMsRdpClient9) then
  begin
    (PageControl1.ActivePage.Components[i] as TMsRdpClient9).UpdateSessionDisplaySettings
    (PageControl1.ActivePage.Width-50,PageControl1.ActivePage.Height-50, //  Ширина рабочего стола. и Высота рабочего стола.
    PageControl1.ActivePage.Width-10,PageControl1.ActivePage.Height-10,  // Физическая ширина. и Физическая высота.
    2, // Ориентация рабочего стола
    100,  //Коэффициент масштабирования рабочего стола.
    100); //Коэффициент масштабирования устройства
  end;

end;


end;


procedure TFormRDP.SpeedButton4MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
SpeedButton4.Hint:='Удалить подключение '+ PageControl1.ActivePage.Caption;
end;


procedure TFormRDP.SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
SpeedButton3.Hint:='Добавить новое подключение с '+ ComboKomp.Text;
end;

procedure TFormRDP.Button1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
Button1.Hint:='Рабочий стол по размеру окна, для '+ PageControl1.ActivePage.Caption;
end;



procedure TFormRDP.SpeedButton1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
SpeedButton1.Hint:='Подключится к '+Combokomp.Text;
end;

procedure TFormRDP.SpeedButton2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
SpeedButton2.Hint:='Закрыть соединение '+ PageControl1.ActivePage.Caption;
end;



initialization  /// инициализация дополнения к стилям, а точнее кнопки закрыть на вкладке
TStyleManager.Engine.RegisterStyleHook(TCustomTabControl, TTabControlStyleHookBtnClose);
TStyleManager.Engine.RegisterStyleHook(TTabControl, TTabControlStyleHookBtnClose);

    
end.
