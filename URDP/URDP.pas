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

 type /////�������� ��� ������ �������
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
    function loadlistpc(sFile:string):Tstringlist; // �������� ������������ ������ ����������

    function FlagsChecked (x:bool):bool;
    function FlagsEnabled(x:bool):bool;
    function readSettingsInDB(s:string):bool; /// ������ �������� �� ��
    //function writesettingToDB(s:string):bool; // ������ �������� � ��
    function ConnectToCurentSetting(NamePC:string;IndexTab:integer):boolean; // ����������� �������� ��� ���������� RDP ������� � ���������� �����������
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

  HM,HML: THandle; //// ������� ��� ����� ���������� ������ ��� ������ ������������� ����������
  RootPatch: HKEY;
  TabShetCount,CountNewForm:integer;// ���������� ������� � ���� � ��������
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
      Memo1.Lines.Add('��������');
      BorderStyle:=bsSingle;
      end;
    SC_RESTORE:
      begin
      Memo1.Lines.Add('����������');
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


function TFormRDP.createnewconnectInTab(namepc:string):bool; // ����� ������� ���� ������� ���� ��������� ��������� ��� ���������� �� ���� ������
begin
try
  if createRDP(namepc, //- ��� �����
 labelededit3.Text,               // - �����
 labelededit1.Text,               //������������
 labelededit2.Text,               // ������
 true,                            /// ������������ ����� ��� ��������
 bitcolorDepth(ComboBoxColorDepth.ItemIndex), // ���������� ��� �� �������
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
   Memo1.Lines.Add('������ �������� ����������� '+namePC+' - '+e.Message)
  end;
end;
 end;



procedure TFormRDP.findComponent; ///frmDomainInfo.findElement
var
i:integer;
CurentNamePC:string;
begin
CountTab:=3;
FormRDP.caption:=FormRDP.caption+' ��������!!! �� ������������������ ������, �������� 3 �����������.';
end;

function regeditread(patch,Section:string):string;
var      /// ������ �� �������
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

function existreg(patch:string):bool;  //// ������������� ���� � �������
var
regFile:TregInifile;
begin
regFile:=Treginifile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then result:=true
else result :=false;
end;

procedure TFormRDP.ButtonOnForForm(Sender: TObject); //// ����������� �� RDP �� ��������� �����
var
i:integer;
namePC:string;
LevelAut:integer;
begin
try
LevelAut:=2; /// �������������
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
   Memo1.Lines.Add('������ ��� ����������� � ���������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonOffForForm(Sender: TObject); //// ���������� �� RDP �� ��������� �����
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
   Memo1.Lines.Add('������ ��� ���������� �� ���������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonResizeForForm(Sender: TObject); //// �� ���� �����  RDP �� ��������� �����
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
        (Components[i] as TMsRdpClient9).Width:= Width; /// ������ �� ������ � ������ �����
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
   Memo1.Lines.Add('������ ��� ��������������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonCloseForForm(Sender: TObject); //// ������� ��������� �����
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
 ((TComponent(sender).Owner as TPanel).Owner as Tform).close;
except
 on E:Exception do
  begin
   Memo1.Lines.Add('������ ��� �������� �����������  - '+e.Message)
  end;
end;
end;


procedure TFormRDP.ShowButtonOnClick(Sender: TObject); /// ������ �� ������ GroupBox8 ��� �������������� � ������������ ����
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
   Memo1.Lines.Add('������ ��� ������������/�������������� ���� - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonMinimizeForForm(Sender: TObject); //// �������� ��������� �����
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
   Memo1.Lines.Add('������ ��� ������������ ���� - '+e.Message)
  end;
end;
end;

procedure TFormRDP.ButtonMaxNormalForForm(Sender: TObject); //// ���������� ��� �������� ��������� �����
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
   Memo1.Lines.Add('������ ��� ��������� ������� ���� - '+e.Message)
  end;
end;
end;

procedure TFormRDP.OtherWinForRDPClientClose(Sender: TObject; var Action: TCloseAction); /// ����������� ����� ��� ��������
begin
if Sender is Tform then
(Sender as TForm).Release;
end;

procedure TFormRDP.PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
begin
     movingPan:=true;
     x0:=x;
     y0:=y;
end;

procedure TFormRDP.PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// ����������� ������ ��  �����
  Y: Integer);
begin
     if movingPan then
   begin
     (sender as tpanel).Left:=(sender as tpanel).Left+x-x0;
     (sender as tpanel).Top:=(sender as tpanel).Top+y-y0;
   end;
end;

procedure TFormRDP.PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
begin
  movingPan := false;
end;

procedure TFormRDP.ButonMoveWinMouseDown(Sender: TObject;  /// ����������� ����� �� ������
                           Button: TMouseButton;
                            Shift: TShiftState;
                             X, Y: Integer);
var
FrmHND:Thandle;
begin
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
 FrmHND:=((TComponent(sender).Owner as TPanel).Owner as Tform).Handle; // �������� ����� ����� ��� �����������
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
groupbox8.Caption:='����������� ��� '+(sender as Tbutton).Hint;
end;


procedure TFormRDP.ShowBMouseLeave(Sender: TObject);
begin
if not (sender is Tbutton) then exit;
groupbox8.Caption:='����������� � ��������� ����';
end;

procedure TFormRDP.N8Click(Sender: TObject); /// �����, ��������� ����������� � ��������� ����
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
 if not readSettingsInDB(namerdp)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
 readSettingsInDB('SETDEFAULT');
 createnewconnectInTab(namerdp); // �������� �����������
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
PanelForBut.Hint:='����������� � '+NameCapt+'. ������� ��� �����������';
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
ButFP.Hint:='����������� � '+NameCapt;
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
ButFP.Hint:='���������� �� '+NameCapt;
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
ButFP.Hint:='�������� �� ������� ������';
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
AutLevel.Hint:='�������� ����������� (CredSSP) ��� '+NameCapt;
AutLevel.ShowHint:=true;
AutLevel.Items.Add('������������ ��� ��������������');
AutLevel.Items.Add('�� ���������');
AutLevel.Items.Add('�������������');
AutLevel.ItemIndex:=2;
AutLevel.Left:=(ButFP.Width*3)+20;
AutLevel.TabOrder:=4;

ButFP:=TButton.Create(PanelForBut); /// �������������� �����
ButFP.Parent:=PanelForBut;
ButFP.Name:='MoveWin'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='����������� �����';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=12;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*3)+25+AutLevel.Width;
ButFP.OnMouseDown:=ButonMoveWinMouseDown;

ButFP:=TButton.Create(PanelForBut); /// ������������ �����
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPCollapse'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='��������';
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=11;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*4)+30+AutLevel.Width;
ButFP.OnClick:= ButtonMinimizeForForm;  //

ButFP:=TButton.Create(PanelForBut); /// ���������� ��� ��������
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPCollapseWin'+inttostr(CountNewForm);
ButFP.Width:=23;
ButFP.Height:=23;
ButFp.Caption:='';
ButFP.Hint:='�������� � ����';
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
ButFP.Hint:='������� ����������� � '+NameCapt;
ButFp.ShowHint:=true;
ButFP.ImageAlignment:=iaCenter;
ButFP.Images:=ImageButton;
ButFP.ImageIndex:=9;
ButFP.Top:=4;
ButFP.Left:=(ButFP.Width*6)+40+AutLevel.Width;
ButFP.OnClick:=ButtonCloseForForm;

if GroupBox8.Caption='' then GroupBox8.Caption:='����������� � ��������� ����';
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
MyRDPWin.SendToBack;/// ��� ���� ����� ���� ����� ������

if not readSettingsInDB(NameCapt)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
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
 MyRDPWin.ConnectingText:='��������� ���������� c '+NameCapt+'...';
 MyRDPWin.DisconnectedText:='������ '+NameCapt+' ��������';
 MyRDPWin.Server:=Widestring(NameCapt);
 if Domain<>'' then  MyRDPWin.Domain:=WideString(Domain)
 else MyRDPWin.Domain:=LabeledEdit3.Text;
 if UserName<>'' then MyRDPWin.UserName:=UserName
 else MyRDPWin.UserName:='';



step:='';
MyRDPWin.ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex);
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=booltoint(CheckBoxBitmapPeristence.Checked); // 0 ���������,1 ��������.  ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='CachePersistenceActive';
MyRDPWin.AdvancedSettings9.CachePersistenceActive:=booltoint(CheckboxCachePersistenceActive.Checked); // 0 - ���������, 1 - ��������. ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='BitmapCacheSize';
MyRDPWin.AdvancedSettings9.BitmapCacheSize:=SpinBitmapCacheSize.Value;; // �� 1 �� 32 . ������ ����� ���� ��������� ����������� � ����������, ������������� ��� ��������� ����������� 8 ��� �� �������
step:='VirtualCache16BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=SpinEditCache16BppSize.Value; // ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCache32BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=SpinEditCache32BppSize.Value;// ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCacheSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=SpinBitmapVirtualCacheSize.Value;  //������ ������ ����� ����������� ���� ���������� ����������� � ����������, ������� ������������ ��� ����� 8 ��� �� �������. ���������� �������� �������� ����� �������� - �� 1 �� 32 ������������. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��. ��������� �������� �������� �������� BitmapVirtualCache16BppSize � BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=booltoint(CheckBoxDisableCtrlAltDel.Checked); // 0 - ���������, 1 - ��������.  ���������, ������ �� ������������ ��������� ������������� ����� � Winlogon.
step:='SmartSizing';
MyRDPWin.AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked); //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=booltoint(CheckBoxEnableWindowsKey.Checked);   //0 - ���������, 1 - ��������. ���������, ����� �� ������������ ���� Windows � ��������� ������.
step:='GrabFocusOnConnect';
MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked; //0 - ���������, 1 - ��������. ���������, ������ �� ������� ���������� ������� ����� ����� ��� �����������. ������� ���������� �� ����� �������� ��������� ����� �� ����, ����������� � ������ ��������.
step:='MinutesToIdleTimeout';
MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= SpinMinutesToIdleTimeout.Value;  //�� ��������� 0 ������ �.�. �� ����������� �����. ������ ������������ ���������� ������� � �������, � ������� �������� ������ ������ ���������� ������������ ��� ������� ������������. ���� ��������� ����� �������, ������� ���������� �������� ����� IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= SpinOverallConnectionTimeout.Value; //���������� ����� ����������������� ������� � ��������, � ������� �������� ���������� ������� ���������� ������� ���������� ����������. ������������ ���������� �������� ����� �������� - 600, ��� ������������� 10 �������. ���� ��������� ����� �������� �� ���������� ����������, ������� ���������� ����������� � �������� ����� IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
MyRDPWin.AdvancedSettings9.RdpPort:=strtoint(LabeledEditRdpPort.Text);


{MyRDPWinNoScr:=(MyRDPWin.ControlInterface as IMsRdpClientNonScriptable5);//.Set_AllowPromptingForCredentials(false);
//MyRDPWinNoScr.Set_AllowPromptingForCredentials(false); ////// ���� ��� ��������� �� ���� ����������� �� �������  //////// � ���� �������� �� ������ ���� �� ��� ������������
MyRDPWinNoScr.get_RemoteMonitorCount(CountMonitor); /// ���������� ���������� ��������� ���������
MyRDPWinNoScr.get_RemoteMonitorLayoutMatchesLocal(UseRemotMon); //����������, ��������� �� ����� ���������� �������� ��������� ���������� ��������. ���� ��� �������� �������� VARIANT_TRUE , ���������� ���������� �������� ��������� ���������� ���������� ��������. ���� ��� �������� �������� VARIANT_FALSE , ��������� � ��������� �������� ����� ������ ������.
 //MyRDPWinNoScr.Get_UseMultimon(UseSecondMon); /// �����������, ������������ �� ��������� ���������
 //MyRDPWinNoScr.Set_UseMultimon(True); /// ������������� ������������� ���������� ���������
   showmessage('���������� ��������� ��������� - '+inttostr(CountMonitor));

  z:=MessageBox(Self.Handle, PChar('������������ ��� �������� ��� �����������?')
        , PChar('��������� - '+server ) ,MB_YESNO+MB_ICONQUESTION);
   if z=IDYES then
   begin
    if UseRemotMon then ShowMessage('����� ��������� ���������') else ShowMessage('����� ��������� �� ���������');
    MyRDPWinNoScr.Set_UseMultimon(True); /// ������������� ��������� ���������
   end; }

  if ComboBoxNetworkConnectionType.ItemIndex=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=true;          /// ���� NetworkConnectionType = 0 ��....���, ����� �� ������������� ���������� ��������� ���������� �����������.
//MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end
else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // ��������� ���� ������������ ���� 1- �����, 2-��������������� �������������� ����� (�� 256 ���� / � �� 2 ���� / �), 3������� (�� 2 ���� / � �� 16 ���� / �, � ������� ���������), 4 -���������������� �������������� ������ (�� 2 ���� / � �� 10 ���� / �), 5- ���������� ���� (WAN) (10 ���� / � ��� ����, � ������� ���������), 6 -��������� ���� (LAN) (10 ���� / � ��� ����)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  ������ ����� �������, ������� ����� ���������� �� ������� ��� ��������� ������������������.
end;
//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:= SpinMaxReconnectAttempts.Value;   //������ ���������� ������� ��������������� �� ����� ��������������� ���������������. ���������� �������� ����� �������� - �� 0 �� 200 ������������.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:= CheckEnableAutoReconnect.Checked;       //���������, ������� �� ��������� ����������� �������� ���������� ������������� ������������ � ������ � ������ ���������� ����.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=ComboAudioRedirectionMode.ItemIndex;  // 0-������������� ����� �� �������. 1-��������������� ������ �� ��������� ����������. 2-��������� ��������������� �����; �� �������������� ����� �� �������.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=CheckBoxDisk.Checked;     ///��������� �������������� �����
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=CheckBoxPrint.Checked;    /// ��������� �������������� ��������
if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// ���� ���������� ������������� �������� � ����� ������ ��
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// ���������� ��� ����� ��������� �������� 1, ����� ��������� ���������������, ��� 0, ����� �������� ���������������.
step:='RelativeMouseMode';
MyRDPWin.AdvancedSettings9.RelativeMouseMode:=CheckBoxMouseMode.Checked;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=CheckBoxClipboard.Checked;  /// ���������, ������� �� �������� ��� ��������� ��������������� ������ ������.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=CheckRedirectDevices.Checked;   /// ���������, ������ �� ���������������� ���������� ���� �������� ��� ���������.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=CheckBoxPort.Checked;       /// ��������������� ������
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=CheckBoxConToAdmSrv.Checked; /// ��������� ������ �� ������� ���������� ActiveX �������� ������������ � ������� � ���������������� �����.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=CheckBoxRecAudio.Checked;  ///  ���������, ���������������� �� ���������� ���������� �� ��������� �� ������� � ���������� ������
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=CheckBoxEnSuperPan.Checked;              ///��������� ������������ ����� ������������ �� ���������� �������� ����� � ������������� ������, ����� ������� ���������� �������� ����� ������ �������� �������� ���� �������. ������ ����, ����� ������������ ������ ��������� ��� ��������� �� �������� �����, ������������ ����� ������� �� ������� ����, � ��������� ������� ���� ����� ������������� �������������� � ���� �����������. SuperPan �� ������������ ����� ������ ��������.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=ComboBoxAuthLevel.ItemIndex;           ///������ ������� �������� ����������� ��� �����������.  (0.1.2)0- ��� �������������� �� �������. 1- �������� ����������� ������� ��������� � ������ ������� ����������� ��� ����������� �����������. 2- ������� �������������� �������. ���� �������������� �� �������, ������������ ����� ���������� �������� ���������� ��� ���������� ��� �������������� �� �������.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=CheckBoxCredSsp.Checked;       ///��������� ��� ���������, ������� �� ��� ����� �����������  ��������� ������ ������������ ������� ������ (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=ComboKey.ItemIndex; // 0-���������� ���������� ������ ������ �������� �� ���������� ����������. 1- ���������� ���������� ������ �� ��������� �������. 2- ���������� ���������� ������ � ���������� ������� ������ �����, ����� ������ �������� � ������������� ������. ��� �������� �� ���������.

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
        memo1.Lines.add('������ �������� ���������� ����  :'+E.Message);
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

function ActProc(TipForm,NameForm:string):bool; /// ������� ������ ���������� � ����� �� �������� ���� ���� ��� ��������
var
hwn: THandle;
begin           //'TForm1', 'Form1'
  try
  hwn := FindWindow(pchar(TipForm),nil);
  if hwn<>0 then
  begin
  //OwnerHwn:=GetWindow(hwn,GW_OWNER); /// �������� ��������
  ShowWindow(GetWindow(hwn,GW_OWNER), SW_SHOWNORMAL); //SW_RESTORE   // SW_SHOWMAXIMIZED  //SW_SHOWNORMAL
  SetForegroundWindow(GetWindow(hwn,GW_OWNER));
  end
  else result:=false;
  except
   result:=false;
   end;
end;


function TFormRDP.ConnectToCurentSetting(NamePC:string;IndexTab:integer):boolean;
var     //// ���������� ���� �������� ��� ��������� �� ������������ �����������
i:integer;
begin
for I := 0 to PageControl1.pages[IndexTab].ComponentCount-1 do
begin
 if PageControl1.pages[IndexTab].Components[i] is TMsRdpClient9 then
 with PageControl1.pages[IndexTab].Components[i] as TMsRdpClient9 do
  begin
  ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex); // ���������� ��� �� �������
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
  ConnectingText:='��������� ���������� c '+NamePC+'...';
  DisconnectedText:='������ '+NamePC+' ��������';
  if ComboBoxNetworkConnectionType.ItemIndex=0 then
  begin
  AdvancedSettings9.BandwidthDetection:=true;          /// ���� NetworkConnectionType = 0 ��....���, ����� �� ������������� ���������� ��������� ���������� �����������.
  AdvancedSettings9.PerformanceFlags:=0;
  end
  else
  begin
  AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // ��������� ���� ������������ ���� 1- �����, 2-��������������� �������������� ����� (�� 256 ���� / � �� 2 ���� / �), 3������� (�� 2 ���� / � �� 16 ���� / �, � ������� ���������), 4 -���������������� �������������� ������ (�� 2 ���� / � �� 10 ���� / �), 5- ���������� ���� (WAN) (10 ���� / � ��� ����, � ������� ���������), 6 -��������� ���� (LAN) (10 ���� / � ��� ����)
  AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  ������ ����� �������, ������� ����� ���������� �� ������� ��� ��������� ������������������.
  end;
  AdvancedSettings9.MaxReconnectAttempts:= SpinMaxReconnectAttempts.Value;           // MaxReconnectAttempts
  AdvancedSettings9.AudioRedirectionMode:= ComboAudioRedirectionMode.ItemIndex;      // AudioRedirectionMode
  AdvancedSettings9.AuthenticationLevel:= ComboBoxAuthLevel.ItemIndex;             //  AuthenticationLevel
  SecuredSettings3.KeyboardHookMode := ComboKey.ItemIndex;                        // KeyboardHookMode
  AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked);         //SmartSizing
  AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked;            //GrabFocusOnConnect
  AdvancedSettings9.BandwidthDetection:=true;                                          //BandwidthDetection
  AdvancedSettings9.EnableAutoReconnect:=CheckEnableAutoReconnect.Checked;                  //EnableAutoReconnect
  if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// ���� ���������� ������������� �������� � ����� ������ ��
  AdvancedSettings.DisableRdpdr:=0; //// ���������� ��� ����� ��������� �������� 1, ����� ��������� ���������������, ��� 0, ����� �������� ���������������.
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
begin    //// ���� ��� ��������� ����� �� flgs=0 ������ �� ���������
flgs:=0;

if CheckBox2.checked then  flgs:= flgs + 1;  ///���� ��� �������� ����
if CheckBox3.checked then  flgs:= flgs + 2;  //�������������� � ������ ���� ���������; ��� ����������� ���� ������������ ������ ������ ����.
if CheckBox5.checked then  flgs:= flgs + 4; //�������� ���� ���������.
if CheckBox6.checked then  flgs:= flgs + 8; //���� ���������.
if CheckBox7.checked then  flgs:= flgs + 16; //�������� ���������� ������� ��� ����������� ���� � ���� (+10)
if CheckBox8.checked then  flgs:= flgs + 32; // ��������� ���� ��� �������
if CheckBox9.checked then  flgs:= flgs + 64; // ��������� ������� �������
if CheckBox1.checked then  flgs:= flgs + 128; //�������� ����������� �������
if CheckBox10.checked then  flgs:= flgs + 256; //�������� ���������� �� ������� �����
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
        memo1.Lines.add('������ SmartSizing  :'+E.Message);
      end;
   end;
end;


///////////////////////////////////////////////////////////////////////////////////////
//////////////////������ �������� � �������� ������ ������� �� ������ ��������//////////////////////////////////
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

  //Result.Left :=Result.Right - (ButtonR.Width) - 5; ///������������  ������ �� �������
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

///////////////// ����� �������� � �������� ������ ��� �������� ��������///////////////////////////////////////
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
memo1.Lines.add('�� ���������� ���� ������');
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
end else result:=false; /// ���� ��������� ��� ����� �� ����� �� false
DataMod.FDQueryRead.Close;
Except
  on E: Exception do
  begin
  result:=false;
  memo1.Lines.add('��� ������ �������� ��� '+s+' ��������� ������ '+E.Message);
  end;
end;
end;

function TFormRDP.sizeRegion(i,z:integer):bool; /// �������� �������� ���� ���������
var
res:integer;
begin   /// ���� ������ ���� ��������� ������ �� 15 �������� ������ �� ������� �� ��������� �� ������ ��� ������
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
  memo1.Lines.add('������ �������� ������ ����������� :'+E.Message);
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
  Memo1.Lines.Add(s+': �������� �������� �������� �������');
  end
else
  begin
  result:=true; ///��������
 Memo1.Lines.Add('IP ����� - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  �����='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'��'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;
   except
    begin
    result:=false;
    Memo1.Lines.Add(s+': ���� �� ��������');
    end;
   end;
Myidicmpclient.Free;
end;





procedure TFormRDP.PopupMenu1Popup(Sender: TObject);
begin
if PageControl1.MultiLine then PopupMenu1.Items[7].Caption:='������� � ���� ���'
else PopupMenu1.Items[7].Caption:='������� Multiline'

end;



procedure TFormRDP.Button1Click(Sender: TObject);
var
i,z:integer;
begin
try
//Memo1.Lines.Add('�������� �������� - '+inttostr(PageControl1.ActivePageIndex));
//PageControl1.ActivePage
if PageControl1.PageCount=0 then  exit;

for I := 0 to PageControl1.ActivePage.ComponentCount-1 do //// ������� ���������� � �����
begin
 if PageControl1.ActivePage.Components[i] is TMsRdpClient9 then /// ���� ����� TMsRdpClient9
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
  memo1.Lines.add('������ ��������� ������� ���� :'+E.Message);
  end;
end;
end;



procedure TFormRDP.ComboBoxNetworkConnectionTypeChange(Sender: TObject);
var
i:integer;
begin   ///// ��������� ���� ���� �� ������� ������������ �� RDP
case (sender as Tcombobox).ItemIndex of
0: // ���������� ������������� , ���������� ���� (10 ����/� ��� ���� � ������� ���������),��������� ���� (10 ����/� � ����)
  begin
   FlagsEnabled(false);
  end;
1:  //����� (56 ����/�)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // ������� �����
  checkbox2.Checked:=true;
  checkbox3.Checked:=true;
  checkbox5.Checked:=true;
  checkbox6.Checked:=true;
  checkbox8.Checked:=true;
  checkbox9.Checked:=true;
end;
2: //��������������� �������������� ����������� (256 ����/� - 2 ����/�)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // ������� �����
  checkbox2.Checked:=true;
  checkbox3.Checked:=true;
  checkbox5.Checked:=true;
  checkbox8.Checked:=true;
  checkbox9.Checked:=true;
end;
3,4: // ������� (2-16 ����/� � ������� ���������), ���������������� �������������� ����������� (2-10 ����/�)
begin
  FlagsEnabled(true);
  FlagsChecked(false); // ������� �����
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
  FlagsChecked(false); // ������� �����
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
if not readSettingsInDB(ComboKomp.Text)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
readSettingsInDB('SETDEFAULT');
SpeedButton3.Click; /// ���� ������ enter
end;
end;


procedure TFormRDP.ComboKompSelect(Sender: TObject);
begin /// ��� ������ ���������� ��������� ��������� �� ��
if not readSettingsInDB(ComboKomp.Text)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
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
        on E: Exception do Memo1.Lines.Add('������ ��� LogonUser '+NamePC+' - '+e.Message)
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
 MyRDPWin.ConnectingText:='��������� ���������� � '+NamePC+'...';
 MyRDPWin.DisconnectedText:='������ '+NamePC+' ��������';
 MyRDPWin.Server:=Widestring(NamePC);
 MyRDPWin.Domain:=WideString(Domain);
 if UserName<>'' then MyRDPWin.UserName:=Widestring(UserName);
 if Passwd<>'' then MyRDPWin.AdvancedSettings9.ClearTextPassword:=Widestring(Passwd);

//MyRDPWin.SecuredSettings.StartProgram:='notepad.exe';
step:='';
//MyRDPWin.AdvancedSettings9.RedirectPOSDevices:=true; // ������������� ��� ��������� ������������ ��� ��������������� ���������� Point of Service.
MyRDPWin.ColorDepth:=ColorDepth; //// ������� ����� 32-24-16-8 ��� �� �������
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=BitmapPeristence; // 0 ���������,1 ��������.  ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='CachePersistenceActive';
MyRDPWin.AdvancedSettings9.CachePersistenceActive:=CachePersistenceActive; // 0 - ���������, 1 - ��������. ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='BitmapCacheSize';
MyRDPWin.AdvancedSettings9.BitmapCacheSize:=BitmapCacheSize; // �� 1 �� 32 . ������ ����� ���� ��������� ����������� � ����������, ������������� ��� ��������� ����������� 8 ��� �� �������
step:='VirtualCache16BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=VirtualCache16BppSize;// ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCache32BppSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=VirtualCache32BppSize;// ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCacheSize';
MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=VirtualCacheSize; //������ ������ ����� ����������� ���� ���������� ����������� � ����������, ������� ������������ ��� ����� 8 ��� �� �������. ���������� �������� �������� ����� �������� - �� 1 �� 32 ������������. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��. ��������� �������� �������� �������� BitmapVirtualCache16BppSize � BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=DisableCtrlAltDel; // 0 - ���������, 1 - ��������.  ���������, ������ �� ������������ ��������� ������������� ����� � Winlogon.
step:='SmartSizing';
MyRDPWin.AdvancedSettings9.SmartSizing:= not(SmartSizing); //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=EnableWindowsKey;   //0 - ���������, 1 - ��������. ���������, ����� �� ������������ ���� Windows � ��������� ������.
step:='GrabFocusOnConnect';
MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=GrabFocusOnConnect; //0 - ���������, 1 - ��������. ���������, ������ �� ������� ���������� ������� ����� ����� ��� �����������. ������� ���������� �� ����� �������� ��������� ����� �� ����, ����������� � ������ ��������.
step:='MinutesToIdleTimeout';
MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= MinutesToIdleTimeout; //�� ��������� 0 ������ �.�. �� ����������� �����. ������ ������������ ���������� ������� � �������, � ������� �������� ������ ������ ���������� ������������ ��� ������� ������������. ���� ��������� ����� �������, ������� ���������� �������� ����� IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= OverallConnectionTimeout; //���������� ����� ����������������� ������� � ��������, � ������� �������� ���������� ������� ���������� ������� ���������� ����������. ������������ ���������� �������� ����� �������� - 600, ��� ������������� 10 �������. ���� ��������� ����� �������� �� ���������� ����������, ������� ���������� ����������� � �������� ����� IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
MyRDPWin.AdvancedSettings9.RdpPort:=RdpPort;


if NetworkConnectionType=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=BandwidthDetection;          /// ���� NetworkConnectionType = 0 ��....���, ����� �� ������������� ���������� ��������� ���������� �����������.
MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end
else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= NetworkConnectionType; // ��������� ���� ������������ ���� 1- �����, 2-��������������� �������������� ����� (�� 256 ���� / � �� 2 ���� / �), 3������� (�� 2 ���� / � �� 16 ���� / �, � ������� ���������), 4 -���������������� �������������� ������ (�� 2 ���� / � �� 10 ���� / �), 5- ���������� ���� (WAN) (10 ���� / � ��� ����, � ������� ���������), 6 -��������� ���� (LAN) (10 ���� / � ��� ����)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=PerformanceFlags; //  ������ ����� �������, ������� ����� ���������� �� ������� ��� ��������� ������������������.
//Memo1.Lines.add('PerformanceFlags - '+inttostr(PerformanceFlags));
//Memo1.Lines.add('NetworkConnectionType - '+inttostr(NetworkConnectionType));
end;
//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:=MaxReconnectAttempts; //������ ���������� ������� ��������������� �� ����� ��������������� ���������������. ���������� �������� ����� �������� - �� 0 �� 200 ������������.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:= EnableAutoReconnect;       //���������, ������� �� ��������� ����������� �������� ���������� ������������� ������������ � ������ � ������ ���������� ����.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=AudioRedirectionMode;  // 0-������������� ����� �� �������. 1-��������������� ������ �� ��������� ����������. 2-��������� ��������������� �����; �� �������������� ����� �� �������.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=RedirectDrives;     ///��������� �������������� �����
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=RedirectPrinters;   /// ��������� �������������� ��������

if RedirectClipboard or RedirectPrinters then /// ���� ���������� ������������� �������� � ����� ������ ��
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// ���������� ��� ����� ��������� �������� 1, ����� ��������� ���������������, ��� 0, ����� �������� ���������������.

step:='RelativeMouseMode';
MyRDPWin.AdvancedSettings9.RelativeMouseMode:=RelativeMouseMode;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=RedirectClipboard;  /// ���������, ������� �� �������� ��� ��������� ��������������� ������ ������.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=RedirectDevices;    /// ���������, ������ �� ���������������� ���������� ���� �������� ��� ���������.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=RedirectPorts;      /// ��������������� ������
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=ConnectToAdministerServer; /// ��������� ������ �� ������� ���������� ActiveX �������� ������������ � ������� � ���������������� �����.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=AudioCaptureRedirectionMode;  ///  ���������, ���������������� �� ���������� ���������� �� ��������� �� ������� � ���������� ������
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=EnableSuperPan;              ///��������� ������������ ����� ������������ �� ���������� �������� ����� � ������������� ������, ����� ������� ���������� �������� ����� ������ �������� �������� ���� �������. ������ ����, ����� ������������ ������ ��������� ��� ��������� �� �������� �����, ������������ ����� ������� �� ������� ����, � ��������� ������� ���� ����� ������������� �������������� � ���� �����������. SuperPan �� ������������ ����� ������ ��������.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=AuthenticationLevel;            ///������ ������� �������� ����������� ��� �����������.  (0.1.2)0- ��� �������������� �� �������. 1- �������� ����������� ������� ��������� � ������ ������� ����������� ��� ����������� �����������. 2- ������� �������������� �������. ���� �������������� �� �������, ������������ ����� ���������� �������� ���������� ��� ���������� ��� �������������� �� �������.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=EnableCredSspSupport;      ///��������� ��� ���������, ������� �� ��� ����� �����������  ��������� ������ ������������ ������� ������ (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=KeyboardHookMode; // 0-���������� ���������� ������ ������ �������� �� ���������� ����������. 1- ���������� ���������� ������ �� ��������� �������. 2- ���������� ���������� ������ � ���������� ������� ������ �����, ����� ������ �������� � ������������� ������. ��� �������� �� ���������.

PageControl1.ActivePageIndex:=PageControl1.PageCount-1; //// ��������� ��������� ��������� ��������
//MyRDPWin.StartConnected:=1;   //���������� ��� ����� ��������� �������� TRUE, ���� ������� ���������� ������ ������������ ����� ��� �������, ��� FALSE � ��������� ������.

if AutoConnect then
begin
 ////// ���� ��� ��������� �� ���� ����������� �� �������  //////// � ���� �������� �� ������ ���� �� ��� ������������
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
   Memo1.Lines.Add('������ ��� �������� ����������� � ���������� '+NamePC+' - '+e.Message) ;
   Memo1.Lines.Add('������ �� ����� ��������� - '+step) ;
  end;
end;
end;



procedure TFormRDP.SpeedButton1Click(Sender: TObject); //// ����������� �� RDP
var
i,z:integer;
TabTrue:Boolean;
begin
try
  ////////////////////// ���� ������ ������� � ������ �� ��� ��������
if PageControl1.TabIndex<>-1 then    /// ���� ���� �������� �������
begin
  if PageControl1.ActivePage.ImageIndex<>4 then /// ���� ������ ���������� �� ����� ok
   begin
   for I :=0 to PageControl1.ActivePage.ComponentCount-1 do
     begin
       if PageControl1.ActivePage.Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Connected<>1 then /// ���� ���������� �� �����������
        begin
        ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// ���������� �������� �� �������� �������
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
ShowMessage('�� ������� ��� ����������');
ComboKomp.SetFocus; /// ��������� ����� �� ���������
exit;
end;

if not ping(ComboKomp.Text) then  /// ���� ��������� �� ��������
begin
i:=MessageBox(Self.Handle, PChar('���������� �����������?')
        , PChar('��������� - '+ComboKomp.Text+' �� ��������' ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
end;




///////////////////// ���� ������� � �����������
TabTrue:=false; /// ������� ��� �������� ������ �����������
if PageControl1.TabIndex<>-1 then    /// ���� ���� �������� �������
begin
for I := 0 to PageControl1.PageCount-1 do   /// ������� ��� � �����
Begin
 if PageControl1.Pages[i].Caption=ComboKomp.Text then /// ������� ����, ���� ��� ���������
 begin
  PageControl1.ActivePageIndex:=i; /// ���� �����, ��������� �� �������
  for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ���� ������ ��� �� �������
   begin
     if PageControl1.ActivePage.Components[z] is TMsRdpClient9  then  ///���� �����
     begin
     if (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connected<>1 then /// ���� ���������� �� �����������
      begin
      ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// ���������� �������� �� �������� �������
      (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connect;
      end;
      TabTrue:=true;
      break; // ���� ����� � ���������� �� ������� �� �����
     end;
   end;
 break; /// ���� ����� ������� �� ������� �� �����.
 end;
End;
end;

if not TabTrue then  /// ���� ����������� � ��������� ����
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


if  (PageControl1.PageCount<CountTab) and (not TabTrue) then /// ��� �� ����� ����������� �� ������� �����
 begin
 i:=MessageBox(Self.Handle, PChar('������� ����� �����������?')
 , PChar('����������� ��� '+ComboKomp.Text ) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then exit;
 createnewconnectInTab(ComboKomp.Text);
 end
except
 on E:Exception do
  begin
   Memo1.Lines.Add('������ ��� ����������� � ���������� '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;

end;



procedure TFormRDP.SpeedButton2Click(Sender: TObject); //// ���������� RDp �� �������
var
i:integer;
begin
try
if PageControl1.TabIndex<>-1 then  /// ���� ���� �������
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
   Memo1.Lines.Add('������ ��� ���������� �� ���������� '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;
end;



procedure TFormRDP.SpeedButton3Click(Sender: TObject); /// �������� �������
var
i,z:integer;
TabTrue:Boolean;
begin
if ComboKomp.Text='' then
  begin
  ShowMessage('�� ������� ��� ����������');
  ComboKomp.SetFocus; /// ��������� ����� �� ���������
  exit;
  end;
TabTrue:=false;
for I := 0 to PageControl1.PageCount-1 do /// ������� ���� �� ��� ����������� ������ �� ��������
begin
  if PageControl1.Pages[i].Caption=ComboKomp.Text then  /// ���� ����� 
  begin
  PageControl1.ActivePageIndex:=i;  /// ��������� �� �������
  for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ���� ������ ���
    begin
    if PageControl1.ActivePage.Components[z] is TMsRdpClient9  then  /// ���� �����
      begin
       if (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connected<>1 then /// ���� ���������� �� �����������
       begin
       ConnectToCurentSetting(Pagecontrol1.ActivePage.Caption,Pagecontrol1.TabIndex); /// ���������� �������� �� �������� �������
      (PageControl1.ActivePage.Components[z] as TMsRdpClient9).Connect;
       end;
      TabTrue:=true;
      break; // ���� ����� � ���������� �� ������� �� �����
      end;           
    end;
   break;
   end;
end;

if not TabTrue then  ///������� ���� ����������� � ��������� ����
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


if not ping(ComboKomp.Text) then  /// ���� ��������� �� ��������
begin
i:=MessageBox(Self.Handle, PChar('���������� �����������?')
        , PChar('��������� - '+ComboKomp.Text+' �� ��������' ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
end;
/// ��� ��� ��������� � ������� ������ 1000 �� ������ 1000 ������� �� ���������
if (PageControl1.PageCount<CountTab) and (not TabTrue) then
 begin  //// ������� �����������
 createnewconnectInTab(ComboKomp.Text);
 end;

end;



procedure TFormRDP.SpeedButton4Click(Sender: TObject); /// ������� �������
var
z:integer;
begin
try
if PageControl1.PageCount>0 then
 begin
    for z :=0 to PageControl1.ActivePage.ComponentCount-1 do //// ���� ������ ���
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
        memo1.Lines.add('������ �������� �������:'+E.Message);
      end;
   end;
end;







procedure TFormRDP.TabSheet1Resize(Sender: TObject); /// ��������� ������� ������� ��������� ��� ��������� ������� �������
var
i:integer;
begin
try

 if (PageControl1.PageCount<1)or(PageControl1.TabIndex=-1) then exit; /// ���� ������� ��� �� � ������ ������ �� � ����
 for I := 0 to PageControl1.ActivePage.ComponentCount-1 do //// ������� ���������� � �����
begin
 if PageControl1.ActivePage.Components[i] is TMsRdpClient9 then /// ���� ����� TMsRdpClient9
 begin 
 if(sizeRegion((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopWidth,PageControl1.ActivePage.Width)) or  /// ��������� ������ � ������ ����
 (sizeRegion((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopHeight,PageControl1.ActivePage.Height)) then   /// ���� ���������� ������ reconect
   begin
    (PageControl1.ActivePage.Components[i] as TMsRdpClient9).Reconnect(PageControl1.ActivePage.Width,PageControl1.ActivePage.Height);
   break;
   end;
 {Memo1.Lines.add('������ TMsRdpClient9 - '+inttostr((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopWidth));  
 Memo1.Lines.add('������ ActivePage - '+inttostr(PageControl1.ActivePage.Width)); 
 Memo1.Lines.add('������ TMsRdpClient9 - '+inttostr((PageControl1.ActivePage.Components[i] as TMsRdpClient9).DesktopHeight)); 
 Memo1.Lines.add('������ ActivePage - '+inttostr(PageControl1.ActivePage.Height)); 
 Memo1.Lines.Add('��������� - '+(PageControl1.ActivePage.Components[i] as TMsRdpClient9).Server)};
 end;

 
end;
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ ��� ��������� ������� �������� :'+E.Message);
  end;
end;
end;

procedure TFormRDP.WMCopyData(var MessageData: TWMCopyData); /// ��������� ���������
var
ListStr:string;
ListPC:TstringList;
i:integer;

const
 SEND_SETLISTTEXT = 1;  /// Receiv ���������
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv ������ �� ��������� ������
 RECEIVE_REQ_USPASS=3;    ///Receiv ������ ������ � ������
begin
try
  //���� ������ ������ �������� ������� ���������
  if MessageData.CopyDataStruct.dwData = SEND_SETLISTTEXT then
  Begin
    //������������� ����� �� ���������� ������
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
    //������������� ����� �� ���������� ������
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
        memo1.Lines.add('������ ��������� ������ �� ������� ��������� :'+E.Message);
      end;
   end;
//Memo1.Lines.Add('������ - '+ListStr);
end;

Function TFormRDP.RequestStr(nameStr:string;Receiv:integer):bool; /// ������� ������� ������ ������ ��� ������ � ������ ��������� �� �������� ����������
var
  CDS: TCopyDataStruct;
  AllStr:TstringList;
  i:integer;
const
 SEND_SETLISTTEXT = 1;  /// Receiv ���������
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv ������ �� ��������� ������
 RECEIVE_REQ_USPASS=3;    ///Receiv ������ ������ � ������
 begin
 try
//������������� ��� �������
  CDS.dwData := Receiv;
  //������������� ����� ������������ ������
  CDS.cbData := Length('') + 1;
  //�������� ������ ������ ��� �������� ������
  GetMem(CDS.lpData, CDS.cbData);
  try
    //�������� ������ � �����
    StrPCopy(CDS.lpData, AnsiString(''));
    //�������� ��������� � ���� � ���������� StringReceiver
    SendMessage(FindWindow(Pchar(nameStr), nil),
                  WM_COPYDATA, Handle, Integer(@CDS));
  finally
    //������������ �����
    FreeMem(CDS.lpData, CDS.cbData);
  end;
 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ ������� ������ �� ������� ��������� :'+E.Message);
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

 TabShetCount:=0; //// ���������� ��������� ���������. ��� ���������� ������� ������������ ����������
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
FormRDP.caption:=FormRDP.caption+' ��������!!! �� ������������������ ������, �������� 3 �����������.';
end;}

end;

procedure TFormRDP.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (ssCtrl in Shift) AND (key=38) then    //38- ������� ����� 40- ������� ����
  FormRDP.Align:=alClient;
 if (ssCtrl in Shift) AND (key=40) then    //38- ������� ����� 40- ������� ����
  FormRDP.Align:=alNone;
end;



procedure TFormRDP.FormShow(Sender: TObject);
var
listpc:TstringList;
i:integer;
const
 SEND_SETLISTTEXT = 1;  /// Receiv ���������
 RECEIVE_REQ_LISTTEXT=2;  ///Receiv ������ �� ��������� ������
 RECEIVE_REQ_USPASS=3;    ///Receiv ������ ������ � ������
begin
try
RequestStr('TfrmDomainInfo',2); /// ������ ������ ������ �� �������� ����������
RequestStr('TfrmDomainInfo',3); /// ������ ������, ������ � ������ �� �������� ����������
//////////////////////////////////////////////////////////////////////////////
readSettingsInDB('SETDEFAULT'); /// ������ ��� ���������� �������� ����������� �� ���������
//////////////////////////////////////////////////////////////////////////

 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ �������� ���������� : '+E.Message);
      end;
   end;
end;



procedure TFormRDP.MsRdpClient91AuthenticationWarningDismissed(Sender: TObject);
begin
   ///���������� ����� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).
end;

procedure TFormRDP.PopupOsherWinPanelPopup(Sender: TObject);  ////��� ����������� ����� ���� �� ������ � ��������� ����
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

procedure TFormRDP.N12Click(Sender: TObject); //// ����� ���� �� ������ � ��������� ����, ������ ���������
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

procedure TFormRDP.N1Click(Sender: TObject); /// ����� ����, ���������� �� �������
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;

//if PageControl1.TabIndex<>-1 then    /// ���� ���� �������� �������
begin
  if PageControl1.Pages[tabInd].ImageIndex<>4 then /// ���� ������ ���������� �� ����� ok
   begin
   for I :=0 to PageControl1.Pages[tabInd].ComponentCount-1 do
     begin
       if PageControl1.Pages[tabInd].Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Connected<>1 then /// ���� ���������� �� �����������
        begin
         //PageControl1.TabIndex:=tabInd;
         ConnectToCurentSetting(PageControl1.Pages[tabInd].caption,tabInd); /// ���������� �������� �� �������� �������
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
        memo1.Lines.add('������ ����������� :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N2Click(Sender: TObject); //// ����� ���� ���������� �������
var
tabInd,i:integer;
tabCapt:string;
begin
try
tabInd:=PageControl1.IndexOfTabAt(PMouse.x,PMouse.y);
tabCapt:=PageControl1.Pages[tabInd].Caption;
//if PageControl1.TabIndex<>-1 then    /// ���� ���� �������� �������
begin
  if PageControl1.Pages[tabInd].ImageIndex<>5 then /// ���� ������ ���������� �� ����� ���������
   begin
   for I :=0 to PageControl1.Pages[tabInd].ComponentCount-1 do
     begin
       if PageControl1.Pages[tabInd].Components[i] is TMsRdpClient9  then
       begin
       if (PageControl1.Pages[tabInd].Components[i] as TMsRdpClient9).Connected<>0 then /// ���� ����������  ����������� (1) ��� ��� ������������ (2)
        //if PageControl1.Pages[tabInd].ImageIndex<>5 then /// ���� ���������� �� �����������, ��� ��������� � ������ ���������
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
        memo1.Lines.add('������ ���������� :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N3Click(Sender: TObject); /// ����� ����, ������� ������ �������
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
        memo1.Lines.add('������ �������� ����������� :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N4Click(Sender: TObject);  ///����� ����, ������� ������� �����������
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
        memo1.Lines.add('������ �������� ����������� :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.N5Click(Sender: TObject);  /// �����  � ����� ����
var
tabInd,i:integer;
tabCapt:string;
begin
if (CountTab=3) and  (GroupBox8.ControlCount>=1) then
begin
ShowMessage('������ ����������� � ������������������ ������ ���������.');
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
    showmessage('������������� ������� ��� ���������� ��������');
    end;
end;

procedure TFormRDP.N7Click(Sender: TObject); // ������ /���������� ���
var
SetInI:Tinifile;
begin
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    if SetInI.readBool('MRRDP','ViewLog',true) then // ���� ��� �������
    begin
    SetInI.WriteBool('MRRDP','ViewLog',false);
    SetInI.Free;
    exit;
    end;
    if SetInI.readBool('MRRDP','ViewLog',true)=false then /// ���� ��� ������������
    begin
    SetInI.WriteBool('MRRDP','ViewLog',true);
    JvNetscapeSplitter2.Maximized:=true; // ������ ���
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
    PopupLog.Items[0].Caption:='���������� ���';

    if SetInI.readBool('MRRDP','ViewLog',true)=false then
    PopupLog.Items[0].Caption:='�������� ���';

    SetInI.Free;
    end;
end;

procedure TFormRDP.NewAuthenticationWarningDisplayed(Sender: TObject);
begin
 //���������� �� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).

end;

procedure TFormRDP.MsRdpClient91AutoReconnected(Sender: TObject);
begin
 //����������, ����� ���������� ������� ���������� ������������� �������� ������������ � ���������� ������.
end;

procedure TFormRDP.MsRdpClient91AutoReconnecting(ASender: TObject;
  disconnectReason, attemptCount: Integer);  //����������, ����� ������ ��������� � �������� ��������������� ���������� ����������� ������ � ������� ���� ������� ��������� ������� ������ (RD Session Host).
begin
//Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������������, ������� - '+inttostr(disconnectReason)+'/ ������� - '+ SysErrorMessage(attemptCount));
end;

procedure TFormRDP.MsRdpClient91AutoReconnecting2(ASender: TObject;   //����������, ����� ������ ��������� � �������� ��������������� ���������� ����������� ������ � ������� ���� ������� ��������� ������� ������ (RD Session Host).
  disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
  maxAttemptCount: Integer);
var
errorstr:string;
id2:integer;
 function NetW(s:bool):string;
 begin
   if s then result:='����������'
   else result:='���������';
 end;
begin
errorstr:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������������:  '+errorstr+
'/ ��������� ���� :'+netW(networkAvailable)+' / ������� - '+inttostr(attemptCount)+' �� '+inttostr(maxAttemptCount));
end;


procedure TFormRDP.MsRdpClient91DevicesButtonPressed(Sender: TObject);
begin
////����� ��� ������ ������ �� �����
end;



procedure TFormRDP.MsRdpClient91ServiceMessageReceived(ASender: TObject;
  const serviceMessage: WideString);
begin
Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP : '+serviceMessage);
end;


procedure TFormRDP.NewRdpAuthenticationWarningDismissed(Sender: TObject);
begin  //���������� ����� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).
try
 if not (sender is TMsRdpClient9) then exit;

{if  ((sender as TMsRdpClient9).Connected=0)and   /// tckb ���������� �� �����������
 ((sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport=false) then  // � ���������  EnableCredSspSupport
 begin
   (sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport:=true;
   (sender as TMsRdpClient9).Connect;
 end;}


Memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������������� ��������������');
 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - �������������� ��������������." :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpAuthenticationWarningDisplayed(Sender: TObject);
 const  //���������� �� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).
  AutType: array [0..4] of string =('No authentication is used.',
  '������������ �������� ����������� �����������','������������ �������� ����������� Kerberos',
  '������������ ��� ���� �������� �����������, ����������� � �� ���������','����������� ���������');
begin
try
  if not (sender is TMsRdpClient9) then  exit;
 if ((sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType<5) then
 memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������� �����������. - '
 + AutType[(sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType]);

 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - �������� �����������." :'+E.Message);
      end;
   end;
end;





procedure TFormRDP.NewRdpAutoReconnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
Memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ��������������� � ���������� ����������.- ');
 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - ��������������� � ���������� ����������" :'+E.Message);
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
+' : RDP - �������������� ��������������� ������ ��������� ������� ������: '+strerror);
 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - �������������� ��������������� ������ ��������� ������� ������" :'+E.Message);
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
+' : RDP - ���������� ������� ���������� ���������������� � ���������� ������ :'+strerror);
 Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - ���������� ������� ���������� ���������������� � ���������� ������" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpChannelReceivedData(ASender: TObject; const chanName,
  data: WideString);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ������ �������� ������ �� ����������� ������ � ������������ ��������');
Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - ������ �������� ������ �� ����������� ������ � ������������ ��������" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConfirmClose(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ���������� ������ ����� ���������� �������� �����');
Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - ���������� ������ ����� ���������� �������� �����" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// ����� ����� ������
memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ��������� ����������...');
Except
   on E: Exception do
      begin
        memo1.Lines.add('������ "RDP - ��������� ����������" :'+E.Message);
      end;
   end;
end;

procedure TFormRDP.NewRdpConnecting(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// ���������� ������ ������
memo1.Lines.Add((sender as TMsRdpClient9).Server+
' : RDP - ������������� ����������...');
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
        memo1.Lines.add('������ "RDP - ������������� ����������" :'+E.Message);
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
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������� ������ - '+errorstr);
if TComponent(Asender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
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
  memo1.Lines.add('������ ���������� ������� :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpExit(Sender: TObject);
begin
try
{if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - ����� ���������'); }

Except
  on E: Exception do
  begin
  memo1.Lines.add('������ "RDP - ����� ���������" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpFatalError(ASender: TObject; errorCode: Integer);
const
StrError : array [0..7] of string=('��������� ����������� ������.','��� ���������� ������ 1.'
,'��������� ������ �������� ������.','��������� ������ �������� ����.','��� ���������� ������ 2.'
,'���������� ��� ������ 3','��� ���������� ������ 4.','�� ����� ����������� � ������� ��������� ������������ ������.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if errorCode=100 then memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ���������� ������������ ������ - ������ ������������� Winsock.')
else
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP - ���������� ������������ ������ - '+StrError[errorCode]);
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ "RDP - ���������� ������������ ������" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpLoginComplete(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������� ���� � �������');

if TComponent(sender).Owner is TTabSheet then   /// ���� �������� ��� ������� TabSheet � �� ��������� ����
begin
TTabSheet(TComponent(sender).Owner).ImageIndex:=4;
DataMod.WriteItemDB((sender as TMsRdpClient9).Server); /// ���� ������ ������������ ���������� ��������� ����������� � ����
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
  memo1.Lines.add('������ ���������� ������� :'+E.Message);
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
' : RDP - ������ ����� � ������� - '+ SysErrorMessage(lError));
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ "������ ����� � �������" :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpNetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
const
qLevel: array [0..5] of string =('�������� �� ��������...','�������� ����� 512 ����/�',
'�������� �� 512 ����/� �� 2 ����/�.','�������� �� 2 �� 10 ����/�.'
,'�������� ������ 10 ����/�.','�������� ������ 100 ����/�.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ��������� ��������� ��������� ����');
memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : '+qLevel[qualityLevel]
+' / ���������� ����������� (Bandwidth) - '+inttostr(bandwidth)+' ����/� / �������� ���������� (RTT) - '+inttostr(rtt)+' ��.')
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ ��� ������ ������������� ��������� ��������� ���� :'+E.Message);
  end;
end;
end;

procedure TFormRDP.NewRdpWarning(ASender: TObject; warningCode: Integer);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if warningCode=1 then
memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP -  ��� ���������� ����������� ���������.');
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ ��� ���������� OnWarning :'+E.Message);
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

if PageControl1.ActivePage.ImageIndex=5 then  /// ���� ������ ��������, �� ������ ��������� �� ��
begin
if not readSettingsInDB(PageControl1.ActivePage.Caption)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
readSettingsInDB('SETDEFAULT');
end;
Except
  on E: Exception do
  begin
  memo1.Lines.add('������ ��� ������������ :'+E.Message);
  end;
end;
end;





procedure TFormRDP.PageControl1DragDrop(Sender, Source: TObject; X, Y: Integer); /// ��� ����������� �������
const
  TCM_GETITEMRECT = $130A;
var
  i: Integer;
  r: TRect;
begin
try
if PageControl1.PageCount=0 then exit;
/// ����� ���
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
  // ��� ���


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
  memo1.Lines.add('������ PageControlDragDrop :'+E.Message);
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
  memo1.Lines.add('������ PageControlMouseDown :'+E.Message);
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
  memo1.Lines.add('������ PageControlMouseUp :'+E.Message);
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
  memo1.Lines.add('������ PageControl1DragOver :'+E.Message);
  end;
end;
end;



procedure TFormRDP.Button2Click(Sender: TObject); //// ���������� ������ �����������
var
i:integer;
ListIni:Tinifile;
NewDescription:string;
begin
try
if PageControl1.PageCount<=0 then
begin  
  showmessage('��� �����������');
  exit;
end;

if not InputQuery('�������� ������ �����������', '��������', NewDescription)
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
  memo1.Lines.add('������ ���������� ������ ����������� :'+E.Message);
  end;
end;
end;


procedure TFormRDP.Button3Click(Sender: TObject);
begin
FormLoadList.Show;
end;

procedure TFormRDP.Button4Click(Sender: TObject);
var   //// ���������� ������� ���� ����������� ��� ��������� ������
  CDS: TCopyDataStruct;
  AllStr:TstringList;
  i:integer;
  const
 SEND_SETLISTTEXT = 1;  /// ���������
  RECEIVE_REQ_LISTTEXT=2;  /// ������ �� ��������� ������
begin
try
//������������� ��� �������
  CDS.dwData := RECEIVE_REQ_LISTTEXT;
  //������������� ����� ������������ ������
  CDS.cbData := Length('') + 1;
  //�������� ������ ������ ��� �������� ������
  GetMem(CDS.lpData, CDS.cbData);
  try
    //�������� ������ � �����
    StrPCopy(CDS.lpData, AnsiString(''));
    //�������� ��������� � ���� � ���������� StringReceiver
   // SendMessage(GetWindow(FindWindow('TfrmDomainInfo', nil),GW_OWNER),
   // WM_COPYDATA, Handle, Integer(@CDS));
   SendMessage(FindWindow('TfrmDomainInfo', nil),
    WM_COPYDATA, Handle, Integer(@CDS));
  finally
    //������������ �����
    FreeMem(CDS.lpData, CDS.cbData);
  end;

Except
   on E: Exception do
      begin
        memo1.Lines.add('������ ���������� ������� �����������:'+E.Message);
      end;
   end;
end;


procedure TFormRDP.Button5Click(Sender: TObject);
begin
Memo1.Lines.Add('Groupbox8 - '+inttostr(Groupbox8.controlcount));
Memo1.Lines.Add('CountTab - '+inttostr(CountTab));
//if not readSettingsInDB(ComboKomp.Text)then /// ���� ��� �������� ��� ����� �����, �� ������ ��������� �� ���������
//readSettingsInDB('SETDEFAULT');
end;


procedure TFormRDP.Button6Click(Sender: TObject);
var
i:integer;
begin
i:=MessageBox(Self.Handle, PChar('���������� ������� ��������� �� ��������� '+#10#13+' ��� ���� ����� �����������?')
        , PChar('��������� ��� RDP �����������' ) ,MB_YESNO+MB_ICONQUESTION);
if i=IDYES then
 begin
 if not DataMod.writesetDB('SETDEFAULT', //- ��� �����
 bitcolorDepth(ComboBoxColorDepth.ItemIndex), // ���������� ��� �� �������
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
 ShowMessage('������ ������ �������� �� ���������');
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
    (PageControl1.ActivePage.Width-50,PageControl1.ActivePage.Height-50, //  ������ �������� �����. � ������ �������� �����.
    PageControl1.ActivePage.Width-10,PageControl1.ActivePage.Height-10,  // ���������� ������. � ���������� ������.
    2, // ���������� �������� �����
    100,  //����������� ��������������� �������� �����.
    100); //����������� ��������������� ����������
  end;

end;


end;


procedure TFormRDP.SpeedButton4MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
SpeedButton4.Hint:='������� ����������� '+ PageControl1.ActivePage.Caption;
end;


procedure TFormRDP.SpeedButton3MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
SpeedButton3.Hint:='�������� ����� ����������� � '+ ComboKomp.Text;
end;

procedure TFormRDP.Button1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
Button1.Hint:='������� ���� �� ������� ����, ��� '+ PageControl1.ActivePage.Caption;
end;



procedure TFormRDP.SpeedButton1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
SpeedButton1.Hint:='����������� � '+Combokomp.Text;
end;

procedure TFormRDP.SpeedButton2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if PageControl1.PageCount>0 then
SpeedButton2.Hint:='������� ���������� '+ PageControl1.ActivePage.Caption;
end;



initialization  /// ������������� ���������� � ������, � ������ ������ ������� �� �������
TStyleManager.Engine.RegisterStyleHook(TCustomTabControl, TTabControlStyleHookBtnClose);
TStyleManager.Engine.RegisterStyleHook(TTabControl, TTabControlStyleHookBtnClose);

    
end.
