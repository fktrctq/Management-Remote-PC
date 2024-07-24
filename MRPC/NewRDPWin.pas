unit NewRDPWin;

interface

uses
  Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
   MSTSCLib_TLB, Vcl.ComCtrls,StrUtils,Winapi.Windows, ActiveX, Vcl.ImgList,
  Vcl.Menus;

type
  TRDPWin = class(TForm)
    BigImage: TImageList;
    ImageButton: TImageList;
    SmallImage: TImageList;
    PopupOtherWin: TPopupMenu;
    N8: TMenuItem;
    Button1: TButton;
    FlowPanel1: TFlowPanel;
    procedure N8Click(Sender: TObject);
  private
    { Private declarations }
  public
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
    procedure NewAuthenticationWarningDisplayed(Sender: TObject);
    procedure MsRdpClient91AutoReconnected(Sender: TObject);
    procedure MsRdpClient91AutoReconnecting(ASender: TObject; disconnectReason,
      attemptCount: Integer);
    procedure MsRdpClient91AutoReconnecting2(ASender: TObject;
      disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
      maxAttemptCount: Integer);
   procedure MsRdpClient91ServiceMessageReceived(ASender: TObject;
      const serviceMessage: WideString);
    procedure MsRdpClient91DevicesButtonPressed(Sender: TObject);
    procedure NewRdpWarning(ASender: TObject; warningCode: Integer);

    //MyOnClickFunc:procedure (Sender: TObject) as object;

   procedure OtherWinForRDPClientClose (Sender: TObject; var Action: TCloseAction);
   procedure PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
   procedure PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// ����������� ������ ��  �����
  Y: Integer);
   procedure PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
   procedure ButtonOnForForm(Sender: TObject);
   procedure ButtonOffForForm(Sender: TObject);
   procedure ButtonMinimizeForForm(Sender: TObject); //// �������� ��������� �����
   procedure ButtonResizeForForm(Sender: TObject); //// �� ���� �����  RDP �� ��������� �����
   procedure ButtonCloseForForm(Sender: TObject); //// ������� ��������� �����
   procedure ShowButtonOnClick(Sender: TObject); /// ������ �� ������ GroupBox8 ��� �������������� � ������������ ����
   procedure ButtonMaxNormalForForm(Sender: TObject); //// ���������� ��� �������� ��������� �����
   procedure ButonMoveWinMouseDown(Sender: TObject;  /// ����������� ����� �� ������
                           Button: TMouseButton;
                            Shift: TShiftState;
                             X, Y: Integer);
   procedure RDPFormShow(Sender: TObject);  // ������������ ���� ���
   function OtherWinForRDPClient(indTab:integer;namePC,Domain,UserName,Passwd:string;OwnerForm:Tform):bool;
  end;

var
  RDPWin: TRDPWin;

implementation
uses umain;
var
movingPan:boolean;
y0,x0:integer;
{$R *.dfm}

procedure TRDPWin.OtherWinForRDPClientClose(Sender: TObject; var Action: TCloseAction); /// ����������� ����� ��� ��������
begin
if Sender is Tform then
(Sender as TForm).Release;
end;

procedure TRDPWin.PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
begin
     movingPan:=true;
     x0:=x;
     y0:=y;
end;

procedure TRDPWin.PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// ����������� ������ ��  �����
  Y: Integer);
begin
     if movingPan then
   begin
     (sender as tpanel).Left:=(sender as tpanel).Left+x-x0;
     (sender as tpanel).Top:=(sender as tpanel).Top+y-y0;
   end;
end;

procedure TRDPWin.PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // ����������� ������ ��  �����
  Shift: TShiftState; X, Y: Integer);
begin
  movingPan := false;
end;

procedure TRDPWin.ButtonOnForForm(Sender: TObject); //// ����������� �� RDP �� ��������� �����
var
i:integer;
namePC:string;
LevelAut:integer;
begin
try
LevelAut:=2; /// �������������
if not( TComponent(sender).Owner is TPanel) then exit;
for I := 0 to ( TComponent(sender).Owner as TPanel).ComponentCount-1 do
if  ( TComponent(sender).Owner as TPanel).Components[i] is Tcombobox then
 LevelAut:=( ( TComponent(sender).Owner as TPanel).Components[i] as Tcombobox).ItemIndex;

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
   frmDomainInfo.memo1.Lines.Add('������ ��� ����������� � ���������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonOffForForm(Sender: TObject); //// ���������� �� RDP �� ��������� �����
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
   frmDomainInfo.Memo1.Lines.Add('������ ��� ���������� �� ���������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonResizeForForm(Sender: TObject); //// �� ���� �����  RDP �� ��������� �����
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
   frmDomainInfo.Memo1.Lines.Add('������ ��� ��������������� '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonCloseForForm(Sender: TObject); //// ������� ��������� �����
var
i:integer;
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
  for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do // ���� rdp
    if ((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] is TMsRdpClient9  then  // ���� �����
     begin
     if (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Connected<>0 then // ���� ���������� �����������
     begin
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Disconnect;         // �����������
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Free;
      break;
      end;                                                                                                   // ����� �� �����
     end;
 ((TComponent(sender).Owner as TPanel).Owner as Tform).Release;   // ���������� �����
end;
except
 on E:Exception do
  begin
   frmDomainInfo.Memo1.Lines.Add('������ ��� �������� �����������  - '+e.Message)
  end;
end;
end;


procedure TRDPWin.ShowButtonOnClick(Sender: TObject); /// ������ �� ������ panel7 ��� �������������� � ������������ ����
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
   frmDomainInfo.Memo1.Lines.Add('������ ��� ������������/�������������� ���� - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonMinimizeForForm(Sender: TObject); //// �������� ��������� �����
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
   frmDomainInfo.Memo1.Lines.Add('������ ��� ������������ ���� - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonMaxNormalForForm(Sender: TObject); //// ���������� ��� �������� ��������� �����
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
   frmDomainInfo.Memo1.Lines.Add('������ ��� ��������� ������� ���� - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButonMoveWinMouseDown(Sender: TObject;  /// ����������� ����� �� ������
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

procedure TRDPWin.RDPFormShow(Sender: TObject);
begin
 // ShowWindow(Application.Handle, SW_HIDE);
end;

function TRDPWin.OtherWinForRDPClient(indTab:integer;namePC,Domain,UserName,Passwd:string;OwnerForm:Tform):bool;
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
GroupConnect:TgroupBox;
step:string;
//MyRDPWinNoScr:IMsRdpClientNonScriptable5;

begin
try

CoInitialize(nil);
NameCapt:=namePC;
FormForRDP:=Tform.Create(OwnerForm);
FormForRDP.Name:='FormForRDP'+inttostr(indTab);
FormForRDP.Position:=poMainFormCenter;
FormForRDP.OnClose:=OtherWinForRDPClientClose;
FormForRDP.Width:=Screen.Width;
FormForRDP.Height:=Screen.Height;
//FormForRDP.WindowState:=wsMaximized;
FormForRDP.Tag:= indTab;
FormForRDP.FormStyle:=fsStayOnTop;//fsMDIChild;//fsNormal;//fsMDIForm;//fsStayOnTop;
FormForRDP.Caption:=NameCapt;
FormForRDP.BorderStyle:=bsnone;
FormForRDP.PopupMode:=pmAuto;
//FormForRDP.OnShow:=RDPFormShow;
//FormForRDP.OnResize:=NewRDPFormResize;


PanelForBut:=Tpanel.Create(FormForRDP);
PanelForBut.Parent:=FormForRDP;
PanelForBut.Name:='PanelForBut'+ inttostr(indTab);
PanelForBut.Caption:='';//NameCapt;
//PanelForBut.Alignment:=taRightJustify;
PanelForBut.ShowHint:=true;
PanelForBut.Hint:='����������� � '+NameCapt+'. ������� ��� �����������';
PanelForBut.Top:=0;
PanelForBut.Left:=0;
PanelForBut.Width:=380;
//PanelForBut.Align:=alTop;
PanelForBut.Height:=30;
PanelForBut.OnMouseDown:=PanelForButMouseDown;
PanelForBut.OnMouseMove:=PanelForButMouseMove;
PanelForBut.OnMouseUp:=PanelForButMouseUp;

ButFP:=TButton.Create(PanelForBut);
ButFP.Parent:=PanelForBut;
ButFP.Name:='ButFPOn'+inttostr(indTab);
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
ButFP.Name:='ButFPOFF'+inttostr(indTab);
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
ButFP.Name:='ButFPResize'+inttostr(indTab);
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
AutLevel.Name:='Level'+inttostr(indTab);
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
ButFP.Name:='MoveWin'+inttostr(indTab);
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
ButFP.Name:='ButFPCollapse'+inttostr(indTab);
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
ButFP.Name:='ButFPCollapseWin'+inttostr(indTab);
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
ButFP.Name:='ButFPClose'+inttostr(indTab);
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

//{if GroupBox8.Caption='' then GroupBox8.Caption:='����������� � ��������� ����';

ButFP:=TButton.Create(FormForRDP);
ButFP.Parent:=frmDomainInfo.GroupRDPWin;
ButFP.Name:='ShowB'+inttostr(indTab);
ButFP.Caption:=namePC;
ButFP.Hint:='����������� � '+NameCapt;
ButFP.ShowHint:=true;
ButFP.Images:=SmallImage;
ButFP.ImageIndex:=0;
ButFP.Tag:=indTab;
//ButFP.Width:=frmDomainInfo.GroupRDPWin.Width;
ButFP.Height:=20;
ButFP.Align:=alBottom;
ButFP.OnClick:=ShowButtonOnClick;

//ButFP.OnMouseMove:=ShowBMouseMove;
//ButFP.OnMouseLeave:=ShowBMouseLeave;
ButFP.PopupMenu:=PopupOtherWin;

MyRDPWin:=TMsRdpClient9.Create(FormForRDP);
MyRDPWin.Parent:= FormForRDP;
MyRDPWin.Name:='NewRDPWin'+inttostr(indTab);
MyRDPWin.SendToBack;/// ��� ���� ����� ���� ����� ������
MyRDPWin.Name:='MyRDPWin'+inttostr(indTab);
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
MyRDPWin.Server:=Widestring(NameCapt);
if Domain<>'' then MyRDPWin.Domain:=WideString(Domain);
if UserName<>'' then MyRDPWin.UserName:=WideString(UserName) ;
MyRDPWin.ConnectingText:='��������� ���������� c ' +namePC+'...';
MyRDPWin.DisconnectedText:='������ '+namePC+' ��������';

step:='';
//MyRDPWin.ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex);
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=1;//booltoint(CheckBoxBitmapPeristence.Checked); // 0 ���������,1 ��������.  ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='CachePersistenceActive';
//MyRDPWin.AdvancedSettings9.CachePersistenceActive:=booltoint(CheckboxCachePersistenceActive.Checked); // 0 - ���������, 1 - ��������. ���������, ������� �� ������������ ���������� ��������� �����������. ���������� ����������� ����� �������� ������������������, �� ������� ��������������� ��������� ������������.
step:='BitmapCacheSize';
//MyRDPWin.AdvancedSettings9.BitmapCacheSize:=SpinBitmapCacheSize.Value;; // �� 1 �� 32 . ������ ����� ���� ��������� ����������� � ����������, ������������� ��� ��������� ����������� 8 ��� �� �������
step:='VirtualCache16BppSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=SpinEditCache16BppSize.Value; // ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCache32BppSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=SpinEditCache32BppSize.Value;// ����� ������ ����. ���������� �������� - �� 1 �� 32 ������������, � �������� �� ��������� - 20. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��.
step:='VirtualCacheSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=SpinBitmapVirtualCacheSize.Value;  //������ ������ ����� ����������� ���� ���������� ����������� � ����������, ������� ������������ ��� ����� 8 ��� �� �������. ���������� �������� �������� ����� �������� - �� 1 �� 32 ������������. �������� ��������, ��� ������������ ������ ���� ������ ������������ ���� ���������� 128 ��. ��������� �������� �������� �������� BitmapVirtualCache16BppSize � BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
//MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=booltoint(CheckBoxDisableCtrlAltDel.Checked); // 0 - ���������, 1 - ��������.  ���������, ������ �� ������������ ��������� ������������� ����� � Winlogon.
step:='SmartSizing';
//MyRDPWin.AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked); //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=1;//booltoint(CheckBoxEnableWindowsKey.Checked);   //0 - ���������, 1 - ��������. ���������, ����� �� ������������ ���� Windows � ��������� ������.
step:='GrabFocusOnConnect';
//MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked; //0 - ���������, 1 - ��������. ���������, ������ �� ������� ���������� ������� ����� ����� ��� �����������. ������� ���������� �� ����� �������� ��������� ����� �� ����, ����������� � ������ ��������.
step:='MinutesToIdleTimeout';
//MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= SpinMinutesToIdleTimeout.Value;  //�� ��������� 0 ������ �.�. �� ����������� �����. ������ ������������ ���������� ������� � �������, � ������� �������� ������ ������ ���������� ������������ ��� ������� ������������. ���� ��������� ����� �������, ������� ���������� �������� ����� IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
//MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= SpinOverallConnectionTimeout.Value; //���������� ����� ����������������� ������� � ��������, � ������� �������� ���������� ������� ���������� ������� ���������� ����������. ������������ ���������� �������� ����� �������� - 600, ��� ������������� 10 �������. ���� ��������� ����� �������� �� ���������� ����������, ������� ���������� ����������� � �������� ����� IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
//MyRDPWin.AdvancedSettings9.RdpPort:=strtoint(LabeledEditRdpPort.Text);


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

//  if ComboBoxNetworkConnectionType.ItemIndex=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=true;          /// ���� NetworkConnectionType = 0 ��....���, ����� �� ������������� ���������� ��������� ���������� �����������.
//MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end ;
{else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // ��������� ���� ������������ ���� 1- �����, 2-��������������� �������������� ����� (�� 256 ���� / � �� 2 ���� / �), 3������� (�� 2 ���� / � �� 16 ���� / �, � ������� ���������), 4 -���������������� �������������� ������ (�� 2 ���� / � �� 10 ���� / �), 5- ���������� ���� (WAN) (10 ���� / � ��� ����, � ������� ���������), 6 -��������� ���� (LAN) (10 ���� / � ��� ����)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  ������ ����� �������, ������� ����� ���������� �� ������� ��� ��������� ������������������.
end; }

//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //���������, ������� �� �������������� �������, ����� �� �������������� ���������� ������� �������� ����������
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:= 20;//SpinMaxReconnectAttempts.Value;   //������ ���������� ������� ��������������� �� ����� ��������������� ���������������. ���������� �������� ����� �������� - �� 0 �� 200 ������������.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:=true;// CheckEnableAutoReconnect.Checked;       //���������, ������� �� ��������� ����������� �������� ���������� ������������� ������������ � ������ � ������ ���������� ����.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=0;//ComboAudioRedirectionMode.ItemIndex;  // 0-������������� ����� �� �������. 1-��������������� ������ �� ��������� ����������. 2-��������� ��������������� �����; �� �������������� ����� �� �������.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=true;//CheckBoxDisk.Checked;     ///��������� �������������� �����
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=true;//CheckBoxPrint.Checked;    /// ��������� �������������� ��������
//if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// ���� ���������� ������������� �������� � ����� ������ ��
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// ���������� ��� ����� ��������� �������� 1, ����� ��������� ���������������, ��� 0, ����� �������� ���������������.
step:='RelativeMouseMode';
//MyRDPWin.AdvancedSettings9.RelativeMouseMode:=CheckBoxMouseMode.Checked;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=true;//CheckBoxClipboard.Checked;  /// ���������, ������� �� �������� ��� ��������� ��������������� ������ ������.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=true;//CheckRedirectDevices.Checked;   /// ���������, ������ �� ���������������� ���������� ���� �������� ��� ���������.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=true;//CheckBoxPort.Checked;       /// ��������������� ������
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=true;//CheckBoxConToAdmSrv.Checked; /// ��������� ������ �� ������� ���������� ActiveX �������� ������������ � ������� � ���������������� �����.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=true;//CheckBoxRecAudio.Checked;  ///  ���������, ���������������� �� ���������� ���������� �� ��������� �� ������� � ���������� ������
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=true;//CheckBoxEnSuperPan.Checked;              ///��������� ������������ ����� ������������ �� ���������� �������� ����� � ������������� ������, ����� ������� ���������� �������� ����� ������ �������� �������� ���� �������. ������ ����, ����� ������������ ������ ��������� ��� ��������� �� �������� �����, ������������ ����� ������� �� ������� ����, � ��������� ������� ���� ����� ������������� �������������� � ���� �����������. SuperPan �� ������������ ����� ������ ��������.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=2;//ComboBoxAuthLevel.ItemIndex;           ///������ ������� �������� ����������� ��� �����������.  (0.1.2)0- ��� �������������� �� �������. 1- �������� ����������� ������� ��������� � ������ ������� ����������� ��� ����������� �����������. 2- ������� �������������� �������. ���� �������������� �� �������, ������������ ����� ���������� �������� ���������� ��� ���������� ��� �������������� �� �������.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=true;//CheckBoxCredSsp.Checked;       ///��������� ��� ���������, ������� �� ��� ����� �����������  ��������� ������ ������������ ������� ������ (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=1;//ComboKey.ItemIndex; // 0-���������� ���������� ������ ������ �������� �� ���������� ����������. 1- ���������� ���������� ������ �� ��������� �������. 2- ���������� ���������� ������ � ���������� ������� ������ �����, ����� ������ �������� � ������������� ������. ��� �������� �� ���������.
//FormForRDP.WindowState:=wsNormal;
FormForRDP.Show;
MyRDPWin.Align:=alClient;
//MyRDPWin.Width:=FormForRDP.Width;
//MyRDPWin.Height:=FormForRDP.Height;
MyRDPWin.Connect;
CoUninitialize;
 Except
   on E: Exception do
      begin
       frmdomaininfo.memo1.Lines.add('������ �������� ���������� ����  :'+E.Message);
        if assigned(FormForRDP) then
         if FormForRDP.visible=false then  FormForRDP.release;

      end;
   end;
end;



procedure TRDPWin.N8Click(Sender: TObject);
var
namerdp:string;
i:integer;
begin
try
namerdp:=(Tbutton(PopupOtherWin.PopupComponent).Hint);
if (Tbutton(PopupOtherWin.PopupComponent).owner is TForm) then
begin
   for I := 0 to (Tbutton(PopupOtherWin.PopupComponent).owner as Tform).ComponentCount-1 do // ���� rdp
    if ((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] is TMsRdpClient9  then  // ���� �����
     begin
     if (((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] as TMsRdpClient9).Connected<>0 then // ���� ���������� �����������
      (((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] as TMsRdpClient9).Disconnect;         // �����������
      break;                                                                                                   // ����� �� �����
     end;
 (Tbutton(PopupOtherWin.PopupComponent).owner as TForm).Release;
end;
 Except
   on E: Exception do
       frmdomaininfo.memo1.Lines.add(namerdp+': ������ �������� ���������� ����  :'+E.Message);
   end;
end;


procedure TRDPWin.NewAuthenticationWarningDisplayed(Sender: TObject);
begin
 //���������� �� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).

end;

procedure TRDPWin.MsRdpClient91AutoReconnected(Sender: TObject);
begin
 //����������, ����� ���������� ������� ���������� ������������� �������� ������������ � ���������� ������.
end;

procedure TRDPWin.MsRdpClient91AutoReconnecting(ASender: TObject;
  disconnectReason, attemptCount: Integer);  //����������, ����� ������ ��������� � �������� ��������������� ���������� ����������� ������ � ������� ���� ������� ��������� ������� ������ (RD Session Host).
begin
//Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������������, ������� - '+inttostr(disconnectReason)+'/ ������� - '+ SysErrorMessage(attemptCount));
end;

procedure TRDPWin.MsRdpClient91AutoReconnecting2(ASender: TObject;   //����������, ����� ������ ��������� � �������� ��������������� ���������� ����������� ������ � ������� ���� ������� ��������� ������� ������ (RD Session Host).
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
frmdomaininfo.Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������������:  '+errorstr+
'/ ��������� ���� :'+netW(networkAvailable)+' / ������� - '+inttostr(attemptCount)+' �� '+inttostr(maxAttemptCount));
end;


procedure TRDPWin.MsRdpClient91DevicesButtonPressed(Sender: TObject);
begin
////����� ��� ������ ������ �� �����
end;



procedure TRDPWin.MsRdpClient91ServiceMessageReceived(ASender: TObject;
  const serviceMessage: WideString);
begin
frmdomaininfo.Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP : '+serviceMessage);
end;


procedure TRDPWin.NewRdpAuthenticationWarningDismissed(Sender: TObject);
begin  //���������� ����� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).
try
 if not (sender is TMsRdpClient9) then exit;

{if  ((sender as TMsRdpClient9).Connected=0)and   /// tckb ���������� �� �����������
 ((sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport=false) then  // � ���������  EnableCredSspSupport
 begin
   (sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport:=true;
   (sender as TMsRdpClient9).Connect;
 end;}


frmdomaininfo.Memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������������� ��������������');
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - �������������� ��������������." :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpAuthenticationWarningDisplayed(Sender: TObject);
 const  //���������� �� ����, ��� ������� ���������� ActiveX ���������� ���������� ���� �������������� (��������, ���������� ���� ������ �����������).
  AutType: array [0..4] of string =('No authentication is used.',
  '������������ �������� ����������� �����������','������������ �������� ����������� Kerberos',
  '������������ ��� ���� �������� �����������, ����������� � �� ���������','����������� ���������');
begin
try
  if not (sender is TMsRdpClient9) then  exit;
 if ((sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType<5) then
 frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������� �����������. - '
 + AutType[(sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType]);

 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - �������� �����������." :'+E.Message);
      end;
   end;
end;





procedure TRDPWin.NewRdpAutoReconnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.Memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ��������������� � ���������� ����������.- ');
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - ��������������� � ���������� ����������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpAutoReconnecting(ASender: TObject; disconnectReason,
  attemptCount: Integer);
  var
strerror:string;
id2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
strerror:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - �������������� ��������������� ������ ��������� ������� ������: '+strerror);
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - �������������� ��������������� ������ ��������� ������� ������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpAutoReconnecting2(ASender: TObject;
  disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
  maxAttemptCount: Integer);
var
strerror:string;
id2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
strerror:= (Asender as TMsRdpClient9).GetErrorDescription(id2,disconnectReason);
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ���������� ������� ���������� ���������������� � ���������� ������ :'+strerror);
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - ���������� ������� ���������� ���������������� � ���������� ������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpChannelReceivedData(ASender: TObject; const chanName,
  data: WideString);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ������ �������� ������ �� ����������� ������ � ������������ ��������');
Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - ������ �������� ������ �� ����������� ������ � ������������ ��������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConfirmClose(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ���������� ������ ����� ���������� �������� �����');
Except
   on E: Exception do
      begin
       frmdomaininfo.memo1.Lines.add('������ "RDP - ���������� ������ ����� ���������� �������� �����" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// ����� ����� ������
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - ��������� ����������...');
Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - ��������� ����������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConnecting(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// ���������� ������ ������
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+
' : RDP - ������������� ����������...');
for I := 0 to frmdomaininfo.GroupRDPWin.ControlCount-1 do
begin
  if frmdomaininfo.GroupRDPWin.Controls[i] is TButton then
  if frmdomaininfo.GroupRDPWin.Controls[i].Owner.Name=(sender as TMsRdpClient9).Owner.Name then
   begin
    (frmdomaininfo.GroupRDPWin.Controls[i] as Tbutton).ImageIndex:=7;
   end;
end;

Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('������ "RDP - ������������� ����������" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpDisconnected(ASender: TObject; discReason: Integer);
var
errorstr:string;
disconnectReason:integer;
i:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
errorstr:=(Asender as TMsRdpClient9).GetErrorDescription(disconnectReason,discReason);
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : ���������� ������ - '+errorstr);
if TComponent(Asender).Owner is TTabSheet then  /// ���� �������� ��� ������� TabSheet � �� ��������� ����
TTabSheet(TComponent(Asender).Owner).ImageIndex:=5;
for I := 0 to frmdomaininfo.GroupRDPWin.ControlCount-1 do
begin
  if frmdomaininfo.GroupRDPWin.Controls[i] is TButton then
  if frmdomaininfo.GroupRDPWin.Controls[i].Owner.Name=(Asender as TMsRdpClient9).Owner.Name then
   begin
    (frmdomaininfo.GroupRDPWin.Controls[i] as Tbutton).ImageIndex:=5;
   end;
end;

Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ ���������� ������� :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpExit(Sender: TObject);
begin
try
{if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - ����� ���������'); }

Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ "RDP - ����� ���������" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpFatalError(ASender: TObject; errorCode: Integer);
const
StrError : array [0..7] of string=('��������� ����������� ������.','��� ���������� ������ 1.'
,'��������� ������ �������� ������.','��������� ������ �������� ����.','��� ���������� ������ 2.'
,'���������� ��� ������ 3','��� ���������� ������ 4.','�� ����� ����������� � ������� ��������� ������������ ������.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if errorCode=100 then frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ���������� ������������ ������ - ������ ������������� Winsock.')
else
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP - ���������� ������������ ������ - '+StrError[errorCode]);
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ "RDP - ���������� ������������ ������" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpLoginComplete(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - �������� ���� � �������');

if TComponent(sender).Owner is TTabSheet then   /// ���� �������� ��� ������� TabSheet � �� ��������� ����
begin
TTabSheet(TComponent(sender).Owner).ImageIndex:=4;
 /// ���� ������ ������������ ���������� ��������� ����������� � ����
end;
for I := 0 to frmdomaininfo.GroupRDPWin.ControlCount-1 do
begin
  if frmdomaininfo.GroupRDPWin.Controls[i] is TButton then
  if frmdomaininfo.GroupRDPWin.Controls[i].Owner.Name=(sender as TMsRdpClient9).Owner.Name then
   begin
    (frmdomaininfo.GroupRDPWin.Controls[i] as Tbutton).ImageIndex:=4;
   end;
end;
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ ���������� ������� :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpLogonError(ASender: TObject; lError: Integer);
var
StatusText,ErrorDesc:string;
ID2:integer;
begin
try
if not (Asender is TMsRdpClient9) then  exit;
StatusText:=(Asender as TMsRdpClient9).GetStatusText(lError);
ErrorDesc:=(Asender as TMsRdpClient9).GetErrorDescription(ID2,lError);
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - StatusText - '+ StatusText);
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - ErrorDesc - '+ ErrorDesc);

frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+
' : RDP - ������ ����� � ������� - '+ SysErrorMessage(lError));
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ "������ ����� � �������" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpNetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
const
qLevel: array [0..5] of string =('�������� �� ��������...','�������� ����� 512 ����/�',
'�������� �� 512 ����/� �� 2 ����/�.','�������� �� 2 �� 10 ����/�.'
,'�������� ������ 10 ����/�.','�������� ������ 100 ����/�.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - ��������� ��������� ��������� ����');
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : '+qLevel[qualityLevel]
+' / ���������� ����������� (Bandwidth) - '+inttostr(bandwidth)+' ����/� / �������� ���������� (RTT) - '+inttostr(rtt)+' ��.')
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ ��� ������ ������������� ��������� ��������� ���� :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpWarning(ASender: TObject; warningCode: Integer);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if warningCode=1 then
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP -  ��� ���������� ����������� ���������.');
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('������ ��� ���������� OnWarning :'+E.Message);
  end;
end;

end;

end.
