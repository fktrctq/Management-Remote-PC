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
   procedure PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
   procedure PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// перемещение панели по  форме
  Y: Integer);
   procedure PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
   procedure ButtonOnForForm(Sender: TObject);
   procedure ButtonOffForForm(Sender: TObject);
   procedure ButtonMinimizeForForm(Sender: TObject); //// свернуть отдельную форму
   procedure ButtonResizeForForm(Sender: TObject); //// на весь экран  RDP на отдельной форму
   procedure ButtonCloseForForm(Sender: TObject); //// закрыть отдельную форму
   procedure ShowButtonOnClick(Sender: TObject); /// кнопка на панели GroupBox8 для разворачивания и сворачивания окон
   procedure ButtonMaxNormalForForm(Sender: TObject); //// развернуть или свернуть отдельную форму
   procedure ButonMoveWinMouseDown(Sender: TObject;  /// перемещение формы за кнопку
                           Button: TMouseButton;
                            Shift: TShiftState;
                             X, Y: Integer);
   procedure RDPFormShow(Sender: TObject);  // сворачивание окна рдп
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

procedure TRDPWin.OtherWinForRDPClientClose(Sender: TObject; var Action: TCloseAction); /// уничтожение формы при закрытии
begin
if Sender is Tform then
(Sender as TForm).Release;
end;

procedure TRDPWin.PanelForButMouseDown(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
begin
     movingPan:=true;
     x0:=x;
     y0:=y;
end;

procedure TRDPWin.PanelForButMouseMove(Sender: TObject; Shift: TShiftState; X,// перемещение панели по  форме
  Y: Integer);
begin
     if movingPan then
   begin
     (sender as tpanel).Left:=(sender as tpanel).Left+x-x0;
     (sender as tpanel).Top:=(sender as tpanel).Top+y-y0;
   end;
end;

procedure TRDPWin.PanelForButMouseUp(Sender: TObject; Button: TMouseButton; // перемещение панели по  форме
  Shift: TShiftState; X, Y: Integer);
begin
  movingPan := false;
end;

procedure TRDPWin.ButtonOnForForm(Sender: TObject); //// подключение по RDP на отдельной форму
var
i:integer;
namePC:string;
LevelAut:integer;
begin
try
LevelAut:=2; /// предупреждать
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
   frmDomainInfo.memo1.Lines.Add('Ошибка при подключении к компьютеру '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonOffForForm(Sender: TObject); //// отключение от RDP на отдельной форму
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
   frmDomainInfo.Memo1.Lines.Add('Ошибка при отключении от компьютера '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonResizeForForm(Sender: TObject); //// на весь экран  RDP на отдельной форму
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
   frmDomainInfo.Memo1.Lines.Add('Ошибка при переподключении '+namePC+' - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonCloseForForm(Sender: TObject); //// закрыть отдельную форму
var
i:integer;
begin
try
if TComponent(sender).Owner is TPanel then
if (TComponent(sender).Owner as TPanel).Owner is Tform then
begin
  for I := 0 to ((TComponent(sender).Owner as TPanel).Owner as Tform).ComponentCount-1 do // ищем rdp
    if ((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] is TMsRdpClient9  then  // если нашли
     begin
     if (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Connected<>0 then // если соединение установлено
     begin
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Disconnect;         // отключаемся
      (((TComponent(sender).Owner as TPanel).Owner as Tform).Components[i] as TMsRdpClient9).Free;
      break;
      end;                                                                                                   // выход из цикла
     end;
 ((TComponent(sender).Owner as TPanel).Owner as Tform).Release;   // уничтожаем форму
end;
except
 on E:Exception do
  begin
   frmDomainInfo.Memo1.Lines.Add('Ошибка при закрытии подключения  - '+e.Message)
  end;
end;
end;


procedure TRDPWin.ShowButtonOnClick(Sender: TObject); /// кнопка на панели panel7 для разворачивания и сворачивания окон
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
   frmDomainInfo.Memo1.Lines.Add('Ошибка при сворачивании/разворачивании окна - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonMinimizeForForm(Sender: TObject); //// свернуть отдельную форму
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
   frmDomainInfo.Memo1.Lines.Add('Ошибка при сворачивании окна - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButtonMaxNormalForForm(Sender: TObject); //// развернуть или свернуть отдельную форму
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
   frmDomainInfo.Memo1.Lines.Add('Ошибка при изменении размера окна - '+e.Message)
  end;
end;
end;

procedure TRDPWin.ButonMoveWinMouseDown(Sender: TObject;  /// перемещение формы за кнопку
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
PanelForBut.Hint:='Подключение к '+NameCapt+'. Нажмите для перемещения';
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
ButFP.Name:='ButFPOFF'+inttostr(indTab);
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
ButFP.Name:='ButFPResize'+inttostr(indTab);
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
AutLevel.Name:='Level'+inttostr(indTab);
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
ButFP.Name:='MoveWin'+inttostr(indTab);
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
ButFP.Name:='ButFPCollapse'+inttostr(indTab);
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
ButFP.Name:='ButFPCollapseWin'+inttostr(indTab);
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
ButFP.Name:='ButFPClose'+inttostr(indTab);
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

//{if GroupBox8.Caption='' then GroupBox8.Caption:='Подключения в отдельном окне';

ButFP:=TButton.Create(FormForRDP);
ButFP.Parent:=frmDomainInfo.GroupRDPWin;
ButFP.Name:='ShowB'+inttostr(indTab);
ButFP.Caption:=namePC;
ButFP.Hint:='подключение к '+NameCapt;
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
MyRDPWin.SendToBack;/// Для того чтобы было видно панель
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
MyRDPWin.ConnectingText:='Установка соединения c ' +namePC+'...';
MyRDPWin.DisconnectedText:='Клиент '+namePC+' отключен';

step:='';
//MyRDPWin.ColorDepth:=bitcolorDepth(ComboBoxColorDepth.ItemIndex);
step:='BitmapPeristence';
MyRDPWin.AdvancedSettings9.BitmapPeristence:=1;//booltoint(CheckBoxBitmapPeristence.Checked); // 0 выключить,1 включить.  Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='CachePersistenceActive';
//MyRDPWin.AdvancedSettings9.CachePersistenceActive:=booltoint(CheckboxCachePersistenceActive.Checked); // 0 - отключить, 1 - включить. Указывает, следует ли использовать постоянное растровое кэширование. Постоянное кэширование может повысить производительность, но требует дополнительного дискового пространства.
step:='BitmapCacheSize';
//MyRDPWin.AdvancedSettings9.BitmapCacheSize:=SpinBitmapCacheSize.Value;; // от 1 до 32 . Размер файла кэша растровых изображений в килобайтах, используемого для растровых изображений 8 бит на пиксель
step:='VirtualCache16BppSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCache16BppSize:=SpinEditCache16BppSize.Value; // Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCache32BppSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCache32BppSize:=SpinEditCache32BppSize.Value;// Новый размер кеша. Допустимые значения - от 1 до 32 включительно, а значение по умолчанию - 20. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ.
step:='VirtualCacheSize';
//MyRDPWin.AdvancedSettings9.BitmapVirtualCacheSize:=SpinBitmapVirtualCacheSize.Value;  //Задает размер файла постоянного кэша растрового изображения в мегабайтах, который используется для цвета 8 бит на пиксель. Допустимые числовые значения этого свойства - от 1 до 32 включительно. Обратите внимание, что максимальный размер всех файлов виртуального кэша составляет 128 МБ. Связанные свойства включают свойства BitmapVirtualCache16BppSize и BitmapVirtualCache24BppSize .
step:='DisableCtrlAltDel';
//MyRDPWin.AdvancedSettings9.DisableCtrlAltDel:=booltoint(CheckBoxDisableCtrlAltDel.Checked); // 0 - отключить, 1 - включить.  Указывает, должен ли отображаться начальный пояснительный экран в Winlogon.
step:='SmartSizing';
//MyRDPWin.AdvancedSettings9.SmartSizing:=not(CheckBoxSmartSizing.Checked); //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления.
step:='EnableWindowsKey';
MyRDPWin.AdvancedSettings9.EnableWindowsKey:=1;//booltoint(CheckBoxEnableWindowsKey.Checked);   //0 - отключить, 1 - включить. Указывает, можно ли использовать ключ Windows в удаленном сеансе.
step:='GrabFocusOnConnect';
//MyRDPWin.AdvancedSettings9.GrabFocusOnConnect:=checkboxGrabFocusOnConnect.Checked; //0 - отключить, 1 - включить. Указывает, должен ли элемент управления клиента иметь фокус при подключении. Элемент управления не будет пытаться захватить фокус из окна, запущенного в другом процессе.
step:='MinutesToIdleTimeout';
//MyRDPWin.AdvancedSettings9.MinutesToIdleTimeout:= SpinMinutesToIdleTimeout.Value;  //По умолчанию 0 только т.е. не отслеживать время. Задает максимальный промежуток времени в минутах, в течение которого клиент должен оставаться подключенным без участия пользователя. Если указанное время истекло, элемент управления вызывает метод IMsTscAxEvents :: OnIdleTimeoutNotification .
step:='OverallConnectionTimeout';
//MyRDPWin.AdvancedSettings9.OverallConnectionTimeout:= SpinOverallConnectionTimeout.Value; //Определяет общую продолжительность времени в секундах, в течение которого клиентский элемент управления ожидает завершения соединения. Максимальное допустимое значение этого свойства - 600, что соответствует 10 минутам. Если указанное время истекает до завершения соединения, элемент управления отключается и вызывает метод IMsTscAxEvents :: OnDisconnected
step:='RdpPort';
//MyRDPWin.AdvancedSettings9.RdpPort:=strtoint(LabeledEditRdpPort.Text);


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

//  if ComboBoxNetworkConnectionType.ItemIndex=0 then
begin
MyRDPWin.AdvancedSettings9.BandwidthDetection:=true;          /// если NetworkConnectionType = 0 то....яет, будут ли автоматически обнаружены изменения пропускной способности.
//MyRDPWin.AdvancedSettings9.PerformanceFlags:=0;
end ;
{else
begin
step:='NetworkConnectionType';
MyRDPWin.AdvancedSettings9.NetworkConnectionType:= ComboBoxNetworkConnectionType.ItemIndex; // установка типа используемой сети 1- модем, 2-Низкоскоростная широкополосная связь (от 256 Кбит / с до 2 Мбит / с), 3Спутник (от 2 Мбит / с до 16 Мбит / с, с высокой задержкой), 4 -Высокоскоростной широкополосный доступ (от 2 Мбит / с до 10 Мбит / с), 5- Глобальная сеть (WAN) (10 Мбит / с или выше, с высокой задержкой), 6 -Локальная сеть (LAN) (10 Мбит / с или выше)
step:='PerformanceFlags';
MyRDPWin.AdvancedSettings9.PerformanceFlags:=CalcPerformanceFlags(''); //  Задает набор функций, которые можно установить на сервере для повышения производительности.
end; }

//MyRDPWin.AdvancedSettings9.SmartSizing:=true; //Указывает, следует ли масштабировать дисплей, чтобы он соответствовал клиентской области элемента управления
step:='MaxReconnectAttempts';
MyRDPWin.AdvancedSettings9.MaxReconnectAttempts:= 20;//SpinMaxReconnectAttempts.Value;   //Задает количество попыток переподключения во время автоматического переподключения. Допустимые значения этого свойства - от 0 до 200 включительно.
step:='EnableAutoReconnect';
MyRDPWin.AdvancedSettings9.EnableAutoReconnect:=true;// CheckEnableAutoReconnect.Checked;       //Указывает, следует ли разрешить клиентскому элементу управления автоматически подключаться к сеансу в случае отключения сети.
step:='AudioRedirectionMode';
MyRDPWin.SecuredSettings3.AudioRedirectionMode:=0;//ComboAudioRedirectionMode.ItemIndex;  // 0-Перенаправить звуки на клиента. 1-Воспроизведение звуков на удаленном компьютере. 2-Отключить перенаправление звука; не воспроизводить звуки на сервере.
step:='RedirectDrives';
MyRDPWin.AdvancedSettings9.RedirectDrives:=true;//CheckBoxDisk.Checked;     ///разрешить перенаправлять диски
step:='RedirectPrinters';
MyRDPWin.AdvancedSettings9.RedirectPrinters:=true;//CheckBoxPrint.Checked;    /// разрешить перенаправлять принтеры
//if CheckBoxClipboard.Checked or CheckBoxPrint.Checked then /// если необходимо перенаправить принтеры и буфер обмена то
MyRDPWin.AdvancedSettings.DisableRdpdr:=0; //// Установите для этого параметра значение 1, чтобы отключить перенаправление, или 0, чтобы включить перенаправление.
step:='RelativeMouseMode';
//MyRDPWin.AdvancedSettings9.RelativeMouseMode:=CheckBoxMouseMode.Checked;  ///
step:='RedirectClipboard';
MyRDPWin.AdvancedSettings9.RedirectClipboard:=true;//CheckBoxClipboard.Checked;  /// Указывает, следует ли включить или отключить перенаправление буфера обмена.
step:='RedirectDevices';
MyRDPWin.AdvancedSettings9.RedirectDevices:=true;//CheckRedirectDevices.Checked;   /// Указывает, должны ли перенаправленные устройства быть включены или отключены.
 step:='RedirectPorts';
MyRDPWin.AdvancedSettings9.RedirectPorts:=true;//CheckBoxPort.Checked;       /// перенаправление портов
step:='ConnectToAdministerServer';
MyRDPWin.AdvancedSettings9.ConnectToAdministerServer:=true;//CheckBoxConToAdmSrv.Checked; /// указывает должен ли элемент управления ActiveX пытаться подключиться к серверу в административных целях.
step:='AudioCaptureRedirectionMode';
MyRDPWin.AdvancedSettings9.AudioCaptureRedirectionMode:=true;//CheckBoxRecAudio.Checked;  ///  указывает, перенаправляется ли устройство аудиовхода по умолчанию от клиента к удаленному сеансу
step:='EnableSuperPan';
MyRDPWin.AdvancedSettings9.EnableSuperPan:=true;//CheckBoxEnSuperPan.Checked;              ///позволяет пользователю легко перемещаться по удаленному рабочему столу в полноэкранном режиме, когда размеры удаленного рабочего стола больше размеров текущего окна клиента. Вместо того, чтобы использовать полосы прокрутки для навигации по рабочему столу, пользователь может указать на границу окна, и удаленный рабочий стол будет автоматически прокручиваться в этом направлении. SuperPan не поддерживает более одного монитора.
step:='AuthenticationLevel';
MyRDPWin.AdvancedSettings9.AuthenticationLevel:=2;//ComboBoxAuthLevel.ItemIndex;           ///Задает уровень проверки подлинности для подключения.  (0.1.2)0- Нет аутентификации на сервере. 1- Проверка подлинности сервера требуется и должна успешно завершиться для продолжения подключения. 2- Попытка аутентификации сервера. Если аутентификация не удалась, пользователю будет предложено отменить соединение или продолжить без аутентификации на сервере.
step:='EnableCredSspSupport';
MyRDPWin.AdvancedSettings9.EnableCredSspSupport:=true;//CheckBoxCredSsp.Checked;       ///Указывает или извлекает, включен ли для этого подключения  Поставщик службы безопасности учетных данных (CredSSP)
step:='KeyboardHookMode';
MyRDPWin.SecuredSettings3.KeyboardHookMode:=1;//ComboKey.ItemIndex; // 0-Применяйте комбинации клавиш только локально на клиентском компьютере. 1- Применяйте комбинации клавиш на удаленном сервере. 2- Применяйте комбинации клавиш к удаленному серверу только тогда, когда клиент работает в полноэкранном режиме. Это значение по умолчанию.
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
       frmdomaininfo.memo1.Lines.add('Ошибка создания отдельного окна  :'+E.Message);
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
   for I := 0 to (Tbutton(PopupOtherWin.PopupComponent).owner as Tform).ComponentCount-1 do // ищем rdp
    if ((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] is TMsRdpClient9  then  // если нашли
     begin
     if (((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] as TMsRdpClient9).Connected<>0 then // если соединение установлено
      (((Tbutton(PopupOtherWin.PopupComponent).owner) as Tform).Components[i] as TMsRdpClient9).Disconnect;         // отключаемся
      break;                                                                                                   // выход из цикла
     end;
 (Tbutton(PopupOtherWin.PopupComponent).owner as TForm).Release;
end;
 Except
   on E: Exception do
       frmdomaininfo.memo1.Lines.add(namerdp+': Ошибка закрытия отдельного окна  :'+E.Message);
   end;
end;


procedure TRDPWin.NewAuthenticationWarningDisplayed(Sender: TObject);
begin
 //Вызывается до того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).

end;

procedure TRDPWin.MsRdpClient91AutoReconnected(Sender: TObject);
begin
 //Вызывается, когда клиентский элемент управления автоматически повторно подключается к удаленному сеансу.
end;

procedure TRDPWin.MsRdpClient91AutoReconnecting(ASender: TObject;
  disconnectReason, attemptCount: Integer);  //Вызывается, когда клиент находится в процессе автоматического повторного подключения сеанса к серверу узла сеансов удаленных рабочих столов (RD Session Host).
begin
//Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Переподключение, попытка - '+inttostr(disconnectReason)+'/ Причина - '+ SysErrorMessage(attemptCount));
end;

procedure TRDPWin.MsRdpClient91AutoReconnecting2(ASender: TObject;   //Вызывается, когда клиент находится в процессе автоматического повторного подключения сеанса к серверу узла сеансов удаленных рабочих столов (RD Session Host).
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
frmdomaininfo.Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Переподключение:  '+errorstr+
'/ Состояние сети :'+netW(networkAvailable)+' / Попытка - '+inttostr(attemptCount)+' из '+inttostr(maxAttemptCount));
end;


procedure TRDPWin.MsRdpClient91DevicesButtonPressed(Sender: TObject);
begin
////вроде как нажати кнопок на клаве
end;



procedure TRDPWin.MsRdpClient91ServiceMessageReceived(ASender: TObject;
  const serviceMessage: WideString);
begin
frmdomaininfo.Memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP : '+serviceMessage);
end;


procedure TRDPWin.NewRdpAuthenticationWarningDismissed(Sender: TObject);
begin  //Вызывается после того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).
try
 if not (sender is TMsRdpClient9) then exit;

{if  ((sender as TMsRdpClient9).Connected=0)and   /// tckb соединение не установлено
 ((sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport=false) then  // и отключена  EnableCredSspSupport
 begin
   (sender as TMsRdpClient9).AdvancedSettings9.EnableCredSspSupport:=true;
   (sender as TMsRdpClient9).Connect;
 end;}


frmdomaininfo.Memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - Предупреждение Аутентификации');
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - Предупреждение Аутентификации." :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpAuthenticationWarningDisplayed(Sender: TObject);
 const  //Вызывается до того, как элемент управления ActiveX отображает диалоговое окно аутентификации (например, диалоговое окно ошибки сертификата).
  AutType: array [0..4] of string =('No authentication is used.',
  'Используется проверка подлинности сертификата','Используется проверка подлинности Kerberos',
  'Используется оба типа проверки подлинности, сертификата и по протоколу','неизвестное состояние');
begin
try
  if not (sender is TMsRdpClient9) then  exit;
 if ((sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType<5) then
 frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - Проверка подлинности. - '
 + AutType[(sender as TMsRdpClient9).AdvancedSettings9.AuthenticationType]);

 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - Проверка подлинности." :'+E.Message);
      end;
   end;
end;





procedure TRDPWin.NewRdpAutoReconnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.Memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - Переподключение к клиентсому компьютеру.- ');
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - Переподключение к клиентсому компьютеру" :'+E.Message);
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
+' : RDP - Автоматическое переподключение сеанса удаленных рабочих столов: '+strerror);
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - Автоматическое переподключение сеанса удаленных рабочих столов" :'+E.Message);
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
+' : RDP - клиентский элемент управления переподключается к удаленному сеансу :'+strerror);
 Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - клиентский элемент управления переподключается к удаленному сеансу" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpChannelReceivedData(ASender: TObject; const chanName,
  data: WideString);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - клиент получает данные на виртуальном канале с возможностью сценария');
Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - клиент получает данные на виртуальном канале с возможностью сценария" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConfirmClose(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - Завершение сеанса служб удаленного рабочего стола');
Except
   on E: Exception do
      begin
       frmdomaininfo.memo1.Lines.add('Ошибка "RDP - Завершение сеанса служб удаленного рабочего стола" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConnected(Sender: TObject);
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// после ввода пароля
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server
+' : RDP - установка соединения...');
Except
   on E: Exception do
      begin
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - установка соединения" :'+E.Message);
      end;
   end;
end;

procedure TRDPWin.NewRdpConnecting(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
if TComponent(sender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
TTabSheet(TComponent(sender).Owner).ImageIndex:=7; /// сдесьперед вводом пароля
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+
' : RDP - инициализация соединения...');
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
        frmdomaininfo.memo1.Lines.add('Ошибка "RDP - инициализация соединения" :'+E.Message);
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
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : Завершение сеанса - '+errorstr);
if TComponent(Asender).Owner is TTabSheet then  /// если родитель рдп клиента TabSheet а не отдельное окно
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
  frmdomaininfo.memo1.Lines.add('Ошибка обновления статуса :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpExit(Sender: TObject);
begin
try
{if not (sender is TMsRdpClient9) then  exit;
memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - фокус перемещен'); }

Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('Ошибка "RDP - фокус перемещен" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpFatalError(ASender: TObject; errorCode: Integer);
const
StrError : array [0..7] of string=('Произошла неизвестная ошибка.','Код внутренней ошибки 1.'
,'Произошла ошибка нехватки памяти.','Произошла ошибка создания окна.','Код внутренней ошибки 2.'
,'Внутренний код ошибки 3','Код внутренней ошибки 4.','Во время подключения к клиенту произошла неисправимая ошибка.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if errorCode=100 then frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - произошщла неустранимая ошибка - Ошибка инициализации Winsock.')
else
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : RDP - произошщла неустранимая ошибка - '+StrError[errorCode]);
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('Ошибка "RDP - произошщла неустранимая ошибка" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpLoginComplete(Sender: TObject);
var
i:integer;
begin
try
if not (sender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((sender as TMsRdpClient9).Server+' : RDP - успешный вход в систему');

if TComponent(sender).Owner is TTabSheet then   /// если родитель рдп клиента TabSheet а не отдельное окно
begin
TTabSheet(TComponent(sender).Owner).ImageIndex:=4;
 /// если удачно подключились записываем настройки подключения в базу
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
  frmdomaininfo.memo1.Lines.add('Ошибка обновления статуса :'+E.Message);
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
' : RDP - ошибка входа в систему - '+ SysErrorMessage(lError));
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('Ошибка "ошибка входа в систему" :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpNetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
const
qLevel: array [0..5] of string =('Скорость не известна...','Скорость менее 512 Кбит/с',
'Скорость от 512 Кбит/с до 2 Мбит/с.','Скорость от 2 до 10 Мбит/с.'
,'Скорость больше 10 Мбит/с.','Скорость больше 100 Мбит/с.');
begin
try
if not (Asender is TMsRdpClient9) then  exit;
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP - произошло изменение состояния сети');
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server+' : '+qLevel[qualityLevel]
+' / Пропускная способность (Bandwidth) - '+inttostr(bandwidth)+' Кбит/с / Задержка соединения (RTT) - '+inttostr(rtt)+' мс.')
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('Ошибка при чтении характеристик изменения состояния сети :'+E.Message);
  end;
end;
end;

procedure TRDPWin.NewRdpWarning(ASender: TObject; warningCode: Integer);
begin
try
if not (Asender is TMsRdpClient9) then  exit;
if warningCode=1 then
frmdomaininfo.memo1.Lines.Add((Asender as TMsRdpClient9).Server
+' : RDP -  Кэш растрового изображения поврежден.');
Except
  on E: Exception do
  begin
  frmdomaininfo.memo1.Lines.add('Ошибка при оповещении OnWarning :'+E.Message);
  end;
end;

end;

end.
