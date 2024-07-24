unit MRDUNew;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, RDPCOMAPILib_TLB,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, System.Win.ScktComp,IniFiles,
  Vcl.Menus, Vcl.Samples.Spin, Vcl.ComCtrls, Vcl.ImgList;

type
  TMRDForm = class(TForm)
    Panel1: TPanel;
    SpeedButton10: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    ClientSocket1: TClientSocket;
    SpeedButton11: TSpeedButton;
    SpeedButton1: TSpeedButton;
    EditUser: TEdit;
    EditPass: TEdit;
    EditPort: TEdit;
    Panel2: TPanel;
    LabelEdEdit1: TComboBoxEx;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    ImageButton: TImageList;
    Panel3: TPanel;
    ImageList1: TImageList;
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure ClientSocket1Connect(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocket1Connecting(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Disconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Error(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure ClientSocket1Read(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocket1Write(Sender: TObject; Socket: TCustomWinSocket);
    procedure RDPViewer1AttendeeConnected(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPViewer1AttendeeDisconnected(ASender: TObject;
      const pDisconnectInfo: IDispatch);
    procedure RDPViewer1AttendeeUpdate(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPViewer1ChannelDataReceived(ASender: TObject;
      const pChannel: IInterface; lAttendeeId: Integer;
      const bstrData: WideString);
    procedure RDPViewer1ChannelDataSent(ASender: TObject;
      const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
    procedure RDPViewer1ConnectionAuthenticated(Sender: TObject);
    procedure RDPViewer1ConnectionEstablished(Sender: TObject);
    procedure RDPViewer1ConnectionFailed(Sender: TObject);
    procedure RDPViewer1ConnectionTerminated(ASender: TObject; discReason,
      ExtendedInfo: Integer);
    procedure RDPViewer1ControlLevelChangeRequest(ASender: TObject;
      const pAttendee: IDispatch; RequestedLevel: TOleEnum);
    procedure RDPViewer1FocusReleased(ASender: TObject; iDirection: Integer);
    procedure RDPViewer1GraphicsStreamPaused(Sender: TObject);
    procedure RDPViewer1GraphicsStreamResumed(Sender: TObject);
    procedure RDPViewer1SharedDesktopSettingsChanged(ASender: TObject; width,
      height, colordepth: Integer);
    procedure RDPViewer1SharedRectChanged(ASender: TObject; left, top, right,
      bottom: Integer);
    procedure RDPViewer1StartDrag(Sender: TObject; var DragObject: TDragObject);
    procedure RDPViewer1WindowClose(ASender: TObject; const pWindow: IDispatch);
    procedure RDPViewer1WindowOpen(ASender: TObject; const pWindow: IDispatch);
    procedure RDPViewer1WindowUpdate(ASender: TObject;
      const pWindow: IDispatch);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure createviewer;
    procedure N161Click(Sender: TObject);
    procedure N241Click(Sender: TObject);
    procedure LabeledEdit5KeyPress(Sender: TObject; var Key: Char);
    procedure EditPassKeyPress(Sender: TObject; var Key: Char);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Panel2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure createbutinpanel7;
    procedure ShowButtonOnClick(Sender: TObject);
    procedure Panel2DblClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;
var
  MRDForm: TMRDForm;
  UnloadFile:TMemInifile;
  MRDuRDPViewer:TRDPViewer;
  ButShowFP:Tbutton;
implementation
uses umain,RemoteDesktopSettingDialog;
{$R *.dfm}
var
  ConnectionEnable:bool;
  AccessSettingLevel,levelCntrl,UnloadFileSetting:boolean;
  Stt:twindowstate;

function Crypt(varStr: WideString):WideString;
var
 k: integer;
 s: WideString;
begin
   RandSeed:=100;
   s:=varStr;
   for k:=1 to Length(s) do
    s[k]:=Chr(ord(s[k]) xor (Random(127)+1));

 Crypt:=s;
end;

procedure TMRDForm.ShowButtonOnClick(Sender: TObject); /// кнопка на панели panel7 для разворачивания и сворачивания окон
begin
try
if sender is Tbutton then
begin
if  (sender as Tbutton).Owner is Tform then
begin
  if ((sender as Tbutton).Owner as Tform).WindowState=wsMinimized then
    begin
    ((sender as Tbutton).Owner as Tform).WindowState:=Stt;
    end
    else ((sender as Tbutton).Owner as Tform).WindowState:=wsMinimized;
end;
end;
except
 on E:Exception do
  begin
   frmDomainInfo.Memo1.Lines.Add('Ошибка при разворачивании окна uRDM - '+e.Message)
  end;
end;
end;

procedure TMRDForm.createbutinpanel7; // процедура создания кнопки на панели при сворачивании формы
begin
try
ButShowFP:=TButton.Create(MRDForm);
ButShowFP.Parent:=frmDomainInfo.GroupRDPWin;
ButShowFP.Name:='ShowBMRDP';
ButShowFP.Caption:=LabelEdEdit1.Text;
ButShowFP.Hint:='подключение к '+LabelEdEdit1.Text;
ButShowFP.ShowHint:=true;
ButShowFP.Images:=ImageList1;
ButShowFP.ImageIndex:=0;
ButShowFP.Height:=20;
ButShowFP.Align:=alBottom;
ButShowFP.OnClick:=ShowButtonOnClick;
except
 on E:Exception do
  begin
   frmDomainInfo.Memo1.Lines.Add('Ошибка создания кнопки для uRDM - '+e.Message)
  end;
end;
end;


procedure TMRDForm.Button1Click(Sender: TObject); // свернуть форму
begin
stt:=MRDForm.WindowState;
MRDForm.WindowState:=wsMinimized;


end;

procedure TMRDForm.Button2Click(Sender: TObject);  // развернуть форму   к окну и максимально
begin
if MRDForm.WindowState=wsMaximized then
 MRDForm.WindowState:=wsnormal
else
 MRDForm.WindowState:=wsMaximized;
end;

procedure TMRDForm.Button3Click(Sender: TObject);
begin
try
close;
freeandnil( ButShowFP);
 Except
   on E: Exception do  FrmDomainInfo.memo1.Lines.Add('Упс... ошибка закрытия окна');
 end;
end;

procedure TMRDForm.ClientSocket1Connect(Sender: TObject;
  Socket: TCustomWinSocket);
var
  s:WideString;
begin
FrmDomainInfo.memo1.Lines.Add('Соединение установлено, отправляю запрос на сервер');
s:=Widestring('<Client>'+EditUser.Text+'<passwd>'+EditPass.Text);
s:=crypt(s);//// шифрация текста
ClientSocket1.Socket.SendText(s);  /// отправляем зашифрованное сообщение
ConnectionEnable:=true;
end;

procedure TMRDForm.ClientSocket1Connecting(Sender: TObject;
  Socket: TCustomWinSocket);
begin
FrmDomainInfo.memo1.Lines.Add('Устанавливаю соединение');
end;

procedure TMRDForm.ClientSocket1Disconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
FrmDomainInfo.Memo1.Lines.Add('Отключение от сервера');
AccessSettingLevel:=false;
ConnectionEnable:=false;
end;

procedure TMRDForm.ClientSocket1Error(Sender: TObject; Socket: TCustomWinSocket;
  ErrorEvent: TErrorEvent; var ErrorCode: Integer);
begin
FrmDomainInfo.Memo1.Lines.add('Ошибка подключения к серверу: ' + SysErrorMessage(ErrorCode));
ErrorCode:=0;
//speedbutton9.Enabled:=true;
ClientSocket1.Close;       ///// ???????????
//MRDuRDPViewer.Disconnect;
// AccessSettingLevel:=false;
ConnectionEnable:=false;
end;

procedure TMRDForm.ClientSocket1Read(Sender: TObject; Socket: TCustomWinSocket);
var
  S:Widestring;
  UnloadList:TstringList;
  user:Widestring;
  z:integer;
begin
try
 z:=0;
 user:=WideString(EditUser.Text);
  FrmDomainInfo.Memo1.Lines.Add('Пришел ответ от сервера - '+(sender as TClientSocket).Host);
  s:=(Socket.ReceiveText);  /// ответ от сервера,  строка соединения
  s:=crypt(s);  //// дешифрация ответа
  if s=WideString('<ErrorUser>') then    /// ответ от сервера если ошибка
    begin
    FrmDomainInfo.Memo1.Lines.Add('Неверный пароль или имя пользователя');
    ClientSocket1.Active:=False;
    MRDuRDPViewer.Disconnect;
    ConnectionEnable:=false;
    exit;
    end;
  if s=WideString('<unloadlist>') then   //// признак получения файла настроек
    begin
    UnloadFileSetting:=true;
    exit;
    end;
  if UnloadFileSetting then   //// получаем файл настроек
    begin
     UnloadList:=TstringList.Create;
     UnloadList.SetText(pchar(s));
     UnLoadFile:=TMemIniFile.Create(ChangeFileExt( Application.Exename,'UnLF.ini'));
     UnLoadFile.SetStrings(UnloadList);
     UnloadList.Free;
     UnloadFileSetting:=false;
     AccessSettingLevel:=true;
     exit;
    end;


z:=pos('RemoteComputer',string(s));
if (z<>0) then
   begin
   FrmDomainInfo.Memo1.Lines.Add('Подключаю удаленный рабочий стол - '+(sender as TClientSocket).Host);
   MRDuRDPViewer.Connect(s,user,WideString('ADMIN')); /// подключаюсь
   MRDuRDPViewer.SmartSizing:=true;
   end;

  Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Ошибка получения данных от сервера :'+E.Message);
       if ConnectionEnable=true then
         begin
           ClientSocket1.Close;
           ConnectionEnable:=false;
         end;

      end;
   end;


end;

procedure TMRDForm.ClientSocket1Write(Sender: TObject;
  Socket: TCustomWinSocket);
begin
FrmDomainInfo.Memo1.Lines.Add('Отправляю данные на сервер - '+(sender as TClientSocket).Host );
end;

procedure TMRDForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
try
if ConnectionEnable then
  begin
  ClientSocket1.Close;
  MRDuRDPViewer.Disconnect;
  AccessSettingLevel:=false;
if Assigned(UnLoadFile) then UnLoadFile.Free;
 if Assigned(MRDuRDPViewer) then FreeAndNil(MRDuRDPViewer); /// уничтожаем окно если оно есть

  end;
 Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Непредвиденная ошибка :'+E.Message);
      end;
   end;
end;

procedure TMRDForm.EditPassKeyPress(Sender: TObject; var Key: Char);
begin
if key=#13 then SpeedButton9.Click;
end;

procedure TMRDForm.FormCreate(Sender: TObject);
var
SetIni:TMeminifile;
begin
 speedButton6.Hint:='Масштаб рабочего стола';
 speedButton7.Hint:='Управление Вкл/Выкл';
 speedButton8.Hint:='Настройка прав пользователей';
 speedButton9.Hint:='Подключится';
 speedButton10.Hint:='Отключится';
 speedButton1.Hint:='Открыть Диспетчер задач';
end;
procedure TMRDForm.createviewer;
var
propRdp:IRDPSRAPISessionProperties;
begin
 try
  if not Assigned(MRDuRDPViewer) then
 begin
 MRDuRDPViewer:=TRDPViewer.Create(panel1);
 MRDuRDPViewer.Parent:=Panel1;
 MRDuRDPViewer.Name:='MRDuRDPViewerNew';
 MRDuRDPViewer.Align:=alClient;
 MRDuRDPViewer.OnAttendeeConnected:=RDPViewer1AttendeeConnected;
 MRDuRDPViewer.OnAttendeeDisconnected:=RDPViewer1AttendeeDisconnected;
 MRDuRDPViewer.OnAttendeeUpdate:=RDPViewer1AttendeeUpdate;
 MRDuRDPViewer.OnChannelDataReceived:=RDPViewer1ChannelDataReceived;
 MRDuRDPViewer.OnChannelDataSent:= RDPViewer1ChannelDataSent;
 MRDuRDPViewer.OnConnectionAuthenticated:=RDPViewer1ConnectionAuthenticated;
 MRDuRDPViewer.OnConnectionEstablished:=RDPViewer1ConnectionEstablished;
 MRDuRDPViewer.OnConnectionFailed:=RDPViewer1ConnectionFailed;
 MRDuRDPViewer.OnConnectionTerminated:=RDPViewer1ConnectionTerminated;
 MRDuRDPViewer.OnControlLevelChangeRequest:=RDPViewer1ControlLevelChangeRequest;
 MRDuRDPViewer.OnFocusReleased:=RDPViewer1FocusReleased;
 MRDuRDPViewer.OnGraphicsStreamPaused:=RDPViewer1GraphicsStreamPaused;
 MRDuRDPViewer.OnGraphicsStreamResumed:=RDPViewer1GraphicsStreamResumed;
 MRDuRDPViewer.OnSharedDesktopSettingsChanged:=RDPViewer1SharedDesktopSettingsChanged;
 MRDuRDPViewer.OnSharedRectChanged:=RDPViewer1SharedRectChanged;
 MRDuRDPViewer.OnStartDrag:=RDPViewer1StartDrag;
 MRDuRDPViewer.OnWindowClose:=RDPViewer1WindowClose;
 MRDuRDPViewer.OnWindowOpen:= RDPViewer1WindowOpen;
 MRDuRDPViewer.OnWindowUpdate:=RDPViewer1WindowUpdate;
 MRDuRDPViewer.DisconnectedText:='Клиент отключен';
  //MRDuRDPViewer.ControlInterface.RequestColorDepthChange(16);

 end;
except
 on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Ошибка создания окна :'+E.Message);
      end;
 end;
end;



procedure TMRDForm.FormShow(Sender: TObject);
begin
 MRDForm.Caption:='uRDM';
 ConnectionEnable:=false;
 createbutinpanel7;/// создаем кнопку на панели
 end;


procedure TMRDForm.LabeledEdit5KeyPress(Sender: TObject; var Key: Char);
begin
if Key=#13 then speedButton9.Click;
end;

procedure TMRDForm.N161Click(Sender: TObject);
begin
if ConnectionEnable then
  begin
  MRDuRDPViewer.ControlInterface.RequestColorDepthChange(16);
  end;
end;

procedure TMRDForm.N241Click(Sender: TObject);
begin
if ConnectionEnable then
  begin
  MRDuRDPViewer.ControlInterface.RequestColorDepthChange(24);
  end;
end;

procedure TMRDForm.Panel2DblClick(Sender: TObject);
begin
if MRDForm.BorderStyle=bsNone then MRDForm.BorderStyle:=bsSizeToolWin
else MRDForm.BorderStyle:=bsNone;

end;

procedure TMRDForm.Panel2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
FrmHND:Thandle;
begin

 FrmHND:=MRDForm.Handle; // получаем хендл формы для перемешения
 ReleaseCapture;
 SendMessage(FrmHND, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TMRDForm.RDPViewer1AttendeeConnected(ASender: TObject;
  const pAttendee: IDispatch);
begin
FrmDomainInfo.memo1.Lines.Add('Подключаюсь к сеансу');
ConnectionEnable:=true;
end;

procedure TMRDForm.RDPViewer1AttendeeDisconnected(ASender: TObject;
  const pDisconnectInfo: IDispatch);
 var
  myclient:IRDPSRAPIAttendeeDisconnectInfo;
begin
///// какой пользователь отключился от сеанса
pDisconnectInfo.QueryInterface(IID_IRDPSRAPIAttendeeDisconnectInfo,Myclient);
//if LabelEdEdit4.Text<>string(Myclient.Attendee.RemoteName) then
FrmDomainInfo.memo1.Lines.Add(string(Myclient.Attendee.RemoteName)+' отключился от сеанса');
end;

procedure TMRDForm.RDPViewer1AttendeeUpdate(ASender: TObject;
  const pAttendee: IDispatch);
begin
//FrmDomainInfo.memo1.Lines.Add('изменяется одно из значений свойств для участника');
end;

procedure TMRDForm.RDPViewer1ChannelDataReceived(ASender: TObject;
  const pChannel: IInterface; lAttendeeId: Integer; const bstrData: WideString);
begin
//FrmDomainInfo.memo1.Lines.Add('данные принимаются от участника');
end;

procedure TMRDForm.RDPViewer1ChannelDataSent(ASender: TObject;
  const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
begin
//FrmDomainInfo.memo1.Lines.Add('данные отправляются клиенту');
end;

procedure TMRDForm.RDPViewer1ConnectionAuthenticated(Sender: TObject);
begin
//FrmDomainInfo.memo1.Lines.Add('соединение аутентифицировано');
end;

procedure TMRDForm.RDPViewer1ConnectionEstablished(Sender: TObject);
begin
FrmDomainInfo.memo1.Lines.Add('Соединение с сервером установлено');
//speedButton10.Enabled:=true;
ConnectionEnable:=true;
end;

procedure TMRDForm.RDPViewer1ConnectionFailed(Sender: TObject);
begin
FrmDomainInfo.memo1.Lines.Add('клиент не может подключиться к серверу');
try
  if ConnectionEnable then
    begin
    UnLoadFile.Free;
    clientSocket1.Close;
    MRDuRDPViewer.Disconnect;
    ConnectionEnable:=False;
    end;
Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Непредвиденная ошибка :'+E.Message);
      end;
   end;
end;

procedure TMRDForm.RDPViewer1ConnectionTerminated(ASender: TObject; discReason,
  ExtendedInfo: Integer);
begin
case discReason of
0: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Информайия отсутствует');
1: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Локальное отключение');
2: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Удаленное отключение пользователем');
3:  FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Удаленное отключение сервером');
260:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка поиска имени DNS');
262: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
264: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Время соединения истекло.');
516: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Сбой подключения к Windows.');
518:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
520: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка хост не найден.');
772: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Windows Sockets ошибка вызова.');
774: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
776: FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Указан неверный IP-адрес');
1028:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка вызова Windows Sockets');
1030:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Недействительные данные безопасности.');
1032:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка.');
1286:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Указан недопустимый метод шифрования.');
1288:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка поиска DNS.');
1540:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка сокета Windows Sockets gethostbyname.');
1542:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Недействительные данные безопасности сервера..');
1544:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка внутреннего таймера.');
1796:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Произошел тайм-аут.');
1798:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Не удалось распечатать сертификат сервера.');
2052:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Указан неверный IP-адрес.');
2056:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Согласование лицензии не удалось.');
2310:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка безопасности.');
2308:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Server Winsock закрыт.');
2312:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Тайм-аут лицензирования.');
2566:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка безопасности.');
2822:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка шифрования.');
3078:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка расшифровки.');
3080:FrmDomainInfo.memo1.Lines.Add('Соединение закрыто - Ошибка декомпрессии.');
end;
case ExtendedInfo of
1:FrmDomainInfo.memo1.Lines.Add('Причина - Приложение инициировало отключение.');
2:FrmDomainInfo.memo1.Lines.Add('Причина - Приложение инициировало выход клиента.');
3:FrmDomainInfo.memo1.Lines.Add('Причина - Сервер отключил клиент, потому что клиент простаивал в течение периода времени, превышающего указанный период тайм-аута.');
4:FrmDomainInfo.memo1.Lines.Add('Причина - Сервер отключил клиент, потому что клиент превысил период, указанный для подключения.');
5:FrmDomainInfo.memo1.Lines.Add('Причина - Соединение клиента было заменено другим соединением.');
6:FrmDomainInfo.memo1.Lines.Add('Причина - Нет памяти.');
7:FrmDomainInfo.memo1.Lines.Add('Причина - Сервер отказал в подключении.');
8:FrmDomainInfo.memo1.Lines.Add('Причина - Сервер отказал в соединении по соображениям безопасности.');
9:FrmDomainInfo.memo1.Lines.Add('Причина - Соединение было отклонено, поскольку учетная запись пользователя не авторизована для удаленного входа в систему.');
11:FrmDomainInfo.memo1.Lines.Add('Причина - Ошибка внутреннего лицензирования.');
12:FrmDomainInfo.memo1.Lines.Add('Причина - Сервер лицензий недоступен.');
13:FrmDomainInfo.memo1.Lines.Add('Причина - Лицензия на программное обеспечение не была доступна.');
14:FrmDomainInfo.memo1.Lines.Add('Причина - Удаленный компьютер получил недопустимое лицензионное сообщение.');
15:FrmDomainInfo.memo1.Lines.Add('Причина - Идентификатор оборудования не совпадает с идентификатором, указанным в лицензии на программное обеспечение.');
16:FrmDomainInfo.memo1.Lines.Add('Причина - Ошибка лицензии клиента.');
17:FrmDomainInfo.memo1.Lines.Add('Причина - Сетевые проблемы возникли во время протокола лицензирования.');
18:FrmDomainInfo.memo1.Lines.Add('Причина - Клиент преждевременно завершил протокол лицензирования.');
20:FrmDomainInfo.memo1.Lines.Add('Причина - Лицензионное сообщение было зашифровано неправильно.');
21:FrmDomainInfo.memo1.Lines.Add('Причина - Лицензия на доступ к локальному компьютеру не может быть обновлена ​​или обновлена.');
22:FrmDomainInfo.memo1.Lines.Add('Причина - Удаленный компьютер не имеет лицензии на прием удаленных подключений.');
23:FrmDomainInfo.memo1.Lines.Add('Причина - Соединение было отклонено, поскольку продавец не смог аутентифицировать зрителя.');
24:FrmDomainInfo.memo1.Lines.Add('Причина - внутренние ошибки протокола');
end;
 try
if ConnectionEnable then
  begin
  ClientSocket1.Close;
  MRDuRDPViewer.Disconnect;
  AccessSettingLevel:=false;
  ConnectionEnable:=false;
  UnLoadFile.Free;
  end;
Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Непредвиденная ошибка :'+E.Message);
      end;
   end;
end;

procedure TMRDForm.RDPViewer1ControlLevelChangeRequest(ASender: TObject;
  const pAttendee: IDispatch; RequestedLevel: TOleEnum);
begin
//// если сервер изменил уровень привелегий
//FrmDomainInfo.memo1.Lines.Add('Уровень привилегий - CTRL_LEVEL_INTERACTIVE - '+ string(RequestedLevel)) ;
//FrmDomainInfo.memo1.Lines.Add('Уровень привилегий - CTRL_LEVEL_VIEW - '+ string(RequestedLevel)) ;

end;

procedure TMRDForm.RDPViewer1FocusReleased(ASender: TObject;
  iDirection: Integer);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1FocusReleased');
end;

procedure TMRDForm.RDPViewer1GraphicsStreamPaused(Sender: TObject);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1GraphicsStreamPaused');
//SpeedButton9.Enabled:=false;
//SpeedButton10.Enabled:=true;
end;

procedure TMRDForm.RDPViewer1GraphicsStreamResumed(Sender: TObject);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1GraphicsStreamResumed');
end;

procedure TMRDForm.RDPViewer1SharedDesktopSettingsChanged(ASender: TObject;
  width, height, colordepth: Integer);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1SharedDesktopSettingsChanged');
end;

procedure TMRDForm.RDPViewer1SharedRectChanged(ASender: TObject; left, top,
  right, bottom: Integer);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1SharedRectChanged');
end;

procedure TMRDForm.RDPViewer1StartDrag(Sender: TObject;
  var DragObject: TDragObject);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1StartDrag');
end;

procedure TMRDForm.RDPViewer1WindowClose(ASender: TObject;
  const pWindow: IDispatch);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1WindowClose');
end;

procedure TMRDForm.RDPViewer1WindowOpen(ASender: TObject;
  const pWindow: IDispatch);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1WindowOpen');
end;

procedure TMRDForm.RDPViewer1WindowUpdate(ASender: TObject;
  const pWindow: IDispatch);
begin
//FrmDomainInfo.memo1.Lines.Add('RDPViewer1WindowUpdate');
end;

procedure TMRDForm.SpeedButton10Click(Sender: TObject);
begin                            /// отключение
try
if ConnectionEnable then
  begin
  ClientSocket1.Close;
  MRDuRDPViewer.Disconnect;
  AccessSettingLevel:=false;
  ConnectionEnable:=false;
  UnLoadFile.Free;
  if Assigned(MRDuRDPViewer) then FreeAndNil(MRDuRDPViewer);  /// уничтожаем окно если оно есть
  end;
  Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Непредвиденная ошибка :'+E.Message);
      end;
   end;
end;

procedure TMRDForm.SpeedButton1Click(Sender: TObject);
var                  //// отправить сочетание клавиш Alt+Ctrl+Del
s:string;
begin
if ClientSocket1.Active then
  begin
  s:=Widestring('<Shift+Ctrl+Esc>');
  s:=crypt(s);//// шифрация текста
  ClientSocket1.Socket.SendText(s);  /// отправляем зашифрованное сообщение
  end
  else
  showmessage('установите соединение с сервером!');

end;

procedure TMRDForm.SpeedButton6Click(Sender: TObject);
begin                             /// масштаб
if not ClientSocket1.Active then
begin
  showmessage('Установите соединение с сервером!');
  exit;
end;
MRDuRDPViewer.SmartSizing:=not MRDuRDPViewer.SmartSizing;
end;

procedure TMRDForm.SpeedButton7Click(Sender: TObject);
begin
if ClientSocket1.Active then
  begin
    if levelCntrl then            ////// изменение уровня привелегий
      begin
      MRDuRDPViewer.RequestControl(CTRL_LEVEL_VIEW);
      levelCntrl:=false;
      end
    else
      begin
      MRDuRDPViewer.RequestControl(CTRL_LEVEL_MAX);
      levelCntrl:=true;
      end;
  end
else showmessage('Установите соединение с сервером!');
end;

procedure TMRDForm.SpeedButton8Click(Sender: TObject);
begin                          //// окно настройки пользователей
if not ClientSocket1.Active then
begin
  showmessage('Установите соединение с сервером!');
  exit;
end;
if AccessSettingLevel=true then
  begin
   RemoteDesktopSetting.memo1.Clear;
   UnloadFile.GetStrings(RemoteDesktopSetting.memo1.Lines);
   RemoteDesktopSetting.Tag:=0;
   RemoteDesktopSetting.ShowModal
  end
else
  showmessage('Не достаточно привелегий'+#13#10+' учетной записи');
end;

procedure TMRDForm.SpeedButton9Click(Sender: TObject);
begin
try
if not ConnectionEnable then
  begin
  levelCntrl:=false; /// уровень привелегий CTRL_LEVEL_VIEW  , если true то  CTRL_LEVEL_INTERACTIVE
  if LabelEdEdit1.Text='' then
   begin
    showmessage('Укажите имя компьютера');
    exit;
   end;
  createviewer; // создаем окно
  ClientSocket1.Port:=strtoint(Editport.text);   /// порт клиента
  ClientSocket1.Host:=LabelEdEdit1.Text;   //// имя сервера
  ClientSocket1.Active:=true;  /// активация клиента
  UnloadFileSetting:=false;
  ConnectionEnable:=true;
  end;
  Except
   on E: Exception do
      begin
        FrmDomainInfo.memo1.Lines.add('Ошибка в самом начале :'+E.Message);
      end;
   end;
end;

end.
