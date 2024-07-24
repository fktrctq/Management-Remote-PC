unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
   Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleServer, RDPCOMAPILib_TLB,
  Vcl.StdCtrls, Vcl.OleCtrls,ActiveX, System.Win.ScktComp,inifiles, Vcl.ExtCtrls,
  System.Win.TaskbarCore, Vcl.Taskbar;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    ServerSocket1: TServerSocket;
    Label1: TLabel;
    TrayIcon1: TTrayIcon;
    Timer1: TTimer;
    procedure RDPSession1AttendeeConnected(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPSession1AttendeeDisconnected(ASender: TObject;
      const pDisconnectInfo: IDispatch);
    procedure RDPSession1ApplicationOpen(ASender: TObject;
      const pApplication: IDispatch);
    procedure RDPSession1ApplicationUpdate(ASender: TObject;
      const pApplication: IDispatch);
    procedure RDPSession1AttendeeUpdate(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPSession1ChannelDataReceived(ASender: TObject;
      const pChannel: IInterface; lAttendeeId: Integer;
      const bstrData: WideString);
    procedure RDPSession1ChannelDataSent(ASender: TObject;
      const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
    procedure RDPSession1ConnectionAuthenticated(Sender: TObject);
    procedure RDPSession1ConnectionEstablished(Sender: TObject);
    procedure RDPSession1ConnectionFailed(Sender: TObject);
    procedure RDPSession1ConnectionTerminated(ASender: TObject; discReason,
      ExtendedInfo: Integer);
    procedure RDPSession1ControlLevelChangeRequest(ASender: TObject;
      const pAttendee: IDispatch; RequestedLevel: TOleEnum);
    procedure RDPSession1Error(ASender: TObject; ErrorInfo: OleVariant);
    procedure RDPSession1FocusReleased(ASender: TObject; iDirection: Integer);
    procedure RDPSession1GraphicsStreamPaused(Sender: TObject);
    procedure RDPSession1GraphicsStreamResumed(Sender: TObject);
    procedure RDPSession1SharedDesktopSettingsChanged(ASender: TObject; width,
      height, colordepth: Integer);
    procedure RDPSession1SharedRectChanged(ASender: TObject; left, top, right,
      bottom: Integer);
    procedure RDPSession1WindowClose(ASender: TObject;
      const pWindow: IDispatch);
    procedure RDPSession1WindowOpen(ASender: TObject; const pWindow: IDispatch);
    procedure RDPSession1WindowUpdate(ASender: TObject;
      const pWindow: IDispatch);
    function NewConnectString():WideString;  //// создаем новую строку
    function findLevel(s:string):integer;    //// поиск уровня доступа клиента
    function Log_write(fname, text:string):string;
    function MyExtendedInfo (i:integer):string;
    function MydiscReason (i:integer):string;
    procedure ServerSocket1Listen(Sender: TObject; Socket: TCustomWinSocket);
    procedure ServerSocket1Accept(Sender: TObject; Socket: TCustomWinSocket);
    procedure ServerSocket1ClientConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientError(Sender: TObject;
      Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
      var ErrorCode: Integer);
    procedure ServerSocket1ClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientWrite(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1GetSocket(Sender: TObject; Socket: NativeInt;
      var ClientSocket: TServerClientWinSocket);
    procedure ServerSocket1GetThread(Sender: TObject;
      ClientSocket: TServerClientWinSocket;
      var SocketThread: TServerClientThread);
    procedure ServerSocket1ThreadEnd(Sender: TObject;
      Thread: TServerClientThread);
    procedure ServerSocket1ThreadStart(Sender: TObject;
      Thread: TServerClientThread);
    procedure FormCreate(Sender: TObject);
    procedure Code(var text: WideString; password: WideString;  //// процедура кодирования и декодирования файла
decode: boolean);
    procedure createserv;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RDPSession1ApplicationClose(ASender: TObject;
      const pApplication: IDispatch);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    ////////////////////////
    procedure WMQueryEndSession(var Message : TMessage); message WM_QUERYENDSESSION;
    procedure PressAltCtrlDel;
     //// завершение программы при перезагрузке или выключениии Windows

  private
       { Private declarations }
  public
    { Public declarations }

  end;

var
  Form1: TForm1;

  MyStringConnect:WideString;
  MyRDPSession:TRDPSession;
  ClientPass,ClientLogin:TstringList;
  LevelUser:TstringList;
  ServerPort:integer;
  LoadFile:bool;    ////  загрузка файла от администратора
  UnloadFile:bool; //// выгрузка файла администратору
  Mini: TMemIniFile;  //// дешифрованный файл
  CurrentLevelUser:byte;
  ClientCount:integer; /// кол-во пользователей в ини файле
  UserConnect:bool; //// есть ли подключения
  ConnectUser:integer;  //// текущий подключившийся пользователь (кол-во подключившихся)
  MyThread:TThread;
  RDPSession:boolean;
  WindowsVer:string;
  closeserver:bool;
implementation
uses VersionCheck;
{$R *.dfm}




function Tform1.MyExtendedInfo (i:integer):string;
begin
case i of
1:result:=('Причина - Приложение инициировало отключение.');
2:result:=('Причина - Приложение инициировало выход клиента.');
3:result:=('Причина - Сервер отключил клиент, потому что клиент простаивал в течение периода времени, превышающего указанный период тайм-аута.');
4:result:=('Причина - Сервер отключил клиент, потому что клиент превысил период, указанный для подключения.');
5:result:=('Причина - Соединение клиента было заменено другим соединением.');
6:result:=('Причина - Нет памяти.');
7:result:=('Причина - Сервер отказал в подключении.');
8:result:=('Причина - Сервер отказал в соединении по соображениям безопасности.');
9:result:=('Причина - Соединение было отклонено, поскольку учетная запись пользователя не авторизована для удаленного входа в систему.');
11:result:=('Причина - Ошибка внутреннего лицензирования.');
12:result:=('Причина - Сервер лицензий недоступен.');
13:result:=('Причина - Лицензия на программное обеспечение не была доступна.');
14:result:=('Причина - Удаленный компьютер получил недопустимое лицензионное сообщение.');
15:result:=('Причина - Идентификатор оборудования не совпадает с идентификатором, указанным в лицензии на программное обеспечение.');
16:result:=('Причина - Ошибка лицензии клиента.');
17:result:=('Причина - Сетевые проблемы возникли во время протокола лицензирования.');
18:result:=('Причина - Клиент преждевременно завершил протокол лицензирования.');
20:result:=('Причина - Лицензионное сообщение было зашифровано неправильно.');
21:result:=('Причина - Лицензия на доступ к локальному компьютеру не может быть обновлена ​​или обновлена.');
22:result:=('Причина - Удаленный компьютер не имеет лицензии на прием удаленных подключений.');
23:result:=('Причина - Соединение было отклонено, поскольку продавец не смог аутентифицировать зрителя.');
24:result:=('Причина - внутренние ошибки протокола');
end;
end;
function Tform1.MydiscReason (i:integer):string;
begin
case i of
0: result:=('Соединение закрыто - Информайия отсутствует');
1: result:=('Соединение закрыто - Локальное отключение');
2: result:=('Соединение закрыто - Удаленное отключение пользователем');
3: result:=('Соединение закрыто - Удаленное отключение сервером');
260:result:=('Соединение закрыто - Ошибка поиска имени DNS');
262: result:=('Соединение закрыто - Недостаточно памяти.');
264: result:=('Соединение закрыто - Время соединения истекло.');
516: result:=('Соединение закрыто - Сбой подключения к Windows.');
518:result:=('Соединение закрыто - Недостаточно памяти.');
520: result:=('Соединение закрыто - Ошибка хост не найден.');
772: result:=('Соединение закрыто - Windows Sockets ошибка вызова.');
774: result:=('Соединение закрыто - Недостаточно памяти.');
776: result:=('Соединение закрыто - Указан неверный IP-адрес');
1028:result:=('Соединение закрыто - Ошибка вызова Windows Sockets');
1030:result:=('Соединение закрыто - Недействительные данные безопасности.');
1032:result:=('Соединение закрыто - Внутренняя ошибка.');
1286:result:=('Соединение закрыто - Указан недопустимый метод шифрования.');
1288:result:=('Соединение закрыто - Ошибка поиска DNS.');
1540:result:=('Соединение закрыто - Ошибка сокета Windows Sockets gethostbyname.');
1542:result:=('Соединение закрыто - Недействительные данные безопасности сервера..');
1544:result:=('Соединение закрыто - Ошибка внутреннего таймера.');
1796:result:=('Соединение закрыто - Произошел тайм-аут.');
1798:result:=('Соединение закрыто - Не удалось распечатать сертификат сервера.');
2052:result:=('Соединение закрыто - Указан неверный IP-адрес.');
2056:result:=('Соединение закрыто - Согласование лицензии не удалось.');
2310:result:=('Соединение закрыто - Внутренняя ошибка безопасности.');
2308:result:=('Соединение закрыто - Server Winsock закрыт.');
2312:result:=('Соединение закрыто - Тайм-аут лицензирования.');
2566:result:=('Соединение закрыто - Внутренняя ошибка безопасности.');
2822:result:=('Соединение закрыто - Ошибка шифрования.');
3078:result:=('Соединение закрыто - Ошибка расшифровки.');
3080:result:=('Соединение закрыто - Ошибка декомпрессии.');
end;
end;

function TForm1.Log_write(fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
  //if not DirectoryExists('logs') then CreateDir('logs');
  if TrayIcon1.Visible then
   begin
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+fname+'.log');
        f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+fname+'.log');
      finally
        f.Destroy;
      end;
   end;
end;

function TForm1.NewConnectString():WideString;
var    //// создание новой строки для подключения
StAuthString,StGroupName,StPassword,stringConnect:string;
MyInvation:IRDPSRAPIInvitation;

begin
try
if RDPSession=false  then  createserv;
StAuthString:='RemoteComputer';
StGroupName:='RemoteAssistent';
StPassword:='ADMIN';
 //RDPsession1.Properties.Property_['PortId']:=48998;     /// устанавливаем номер порта для соединения
 //RDPsession1.Properties.Property_['PortProtocol']:=2;  /// устанавливаем тип протокола 0- ipv4,ipv6  2 - ipv4  23 - ipv6
 if copy(WindowsVer,1,2)='10' then /// если windows 10
 begin
  MyRDPSession.Properties.Property_['EnableClipboardRedirect']:=TRUE;  /// включаем буфер
 end;
MyRDPSession.Open;
MyInvation:=MyRDPSession.Invitations.CreateInvitation((WideString(StAuthString)),
(WideString(StGroupName)),(WideString(StPassword)),10);
result:= WideString(MyInvation.ConnectionString);
    Except
     on E:Exception do
      begin
        Log_Write('debug','Ошибка создания строки подключения - "' +E.Message+'"');
      end;
    End;
end;

function Crypt(varStr: WideString):WideString; /// функция шифрования и дешифровки текста для передачи
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

 procedure TForm1.WMQueryEndSession(var Message : TMessage);
 begin   //// завершение программы при перезагрузке или выключениии Windows
  //inherited;
 // form1.Close;
 //  Message.Result := Integer(CloseQuery and CallTerminateProcs);
  closeserver:=true; /// признак перезагрузки или выключения компьютера
end;


procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin

//ServerSocket1.Active := False; //// останавливаем сервер при закрытии
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if closeserver=false then  ////  если пришел сигнал перезагрузки или выключения компьютера то  closeserver=true
begin                      //// после этого программа закрывается и компьютер перезагружается
CanClose:=false;           /// иначе closeserver=false и при попытке закрыть прогу, программа свертывается в трей
Application.Minimize;
form1.Visible:=false;
end;

end;

procedure TForm1.PressAltCtrlDel;
begin
keybd_event(VK_CONTROL, 0, 0, 0); //Нажатие Ctrl.
keybd_event(VK_LSHIFT, 0, 0, 0); //Нажатие Shift.
keybd_event(VK_ESCAPE, 0, 0, 0); //Нажатие Esc.
keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0); //Отпускание Ctrl.
keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0); //Отпускание Shift.
keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0); //Отпускание Esc.
Log_Write('debug','Открытие диспетчера задач');
end;

procedure TForm1.Code(var text: Widestring; password: WideString;   //// шифровка дешифрация текста для файлов
decode: boolean);
var
i, PasswordLength: integer;
sign: shortint;
begin
PasswordLength := length(password);
if PasswordLength = 0 then Exit;
if decode
then sign := -1
else sign := 1;
for i := 1 to Length(text) do
text[i] := chr(ord(text[i]) + sign *
ord(password[i mod PasswordLength + 1]));
end;
////////////////////////////////////////////////////////////////////////////



procedure TForm1.ServerSocket1ClientRead(Sender: TObject; //// socket получаем от клиента сообщение
  Socket: TCustomWinSocket);                              //// отправляем клиенту ключ в сообщении
  var
  i,z,l:integer;
  ClientMessage,answer,UnloadString:Widestring;
  ListFile,ListFile2:TstringList;
  UnloadList:TStringList;
  s,SettingFile:Widestring;
begin
  try
   if LoadFile=true then     //// замена файла настроек
        begin
        try
          begin
          ClientMessage:= (Socket.ReceiveText);
          ListFile:=TstringList.Create;
          ListFile2:=TstringList.Create;
          ClientMessage:=crypt(ClientMessage);
          ListFile.SetText(Pchar(ClientMessage));
          Mini.SetStrings(ListFile);/// Обновляем файл настроек в памяти
          ///  обновляем списки доступа
           ClientLogin.Clear;
           ClientPass.Clear;
           LevelUser.Clear;
           for I := 0 to Mini.ReadInteger('RemoteClient','Count',0) do
             begin
             ClientLogin.Add(Mini.ReadString('RemoteClient','ClientName'+inttostr(i),'Administrator'));
             ClientPass.Add(Mini.ReadString('RemoteClient','ClientPasswd'+inttostr(i),'159753'));
             LevelUser.Add(Mini.ReadString('RemoteClient','Access'+inttostr(i),'3'));
             end;
           ClientCount:=Mini.ReadInteger('RemoteClient','Count',0);
           TrayIcon1.Visible:=Mini.ReadBool('Server','View',true);
           if serverport<>Mini.ReadInteger('Server','Port',48999) then
            begin
            serverport:=Mini.ReadInteger('Server','Port',48999);
            ServerSocket1.Active:=false;
            ServerSocket1.Port := ServerPort;   /// порт сервера
            ServerSocket1.Active:=true;
            end;
          ////////////////////////////////////////////////////////////////////////////
           for i := 0 to ListFile.Count-1 do
           begin
           s:=ListFile[i];
           code(s, '12345', false);
           ListFile2.Add(s);
           end;
          ListFile2.SaveToFile(ExtractFilePath(Application.ExeName)+'client.dat',TEncoding.Unicode);
          Log_Write('debug','Произошло изменение настроек программы');
          ListFile.Free;
          ListFile2.Free;
          LoadFile:=false;
          exit;
          end;
            Except
             on E:Exception do
              begin
                Log_Write('debug','Ошибка сохранения настроек  - "' +E.Message+'"');
              end;
            End;
        end;

  {От клиента получено сообщение - выводим его в Memo1}
  try
 // Memo1.Lines.add( 'Сообщение от клиента');
  ClientMessage:= (Socket.ReceiveText);
 // Memo1.Lines.add('CODE in client - '+ClientMessage );
  ClientMessage:=Crypt(ClientMessage); /// дешифрация входящего сообщения
 // Memo1.Lines.add('DECODE in client - r - '+ClientMessage );
  Except
   on E:Exception do
    begin
    Log_Write('debug','Ошибка получения сообщения от клиента - "' +E.Message+'"');
    end;
  End;

  if clientMessage='<Shift+Ctrl+Esc>' then
begin
    PressAltCtrlDel;
    exit;
    end;

  for I := 0 to ClientLogin.Count-1 do
    begin
     if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])+'<settingFile>') and (Leveluser[i]='0') then
       begin
       LoadFile:=true; //// включаем загрузку файлов
       memo1.Lines.Add('Применение новых настроек');
       exit;
       end;
     end;


  for I := 0 to ClientLogin.Count-1 do  /////// если логин и пароль верны
  begin
      try
       begin
        if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])) then
          Begin
           if UserConnect=false then  //// если нет подключенных пользователей то создаем новую строку
              begin
              Memo1.Lines.add( 'Настройка параметров нового сеанса');
              Log_write('debug','Настройка параметров нового сеанса');
              MyStringConnect:=(NewConnectString);
              MyStringConnect:=Crypt(MyStringConnect);  /// шифрация
              end;
         // for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do    //// отправляем
             ServerSocket1.Socket.Connections[Connectuser-1].SendText(MyStringConnect);
             currentLevelUser:=strtoint(LevelUser[i]);
             if currentLevelUser=0 then
                begin
                  sleep(1000);
                  SettingFile:=Crypt('<unloadlist>');
                        // for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
                   ServerSocket1.Socket.Connections[Connectuser-1].SendText(SettingFile);
                  sleep(1000);
                  unloadList:=TstringList.Create;
                  Mini.GetStrings(unloadList);
                  for z := 0 to unloadList.Count-1 do
                    begin
                      UnloadString:=UnloadString+(WideString(unloadList[z]+#13#10));
                    end;
                   UnloadString:=crypt(UnloadString);
                   //memo1.Lines.Add('отправляю файл настроек');
                   //for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
                   ServerSocket1.Socket.Connections[Connectuser-1].SendText(UnloadString);
                  unloadList.Free;
                end;
           break;
          End
          else ///клиент указал не верный пароль
          begin
          answer:=WideString('<ErrorUser>');
          answer:=crypt(answer);  /// шифрация
          currentLevelUser:=6;
          ServerSocket1.Socket.Connections[Connectuser-1].SendText(answer);
          Memo1.Lines.add( 'Клиент указал неверный логин или пароль');
          Log_Write('debug','Клиент указал неверный логин или пароль');
          end;

        end;
        Except
        on E:Exception do
          begin
          Log_Write('debug','Ошибка настройки параметров нового сеанса  - "' +E.Message+'"');
          end;
        End;
    end;
Except
on E:Exception do
  begin
  Log_Write('debug','Общая ошибка получения данных от клиента - "' +E.Message+'"');
  end;
End;

end;


procedure Tform1.createserv;  ///// запуск/перезапуск севрера
var
SettingIni:Tinifile;
mylist,mylist2:TstringList;
s:WideString;
a,b,i:integer;
begin
try
 if RDPSession then //// проверка , создана ли переменная MyRDPSession
   begin
     Memo1.Clear;
     if ClientCount<>0 then MyRDPSession.close; //// если пользователи подключены ,то отключаем
     MyRDPSession.Destroy;
     RDPSession:=false;
     if ServerSocket1.Active then ServerSocket1.Active:=false;
     ClientLogin.Free;
     ClientPass.Free;
     LevelUser.Free;
     memo1.Lines.Add('Перезапуск сервера');
     Log_write('debug','Перезапуск сервера');
   end;

 ClientLogin:=TstringList.Create;
 ClientPass:=TstringList.Create;
 LevelUser:=TstringList.Create;
 LoadFile:=false;      //// отключем загрузку файлов
 CurrentLevelUser:=6; //// уровень доступа пользователя, устанавливаем никакой
 UserConnect:=false; ///// нет подключенных пользователей
 if not FileExists(ExtractFilePath(Application.ExeName) + 'client.dat') then
 begin
 SettingIni:=Tinifile.Create(ExtractFilePath(Application.ExeName) + 'client.dat');
 SettingIni.WriteString('RemoteClient','ClientName0','Administrator');
 SettingIni.WriteString('RemoteClient','ClientPasswd0','');
 SettingIni.WriteInteger('RemoteClient','Access0',0);
 SettingIni.Writeinteger('RemoteClient','Count',0);
 SettingIni.Writeinteger('Server','Port',48999);
 SettingIni.WriteBool('Server','View',true);
 SettingIni.Free;
 ///////////////////////// начало шифрации файла
 mylist:=Tstringlist.Create;
 mylist2:=Tstringlist.Create;
 mylist.Clear;
 mylist2.Clear;
 mylist.LoadFromFile(ExtractFilePath(Application.ExeName) + 'client.dat');
 for b := 0 to mylist.Count-1 do
   begin
   s:=mylist[b];
   code(s, '12345', false);
   mylist2.Add(s);
   end;
 mylist2.SaveToFile(ExtractFilePath(Application.ExeName) + 'client.dat',TEncoding.Unicode);
 mylist.Free;
 mylist2.Free;
 Log_write('debug','Создаю файл настроек');
 ////////////////////// окончание шифрации файла
 end;

 ///////////////////////////   начало дешифрации файлов
    mylist:=Tstringlist.Create;
    mylist2:=Tstringlist.Create;
    Mini:=TMemIniFile.Create(ChangeFileExt( Application.Exename,'.ini'));
    begin
    mylist.LoadFromFile(ExtractFilePath(Application.ExeName) + 'client.dat',TEncoding.Unicode);
    for b := 0 to mylist.Count-1 do
       begin
       s:=mylist[b];
       code(s, '12345', true);
       mylist2.Add(s);
       end;
    Mini.SetStrings(mylist2);
    end;
    mylist.Free;
  ///////////////////////////  окончание дешифрации файлов
 for I := 0 to Mini.ReadInteger('RemoteClient','Count',0) do
   begin
   ClientLogin.Add(Mini.ReadString('RemoteClient','ClientName'+inttostr(i),'Administrator'));
   ClientPass.Add(Mini.ReadString('RemoteClient','ClientPasswd'+inttostr(i),'159753'));
   LevelUser.Add(Mini.ReadString('RemoteClient','Access'+inttostr(i),'3'));
   end;
  serverport:=Mini.ReadInteger('Server','Port',48999);
  ClientCount:=Mini.ReadInteger('RemoteClient','Count',0);
  if TrayIcon1.Visible<>Mini.ReadBool('Server','View',true) then TrayIcon1.Visible:=Mini.ReadBool('Server','View',true);

 MyRDPSession:= TRDPSession.Create(self);
 MyRDPSession.OnAttendeeConnected:=RDPSession1AttendeeConnected;
 MyRDPSession.OnApplicationOpen:= RDPSession1ApplicationOpen;
 MyRDPSession.OnApplicationClose:= RDPSession1ApplicationClose;
 MyRDPSession.OnAttendeeUpdate:= RDPSession1AttendeeUpdate;
 MyRDPSession.OnAttendeeDisconnected:= RDPSession1AttendeeDisconnected;
 MyRDPSession.OnApplicationUpdate:=  RDPSession1ApplicationUpdate;
 MyRDPSession.OnConnectionEstablished:= RDPSession1ConnectionEstablished;
 MyRDPSession.OnConnectionFailed:= RDPSession1ConnectionFailed;
 MyRDPSession.OnConnectionTerminated:= RDPSession1ConnectionTerminated;
 MyRDPSession.OnConnectionAuthenticated:= RDPSession1ConnectionAuthenticated;
 MyRDPSession.OnError:=RDPSession1Error;
 MyRDPSession.OnWindowOpen:=RDPSession1WindowOpen;
 MyRDPSession.OnWindowClose:=RDPSession1WindowClose;
 MyRDPSession.OnWindowUpdate:=RDPSession1WindowUpdate;
 MyRDPSession.OnControlLevelChangeRequest:=RDPSession1ControlLevelChangeRequest;
 MyRDPSession.OnGraphicsStreamPaused:=RDPSession1GraphicsStreamPaused;
 MyRDPSession.OnGraphicsStreamResumed:=RDPSession1GraphicsStreamResumed;
 MyRDPSession.OnChannelDataReceived:=RDPSession1ChannelDataReceived;
 MyRDPSession.OnChannelDataSent:=RDPSession1ChannelDataSent;
 MyRDPSession.OnSharedRectChanged:=RDPSession1SharedRectChanged;
 MyRDPSession.OnFocusReleased:=RDPSession1FocusReleased;
 MyRDPSession.OnSharedDesktopSettingsChanged:=RDPSession1SharedDesktopSettingsChanged;
// MyRDPSession.ApplicationFilter.Enabled:=true;    /// использовать если необходимо расшарить только приложения
 try
  Log_Write('debug','Colordepth из файла :'+(Mini.ReadString('Server','colordepth',''))+' bit');
 MyRDPSession.colordepth:=Mini.ReadInteger('Server','colordepth',24);
 Log_Write('debug','Colordepth система :'+inttostr(MyRDPSession.colordepth)+' bit');
 //MRDuRDPViewer.ControlInterface.RequestColorDepthChange(16);
 except
 on E: Exception do  Log_Write('debug','Colordepth :'+E.Message);
 end;

 try
  //MRDuRDPViewer.Properties.Property_['FrameCaptureIntervalInMs']:=1000 div EditRDPFPS.Value;  //https://docs.microsoft.com/ru-ru/windows/win32/api/rdpencomapi/nf-rdpencomapi-irdpsrapisessionproperties-get_property
  Log_Write('debug','FPS  из файла (кадр/сек) :'+(Mini.Readstring('Server','FPS','')));
  MyRDPSession.Properties.set_Property_('FrameCaptureIntervalInMs',variant(1000 div (Mini.ReadInteger('Server','FPS',30))));
  Log_Write('debug','FPS система (мл.сек) :'+vartostr(MyRDPSession.Properties.get_Property_('FrameCaptureIntervalInMs')));
 except
 on E: Exception do  Log_Write('debug','FPS :'+E.Message);
 end;

 RDPSession:=true;
 ServerSocket1.Port := ServerPort;   /// порт сервера
 ServerSocket1.Active:=true;
 mylist2.Free;
 Log_write('debug','Запуск сервера');
 Memo1.Lines.Add('Старт сервера');
   Except
  on E:Exception do
    begin
     Log_Write('debug','Произошла ошибка при запуске/перезапуске сервера - "' +E.Message+'"');
    end;
  End;
end;




procedure TForm1.FormCreate(Sender: TObject);  /// загрузка программы
begin
Application.Title:='uRDMServer';
SetWindowLong(Application.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
timer1.Enabled:=true;
WindowsVer:=wmiver;
memo1.Lines.Add(WindowsVer);
createserv; /// создание
closeserver:=false;  /// признак перезагрузки или выключения компьютера
end;


procedure TForm1.FormShow(Sender: TObject);
begin
SetWindowLong(Application.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
form1.Hide;           //// скрывваем форму и выключаем таймер
Timer1.Enabled:=false;
end;

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
//Application.Restore;
form1.Show;
end;

procedure TForm1.RDPSession1ApplicationClose(ASender: TObject;
  const pApplication: IDispatch);
begin
//memo1.Lines.Add('Приложение закрывается');
end;

procedure TForm1.RDPSession1ApplicationOpen(ASender: TObject;
  const pApplication: IDispatch);
  var
  newApplication:IRDPSRAPIApplication;
  MyWindowsList:IRDPSRAPIWindowList;
  MyWindows: IRDPSRAPIWindow;
begin
{memo1.Lines.Add('Создается новое приложение');
pApplication.QueryInterface(IID_IRDPSRAPIApplication,newApplication);
memo1.Lines.Add('Приложение - '+string(newApplication.Name));
newApplication.Windows.QueryInterface(IID_IRDPSRAPIWindowList,MyWindowsList);
memo1.Lines.Add('MyWindowsList');
newApplication.Shared:=true;
memo1.Lines.Add('newApplication.Shared'); }
//MyWindowsList.Item[newApplication.Id].Shared:=true;
//memo1.Lines.Add('MyWindowsList.Item[newApplication.Id].Shared ');
//MyWindowsList.Item[newApplication.Id].Show;
//memo1.Lines.Add('MyWindowsList.Item[newApplication.Id].Show');
//memo1.Lines.Add('MyWindowsList.Item[newApplication.Id].Name - '
//+MyWindowsList.Item[newApplication.Id].Name+'/ ID - '+inttostr(MyWindowsList.Item[newApplication.Id].Id));
end;

procedure TForm1.RDPSession1ApplicationUpdate(ASender: TObject;
  const pApplication: IDispatch);
begin
 //memo1.Lines.Add('изменение общего свойства объекта приложения');
end;


procedure TForm1.RDPSession1AttendeeDisconnected(ASender: TObject;
  const pDisconnectInfo: IDispatch);   /// отключение клиента
  var
  i:integer;
  myclient:IRDPSRAPIAttendeeDisconnectInfo;
begin
try
pDisconnectInfo.QueryInterface(IID_IRDPSRAPIAttendeeDisconnectInfo,Myclient);
memo1.Lines.Add(string(Myclient.Attendee.RemoteName)+' отключается от сеанса');
Log_Write('debug',string(Myclient.Attendee.RemoteName)+' отключается от сеанса');
   Except
  on E:Exception do
    begin
     Log_Write('debug','Произошла ошибка на уровне закрытия сеанса - "' +E.Message+'"');
    end;
  End;

end;

procedure TForm1.RDPSession1AttendeeConnected(ASender: TObject; //// подключение клиента
  const pAttendee: IDispatch);
  var
 m:integer;
 Myclient:IRDPSRAPIAttendee;
begin
try
pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// получаем интерфейс клиента  в  MyClient
///SendMessage(HWND_BROADCAST, WM_SYSCOMMAND,SC_MONITORPOWER, -1); //// включить монитор
begin
     memo1.Lines.Add(string(MyClient.RemoteName)+' - подключается к сеансу');
     Log_Write('debug',string(MyClient.RemoteName)+' подключается к сеансу');
     if findlevel(string(Myclient.RemoteName))=2 then
      begin
       m:=MessageDlg('Разрешить '+string(MyClient.RemoteName)+#13#10+ 'подключится к Вашему компьютеру!!!'
       ,mtConfirmation,[mbYes,mbCancel], 0);
        if m=mrYes then
           begin
           Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
           Log_Write('debug',string(Myclient.RemoteName)+' подключен к сеансу с правами на просмотр');
           //memo1.Lines.Add(string(Myclient.RemoteName)+' подключен к сеансу с правами на просмотр');
           UserConnect:=true;
           end
          else
            begin
             Myclient.ControlLevel:=CTRL_LEVEL_NONE;
             Log_Write('debug',string(Myclient.RemoteName)+' пользователь отклонил подключение к сеансу');
             //memo1.Lines.Add(string(Myclient.RemoteName)+' пользователь отклонил подключение к сеансу');
             Myclient.TerminateConnection;
             if ConnectUser=0 then UserConnect:=false;
            end;
       exit;
      end;

      case currentLevelUser of
        0:
        begin
        //memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_VIEW');
        Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
        Log_Write('debug',string(MyClient.RemoteName)+' - подключился к сеансу');
        memo1.Lines.Add(string(MyClient.RemoteName)+' - подключился к сеансу');
        UserConnect:=true;
        end;
        1:
        begin
       // memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_VIEW');
        Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
         Log_Write('debug',string(MyClient.RemoteName)+' - подключился к сеансу');
        memo1.Lines.Add(string(MyClient.RemoteName)+' - подключился к сеансу');
        UserConnect:=true;
        //MyChanal:=MyRDPSession.VirtualChannelManager.CreateVirtualChannel('ChAdm#',CHANNEL_PRIORITY_HI,0); // создаем виртуальный канал  //  CHANNEL_PRIORITY_LO = $00000000;  CHANNEL_PRIORITY_MED = $00000001;   CHANNEL_PRIORITY_HI = $00000002;
        //MyChanal.SetAccess(MyClient.Id,CHANNEL_ACCESS_ENUM_SENDRECEIVE); //CHANNEL_ACCESS_ENUM_NONE -  Значение: 0 Нет доступа. Участник не может отправлять или получать данные на канале. CHANNEL_ACCESS_ENUM_SENDRECEIVE -  Участник может отправлять или получать данные на канале.
        //MyRDPSession.Properties.set_Property_('FrameCaptureIntervalInMs',variant(1000 div (Mini.ReadInteger('Server','FPS',30))));
        end;
      end;

  end;
   Except
  on E:Exception do
    begin
     Log_Write('debug','Произошла ошибка на уровне предоставления привелегий участнику сеанса - "' +E.Message+'"');
    end;
  End;

end;


procedure TForm1.ServerSocket1Accept(Sender: TObject; Socket: TCustomWinSocket);
begin
  try
 {Здесь сервер принимает клиента ServerSocket1.Socket.ActiveConnections}
 // Memo1.Lines.add('Client connection accepted count user-'+inttostr(ServerSocket1.Socket.ActiveConnections-1));
  Memo1.Lines.add('Компьютер - '+Socket.RemoteHost+' / IP-адрес - '+Socket.RemoteAddress
  +' / Порт - '+inttostr(socket.RemotePort)+'/Участник №'+inttostr(Connectuser));
  Log_Write('debug','подключение от компьютера - '+Socket.RemoteHost+' / IP-адрес - '+Socket.RemoteAddress
  +' / Порт - '+inttostr(socket.RemotePort)+'/Участник №'+inttostr(Connectuser));
   Except
  on E:Exception do
    begin
     Log_Write('debug','Ошибка получения данных о клиенте - "' +E.Message+'"');
    end;
  End;
 end;

function TForm1.findLevel(s:string):integer; //// поиск уровня привелегий пользователя
var
i:integer;
begin
for I := 0 to ClientLogin.Count-1 do
  begin
    if s=ClientLogin[i] then
       begin
       result:=strtoint(LevelUser[i]);
       exit;
       end
    else
       result:=2;
  end;
end;

procedure TForm1.RDPSession1ControlLevelChangeRequest(ASender: TObject;
  const pAttendee: IDispatch; RequestedLevel: TOleEnum);
  var
                /// Смена уровня привелегий в зависимости от CurrentLevelUser
 Myclient:IRDPSRAPIAttendee;
  begin
  try
    pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// получаем интерфейс клиента
    begin
    if findlevel(string(Myclient.RemoteName))=2 then
     begin
     memo1.Lines.Add(string(Myclient.RemoteName)+' - изменил уровень привелегий');
     Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
     Log_Write('debug',string(Myclient.RemoteName)+' - изменил уровень привелегий');
     exit;
     end;

         try
           begin      //// установка привелегий для остальных пользователей
            memo1.Lines.Add(string(Myclient.RemoteName)+' - изменил уровень привелегий');
            Myclient.ControlLevel:=RequestedLevel;
            Log_Write('debug',string(Myclient.RemoteName)+' - изменил уровень привелегий ');
            //RDPsession1.Properties.Property_['PortId']:=48998;     /// устанавливаем номер порта для соединения
            //RDPsession1.Properties.Property_['PortProtocol']:=2;  /// устанавливаем тип протокола 0- ipv4,ipv6  2 - ipv4  23 - ipv6
            // MyRDPSession.Properties.Property_['EnableClipboardRedirect']:=TRUE;
           end
           //// изменение уровня привелегий подключаемого
          Except  /// если ошибка применения уровня привелегий до CTRL_LEVEL_MAX
                  /// то применяем другой уровнь CTRL_LEVEL_INTERACTIVE
           Myclient.ControlLevel:=CTRL_LEVEL_INTERACTIVE;
           Log_Write('debug',string(Myclient.RemoteName)+' - изменил уровень привелегий на INTERACTIVE');
         end;
   end;
  Except
  on E:Exception do
    begin
     Log_Write('debug','Ошибка смены уровня привелегий по запросу пользователя - "' +E.Message+'"');
    end;
  End;

end;



procedure TForm1.RDPSession1AttendeeUpdate(ASender: TObject;
  const pAttendee: IDispatch);
begin
 Log_Write('debug','изменяется одно из значений свойств для участника');
//  memo1.Lines.Add('изменяется одно из значений свойств для участника');
end;

procedure TForm1.RDPSession1ChannelDataReceived(ASender: TObject;
  const pChannel: IInterface; lAttendeeId: Integer; const bstrData: WideString);
begin
 Log_Write('debug','данные принимаются от участника сеанса.');
//  memo1.Lines.Add(' данные принимаются от участника сеанса.');
end;

procedure TForm1.RDPSession1ChannelDataSent(ASender: TObject;
  const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
begin
 Log_Write('debug','данные отправляются участнику сеанса.');
// memo1.Lines.Add(' данные отправляются участнику сеанса.');
end;

procedure TForm1.RDPSession1ConnectionAuthenticated(Sender: TObject);
begin
  Log_Write('debug','Соединение аутентифицированно.');
 // memo1.Lines.Add(' Соединение аутентифицированно.');
end;

procedure TForm1.RDPSession1ConnectionEstablished(Sender: TObject);
begin
   Log_Write('debug','установлено соединение с сервером.');
 // memo1.Lines.Add(' установлено соединение с сервером.');
end;

procedure TForm1.RDPSession1ConnectionFailed(Sender: TObject);
begin
  Log_Write('debug','клиент не может подключиться к серверу');
//  memo1.Lines.Add('клиент не может подключиться к серверу');
end;

procedure TForm1.RDPSession1ConnectionTerminated(ASender: TObject; discReason,
  ExtendedInfo: Integer);
begin
 try
 Log_Write('debug','соединение клиента с сервером закрывается:  ' + MyExtendedInfo(ExtendedInfo));
 Log_Write('debug','Причина закрытия сеанса:  ' + MydiscReason(discReason));
// memo1.Lines.Add(' соединение клиента с сервером закрывается - ' + MyExtendedInfo(ExtendedInfo)+'/'+MydiscReason(discReason));
  Except
  on E:Exception do
    begin
     Log_Write('debug','Ошибка закрытия сеанса - "' +E.Message+'"');
    end;
  End;
end;


procedure TForm1.RDPSession1Error(ASender: TObject; ErrorInfo: OleVariant);
begin
try
 Log_Write('debug','В сеансе возникла критическая ошибка:  ' + SysErrorMessage(ErrorInfo));
// memo1.Lines.Add(' В сеансе возникла критическая ошибка - ' + SysErrorMessage(ErrorInfo));
 //createserv;
 Except
  on E:Exception do
    begin
     Log_Write('debug','В сеансе возникла критическая ошибка - "' +E.Message+'"');
    end;
  End;
 end;

procedure TForm1.RDPSession1FocusReleased(ASender: TObject;
  iDirection: Integer);
begin
 Log_Write('debug','общее окно верхнего уровня набрало или потеряло фокус');
 // memo1.Lines.Add(' общее окно верхнего уровня набрало или потеряло фокус');
end;

procedure TForm1.RDPSession1GraphicsStreamPaused(Sender: TObject);
begin
  Log_Write('debug','графический поток приостановлен');
// memo1.Lines.Add('графический поток приостановлен');
end;

procedure TForm1.RDPSession1GraphicsStreamResumed(Sender: TObject);
begin
   Log_Write('debug','графический поток восстановлен');
 //  memo1.Lines.Add('графический поток восстановлен');
end;

procedure TForm1.RDPSession1SharedDesktopSettingsChanged(ASender: TObject;
  width, height, colordepth: Integer);
begin
   Log_Write('debug','изменении общей настройки рабочего стола');
//  memo1.Lines.Add('изменении общей настройки рабочего стола');
end;

procedure TForm1.RDPSession1SharedRectChanged(ASender: TObject; left, top,
  right, bottom: Integer);
begin
  Log_Write('debug','изменяется размер общего окна верхнего уровня приложения');
 // memo1.Lines.Add('изменяется размер общего окна верхнего уровня приложения');
end;

procedure TForm1.RDPSession1WindowClose(ASender: TObject;
  const pWindow: IDispatch);
begin
  Log_Write('debug','закрывается закрываемое окно верхнего уровня');
 // memo1.Lines.Add('закрывается закрываемое окно верхнего уровня');
end;

procedure TForm1.RDPSession1WindowOpen(ASender: TObject;
  const pWindow: IDispatch);
begin
  Log_Write('debug','приложения создает открытое окно верхнего уровня');
 // memo1.Lines.Add('приложения создает открытое окно верхнего уровня');
end;

procedure TForm1.RDPSession1WindowUpdate(ASender: TObject;
  const pWindow: IDispatch);
begin
  Log_Write('debug','Произошло изменение одного из свойств объекта окна');
 // memo1.Lines.Add('Произошло изменение одного из свойств объекта окна');
end;



procedure TForm1.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  {Здесь клиент подсоединяется}
Memo1.Lines.add('Клиент подключается');
Log_Write('debug','Подключение клиента');
Connectuser:=ServerSocket1.Socket.ActiveConnections;
Label1.Caption:= 'Подключенных клиентов - '+Inttostr(ServerSocket1.Socket.ActiveConnections);
end;

procedure TForm1.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
 try
   {Здесь клиент отсоединяется}
  Log_Write('debug','Отключение клиента');
  ConnectUser:=ServerSocket1.Socket.ActiveConnections-1;
  Label1.Caption:= 'Подключенных клиентов - '+Inttostr(Connectuser);
    if (ConnectUser=0)  then
      begin
      UserConnect:=false;   //// нет подключенных пользователей
      if Assigned(MyRDPSession) then  /// если компонент создан то удаляем
        begin
        MyRDPSession.Close;        /// отключение сессии
        MyRDPSession.Destroy;
        end;
      RDPSession:=false;
      Log_Write('debug','Полное закрытие сеанса');
      end;
  Except
  on E:Exception do
    begin
     Log_Write('debug','Клиент отключается, ошибка - "' +E.Message+'"');
    end;
  End;
end;

procedure TForm1.ServerSocket1ClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent; var ErrorCode: Integer);
begin
 {Произошла ошибка - выводим ее код}
  Log_Write('debug','Ошибка подключения:  ' + SysErrorMessage(ErrorCode));
 // Memo1.Lines.add('Ошибка подключения:  ' + SysErrorMessage(ErrorCode));
  //createserv;
end;



procedure TForm1.ServerSocket1ClientWrite(Sender: TObject;
  Socket: TCustomWinSocket);
  var
  i:integer;
begin
  {Теперь можно слать данные в сокет}
  //Log_Write('debug','Теперь можно слать данные в сокет');
 // Memo1.Lines.add('Теперь можно слать данные в сокет');

end;

procedure TForm1.ServerSocket1GetSocket(Sender: TObject; Socket: NativeInt;
  var ClientSocket: TServerClientWinSocket);
begin
//Memo1.Lines.add('Get socket');
end;

procedure TForm1.ServerSocket1GetThread(Sender: TObject;
  ClientSocket: TServerClientWinSocket; var SocketThread: TServerClientThread);
begin
 //Memo1.Lines.add('GetThread');
end;

procedure TForm1.ServerSocket1Listen(Sender: TObject; Socket: TCustomWinSocket);
begin
try
  {Здесь сервер "прослушивает" сокет на наличие клиентов}
  Log_Write('debug','Прослушиваю порт '+IntToStr(ServerSocket1.Port)+' на наличие клиентов');
 // Memo1.Lines.Add('Прослушиваю порт на наличие клиентов ' + IntToStr(ServerSocket1.Port));

  Except
  on E:Exception do
    begin
     Log_Write('debug','Прослушиваю порт"' +E.Message+'"');
    // Memo1.Lines.Add('Прослушиваю порт на наличие клиентов "' +E.Message+'"');
    end;
  End;

  end;


procedure TForm1.ServerSocket1ThreadEnd(Sender: TObject;
  Thread: TServerClientThread);
begin
// Memo1.Lines.add('ThreadEnd');
end;

procedure TForm1.ServerSocket1ThreadStart(Sender: TObject;
  Thread: TServerClientThread);
begin
// Memo1.Lines.add('ThreadStart');
end;



end.
