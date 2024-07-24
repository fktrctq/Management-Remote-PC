unit MRDThread;

interface

uses
  System.Classes,ActiveX,
   System.Win.ScktComp,inifiles,RDPCOMAPILib_TLB,Vcl.Forms
   ,System.SysUtils,Vcl.Dialogs,VCL.Controls;

type
  MRD = class(TThread)
private


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
    procedure Code(var text: WideString; password: WideString;  //// процедура кодирования и декодирования файла
decode: boolean);
    procedure Start;
    { Private declarations }
  protected
    procedure Execute; override;
  public

  end;


implementation
uses Unit1;
var
   MyStringConnect:WideString;
  ClientPass,ClientLogin:TstringList;
  LevelUser:TstringList;
  ServerPort:integer;
  LoadFile:boolean;    ////  загрузка файла от администратора
  UnloadFile:boolean; //// выгрузка файла администратору
  Mini: TMemIniFile;  //// дешифрованный файл
  CurrentLevelUser:byte;
  ClientCount:integer; /// кол-во пользователей в ини файле
  UserConnect:boolean; //// есть ли подключения
  ConnectUser:integer;  //// текущий подключившийся пользователь (кол-во подключившихся)
  MyInvation:IRDPSRAPIInvitation;
  MyRDPSession:TRDPSession;
  MyServerSocket: TServerSocket;



{ MRD }
////////////////////////////////////////////////////////////
function MRD.NewConnectString():WideString;
var
StAuthString,StGroupName,StPassword,stringConnect:string;
begin

StAuthString:='RemoteComputer';
StGroupName:='RemoteAssistent';
StPassword:='ADMIN';
 //RDPsession1.Properties.Property_['PortId']:=48998;     /// устанавливаем номер порта для соединения
 //RDPsession1.Properties.Property_['PortProtocol']:=2;  /// устанавливаем тип протокола 0- ipv4,ipv6  2 - ipv4  23 - ipv6
Form1.Memo1.Lines.Add('MyRDPSession.Open');
MyRDPSession.Open;
Form1.Memo1.Lines.Add('Создаем string');
MyInvation:=MyRDPSession.Invitations.CreateInvitation((WideString(StAuthString)),
(WideString(StGroupName)),(WideString(StPassword)),10);
result:= WideString(MyInvation.ConnectionString)
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


procedure MRD.Code(var text: Widestring; password: WideString;   //// шифровка дешифрация текста для файлов
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

procedure MRD.ServerSocket1ClientRead(Sender: TObject; //// socket получаем от клиента сообщение
  Socket: TCustomWinSocket);                              //// отправляем клиенту ключ в сообщении
  var
  i,z,l:integer;
  ClientMessage,answer,UnloadString:Widestring;
  ListFile,ListFile2:TstringList;
  UnloadList:TStringList;
  s,SettingFile:Widestring;
begin

 if LoadFile=true then     //// замена файла настроек
  begin
  ClientMessage:= (Socket.ReceiveText);
  ListFile:=TstringList.Create;
  ListFile2:=TstringList.Create;
  Form1.memo1.Lines.Add('Без расшифровки - '+ClientMessage);
  ClientMessage:=crypt(ClientMessage);
  Form1.memo1.Lines.Add('Расшифровка - '+ClientMessage);
  ListFile.SetText(Pchar(ClientMessage));
  Form1.memo1.Lines.Add('memo1.Lines:= ListFile - ');
  Form1.memo1.Lines:= ListFile;

  Mini.SetStrings(ListFile);/// Обновляем файл настроек в памяти
  ///////////////////////////////////////////////////
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
   if serverport<>Mini.ReadInteger('Server','Port',48999) then
    begin
    serverport:=Mini.ReadInteger('Server','Port',48999);
    MyServerSocket.Active:=false;
    MyServerSocket.Port := ServerPort;   /// порт сервера
    MyServerSocket.Active:=true;
    end;
  ////////////////////////////////////////////////////////////////////////////

    for i := 0 to ListFile.Count-1 do
               begin
                s:=ListFile[i];
                code(s, '12345', false);
                ListFile2.Add(s);
                end;
  ListFile2.SaveToFile(ExtractFilePath(Application.ExeName)+'client.dat',TEncoding.Unicode);


  LoadFile:=False;
  ListFile.Free;
  ListFile2.Free;
  LoadFile:=false;
  exit;
  end;

  {От клиента получено сообщение - выводим его в Memo1}
  Form1.Memo1.Lines.add( 'Получаем сообщение от клиента');
  ClientMessage:= (Socket.ReceiveText);
  Form1.Memo1.Lines.add('CODE in client - '+ClientMessage );
  ClientMessage:=Crypt(ClientMessage); /// дешифрация входящего сообщения
  Form1.Memo1.Lines.add('DECODE in client - r - '+ClientMessage );

  if clientMessage='<Alt+Ctrl+Del>' then
    begin


    end;

  for I := 0 to ClientLogin.Count-1 do
    begin
     if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])+'<settingFile>') and (Leveluser[i]='0') then
       begin
       LoadFile:=true; //// включаем загрузку файлов
       Form1.memo1.Lines.Add('Включаю загрузку файла');
       exit;
       end;
     end;


  for I := 0 to ClientCount do  /////// если логин и пароль верны
   begin
    if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])) then
      begin
      Form1.Memo1.Lines.add('> ' + ClientMessage);
      Form1.Memo1.Lines.add( 'Отправляем сообщение Администратору');
       if UserConnect=false then  //// если нет подключенных пользователей то создаем новую строку
          begin
          Form1.Memo1.Lines.add( 'Создаем новую строку');
          MyStringConnect:=(NewConnectString);
          Form1.Memo1.Lines.Add(MyStringConnect);
          MyStringConnect:=Crypt(MyStringConnect);  /// шифрация
          end;
     // for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do    //// отправляем
       MyServerSocket.Socket.Connections[Connectuser-1].SendText(MyStringConnect);
      currentLevelUser:=strtoint(LevelUser[i]);
      Form1.memo1.Lines.Add('Строка - '+MyStringConnect) ;
        if currentLevelUser=0 then
            begin
              sleep(2000);
              SettingFile:=Crypt('<unloadlist>');
                    // for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
               MyServerSocket.Socket.Connections[Connectuser-1].SendText(SettingFile);
              sleep(2000);
              unloadList:=TstringList.Create;
              Mini.GetStrings(unloadList);
              for z := 0 to unloadList.Count-1 do
                begin
                  UnloadString:=UnloadString+(WideString(unloadList[z]+#13#10));
                end;
               Form1.memo1.Lines.Add('Без шифрации - '+UnloadString);
               UnloadString:=crypt(UnloadString);
               Form1.memo1.Lines.Add('Шифрация -  '+UnloadString);
               Form1.memo1.Lines.Add('отправляю файл настроек');
               //for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
               MyServerSocket.Socket.Connections[Connectuser-1].SendText(UnloadString);
              unloadList.Free;
            end;
       exit;
      end;
    end;

  for I := 0 to ClientCount do  /// если клиент указал не верный пароль
    begin
       if (ClientMessage<>WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])) then
        begin
        answer:=WideString('<ErrorUser>');
        answer:=crypt(answer);  /// шифрация
        currentLevelUser:=6;
        //for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
         MyServerSocket.Socket.Connections[Connectuser-1].SendText(answer);
        Form1.Memo1.Lines.add('Ответ  CODE - r - '+answer); /// ответ в зашифрованном виде
        Form1.Memo1.Lines.add( 'Клиент указал неверный пароль');
        exit;
        end;
    end;
end;


procedure MRD.Start;  /// загрузка программы


begin



end;


procedure MRD.RDPSession1ApplicationOpen(ASender: TObject;
  const pApplication: IDispatch);
begin
//memo1.Lines.Add('Создается новое приложение');
end;

procedure MRD.RDPSession1ApplicationUpdate(ASender: TObject;
  const pApplication: IDispatch);
begin
// memo1.Lines.Add('изменение общего свойства объекта приложения');
end;


procedure MRD.RDPSession1AttendeeDisconnected(ASender: TObject;
  const pDisconnectInfo: IDispatch);   /// отключение клиента
  var
  i:integer;
begin
pDisconnectInfo.GetTypeInfoCount(i);
Form1.memo1.Lines.Add('участник отключается от сеанса i='+inttostr(i));
Form1.memo1.Lines.Add('Connectuser='+inttostr(Connectuser));
if ConnectUser=0 then
  begin
  UserConnect:=false;   //// нет подключенных пользователей
  Form1.memo1.Lines.Add('MyRDPSession.Disconnect');
  MyRDPSession.Disconnect;        /// отключение сессии
  end;
end;

procedure MRD.RDPSession1AttendeeConnected(ASender: TObject; //// подключение клиента
  const pAttendee: IDispatch);
  var
  z,m:integer;
 Myclient:IRDPSRAPIAttendee;

begin

pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// получаем интерфейс клиента
z:=Myclient.Id;   /// определяем id клиента


begin
     Form1.memo1.Lines.Add('участник подключается к сеансу z - ' +inttostr(z));
    if currentLevelUser=2 then
      begin
       m:=MessageDlg('Разрешить удаленному помошнику'+#13#10+ 'подключится к Вашему компьютеру!!!'
       ,mtConfirmation,[mbYes,mbCancel], 0);
        if m=mrYes then
           begin
            form1.memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_VIEW');
           MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_VIEW;
           Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
           Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
           UserConnect:=true;
           end
          else
            begin
             Form1.memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_NONE');
             MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_NONE;
             Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
             Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
             MyRDPSession.Disconnect;
             UserConnect:=false;
            end;
      end;

      case currentLevelUser of
        0:
        begin
        Form1.memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_VIEW');
        MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_VIEW;
        Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
        Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
        UserConnect:=true;
        end;
        1:
        begin
        Form1.memo1.Lines.Add('Уровень привелегий  CTRL_LEVEL_VIEW');
        MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_VIEW;
        Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
        Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
        UserConnect:=true;
        end;
      end;
end;

end;


procedure MRD.ServerSocket1Accept(Sender: TObject; Socket: TCustomWinSocket);
begin
 {Здесь сервер принимает клиента ServerSocket1.Socket.ActiveConnections}
  Form1.Memo1.Lines.add('Client connection accepted count user-'+inttostr(MyServerSocket.Socket.ActiveConnections-1));
  Form1.Memo1.Lines.add('Компьютер - '+Socket.RemoteHost+' / IP-адрес - '+Socket.RemoteAddress
  +' / Порт - '+inttostr(socket.RemotePort)+'/Участник №'+inttostr(Connectuser));

end;

function MRD.findLevel(s:string):integer; //// поиск уровня привелегий пользователя
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

procedure MRD.RDPSession1ControlLevelChangeRequest(ASender: TObject;
  const pAttendee: IDispatch; RequestedLevel: TOleEnum);
  var
  z:integer;                 /// Смена уровня привелегий в зависимости от CurrentLevelUser
 Myclient:IRDPSRAPIAttendee;
  begin                          //// как подключить второго участника
 pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// получаем интерфейс клиента




   begin
    if findlevel(string(Myclient.RemoteName))=2 then
     begin
     Form1.memo1.Lines.Add('Уровнь привелегий для пользователя с правми на просмотр');
     Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
     Form1.memo1.Lines.Add('ID - '+inttostr(Myclient.Id));
     Form1.memo1.Lines.Add('RemoteName - '+string(Myclient.RemoteName));
     exit;
     end;
    Form1.memo1.Lines.Add('i='+ inttostr(z));   //// остается вопрос по индексу участника

     begin      //// установка привелегий для остальных пользователей
      Form1.memo1.Lines.Add('Смена уровня привелегий - '+ inttostr(z));
      Myclient.ControlLevel:=RequestedLevel;
      Form1.memo1.Lines.Add('ID - '+inttostr(Myclient.Id));
      Form1.memo1.Lines.Add('RemoteName - '+string(Myclient.RemoteName));
     end
     //// изменение уровня привелегий подключаемого
   end;
end;



procedure MRD.RDPSession1AttendeeUpdate(ASender: TObject;
  const pAttendee: IDispatch);
begin
  //    memo1.Lines.Add('изменяется одно из значений свойств для участника');
end;

procedure MRD.RDPSession1ChannelDataReceived(ASender: TObject;
  const pChannel: IInterface; lAttendeeId: Integer; const bstrData: WideString);
begin
 // memo1.Lines.Add(' данные принимаются от участника сеанса.');
end;

procedure MRD.RDPSession1ChannelDataSent(ASender: TObject;
  const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
begin
 //memo1.Lines.Add(' данные отправляются участнику сеанса.');
end;

procedure MRD.RDPSession1ConnectionAuthenticated(Sender: TObject);
begin
  //memo1.Lines.Add(' Соединение аутентифицированно.');
end;

procedure MRD.RDPSession1ConnectionEstablished(Sender: TObject);
begin

  //memo1.Lines.Add(' установлено соединение с сервером.');
end;

procedure MRD.RDPSession1ConnectionFailed(Sender: TObject);
begin
 // memo1.Lines.Add('клиент не может подключиться к серверу');
end;

procedure MRD.RDPSession1ConnectionTerminated(ASender: TObject; discReason,
  ExtendedInfo: Integer);
begin

  //memo1.Lines.Add('соединение клиента с сервером закрывается');
end;


procedure MRD.RDPSession1Error(ASender: TObject; ErrorInfo: OleVariant);
begin
 //memo1.Lines.Add(' В сеансе возникла критическая ошибка');
end;

procedure MRD.RDPSession1FocusReleased(ASender: TObject;
  iDirection: Integer);
begin
 // memo1.Lines.Add(' общее окно верхнего уровня набрало или потеряло фокус');
end;

procedure MRD.RDPSession1GraphicsStreamPaused(Sender: TObject);
begin
// memo1.Lines.Add('графический поток приостановлен');
end;

procedure MRD.RDPSession1GraphicsStreamResumed(Sender: TObject);
begin
  // memo1.Lines.Add('графический поток восстановлен');
end;

procedure MRD.RDPSession1SharedDesktopSettingsChanged(ASender: TObject;
  width, height, colordepth: Integer);
begin
  //memo1.Lines.Add('изменении общей настройки рабочего стола');
end;

procedure MRD.RDPSession1SharedRectChanged(ASender: TObject; left, top,
  right, bottom: Integer);
begin
 // memo1.Lines.Add('изменяется размер общего окна верхнего уровня приложения');
end;

procedure MRD.RDPSession1WindowClose(ASender: TObject;
  const pWindow: IDispatch);
begin
 // memo1.Lines.Add('закрывается закрываемое окно верхнего уровня');
end;

procedure MRD.RDPSession1WindowOpen(ASender: TObject;
  const pWindow: IDispatch);
begin
  //memo1.Lines.Add('приложения создает открытое окно верхнего уровня');
end;

procedure MRD.RDPSession1WindowUpdate(ASender: TObject;
  const pWindow: IDispatch);
begin
  //memo1.Lines.Add('Произошло изменение одного из свойств объекта окна');
end;



procedure MRD.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin

  {Здесь клиент подсоединяется}
  Form1.Memo1.Lines.add('Client connected');
  Connectuser:=MyServerSocket.Socket.ActiveConnections;

end;

procedure MRD.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   {Здесь клиент отсоединяется}
  //Memo1.Lines.add('Client disconnected');
  Connectuser:=MyServerSocket.Socket.ActiveConnections-1;
  //LabelEdEdit1.Text:= Inttostr(ServerSocket1.Socket.ActiveConnections-1);
end;

procedure MRD.ServerSocket1ClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent; var ErrorCode: Integer);
begin
 {Произошла ошибка - выводим ее код}
  //Memo1.Lines.add('Ошибка подключения:  ' + SysErrorMessage(ErrorCode));
end;



procedure MRD.ServerSocket1ClientWrite(Sender: TObject;
  Socket: TCustomWinSocket);

begin
  {Теперь можно слать данные в сокет}
  //Memo1.Lines.add('Теперь можно слать данные в сокет');

end;

procedure MRD.ServerSocket1GetSocket(Sender: TObject; Socket: NativeInt;
  var ClientSocket: TServerClientWinSocket);
begin
//Memo1.Lines.add('Get socket');
end;

procedure MRD.ServerSocket1GetThread(Sender: TObject;
  ClientSocket: TServerClientWinSocket; var SocketThread: TServerClientThread);
begin
// Memo1.Lines.add('GetThread');
end;

procedure MRD.ServerSocket1Listen(Sender: TObject; Socket: TCustomWinSocket);
begin
  {Здесь сервер "прослушивает" сокет на наличие клиентов}
 // Memo1.Lines.Add('Прослушиваю порт на наличие клиентов ' + IntToStr(ServerSocket1.Port));
end;


procedure MRD.ServerSocket1ThreadEnd(Sender: TObject;
  Thread: TServerClientThread);
begin
// Memo1.Lines.add('ThreadEnd');
end;

procedure MRD.ServerSocket1ThreadStart(Sender: TObject;
  Thread: TServerClientThread);
begin
// Memo1.Lines.add('ThreadStart');
end;




procedure MRD.Execute;
var
SettingIni:Tinifile;
mylist,mylist2:TstringList;
s:WideString;
a,b,i:integer;
begin

   form1.Memo1.Clear;
 MyRDPSession:= TRDPSession.Create(application);
 MyRDPSession.OnAttendeeConnected:=RDPSession1AttendeeConnected;
 MyRDPSession.OnApplicationOpen:= RDPSession1ApplicationOpen;
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

 MyServerSocket:=TServerSocket.Create(Application);
 MyServerSocket.OnClientRead:=ServerSocket1ClientRead;
 MyServerSocket.OnListen:=ServerSocket1Listen;
 MyServerSocket.OnAccept:=ServerSocket1Accept;
 MyServerSocket.OnClientConnect:=ServerSocket1ClientConnect;
 MyServerSocket.OnClientDisconnect:=ServerSocket1ClientDisconnect;
 MyServerSocket.OnClientError:=ServerSocket1ClientError;
 MyServerSocket.OnClientRead:=ServerSocket1ClientRead;
 MyServerSocket.OnClientWrite:=ServerSocket1ClientWrite;
 MyServerSocket.OnGetSocket:=ServerSocket1GetSocket;
 MyServerSocket.OnGetThread:=ServerSocket1GetThread;
 MyServerSocket.OnThreadEnd:=ServerSocket1ThreadEnd;
 MyServerSocket.OnThreadStart:=ServerSocket1ThreadStart;




 ClientLogin:=TstringList.Create;
 ClientPass:=TstringList.Create;
 LevelUser:=TstringList.Create;
 LoadFile:=false;      //// отключем загрузку файлов
 CurrentLevelUser:=6; //// уровень доступа пользователя, устанавливаем никакой
 UserConnect:=false; ///// нет подключенных пользователей
 if FileExists(ExtractFilePath(Application.ExeName) + 'client.dat')=false then
        begin
         SettingIni:=Tinifile.Create(ExtractFilePath(Application.ExeName) + 'client.dat');
         SettingIni.WriteString('RemoteClient','ClientName0','Administrator');
         SettingIni.WriteString('RemoteClient','ClientPasswd0','');
         SettingIni.WriteInteger('RemoteClient','Access0',0);
         SettingIni.Writeinteger('RemoteClient','Count',0);
         SettingIni.Writeinteger('Server','Port',48999);
         SettingIni.Free;
         ///////////////////////// начало шифрации файла
        mylist:=Tstringlist.Create;
        mylist2:=Tstringlist.Create;
           begin
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
          end;
         mylist.Free;
         mylist2.Free;
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

 MyServerSocket.Port := ServerPort;   /// порт сервера
 MyServerSocket.Active:=true;
 Form1.memo1.Lines:=mylist2;
 mylist2.Free;
 Form1.Memo1.Lines.Add('Server starting');
end;

end.
