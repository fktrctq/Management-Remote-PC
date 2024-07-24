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
    function NewConnectString():WideString;  //// ������� ����� ������
    function findLevel(s:string):integer;    //// ����� ������ ������� �������
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
    procedure Code(var text: WideString; password: WideString;  //// ��������� ����������� � ������������� �����
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
  LoadFile:boolean;    ////  �������� ����� �� ��������������
  UnloadFile:boolean; //// �������� ����� ��������������
  Mini: TMemIniFile;  //// ������������� ����
  CurrentLevelUser:byte;
  ClientCount:integer; /// ���-�� ������������� � ��� �����
  UserConnect:boolean; //// ���� �� �����������
  ConnectUser:integer;  //// ������� �������������� ������������ (���-�� ��������������)
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
 //RDPsession1.Properties.Property_['PortId']:=48998;     /// ������������� ����� ����� ��� ����������
 //RDPsession1.Properties.Property_['PortProtocol']:=2;  /// ������������� ��� ��������� 0- ipv4,ipv6  2 - ipv4  23 - ipv6
Form1.Memo1.Lines.Add('MyRDPSession.Open');
MyRDPSession.Open;
Form1.Memo1.Lines.Add('������� string');
MyInvation:=MyRDPSession.Invitations.CreateInvitation((WideString(StAuthString)),
(WideString(StGroupName)),(WideString(StPassword)),10);
result:= WideString(MyInvation.ConnectionString)
end;

function Crypt(varStr: WideString):WideString; /// ������� ���������� � ���������� ������ ��� ��������
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


procedure MRD.Code(var text: Widestring; password: WideString;   //// �������� ���������� ������ ��� ������
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

procedure MRD.ServerSocket1ClientRead(Sender: TObject; //// socket �������� �� ������� ���������
  Socket: TCustomWinSocket);                              //// ���������� ������� ���� � ���������
  var
  i,z,l:integer;
  ClientMessage,answer,UnloadString:Widestring;
  ListFile,ListFile2:TstringList;
  UnloadList:TStringList;
  s,SettingFile:Widestring;
begin

 if LoadFile=true then     //// ������ ����� ��������
  begin
  ClientMessage:= (Socket.ReceiveText);
  ListFile:=TstringList.Create;
  ListFile2:=TstringList.Create;
  Form1.memo1.Lines.Add('��� ����������� - '+ClientMessage);
  ClientMessage:=crypt(ClientMessage);
  Form1.memo1.Lines.Add('����������� - '+ClientMessage);
  ListFile.SetText(Pchar(ClientMessage));
  Form1.memo1.Lines.Add('memo1.Lines:= ListFile - ');
  Form1.memo1.Lines:= ListFile;

  Mini.SetStrings(ListFile);/// ��������� ���� �������� � ������
  ///////////////////////////////////////////////////
  ///  ��������� ������ �������
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
    MyServerSocket.Port := ServerPort;   /// ���� �������
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

  {�� ������� �������� ��������� - ������� ��� � Memo1}
  Form1.Memo1.Lines.add( '�������� ��������� �� �������');
  ClientMessage:= (Socket.ReceiveText);
  Form1.Memo1.Lines.add('CODE in client - '+ClientMessage );
  ClientMessage:=Crypt(ClientMessage); /// ���������� ��������� ���������
  Form1.Memo1.Lines.add('DECODE in client - r - '+ClientMessage );

  if clientMessage='<Alt+Ctrl+Del>' then
    begin


    end;

  for I := 0 to ClientLogin.Count-1 do
    begin
     if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])+'<settingFile>') and (Leveluser[i]='0') then
       begin
       LoadFile:=true; //// �������� �������� ������
       Form1.memo1.Lines.Add('������� �������� �����');
       exit;
       end;
     end;


  for I := 0 to ClientCount do  /////// ���� ����� � ������ �����
   begin
    if (ClientMessage=WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])) then
      begin
      Form1.Memo1.Lines.add('> ' + ClientMessage);
      Form1.Memo1.Lines.add( '���������� ��������� ��������������');
       if UserConnect=false then  //// ���� ��� ������������ ������������� �� ������� ����� ������
          begin
          Form1.Memo1.Lines.add( '������� ����� ������');
          MyStringConnect:=(NewConnectString);
          Form1.Memo1.Lines.Add(MyStringConnect);
          MyStringConnect:=Crypt(MyStringConnect);  /// ��������
          end;
     // for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do    //// ����������
       MyServerSocket.Socket.Connections[Connectuser-1].SendText(MyStringConnect);
      currentLevelUser:=strtoint(LevelUser[i]);
      Form1.memo1.Lines.Add('������ - '+MyStringConnect) ;
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
               Form1.memo1.Lines.Add('��� �������� - '+UnloadString);
               UnloadString:=crypt(UnloadString);
               Form1.memo1.Lines.Add('�������� -  '+UnloadString);
               Form1.memo1.Lines.Add('��������� ���� ��������');
               //for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
               MyServerSocket.Socket.Connections[Connectuser-1].SendText(UnloadString);
              unloadList.Free;
            end;
       exit;
      end;
    end;

  for I := 0 to ClientCount do  /// ���� ������ ������ �� ������ ������
    begin
       if (ClientMessage<>WideString('<Client>'+ClientLogin[i]+'<passwd>'+ClientPass[i])) then
        begin
        answer:=WideString('<ErrorUser>');
        answer:=crypt(answer);  /// ��������
        currentLevelUser:=6;
        //for z := 0 to ServerSocket1.Socket.ActiveConnections - 1 do
         MyServerSocket.Socket.Connections[Connectuser-1].SendText(answer);
        Form1.Memo1.Lines.add('�����  CODE - r - '+answer); /// ����� � ������������� ����
        Form1.Memo1.Lines.add( '������ ������ �������� ������');
        exit;
        end;
    end;
end;


procedure MRD.Start;  /// �������� ���������


begin



end;


procedure MRD.RDPSession1ApplicationOpen(ASender: TObject;
  const pApplication: IDispatch);
begin
//memo1.Lines.Add('��������� ����� ����������');
end;

procedure MRD.RDPSession1ApplicationUpdate(ASender: TObject;
  const pApplication: IDispatch);
begin
// memo1.Lines.Add('��������� ������ �������� ������� ����������');
end;


procedure MRD.RDPSession1AttendeeDisconnected(ASender: TObject;
  const pDisconnectInfo: IDispatch);   /// ���������� �������
  var
  i:integer;
begin
pDisconnectInfo.GetTypeInfoCount(i);
Form1.memo1.Lines.Add('�������� ����������� �� ������ i='+inttostr(i));
Form1.memo1.Lines.Add('Connectuser='+inttostr(Connectuser));
if ConnectUser=0 then
  begin
  UserConnect:=false;   //// ��� ������������ �������������
  Form1.memo1.Lines.Add('MyRDPSession.Disconnect');
  MyRDPSession.Disconnect;        /// ���������� ������
  end;
end;

procedure MRD.RDPSession1AttendeeConnected(ASender: TObject; //// ����������� �������
  const pAttendee: IDispatch);
  var
  z,m:integer;
 Myclient:IRDPSRAPIAttendee;

begin

pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// �������� ��������� �������
z:=Myclient.Id;   /// ���������� id �������


begin
     Form1.memo1.Lines.Add('�������� ������������ � ������ z - ' +inttostr(z));
    if currentLevelUser=2 then
      begin
       m:=MessageDlg('��������� ���������� ���������'+#13#10+ '����������� � ������ ����������!!!'
       ,mtConfirmation,[mbYes,mbCancel], 0);
        if m=mrYes then
           begin
            form1.memo1.Lines.Add('������� ����������  CTRL_LEVEL_VIEW');
           MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_VIEW;
           Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
           Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
           UserConnect:=true;
           end
          else
            begin
             Form1.memo1.Lines.Add('������� ����������  CTRL_LEVEL_NONE');
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
        Form1.memo1.Lines.Add('������� ����������  CTRL_LEVEL_VIEW');
        MyRDPSession.Attendees.Item[z].ControlLevel:=CTRL_LEVEL_VIEW;
        Form1.memo1.Lines.Add('ID - '+inttostr(MyRDPSession.Attendees.Item[z].Id));
        Form1.memo1.Lines.Add('RemoteName - '+string(MyRDPSession.Attendees.Item[z].RemoteName));
        UserConnect:=true;
        end;
        1:
        begin
        Form1.memo1.Lines.Add('������� ����������  CTRL_LEVEL_VIEW');
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
 {����� ������ ��������� ������� ServerSocket1.Socket.ActiveConnections}
  Form1.Memo1.Lines.add('Client connection accepted count user-'+inttostr(MyServerSocket.Socket.ActiveConnections-1));
  Form1.Memo1.Lines.add('��������� - '+Socket.RemoteHost+' / IP-����� - '+Socket.RemoteAddress
  +' / ���� - '+inttostr(socket.RemotePort)+'/�������� �'+inttostr(Connectuser));

end;

function MRD.findLevel(s:string):integer; //// ����� ������ ���������� ������������
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
  z:integer;                 /// ����� ������ ���������� � ����������� �� CurrentLevelUser
 Myclient:IRDPSRAPIAttendee;
  begin                          //// ��� ���������� ������� ���������
 pAttendee.QueryInterface(IID_IRDPSRAPIAttendee,Myclient) ;   ///// �������� ��������� �������




   begin
    if findlevel(string(Myclient.RemoteName))=2 then
     begin
     Form1.memo1.Lines.Add('������ ���������� ��� ������������ � ������ �� ��������');
     Myclient.ControlLevel:=CTRL_LEVEL_VIEW;
     Form1.memo1.Lines.Add('ID - '+inttostr(Myclient.Id));
     Form1.memo1.Lines.Add('RemoteName - '+string(Myclient.RemoteName));
     exit;
     end;
    Form1.memo1.Lines.Add('i='+ inttostr(z));   //// �������� ������ �� ������� ���������

     begin      //// ��������� ���������� ��� ��������� �������������
      Form1.memo1.Lines.Add('����� ������ ���������� - '+ inttostr(z));
      Myclient.ControlLevel:=RequestedLevel;
      Form1.memo1.Lines.Add('ID - '+inttostr(Myclient.Id));
      Form1.memo1.Lines.Add('RemoteName - '+string(Myclient.RemoteName));
     end
     //// ��������� ������ ���������� �������������
   end;
end;



procedure MRD.RDPSession1AttendeeUpdate(ASender: TObject;
  const pAttendee: IDispatch);
begin
  //    memo1.Lines.Add('���������� ���� �� �������� ������� ��� ���������');
end;

procedure MRD.RDPSession1ChannelDataReceived(ASender: TObject;
  const pChannel: IInterface; lAttendeeId: Integer; const bstrData: WideString);
begin
 // memo1.Lines.Add(' ������ ����������� �� ��������� ������.');
end;

procedure MRD.RDPSession1ChannelDataSent(ASender: TObject;
  const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
begin
 //memo1.Lines.Add(' ������ ������������ ��������� ������.');
end;

procedure MRD.RDPSession1ConnectionAuthenticated(Sender: TObject);
begin
  //memo1.Lines.Add(' ���������� ������������������.');
end;

procedure MRD.RDPSession1ConnectionEstablished(Sender: TObject);
begin

  //memo1.Lines.Add(' ����������� ���������� � ��������.');
end;

procedure MRD.RDPSession1ConnectionFailed(Sender: TObject);
begin
 // memo1.Lines.Add('������ �� ����� ������������ � �������');
end;

procedure MRD.RDPSession1ConnectionTerminated(ASender: TObject; discReason,
  ExtendedInfo: Integer);
begin

  //memo1.Lines.Add('���������� ������� � �������� �����������');
end;


procedure MRD.RDPSession1Error(ASender: TObject; ErrorInfo: OleVariant);
begin
 //memo1.Lines.Add(' � ������ �������� ����������� ������');
end;

procedure MRD.RDPSession1FocusReleased(ASender: TObject;
  iDirection: Integer);
begin
 // memo1.Lines.Add(' ����� ���� �������� ������ ������� ��� �������� �����');
end;

procedure MRD.RDPSession1GraphicsStreamPaused(Sender: TObject);
begin
// memo1.Lines.Add('����������� ����� �������������');
end;

procedure MRD.RDPSession1GraphicsStreamResumed(Sender: TObject);
begin
  // memo1.Lines.Add('����������� ����� ������������');
end;

procedure MRD.RDPSession1SharedDesktopSettingsChanged(ASender: TObject;
  width, height, colordepth: Integer);
begin
  //memo1.Lines.Add('��������� ����� ��������� �������� �����');
end;

procedure MRD.RDPSession1SharedRectChanged(ASender: TObject; left, top,
  right, bottom: Integer);
begin
 // memo1.Lines.Add('���������� ������ ������ ���� �������� ������ ����������');
end;

procedure MRD.RDPSession1WindowClose(ASender: TObject;
  const pWindow: IDispatch);
begin
 // memo1.Lines.Add('����������� ����������� ���� �������� ������');
end;

procedure MRD.RDPSession1WindowOpen(ASender: TObject;
  const pWindow: IDispatch);
begin
  //memo1.Lines.Add('���������� ������� �������� ���� �������� ������');
end;

procedure MRD.RDPSession1WindowUpdate(ASender: TObject;
  const pWindow: IDispatch);
begin
  //memo1.Lines.Add('��������� ��������� ������ �� ������� ������� ����');
end;



procedure MRD.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin

  {����� ������ ��������������}
  Form1.Memo1.Lines.add('Client connected');
  Connectuser:=MyServerSocket.Socket.ActiveConnections;

end;

procedure MRD.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   {����� ������ �������������}
  //Memo1.Lines.add('Client disconnected');
  Connectuser:=MyServerSocket.Socket.ActiveConnections-1;
  //LabelEdEdit1.Text:= Inttostr(ServerSocket1.Socket.ActiveConnections-1);
end;

procedure MRD.ServerSocket1ClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent; var ErrorCode: Integer);
begin
 {��������� ������ - ������� �� ���}
  //Memo1.Lines.add('������ �����������:  ' + SysErrorMessage(ErrorCode));
end;



procedure MRD.ServerSocket1ClientWrite(Sender: TObject;
  Socket: TCustomWinSocket);

begin
  {������ ����� ����� ������ � �����}
  //Memo1.Lines.add('������ ����� ����� ������ � �����');

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
  {����� ������ "������������" ����� �� ������� ��������}
 // Memo1.Lines.Add('����������� ���� �� ������� �������� ' + IntToStr(ServerSocket1.Port));
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
 LoadFile:=false;      //// �������� �������� ������
 CurrentLevelUser:=6; //// ������� ������� ������������, ������������� �������
 UserConnect:=false; ///// ��� ������������ �������������
 if FileExists(ExtractFilePath(Application.ExeName) + 'client.dat')=false then
        begin
         SettingIni:=Tinifile.Create(ExtractFilePath(Application.ExeName) + 'client.dat');
         SettingIni.WriteString('RemoteClient','ClientName0','Administrator');
         SettingIni.WriteString('RemoteClient','ClientPasswd0','');
         SettingIni.WriteInteger('RemoteClient','Access0',0);
         SettingIni.Writeinteger('RemoteClient','Count',0);
         SettingIni.Writeinteger('Server','Port',48999);
         SettingIni.Free;
         ///////////////////////// ������ �������� �����
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
        ////////////////////// ��������� �������� �����
        end;

 ///////////////////////////   ������ ���������� ������
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
  ///////////////////////////  ��������� ���������� ������
 for I := 0 to Mini.ReadInteger('RemoteClient','Count',0) do
   begin
   ClientLogin.Add(Mini.ReadString('RemoteClient','ClientName'+inttostr(i),'Administrator'));
   ClientPass.Add(Mini.ReadString('RemoteClient','ClientPasswd'+inttostr(i),'159753'));
   LevelUser.Add(Mini.ReadString('RemoteClient','Access'+inttostr(i),'3'));
   end;
  serverport:=Mini.ReadInteger('Server','Port',48999);
  ClientCount:=Mini.ReadInteger('RemoteClient','Count',0);

 MyServerSocket.Port := ServerPort;   /// ���� �������
 MyServerSocket.Active:=true;
 Form1.memo1.Lines:=mylist2;
 mylist2.Free;
 Form1.Memo1.Lines.Add('Server starting');
end;

end.
