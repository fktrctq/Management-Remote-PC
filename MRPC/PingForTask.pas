unit PingForTask;


interface

function PingGetaddrinfo( var resIp:string; HostName: AnsiString;TimeoutMS: cardinal): boolean;
function PingIdIcmp(var resIP:string; HostName:string;timeout:integer):boolean;
function PingGetHostByName( var resIp:string; HostName: AnsiString; TimeoutMS: cardinal): boolean;
implementation

uses Windows, SysUtils, WinSock2,IdIcmpClient,System.Classes,Vcl.forms;

function IcmpCreateFile: THandle; stdcall; external 'iphlpapi.dll';
function IcmpCloseHandle(icmpHandle: THandle): boolean; stdcall;
  external 'iphlpapi.dll';
function IcmpSendEcho(icmpHandle: THandle; DestinationAddress: In_Addr;
  RequestData: Pointer; RequestSize: Smallint; RequestOptions: Pointer;
  ReplyBuffer: Pointer; ReplySize: DWORD; Timeout: DWORD): DWORD; stdcall;
  external 'iphlpapi.dll';






function Log_write(Level:integer;fname,text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        FreeAndNil(f);
      end;
except
exit;
end;

end;

//######################################################################################################################
function PingGetHostByName( var resIp:string; HostName: AnsiString; TimeoutMS: cardinal): boolean;
const
  rSize = $400;
type
  TEchoReply = packed record
    Addr: In_Addr;
    Status: DWORD;
    RoundTripTime: DWORD;
  end;
  PEchoReply = ^TEchoReply;

var
  e: PHostEnt;
  a: PInAddr;
  h: THandle;
  d: string;
  r: array [0 .. rSize - 1] of byte;
  i: cardinal;
  WSAData: TWSAData;
begin
try
  if WSAStartup($0101, WSAData) <> 0 then
  begin
   // raise Exception.Create('WSAStartup');
   Log_write(1,'TASK',HostName+' WSAStartup. '+SysErrorMessage(WSAGetLastError));
   result:=false;
  end;

  e := gethostbyname(PAnsiChar(HostName));
  if e<>nil then
  Begin
    if e.h_addrtype = AF_INET then
    begin
      Pointer(a) := e.h_addr^;
      resIp:= inet_ntoa(PInAddr(e.h_addr_list^)^);
      d := FormatDateTime('yyyymmddhhnnsszzz', Now);
      h := IcmpCreateFile;
      if h <> INVALID_HANDLE_VALUE then
      begin
        try
        i := IcmpSendEcho(h, a^, PChar(d), Length(d), nil, @r[0], rSize, TimeoutMS);
        Result := (i <> 0) and (PEchoReply(@r[0]).Status = 0);
        finally
        IcmpCloseHandle(h);
        end;
      end
      else result:=false;
    end
    else
    begin
     //Log_write(2,'Hardware',HostName+' Ошибка Gethostbyname. '+SysErrorMessage(WSAGetLastError));
     result:=false;
    end;
  End;
  if WSACleanup <> 0 then
  begin
  Log_write(1,'TASK',HostName+' WSACleanup. '+SysErrorMessage(WSAGetLastError));
  //raise Exception.Create('WSACleanup');
  end;

except on E: Exception do
begin
Log_write(2,'TASK',HostName+' Общая ошибка GetHostByName. '+e.Message);
result:=false;
end;
end;
end;

//##############################################################################################
function PingGetaddrinfo( var resIp:string; HostName: AnsiString;TimeoutMS: cardinal): boolean;
const
  rSize = $400;
type
    //http://msdn.microsoft.com/en-us/library/ms737529%28v=VS.85%29.aspx
    Paddrinfo = ^addrinfo;
    addrinfo = packed record
      ai_flags     : Integer;
      ai_family    : Integer;
      ai_socktype  : Integer;
      ai_protocol  : Integer;
      ai_addrlen   : Cardinal;
      ai_canonname : PPAnsiChar;
      ai_addr      : LPSOCKADDR;
      ai_next      : ^addrinfo;
    end;

    TEchoReply = packed record
    Addr: In_Addr;
    Status: DWORD;
    RoundTripTime: DWORD;
   end;
  PEchoReply = ^TEchoReply;

var  rez1, rez2 : Paddrinfo;
     lwsaData : WSAData;
     error : DWORD;
     hWs2_32 : THandle;
     i: cardinal;
     d: string;
     h: THandle;
     r: array [0 .. rSize - 1] of byte;
     getaddrinfo : function ( pNodeName : PAnsiChar;
                              pServiceName : PAnsiChar;
                              const pHints : Paddrinfo;
                              out ppResult : Paddrinfo
                             ) : DWORD; stdcall;
     freeaddrinfo : function ( ai : Paddrinfo) : DWORD; stdcall;
begin
try
  hWs2_32 := LoadLibraryA('Ws2_32.dll');
  getaddrinfo := GetProcAddress(hWs2_32,'getaddrinfo');
  freeaddrinfo := GetProcAddress(hWs2_32,'freeaddrinfo');
  if not (Assigned(getaddrinfo) and Assigned (freeaddrinfo)) then
   begin
    if hWs2_32 <> 0 then FreeLibrary(hWs2_32);
    Log_write(2,'TASK',HostName+' Необходимые процедуры или функции не найдены в библиотеке Ws2_32.dll - '+SysErrorMessage(GetLastError));
    result:=false;
    Exit;
  end;
  try
    if WSAStartup(MakeWord(2,2),lwsaData) <> 0 then
    begin
    Log_write(2,'TASK',HostName+' Ошибка при вызове функции WSAStartup. '+SysErrorMessage(WSAGetLastError));
    result:=false;
    end;
    rez1 := nil;
    rez2 := nil;
    New(rez1);
    with rez1^ do
    begin
      ai_flags:=0;
      ai_family:=af_inet;
      ai_socktype:=SOCK_STREAM;
      ai_protocol:=IPPROTO_TCP;
      ai_addrlen:=0;
      ai_canonname:=nil;
      ai_addr:=nil;
      ai_next:=nil;
    end;
    error := getaddrinfo(PAnsiChar(HostName),'hostnames',rez1,rez2);  // hostnames
    if error = 0 then
     begin
      resIp:= inet_ntoa(SOCKADDR_IN(rez2^.ai_addr^).sin_addr);
      if rez2^.ai_family = AF_INET then
        Begin
            d := FormatDateTime('yyyymmddhhnnsszzz', Now);
            h := IcmpCreateFile;
            if h <> INVALID_HANDLE_VALUE then
             begin
              try
              i := IcmpSendEcho(h, SOCKADDR_IN(rez2^.ai_addr^).sin_addr, PChar(d), Length(d), nil, @r[0], rSize, TimeoutMS);
              Result := (i <> 0) and (PEchoReply(@r[0]).Status = 0);
              finally
              IcmpCloseHandle(h);
              end;
             end
            else result:=false;
        End else result:=false;
      end
     else
       begin
       //Log_write(1,'Hardware',HostName+' Getaddrinfo. '+SysErrorMessage(WSAGetLastError));
       result:=false;
       end;

  finally
    WSACleanup;
    if rez1 <> nil then begin
      Dispose(rez1);
      rez1 := nil;
    end;
    if rez2 <> nil then begin
      freeaddrinfo(rez2);
      rez2 := nil;
    end;
    if hWs2_32 <> 0 then FreeLibrary(hWs2_32);
  end;

except on E: Exception do
begin
Log_write(2,'TASK',HostName+' Общая ошибка Getaddrinfo. '+e.Message);
result:=false;
end;
end;
end;


//############################################################################

function PingIdIcmp(var resIP:string; HostName:string;timeout:integer):boolean;
var
  IdIcmpClient: TIdIcmpClient;
begin
try
  try
    Result := false;
    IdIcmpClient := TIdIcmpClient.Create;
    IdIcmpClient.PacketSize := 32;
    IdIcmpClient.Port := 0;
    IdIcmpClient.Protocol := 1;
    IdIcmpClient.ReceiveTimeout := timeout;
    IdIcmpClient.host := HostName;
    try
      IdIcmpClient.ping;
      if (IdIcmpClient.ReplyStatus.Msg <> 'Echo') or
        (IdIcmpClient.ReplyStatus.FromIpAddress = '0.0.0.0') then
        begin
        resIP:='';
        Result := false;
        end
      else
        begin
          resIP:=IdIcmpClient.ReplyStatus.FromIpAddress;
          Result := true;
        end;
    except
      begin
      resIP:='';
      Result := false;
      end;
    end;
  finally
    IdIcmpClient.Free;
  end;
except on E: Exception do
begin
Log_write(2,'TASK',HostName+' Общая ошибка IdIcmp. '+e.Message);
result:=false;
end;
end;
end;




end.

