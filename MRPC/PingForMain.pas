unit PingForMain;

interface

function PingGetaddrinfo( var resIp:string; HostName: AnsiString;TimeoutMS: cardinal): boolean;
function PingIdIcmp(HostName:string;timeout:integer):boolean;
function PingGetHostByName( var resIp:string; HostName: AnsiString; TimeoutMS: cardinal): boolean;
implementation

uses Windows, SysUtils, WinSock2,uMain,IdIcmpClient;

function IcmpCreateFile: THandle; stdcall; external 'iphlpapi.dll';
function IcmpCloseHandle(icmpHandle: THandle): boolean; stdcall;
  external 'iphlpapi.dll';
function IcmpSendEcho(icmpHandle: THandle; DestinationAddress: In_Addr;
  RequestData: Pointer; RequestSize: Smallint; RequestOptions: Pointer;
  ReplyBuffer: Pointer; ReplySize: DWORD; Timeout: DWORD): DWORD; stdcall;
  external 'iphlpapi.dll';








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

  if WSAStartup($0101, WSAData) <> 0 then
    raise Exception.Create('WSAStartup');

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
    else result:=false;
  End;
  if WSACleanup <> 0 then
    raise Exception.Create('WSACleanup');
end;


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
 hWs2_32 := LoadLibraryA('Ws2_32.dll');
  getaddrinfo := GetProcAddress(hWs2_32,'getaddrinfo');
  freeaddrinfo := GetProcAddress(hWs2_32,'freeaddrinfo');
  if not (Assigned(getaddrinfo) and Assigned (freeaddrinfo)) then
   begin
    if hWs2_32 <> 0 then FreeLibrary(hWs2_32);
     frmdomaininfo.Memo1.Lines.Add('Необходимые процедуры или функции не найдены в библиотеке Ws2_32.dll.');
     frmdomaininfo.Memo1.Lines.Add(SysErrorMessage(GetLastError));
     result:=false;
    Exit;
  end;
  try
    if WSAStartup(MakeWord(2,2),lwsaData) <> 0 then
    begin
      frmdomaininfo.Memo1.Lines.Add('Ошибка при вызове функции "WSAStartup".');
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
      //frmdomaininfo.Memo1.Lines.Add('----------------------------------------------');
      //frmdomaininfo.Memo1.Lines.Add('ai_family: '+ IntToStr(rez2^.ai_family));
      //frmdomaininfo.Memo1.Lines.Add('ai_socktype: '+ IntToStr(rez2^.ai_socktype));
      //frmdomaininfo.Memo1.Lines.Add('ai_addrlen: '+ IntToStr(rez2^.ai_addrlen));
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
        WSAGetLastError;
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

end;



// =============================================================================

function PingIdIcmp(HostName:string;timeout:integer):boolean;
var
  IdIcmpClient: TIdIcmpClient;
begin
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
          Result := false;
        end
      else
        begin
          frmdomaininfo.Memo1.Lines.Add('----------------------------------------------');
          frmdomaininfo.Memo1.Lines.Add(IdIcmpClient.ReplyStatus.Msg + ' - IP хоста - ' +
            IdIcmpClient.ReplyStatus.FromIpAddress + '  Время=' +
            inttostr(IdIcmpClient.ReplyStatus.MsRoundTripTime) + 'мс' + '  TTL='
            + inttostr(IdIcmpClient.ReplyStatus.TimeToLive));
          Result := true;
        end;
    except
      on E: Exception do
        begin
          Result := false;
        end;
    end;
  finally
    IdIcmpClient.Free;
  end;
end;




end.

