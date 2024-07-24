unit ADWork;

interface
uses Windows, Messages, System.SysUtils, System.Variants, System.Classes, Graphics,
   Controls,ShellApi,Vcl.ComCtrls;
const
  netapi32lib = 'netapi32.dll';

type
   // ��������� ��� ��������� ���������� � ������� ������� WKSTA_INFO_100
  PWkstaInfo100 = ^TWkstaInfo100;
  TWkstaInfo100 = record
    wki100_platform_id  : DWORD;
    wki100_computername : PWideChar;
    wki100_langroup     : PWideChar;
    wki100_ver_major    : DWORD;
    wki100_ver_minor    : DWORD;
  end;
   // ��������� ��� ��������� ���������� � ������� ������� WKSTA_INFO_102
  PWkstaInfo102 = ^TWkstaInfo102;
  TWkstaInfo102 = record
    wki102_platform_id  : DWORD;
    wki102_computername : PWideChar;
    wki102_langroup     : PWideChar;
    wki102_ver_major    : DWORD;
    wki102_ver_minor    : DWORD;
    wki102_lanroot      : PWideChar;
    wki102_logged_on_users :DWORD;
  end;

  // C�������� ��� ����������� DNS ����� ���������� ������
  TDomainControllerInfoA = record
    DomainControllerName: LPSTR;
    DomainControllerAddress: LPSTR;
    DomainControllerAddressType: ULONG;
    DomainGuid: TGUID;
    DomainName: LPSTR;
    DnsForestName: LPSTR;
    Flags: ULONG;
    DcSiteName: LPSTR;
    ClientSiteName: LPSTR;
  end;
  PDomainControllerInfoA = ^TDomainControllerInfoA;

     // ��������� ��� ����������� �����
  PNetDisplayGroup = ^TNetDisplayGroup;
  TNetDisplayGroup = record
    grpi3_name: LPWSTR;
    grpi3_comment: LPWSTR;
    grpi3_group_id: DWORD;
    grpi3_attributes: DWORD;
    grpi3_next_index: DWORD;
  end;


  // ��������� ��� ����������� ������� �������
  PNetDisplayMachine = ^TNetDisplayMachine;
  TNetDisplayMachine = record
    usri2_name: LPWSTR;
    usri2_comment: LPWSTR;
    usri2_flags: DWORD;
    usri2_user_id: DWORD;
    usri2_next_index: DWORD;
  end;

  // ��������� ��� ����������� �������������
  PNetDisplayUser = ^TNetDisplayUser;
  TNetDisplayUser = record
    usri1_name: LPWSTR;
    usri1_comment: LPWSTR;
    usri1_flags: DWORD;
    usri1_full_name: LPWSTR;
    usri1_user_id: DWORD;
    usri1_next_index: DWORD;
  end;

    // ��������� ��� ����������� ������������� ������������� ������
  // ��� ����� � ������� ������ ������������
  PGroupUsersInfo0 = ^TGroupUsersInfo0;
  TGroupUsersInfo0 = record
    grui0_name: LPWSTR;
  end;
  TStringArray = array of string;
////////////////////////////////////////////////
///  ��������� ��� ��������� ���������� � ������������� �� �����������

  THostInfo = record
     username     : PWideChar;
     logon_domain : PWideChar;
     other_domains: PWideChar;
     logon_server : PWideChar;
   end;

   WKSTA_USER_INFO_0 = packed record
     wkui0_username: PWideChar;
   end;
   PWKSTA_USER_INFO_0 = ^WKSTA_USER_INFO_0;

     // ������� ������� ����������� ��� ����������� ��������� ����������
//  function NetWkstaUserEnum(ServerName: PWideChar; Level: DWORD;
 //   Bufptr: Pointer; prefmaxlen:DWORD ;entriesread:LPDWORD;
 //   totalentries:LPDWORD; resumehandle:LPDWORD):DWORD; stdcall;
 //    external netapi32lib;
 function NetWkstaUserEnum(servername: PWideChar;
   // ��������� �� ������, ������� ��������� DNS ��� NetBIOS-���
  // ��������� ������, �� ������� ������ ����������� �������.
  // ���� ���� �������� nil, ������������ ��������� ���������.
  level: DWORD;
   // Level = 0: ������� ����� �������������, ������� � ������ ������ ����� �� ������� �������.
  var bufptr: Pointer;   // ��������� �� �����, ������� �������� ������
  prefmaxlen: DWORD;
   //  ��������� ���������������� ������������ ����� ������������ ������ � ������.
  var entriesread: PDWord;
   // ��������� �� ��������, ������� �������� ���������� ���������� ������������� ���������.
  var totalentries: PDWord;  // ����� ���������� �������
  var resumehandle: PDWord)
   // �������� ���������� ������, ������� ������������ ��� ����������� ������������� ������
  : Longint;
   stdcall; external 'netapi32.dll' Name 'NetWkstaUserEnum';

  function NetApiBufferFree(Buffer: Pointer): DWORD; stdcall;
    external netapi32lib;
  function NetWkstaGetInfo(ServerName: PWideChar; Level: DWORD;
    Bufptr: Pointer): DWORD; stdcall; external netapi32lib;
  function NetGetDCName(ServerName: PWideChar; DomainName: PWideChar;
    var Bufptr: PWideChar): DWORD; stdcall; external netapi32lib;
  function DsGetDcName(ComputerName, DomainName: PChar; DomainGuid: PGUID;
    SiteName: PChar; Flags: ULONG;
    var DomainControllerInfo: PDomainControllerInfoA): DWORD; stdcall;
    external netapi32lib name 'DsGetDcNameA';
  function NetQueryDisplayInformation(ServerName: PWideChar; Level: DWORD;
    Index: DWORD; EntriesRequested: DWORD; PreferredMaximumLength: DWORD;
    var ReturnedEntryCount: DWORD; SortedBuffer: Pointer): DWORD; stdcall;
    external netapi32lib;
  function NetGroupGetUsers(ServerName: PWideChar; GroupName: PWideChar; Level: DWORD;
    var Bufptr: Pointer; PrefMaxLen: DWORD; var EntriesRead: DWORD;
    var TotalEntries: DWORD; ResumeHandle: PDWORD): DWORD; stdcall;
    external netapi32lib;
  function NetUserGetGroups(ServerName: PWideChar; UserName: PWideChar; Level: DWORD;
    var Bufptr: Pointer; PrefMaxLen: DWORD; var EntriesRead: DWORD;
    var TotalEntries: DWORD): DWORD; stdcall; external netapi32lib;
  function NetEnumerateTrustedDomains(ServerName: PWideChar;
    DomainNames: PWideChar): DWORD; stdcall; external netapi32lib;

// function ConvertStringSidToSid(StringSid: PChar; out Sid: PSID):BOOL; stdcall;
 //   external advapi32 name {$IFDEF UNICODE} 'ConvertStringSidToSidW' {$ELSE}
  // 'ConvertStringSidToSidA'{$ENDIF};

 function ConvertSidToStringSid(Sid: PSID; out StringSid: PChar):BOOL; stdcall;
   external advapi32 name {$IFDEF UNICODE} 'ConvertSidToStringSidW' {$ELSE}
   'ConvertSidToStringSidA'{$ENDIF};

   function GetDomainController(const DomainName: String): String; ///������� ���������� ����� ����������� ������
   function EnumAllTrustedDomains: Boolean;
   function EnumAllWorkStation: TstringList;
   function EnumAllGroups: TstringList;
   function GetCurrentComputerName: String;
   function GetCurrentComputerOS(s:PWideChar): String;
   function GetAllGroupUsers(const GroupName: String): TstringList;
   function GetDNSDomainName(const DomainName: String): String;
   var
    CurrentDomainName:string;
   const
    NERR_Success = NO_ERROR;

implementation

  //////////////////////////////////////////////////////////////////////////
function GetCurrentComputerName: String; /// ��������� ���������� � ���������� � ���� � ������
var
  Info: PWkstaInfo100;
  Error: DWORD;
begin
try
  // � ��� ����� �� ������������� ��������� ��������
  Error := NetWkstaGetInfo(nil, 100, @Info);
  if Error <> 0 then
    raise Exception.Create(SysErrorMessage(Error));
  // ��� �����, ����� ������� ���������� ������� ���������, �� ������� � ���������, ��� ��� ����� :)

  // � ������ ��� ���������� � ����
  Result := Info^.wki100_computername;
  // � ��� �� ����������
  CurrentDomainName := info^.wki100_langroup;
Except
   on E: Exception do
      begin
        //memo1.Lines.add('������ �������� ������ �����������  :'+E.Message);
      end;
end;
end;

///////////////////////////////////////////////////////  ������ �� (5.0,6.0,6.1 � �.�)
function GetCurrentComputerOS(s:PWideChar): String;
var
  Info: PWkstaInfo100;
  Error: DWORD;
  begin
  // � ��� ����� �� ������������� ��������� ��������
   try
  Error := NetWkstaGetInfo((s), 101, @Info);   /// s ��� ����������
  if Error <> 0 then
    raise Exception.Create(SysErrorMessage(Error));
   Result:=(inttostr(Info^.wki100_ver_major)+'.'+inttostr(Info^.wki100_ver_minor));
except on E:Exception do result:=(E.Message);
  end;
end;
///////////////////////////////////////////////////////////////////////////////////////////////
function EnumAllTrustedDomains: Boolean;
var
  Tmp, DomainList: PWideChar;
begin
  // ���������� ������������������� ������� NetEnumerateTrustedDomains
  // (������ �� �����, � ������ �������� ��� �� ���������������?)
  // ��� ��� ����� ������, �� ���� ��� ���������� ������, �� ����� - ������ ���������� �������
  Result := NetEnumerateTrustedDomains(StringToOleStr(GetDomainController(CurrentDomainName)),
    @DomainList) = NERR_Success;
  // ���� ����� ������� �������, ��...
  if Result then
  try
    Tmp := DomainList;
    while Length(Tmp) > 0 do
    begin
      //memTrustedDomains.Lines.Add(Tmp); // �������� ������� ������ �� �����
      Tmp := Tmp + Length(Tmp) + 1;
    end;
  finally
    // �� �������� ��� ������
    NetApiBufferFree(DomainList);
  end;
end;


//  ������ ������� �������� ���������� � ���� ������� �������� �������������� � ������
//  �������� ��� ������ ������� �� �����, ���� � ��� ��� ������� ������� �����
//  �������������� � ������ �� ������ ��, ������� ����� �������� (�� ��� ����� ������ � ���)
// =============================================================================
function EnumAllWorkStation: TstringList;
var
  Tmp, Info: PNetDisplayMachine;
  I, CurrIndex, EntriesRequest,
  PreferredMaximumLength,
  ReturnedEntryCount: Cardinal;
  Error: DWORD;
  s:string;
  z:integer;
begin
  CurrIndex := 0;
  repeat
    Info := nil;
    result:=Tstringlist.Create;
    // NetQueryDisplayInformation ���������� ���������� ������ � 100-� �������
    // ��� ���� ����� �������� ��� ���������� ������������ ������ ��������,
    // ������������ �������, ������� ���������� � ����� ������ ����������
    // ����� ����������
    EntriesRequest := 100;
    PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayMachine);
    ReturnedEntryCount := 0;
    // ��� ���������� �������, � ��� ����� �������� DNS ��� ���������� ������
    // (��� ��� IP �����), � �������� �� ����� �������� ����������
    // ��� ��������� ���������� � ������� �������� ������������ ��������� NetDisplayMachine
    // � �� ������������� 2 (������) �� ������ ���������
    Error := NetQueryDisplayInformation(StringToOleStr(GetDomainController(CurrentDomainName)), 2, CurrIndex,
      EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
    // ��� ������������ ���������� ������ ����� ��������� ����
    // 1. NERR_Success - ��� ������ ����������
    // 2. ERROR_MORE_DATA - ������ ����������, �� �������� ��� � ����� �������� ������� ��������
    if Error in [NERR_Success, ERROR_MORE_DATA] then
    try
      Tmp := Info;
      // ������� ���������� ������� ������� ������� � ���������
      for I := 0 to ReturnedEntryCount - 1 do
      begin
      result.add(copy(Tmp^.usri2_name,1,Length(Tmp^.usri2_name)-1));
      CurrIndex := Tmp^.usri2_next_index;
      Inc(Tmp);
      end;
    finally
      // ���� ������ ������
      NetApiBufferFree(Info);
    end;
  // ���� ��������� ���������� ������� ERROR_MORE_DATA
  // (�.�. ���� ��� ������) - �������� ������� ��������
  until Error in [NERR_Success, ERROR_ACCESS_DENIED];
  // �� � ���������� ��������� ����� ��� �� ��� ��������
 // Result := Error = NERR_Success;
end;

//////////////////////////////////////////////////////////////////////////////
//  �������� ������� �������, ���������� ������ ����� ����������� �������������� ������
// =============================================================================
function GetAllGroupUsers(const GroupName: String): TstringList;
var
  Tmp, Info: PGroupUsersInfo0;
  PrefMaxLen, EntriesRead,
  TotalEntries, ResumeHandle: DWORD;
  I: Integer;
  rebool:boolean;
begin
  result:=TstringList.Create;
  // ������������ �������������
  ResumeHandle := 0;
  PrefMaxLen := DWORD(-1);
  // ���������
  rebool := NetGroupGetUsers(StringToOleStr(GetDomainController(CurrentDomainName)),
    StringToOleStr(GroupName), 0, Pointer(Info), PrefMaxLen,
    EntriesRead, TotalEntries, @ResumeHandle) = NERR_Success;
  // ������� ���������...
  if rebool then
  try
    Tmp := Info;
    for I := 0 to EntriesRead - 1 do
    begin
       // �������� ������� ��������� �� ���������
    if pos('$',(Tmp^.grui0_name))<>0 then
     result.Add(copy(Tmp^.grui0_name,1,Length(Tmp^.grui0_name)-1));
    Inc(Tmp);
    end;
  finally
    // �� ��������, ��� ����� ���� ������� :)
    NetApiBufferFree(Info);
  end;
end;

function EnumAllGroups: TstringList;
var
  Tmp, Info: PNetDisplayGroup;
  I, CurrIndex, EntriesRequest,
  PreferredMaximumLength,
  ReturnedEntryCount: Cardinal;
  Error: DWORD;
begin
  CurrIndex := 0;
  result:=TstringList.Create;
  repeat
    Info := nil;
    // NetQueryDisplayInformation ���������� ���������� ������ � 100-� �������
    // ��� ���� ����� �������� ��� ���������� ������������ ������ ��������,
    // ������������ �������, ������� ���������� � ����� ������ ����������
    // ����� ����������
    EntriesRequest := 100;
    PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayGroup);
    ReturnedEntryCount := 0;
    // ��� ���������� �������, � ��� ����� �������� DNS ��� ���������� ������
    // (��� ��� IP �����), � �������� �� ����� �������� ����������
    // ��� ��������� ���������� � ������� ������������ ��������� NetDisplayGroup
    // � �� ������������� 3 (������) �� ������ ���������
    Error := NetQueryDisplayInformation(StringToOleStr(GetDomainController(CurrentDomainName)), 3, CurrIndex,
    EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
    // ��� ������������ ���������� ������ ����� ��������� ����
    // 1. NERR_Success - ��� ������ ����������
    // 2. ERROR_MORE_DATA - ������ ����������, �� �������� ��� � ����� �������� ������� ��������
    if Error in [NERR_Success, ERROR_MORE_DATA] then
    try
      Tmp := Info;
      for I := 0 to ReturnedEntryCount - 1 do
      begin
      result.add(Tmp^.grpi3_name);         // ��� ������
      // ���������� ������ � ������� ����� �������� �������� ������� (���� �����)
      CurrIndex := Tmp^.grpi3_next_index;
      Inc(Tmp);
      end;
      ///  ////////////////////////////////////////////////////
    finally
      // ����� ������ ������ ��������, ����������� ������ ������� �������� ��� ���������
      NetApiBufferFree(Info);
    end;
  // ���� ��������� ���������� ������� ERROR_MORE_DATA - �������� ������� ��������
  until Error in [NERR_Success, ERROR_ACCESS_DENIED];
  // �� � ���������� ��������� ����� ��� �� ��� ��������
  //Result := Error = NERR_Success;
end;
//  �������� DNS ��� ����������� ������
// =============================================================================
function GetDNSDomainName(const DomainName: String): String;
const
  DS_IS_FLAT_NAME = $00010000;
  DS_RETURN_DNS_NAME  = $40000000;
var
  GUID: PGUID;
  DomainControllerInfo: PDomainControllerInfoA;
begin
  GUID := nil;
  // ��� ����������� �������� ��� ����������� IP ����� ����������� ������
  // ��� ��� DNS ���, ������� �� ������� ��� ���:
  if DsGetDcName(nil, PChar(CurrentDomainName), GUID, nil,
    DS_IS_FLAT_NAME or DS_RETURN_DNS_NAME, DomainControllerInfo) = NERR_Success then
  // ��������� ������� �� �������� ��������:
  // DS_IS_FLAT_NAME - �������� ������ ��� ������
  // DS_RETURN_DNS_NAME - ���� ��������� DNS �����
  try
    Result := DomainControllerInfo^.DomainControllerName; // ��������� ������ ���...
  finally
    // ������� ��� �������, �� ����� ������...
    NetApiBufferFree(DomainControllerInfo);
  end;
end;

// =============================================================================
//  �� ��� ��� ������������ - ������ �������� ��� ����������� ������
// =============================================================================
function GetDomainController(const DomainName: String): String;
var
  Domain: WideString;
  Server: PWideChar;
begin
  Domain := StringToOleStr(DomainName);
  if NetGetDCName(nil, @Domain[1], Server) = NERR_Success then
  try
    Result := Server;
  finally
    NetApiBufferFree(Server);
  end;
end;

 {//ledDNSName.Text := GetDNSDomainName(CurrentDomainName);
  GetCurrentComputerName
  EnumAllTrustedDomains; /// ������ ��������� �������
  EnumAllWorkStation;  /// ��� ������� ������� � ������
  EnumAllGroups;      ///��� ������� ������������ � ������
   GetCurrentComputerName;  ///  ���������� � ����������
  EnumAllUsers; /// ��� ������������ ������ }

end.
