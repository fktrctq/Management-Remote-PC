unit ADWork;

interface
uses Windows, Messages, System.SysUtils, System.Variants, System.Classes, Graphics,
   Controls,ShellApi,Vcl.ComCtrls;
const
  netapi32lib = 'netapi32.dll';

type
   // Структура для получения информации о рабочей станции WKSTA_INFO_100
  PWkstaInfo100 = ^TWkstaInfo100;
  TWkstaInfo100 = record
    wki100_platform_id  : DWORD;
    wki100_computername : PWideChar;
    wki100_langroup     : PWideChar;
    wki100_ver_major    : DWORD;
    wki100_ver_minor    : DWORD;
  end;
   // Структура для получения информации о рабочей станции WKSTA_INFO_102
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

  // Cтруктура для определения DNS имени контролера домена
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

     // Структура для отображения групп
  PNetDisplayGroup = ^TNetDisplayGroup;
  TNetDisplayGroup = record
    grpi3_name: LPWSTR;
    grpi3_comment: LPWSTR;
    grpi3_group_id: DWORD;
    grpi3_attributes: DWORD;
    grpi3_next_index: DWORD;
  end;


  // Структура для отображения рабочих станций
  PNetDisplayMachine = ^TNetDisplayMachine;
  TNetDisplayMachine = record
    usri2_name: LPWSTR;
    usri2_comment: LPWSTR;
    usri2_flags: DWORD;
    usri2_user_id: DWORD;
    usri2_next_index: DWORD;
  end;

  // Структура для отображения пользователей
  PNetDisplayUser = ^TNetDisplayUser;
  TNetDisplayUser = record
    usri1_name: LPWSTR;
    usri1_comment: LPWSTR;
    usri1_flags: DWORD;
    usri1_full_name: LPWSTR;
    usri1_user_id: DWORD;
    usri1_next_index: DWORD;
  end;

    // Структура для отображения пользователей принадлежащих группе
  // или групп в которые входит пользователь
  PGroupUsersInfo0 = ^TGroupUsersInfo0;
  TGroupUsersInfo0 = record
    grui0_name: LPWSTR;
  end;
  TStringArray = array of string;
////////////////////////////////////////////////
///  структура для получения информации о пользователях на компьютерах

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

     // Функции которые предоставят нам возможность получения информации
//  function NetWkstaUserEnum(ServerName: PWideChar; Level: DWORD;
 //   Bufptr: Pointer; prefmaxlen:DWORD ;entriesread:LPDWORD;
 //   totalentries:LPDWORD; resumehandle:LPDWORD):DWORD; stdcall;
 //    external netapi32lib;
 function NetWkstaUserEnum(servername: PWideChar;
   // Указатель на строку, которая указывает DNS или NetBIOS-имя
  // удаленный сервер, на котором должна выполняться функция.
  // Если этот параметр nil, используется локальный компьютер.
  level: DWORD;
   // Level = 0: вернуть имена пользователей, которые в данный момент вошли на рабочую станцию.
  var bufptr: Pointer;   // Указатель на буфер, который получает данные
  prefmaxlen: DWORD;
   //  Указывает предпочтительную максимальную длину возвращаемых данных в байтах.
  var entriesread: PDWord;
   // Указатель на значение, которое получает количество фактически перечисленных элементов.
  var totalentries: PDWord;  // общее количество записей
  var resumehandle: PDWord)
   // содержит дескриптор резюме, который используется для продолжения существующего поиска
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

   function GetDomainController(const DomainName: String): String; ///функция извлечения имени контроллера домена
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
function GetCurrentComputerName: String; /// получение информации о компьютере в сети и домене
var
  Info: PWkstaInfo100;
  Error: DWORD;
begin
try
  // А для этого мы воспользуемся следующей функцией
  Error := NetWkstaGetInfo(nil, 100, @Info);
  if Error <> 0 then
    raise Exception.Create(SysErrorMessage(Error));
  // Как видно, вызов который возвращает обычную структуру, из которой и прочитаем, все что нужно :)

  // А именно имя компьютера в сети
  Result := Info^.wki100_computername;
  // И где он находиться
  CurrentDomainName := info^.wki100_langroup;
Except
   on E: Exception do
      begin
        //memo1.Lines.add('Ошибка загрузки списка компьютеров  :'+E.Message);
      end;
end;
end;

///////////////////////////////////////////////////////  Версия ОС (5.0,6.0,6.1 и т.д)
function GetCurrentComputerOS(s:PWideChar): String;
var
  Info: PWkstaInfo100;
  Error: DWORD;
  begin
  // А для этого мы воспользуемся следующей функцией
   try
  Error := NetWkstaGetInfo((s), 101, @Info);   /// s имя компьютера
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
  // Используем недокументированную функцию NetEnumerateTrustedDomains
  // (только не пойму, с какого перепуга она не документирована?)
  // Тут все очень просто, на вход имя контролера домена, ны выход - список доверенных доменов
  Result := NetEnumerateTrustedDomains(StringToOleStr(GetDomainController(CurrentDomainName)),
    @DomainList) = NERR_Success;
  // Если вызов функции успешен, то...
  if Result then
  try
    Tmp := DomainList;
    while Length(Tmp) > 0 do
    begin
      //memTrustedDomains.Lines.Add(Tmp); // Банально выводим список на экран
      Tmp := Tmp + Length(Tmp) + 1;
    end;
  finally
    // Не забываем про память
    NetApiBufferFree(DomainList);
  end;
end;


//  Данная функция получает информацию о всех рабочих станциях присутствующих в домене
//  Вообщето так делать немного не верно, дело в том что рабочие станции могут
//  присутствовать в списке не только те, которые завел сисадмин (но для демки сойдет и так)
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
    // NetQueryDisplayInformation возвращает информацию только о 100-а записях
    // для того чтобы получить всю информацию используется третий параметр,
    // передаваемый функции, который определяет с какой записи продолжать
    // вывод информации
    EntriesRequest := 100;
    PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayMachine);
    ReturnedEntryCount := 0;
    // Для выполнения функции, в нее нужно передать DNS имя контролера домена
    // (или его IP адрес), с которого мы хочем получить информацию
    // Для получения информации о рабочих станциях используется структура NetDisplayMachine
    // и ее идентификатор 2 (двойка) во втором параметре
    Error := NetQueryDisplayInformation(StringToOleStr(GetDomainController(CurrentDomainName)), 2, CurrIndex,
      EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
    // При безошибочном выполнении фунции будет результат либо
    // 1. NERR_Success - все записи возвращены
    // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
    if Error in [NERR_Success, ERROR_MORE_DATA] then
    try
      Tmp := Info;
      // Выводим информацию которую вернула функция в структуру
      for I := 0 to ReturnedEntryCount - 1 do
      begin
      result.add(copy(Tmp^.usri2_name,1,Length(Tmp^.usri2_name)-1));
      CurrIndex := Tmp^.usri2_next_index;
      Inc(Tmp);
      end;
    finally
      // Дабы небыло утечек
      NetApiBufferFree(Info);
    end;
  // Если результат выполнения функции ERROR_MORE_DATA
  // (т.е. есть еще данные) - вызываем функцию повторно
  until Error in [NERR_Success, ERROR_ACCESS_DENIED];
  // Ну и возвращаем результат всего что мы тут накодили
 // Result := Error = NERR_Success;
end;

//////////////////////////////////////////////////////////////////////////////
//  Довольно простая функция, возвращает только имена компьютеров принадлезжащих группе
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
  // Обязательная инициализация
  ResumeHandle := 0;
  PrefMaxLen := DWORD(-1);
  // Выполняем
  rebool := NetGroupGetUsers(StringToOleStr(GetDomainController(CurrentDomainName)),
    StringToOleStr(GroupName), 0, Pointer(Info), PrefMaxLen,
    EntriesRead, TotalEntries, @ResumeHandle) = NERR_Success;
  // Смотрим результат...
  if rebool then
  try
    Tmp := Info;
    for I := 0 to EntriesRead - 1 do
    begin
       // Банально выводим результат из структуры
    if pos('$',(Tmp^.grui0_name))<>0 then
     result.Add(copy(Tmp^.grui0_name,1,Length(Tmp^.grui0_name)-1));
    Inc(Tmp);
    end;
  finally
    // Не забываем, ибо может быть склероз :)
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
    // NetQueryDisplayInformation возвращает информацию только о 100-а записях
    // для того чтобы получить всю информацию используется третий параметр,
    // передаваемый функции, который определяет с какой записи продолжать
    // вывод информации
    EntriesRequest := 100;
    PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayGroup);
    ReturnedEntryCount := 0;
    // Для выполнения функции, в нее нужно передать DNS имя контролера домена
    // (или его IP адрес), с которого мы хочем получить информацию
    // Для получения информации о группах используется структура NetDisplayGroup
    // и ее идентификатор 3 (тройка) во втором параметре
    Error := NetQueryDisplayInformation(StringToOleStr(GetDomainController(CurrentDomainName)), 3, CurrIndex,
    EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
    // При безошибочном выполнении фунции будет результат либо
    // 1. NERR_Success - все записи возвращены
    // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
    if Error in [NERR_Success, ERROR_MORE_DATA] then
    try
      Tmp := Info;
      for I := 0 to ReturnedEntryCount - 1 do
      begin
      result.add(Tmp^.grpi3_name);         // Имя группы
      // Запоминаем индекс с которым будем вызывать повторно функцию (если нужно)
      CurrIndex := Tmp^.grpi3_next_index;
      Inc(Tmp);
      end;
      ///  ////////////////////////////////////////////////////
    finally
      // Чтобы небыло утечки ресурсов, освобождаем память занятую функцией под структуру
      NetApiBufferFree(Info);
    end;
  // Если результат выполнения функции ERROR_MORE_DATA - вызываем функцию повторно
  until Error in [NERR_Success, ERROR_ACCESS_DENIED];
  // Ну и возвращаем результат всего что мы тут накодили
  //Result := Error = NERR_Success;
end;
//  Получаем DNS имя контроллера домена
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
  // Для большинства операций нам потребуется IP адрес контроллера домена
  // или его DNS имя, которое мы получим вот так:
  if DsGetDcName(nil, PChar(CurrentDomainName), GUID, nil,
    DS_IS_FLAT_NAME or DS_RETURN_DNS_NAME, DomainControllerInfo) = NERR_Success then
  // Параметры которые мы передаем означают:
  // DS_IS_FLAT_NAME - передаем просто имя домена
  // DS_RETURN_DNS_NAME - ждем получения DNS имени
  try
    Result := DomainControllerInfo^.DomainControllerName; // Результат собсно тут...
  finally
    // Склероз это болезнь, ее нужно лечить...
    NetApiBufferFree(DomainControllerInfo);
  end;
end;

// =============================================================================
//  Ну тут без комментариев - просто получаем имя контроллера домена
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
  EnumAllTrustedDomains; /// список довереных доменов
  EnumAllWorkStation;  /// все рабочие станции в домене
  EnumAllGroups;      ///все групппы безопасности в домене
   GetCurrentComputerName;  ///  информация о компьютере
  EnumAllUsers; /// Все пользователи домена }

end.
