unit EventWindows;

interface
uses System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,
WinApi.windows,Vcl.ComCtrls,vcl.forms,System.StrUtils,DateUtils;

type
  TPropEvent = record
    NamePC       :String[255];
    User         :String[255];
    pass         :String[255];
    FileLog      :string[255];
    SourseLog    :string[255];
    DateTimeFrom :string[255];
    DateTimeTo   :string[255];
    TypeEvent    :integer;
    CounEvent    :integer;
    FilterSourse :boolean;
    FilterType   :boolean;
    FilterDate   :boolean;
    FilterLogFile:boolean;
    ListEvent    :TListView;
  end;

//// описываем функции
 function readlistLogsFileEvent(s,User,Pass:string):TstringList;
 function readlistSourceEvent(s,FileLog,User,Pass:string):TstringList;
 function readlistViewEvent(ParamTh:pointer):Integer; //// чтение потоком
 {function FuncReadlistViewEvent(s,FileLog,SourseLog,DateTimeFrom,DateTimeTo:string;TypeEvent,CounEvent:integer
;FilterSourse,FilterType,FilterDate:bool;ListEvent:TListView):Boolean;}
 function readInfoSelEvent(s,FileLog,NumberRecord,User,Pass:string):boolean;
 function LogsFileProperties(s,LogFile,User,Pass:string):TstringList;
 function BackUpLogsFileEvent(s,LogFile,Pach,User,Pass:string):integer;
 function ClearLogsFileEvent(s,LogFile,User,Pass:string):integer;
/// константы
///

var
EventCall: TPropEvent; ///


implementation
uses umain,MyDM;

ThreadVar
PropertyEventCall: ^TPropEvent;

var
TimeZoneUTC:integer;

function UTCTime(FWMIService:Olevariant):integer;
var     ////////////////////////////// функция определения смещения времени по UTC
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  i:integer;
begin
try
OleInitialize(nil);
i:=0;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT bias FROM Win32_TimeZone','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
if VarIsNumeric(FWbemObject.bias)then
result:= integer(FWbemObject.bias) div 60
else result:=0;
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    result:=0;
    //frmDomainInfo.memo1.Lines.Add(' - Ошибка UTC "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;


function WmiDate2DT (S: string; var UtcOffset: integer): TDateTime ;
// yyyymmddhhnnss.zzzzzzsUUU  +60 means 60 mins of UTC time
// 20030709091030.686000+060
// 1234567890123456789012345
var
    yy, mm, dd, hh, nn, ss, zz: integer ;
    timeDT: TDateTime ;

    function GetNum (offset, len: integer): integer ;
    var
        E: Integer;
    begin
        Val (copy (S, offset, len), result, E) ;
    end ;

begin
    result := 0 ;
    UtcOffset := 0 ;
    if length (S) <> 25 then exit ;   // fixed length
    yy := GetNum (1, 4) ;
    mm := GetNum (5, 2) ;
    if (mm = 0) or (mm > 12) then exit ;
    dd := GetNum (7, 2) ;
    if (dd = 0) or (dd > 31) then exit ;
    if NOT TryEncodeDate (yy, mm, dd, result) then     // D6 and later
    begin
        result := -1 ;
        exit ;
    end ;
    hh := GetNum (9, 2) ;
    nn := GetNum (11, 2) ;
    ss := GetNum (13, 2) ;
    zz := 0 ;
    if Length (S) >= 18 then zz := GetNum (16, 3) ;
    if NOT TryEncodeTime (hh, nn, ss, zz, timeDT) then exit ;   // D6 and later
    result := result + timeDT ;
    result:=IncHour(Result,TimeZoneUTC); //// увеличиваем врема на определенное количество часов по UTC
    UtcOffset := GetNum (22, 4) ; // including sign
end ;


function createTempTabl(NamePC:string):boolean;
begin
try
if datam.ConnectionDB.Connected then /// если соединение установленно
begin
dataM.FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Clear;
dataM.FDScriptTempTabl.SQLScripts[0].Name:='CTT';  //;
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE GLOBAL TEMPORARY TABLE EVENT (');   //GLOBAL TEMPORARY
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('ID_EVENT  INTEGER,'); //ID_EVENT  INTEGER NOT NULL PRIMARY KEY
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('PC_NAME     VARCHAR(50) ,');    // имя компа из combobox
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('PC_NAME_LOG     VARCHAR(100) ,');
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('LOG_FILE     VARCHAR(100) ,');   // сам лог файл
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('LOG_SOURCE     VARCHAR(150) ,');  // Источник события
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('RECORD_NUMBER    VARCHAR(50) ,'); /// номер записи
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('TYPE_INT     VARCHAR(5) ,'); /// 0-5
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('EVENT_TYPE     VARCHAR(100) ,');     // описание типа события
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('TIME_GEN    VARCHAR(100) ,');
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('MESSAGE_EVENT     BLOB SUB_TYPE 1 SEGMENT SIZE 80,');
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('INSERT_STR     BLOB SUB_TYPE 1 SEGMENT SIZE 80,');
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('EVENT_CODE     VARCHAR(50) ,');
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('CUR_USER  VARCHAR(100)) ON COMMIT PRESERVE ROWS ;'); //  ON COMMIT PRESERVE ROWS
//dataM.FDScriptTempTabl.ValidateAll;
if dataM.FDScriptTempTabl.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  dataM.FDScriptTempTabl.SQLScripts[0].SQL.Clear;
  dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DELETE FROM EVENT;');
  dataM.FDScriptTempTabl.ExecuteAll;
  end;
end;
except
  on E:Exception do
    begin
    result:=false;
    frmDomainInfo.memo1.Lines.Add('Ошибка создания  таблицы EVENT "'+E.Message+'"');
    OleUnInitialize;
    exit;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////////////////////

function ClearLogsFileEvent(s,LogFile,user,pass:string):integer;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Filename,Name FROM Win32_NTEventlogFile WHERE Filename="'+LogFile+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
 result:=FWbemObject.ClearEventLog(); //string(FWbemObject.name)
 FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
frmDomainInfo.memo1.Lines.Add('Операция очистка журнала: '+SysErrorMessage(result));
except
  on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка операции очистка журнала "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;


function BackUpLogsFileEvent(s,LogFile,Pach,User,Pass:string):integer;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Filename FROM Win32_NTEventlogFile WHERE Filename="'+LogFile+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
 result:=FWbemObject.BackupEventlog(pach);
 FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
frmDomainInfo.memo1.Lines.Add('Операция BackUpLog: '+SysErrorMessage(result));
except
  on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка операции BackUpLog "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

function LogsFileProperties(s,LogFile,user,Pass:string):TstringList;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  utc:integer;
  sourceSTR:string;
  i:integer;
begin
try
OleInitialize(nil);
result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
//if LogFile='Security' then  //// добавляем привелегию на чтение журнала безопасности
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);

TimeZoneUTC:=0;
TimeZoneUTC:=UTCTime(FWMIService); //// определяем смещение времени по UTC
///
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTEventlogFile WHERE FileName ="'+LogFile+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
if (FWbemObject.Archive<>null)then result.Add('Архивный: '+string(FWbemObject.Archive));
if (FWbemObject.Caption<>null)then result.Add('Путь: '+string(FWbemObject.Caption));
if (FWbemObject.Compressed<>null)then
if (FWbemObject.Compressed)then
  begin
  result.Add('Compressed: '+string(FWbemObject.Compressed));
  if (FWbemObject.CompressionMethod<>null)then
  result.Add('CompressionMethod: '+string(FWbemObject.CompressionMethod));
  end;
if (FWbemObject.CreationDate<>null)then result.Add('Дата/время создания: '+datetimetostr(wmidate2dt(string(FWbemObject.CreationDate),utc)));
if (FWbemObject.LastModified<>null)then result.Add('Дата/время изменения: '+datetimetostr(wmidate2dt(string(FWbemObject.LastModified),utc)));
if (FWbemObject.Encrypted<>null)then
if (FWbemObject.Encrypted)then
  begin
  result.Add('Зашифрован: '+string(FWbemObject.Encrypted));
  if (FWbemObject.EncryptionMethod<>null)then
  result.Add('Метод шифрования: '+string(FWbemObject.EncryptionMethod));
  end;
if (FWbemObject.FileSize<>null) then
result.Add('Размер файла: '+string(FWbemObject.FileSize div 1024)+' Кб');
if (FWbemObject.MaxFileSize<>null)then
 result.Add('Максимальный размер файла: '+string(FWbemObject.MaxFileSize div 1024)+' Кб');
if (FWbemObject.NumberOfRecords<>null)then result.Add('Количество записей: '+string(FWbemObject.NumberOfRecords));
if (FWbemObject.Readable<>null)then result.Add('Чтение: '+string(FWbemObject.Readable));
if (FWbemObject.Writeable<>null)then result.Add('Запись: '+string(FWbemObject.Writeable));
if (FWbemObject.status<>null)then result.Add('Статус: '+string(FWbemObject.status));
if (FWbemObject.Version<>null)then result.Add('Версия: '+string(FWbemObject.Version));
if (FWbemObject.System<>null)then result.Add('Системный: '+string(FWbemObject.system));
if (FWbemObject.Hidden<>null)then result.Add('Скрытый: '+string(FWbemObject.Hidden));
if VarIsArray(FWbemObject.Sources) then
begin
sourceSTR:='';
for I := VarArrayLowBound(FWbemObject.Sources,1) to VarArrayHighBound(FWbemObject.Sources,1) do
  sourceSTR:=sourceSTR+string(FWbemObject.properties_.Item('Sources',0).value[i])+', ';
  Result.Add('Источники событий: '+sourceSTR);
end;
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    result.Add('Ошибка');
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка чтения свойств журнала "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

function readlistLogsFileEvent(s,user,pass:string):TstringList;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Filename FROM Win32_NTEventlogFile','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
begin
if (FWbemObject.FileName<>null)then result.Add(string(FWbemObject.FileName));
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    result.Add('Ошибка');
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка чтения списка журналов и источников "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function readlistSourceEvent(s,FileLog,user,Pass:string):TstringList;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  i:integer;
begin
try
OleInitialize(nil);
result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT FileName,Sources FROM Win32_NTEventlogFile WHERE FileName="'+FileLog+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
if VarIsArray(FWbemObject.Sources)then
  begin
  for I := VarArrayLowBound(FWbemObject.Sources,1) to VarArrayHighBound(FWbemObject.Sources,1) do
  Result.Add(FWbemObject.properties_.Item('Sources',0).value[i]);
  end;
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    result.Add('Ошибка');
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка чтения списка источника событий "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
{function FuncReadlistViewEvent(s,FileLog,SourseLog,DateTimeFrom,DateTimeTo:string;TypeEvent,CounEvent:integer
;FilterSourse,FilterType,FilterDate:bool;ListEvent:TListView):Boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  i,z,utc:integer;
  step:string;
  pc_name,log_file,log_source,record_number,event_type,type_int,time_gen,message_event,insert_str,event_code,cur_user:string;
  const
  TypeEv: array[0..5] of string=('Сведения','Ошибка','Предупреждение','Информация','Аудит успеха','Аудит отказа ');
begin
try
OleInitialize(nil);
z:=0;
ListEvent.SmallImages:= frmDomainInfo.ImageListEventWin;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', MyUser, MyPasswd);
//FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', MyUser, MyPasswd,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
///////////////////////////////////////////////////// создаем таблицу
 createTempTabl(s); /// создаем таблицу

/////////////////////////////////////////////////////////
if (FilterSourse) and (not FilterType) and ( not FilterDate) then /// если фильтруем по источнику событий и имени лога
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND SourceName="'+SourseLog+'"','WQL',wbemFlagForwardOnly);
///////////////////////////////////////////////////////////
if (not FilterSourse) and (not FilterType) and ( not FilterDate) then /// если фильтруем только по имени лог файла
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'"','WQL',wbemFlagForwardOnly);
/////////////////////////////////////////////////////////
if (FilterSourse) and (FilterType) and ( not FilterDate) then /// если фильтруем по источнику событий и имени лога и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND SourceName="'+SourseLog+'" AND EventType="'+inttostr(TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////
if (Not FilterSourse) and (FilterType) and ( not FilterDate) then /// еесли фильтруем только по имени лог файла и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND EventType="'+inttostr(TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////

if (not FilterSourse) and (not FilterType) and (FilterDate) then /// если фильтруем только по имени лог файла и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND TimeGenerated>"'+DateTimeFrom+'.000000-000" AND TimeGenerated<"'+DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
if (FilterSourse) and (not FilterType) and (FilterDate) then /// если фильтруем только по имени лог файла, источнику событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND SourceName="'+SourseLog+'" AND TimeGenerated>"'+DateTimeFrom+'.000000-000" AND TimeGenerated<"'+DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
if (not FilterSourse) and (FilterType) and (FilterDate) then /// если фильтруем только по имени лог файла, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND EventType="'+inttostr(TypeEvent)+'" AND TimeGenerated>"'+DateTimeFrom+'.000000-000" AND TimeGenerated<"'+DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
if (FilterSourse) and (FilterType) and (FilterDate) then /// если фильтруем только по имени лог файла, источнику событий, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND EventType="'+inttostr(TypeEvent)+'" AND SourceName="'+SourseLog+'" AND TimeGenerated>"'+DateTimeFrom+'.000000-000" AND TimeGenerated<"'+DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////


oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
frmDomainInfo.GroupBoxEventProperties.Caption:='';
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
begin
 if z>=CounEvent then break;
 log_file:=FileLog;
with ListEvent.Items.Add do
  BEGIN
  if (FWbemObject.EventType<>null) and VarIsNumeric(FWbemObject.EventType) then
  Begin
  inc(z);
   if (FWbemObject.RecordNumber<>null) and VarIsNumeric(FWbemObject.RecordNumber) then
     begin
     caption:=string(FWbemObject.RecordNumber);
     ImageIndex:=integer(FWbemObject.EventType);
     SubItems.Add(TypeEv[strtoint(FWbemObject.EventType)]);
     record_number:=string(FWbemObject.RecordNumber);
     event_type:=TypeEv[strtoint(FWbemObject.EventType)];
     type_int:=string(FWbemObject.EventType);
     end
      else begin Caption:=' '; SubItems.Add(' '); record_number:='no';event_type:='no';type_int:='no';  end;
    step:='0';
  if (FWbemObject.TimeGenerated<>null) then
    begin
    time_gen:=DateTimeToStr(WmiDate2DT(string(FWbemObject.TimeGenerated),utc));
    SubItems.Add(time_gen);
    end
    else begin SubItems.add(' '); time_gen:='no'; end;
    step:='1';
  if (FWbemObject.SourceName<>null) then
    begin
    SubItems.Add(string(FWbemObject.SourceName));
    log_source:=string(FWbemObject.SourceName);
    end
    else begin  SubItems.add(' ');log_source:='no'; end;
     step:='2';
  if (FWbemObject.EventCode<>null) then
    begin
    SubItems.Add(string(FWbemObject.EventCode));
    event_code:=string(FWbemObject.EventCode);
    end
    else begin SubItems.add(' '); event_code:='no';  end;
    step:='3';
  if (FWbemObject.ComputerName<>null) then pc_name:=string(FWbemObject.ComputerName)
  else pc_name:='no';
     step:='4';
  if (FWbemObject.Message<>null) then
    begin
     if string(FWbemObject.Message)<>'' then message_event:=trim(string(FWbemObject.Message))
     else message_event:='no';
    end
   else message_event:='no';
      step:='5';
  if VarIsArray(FWbemObject.InsertionStrings) then
    begin
    insert_str:='';
    for I := VarArrayLowBound(FWbemObject.InsertionStrings,1) to VarArrayHighBound(FWbemObject.InsertionStrings,1) do
    begin
    if FWbemObject.properties_.Item('InsertionStrings',0).value[i]<>'' then
        insert_str:=insert_str+(FWbemObject.properties_.Item('InsertionStrings',0).value[i])+#10;
    end;
    insert_str:=trim(insert_str);
    end
   else insert_str:='no';
       step:='6';
  if (FWbemObject.user<>null) then cur_user:=string(FWbemObject.user)
    else cur_user:='no';
       step:='7';
DataM.FDQueryTempTabl.SQL.clear;
datam.FDQueryTempTabl.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
+'EVENT (PC_NAME,PC_NAME_LOG,LOG_FILE,LOG_SOURCE,RECORD_NUMBER,TYPE_INT,EVENT_TYPE,TIME_GEN ,MESSAGE_EVENT,INSERT_STR,EVENT_CODE,CUR_USER)'
+' VALUES(:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12) MATCHING (PC_NAME,LOG_FILE,RECORD_NUMBER)';
datam.FDQueryTempTabl.params.ParamByName('p1').AsString:=''+leftstr(s,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p2').AsString:=''+leftstr(pc_name,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p3').AsString:=''+leftstr(log_file,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p4').AsString:=''+leftstr(log_source,150)+'';
datam.FDQueryTempTabl.params.ParamByName('p5').AsString:=''+leftstr(record_number,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p6').AsString:=''+leftstr(type_int,5)+'';
datam.FDQueryTempTabl.params.ParamByName('p7').AsString:=''+leftstr(event_type,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p8').AsString:=''+leftstr(time_gen,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p9').AsString:=''+leftstr(message_event,16381)+'';
datam.FDQueryTempTabl.params.ParamByName('p10').AsString:=''+leftstr(insert_str,16381)+'';
datam.FDQueryTempTabl.params.ParamByName('p11').AsString:=''+leftstr(event_code,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p12').AsString:=''+leftstr(cur_user,100)+'';
 step:='8';
datam.FDQueryTempTabl.ExecSQL;
step:='9';
  End;
END;
step:='10';
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;

except
  on E:Exception do
    begin
    //frmDomainInfo.memo1.Lines.Add('z - '+inttostr(z)+'/ Шаг - '+step+' / '+s+' - Ошибка чтения списка событий "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка чтения списка событий "'+E.Message+'"');
    frmDomainInfo.GroupBoxEventProperties.Caption:='';
    OleUnInitialize;
    end;
  end;
end;}
//////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function readlistViewEvent(ParamTh:pointer):Integer;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  i,z,utc:integer;
  step:string;
  pc_name,log_file,log_source,record_number,event_type,type_int
  ,time_gen,message_event,insert_str,event_code,cur_user:string;

  const
  TypeEv: array[0..5] of string=('Сведения','Ошибка','Предупреждение','Информация','Аудит успеха','Аудит отказа ');
begin
try
step:='-10';
result:=0;
PropertyEventCall:=ParamTh;
OleInitialize(nil);
z:=0;
step:='-9';
utc:=5;
{with frmDomainInfo.memo1.Lines do
begin
  add(PropertyEventCall.NamePC);
  add(PropertyEventCall.FileLog);
  add(PropertyEventCall.SourseLog);
  add(PropertyEventCall.DateTimeFrom);
  add(PropertyEventCall.DateTimeTo);
end; }

PropertyEventCall.ListEvent.SmallImages:= frmDomainInfo.ImageListEventWin;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(PropertyEventCall.NamePC, 'root\CIMV2', PropertyEventCall.user, PropertyEventCall.Pass);  //// ,'','',128 - ждем не более 2х минут
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
//if PropertyEventCall.FileLog='Security' then  //// добавляем привелегию на чтение журнала безопасности
begin
//FWMIService.Security_.Privileges.AddAsString('SeCreateTokenPrivilege',true); // Требуется для создания основного объекта токена.
//FWMIService.Security_.Privileges.AddAsString('SeAssignPrimaryTokenPrivilege',true);  // Требуется для замены токена уровня процесса.
//FWMIService.Security_.Privileges.AddAsString('SeLockMemoryPrivilege',true);         //Требуется для блокировки страниц в памяти.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseQuotaPrivilege',true);     //Требуется настроить квоты памяти для процесса.
//FWMIService.Security_.Privileges.AddAsString('SeMachineAccountPrivilege',true);   // Требуется для добавления рабочих станций в домен.
//FWMIService.Security_.Privileges.AddAsString('SeTcbPrivilege',true);             //Требуется действовать как часть операционной системы. Держатель является частью надежной компьютерной базы.
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);        //Требуется для управления аудитом и журналом безопасности NT.
//FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);  //Требуется принять права собственности на файлы или другие объекты, не имея записи контроля доступа (ACE) в списке контроля доступа по усмотрению (DACL).
//FWMIService.Security_.Privileges.AddAsString('SeLoadDriverPrivilege',true);     //Требуется для загрузки или выгрузки драйвера устройства.
//FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);     //Требуется для сбора профильной информации о производительности системы.
//FWMIService.Security_.Privileges.AddAsString('SeSystemtimePrivilege',true);      // Требуется изменить системное время.
//FWMIService.Security_.Privileges.AddAsString('SeProfileSingleProcessPrivilege',true); //Требуется для сбора информации профиля для одного процесса.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseBasePriorityPrivilege',true);  // Требуется для увеличения приоритета планирования
//FWMIService.Security_.Privileges.AddAsString('SeCreatePagefilePrivilege',true);       // Требуется для создания файла подкачки.
//FWMIService.Security_.Privileges.AddAsString('SeCreatePermanentPrivilege',true);      //Требуется для создания постоянных общих объектов.
//FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege',true);               //Требуется для резервного копирования файлов и каталогов, независимо от списка ACL, указанного для файла.
//FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',true);            //Требуется для восстановления файлов и каталогов независимо от списка ACL, указанного для файла.
//FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   Требуется выключить локальную систему
//FWMIService.Security_.Privileges.AddAsString('SeDebugPrivilege',true);             //Требуется для отладки и настройки памяти процесса, принадлежащего другой учетной записи.
//FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);             //Требуется для создания записей аудита в журнале безопасности NT. Только защищенные серверы должны иметь эту привилегию.
//FWMIService.Security_.Privileges.AddAsString('SeSystemEnvironmentPrivilege',true); // Требуется для изменения энергонезависимой оперативной памяти систем, которые используют этот тип памяти для хранения данных конфигурации.
//FWMIService.Security_.Privileges.AddAsString('SeChangeNotifyPrivilege',true);     //  Требуется для получения уведомлений об изменениях файлов или каталогов и обхода проверок доступа. Эта привилегия включена по умолчанию для всех пользователей.
//FWMIService.Security_.Privileges.AddAsString('SeRemoteShutdownPrivilege',true);   //Требуется выключить удаленный компьютер.
//FWMIService.Security_.Privileges.AddAsString(' SeUndockPrivilege',true);           //Требуется снять ноутбук с док-станции.
//FWMIService.Security_.Privileges.AddAsString('SeSyncAgentPrivilege',true);         // Требуется для синхронизации данных службы каталогов.
//FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);   // Требуется, чтобы учетные записи компьютеров и пользователей были доверенными для делегирования.
//FWMIService.Security_.Privileges.AddAsString('SeManageVolumePrivilege',true);       //Требуется для выполнения задач объемного обслуживания.
end;
//FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege');
///////////////////////////////////////////////////// создаем таблицу если её не нашли " "  при поиске таблице нужны только если таблица временна
if not DataM.TableExists('"EVENT"') then   createTempTabl(PropertyEventCall.NamePC); /// создаем таблицу
step:='-8';
TimeZoneUTC:=0;
TimeZoneUTC:=UTCTime(FWMIService); //// определяем смещение времени по UTC
/////////////////////////////////////////////////////////
if PropertyEventCall.FilterLogFile then  //// если фильтруем по лог файлу
begin
if (PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем по источнику событий и имени лога
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND SourceName="'+PropertyEventCall.SourseLog+'"','WQL',wbemFlagForwardOnly);
///////////////////////////////////////////////////////////
step:='-7';
if (not PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'"','WQL',wbemFlagForwardOnly);
/////////////////////////////////////////////////////////
step:='-6';
if (PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем по источнику событий и имени лога и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND SourceName="'+PropertyEventCall.SourseLog+'" AND EventType="'+inttostr(PropertyEventCall.TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////
step:='-5';
if (Not PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// еесли фильтруем только по имени лог файла и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND EventType="'+inttostr(PropertyEventCall.TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////
step:='-4';
if (not PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-3';
if (PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, источнику событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND SourceName="'+PropertyEventCall.SourseLog+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-2';
if (not PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND EventType="'+inttostr(PropertyEventCall.TypeEvent)+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-1';
if (PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, источнику событий, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE LogFile="'+PropertyEventCall.FileLog+'" AND EventType="'+inttostr(PropertyEventCall.TypeEvent)+'" AND SourceName="'+PropertyEventCall.SourseLog+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
end
  else //// если фильтрация без учета логфайла
begin
if (PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем по источнику событий и имени лога
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE SourceName="'+PropertyEventCall.SourseLog+'"','WQL',wbemFlagForwardOnly);
///////////////////////////////////////////////////////////
step:='-7';
if (not PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent ','WQL',wbemFlagForwardOnly);
/////////////////////////////////////////////////////////
step:='-6';
if (PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// если фильтруем по источнику событий и имени лога и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE SourceName="'+PropertyEventCall.SourseLog+'" AND EventType="'+inttostr(PropertyEventCall.TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////
step:='-5';
if (Not PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and ( not PropertyEventCall.FilterDate) then /// еесли фильтруем только по имени лог файла и типу событий -0,,5
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE EventType="'+inttostr(PropertyEventCall.TypeEvent)+'"','WQL',wbemFlagForwardOnly);
////////////////////////////////////////////////////
step:='-4';
if (not PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-3';
if (PropertyEventCall.FilterSourse) and (not PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, источнику событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE SourceName="'+PropertyEventCall.SourseLog+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-2';
if (not PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE EventType="'+inttostr(PropertyEventCall.TypeEvent)+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////
step:='-1';
if (PropertyEventCall.FilterSourse) and (PropertyEventCall.FilterType) and (PropertyEventCall.FilterDate) then /// если фильтруем только по имени лог файла, источнику событий, типу событий и дате
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NTLogEvent WHERE EventType="'+inttostr(PropertyEventCall.TypeEvent)+'" AND SourceName="'+PropertyEventCall.SourseLog+'" AND TimeGenerated>"'+PropertyEventCall.DateTimeFrom+'.000000-000" AND TimeGenerated<"'+PropertyEventCall.DateTimeTo+'.000000-000"','WQL',wbemFlagForwardOnly);
//////////////////////////////////////////////////////////////////////////

end;

step:='-0';
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
frmDomainInfo.GroupBoxEventProperties.Caption:='';
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
begin
 if z>=PropertyEventCall.CounEvent then break;
with PropertyEventCall.ListEvent.Items.Add do
  BEGIN
  if (FWbemObject.EventType<>null) and VarIsNumeric(FWbemObject.EventType) then
  Begin
  inc(z);
   if (FWbemObject.RecordNumber<>null) and VarIsNumeric(FWbemObject.RecordNumber) then
     begin
     caption:=string(FWbemObject.RecordNumber);
     ImageIndex:=integer(FWbemObject.EventType);
     SubItems.Add(TypeEv[strtoint(FWbemObject.EventType)]);
     record_number:=string(FWbemObject.RecordNumber);
     event_type:=TypeEv[strtoint(FWbemObject.EventType)];
     type_int:=string(FWbemObject.EventType);
     end
      else begin Caption:=' '; SubItems.Add(' '); record_number:='no';event_type:='no';type_int:='no';  end;
    step:='0';
  if (FWbemObject.TimeGenerated<>null) then
    begin
    time_gen:=DateTimeToStr(WmiDate2DT(string(FWbemObject.TimeGenerated),utc));
    SubItems.Add(time_gen);
    end
    else begin SubItems.add(' '); time_gen:='no'; end;
    step:='1';
  if (FWbemObject.SourceName<>null) then
    begin
    log_source:=string(FWbemObject.SourceName);
    SubItems.Add(log_source);
    end
    else begin  SubItems.add(' ');log_source:='no'; end;
     step:='2';
  if (FWbemObject.EventCode<>null) then
    begin
    event_code:=string(FWbemObject.EventCode);
    SubItems.Add(event_code);
    end
    else begin SubItems.add(' '); event_code:='no';  end;
    step:='3';
   if (FWbemObject.LogFile<>null) then
    begin
    log_file:=string(FWbemObject.LogFile);
    SubItems.Add(log_file);
    end
    else begin SubItems.add(' '); log_file:=' ';  end;
    step:='3.5';
  //
  if (FWbemObject.ComputerName<>null) then pc_name:=string(FWbemObject.ComputerName)
  else pc_name:='no';
     step:='4';
  if (FWbemObject.Message<>null) then
    begin
     if string(FWbemObject.Message)<>'' then message_event:=trim(string(FWbemObject.Message))
     else message_event:='no';
    end
   else message_event:='no';
      step:='5';
  if VarIsArray(FWbemObject.InsertionStrings) then
    begin
    insert_str:='';
    for I := VarArrayLowBound(FWbemObject.InsertionStrings,1) to VarArrayHighBound(FWbemObject.InsertionStrings,1) do
    begin
    if FWbemObject.properties_.Item('InsertionStrings',0).value[i]<>'' then
        insert_str:=insert_str+(FWbemObject.properties_.Item('InsertionStrings',0).value[i])+#10;
    end;
    insert_str:=trim(insert_str);
    end
   else insert_str:='no';
       step:='6';
  if (FWbemObject.user<>null) then cur_user:=string(FWbemObject.user)
    else cur_user:='no';
       step:='7';
DataM.FDQueryTempTabl.SQL.clear;
datam.FDQueryTempTabl.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
+'EVENT (PC_NAME,PC_NAME_LOG,LOG_FILE,LOG_SOURCE,RECORD_NUMBER,TYPE_INT,EVENT_TYPE,TIME_GEN ,MESSAGE_EVENT,INSERT_STR,EVENT_CODE,CUR_USER)'
+' VALUES(:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12) MATCHING (PC_NAME,LOG_FILE,RECORD_NUMBER)';
datam.FDQueryTempTabl.params.ParamByName('p1').AsString:=''+leftstr(PropertyEventCall.NamePC,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p2').AsString:=''+leftstr(pc_name,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p3').AsString:=''+leftstr(log_file,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p4').AsString:=''+leftstr(log_source,150)+'';
datam.FDQueryTempTabl.params.ParamByName('p5').AsString:=''+leftstr(record_number,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p6').AsString:=''+leftstr(type_int,5)+'';
datam.FDQueryTempTabl.params.ParamByName('p7').AsString:=''+leftstr(event_type,100)+'';
datam.FDQueryTempTabl.params.ParamByName('p8').AsString:=''+leftstr(time_gen,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p9').AsString:=''+leftstr(message_event,16381)+'';
datam.FDQueryTempTabl.params.ParamByName('p10').AsString:=''+leftstr(insert_str,16381)+'';
datam.FDQueryTempTabl.params.ParamByName('p11').AsString:=''+leftstr(event_code,50)+'';
datam.FDQueryTempTabl.params.ParamByName('p12').AsString:=''+leftstr(cur_user,100)+'';
 step:='8';
datam.FDQueryTempTabl.ExecSQL;
step:='9';
  End;
END;
step:='10';
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
frmDomainInfo.SpeedButton82.Enabled:=true;
except
  on E:Exception do
    begin
    ///frmDomainInfo.memo1.Lines.Add('z - '+inttostr(z)+'/ Шаг  '+step+' / '+PropertyEventCall.NamePC+' - Ошибка чтения списка событий "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add(PropertyEventCall.NamePC+' - Ошибка чтения списка событий "'+E.Message+'"');
    frmDomainInfo.GroupBoxEventProperties.Caption:='';
    frmDomainInfo.SpeedButton82.Enabled:=true;
    OleUnInitialize;
    end;
  end;
NetApiBufferFree(ParamTh); /// очищаем память
EndThread(result); /// конец потока
end;
//////////////////////////////////////////////////////////////////////////////////////////////
function readInfoSelEvent(s,FileLog,NumberRecord,user,pass:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);
FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Message,RecordNumber,LogFile,InsertionStrings FROM Win32_NTLogEvent WHERE LogFile="'+FileLog+'" AND RecordNumber="'+NumberRecord+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin
if FWbemObject.Message<>null then frmDomainInfo.MemoEventInfo.Text:= string(FWbemObject.Message);
if FWbemObject.InsertionStrings<>null then
begin
frmDomainInfo.MemoEventInfo.Lines.Add('--------------------------------------------------------------');
frmDomainInfo.MemoEventInfo.Lines.Add( string(FWbemObject.InsertionStrings));
end;

FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
result:=true;
except
  on E:Exception do
    begin
    result:=false;
    frmDomainInfo.memo1.Lines.Add(s+' - Ошибка чтения списка журналов и источников "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;
end.
