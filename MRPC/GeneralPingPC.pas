 unit GeneralPingPC;

interface

uses
 System.Classes,Windows,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,
 SysUtils,IdIcmpClient,IdTCPClient,WinSock,Forms,IniFiles,FireDAC.Phys.FB,
 FireDAC.Comp.Client
 ,FireDAC.Stan.Option,System.StrUtils;

 const  netapi32lib = 'netapi32.dll';

type
  THostInfo = record
     username     : PWideChar;
     logon_domain : PWideChar;
     other_domains: PWideChar;
     logon_server : PWideChar;
   end;

 GeneralPing = class(TThread)
  private
  function Log_write(Level:integer;fname,text:string):string;
  function getUrdmClient(s:string):bool;
  function get135Port(s:string):bool;
  function PcInLan(PC:string;availiblePC:bool):bool;
  function AntivirusUpdateTable(NamePC,AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate,
  AntySpw,AntySpwStatus,AntySpwUpdate,OsName:string):boolean;
  function readUsWinDomIsWmi(s:string):bool;
  function GetInfoComputerInAD(pcName:string): Bool;
  function OsherInfoPC(s:string;Item:integer):bool;
  function ConnectiontDB:boolean;
  function DisconnectDB:boolean;
  procedure ping;
  function EnumNetUsers(HostName: WideString): THostInfo;
  function MySendARP(const MyIPAddress: String): String;
  function WriteAllResult(s:string;Item:integer;StatusLanPC:string):boolean; // запись результатов если данные о компьютере получены
  function EslePcNotLan(ImPing,ImuRdm:integer;NamePC:string):boolean; // если компьютер не в сети
  function SelectedPing(var resIP:string;HostName:string;Timeout:integer):boolean; // Проверка доступности узла 3мя способами на выбор

  protected
procedure Execute; override;
end;


function SendARP(DestIp: DWORD; srcIP: DWORD; pMacAddr: pointer; PhyAddrLen: Pointer): DWORD;stdcall; external 'iphlpapi.dll';

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
 function NetWkstaGetInfo(ServerName: PWideChar; Level: DWORD;
    Bufptr: Pointer): DWORD; stdcall; external netapi32lib;

implementation

uses umain,PingForScan;

 ThReadVar
InvUser,InvPass,nameOS,UserNamePC,PCDomain,LogonSrv,AnswerMAC,IpADRS,ClientuRDM:string;
AntivirProduct:string;
AntivirusIdStatus:integer;
listAvaliblePC      :Tstringlist;
ConnectionThreadPing: TFDConnection;
RenderingPing:boolean; // обновлять или нет listview при блокировке пользователя
TransactionWrite    : TFDTransaction;
FDQueryWrite        : TFDQuery;

////////////////////////////////////////////////////////////
function GeneralPing.SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
begin
resIP:='';
  case PingType of
    1: result:=PingIdIcmp(resIP,HostName,pingtimeout);
    2: result:=PingGetaddrinfo(resIP,HostName,pingtimeout);
    3: result:=PingGetHostByName(resIP,HostName,pingtimeout);
  END;
end;
///////////////////////////////////////////////////////////

function GeneralPing.Log_write(Level:integer;fname,text:string):string;
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
//////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////
function GeneralPing.getUrdmClient(s:string):bool;
var
uRDMscanTCP: TIdTCPClient;
begin
try
uRDMscanTCP:=TIdTCPClient.Create;
uRDMscanTCP.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
uRDMscanTCP.Port:=strtoint(urdmport);
uRDMscanTCP.ConnectTimeout:=pingtimeout div 4;
try
uRDMscanTCP.Port:=strtoint(urdmport);
uRDMscanTCP.Host:=s;
uRDMscanTCP.Connect;
if uRDMscanTCP.Connected then result:=true
else result:=false;
finally
uRDMscanTCP.Disconnect;
freeAndNil(uRDMscanTCP);
end;
 Except
 begin
 result:=false;
 //Log_write(0,'Scan',s+': Scan uRDM port Connect timed out.');
 end;
 end;
end;
//////////////////////////////////////////////////////
function GeneralPing.get135Port(s:string):bool;
var
RPCPortScan : TIdTCPClient;
begin
try
////////////////////////////////////////////////RPCPortScan
RPCPortScan:=TIdTCPClient.Create;
RPCPortScan.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
RPCPortScan.Port:=135;   /// проверяем  135 порт
RPCPortScan.ConnectTimeout:=pingtimeout div 4;
RPCPortScan.Host:=s;
RPCPortScan.Connect;
try
if RPCPortScan.Connected then result:=true
else result:=false;
finally
RPCPortScan.Disconnect;
FreeAndNil(RPCPortScan);
end;
 Except
 begin
 result:=false;
 Log_write(0,'Scan',s+': Scan port 135 Connect timed out.');
 end;
 end;
end;
/////////////////////////////////////////////////////////////////////
function GeneralPing.PcInLan(PC:string;availiblePC:bool):bool;
var
i:integer;
FoundPC:bool;
begin
try
FoundPC:=false;
  case availiblePC of
  true:           //// компьютер доспупен
       begin
         for I := 0 to listAvaliblePC.Count-1 do
         begin
          if PC=listAvaliblePC[i] then
            begin
            FoundPC:=true; /// нашли комп
            result:=true;
            break;
            end;
         end;
         if not FoundPC then
           begin
            listAvaliblePC.Add(pc); /// если не нашли комп то добавляем его в список
            result:=false;
           end;
       end;

  false:          /// компьютер не доступен
        for I := 0 to listAvaliblePC.Count-1 do
        begin
          if pc=listAvaliblePC[i] then /// если нашли комп
          begin
          listAvaliblePC.Delete(i);     /// удаляем его нахер
          result:=true;
          break;
          end
          else result:=false;
        end;
  end;
listAvaliblePC.Sort;  /// сортируем список
frmDomainInfo.StatusBar1.Panels[1].Text:='ПК в сети: '+inttostr(listAvaliblePC.Count);

except on E: Exception do
begin
  result:=false;
 Log_write(2,'Scan',PC+': ListAvaliblePC - '+e.Message);
end;
end;
end;

///////////////////////////////////////////////////////////////////
function GeneralPing.AntivirusUpdateTable(NamePC,AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate,
AntySpw,AntySpwStatus,AntySpwUpdate,OsName:string):boolean;
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.Close;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into ANTIVIRUSPRODUCT '+
  ' (NAMEPC,WIN_DEF,WIN_DEF_STATUS,WIN_DEF_UPDATE,ANTIVIRUS,ANTIVIRUS_STATUS,ANTIVIRUS_UPDATE,FIREWALL,FIREWALL_STATUS,ANTISPYWARE,ANTISPYWARE_STATUS,ANTISPYWARE_UPDATE,OS_NAME)'+
  ' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12,:p13) MATCHING (NAMEPC)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(NamePC)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+(DefProd)+'';
FDQueryWrite.params.ParamByName('p3').AsString:=''+(DefStatus)+'';
FDQueryWrite.params.ParamByName('p4').AsString:=''+(DefUpdate)+'';
FDQueryWrite.params.ParamByName('p5').AsString:=''+(AntProd)+'';
FDQueryWrite.params.ParamByName('p6').AsString:=''+(AntStatus)+'';
FDQueryWrite.params.ParamByName('p7').AsString:=''+(AntUpdate)+'';
FDQueryWrite.params.ParamByName('p8').AsString:=''+(FWProd)+'';
FDQueryWrite.params.ParamByName('p9').AsString:=''+(FWStatus)+'';
FDQueryWrite.params.ParamByName('p10').AsString:=''+(AntySpw)+'';
FDQueryWrite.params.ParamByName('p11').AsString:=''+(AntySpwStatus)+'';
FDQueryWrite.params.ParamByName('p12').AsString:=''+(AntySpwUpdate)+'';
FDQueryWrite.params.ParamByName('p13').AsString:=''+(OsName)+'';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
result:=true;
except on E: Exception do
  begin
  if TransactionWrite.Active then  TransactionWrite.Rollback;
  result:=false;
  Log_write(1,'Scan',NamePC+': Ошибка update ANTIVIRUSPRODUCT :'+E.Message);
  end;
end;
end;
/////////////////////////////////////////////////////////////
function GeneralPing.readUsWinDomIsWmi(s:string):bool;
var
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
z,step:integer;
AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate:string;
AntySpw,AntySpwStatus,AntySpwUpdate,AntySpwDef,AntySpwDefStatus,AntySpwDefUpdate:string;
function statusAntivirus (stateAnt,OnOrUpdate:string):string; // функция определения статуса антивируса
begin
if length(stateAnt)<5 then   // если длинна строки статуса меньше 5ти символов
begin
  result:='Неизвестный статус ('+stateAnt+')';
  exit;
end;
if OnOrUpdate='On' then
 begin
 if (copy(stateAnt,2,2)='10') or (copy(stateAnt,2,2)='11') then
  result:='Включен'
  else
 if (copy(stateAnt,2,2)='00') or (copy(stateAnt,2,2)='01') then
  result:='Выключен'
  else result:='Не определено';
 end;
if OnOrUpdate='Update' then
begin
  if (copy(stateAnt,4,2)='00') then
  result:='Ok'
  else
 if (copy(stateAnt,4,2)='10') then
  result:='Базы устарели'
  else result:='Не определено';
end;
end;
begin

//////////////////////////////////////////////
if RPCport then  //// если проверяем 135 порт
if not get135Port(s) then
begin
  result:=false;
  exit; // в принципе можно валить если порт не доступне
end;

try
step:=0;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');  ////IWbemLocator или   SWbemLocator
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', InvUser, InvPass,'','',128); ///WbemUser, WbemPassword ,0- ждать ответа до посинения 128- ждать 2 минуты
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_OperatingSystem','WQL',wbemFlagReturnImmediately);   //wbemFlagReturnImmediately или wbemFlagForwardOnly
step:=1;
//FWMIService.Security_.impersonationlevel:=3;
//FWMIService.Security_.authenticationLevel := 6;
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 if oEnum.Next(1, FWbemObject, iValue) = s_ok then     //// Операционная система
    begin
    step:=2;
    nameOS:=trim(string(FWbemObject.Caption));
    FWbemObject := Unassigned;
    end;
//VariantClear(FWbemObject);
FWbemObjectSet:=Unassigned;
//VariantClear(FWbemObjectSet);
oEnum:=nil;
step:=3;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT UserName,Domain,workgroup,caption FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
step:=4;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
    begin
    step:=5;
    if FWbemObject.UserName<>null then UserNamePC:=trim(string(FWbemObject.UserName));
    if FWbemObject.Domain<>null then PCDomain:=trim(string(FWbemObject.Domain));
    if PCDomain='' then if FWbemObject.workgroup<>null then PCDomain:=trim(string(FWbemObject.Workgroup));
    if FWbemObject.caption<>null then LogonSrv:=string(FWbemObject.caption);
    FWbemObject := Unassigned;
    step:=6;
    end;
step:=7;
oEnum:=nil;
FWbemObjectSet:=Unassigned;
FWMIService:=Unassigned;
result:=true;
step:=8;
Except
on E:Exception do
  begin
  result:=false;
  Log_write(1,'Scan',s+' get wmi info ('+inttostr(step)+') :'+e.Message);
  if not VarIsEmpty(FWMIService) then
  FWMIService:=Unassigned;
 if not VarIsEmpty(FSWbemLocator) then
  FSWbemLocator:=Unassigned;
  exit;   //если ошибка то дальше делать нехуй
  end;
end;
///////////////////////////////////////////////////////////////////////////////////////////
//получаем стутус антивируса  AntivirProd,AntivirStat
//https://theroadtodelphi.com/2011/02/18/getting-the-installed-antivirus-antispyware-and-firewall-software-using-delphi-and-the-wmi/
//https://social.msdn.microsoft.com/Forums/pt-BR/6501b87e-dda4-4838-93c3-244daa355d7c/wmisecuritycenter2-productstate?forum=vblanguage

if (pos('Server',nameOS)=0)and(pos('Windows XP',nameOS)=0) then  // если продукт не серверный и не ХП
BEGIN
try
step:=9;
AntProd:='';AntStatus:='';AntUpdate:='';
FWProd:='';FWStatus:='';
DefProd:='';DefStatus:='';DefUpdate:='';
AntySpw:='';AntySpwStatus:='';AntySpwUpdate:='';
AntySpwDef:='';AntySpwDefStatus:='';AntySpwDefUpdate:='';
FWMIService   := FSWbemLocator.ConnectServer(s, 'root\SecurityCenter2', InvUser, InvPass,'','',128);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM AntiVirusProduct','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
step:=10;
  //'----AntiVirusProduct--------'
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
  if FWbemObject.displayName<>null then
  Begin
  step:=11;
  if pos(ansiuppercase('Defender'),ansiuppercase(vartostr(FWbemObject.displayName)))<>0 then
    begin
    step:=12;
    DefProd:= vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    DefStatus:=statusAntivirus(IntToHex(z,2),'On');
    DefUpdate:=statusAntivirus(IntToHex(z,2),'Update');
    end
    end
  else
  if (AntProd<>vartostr(FWbemObject.displayName)) and (AntStatus<>'Включен') then //
  begin
    step:=13;
    AntProd:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    step:=14;
    AntStatus:=statusAntivirus(IntToHex(z,2),'On');
    AntUpdate:=statusAntivirus(IntToHex(z,2),'Update');
    end;
  end;
  FWbemObject:=Unassigned;
  End;
  step:=15;
  if (AntProd='')and(DefProd<>'') then
  begin
  AntProd:=DefProd;
  AntStatus:=DefStatus;
  AntUpdate:=DefUpdate;
  end;

 //'-------FirewallProduct------'
  oEnum:=nil;
  FWbemObjectSet:=Unassigned;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM FirewallProduct','WQL',wbemFlagForwardOnly);
  step:=16;
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do
  Begin
  if FWbemObject.displayName<>null then
    if (FWProd<>vartostr(FWbemObject.displayName))and(FWStatus<>'Включен') then
    begin
    step:=17;
    FWProd:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
      begin
      FWStatus:=statusAntivirus(IntToHex(z,2),'On');
      end;
    end;
   FWbemObject:=Unassigned;
  End;
  step:=18;
   //-------AntiSpywareProduct--------
  FWbemObjectSet:=Unassigned;
  FWbemObject:=Unassigned;
  oEnum:=nil;
  step:=19;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM AntiSpywareProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  step:=20;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do
  Begin
  if FWbemObject.displayName<>null then
  Begin
  step:=21;
  if pos(ansiuppercase('Defender'),ansiuppercase(vartostr(FWbemObject.displayName)))<>0 then
    begin
    step:=22;
    AntySpwDef:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
      begin
      step:=23;
      AntySpwDefStatus:=statusAntivirus(IntToHex(z,2),'On');
      AntySpwDefUpdate:=statusAntivirus(IntToHex(z,2),'Update');
      end;
    end
  else
   if (AntySpwStatus<>'Включен')or((AntySpwUpdate<>'Ok')and(AntySpwStatus='Включен')) then
    begin
    step:=24;
    AntySpw:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
      begin
      step:=25;
      AntySpwStatus:=statusAntivirus(IntToHex(z,2),'On');
      AntySpwUpdate:=statusAntivirus(IntToHex(z,2),'Update');
      end;
    end;
  End;
 step:=26;
 FWbemObject:=Unassigned;
 End;

step:=27;
if ((AntySpwStatus='Выключен')or (AntySpwStatus=''))and ((AntySpwDefStatus='Включен')) then
 begin
 AntySpw:=AntySpwDef;
 AntySpwStatus:=AntySpwDefStatus;
 AntySpwUpdate:=AntySpwDefUpdate;
 end;
step:=28;
///////Записываем данные и получаем результат
AntivirusUpdateTable(s,AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate,
AntySpw,AntySpwStatus,AntySpwUpdate,nameOS);
step:=29;
AntivirProduct:=AntProd;
if(AntStatus='Выключен')and (DefStatus='Включен') then // если антивирь есть и он выключен, то
begin // проверка на то что работает, может быть установлен антивирь но он не работает. но работает windows Defender
 AntivirProduct:=DefProd;
 AntStatus:=DefStatus;
 AntUpdate:=DefUpdate;
end;
step:=30;

if (AntStatus='Включен')and(AntUpdate='Ok') then AntivirusIdStatus:=18  // статус ОК
else if (AntStatus='Включен')and(AntUpdate='Базы устарели') then AntivirusIdStatus:=21  // не совсем ОК
else if (AntStatus='Выключен')and(AntUpdate='Ok') then AntivirusIdStatus:=19  // ваще не гуд
else if (AntStatus='Выключен')and(AntUpdate='Базы устарели') then AntivirusIdStatus:=20 //хреного
else  AntivirusIdStatus:=19;  // ваще жопа
//AntivirusIdStatus
step:=31;
except
 on E:Exception do
  begin
   Log_write(1,'Scan',s+': Ошибка AntivirusScan ('+inttostr(step)+') :'+E.Message);
  end;
end;
  //if not VarIsEmpty(FWbemObject) then
  FWbemObject := Unassigned;
  oEnum:=nil;
  //if not VarIsEmpty(FWbemObjectSet) then
  FWbemObjectSet:=Unassigned;
  //if not VarIsEmpty(FWMIService) then
  FWMIService:=Unassigned;
  //if not VarIsEmpty(FSWbemLocator) then
  FSWbemLocator:=Unassigned;
  //VariantClear(FSWbemLocator);
END; // окончание сканирования антивирусов
//////////////////////////////////////////////////////////////////////////////////////////////
end;
///////////////////////////////////////////////////////////////



function GeneralPing.GetInfoComputerInAD(pcName:string): Bool;
var
  Info: PWkstaInfo100;
  Error: DWORD;
  CompPWC:PWideChar;
  htoken:THandle;
begin
try
begin
    try ////////////////////////////////////// авторизация пользователя на удаленном компе
     {if not} (LogonUserA (PAnsiChar(InvUser), PAnsiChar (pcName),  // сначала заходим на комп в сети
     PAnsiChar (InvPass), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     //then GetLastError();
     except on E: Exception do frmdomaininfo.Memo1.Lines.Add(pcName+' : Ошибка LogonUser - '+e.Message)
     end;

  CompPWC:=PWideChar(WideString(pcName));
  Error:=NetWkstaGetInfo(CompPWC, 100, @Info);
  if Error<>0 then
  begin
  //raise Exception.Create(SysErrorMessage(Error));
  Log_write(1,'Scan',pcName+': SysErrorMessage GetCurrentComputerOS  :'+SysErrorMessage(Error));
  PCDomain:='';
  LogonSrv:='';
  result:=false;
  end
  else
  begin
  nameOS:='Windows NT '+inttostr(info^.wki100_ver_major)+'.'+inttostr(info^.wki100_ver_minor);
  PCDomain:= info^.wki100_langroup;
  LogonSrv:=info^.wki100_computername;
  result:=true;
  end;

  end;
  Except
   on E: Exception do
      begin
        Log_write(2,'Scan',pcName+': Ошибка GetCurrentComputerOS  :'+E.Message);
        result:=false;
      end;
end;
CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
//NetApiBufferFree(@info);
NetApiBufferFree(info);
end;

///////////////////////////////////////////////////////////////////////
function GeneralPing.EnumNetUsers(HostName: WideString): THostInfo;
 const
   STR_ERROR_ACCESS_DENIED = 'Пользователь не имеет доступа к запрашиваемой информации.';
   STR_ERROR_MORE_DATA = 'Укажите достаточно большой буфер для приема всех записей.';
   STR_ERROR_INVALID_LEVEL = 'Недопустимый параметр уровня.';
 var
   Info: Pointer;
   ElTotal: PDWord;
   ElCount: PDWord;
   Resume: PDWord;
   Error: Longint;
  const MAX_PREFERRED_LENGTH=DWORD(-1);
begin
try

   Resume := 0;
   Error := NetWkstaUserEnum(
   PWideChar(HostName),  /// имякомпа
     1,                  //   0-WKSTA_USER_INFO_0 , 1- WKSTA_USER_INFO_1
     Info,               // Указатель на буфер, который получает данные -- вся инфа получаемая с компов
     MAX_PREFERRED_LENGTH, //Если вы укажете MAX_PREFERRED_LENGTH, функция выделяет объем памяти, необходимый для данных  //256 * Integer(ElTotal)
     ElCount,              /// Указатель на значение, которое получает количество фактически перечисленных элементов.
     ElTotal,               ///Указатель на значение, которое получает общее количество записей, которые могли быть перечислены с текущей позиции возобновления.
     Resume);              //Указатель на значение, которое содержит дескриптор возобновления, который используется для продолжения существующего поиска. Дескриптор должен быть нулевым при первом вызове и оставлен без изменений для последующих вызовов.

   case Error of
     ERROR_ACCESS_DENIED: Result.UserName := STR_ERROR_ACCESS_DENIED;
     ERROR_MORE_DATA: Result.UserName     := STR_ERROR_MORE_DATA;
     ERROR_INVALID_LEVEL: Result.UserName := STR_ERROR_INVALID_LEVEL
       else
         if Info <> nil then
         begin
           Result := THostInfo(info^);
         end
         else
           begin
            Result.UserName      := Pchar(SysErrorMessage(Error));
            Result.logon_domain  := Pchar(SysErrorMessage(Error));
            Result.other_domains := Pchar(SysErrorMessage(Error));
            Result.logon_server  := Pchar(SysErrorMessage(Error));
           end;
  end;
  Except
   on E: Exception do
      begin
      Log_write(2,'Scan',HostName+': Ошибка NetWkstaUserEnum  :'+E.Message);
      Log_write(2,'Scan',HostName+': Ошибка NetWkstaUserEnum  :'+SysErrorMessage(Error));
      end;
  end;
  NetApiBufferFree(info);
 end;


/////////////////////////////////////// ARP
function GeneralPing.MySendARP(const MyIPAddress: String): String;
var
  DestIP: ULONG;
  MacAddr: Array [0..5] of Byte;
  MacAddrLen: ULONG;
  SendArpResult: Cardinal;
begin
try
if (MyIPAddress='0.0.0.0') or (MyIPAddress='') then
begin
Result := '';
exit;
end;

  DestIP := inet_addr(PAnsiChar(AnsiString(MyIPAddress)));
  MacAddrLen := Length(MacAddr);
  SendArpResult := SendARP(DestIP, 0, @MacAddr, @MacAddrLen);

  if SendArpResult = NO_ERROR then
    Result := Format('%2.2X-%2.2X-%2.2X-%2.2X-%2.2X-%2.2X',
                     [MacAddr[0], MacAddr[1], MacAddr[2],
                      MacAddr[3], MacAddr[4], MacAddr[5]])
  else
  Result := '';
  NetApiBufferFree(@MacAddr);
  NetApiBufferFree(@MacAddrLen);
Except
   on E: Exception do
      begin
      Result := '';
      Log_write(2,'Scan',MyIPAddress+': Ошибка ARP  - '+E.Message);
      end;
  end;
  end;

function GeneralPing.WriteAllResult(s:string;Item:integer;StatusLanPC:string):boolean;
var
step:integer;
begin
//EventUserLogin - пользователь в системе или нет
//Rendering - отрисовка в listview далеж если пользователь не в системе
 try
   if RenderingPing then // если отрисовка в любом случае
   begin
   step:=0;
    with frmDomaininfo.ListView8.Items[Item] do
     begin
     step:=1;
       ImageIndex:=strtoint(StatusLanPC); //статус компьютера в сети
       SubItems[2]:=AnswerMAC;    /// MAC адрес
       SubItems[3]:=nameOS;       /// операционка
       SubItems[4]:=UserNamePC;  /// Пользователь
       SubItems[5]:=PCDomain;  /// Домен
       SubItems[6]:=LogonSrv;  /// имя компа
       SubItems[7]:=IpADRS;    //// IP
       if ClientuRDM='Доступен' then SubItemImages[9]:=1 // доступен
       else SubItemImages[9]:=2; // недоступен
       SubItems[9]:=ClientuRDM; // клиент uRDM
       SubItems[12]:=AntivirProduct;
       SubItemImages[12]:=AntivirusIdStatus;
     end;
   end
   else // иначе если только когда пользователь в системе
   if EventUserLogin then
   begin
   step:=2;
    with frmDomaininfo.ListView8.Items[Item] do
     begin
     step:=3;
       ImageIndex:=strtoint(StatusLanPC); //статус компьютера в сети
       SubItems[2]:=AnswerMAC;    /// MAC адрес
       SubItems[3]:=nameOS;       /// операционка
       SubItems[4]:=UserNamePC;  /// Пользователь
       SubItems[5]:=PCDomain;  /// Домен
       SubItems[6]:=LogonSrv;  /// имя компа
       SubItems[7]:=IpADRS;    //// IP
       if ClientuRDM='Доступен' then SubItemImages[9]:=1 // доступен
       else SubItemImages[9]:=2; // недоступен
       SubItems[9]:=ClientuRDM; // клиент uRDM
       SubItems[12]:=AntivirProduct;
       SubItemImages[12]:=AntivirusIdStatus;
     end;
   end;
    step:=4;
    try //положительные результаты пишем в базу в любом случае не зависимо от того в системе юзер или нет
    TransactionWrite.StartTransaction;
    step:=5;
    FDQueryWrite.Close;
    FDQueryWrite.SQL.Clear;
    FDQueryWrite.SQL.Text:='update or insert into MAIN_PC '+
    '(PC_NAME,PC_OS,ANSWER_MAC,CUR_USER_NAME,CUR_DOMAIN,OTHER_NAME,CUR_IP_ADRS,URDM_CLIENT,ANTIVIRUS_PRODUCT,ANTIVIRUS_STATUS,EXPT_PC) VALUES'+
    ' (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11) MATCHING (PC_NAME)';
    FDQueryWrite.params.ParamByName('p1').AsString:=''+(s)+'';
    FDQueryWrite.params.ParamByName('p2').AsString:=''+leftstr(nameOS,200)+'';
    FDQueryWrite.params.ParamByName('p3').AsString:=''+leftstr(AnswerMAC,30)+'';
    FDQueryWrite.params.ParamByName('p4').AsString:=''+leftstr(UserNamePC,50)+'';
    FDQueryWrite.params.ParamByName('p5').AsString:=''+leftstr(PCDomain,100)+'';
    FDQueryWrite.params.ParamByName('p6').AsString:=''+leftstr(LogonSrv,50)+'';
    FDQueryWrite.params.ParamByName('p7').AsString:=''+leftstr(IpADRS,50)+'';
    FDQueryWrite.params.ParamByName('p8').AsString:=''+leftstr(ClientuRDM,50)+'';
    FDQueryWrite.params.ParamByName('p9').AsString:=''+leftstr(AntivirProduct,250)+'';
    FDQueryWrite.params.ParamByName('p10').AsInteger:=AntivirusIdStatus;
    FDQueryWrite.params.ParamByName('p11').AsString:=statusLanPC;
    FDQueryWrite.ExecSQL;
    TransactionWrite.Commit;
    step:=6;
    except
    TransactionWrite.Rollback;
    end;
  result:=true;
  step:=7;
 except on E: Exception do
  begin
  Log_write(2,'Scan',S+': Ошибка update MAIN_PC ('+inttostr(step)+') :'+E.Message);
  result:=false;
  end;
 end;
end;

//////////////////////////////////////////////////////////////
function GeneralPing.OsherInfoPC(s:string;Item:integer):bool;
var
HostInfo: THostInfo;
StatusLanPC:String;
begin
try
statusLanPC:='0';
if not readUsWinDomIsWmi(s) then  //// считываем инфу через wmi, если не получилось запрашиваем через API
begin
  if RPCport then // если сканируем 135 порт то можно инфу плучить через API
  if GetInfoComputerInAD(s)then
    begin
    HostInfo := EnumNetUsers(s);
    UserNamePC:= HostInfo.username;
    HostInfo.username:='';
    HostInfo.logon_domain:='';
    HostInfo.other_domains:='';
    HostInfo.logon_server:='';
    end;
 StatusLanPC:=inttostr(6); // если  WMI не доступен
 end
else StatusLanPC:=inttostr(4); // если  WMI доступен

if (nameOS<>'') or (AnswerMAC<>'') then  // если знаем операционку или MAC (если комп в сети то мас всегда известен)
  Begin
  WriteAllResult(s,item,StatusLanPC);  //запись в listview и базу данных
  End;
Except
   on E: Exception do
      begin
      Log_write(2,'Scan',S+': Data retrieval :'+E.Message);
      end;
  end;
//HostInfo:=nil;
end;



function GeneralPing.ConnectiontDB:boolean;
begin
try
ConnectionThReadPing:=TFDConnection.Create(nil);
ConnectionThReadPing.DriverName:='FB';
ConnectionThReadPing.Params.database:=databaseName; // расположение базы данных
ConnectionThReadPing.Params.Add('server='+databaseServer);
ConnectionThReadPing.Params.Add('port='+databasePort);
ConnectionThReadPing.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThReadPing.Params.Add('CharacterSet=UTF8');
ConnectionThReadPing.Params.add('sqlDialect=3');
ConnectionThReadPing.Params.DriverID:=databaseDriverID;
ConnectionThReadPing.Params.UserName:=databaseUserName;
ConnectionThReadPing.Params.Password:=databasePassword;
ConnectionThReadPing.Connected:=true;
ConnectionThReadPing.LoginPrompt:= false;  /// отображение диалога user password
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThReadPing;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
FDQueryWrite:=TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=ConnectionThReadPing;
/// ///////////////////////////////////////////////
result:=true;
except on E: Exception do
begin
result:=false;
Log_write(3,'Scan',': Error connection db  :'+E.Message);
end;
end;
end;

function GeneralPing.DisconnectDB:boolean;
begin
try
freeandnil(FDQueryWrite);
if TransactionWrite.Active then TransactionWrite.Rollback;
freeandnil(TransactionWrite);
ConnectionThReadPing.Connected:=false;
ConnectionThReadPing.Close;
freeandnil(ConnectionThReadPing);
result:=true;
except on E: Exception do
begin
result:=false;
Log_write(3,'Scan',': Error disconnect db  :'+E.Message);
end;
end;
end;

function GeneralPing.EslePcNotLan(ImPing,ImuRDM:integer;NamePC:string):boolean;
var
z,step:integer;
begin
try
//EventUserLogin - пользователь в системе или нет
//Rendering - отрисовка в listview далеж если пользователь не в системе
step:=0;
if RenderingPing then // если отрисовка то в любом случае рисуем
begin
step:=1;
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
 if frmDomainInfo.ListView8.Items[z].SubItems[0]=NamePC then  /// комп не в сети Превышен интервал ожидания запроса
  begin
  frmDomaininfo.ListView8.Items[z].ImageIndex:=ImPing;    /// узел не доступен
  frmDomaininfo.ListView8.Items[z].SubItemImages[9]:=ImuRdm; //// urdm не в сети соответственно
  break;
  end;
end
else // иначе отрисовка только когда пользователь в системе
 if EventUserLogin then
 begin
 step:=2;
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
  if frmDomainInfo.ListView8.Items[z].SubItems[0]=NamePC then  /// комп не в сети Превышен интервал ожидания запроса
  begin
  frmDomaininfo.ListView8.Items[z].ImageIndex:=ImPing;    /// узел не доступен
  frmDomaininfo.ListView8.Items[z].SubItemImages[9]:=ImuRdm; //// urdm не в сети соответственно
  break;
  end;
 end;

 if (not RenderingPing) and (not EventUserLogin) then // если нет отрисовки и пользователь не в системе то пишем значения о недоступности пк в базу
 begin
 step:=3;
  TransactionWrite.StartTransaction;
  FDQueryWrite.Close;
  FDQueryWrite.SQL.Clear;
  FDQueryWrite.SQL.Text:='update or insert into MAIN_PC '+
  '(PC_NAME,URDM_CLIENT,EXPT_PC) VALUES'+
  ' (:p1,:p2,:p3) MATCHING (PC_NAME)';
  FDQueryWrite.params.ParamByName('p1').AsString:=''+(NamePC)+'';
  FDQueryWrite.params.ParamByName('p2').AsString:=''+leftstr(ClientuRDM,50)+'';
  FDQueryWrite.params.ParamByName('p3').AsString:=inttostr(ImPing);
  FDQueryWrite.ExecSQL;
  TransactionWrite.Commit;
 end;
PcInLan(NamePC,false); /// проверяем есть ли комп в списке достпуных ПК, если есть то удаляем
result:=true;
step:=4;
except on E: Exception do
begin
if TransactionWrite.Active then TransactionWrite.Rollback;
Log_write(3,'Scan','Update list View pc ('+inttostr(step)+') :'+ e.Message);
result:=false;
end;
end;
end;

procedure GeneralPing.ping;
var
i,z:integer;
predComp:string;
IniSettings:TMeminiFile;
begin
if not ConnectiontDB then  // подключаемся к БД и создаем необходимые компоненты
 begin
  DisconnectDB;
  Log_write(3,'Scan','Не установленно соединение с базой данных инвентаризация завершена');
  SolveExitInvScan:=true;
  OutForPing:=false;
  exit;
 end;
OleInitialize(nil);
Log_write(0,'Scan','Запуск сканирования сети');
SolveExitInvScan:=false; // признак того что поток запущен так же устанавливается при нажатии сканировать во избежании резкой остановки потока если он еще не запущен
RenderingPing:=false;
IniSettings:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
RenderingPing:=IniSettings.Readbool('Scan','Rendering',false); // обновлять или нет listview при блокировке пользователя
if RenderingPing then Log_write(0,'Scan',' Отрисовка включена')
else Log_write(0,'Scan',' Отрисовка выключена');
finally
IniSettings.Free;
end;
while (OutForPing)and (ConnectionThReadPing.Connected) do //// пока нет признака завершения потока и подключение к базе активно повторяем цикл
BEGIN
try
for i := 0 to PingPCList.Count-1 do
Begin
if not ConnectionThReadPing.Connected then
 begin
   Log_write(3,'Scan',' Потеряно соединение с базой данных');
   break;
 end;
  if (OutForPing=false) then break; // если сканирование завершено то выходим из цила
  IpADRS:=''; AnswerMAC:=''; ClientuRDM:='';
  nameOS:=''; UserNamePC:=''; PCDomain:='';
  LogonSrv:=''; AntivirusIdStatus:=22;  AntivirProduct:='';
  ClientuRDM:='Не доступен';
  try
   if not SelectedPing(IpADRS,PingPCList[i],pingtimeout) then
   Begin
    EslePcNotLan(5,2,PingPCList[i]);
    End
   else
    for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
    if frmDomainInfo.ListView8.Items[z].SubItems[0]=PingPCList[i] then
     begin
     AnswerMAC:=MySendARP(IpADRS);
     if uRDMScan then
     if getUrdmClient(IpADRS) then ClientuRDM:='Доступен';/// проверяем доступен ли на узле клиент uRDM
     OsherInfoPC(PingPCList[i],z);  /// получаем и записываем дополнительную информацию с компа и выводим стутусы
     PcInLan(PingPCList[i],true); /// проверяем есть ли комп в списке достпуных если нет то добавляем
     break;                       // валим из цикла
     end;
 Except
  begin
  EslePcNotLan(5,2,PingPCList[i]);
  end;
 End; // исключения в цикле
End; //Цикл по списку ПК

Except
on E: Exception do
  begin
  Log_write(3,'Scan',': Ошибка, сканирования сети. :'+E.Message);
  i:=0;
  frmDomainInfo.StatusBar1.Panels[1].Text:='ПК в сети: 0';
  end;
End; // исключение всего цикла

if OutForPing then  // если сканирование не остановлено то переподключаемся к базе
  begin
  if ConnectionThReadPing.Connected then DisconnectDB; //отключаемся от БД и все компоненты херим
   if PingPCList.Count<3 then   // если компов меньше 3х то засыпаем на 5 сек после цикла
    begin
    sleep(5000);
    end
   else sleep(100000 div PingPCList.Count); //// пауза в потоке, если список компов маленький или вообще комп один
  if not ConnectiontDB then break; // подключаемся к БД и создаем необходимые компоненты
  Log_write(0,'Scan',' Переподключение к БД.');
  end;
END; // пока поток не остановлен
Log_write(0,'Scan',' Сканирование сети завершено.');
listAvaliblePC.Clear; /// очищаем лист со списком доступных компов
freeandnil(listAvaliblePC);
PingPCList.Clear;
freeandnil(PingPCList);
if ConnectionThReadPing.Connected then DisconnectDB; //отключаемся от БД
OleUnInitialize();
nameOS:='';
UserNamePC:='';
PCDomain:='';
LogonSrv:='';
AnswerMAC:='';
IpADRS:='';
OutForPing:=false; //// поток завершен, можно запускать его еще раз
SolveExitInvScan:=true; //// разрешить закрыть программу
end;


procedure GeneralPing.Execute;
begin
InvUser:=frmDomainInfo.LabeledEdit1.Text;
InvPass:=frmDomainInfo.LabeledEdit2.Text;
listAvaliblePC:=TStringList.Create; /// создаем лист, туда будем складывать доступне ПК
listAvaliblePC.Clear;
ping;
//OutForPing:=false;
end;







end.
