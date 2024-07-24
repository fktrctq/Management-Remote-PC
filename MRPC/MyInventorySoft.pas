unit MyInventorySoft;

interface

uses
    System.Classes,ActiveX,ComObj,Variants,FireDAC.Stan.Intf, FireDAC.Stan.Option
  , FireDAC.Stan.Param,System.SysUtils,Vcl.Forms,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,FireDAC.UI.Intf
  ,FireDAC.Comp.UI,FireDAC.Phys.FB,inifiles,SqlTimSt,WinAPI.Windows,IdIcmpClient
  ,System.StrUtils,IdTCPClient;

type
  InventorySoft = class(TThread)
  private
    function SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
    function Log_write(level:integer;fname, text:string):string;
    function get135Port(s:string):bool;
    function WbemTimeToDateTime(const V : OleVariant): string; //перевод даты
    function InstallSoft(namePC:string):bool;
    function ConnectDBSoft:boolean;
    Function DisconnectDBSoft:boolean;
    function writeSoftInDB(SOFTNAME,SOFT_VERSION,MANUFACTURE:string):string;
  protected
    procedure Execute; override;

  end;

implementation
uses uMain,PingForInventorySoft;

ThreadVar
  FSWbemLocatorSoft : OLEVariant;
  FWMIServiceSoft   : OLEVariant;
  ConnectionThreadSoft: TFDConnection;
  TransactionWriteSoft: TFDTransaction;
  FDQueryWriteSoft: TFDQuery;



function InventorySoft.Log_write(level:integer;fname, text:string):string;
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
        f.Destroy;
      end;
except
  exit;
end;
end;

//////////////////////////////////////////////////////
function InventorySoft.get135Port(s:string):bool;
var
ScanTCPPort : TIdTCPClient;
begin
try
if not RPCport then  /// если не проверям 135 порт
begin
result:=true;
exit;
end;
ScanTCPPort:=TIdTCPClient.Create;
ScanTCPPort.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
ScanTCPPort.Port:=135;
ScanTCPPort.ConnectTimeout:=pingtimeout;
ScanTCPPort.Host:=s;
ScanTCPPort.Connect;
try
if ScanTCPPort.Connected then result:=true
else result:=false;
ScanTCPPort.Disconnect;
finally
if ScanTCPPort.Connected then ScanTCPPort.Disconnect;
ScanTCPPort.free;
end;
 Except
 begin
 result:=false;
 Log_write(0,'Software',s+': Ошибка scan 135 potr ');
 end;
 end;
end;
/////////////////////////////////////////////////////////////////////

function InventorySoft.SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
var
avalible:boolean;
begin
resIP:='';
  case PingType of
    1:   avalible:=PingIdIcmp(resIP,HostName,pingtimeout);
    2:   avalible:=PingGetaddrinfo(resIP,HostName,pingtimeout);
    3:   avalible:=PingGetHostByName(resIP,HostName,pingtimeout);
  END;
  resIP:='';
  Result:=avalible;
  if not avalible then
  begin
  try
  TransactionWriteSoft.StartTransaction;
  FDQueryWriteSoft.close;
  FDQueryWriteSoft.SQL.Clear;
  FDQueryWriteSoft.SQL.Text:='update or insert into MAIN_PC_SOFT (PC_NAME,DATE_INV,TIME_INV,ERROR_INV) VALUES'
     +' ('''+HostName+''','''+datetostr(date)+''','''+TimeToStr(time)+''',''Узел не доступен'') MATCHING (PC_NAME)';
  FDQueryWriteSoft.ExecSQL;
  TransactionWriteSoft.Commit;
  frmDomainInfo.statusbarSoft.Panels[3].Text:=HostName+' - не доступен';
  //Log_write(0,'Software',HostName+' Узел не доступен') ;
  except on E: Exception do
  Log_write(1,'Software',HostName+' Ошибка SelectedPing '+e.Message) ;
  end;
  end;

end;


function InventorySoft.WbemTimeToDateTime(const V : OleVariant): string; //перевод даты
var
  Dt: string;
  FormatSettings,FormatSettings1: TFormatSettings;
begin
  try
  if VarIsNull(V) then exit;
  Dt:=string(v);
  insert('/',dt,5);
  insert('/',dt,8);
  FormatSettings.ShortDateFormat:='yyyy-mm-dd';
  FormatSettings.DateSeparator:='/';
  FormatSettings1.ShortDateFormat:='dd.mm.yyyy';
  FormatSettings1.DateSeparator:=':';
  result:=datetostr(strtodate(dt,FormatSettings),FormatSettings1)
  except
   result:='01.01.2001';
  end;
end;

function scantext(s:string):string;
var
ss:string;
begin
result:=StringReplace(s,'''','',[rfReplaceAll]);
end;

function InventorySoft.writeSoftInDB(SOFTNAME,SOFT_VERSION,MANUFACTURE:string):string;
begin
try  
TransactionWriteSoft.StartTransaction;
FDQueryWriteSoft.close;
FDQueryWriteSoft.SQL.Text:=
'update or insert into SOFT_PC (SOFT_NAME,SOFT_VERSION,MANUFACTURE)'
+'VALUES(:p1,:p2,:p3) MATCHING ('+InvSoftWarning+')'
+' RETURNING ID';
FDQueryWriteSoft.Params.ParamByName('p1').Value:=''+leftstr(SOFTNAME,400)+'';
FDQueryWriteSoft.Params.ParamByName('p2').Value:=''+leftstr(SOFT_VERSION,50)+'';
FDQueryWriteSoft.Params.ParamByName('p3').Value:=''+leftstr(MANUFACTURE,400)+'';
FDQueryWriteSoft.Open;
if FDQueryWriteSoft.Fields[0].Value<>null then
result:=(FDQueryWriteSoft.Fields[0].AsString)
else result:='';
TransactionWriteSoft.Commit;
except //// ловим ошибку записи в БД
on E:Exception do
  begin
  if TransactionWriteSoft.Active then
  begin
   TransactionWriteSoft.Rollback;  
  end;
  sleep(500);
  Log_write(1,'Software',' Ошибка записи/чтения данных о программе - '+E.Message);
  result:='';
  end;
end;
end;

function InventorySoft.InstallSoft(namePC:string):bool;
var
i,id:integer;
SoftName,DIRECT_INSTALL,SOURCE_INST,DATE_INSTALL,
SOFT_VERSION,UNINSTALL_STR,MANUFACTURE,invok,ErrorSanSoft:string;
FWbemObjectSet: OLEVariant;
FInParams,FInParams1                    : OLEVariant;
FOutParams,FOutParams1                  : OLEVariant;
stringSoftID,CurrentID:string;
begin
try /////// ловим общую ошибку при сканированиии компа
    ErrorSanSoft:='';
    id:=0; /// количество программ
    invok:='OK';
    stringSoftID:='';
    CurrentID:='';
    FWbemObjectSet:= FWMIServiceSoft.Get('StdRegProv');
    FInParams     := FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
    FInParams.hDefKey:=HKEY_LOCAL_MACHINE;
    FInParams.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall';
    FoutParams    := FWMIServiceSoft.ExecMethod('StdRegProv', 'EnumKey', FInParams);
    FInParams1     := FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_();
    FInParams1.hDefKey:=HKEY_LOCAL_MACHINE;
    ////////////////////////////////
for i:= VarArrayLowBound(FOutParams.sNames,1) to VarArrayHighBound(FOutParams.sNames,1) do
begin
 try  /// ловим ошибку при сканировании реестар в цикле
        SOFT_VERSION:='';
        SoftName:='';
        MANUFACTURE:='';
        FInParams1.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall\'+(string(FOutParams.sNames[i]));
        FInParams1.sValueName :='DisplayName';
        FOutParams1    := FWMIServiceSoft.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
        if (FoutParams1.sValue<>null) and (string(FoutParams1.sValue)<>'') then
         begin
         SoftName:=string(FoutParams1.sValue);
        ///////////////////////////////////////////////////////////////////////////
         FInParams1.sValueName :='DisplayVersion';
         FOutParams1:=Unassigned;
         FOutParams1    := FWMIServiceSoft.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
         if FoutParams1.sValue<>null then
         SOFT_VERSION:=string(FoutParams1.sValue)
         else SOFT_VERSION:='';
            ///////////////////////////////////////////////////////////////////////////////////
         FInParams1.sValueName :='Publisher';
         FOutParams1:=Unassigned;
         FOutParams1    := FWMIServiceSoft.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
         if FoutParams1.sValue<>null then MANUFACTURE:=string(FoutParams1.sValue)
         else  MANUFACTURE:='';
             /////////////////////////////////////////////////////////////////////////////////
         CurrentID:=writeSoftInDB(SoftName,SOFT_VERSION,MANUFACTURE); // запись чтение из БД
         if CurrentID<>'' then stringSoftID:=stringSoftID+CurrentID+',';
         inc(id);
        end; // if
       FOutParams1:=Unassigned;
   except /// ловим ошибку при сканировании реестра в цикле
   on E:Exception do
     begin
     FOutParams1:=Unassigned;
     ErrorSanSoft:='Ошибка сканирования компьютера';
     result:=false;
     Log_write(1,'Software','Ошибка сканирования компьютера '+namePC+' - id='+inttostr(id) +' - '+ E.Message);
     end;
   end;
end; //цикл

    if stringSoftID<>'' then //если есть что записать то записываем в базу
    begin
    TransactionWriteSoft.StartTransaction; //запись в базу после цикла
    FDQueryWriteSoft.close;
    FDQueryWriteSoft.SQL.Clear;
    FDQueryWriteSoft.SQL.Text:='update or insert into '
    +' MAIN_PC_SOFT (PC_NAME,DATE_INV,TIME_INV,RESUL_INV,INST_SOFT,COUNT_SOFT,ERROR_INV) VALUES'
    +' ('''+namePC+''','''+datetostr(date)+''','''+TimeToStr(time)+''','''+invOK+''''
    +','''+stringSoftID+''','''+inttostr(id)+''','''+leftstr(ErrorSanSoft,700)+''')'
    +' MATCHING (PC_NAME)';
    FDQueryWriteSoft.ExecSQL;
    TransactionWriteSoft.Commit;
    end;
    ////////////////////////////////////////////////////////////////////////////////////
    result:=true;
    frmDomainInfo.statusbarSoft.Panels[3].Text:=namePC+' - OK';

except   /////// ловим общую ошибку при сканированиии компа
  on E:Exception do
    begin
    invok:='Error';
    frmDomainInfo.statusbarSoft.Panels[3].Text:=namePC+' - Error';
    if TransactionWriteSoft.Active then TransactionWriteSoft.Rollback;
    {FDQueryWriteSoft.Close;
    TransactionWriteSoft.StartTransaction;
    FDQueryWriteSoft.SQL.Clear;
    FDQueryWriteSoft.SQL.Text:='update or insert into '
    +' MAIN_PC_SOFT (PC_NAME,DATE_INV,TIME_INV,RESUL_INV,INST_SOFT,COUNT_SOFT,ERROR_INV) VALUES'
    +' ('''+namePC+''','''+datetostr(date)+''','''+TimeToStr(time)+''','''+invok+''','''+stringSoftID+''','''+inttostr(id)
    +''','''+E.Message+''') MATCHING (PC_NAME)';
    FDQueryWriteSoft.ExecSQL;
    TransactionWriteSoft.Commit;
    FDQueryWriteSoft.Close;}
    Log_write(2,'Software',namePC+ ' - Общая ошибка получения данных: '+E.Message);
    result:=false;
    end;
  end;
FInParams:=Unassigned;
FInParams1:=Unassigned;
FoutParams:=Unassigned;
FoutParams1:=Unassigned;
FWbemObjectSet:=Unassigned;
end;


function InventorySoft.ConnectDBSoft:boolean;
begin
try
ConnectionThReadSoft:=TFDConnection.Create(nil);
ConnectionThReadSoft.DriverName:='FB';
ConnectionThReadSoft.Params.database:=databaseName; // расположение базы данных
ConnectionThReadSoft.Params.Add('server='+databaseServer);
ConnectionThReadSoft.Params.Add('port='+databasePort);
ConnectionThReadSoft.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThReadSoft.Params.Add('CharacterSet=UTF8');
ConnectionThReadSoft.Params.add('sqlDialect=3');
ConnectionThReadSoft.Params.DriverID:=databaseDriverID;
ConnectionThReadSoft.Params.UserName:=databaseUserName;
ConnectionThReadSoft.Params.Password:=databasePassword;
ConnectionThReadSoft.Connected:=true;
ConnectionThReadSoft.LoginPrompt:= false;  /// отображение диалога user password
TransactionWriteSoft:= TFDTransaction.Create(nil);
TransactionWriteSoft.Connection:=ConnectionThReadSoft;
TransactionWriteSoft.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
TransactionWriteSoft.Options.AutoCommit:=false;
TransactionWriteSoft.Options.AutoStart:=false;
TransactionWriteSoft.Options.AutoStop:=false;
FDQueryWriteSoft:=TFDQuery.Create(nil);
FDQueryWriteSoft.Transaction:=TransactionWriteSoft;
FDQueryWriteSoft.Connection:=ConnectionThReadSoft;
result:=true;
except on E: Exception do
begin
result:=false;
log_write(3,'Software','Ошибка подключения к базе данных: '+e.Message);
end;
end;
end;

Function InventorySoft.DisconnectDBSoft:boolean;
begin
try
if Assigned(FDQueryWriteSoft) then  FDQueryWriteSoft.free;
if TransactionWriteSoft.Active then TransactionWriteSoft.Rollback;
if Assigned(TransactionWriteSoft) then TransactionWriteSoft.free;
if Assigned(ConnectionThReadSoft) then
begin
ConnectionThReadSoft.Close;
ConnectionThReadSoft.Free;
end;
result:=true;
except on E: Exception do
begin
result:=false;
log_write(3,'Software','Ошибка отключения от базы данных: '+e.Message);
end;
end;
end;


procedure InventorySoft.Execute;
var
i:integer;
ListPCSoft:TstringList;
InvUser,InvPass:string;
FWbemObject:olevariant;
oEnum      : IEnumvariant;
FWbemObjectSet: OLEVariant;
iValue        : LongWord;
CountPCOKSoft:integer;
resIP:string;
begin
try
if not ConnectDBSoft then
begin
DisconnectDBSoft;
InventSoft:=false;
SolveExitInvSoft:=true; //// разрешить закрыть программу
exit;
end;
///// пользователь и пароль
InvUser:=MyUser;
InvPass:=MyPasswd;
////// список компьютеров
ListPCSoft:=TstringList.Create;
ListPCSoft.Text:=ListPCConf.Text;
ListPCConf.Free;
////////////////////////////
CountPCOKSoft:=0; // количество ПК со статусом ОК
SolveExitInvSoft:=false;
OleInitialize(nil);
log_write(0,'Software','Запуск инвентаризации');
for I := 0 to ListPCSoft.Count-1 do
  begin
  if not InventSoft then break;
    if (SelectedPing(resIP,ListPCSoft[i],pingtimeout)) and (get135Port(ListPCSoft[i])) then
      begin
        try
        if not ConnectionThReadSoft.Connected then
        begin
        log_write(3,'Software','Потеряно соединение с базой данных');
        break;  
        end;        
        FSWbemLocatorSoft := CreateOleObject('WbemScripting.SWbemLocator');
        FWMIServiceSoft   := FSWbemLocatorSoft.ConnectServer(ListPCSoft[i], 'root\CIMV2', InvUser, InvPass,'','',128);
        /////////////////////////////////////////////////   проверка на windows XP , её мы исключаем
        FWbemObjectSet:= FWMIServiceSoft.ExecQuery('SELECT caption FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
        oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
        if oEnum.Next(1, FWbemObject, iValue) = 0 then    //// Пользователь
          begin
          if (FWbemObject.caption)<>null then
          if pos('XP',vartostr(FWbemObject.caption))=0 then  // если операционка не XP то проводим инвентаризацию 
          if InstallSoft(ListPCSoft[i])then inc(CountPCOKSoft); //  если удачно то увеличиваем количестве удачных ПК        
          FWbemObject:=Unassigned;
          end;
        //////////////////////////////////////////////////////
        except   //// если пинг есть но RPC не доступен или еще какая фигня, то записываем ошибку
          on E:Exception do
            begin
            TransactionWriteSoft.StartTransaction;
            FDQueryWriteSoft.close;
            FDQueryWriteSoft.SQL.Clear;
            FDQueryWriteSoft.SQL.Text:='update or insert into MAIN_PC_SOFT (PC_NAME,DATE_INV,TIME_INV,ERROR_INV) VALUES'
            +' ('''+ListPCSoft[i]+''','''+datetostr(date)+''','''+timetostr(time)+''','''+e.Message+''') MATCHING (PC_NAME)';
            FDQueryWriteSoft.ExecSQL;
            TransactionWriteSoft.Commit;
            frmDomainInfo.statusbarSoft.Panels[4].Text:=ListPCSoft[i]+'-'+e.Message;
            end;
          end;
      FWbemObject:=Unassigned;
      FWbemObjectSet:=Unassigned;
      FWMIServiceSoft:=Unassigned;
      FSWbemLocatorSoft:=Unassigned;
      oEnum:=nil;
      end;
frmDomainInfo.statusbarSoft.Panels[4].Text:=inttostr(ListPCSoft.Count)+'/'+inttostr(i+1)+'/'+inttostr(CountPCOKSoft)+' - OK';
if not InventSoft then break;
end;  // цикл по ПК

Except
on E:Exception do
  begin
  log_write(3,'Software','Общая ошибка сканирования ПО:'+e.Message);
  SolveExitInvSoft:=true; //// разрешить закрыть программу
  InventSoft:=false;
  end;
end;
DisconnectDBSoft; // если соединение установлено, то отключаемся от базы
frmDomainInfo.statusbarSoft.Panels[3].Text:='Инвентаризация завершена';
log_write(0,'Software','Инвентаризация завершена');
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
FWMIServiceSoft:=Unassigned;
FSWbemLocatorSoft:=Unassigned;;
OleUnInitialize();
InventSoft:=false;
SolveExitInvSoft:=true; //// разрешить закрыть программу
end;


end.

