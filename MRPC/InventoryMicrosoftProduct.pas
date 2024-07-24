unit InventoryMicrosoftProduct;


interface
uses System.Classes,ActiveX,ComObj,Variants,FireDAC.Stan.Intf, FireDAC.Stan.Option
  , FireDAC.Stan.Param,System.SysUtils,Vcl.Forms,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,FireDAC.UI.Intf
  ,FireDAC.Comp.UI,FireDAC.Phys.FB,inifiles,SqlTimSt,IdIcmpClient,System.StrUtils,
  WinSock,Windows,System.TimeSpan,IdTCPClient;

type
  TPropMicProduct = record
    NamePC       :TstringList;
    User         :String[255];
    Pass         :String[255];
end;

function ScanMicrosoftProduct(MicProduct:pointer):integer; /// Поток для сканирования продуктов Microsoft

Var
CallMicroProduct:TPropMicProduct;

implementation
uses umain;

ThreadVar
MicrosoftProductP: ^TPropMicProduct;


function ScanMicrosoftProduct(MicProduct:pointer):integer; /// Поток для сканирования продуктов Microsoft
var
  ListPC:TstringList;
  User,Pass,CurrentPC:string;
  i,CoutPCOK:integer;
  WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptWin:string;
  OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList;
  sOfficeProduct,sStatOfLic,sKeyOfc,sTypelicOffice,sDescriptLicOffice,nameOS,VerOS:string;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  ConnectionThread    : TFDConnection;
  TransactionWrite    : TFDTransaction;
  FDQueryWrite        : TFDQuery;
///////////////////////////////////////////////////////////
/// ниже описаны необходимые функции для потока///////////
function Log_write(fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
  exit;
end;
end;

function pingWin (s:string):boolean;
var
BufHost:string;
IdIcmpClientConf    : TIdIcmpClient;
begin
try
BufHost:='';
result:=false;
/////////////////////////////////////////////////////
IdIcmpClientConf:=TIdIcmpClient.Create;
IdIcmpClientConf.PacketSize:=32;
IdIcmpClientConf.Port:=0;
IdIcmpClientConf.Protocol:=1;
IdIcmpClientConf.ReceiveTimeout:=pingtimeout;
//////////////////////////////////////////
IdIcmpClientConf.ReplyStatus.FromIpAddress:='';
IdIcmpClientConf.ReplyStatus.Msg:='';
IdIcmpClientConf.host:=s;
IdIcmpClientConf.Port:= 0;
BufHost:= s + StringOfChar (' ', 255);
IdIcmpClientConf.Ping(BufHost);
if  (IdIcmpClientConf.ReplyStatus.Msg<>'Echo') 
or (IdIcmpClientConf.ReplyStatus.FromIpAddress='0.0.0.0') then result:=false
else result:=true;
except
on E: Exception do
begin
  Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+s+' - Ping - '+e.Message);
  result:=false;
end;
end;
s:='';
if Assigned(IdIcmpClientConf) then FreeAndNil(IdIcmpClientConf); /// уничтожение клиента ICMP
end;

function TypeLicinAndOffice(Des:string):string;
begin
if pos('OEM_DM',des)<>0 then result:='OEM_DM' else
if pos('OEM_COA_NSLP',des)<>0 then result:='OEM_COA_NSLP' else
if pos('OEM_SLP',des)<>0 then result:='OEM_SLP' else 
if pos('OEM_COA_SLP',des)<>0 then result:='OEM_COA_SLP' else 
if pos('OEM_NO_NSLP',des)<>0 then result:='OEM_NO_NSLP' else 
if pos('RETAIL',des)<>0 then result:='RETAIL' else  
if pos('KMSCLIENT',des)<>0 then result:='KMS' else  
if pos('RETAIL(Free)',des)<>0 then result:='RETAIL(Free)' else  
if pos('TIMEBASED_SUB',des)<>0 then result:='TIMEBASED_SUB' else  
if pos('TIMEBASED_EVAL',des)<>0 then result:='TIMEBASED_EVAL' else  
if pos('KMS_W7',des)<>0 then result:='KMS' else  
if pos('KMS_R2_C',des)<>0 then result:='KMS' else  
if pos('RETAIL(MAK)',des)<>0 then result:='RETAIL(MAK)' else 
if pos('KMS_WS19',des)<>0 then result:='KMS' else  
if pos('KMS_2012-R2',des)<>0 then result:='KMS' else  
if pos('KMS_WS12_R2',des)<>0 then result:='KMS' else  
if pos('KMS_WS16',des)<>0 then result:='KMS' else  
if pos('VIRTUAL_MACHINE_ACTIVATION',des)<>0 then result:='VIRTUAL_MACHINE_ACTIVATION' else  
if pos('KMS_W8.1',des)<>0 then result:='KMS' else
if pos('GVLK',des)<>0 then result:='GVLK' else 
if pos('OEM',des)<>0 then result:='OEM' else
if pos('MAK',des)<>0 then result:='MAK' else
if pos('RETAIL',des)<>0 then result:='RETAIL' else
if pos('KMS',des)<>0 then result:='KMS'
else result:=des;
end;

function DescriptionLicStatus(num:string):string;
var   /// передаем код ошибки и из таблицы получаем описание этой ошибки
z:integer;
he,s,A:string;
begin
try
a:=num;
if pos('OLE error ',num)<>0 then
begin
s:=copy(num,11,length(num));
a:=s;
end;
//https://support.microsoft.com/ru-ru/windows/%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BA%D0%B0-%D0%BF%D0%BE-%D0%BE%D1%88%D0%B8%D0%B1%D0%BA%D0%B0%D0%BC-%D0%B0%D0%BA%D1%82%D0%B8%D0%B2%D0%B0%D1%86%D0%B8%D0%B8-windows-09d8fb64-6768-4815-0c30-159fa7d89d85
if TryStrToInt(a,z) then
begin
he:=IntToHex(z,1);
end
else he:=a;
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='SELECT * FROM LIC_ERROR WHERE CODE='''+he+'''';
FDQueryWrite.Open;
result:=FDQueryWrite.FieldByName('DESCRIPTION').AsString;
TransactionWrite.Commit;
if result='' then
begin // иначе записываем этот код в таблицу
result:=he;
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into LIC_ERROR (CODE,DESCRIPTION) VALUES (:p1,:p2) MATCHING (CODE,DESCRIPTION)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(he)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+('Unknown')+'';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
end;
except on E: Exception do
begin
 result:='Unknown';
 Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - DescriptionLicStatus - '+e.Message);
end;
end;
end;

function ReadNameProductMicrosoft(sID:string):string;  // чтение имени продукта
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='SELECT * FROM MICROSOFT_PRODUCT WHERE PRODUCT_ID='''+sID+'''';
FDQueryWrite.Open;
Result:=(FDQueryWrite.FieldByName('DESCRIPTION').AsString);
TransactionWrite.Commit;
except on E: Exception do
begin
 result:='';
 Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - ReadNameProductMicrosoft - '+e.Message);
end;
end;
end;
///////////////////////////////////////////////////////////////////////////////////////
function WriteNameIDProductMicrosoft(sID,sName,KeyProduct,TypeLic,PartProdKey,Product:string):boolean;  // Запись нового ID и Имени продукта и Ключа
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into MICROSOFT_PRODUCT'+
  ' (PRODUCT_ID,DESCRIPTION,KEY_PRODUCT,TYPE_LIC,PARTIALPRODUCTKEY,PRODUCT)'+
  ' VALUES (:p1,:p2,:p3,:p4,:p5,:p6) MATCHING (DESCRIPTION,KEY_PRODUCT,TYPE_LIC,PRODUCT_ID)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(sID)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+(sName)+'';
FDQueryWrite.params.ParamByName('p3').AsString:=''+(KeyProduct)+'';
FDQueryWrite.params.ParamByName('p4').AsString:=''+(TypeLic)+'';
FDQueryWrite.params.ParamByName('p5').AsString:=''+(PartProdKey)+'';
FDQueryWrite.params.ParamByName('p6').AsString:=''+(Product)+'';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
result:=true;
except on E: Exception do
begin
 result:=false;
 Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - WriteNameIDProductMicrosoft - '+e.Message);
end;
end;
end;
//////////////////////////////////////////////////////////////////////////////////////
function keyOEM:string;   ///// Windows 10
var
FWbemObjectSet                          : OLEVariant;
FWbemObject,KeyObject                   : OLEVariant;
oEnum                                   : IEnumvariant;
iValue        : LongWord;
yes:boolean;
Begin
try
yes:=false;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT OA3xOriginalProductKey FROM SoftwareLicensingService');
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
begin
if FWbemObject.OA3xOriginalProductKey<>'' then result:=vartostr(FWbemObject.OA3xOriginalProductKey)
else result:='';
yes:=true;
end;
if not yes then result:='';
except
 on E:Exception do
  begin
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - KeyOEM - '+e.Message);
  result:='';
  end;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
end;
///////////////////////////////////////////////////////////////////////////////////////
Procedure ViewLicensingOffice(var OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList);
  var
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
  i:integer;
  ProdKeyID:string;
const
 LicStatus: array [0..6] of string =('Нет лицензии','Лицензия активна',
  'Льготный период','Допустимое отклонение от льготного периода','Льготный период для неподлинных версий','Предупреждение','Расширенное льготное время');
Begin
try
 i:=0;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,Description,ProductKeyID,LicenseStatus,PartialProductKey,ID,LicenseStatusReason FROM OfficeSoftwareProtectionProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
  begin
  if (FWbemObject.ID<>null)and (FWbemObject.Name<>null)and(integer(FWbemObject.LicenseStatus)<>0) then
    begin
    OfficeProduct.add(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
    if OfficeProduct[i]='' then
    begin
    OfficeProduct[i]:=(vartostr(FWbemObject.Name));
    end;
    if FWbemObject.LicenseStatus<>null then StatOfLic.add(LicStatus[(integer(FWbemObject.LicenseStatus))]) else StatOfLic.add('');
    if FWbemObject.PartialProductKey<>null then KeyOfc.add( vartostr(FWbemObject.PartialProductKey)) else KeyOfc.add('');
    if FWbemObject.LicenseStatusReason<>null then DescriptLicOffice.Add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason))) else DescriptLicOffice.Add('');
    if FWbemObject.Description<>null then  TypelicOffice.Add(TypeLicinAndOffice(vartostr(FWbemObject.Description))) else TypelicOffice.Add('');
    if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
     WriteNameIDProductMicrosoft((vartostr(FWbemObject.ID)),OfficeProduct[i],KeyOfc[i],TypelicOffice[i],ProdKeyID,'Office'); // запись в общую таблицу продуктов
    inc(i);
    end;
  end;

except
 on E:Exception do
  begin
   Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - ViewLicensingOffice - '+e.Message);
   VariantClear(FWbemObject);
   oEnum:=nil;
   VariantClear(FWbemObjectSet);
  end;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
End;

///////////////////////////////////////////////////////////////////////////////////////
procedure LicensingStatus(var WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptLicWin:string;OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList);
  var
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
  i:integer;
  OfficeBool:boolean;
  OEMKey,ProdKeyID:string;
const
 LicStatus: array [0..6] of string =('Нет лицензии','Лицензия активна',
  'Льготный период','Допустимое отклонение от льготного периода','Льготный период для неподлинных версий','Предупреждение','Расширенное льготное время');

function Win7:boolean;
begin
if pos('6.1',VerOS)=1 then result:=true
else Result:=false;
end;
function OfficeBoll(s:string):boolean;
begin
if pos('Office',s)<>0 then Result:=true
else Result:=false;
end;
function WindowsBoll(s:string):boolean;
begin
if pos('Windows',s)<>0 then Result:=true
else Result:=false;
end;
Begin
try
i:=0;
OfficeBool:=false;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,Description,LicenseStatus,ProductKeyID,PartialProductKey,ID,LicenseStatusReason FROM SoftwareLicensingProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
  if integer(FWbemObject.LicenseStatus)<>0 then
 BEGIN
   if WindowsBoll(vartostr(FWbemObject.Name)) then
   Begin
    if (FWbemObject.ID<>null) then
     begin
     WinProduct:=(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
     if WinProduct='' then WinProduct:=nameOS+' ('+VerOS+')';
     if FWbemObject.LicenseStatus<>null then StatWinLic:=LicStatus[(integer(FWbemObject.LicenseStatus))];
     if FWbemObject.PartialProductKey<>null then KeyWin:=vartostr(FWbemObject.PartialProductKey);
     if FWbemObject.LicenseStatusReason<>null then DescriptLicWin:=DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason));
     if FWbemObject.Description<>null then  TypelicWin:=TypeLicinAndOffice(vartostr(FWbemObject.Description));
     if pos('OEM',TypelicWin)<>0 then //если тип лицензии OEM то пытаемся получить ключ из UEFI
     begin
      OEMKey:=keyOEM;
      if Length(OEMKey)>5 then KeyWin:=OEMKey;
     end;
     if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
     WriteNameIDProductMicrosoft(vartostr(FWbemObject.ID),WinProduct,KeyWin,TypelicWin,ProdKeyID,'Windows'); // запись в общую таблицу продуктов
     end;
   End;
   if OfficeBoll(vartostr(FWbemObject.Name)) then
   Begin
    if (FWbemObject.ID<>null) then
    begin
     OfficeProduct.add(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
      if OfficeProduct[i]='' then
      begin
       OfficeProduct[i]:=(vartostr(FWbemObject.Name));
      end;
    if FWbemObject.LicenseStatus<>null then StatOfLic.add(LicStatus[(integer(FWbemObject.LicenseStatus))]);
    if FWbemObject.PartialProductKey<>null then KeyOfc.add( vartostr(FWbemObject.PartialProductKey));
    if FWbemObject.LicenseStatusReason<>null then DescriptLicOffice.Add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason))) else DescriptLicOffice.Add('');
    if FWbemObject.Description<>null then  TypelicOffice.Add(TypeLicinAndOffice(vartostr(FWbemObject.Description))) else TypelicOffice.Add('');
    if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
     WriteNameIDProductMicrosoft((vartostr(FWbemObject.ID)),OfficeProduct[i],KeyOfc[i],TypelicOffice[i],ProdKeyID,'Office'); // запись в общую таблицу продуктов
    inc(i);
    end;
    OfficeBool:=true;
   End;
 End;

if not OfficeBool then
begin
OfficeProduct.Clear;
StatOfLic.Clear;
KeyOfc.Clear;
ViewLicensingOffice(OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice);
end;


except
 on E:Exception do
  begin
  Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - LicensingStatus - '+e.Message);
  end;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
End;

Procedure osconfig(var nameOS,VerOS:string);
var
FWbemObjectSet                          : OLEVariant;
FWbemObject                             : OLEVariant;
oEnum                                   : IEnumvariant;
iValue        : LongWord;
begin
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,Version FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Операционная система
    begin
      try
      if FWbemObject.Caption<>null then
      begin
       nameOS:=trim(string(FWbemObject.Caption));
      end;
      except
      nameOS:='unknown';
      end;
      try
      if FWbemObject.version<>null then   VerOS:=trim(string(FWbemObject.version));
      except
      VerOS:= 'unknown';
      end;
      FWbemObject:=Unassigned;
    end;
  FWbemObject:=Unassigned;
 Except
  on E:Exception do
     begin
     Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - OSconfig - '+e.Message);
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet)
     end;
   end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
end;

function ConnectWMI(NamePC,User,Pass:string): Boolean;
begin
 try
    FSWbemLocator:=Unassigned;
    FWMIService:=Unassigned;
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128); ///WbemUser, WbemPassword
    FWMIService.Security_.impersonationlevel:=3;
    FWMIService.Security_.authenticationLevel := 6;
    Result:=True;
 except
    on E:EOleException do
    begin
    result:=false;
    Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+NamePC+' - '+Format(' Ошибка ConnectWMI %s %x', [E.Message,E.ErrorCode]));
    end;
 end;
end;
//////////////////////////////// Ниже начинается основное тело потока из которого будут вызваны необходимые функции
BEGIN
try
Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - Запуск сканирования');
MicrosoftProductP:=MicProduct;
ListPC:=TStringList.Create;
ListPC:=MicrosoftProductP.NamePC;
user:=MicrosoftProductP.User;
Pass:=MicrosoftProductP.Pass;
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // расположение базы данных
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.Connected:=true;
ConnectionThRead.LoginPrompt:= false;  /// отображение диалога user password
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThRead;
TransactionWrite.Options.DisconnectAction:=xdCommit; /// или  xdCommit
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
TransactionWrite.Options.AutoCommit:=false;
TransactionWrite.Options.AutoStart:=false;
TransactionWrite.Options.AutoStop:=false;
FDQueryWrite:=TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=ConnectionThRead;
try
InventMicrosoft:=true;// признак запуска инвентаризации продуктов мелкософта
OleInitialize(nil);
SolveExitInvMicrosoft:=false; // признак того что инвентаризация запущена
CoutPCOK:=0;
for I := 0 to ListPC.Count-1 do
BEGIN
if not InventMicrosoft then break // если остановили инвентаризацию то выходим
else
if (pingWin(ListPC[i])) and (ConnectWMI(ListPC[i],User,Pass)) then
Begin
CurrentPC:=ListPC[i];
nameOS:='';
VerOS:='';
osconfig(nameOS,VerOS); // определение версии ОС
WinProduct:='';
StatWinLic:='';
KeyWin:='';
TypelicWin:='';
DescriptWin:='';
sOfficeProduct:='';
sStatOfLic:='';
sKeyOfc:='';
sTypelicOffice:='';
sDescriptLicOffice:='';
OfficeProduct:=TstringList.Create;
StatOfLic:=TstringList.Create;
KeyOfc:=TstringList.Create;
TypelicOffice:=TStringList.Create;
DescriptLicOffice:=TStringList.Create;
try
LicensingStatus(WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptWin,OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice); // получаем информацию по лицензиям

sOfficeProduct:=OfficeProduct.CommaText;
sStatOfLic:=StatOfLic.CommaText;
sKeyOfc:=KeyOfc.CommaText;
sTypelicOffice:=TypelicOffice.CommaText;
sDescriptLicOffice:=DescriptLicOffice.CommaText;
finally
OfficeProduct.Free;
StatOfLic.Free;
KeyOfc.Free;
TypelicOffice.Free;
DescriptLicOffice.Free;
end;
  try
  TransactionWrite.StartTransaction;
  FDQueryWrite.SQL.Clear;                                                              /// ERROR_INV Зачищаем ошибки прошлой инвентаризации
  FDQueryWrite.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
  +' MAIN_PC (PC_NAME,STATWINLIC,SSTATOFLIC) VALUES'
  +' ('''+ListPC[i]+''','''+leftstr(STATWINLIC,50)+''','''+leftstr(SSTATOFLIC,500)+''')'
  +' MATCHING (PC_NAME)';
  FDQueryWrite.ExecSQL;
  TransactionWrite.commit;
  except on E: Exception do
  begin
  Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+ListPC[i]+' - MAIN_PC - '+e.Message);
  TransactionWrite.commit;
  end;
  end;

  try
  TransactionWrite.StartTransaction;
  FDQueryWrite.SQL.Clear;                                                              /// ERROR_INV Зачищаем ошибки прошлой инвентаризации
  FDQueryWrite.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
  +' MICROSOFT_LIC (NAMEPC,WINPRODUCT,STATWINLIC,KEYWIN,TYPE_LIC_WIN,DESCRIPT_LIC_WIN,SOFFICEPRODUCT,SSTATOFLIC,SKEYOFC,TYPE_LIC_OFFICE,DESCRIPT_LIC_OFFICE) VALUES'
  +' ('''+ListPC[i]+''','''+leftstr(WINPRODUCT,250)+''','''+leftstr(STATWINLIC,100)+''','''+leftstr(KEYWIN,50)+''','''+leftstr(TypelicWin,50)+''','''+leftstr(DescriptWin,250)+''','''+leftstr(SOFFICEPRODUCT,500)+''','''+leftstr(SSTATOFLIC,500)+''','''+leftstr(SKEYOFC,200)+''','''+leftstr(sTypelicOffice,150)+''','''+leftstr(sDescriptLicOffice,1500)+''')'
  +' MATCHING (NAMEPC)';
  FDQueryWrite.ExecSQL;
  TransactionWrite.commit;
  inc(CoutPCOK);
  except on E: Exception do
  begin
  Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+ListPC[i]+' - MICROSOFT_LIC - '+e.Message);
  TransactionWrite.commit;
  end;
  end;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:=ListPC[i]+' - OK';
frmDomainInfo.StatusInvMicrosoft.Panels[2].Text:=inttostr(ListPC.Count)+'/'+IntToStr(i+1)+'/'+inttostr(CoutPCOK)+' - OK';
if not InventMicrosoft then break; // если остановили инвентаризацию то выходим
End // если соединение установлено
else // иначе информацию в statusBar
  begin
  frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:=ListPC[i]+' - не доступен';
  frmDomainInfo.StatusInvMicrosoft.Panels[2].Text:=inttostr(ListPC.Count)+'/'+IntToStr(i+1)+'/'+inttostr(CoutPCOK)+' - OK';
  end;
END;// цикл

finally
frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:='Инвентаризация завершена';
FDQueryWrite.free;
TransactionWrite.Free;
ListPC.Free;
ConnectionThRead.Close;
ConnectionThRead.Free;
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
SolveExitInvMicrosoft:=true; // призна того что этот поток заверился
InventMicrosoft:=false;
NetApiBufferFree(MicProduct); /// очищаем память
EndThread(0); /// конец потока
end;


except on E: Exception do 
begin
Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+CurrentPC+' - Общая ошибка - '+e.Message);
frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:='Ошибка инвентаризации';
SolveExitInvMicrosoft:=true; // призна того что этот поток заверился
InventMicrosoft:=false;
NetApiBufferFree(MicProduct); /// очищаем память
EndThread(0); /// конец потока
end;
end;

END;







end.
