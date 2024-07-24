unit RunTaskOsher;

interface
uses Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,ShellAPI,
  ActiveX,ComObj, FireDAC.Stan.Option, FireDAC.Comp.Client,VCL.Forms;

  FUNCTION RestoreOrNewPoint(namePC,user,pass,Description:string;operation:integer):integer; // функция запуска или создания точки восстановления

implementation

{except
  on E: EOLEException do
  begin
    if E.ErrorCode = HRESULT(wbemErrInvalidQuery) then
    begin
      // do something
    end;
  end;
end;}
FUNCTION RestoreOrNewPoint(namePC,user,pass,Description:string;operation:integer):integer;
var
FSWbemLocatorRE : OLEVariant;
FWMIServiceRE   : OLEVariant;
ResOut:integer;
const
wbemFlagForwardOnly = $00000020;

function Log_write(Level:integer;fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'\Restore.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'\Restore.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'\Restore.log');
      finally
        f.Destroy;
      end;
except
exit;
end;
end;

function CloseConnectWMI:boolean;
begin
  try
  FSWbemLocatorRE:=Unassigned;
  VariantClear(FSWbemLocatorRE);
  FWMIServiceRE:=Unassigned;
  VariantClear(FWMIServiceRE);
  result:=true;
  except
   result:=false;
  end;
end;

function ConnectWMI(NamePC,User,Pass:string):integer;
begin
 try
    FSWbemLocatorRE := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIServiceRE   := FSWbemLocatorRE.ConnectServer(NamePC, 'root\DEFAULT', User, Pass,'','',128);
    FWMIServiceRE.Security_.impersonationlevel:=3;
    FWMIServiceRE.Security_.authenticationLevel := 6;
    Result:=0;
 except
   on E:Exception do
   begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка подключения к компьютеру '+namePC+' (1): '+e.Message);
     result:=666;
     end;
   end;
end;

Function RestoreSelectedPoint(SequenceNumber:integer):integer; // запуск восстановления системы.
var                     // вроде как надо запустить перезагрузку после этой операции
 FWbemObjectSet  : OLEVariant;
begin
 try
 FWbemObjectSet:= FWMIServiceRE.Get('SystemRestore');
 result:=FWbemObjectSet.Restore(SequenceNumber);
 except
 on e:EInOutError do result:=e.ErrorCode;
 end;
 VariantClear(FWbemObjectSet);
end;

function  WbemTimeToDateTime(vDate : OleVariant) : string;  // перево даты и времени
var
  FWbemDateObj  : OleVariant;
  TDT:TdateTime;
begin;
  FWbemDateObj  := CreateOleObject('WbemScripting.SWbemDateTime');
  try
  FWbemDateObj.Value:=vDate;
  TDT:=FWbemDateObj.GetVarDate;
  result:=DateToStr(TDT);
  finally
  VariantClear(FWbemDateObj);
  end;
end;

function LoadPointsRestoreWin(DatePointRestore,Description:string):integer; // загрузка точек восстановления с последующим запуском восстановления до точки
var
 FWbemObjectSet  : OLEVariant;
 FWbemObject     : OLEVariant;
 oEnum           : IEnumvariant;
 iValue          : LongWord;
 pointFind       :boolean;
begin
 try
 pointFind:=false;
 FWbemObjectSet:= FWMIServiceRE.ExecQuery('SELECT * FROM SystemRestore','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// профили
Begin
 if (FWbemObject.SequenceNumber<>null) and (FWbemObject.CreationTime<>null) then
  Begin
   if Description='' then
   begin
   if  DatePointRestore=WbemTimeToDateTime(FWbemObject.CreationTime) then   // если дата точки совпадает с необходимой то восстанавливаем.
     begin
     pointFind:=true;
     result:=RestoreSelectedPoint(integer(FWbemObject.SequenceNumber));
     break;
     end;
   end
   else
   if FWbemObject.Description<>null then
   begin
   if  (DatePointRestore=WbemTimeToDateTime(FWbemObject.CreationTime)) and (vartostr(FWbemObject.Description)=Description) then   // если дата и описание точки совпадает с необходимой то восстанавливаем.
     begin
     pointFind:=true;
     result:=RestoreSelectedPoint(integer(FWbemObject.SequenceNumber));
     break;
     end;
   end;
  End;
VariantClear(FWbemObject);
FWbemObject:=Unassigned;
End;
except
 on E:Exception do
  begin
  Log_write(1,'TASK',timetostr(now)+ ' - Ошибка загрузки точек восстановления '+namePC+' (2): '+e.Message);
  result:=999;
  end;
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 FWbemObjectSet:=Unassigned;
 VariantClear(FWbemObjectSet);
if not pointFind then  result:=777; // иначе точка для восстановления не найддена
end;


function EnableRestore(nameDisk:string):integer;   // включение восстановления на диске
var
FWbemObjectSet  : OLEVariant;
begin
try
FWbemObjectSet:= FWMIServiceRE.Get('SystemRestore');
result:=FWbemObjectSet.Enable(nameDisk);
FWbemObjectSet:=Unassigned;
except
   on E:Exception do
   begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка включения восстановления на диске '+nameDisk+' - '+namePC+' (3): '+e.Message);
     result:=1111;
     end;
   end;
VariantClear(FWbemObjectSet);
end;

Function CreateNewPointRestore(Description:string;RestorePointType,EventType:integer):integer; // создание новой точки восстановления
var
 FWbemObjectSet  : OLEVariant;
 res:integer;
 begin
try
 FWbemObjectSet:= FWMIServiceRE.Get('SystemRestore');
 res:=FWbemObjectSet.CreateRestorePoint(Description,RestorePointType,EventType);
 result:=res;
 except
 on E:Exception do
   begin
   Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания точки восстановления '+namePC+' (4): '+e.Message);
   result:=2222;
   end;
   end;
 FWbemObjectSet:=Unassigned;
 VariantClear(FWbemObjectSet);
 end;

function StrInDescription(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=br;
end;

function StrCopySourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=copy(br,1,pos(' -> ',br)-1);
end;

function StrCopyDestination(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(' -> ',br)+3);
result:=br;
end;

////////////////////// по очереди запускаем необходимые функции
BEGIN //
try
OleInitialize(nil);
ResOut:=ConnectWMI(namePC,user,pass); // установка соединения с WMI
try
if ResOut=0 then // если соединение установлено
Begin
 if operation=1 then // создать точку восстановления
 begin
 ResOut:=EnableRestore('C:\'); // включение восстановления
 if ResOut=0 then // если восстановление включено то создаем точку восстановления
  begin
  ResOut:=CreateNewPointRestore(StrInDescription(Description),0,100);
  end;
 End;
////////////////////////////////////////////////////////////
if operation=2 then // запустить восстановление по дате
 Begin
 ResOut:=EnableRestore('C:\'); // включение восстановления если оно выключено
 if ResOut=0 then // если восстановление включено то пытаемся восстановить до точки
  begin
  ResOut:=LoadPointsRestoreWin(StrInDescription(Description),'');
  end;
 End;
///////////////////////////////////////////////////////////

if operation=3 then // запустить восстановление по дате и описанию точки восстановления
 Begin
 ResOut:=EnableRestore('C:\'); // включение восстановления если оно выключено
 if ResOut=0 then // если восстановление включено то пытаемся восстановить до точки
  begin
  ResOut:=LoadPointsRestoreWin(StrCopySourse(Description),StrCopyDestination(Description)); //первая переменная это дата, вторая это описание
  end;
 End;
End;
result:=ResOut; // иначе соединение не установлено
finally
CloseConnectWMI;
OleUnInitialize;
end;
 except
 on E:Exception do
   begin
   Log_write(2,'TASK',timetostr(now)+ ' - Общая ошибка при работе с точками восстановления '+namePC+' (5): '+e.Message);
   result:=888;
   end;
 end;
END;
///////////////////////////////////////////////////////////////////////////////////////////
//******************* окончание функций восстановления точек доступа********************//



end.
