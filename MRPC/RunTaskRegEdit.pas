unit RunTaskRegEdit;

interface
uses Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,ShellAPI,
  ActiveX,ComObj, FireDAC.Stan.Option, FireDAC.Comp.Client,VCL.Forms;

function ConnectWMI(NamePC,User,Pass:string):integer;
function CreatNewRazdel(hDefKey:integer;sSubKeyName:string):integer;  // создает раздел в указаном подразделе
function DeleteRazdel(hDefKey:integer;sSubKeyName:string):integer;  // удаляет раздел в выеленном подразделе
function DeleteKeyValue(hDefKey:integer;sSubKeyName,sValueName:string):integer;  // удаляет ключ в реестре
function CreateSetStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;  // создание строкового ключа реестра
function CreateBinarySaveKey(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // импортирование бинарного ключа в реестре
function CreateSetDWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:integer):integer; // создание ключа DWORD
function CreateSetQWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:int64):integer; //создание ключа QWORD
function CreateSetMultiStringValue(hDefKey:integer;sSubKeyName,sValueName:string;sValue:TStringList):integer; //новый Мультистринговый ключ
function CreateSetExpandedStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;   // расширенный строковый параметр
function ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;  /// импорт ключей в реестр из задания
Function ReadTabImportKey(idKey:integer):integer; // импорт ключей из таблицы и запуск функции ImportSaveKeyhDefKey
function RunEditRegeditWindows(NamePC,User,Pass,TYPETASK,NUMTASK,NAMETASK:string):integer; // функция получает данные и в соответствии с TYPETASK раскидывает по функциям

implementation
uses RunTask;


{except
  on E: EOLEException do
  begin
    if E.ErrorCode = HRESULT(wbemErrInvalidQuery) then
    begin
      // do something
    end;
  end;
end;}
var
FSWbemLocatorRE : OLEVariant;
FWMIServiceRE   : OLEVariant;

function Log_write(Level:integer;fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'\regedit.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'\regedit.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'\regedit.log');
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
  FWMIServiceRE:=Unassigned;
  result:=true;
  except on E: Exception do result:=false;
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
    VariantClear(FSWbemLocatorRE);
    VariantClear(FWMIServiceRE);
    Log_write(1,'TASK',namePC+' - RPС сервер не доступен (21)');
    result:=666;
 end;
end;


function CreatNewRazdel(hDefKey:integer;sSubKeyName:string):integer;  // создает раздел в выеленном подразделе
Var
FWbemObjectSet: OLEVariant;
res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.CreateKey(hDefKey,sSubKeyName);
  VariantClear(FWbemObjectSet);
  result:=res;
  except
    on E: Exception do
    begin
     Log_write(1,'TASK',' - Ошибка создания раздела реестра (31): '+e.Message);
     result:=631;
    end;

  end;
end;

function DeleteRazdel(hDefKey:integer;sSubKeyName:string):integer;  // удаляет раздел в выеленном подразделе
Var
FWbemObjectSet: OLEVariant;
res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.DeleteKey(hDefKey,sSubKeyName);
  VariantClear(FWbemObjectSet);
  result:=res;
  except
    on E: Exception do
    begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка удаления раздела реестра (32): '+e.Message);
     result:=632;
    end;
    end;
end;


function DeleteKeyValue(hDefKey:integer;sSubKeyName,sValueName:string):integer;  // удаляет ключ в реестре
Var
FWbemObjectSet: OLEVariant;
res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.DeleteValue(hDefKey,sSubKeyName,sValueName);
  VariantClear(FWbemObjectSet);
  result:=res;
  except
    on E: Exception do
    begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка удаления ключа реестра (33): '+e.Message);
    result:=633;
    end;
  end;
end;


function CreateSetStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;  // сщздание строкового ключа реестра
var                   // строковый параметр
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.SetStringValue(hDefKey,sSubKeyName,sValueName,sValue);
  if res=2 then //не удается найти указанный файл - т.е. если раздел для ключа не найден
   begin    // пытаемся создать раздел
   res:=CreatNewRazdel(hDefKey,sSubKeyName);
     if res=0 then // если раздел успешно создан
     begin
     VariantClear(FWbemObjectSet); // очищаем и пытаемся создать повторно
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     res:=FWbemObjectSet.SetStringValue(hDefKey,sSubKeyName,sValueName,sValue);
     end;
   end;
  VariantClear(FWbemObjectSet);
  result:=res;
except
    on E: Exception do
    begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания ключа реестра (37): '+e.Message);
     result:=637;
    end;
  end;
  end;

function CreateBinarySaveKey(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // импортирование бинарного ключа в реестре
var
FWbemObjectSet: OLEVariant;
i,z,OutParam:integer;
mas:variant;
res:TstringList;
newstr:string;
function MemoHexToInt(StrBin:string):string;
var
i,z:integer;res:String;newstr:string;
begin
try
res:=''; i:=0;
while StrBin<>'' do
  begin
  newstr:='';
  if pos(' ',StrBin)=1 then delete(StrBin,1,1);
  if StrBin<>'' then
    begin
    inc(i);
    newstr:=copy(StrBin,1,2);
    if TryStrToInt('$'+newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res:=res+inttostr(z)+' ';
    delete(StrBin,1,2);
    end;
  end;
result:=res;
except on E: Exception do result:='';
end;
end;
begin
try
res:=TStringList.Create;
i:=0;
if HexOrInt='HEX' then // если числа были 16ти ричные то получаем чистый integer
begin
FullStr:=MemoHexToInt(FullStr);
end;
if FullStr[length(FullStr)]<>' ' then // если в конце строки нет пробела то добавляем его
Fullstr:=Fullstr+' ';
while FullStr<>'' do /// при условии
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    newstr:=copy(FullStr,1,pos(' ',FullStr)-1);
    if TryStrToInt(newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res.Add(inttostr(z));
    delete(FullStr,1,pos(' ',FullStr));
    end;
  end;
   mas:=VarArrayCreate([0,res.Count-1],varByte);
   for I := 0 to res.Count-1 do
   mas[i]:=strtoint(res[i]);
  res.Free;
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  OutParam:=FWbemObjectSet.SetBinaryValue(hDefKey,sSubKeyName,sValueName,mas);
   if OutParam=2 then //не удается найти указанный файл - т.е. если раздел для ключа не найден
   begin    // пытаемся создать раздел
   OutParam:=CreatNewRazdel(hDefKey,sSubKeyName);
   if OutParam=0 then // если раздел успешно создан
     begin
     VariantClear(FWbemObjectSet); // очищаем и пытаемся создать повторно
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     OutParam:=FWbemObjectSet.SetBinaryValue(hDefKey,sSubKeyName,sValueName,mas);
     end;
   end;
  VariantClear(FWbemObjectSet);
  VarClear(mas);
  result:=OutParam;
except
    on E: Exception do
    begin
    VariantClear(FWbemObjectSet);
    VarClear(mas);
    Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания ключа реестра (46): '+e.Message);
    result:=646;
    end;
  end;
  end;

function CreateSetDWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:integer):integer; // создание ключа DWORD
var               // новый DWORD (int32) ключ
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.SetDWORDValue(hDefKey,sSubKeyName,sValueName,uValue);
   if res=2 then //не удается найти указанный файл - т.е. если раздел для ключа не найден
   begin    // пытаемся создать раздел
   res:=CreatNewRazdel(hDefKey,sSubKeyName);
   if res=0 then // если раздел успешно создан
     begin
     VariantClear(FWbemObjectSet); // очищаем и пытаемся создать ключ повторно
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     res:=FWbemObjectSet.SetDWORDValue(hDefKey,sSubKeyName,sValueName,uValue);
     end;
   end;
  VariantClear(FWbemObjectSet);
  result:=res;
except
    on E: Exception do
    begin
     VariantClear(FWbemObjectSet);
    Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания ключа реестра (39): '+e.Message);
    result:=639;
    end;
  end;
  end;


function CreateSetQWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:int64):integer; //создание ключа QWORD
var               // новый QWORD (int64) ключ
FWbemObjectSet: OLEVariant;
InParam: OLEVariant;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  InParam:= FWbemObjectSet.Methods_.Item('SetQWORDValue').InParameters.SpawnInstance_();
  InParam.hDefKey:=hDefKey;  // root
  InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь для создаваемого раздела
  InParam.sValueName:=sValueName;
  InParam.uValue:=uValue;
  FWMIServiceRE.ExecMethod('StdRegProv', 'SetQWORDValue', InParam);
  VariantClear(InParam);
  VariantClear(FWbemObjectSet);
  result:=0;
except
    on E: Exception do
    begin
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания ключа реестра (40): '+e.Message);
    result:=640;
    end;
  end;
  end;


function CreateSetMultiStringValue(hDefKey:integer;sSubKeyName,sValueName:string;sValue:TStringList):integer; //новый Мультистринговый ключ
var
FWbemObjectSet: OLEVariant;
res,i:integer;
mas:variant;
begin
try
mas:=VarArrayCreate([0,sValue.count-1],varOleStr);
for I := 0 to sValue.count-1 do mas[i]:=sValue[i];
FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
res:=FWbemObjectSet.SetMultiStringValue(hDefKey,sSubKeyName,sValueName,mas);
 if res=2 then //не удается найти указанный файл - т.е. если раздел для ключа не найден
   begin    // пытаемся создать раздел
   res:=CreatNewRazdel(hDefKey,sSubKeyName);
   if res=0 then // если раздел успешно создан
     begin
     VariantClear(FWbemObjectSet); // очищаем и пытаемся создать ключ повторно
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     res:=FWbemObjectSet.SetMultiStringValue(hDefKey,sSubKeyName,sValueName,mas);
     end;
   end;
VariantClear(FWbemObjectSet);
VarClear(mas);
result:=res;
except
    on E: Exception do
    begin
     VariantClear(FWbemObjectSet);
     VarClear(mas);
    Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания ключа реестра (41): '+e.Message);
    result:=641;
    end;
  end;
  end;


function CreateSetExpandedStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;   // расширенный строковый параметр
var
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  res:=FWbemObjectSet.SetExpandedStringValue(hDefKey,sSubKeyName,sValueName,sValue);
   if res=2 then //не удается найти указанный файл - т.е. если раздел для ключа не найден
   begin    // пытаемся создать раздел
   res:=CreatNewRazdel(hDefKey,sSubKeyName);
   if res=0 then // если раздел успешно создан
     begin
     VariantClear(FWbemObjectSet); // очищаем и пытаемся создать ключ повторно
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     res:=FWbemObjectSet.SetExpandedStringValue(hDefKey,sSubKeyName,sValueName,sValue);
     end;
   end;
  VariantClear(FWbemObjectSet);
  result:=res;
except
    on E: Exception do
    begin
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка создания расширенного ключа реестра (42): '+e.Message);
     result:=642;
    end;

  end;
  end;



function StringInGetRooT(Node:string):integer; // передаем имя раздела, получаем значение в Integer
var
val:integer;
begin
val:=0;
try
if node='HKEY_CLASSES_ROOT' then val:=HKEY_CLASSES_ROOT;
if node='HKEY_CURRENT_USER' then val:= HKEY_CURRENT_USER;
if node='HKEY_LOCAL_MACHINE' then val:= HKEY_LOCAL_MACHINE;
if node='HKEY_USERS' then val:= HKEY_USERS;
if node='HKEY_CURRENT_CONFIG' then val:= HKEY_CURRENT_CONFIG;
result:=val;
except
on E: Exception do begin result:=val; end;
end;
end;

function StrTointTypeKey(TypeKey:string):integer; //передаем тип ключа в string получаем его в integer
begin
try
if TypeKey='REG_SZ' then result:=1;
if TypeKey='REG_EXPAND_SZ' then result:=2;
if TypeKey='REG_BINARY' then result:=3;
if TypeKey='REG_DWORD' then result:=4;
if TypeKey='REG_MULTI_SZ' then result:=7;
if TypeKey='REG_QWORD' then result:=11;
except on E: Exception do begin result:=0; end;
end;
end;


function ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;  /// импорт ключей в реестр из задания
var
i,IntTypeKey:integer;
listT:TstringList;
function StringToList(sValue:string):Tstringlist; // из строки делаем список строк
var
s:string;
begin
result:=TStringList.Create;
result.Text:=sValue;
end;
function StrQDWord(s:string):string; // из строки со значением HEX и Integer выделяем только Int
begin
delete(s,1,pos('(',s));
result:=Copy(s,1,pos(')',s)-1);
end;
begin
try
IntTypeKey:=StrTointTypeKey(TypeKey);
case IntTypeKey of
   1:result:=CreateSetStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   2:result:=CreateSetExpandedStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   3:result:=CreateBinarySaveKey(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue,'HEX');
   4:result:=CreateSetDWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint(StrQDWord(sValue)));
   7:begin
    try
     listT:=TStringList.Create;
     listT:=StringToList(sValue);
     result:=CreateSetMultiStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,StringToList(sValue));
     finally
     listT.Free;
     end;
     end;
   11:result:=CreateSetQWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint64(StrQDWord(sValue)));
end;
except
on E: Exception do
begin
 Log_write(1,'TASK',timetostr(now)+ ' - Ошибка запуска импорта ключей на машины назначения (43)'+e.Message);
 result:=643;
 end;
end;
end;



Function ReadTabImportKey(idKey:integer):integer;  //
var
FDQueryRun: TFDQuery;
TransactionRun    : TFDTransaction;
begin
try
TransactionRun:= TFDTransaction.Create(nil);
TransactionRun.Connection:=ConnectionThReadTask;
TransactionRun.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRun.Options.ReadOnly:=true;
FDQueryRun:= TFDQuery.Create(nil);
FDQueryRun.Transaction:=TransactionRun;
FDQueryRun.Connection:=ConnectionThReadTask;
TransactionRun.StartTransaction;
FDQueryRun.sql.Text:='SELECT * FROM REGEDIT_KEY WHERE ID_KEY=:n';
FDQueryRun.ParamByName('n').asInteger:=idKey;
FDQueryRun.Open;
if FDQueryRun.FieldByName('SSUBKEYNAME').AsString<>'' then
begin
result:=ImportSaveKeyhDefKey(
                     FDQueryRun.FieldByName('HDEFKEY').AsString, // root раздел реестра
                     FDQueryRun.FieldByName('SSUBKEYNAME').AsString, // полный путь к ключу реестра
                     FDQueryRun.FieldByName('SVALUENAME').AsString,  // имя ключа реестра
                     FDQueryRun.FieldByName('SVALUE').AsString,      // значение ключа реестра
                     FDQueryRun.FieldByName('TYPEKEY').AsString );   // тип ключа реестра
end;
TransactionRun.Commit;
FDQueryRun.SQL.clear;   //очистить
FDQueryRun.Close;  /// закрыть нах после чтения
FDQueryRun.Free;
TransactionRun.Free;
Except
  on E:Exception do
     begin
     if Assigned(FDQueryRun) then FDQueryRun.Free;
     if Assigned(TransactionRun) then
     begin
     TransactionRun.Commit;
     TransactionRun.Free;
     end;
     Log_write(1,'TASK',timetostr(now)+ ' - Ошибка импорта ключей из таблицы (44)'+e.Message);
     result:=644;
     end;
    end;
end;

function strtohDefKey(s:string):string;
begin //Удалить ключ :HKEY_LOCAL_MACHINE:EEEEEEEEEEEEE:Имя ключа реестра
delete(s,1,pos(':',s));
result:=copy(s,1,pos(':',s)-1);
end;

function strtosSubKeyName(s:string):string;
begin
delete(s,1,pos(':',s));
delete(s,1,pos(':',s));
if pos(':',s)<>0 then result:=copy(s,1,pos(':',s)-1)
else result:=copy(s,1,length(s));
end;

function strtosNameKey(s:string):string;
begin
delete(s,1,pos(':',s));
delete(s,1,pos(':',s));
delete(s,1,pos(':',s));
result:=s;
end;

function RunEditRegeditWindows(NamePC,User,Pass,TYPETASK,NUMTASK,NAMETASK:string):integer; // функция получает данные и в соответствии с TYPETASK раскидывает по функциям
var
res:integer;
begin
try
OleInitialize(nil);
res:=ConnectWMI(NamePC,User,Pass);
if res=0 then // если установили соединение с компьютером
Begin
 if TYPETASK='CreateKey' then
 begin
 res:=ReadTabImportKey(strtoint(NUMTASK)); // импорт ключа реестра на компьютер из таблицы сохраненных ключей
 end;

 if TYPETASK='sNameKey' then  // удаление ключа реестра
 begin
  res:=DeleteKeyValue(StringInGetRooT(strtohDefKey(NAMETASK)), // передаем root раздел
                 strtosSubKeyName(NAMETASK),              // раздел ключа реестра
                 strtosNameKey(NAMETASK));                // имя ключа реестра
 Log_write(0,'TASK',timetostr(now)+ ' Удаление ключа реестра '+strtohDefKey(NAMETASK)+':'+strtosSubKeyName(NAMETASK)+':'+strtosNameKey(NAMETASK));
 end;

 if TYPETASK='sSubKey' then  // создание или удаление раздела реестра
  begin
   if strtoint(NUMTASK)=0 then
   begin
   res:=DeleteRazdel(StringInGetRooT(strtohDefKey(NAMETASK)), strtosSubKeyName(NAMETASK));//удаляем раздел
   Log_write(0,'TASK',timetostr(now)+ ' Удаление раздела '+strtohDefKey(NAMETASK)+':'+strtosSubKeyName(NAMETASK));
   end;
   if strtoint(NUMTASK)=1 then
   begin
   res:=CreatNewRazdel (StringInGetRooT(strtohDefKey(NAMETASK)), strtosSubKeyName(NAMETASK)); // создаем раздел
   Log_write(0,'TASK',timetostr(now)+ ' Создание раздела '+strtohDefKey(NAMETASK)+':'+strtosSubKeyName(NAMETASK));
   end;
   end;
End;
CloseConnectWMI; // закрываем установленное соединение
OleUnInitialize;
result:=res;
Except
  on E:Exception do
  begin
  Log_write(2,'TASK',timetostr(now)+ ' - Общая ошибка запуска функций работы с реестром: '+e.Message);
  result:=888;
  end;

end;
end;

end.
