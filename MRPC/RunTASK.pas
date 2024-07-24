unit RunTASK;

interface

uses
  System.Classes,IdICMPClient, system.sysutils,ActiveX,ComObj,CommCtrl,system.Variants,VCL.Forms
  ,ShellAPI,Winapi.Windows,FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet,system.StrUtils,IdUDPClient,idGlobal,inifiles,vcl.ExtCtrls,DateUtils,FireDAC.Phys.FB;

type




  TaskRun = class(TThread)
  private
    function createQuery:boolean;
    function ReadAndRunMsiOrProc (numberProc,NumberTask:integer;namePC,TYPETASK,UserName,PassWd:string):boolean; // функция запуска установки программы или запуска процесса
    Function InstallKeyOfficeWin7(StringKey,namePC,User,Passwd:string;NumberTask:integer):boolean;  // установка ключа office в windows 7
    function ActivateOfficeWin7(namePC,User,Passwd:string;NumberTask:integer):boolean;  // активация офиса в windows 7
    Function InstallKeyWinOffice(StringKey,Product,namePC,User,Passwd:string;NumberTask:integer):boolean; // установка ключа Windows для всех версии , Office все кроме Windows 7
    function ActivateProduct (Product,namePC,User,Passwd:string;NumberTask:integer):boolean; // активация Windows для всех версии , Office все кроме Windows 7
    function ControlService(namePC,User,Passwd,NameService:string;Oper,NumberTask:integer):boolean; // управление службами
    function DescriptionLicStatus(num:string):string;
    function RefreshLicinseWin(namePC,user,passwd:string;NumberTask:integer):boolean;  // обновление статуса лицензии windows
    function refreshOfficeLic(NamePC,User,Passwd:string):boolean;     /// обновление статуса лицензии офис
    function MyNewProcess(FSource :String;   ///что копировать источник файл или каталог
    FDest:string; // куда копировать
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    PathCreate:boolean;  // проверка и создания директории
        CancelCopyFF:boolean;   // отменить или нет операцию копирования
    BeforeInstallCopy:boolean;     // копировать или нет перед установкой
    DeleteAfterInstall:boolean; // удалить дистрибутив после установки
    PathDelete:string;         // Какой каталог или файл удалять после установки
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    FileToRun:String;   // файл для запуска
    NumberTask:integer):boolean;
    function RunInstallMSI( FSource :String;   /// источник файл
    FDest:string;      //файл назначение
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    PathCreate:boolean;  // проверка и создания директории
    CancelCopyFF:boolean;   // отменить или нет операцию копирования
    BeforeInstallCopy:boolean;     // копировать или нет перед установкой
    DeleteAfterInstall:boolean; // удалить дистрибутив после установки
    PathDelete:string;         // Какой каталог или файл удалять после установки
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    ProgramInstall:String;   // программа для установки
    KeyNewInstallProgram:string; // ключи установки программы
    InstallAllUsers:boolean; // устанавливать для всех пользователей
    NumberTask:integer):boolean;  // Номер задания в задаче
    function WakeOnLan(NamePC,BRaddress:string;NumTASK:integer):boolean;
    function KillProcess(NamePC,UserName,Passwd,NameProcess:string;equallyName:boolean;
NumTASK:integer):boolean;
    function DeleteProgramMSI(NamePC,UserName,Passwd,NameProgram:string;
equallyName:boolean;NumTASK:integer):boolean;    ///// удаление программы msi
    function ShotDownResetCloseSession(namePC,Myuser,MyPasswd:string;myShutdown,NumTask:integer):boolean; // перезагрузка/завершение сеанса/завершение работы
    function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // копирование в потоке
     function CopyDeleteFF(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // копирование или удалить
     function FindAddcreateDir(path,NamePC:string;NumTask:integer):boolean;// проверка и создание директории
     function Log_write(Level:integer;fname, text:string):string;
     function finditem(StatusTask,ResultTask,namePC:string;NumTask:integer):boolean;
     function UpdateTableTask(NameTable,CurrentPC,CurrentTask,StatusTask:string;Operation:integer;StartStop,StartThread:boolean):boolean;
     function StartStopTask(nametable:string):boolean; /// читам статус задачи, остановить или продолжать выполнение
     function SelectedPing(HostName:string;Timeout,NumTask:integer):boolean;
     function waitPing(namePC,User,passwd:string;NumTask,WaitTime:integer):boolean; // функция ожидания появления компьютера в сети
     function MinPreOper(namePC:string;PreOper,min:integer):integer; //сравниваем время которое прошло с момента выполнения предыдущей операции и время ожидания
     function ConnectWMI (NamePC,User,Passwd:string):Boolean; // проверка доступности WMI
     function CloseConnect:boolean;
     Function renewPath(s:string):string;
     function StrInDescription(s:string):string;
     function StrCopyDestination(s:string):string;
     function StrDeleteSourse(s:string):string;
     function StrCopySourse(s:string):string;
  protected
    procedure Execute; override;
    
  end;

  ThreadVar
  RunNewTask,DesTask:string; // типа имя таблицы  и описание таблицы
  FDQueryWrite,FDQueryWriteTaskTable: TFDQuery;
  TransactionWriteTT  : TFDTransaction;
  ConnectionThreadTask: TFDConnection;
  const
  wbemFlagForwardOnly = $00000020;

implementation
uses umain,TaskEdit,RunTaskRegEdit,RunTaskOsher,PingForTask;
//////////////////////////////////////////////
function bToSTR(b:boolean):string;  //true - 1 иначе false - 0
begin
  if b then result:='1'
  else result:='0';
end;

function STRtoB(b:string):boolean;  //true - 1 иначе false - 0
begin
  if b='0' then result:=false
  else result:=true;
end;
///////////////////////////////////////////////////////////
function TaskRun.Log_write(Level:integer;fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists(ExtractFilePath(Application.ExeName)+'log') then CreateDir(ExtractFilePath(Application.ExeName)+'log');
  if not DirectoryExists(ExtractFilePath(Application.ExeName)+'log\TASK') then CreateDir(ExtractFilePath(Application.ExeName)+'log\TASK');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\TASK\'+DesTask+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\TASK\'+DesTask+'.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\TASK\'+DesTask+'.log');
      finally
        f.Destroy;
      end;
except
exit;
end;

end;


 ////////////////////////////////////////////////////  функция записи результата в таблицу
 function TaskRun.finditem(StatusTask,ResultTask,namePC:string;NumTask:integer):boolean;
var
step:integer;
begin
try
step:=0;
FDQueryWrite.SQL.Clear;
TransactionWriteTT.StartTransaction;
step:=2;
if StatusTask='' then
begin
step:=3;
FDQueryWrite.SQL.Text:='update or insert into '+RunNewTask+' (PC_NAME, ResultTask'+inttostr(NumTask)+') VALUES (:p1,:p2)  MATCHING (PC_NAME)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+namePC+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+leftstr(ResultTask,220)+' ('+DateTimetostr(now)+')'+'';
end
else
begin
step:=4;
FDQueryWrite.SQL.Text:='update or insert into '+RunNewTask+' (PC_NAME, ResultTask'+inttostr(NumTask)+',StatusTask'+inttostr(NumTask)+') VALUES (:p1,:p2,:p3)  MATCHING (PC_NAME)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+namePC+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+leftstr(ResultTask,220)+' ('+DateTimetostr(now)+')'+'';
FDQueryWrite.params.ParamByName('p3').AsString:=''+StatusTask+'';
end;
step:=5;
FDQueryWrite.ExecSQL;
TransactionWriteTT.Commit;
FDQueryWrite.Close;
result:=true;
step:=6;
  except on E: Exception do
     begin
     TransactionWriteTT.Rollback;
     Log_write(2,'TASK',' Ошибка записи результатов на шаге '+inttostr(step)+' :' +e.Message);
     result:=false;
     end;
   end;
end;
////////////////////////////////////////////////////////////////////////// функция обновления таблицы в которой содержаться записи о всех задачах
 function TaskRun.UpdateTableTask(NameTable,CurrentPC,CurrentTask,StatusTask:string;Operation:integer;StartStop,StartThread:boolean):boolean;
var
 //Operation : 0- обновление столцов текущий компьютер и имя выполняемой задачи
 //            1 - STATUS_TASK - выполняется или остановлена по окончанию  + статус остановки задачи по требованию START_STOP
step:integer;
begin
try
step:=1;
FDQueryWriteTaskTable.SQL.Clear;
TransactionWriteTT.StartTransaction;
step:=2;
case operation of
0: begin
   step:=3;
   FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET PC_RUN=:a, CURRENT_TASK=:b, WORKS_THREAD=:c  WHERE NAME_TABLE='''+NameTable+'''';
   FDQueryWriteTaskTable.ParamByName('a').asstring:=CurrentPC;
   FDQueryWriteTaskTable.ParamByName('b').asstring:=CurrentTask;
   FDQueryWriteTaskTable.ParamByName('c').asstring:=booltostr(StartThread); ///признак того что поток выполняется  или остановлен
   end;
1:begin
  step:=4;
  FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET STATUS_TASK=:a, START_STOP=:b, WORKS_THREAD=:c  WHERE NAME_TABLE='''+NameTable+'''';
  FDQueryWriteTaskTable.ParamByName('a').asstring:=StatusTask;
  FDQueryWriteTaskTable.ParamByName('b').asstring:=booltostr(StartStop);
  FDQueryWriteTaskTable.ParamByName('c').asstring:=booltostr(StartThread); // признак того что поток выполняется или остановлен
  end;
2:begin
  step:=5;
  FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET STATUS_TASK=:a, START_STOP=:b, PC_RUN=:c, CURRENT_TASK=:d, WORKS_THREAD=:e   WHERE NAME_TABLE='''+NameTable+'''';
  FDQueryWriteTaskTable.ParamByName('a').asstring:=StatusTask;
  FDQueryWriteTaskTable.ParamByName('b').asstring:=booltostr(StartStop);
  FDQueryWriteTaskTable.ParamByName('c').asstring:='';
  FDQueryWriteTaskTable.ParamByName('d').asstring:='';
  FDQueryWriteTaskTable.ParamByName('e').asstring:=booltostr(StartThread); // признак того что поток выполняется или остановлен
  end;
3:begin  // сброс пароля в таблице
  FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET USER_NAME=:a ,PASS_USER=:b  WHERE NAME_TABLE='''+NameTable+'''';
  FDQueryWriteTaskTable.ParamByName('a').asstring:='';
  FDQueryWriteTaskTable.ParamByName('b').asstring:='';
  end;

end;
step:=6;
FDQueryWriteTaskTable.ExecSQL;
step:=7;
result:=true;
TransactionWriteTT.Commit;
FDQueryWriteTaskTable.Close;
 except on E: Exception do
     begin
     TransactionWriteTT.Rollback;
     Log_write(2,'TASK',' Ошибка UPDATE TABLE_TASK на шаге '+inttostr(step)+' :' +e.Message);
     result:=false;
     end;
   end;
end;
///////////////////////////////////////////////////////////////////////////
function TaskRun.StartStopTask(nametable:string):boolean; /// читам статус задачи, остановить или продолжать выполнение
var
FDQueryRead:TFDQuery;
begin
try
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionWriteTT;
FDQueryRead.Connection:=ConnectionThReadTask;
TransactionWriteTT.StartTransaction; // стартуем
FDQueryRead.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE NAME_TABLE='''+nametable+'''';
FDQueryRead.open;
result:=strtobool(FDQueryRead.FieldByName('START_STOP').AsString);
TransactionWriteTT.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryRead.SQL.clear;   //очистить
FDQueryRead.Close;  /// закрыть нах
FDQueryRead.Free;
 except on E: Exception do
     begin
     result:=true;
     TransactionWriteTT.Rollback;
     if Assigned(FDQueryRead) then FDQueryRead.Free;
     Log_write(1,'TASK',' Ошибка READ StartStopTask TABLE_TASK :' +e.Message);
     end;
   end;
end;
///////////////////////////////////////////////////////////////////////////////////

function TaskRun.SelectedPing(HostName:string;Timeout,NumTask:integer):boolean;
var
avalible:boolean;
resIP:string;
begin
resIP:='';
avalible:=false;
  case PingType of
    1:
    begin
    if PingIdIcmp(resIP,HostName,pingtimeout) then avalible:=true
    else avalible:=false;
    end;
    2:
    begin
    if PingGetaddrinfo(resIP,HostName,pingtimeout) then avalible:=true
    else avalible:=false;
    end;
    3:
    begin
     if PingGetHostByName(resIP,HostName,pingtimeout) then avalible:=true
    else  avalible:=false;
    end;
  END;
  if not avalible then finditem('NO','Узел не доступен',HostName,NumTask);;
resIP:='';
Result:=avalible;
end;

////////////////////////////////////////////////////
 function TaskRun.FindAddcreateDir(path,NamePC:string;NumTask:integer):boolean;// проверка и создание директории
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //если нет каталога
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// создает весь путь до папки
     begin
     finditem('','Создание директории ' +ExtractFileDir(path)+' : произошла ошибка ',NamePC,NumTask);
     result:=false;
     end
     else result:=true; // директория создана
    end
    else result:=true; // директория есть
  except on E: Exception do
     begin
     finditem('','Создание директории ' +ExtractFileDir(path)+' :' +e.Message,NamePC,NumTask);
     result:=false;
     end;
   end;
end;


////////////////////////////////////////////////////////////////
function TaskRun.CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // копирование
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// новая функция начиная с vista , копирует файлы и каталоги для группы компьютеров
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //Когда функция возвращается, этот член содержит дескриптор объекта сопоставления имен, который содержит старые и новые имена переименованных файлов
      lpszProgressTitle :=nil;            // Указатель на заголовок диалогового окна прогресса
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // от куда копируем если операция копирования
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // куда копируем
      finditem('','Запущена операция копирования',CurentPC,NumTask);
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // что удалять  если операция удаление
      pTo := pchar('');   // куда копируем не используется
      //finditem('','Запущена операция удаления дистрибутива',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// авторизация пользователя на удаленном компе
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // сначала заходим на комп в сети
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do finditem('',' : Ошибка LogonUser - '+e.Message,CurentPC,NumTask)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC,NumTask); //проверяем и создаем каталог если его нет.
     rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения /удаления
     if rescopy=0 then
       begin
        if TypeOperation=2 then finditem('','Операция копирования успешно завершена',CurentPC,NumTask);
        //if TypeOperation=3 then  finditem('','Операция удаления дистрибутива успешно завершена',CurentPC,NumTask); //finditem('Операция удаления успешно завершена',CurentPC,1);
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then finditem('','Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);
        //if TypeOperation=3 then finditem('','Оперция удаления дистрибутива завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);//finditem('Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        result:=false;
       end;
     CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
       begin
       finditem('','Ошибка оперции копирования: '+E.Message,CurentPC,NumTask);
       end;
      if TypeOperation=3 then
       begin
       finditem('','Ошибка оперции удаления: '+E.Message,CurentPC,NumTask);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then finditem('','Общая ошибка функции копирования перед установкой программы  : '+E.Message,CurentPC,NumTask);
     if TypeOperation=3 then finditem('','Общая ошибка функции удаления дистрибутива после установки программы  : '+E.Message,CurentPC,NumTask);
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////

function TaskRun.ShotDownResetCloseSession(namePC,Myuser,MyPasswd:string;myShutdown,NumTask:integer):boolean; // перезагрузка/завершение сеанса/завершение работы

var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown:integer;
  Operation     :string;
begin
  ///  if (namePC<>'') and (ping(namePC)) then
      begin
          try
              OleInitialize(nil);
              FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
              FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', MyUser, MyPasswd,'','',128);
              FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
              oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
              if oEnum.Next(1, FWbemObject, iValue) = 0 then
                case myShutdown of
                 0:
                  begin
                    ResShutdown:=FWbemObject.Win32Shutdown(0);
                    Operation:='Завершение сеанса';
                    if ResShutdown<>0 then
                      begin
                      ResShutdown:=FWbemObject.Win32Shutdown(4);
                      Operation:='Принудительное завершение сеанса';
                      end;
                  end;
                  1:
                  begin
                   ResShutdown:=FWbemObject.Win32Shutdown(4);
                   Operation:='Принудительное завершение сеанса';
                  end;
                   2:
                   begin
                     ResShutdown:=FWbemObject.Win32Shutdown(2);
                     Operation:='Перезагрузка';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(6);
                        Operation:='Принудительная перезагрузка';
                        end;
                   end;
                   3:
                   begin
                   ResShutdown:=FWbemObject.Win32Shutdown(6);
                   Operation:='Принудительная перезагрузка';
                   end;
                   4:
                   begin
                     ResShutdown:=FWbemObject.Win32Shutdown(1);
                     Operation:='Завершение работы';
                     if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(5);
                        Operation:='Принудительное завершение работы';
                        end;
                    end;
                    5:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(5);
                      Operation:='Принудительное завершение работы';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(12);
                        Operation:='Форсируем завершение работы';
                        end;
                    end;
                    6:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(12);
                      Operation:='Форсируем завершение работы';
                    end;
                 end;
////////////////////////////////////////////////////////////////////////////////////////////////
          if ResShutdown=0 then
            begin
             result:=true;
             finditem('OK',Operation+' :операция успешно завершена',namePC,NumTask);
            end
          else
            begin
            result:=false;
            finditem('WARNING',Operation+': Ошибка - '+SysErrorMessage(ResShutdown),namePC,NumTask);
            end;

///////////////////////////////////////////////////////////////////////////////
              FWbemObject:=Unassigned;
              VariantClear(FWbemObject);
              oEnum:=nil;
              VariantClear(FWbemObjectSet);
              VariantClear(FWMIService);
              VariantClear(FSWbemLocator);
              OleUnInitialize;
              except
            on E:Exception do
             begin
             finditem('ERROR','Общая ошибка управления питанием : '+E.Message,namePC,NumTask);
             result:=false;
             OleUnInitialize;
             end;
      end;
end;
end;
/////////////////////////////////////////////////////////////////
function TaskRun.DeleteProgramMSI(NamePC,UserName,Passwd,NameProgram:string;
equallyName:boolean;NumTASK:integer):boolean;    ///// удаление программы msi
var
  FSWbemLocator   : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResDelProgram,i,z: integer;
  YesProg:boolean;
begin
try
OleInitialize(nil);
YesProg:=false;
finditem('RUN','Поиск программы для удаления :'+NameProgram,NamePC,NumTASK) ;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', UserName, Passwd,'','',128);
if not equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+NameProgram+'%"','WQL',wbemFlagForwardOnly); // Все программы по совпадению
if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+NameProgram+'"','WQL',wbemFlagForwardOnly);          // точное совпадение имени программы
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
  begin
  finditem('RUN','Запускаю процесс удаления программы :'+NameProgram,NamePC,NumTASK) ;
  ResDelProgram:=FWbemObject.Uninstall(); //// Удалениме программы
  FWbemObject:=Unassigned;
  YesProg:=true;
  end;
if YesProg then
begin
  if (ResDelProgram=0)  then // 0- Операция завершена, 3010 - после установки требует перезагрузки???
  begin
  finditem('OK','Удаление программы '+NameProgram+' :'+SysErrorMessage(ResDelProgram),NamePC,NumTASK);
  result:=true;
  end
  else
  begin
  finditem('WARNING','При удалении программы '+NameProgram+' возникли проблемы : ('+inttostr(ResDelProgram)+') '+SysErrorMessage(ResDelProgram),NamePC,NumTASK);
  result:=false;
  end;
end;
if not YesProg then
  begin
  finditem('OK','Программа '+NameProgram+' не найдена',NamePC,NumTASK);
  result:=true;
  end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
  except
  on E:Exception do
    begin
    finditem('ERROR','При удалении программы '+NameProgram+' произошла ошибка. - "'+E.Message+'"',NamePC,NumTASK);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    OleUnInitialize;
    result:=false;
    end;
  end;
end;
///////////////////////////////////////////////////////////////////////

function TaskRun.KillProcess(NamePC,UserName,Passwd,NameProcess:string;equallyName:boolean;
NumTASK:integer):boolean;
var ///////////////////////////////////// завершение процесса на выбранных машинах в одном потоке
YesProc:boolean;
i,z,ErrorProcKill:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
BEGIN
try
Begin
YesProc:=false;
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', UserName, Passwd,'','',128);
if not equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Process WHERE Caption LIKE'+'"%'+NameProcess+'%"','WQL',wbemFlagForwardOnly);
if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Process WHERE Caption ='+'"'+NameProcess+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
begin
finditem('RUN','Завершения процесса '+NameProcess,NamePC,NumTASK);
ErrorProcKill:=FWbemObject.terminate();
FWbemObject:=Unassigned;
YesProc:=true;
end;
if (YesProc=true) and (ErrorProcKill=0) then
begin
finditem('OK','Завершение процесса '+NameProcess+' прошло успешно',NamePC,NumTASK);
end;
if (YesProc=true)and (ErrorProcKill<>0) then
begin
finditem('WARNING','При завершении процесса '+NameProcess+' возникли проблемы: '+SysErrorMessage(ErrorProcKill),NamePC,NumTASK);
end;
if (YesProc=false)then
finditem('OK','Процесс '+NameProcess+' не найден',NamePC,NumTASK);
End;

VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;

except
on E:Exception do
  begin
  finditem('ERROR','Общая ощибка завершения процесса '+NameProcess+' : "'+E.Message+'"',NamePC,NumTASK);
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  end;
end;
END;
////////////////////////////////////////////////////////
function TaskRun.WakeOnLan(NamePC,BRaddress:string;NumTASK:integer):boolean;
var
  i, j: Byte;
  lBuffer: tbytes;
  lUDPClient: TIDUDPClient;
  setini:Tmeminifile;
  aMacAddress,RealMAC:string;
  FDQueryRead:TFDQuery;
begin
aMacAddress:='';
Sleep(1500); // остановка на 1,5 секунды, чтобы разгрузить БД

FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionWriteTT;
FDQueryRead.Connection:=ConnectionThReadTask;
TransactionWriteTT.StartTransaction;
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT ANSWER_MAC,PC_NAME FROM MAIN_PC WHERE PC_NAME='''+NamePC+'''';
FDQueryRead.Open;
if FDQueryRead.FieldByName('ANSWER_MAC').Value<>null then // иначе MAC нет
aMacAddress:=FDQueryRead.FieldByName('ANSWER_MAC').asString;
TransactionWriteTT.Commit;
FDQueryRead.Close;
FDQueryRead.Free;

if aMacAddress='' then
begin
finditem('WARNING','MAC адрес компьютера '+NamePC+' не найден, включите сканирование сети',NamePC,NumTASK);
result:=false;
exit;
end;
//finditem('RUN','Отправляю пакет WOL на MAC адрес '+aMacAddress,NamePC,NumTASK);
RealMAC:=aMacAddress;
aMacAddress := StringReplace(uppercase(aMacAddress), '-', '', [rfReplaceAll]);
aMacAddress := StringReplace(aMacAddress, ':', '', [rfReplaceAll]);
  try
    SetLength(lbuffer,117);
    for i := 1 to 6 do
    begin
      lBuffer[i] :=StrToIntDef('$' + aMacAddress[(i * 2) - 1] + aMacAddress[i * 2],0);
    end;
    lBuffer[7] := $00;
    lBuffer[8] := $74;
    lBuffer[9] := $FF;
    lBuffer[10] := $FF;
    lBuffer[11] := $FF;
    lBuffer[12] := $FF;
    lBuffer[13] := $FF;
    lBuffer[14] := $FF;
    for j := 1 to 16 do
    begin
      for i := 1 to 6 do
      begin
        lBuffer[15 + (j - 1) * 6 + (i - 1)] := lBuffer[i];
      end;
    end;
    lBuffer[116] := $00;
    lBuffer[115] := $40;
    lBuffer[114] := $90;
    lBuffer[113] := $90;
    lBuffer[112] := $00;
    lBuffer[111] := $40;
    try
      lUDPClient := TIdUDPClient.Create(nil);
      lUDPClient.BroadcastEnabled := true;
      lUDPClient.Host := BRaddress;   ///// траблы с указанием IP адреса (Необходимо определять подсеть в какую слать пакет, УТОЧНИТЬ)
      lUDPClient.Port := 9;
      lUDPClient.SendBuffer(lUDPClient.Host,lUDPClient.Port ,tidbytes(lBuffer));
      finditem('OK','WOL пакет отправлен на MAC адрес '+RealMAC,NamePC,NumTASK);
      result:=true;
    finally
      lUDPClient.Free;
    end;
  except
   on E: Exception do
   begin
    finditem('ERROR','Ошибка отправки WOL на MAC адрес '+RealMAC+' :'+E.Message,NamePC,NumTASK);
    result:=false;
   end;
  end;
end;



////////////////////////////////////////////////////
function TaskRun.RunInstallMSI( FSource :String;   /// источник файл
    FDest:string;      //файл назначение
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    PathCreate:boolean;  // проверка и создания директории
    CancelCopyFF:boolean;   // отменить или нет операцию копирования
    BeforeInstallCopy:boolean;     // копировать или нет перед установкой
    DeleteAfterInstall:boolean; // удалить дистрибутив после установки
    PathDelete:string;         // Какой каталог или файл удалять после установки
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    ProgramInstall:String;   // программа для установки
    KeyNewInstallProgram:string; // ключи установки программы
    InstallAllUsers:boolean; // устанавливать для всех пользователей
    NumberTask:integer):boolean;  // Номер задания в задаче
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  Resinstall,b,z,i :integer;
  CopyOrNo:boolean;
  ListPC:TstringList;
BEGIN
ListPC:=Tstringlist.Create;
ListPC.CommaText:=NamePC;
for I := 0 to ListPC.Count-1 do
Begin
// if ping(ListPC[i]) then
  begin
  CopyOrNo:=false;
  if BeforeInstallCopy then // если необходимо скопировать перед установкой
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // имя компа
  UserName,//логин
  PassWd,  // пароль
  FSource, // что копировать
  FDest,   // куда копировать
  OwnerForm,// родительская форма
  TypeOperation, //тип операции
  CancelCopyFF,
  PathCreate,
  NumberTask) ; // порядковый номер выполняемого задания
  end;
  try
      finditem('RUN','Запускаю процесс установки программы '+ExtractFileName(ProgramInstall),ListPC[i],NumberTask);
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(ListPC[i], 'root\CIMV2', UserName, PassWd,'','',128);
      FWbemObject   := FWMIService.Get('Win32_Product');
      Resinstall:=FWbemObject.install(ProgramInstall,KeyNewInstallProgram,InstallAllUsers); //файл программы, ключи установки, для всех пользователей

       if (Resinstall=0) or (Resinstall=3010) or (Resinstall=1638) then  // 0- Операция завершена, 3010 - после установки требует перезагрузки, 1638 - Уже установлена другая версия этого продукта. Продолжение установки невозможно.
        Begin
        finditem('OK',leftstr('Установка программы ' +ExtractFileName(ProgramInstall)
        +' : '+SysErrorMessage(Resinstall),499),ListPC[i],NumberTask);

        End
       else
        Begin
        finditem('WARNING',leftstr('При установке программы '
        +ExtractFileName(ProgramInstall) +' возникли ошибки : '+' ('+inttostr(Resinstall)+') '
        +SysErrorMessage(Resinstall),499),ListPC[i],NumberTask);

        End;
          {frmDomainInfo.memo1.Lines.Add('Установка программы '
          +ExtractFileName(PointForInstallMSI.ProgramInstall)+' на '
          +ListPC[i]+' : '+SysErrorMessage(Resinstall));}
        FWbemObject:=Unassigned;
        VariantClear(FWbemObject);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
        if BeforeInstallCopy and  //если указано скопировать перед установкой
         DeleteAfterInstall and    // если необходимо удалить дистрибутив после установки
         CopyOrNo then // если операция копирования перед установкой прошла успешно
          begin        // удаляем
          CopyFFSelectPC(ListPC[i], // имя компа
          UserName,//логин
          PassWd,  // пароль
          '', // что копировать можно не укузывать т.к. удаляем
          PathDelete,   // что удалять
          OwnerForm,// родительская форма
          3, //тип операции  (3)- удалить, FO_MOVE
          CancelCopyFF,
          false,  // не проверять наличие каталога т.к. удаляем после копирования
          NumberTask) ; // порядковый номер выполняемого задания
          end;
      OleUnInitialize;
        except
          on E:Exception do
           Begin
           finditem('ERROR','Ошибка установки программы '
           +ExtractFileName(ProgramInstall) +' : '+E.Message,ListPC[i],NumberTask);
           VariantClear(FWbemObject);
           VariantClear(FWMIService);
           VariantClear(FSWbemLocator);
           OleUnInitialize;
           exit;
           End;
       end; // except
  end; //ping
  End; // цикл
ListPC.Free;
END;
///////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////
//запуск процесса
function TaskRun.MyNewProcess(FSource :String;   ///что копировать источник файл или каталог
    FDest:string; // куда копировать
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    PathCreate:boolean;  // проверка и создания директории
        CancelCopyFF:boolean;   // отменить или нет операцию копирования
    BeforeInstallCopy:boolean;     // копировать или нет перед установкой
    DeleteAfterInstall:boolean; // удалить дистрибутив после установки
    PathDelete:string;         // Какой каталог или файл удалять после установки
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    FileToRun:String;   // файл для запуска
    NumberTask:integer):boolean;
var
MyError:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObject   : OLEVariant;
objProcess    : OLEVariant;
objConfig     : OLEVariant;
ProcessID,z,i   : Integer;
CopyOrNo:bool;
const
wbemFlagForwardOnly = $00000020;
HIDDEN_WINDOW       = 1;
/////////////////////////////////////////////////////////////////////
BEGIN
  try
  CopyOrNo:=false;
  if BeforeInstallCopy then // если необходимо скопировать перед установкой
  begin
  CopyOrNo:=CopyFFSelectPC(NamePC, // имя компа
  UserName,//логин
  PassWd,  // пароль
  FSource, // что копировать
  FDest,   // куда копировать
  OwnerForm,// родительская форма
  TypeOperation, //тип операции
  CancelCopyFF,
  PathCreate,
  NumberTask); // порядковый номер выполняемого задания                                                                                  end;
  end;
  OleInitialize(nil);
  finditem('RUN','Запускаю процесс :'+FileToRun,NamePC,NumberTask);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', UserName, PassWd,'','',128);
  FWMIService.security_.AuthenticationLevel:=6;
  FWMIService.security_.ImpersonationLevel:=3;
  //FWMIService.security_.Privileges.AddAsString('SeEnableDelegationPrivilege');
  FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
  objConfig     := FWbemObject.SpawnInstance_;
  objConfig.ShowWindow := HIDDEN_WINDOW;
  objProcess    := FWMIService.Get('Win32_Process');
  MyError:=objProcess.Create(FileToRun, null, objConfig, (ProcessID));
  if MyError=0 then
    begin
    finditem('OK','Запуск процесса '+FileToRun+' :  '+SysErrorMessage(MyError),NamePC,NumberTask);
    end
  else
    begin
    finditem('WARNING','При запуске процесса '+FileToRun+' возникли ошибки : ('+inttostr(MyError)+') '+SysErrorMessage(MyError),NamePC,NumberTask);
    end;
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  VariantClear(FWbemObject);
  VariantClear(objConfig);
  VariantClear(objProcess);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  if BeforeInstallCopy and  //если указано скопировать перед установкой
  DeleteAfterInstall and    // если необходимо удалить дистрибутив после установки
  CopyOrNo then // если операция копирования перед установкой прошла успешно
    begin        // удаляем
    CopyFFSelectPC(NamePC, // имя компа
    UserName,//логин
    PassWd,  // пароль
    '', // что копировать можно не укузывать т.к. удаляем
    PathDelete,   // что удалять
    OwnerForm,// родительская форма
    3, //тип операции  (3)- удалить, FO_MOVE
    CancelCopyFF,
    false,  // не проверять наличие каталога т.к. удаляем после копирования
    NumberTask) ; // порядковый номер выполняемого задания
    end;

    except
      on E:Exception do
      begin
      finditem('ERROR','При запуске процесса '+FileToRun+' возникли проблемы : "'+E.Message+'"',NamePC,NumberTask);
      OleUnInitialize;
      end;
    end;
END;
///////////////////////////////////////////////////////////
function TaskRun.ControlService(namePC,User,Passwd,NameService:string;Oper,NumberTask:integer):boolean;
var
ResServic:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
str:string;
findService:boolean;
Const OWN_PROCESS = 16;
 NOT_INTERACTIVE = True;
 ErrorCodeService: array [0..24] of string=('Операция успешно завершена',
 'Операция не поддерживается','У пользователя нет необходимого доступа',
 'Службу невозможно остановить, поскольку от нее зависят другие запущенные службы',
 'Запрошенный контрольный код недействителен или неприемлем для службы.'
 ,'Запрошенный код элемента управления не может быть отправлен службе из-за состояния службы '
 ,'Служба не запущена','Служба не ответила на запрос своевременно',
 'Неизвестный сбой при запуске службы','Путь к исполняемому файлу службы не найден'
 ,'Служба уже запущена.','База данных для добавления новой службы заблокирована',
 'Зависимость, на которую опирается эта служба, была удалена из системы.',
 'Службе не удалось найти службу, необходимую для зависимой службы.',
 ' Служба была изключена из системы.','Служба не имеет правильную проверку подлинности для запуска в системе.'
 ,'Эта служба удаляется из системы','Служба не имеет потока выполнения',' При запуске служба имеет циклические зависимости'
 ,'Служба выполняется под тем же именем','Имя службы содержит недопустимые символы'
 ,'Службе переданы недопустимые параметры.','Учетная запись, под которой выполняется эта служба, является недопустимой или не имеет разрешений на запуск службы.'
 ,'Служба существует в базе данных сервисов, доступных из системы.',
 'В настоящее время служба приостановлена в системе');
begin
try
OleInitialize(nil);
case oper of
1:str:='Запуск службы';
2:str:='Остановка службы';
3:str:='Удаление службы';
5:str:='Смена типа запуска службы';
end;
findService:=false;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
//if Oper=4 then FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Service','WQL',wbemFlagForwardOnly)
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name FROM Win32_Service WHERE Name = '+'"'+NameService+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
Begin
findService:=true;
finditem('RUN',str+' :'+NameService,NamePC,NumberTask);
case Oper of
1:ResServic:=FWbemObject.StartService();  ////// запустить службу
2:ResServic:=FWbemObject.StopService();  ////// остановить службу
3: ResServic:=FWbemObject.Delete(); //// Удалениме службы
//4://ResServic:=FWbemObject3.create;//('"DbService"', '"Personnel Database"', '"C:\Program Files (x86)\2gis\3.0\2GISUpdateService.exe"',OWN_PROCESS ,1 ,'"Automatic"' ,NOT_INTERACTIVE,'".\LocalSystem"','');  ////// Создать службу
5:ResServic:=FWbemObject.ChangeStartMode(TypeRunService); //// смена типа запуска
end;
FWbemObject:=Unassigned;
End;

if findService then   // если служба найдена
begin
  if ResServic=0 then finditem('OK',str+' :'+NameService+' :  Операция успешно завершена',NamePC,NumberTask)
  else
  begin
  if ResServic>24 then finditem('WARNING',str+' :'+NameService+' :  '+SysErrorMessage(ResServic),NamePC,NumberTask)
  else finditem('WARNING',str+' :'+NameService+' :  '+ErrorCodeService[ResServic],NamePC,NumberTask);
  end;
end
else  // иначе такая служба не найдена на компе
finditem('OK','Cлужба :'+NameService+' не найдена :  Операция успешно завершена',NamePC,NumberTask);

VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
result:=true;
except
on E:Exception do
begin
finditem('ERROR',str+' :'+NameService+' произошла ошибка : '+E.Message,NamePC,NumberTask);
result:=false;
end;
end;

end;
/////////////////////////////////////////////////////////////////////////////////////////////++
//////////////////////////////////активация windows и office //////////////////////////////////////
function TaskRun.DescriptionLicStatus(num:string):string;
var   /// передаем код ошибки и из таблицы получаем описание этой ошибки
z:integer;  //slui.exe 0x2a 0xC004FE00
he,s,A:string;
TransactionWrite:TFDTransaction;
FDQueryWrite:TFDQuery;
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
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThReadTask;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=false;
FDQueryWrite:= TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=ConnectionThReadTask;
try
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
FDQueryWrite.SQL.Text:='update or insert into LIC_ERROR (CODE,DESCRIPTION,ACTIVATE) VALUES (:p1,:p2,:p3) MATCHING (CODE,DESCRIPTION)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(he)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+('Unknown')+'';
FDQueryWrite.params.ParamByName('p3').AsBoolean:=false;
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
end;
finally
FDQueryWrite.Close;
FDQueryWrite.Free;
TransactionWrite.Free;
end;
except on E: Exception do
begin
 result:=num;
 Log_write(1,'TASK',' DescriptionLicStatus - '+e.Message);
end;
end;
end;
////////////////////////////////////////////////////////////////////Product - Office и  Operating System
function TaskRun.ActivateProduct (Product,namePC,User,Passwd:string;
NumberTask:integer):boolean; // активация Windows для всех версии , активация Office кроме Windows 7
var
FWMIService         : OLEVariant;
FSWbemLocator         : OLEVariant;
ObjectSet,FWbemObject      : OLEVariant;
oEnum                : IEnumvariant;
iValue        : LongWord;
ActProduct:boolean;
step:integer;
NameProdukt:string;
begin
try
 OleInitialize(nil);
 step:=0;
 ActProduct:=false;
 finditem('RUN','Активация продукта',NamePC,NumberTask);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 step:=1;
 FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
 step:=2;
 FWMIService.Security_.impersonationlevel:=3;
 step:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy                            WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly
 step:=4;
 {LicenseStatus	DWORD	4	Состояние лицензирования
0 - Нелицензировано
1 - Лицензировано (активировано)
2 – Льготный период OOB
3 – Льготный период OOT
4 - NonGenuineGrace
}                                  //'SELECT  ProductKeyID,ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus,LicenseStatusReason FROM SoftwareLicensingProduct WHERE Description LIKE "%'+Product+'%"'+' and PartialProductKey <> null and LicenseStatus<>1 and LicenseStatus<>0'                                                                                                               //
 ObjectSet:= FWMIService.ExecQuery('SELECT  ProductKeyID,ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus,LicenseStatusReason FROM SoftwareLicensingProduct WHERE Description LIKE "%'+Product+'%"'+' and PartialProductKey <> null and LicenseStatus<>1 and LicenseStatus<>0');   ///выбор продукта с номером и статусом не равным лицензия активна  и нет лицензии(ключа)
 oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   step:=5;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
     step:=6;
      if FWbemObject.name<>null then NameProdukt:=string(FWbemObject.name);
      FWbemObject.Activate();
      step:=7;
      ActProduct:=true;
     end;
     step:=8;
    if ActProduct then // если ключ установлен, то можно и активировать продукт
      begin
      finditem('OK','Успешная активация продукта '+NameProdukt,NamePC,NumberTask);
      step:=9;
      end
      else finditem('OK','Продукт активирован',NamePC,NumberTask);

    VariantClear(FWMIService);
    VariantClear(ObjectSet);
    VariantClear(FWbemObject);
    oEnum:=nil;
    result:=true;
 except
    on E:Exception do
      begin
      finditem('ERROR','('+inttostr(step)+') Ошибка активации продукта ('+NameProdukt+'): '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
      VariantClear(FWMIService);
      VariantClear(ObjectSet);
      VariantClear(FWbemObject);
      oEnum:=nil;
      result:=false;
      end;
    end;
OleUnInitialize;
end;
                                  // Product - Office и  Operating System
Function TaskRun.InstallKeyWinOffice(StringKey,Product,namePC,User,Passwd:string;
NumberTask:integer):boolean; // установка ключа Windows для всех версии , Office все кроме Windows 7
var
  FWMIService         : OLEVariant;
  FSWbemLocator       : OLEVariant;
  ObjectSet,FWbemObject: OLEVariant;
  oEnum                : IEnumvariant;
  iValue        : LongWord;
  InstKey:boolean;
Begin
try
  OleInitialize(nil);
  InstKey:=false;
  finditem('RUN','Установка ключа активации',NamePC,NumberTask);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy   WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly
  ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM SoftwareLicensingService');
  oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
      FWbemObject.InstallProductKey(StringKey); //// установка ключа windows
      FWbemObject:=Unassigned;
      InstKey:=true;
     end;
   VariantClear(ObjectSet);
   VariantClear(FWMIService);
   VariantClear(FWbemObject);
   oEnum:=nil;   
  if instkey then // если ключ установлен, то можно и активировать продукт
  begin
  finditem('OK','Установка ключа завершена',NamePC,NumberTask);
  RefreshLicinseWin(namePC,User,Passwd,NumberTask); // обновление статуса лицензии
  ActivateProduct(Product,namePC,User,Passwd,NumberTask);  // активация продукта
  end
  else finditem('WARNING','Ключ продукта не установлен.',NamePC,NumberTask);
  result:=true;
 except
   on E:Exception do
   begin
    finditem('ERROR','Ошибка установки ключа активации: '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
    VariantClear(ObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    oEnum:=nil;
    result:=false;
   end;
end;
OleUnInitialize;
End;
/////////////////////////////////////////////////////////////////////////////////
function TaskRun.RefreshLicinseWin(namePC,user,passwd:string;NumberTask:integer):boolean;    // обновление статуса лицензии Windows
var
  FWMIService         : OLEVariant;
  ObjectSet,FWbemObject      : OLEVariant;
  oEnum                : IEnumvariant;
  FSWbemLocator : OLEVariant;
  iValue        : LongWord;
Begin
  OleInitialize(nil);
  finditem('RUN','Обновление списка лицензий',NamePC,NumberTask);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
   try
  ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM SoftwareLicensingService');
   oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
      FWbemObject.RefreshLicenseStatus(); /// Обновления статуса лицензирования продукта
     end;
    result:=true;
   except
    on E:Exception do
      begin
      result:=false;
      OleUnInitialize;
      end;
    end;
    VariantClear(ObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;
    OleUnInitialize;
End;
///////////////////////////////////////////////////////////////////////////
function TaskRun.refreshOfficeLic(NamePC,User,Passwd:string):boolean;     /// обновление статуса лицензии офис
 var
  FSWbemLocator                           : OLEVariant;
  FWMIService                             : OLEVariant;
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
Begin
try
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Passwd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT version FROM OfficeSoftwareProtectionService','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
            begin
            FWbemObject.RefreshLicenseStatus();
            end;
   VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    result:=true;
    except
    on E:Exception do
     begin
     result:=false;
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet);
     VariantClear(FWMIService);
     VariantClear(FSWbemLocator);
     end;
    end;

end;

//////////////////////////////////////////////////////////////////////////////Активация для office в windows 7
function TaskRun.ActivateOfficeWin7(namePC,User,Passwd:string;NumberTask:integer):boolean;
var
FSWbemLocator : OLEVariant;
FWMIService         : OLEVariant;
ObjectSet,FWbemObject      : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
step:integer;
begin
try
step:=0;
 finditem('RUN','Активация продукта',NamePC,NumberTask);
 step:=1;
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Passwd,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
 step:=2;
 ObjectSet:= FWMIService.ExecQuery('SELECT Description FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey <> null and LicenseStatus<>1 and LicenseStatus<>0','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
 step:=3;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
     step:=4;
     FWbemObject.Activate();
     step:=5;
     FWbemObject:=Unassigned;
     end;
     step:=6;
    VariantClear(FWMIService);
    VariantClear(ObjectSet);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;
    step:=7;
    finditem('OK','Успешная активация продукта',NamePC,NumberTask);
    result:=true;
 except
    on E:Exception do
      begin
      result:=false;
      finditem('ERROR',inttostr(step)+') Ошибка активации продукта : '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
      VariantClear(FWMIService);
      VariantClear(ObjectSet);
      VariantClear(FWbemObject);
      VariantClear(FSWbemLocator);
      oEnum:=nil;
      end;
    end;
end;

function TaskRun.InstallKeyOfficeWin7(StringKey,namePC,User,Passwd:string;NumberTask:integer):boolean; //Установка ключа для office в windows 7
var
  FWMIService         : OLEVariant;
  ObjectSet,FWbemObject: OLEVariant;
  oEnum                : IEnumvariant;
  FSWbemLocator : OLEVariant;
  iValue        : LongWord;
  InstKey:boolean;
Begin
try
 OleInitialize(nil);
 InstKey:=false;
 finditem('RUN','Установка ключа активации',NamePC,NumberTask);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User,Passwd,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
 ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM OfficeSoftwareProtectionService','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do
  begin
  FWbemObject.InstallProductKey(StringKey); //// установка ключа office для windows 7
  FWbemObject:=Unassigned;
  InstKey:=true;
  end;
  VariantClear(ObjectSet);
  VariantClear(FWMIService);
  VariantClear(FWbemObject);
  VariantClear(FSWbemLocator);
  oEnum:=nil;
  if instkey then // если ключ установлен, то можно и активировать продукт
  begin 
  finditem('OK','Установка ключа завершена',NamePC,NumberTask);
  refreshOfficeLic(namePC,User,Passwd); // обновление статуса лицензии
  ActivateOfficeWin7(namePC,User,Passwd,NumberTask); // активация лицензии
  end
  else finditem('WARNING','Ключ продукта не установлен.',NamePC,NumberTask);
  result:=true;
 except
   on E:Exception do
   begin
   result:=false;
    finditem('ERROR','Ошибка установки ключа активации: '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
    VariantClear(ObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;
   end;
end;
OleUnInitialize;
End;
////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////
function TaskRun.CopyDeleteFF(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // копирование или удалить
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// новая функция начиная с vista , копирует файлы и каталоги для группы компьютеров
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //Когда функция возвращается, этот член содержит дескриптор объекта сопоставления имен, который содержит старые и новые имена переименованных файлов
      lpszProgressTitle :=nil;            // Указатель на заголовок диалогового окна прогресса
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // от куда копируем если операция копирования
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // куда копируем
      finditem('RUN','Запущена операция копирования',CurentPC,NumTask);
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // что удалять  если операция удаление
      pTo := pchar('');   // куда копируем не используется
      finditem('RUN','Запущена операция удаления',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// авторизация пользователя на удаленном компе
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // сначала заходим на комп в сети
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do finditem('',' : Ошибка LogonUser - '+e.Message,CurentPC,NumTask)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC,NumTask); //проверяем и создаем каталог если его нет.
     rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения /удаления
     if rescopy=0 then
       begin
        if TypeOperation=2 then finditem('OK','Операция копирования успешно завершена',CurentPC,NumTask);
        if TypeOperation=3 then  finditem('OK','Операция удаления успешно завершена',CurentPC,NumTask); //finditem('Операция удаления успешно завершена',CurentPC,1);
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then finditem('WARNING','Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);
        if TypeOperation=3 then finditem('WARNING','Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);//finditem('Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        result:=false;
       end;
     CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
       begin
       finditem('ERROR','Ошибка оперции копирования: '+E.Message,CurentPC,NumTask);
       end;
      if TypeOperation=3 then
       begin
       finditem('ERROR','Ошибка оперции удаления: '+E.Message,CurentPC,NumTask);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then finditem('ERROR','Общая ошибка функции копирования: '+E.Message,CurentPC,NumTask);
     if TypeOperation=3 then finditem('ERROR','Общая ошибка функции удаления: '+E.Message,CurentPC,NumTask);
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
Function TaskRun.renewPath(s:string):string;
begin
if AnsiPos('$',s)=2 then
 begin
 delete(s,2,1); // удаляем символ $
 insert(':',s,2); // вставляем на его место :
 end;
result:=s;
end;

function GetLastDir(path:String):String;  // извлечение корневой папки
var
i,j:integer;
begin
path:=ExtractFilePath(path); //Путь до файла меняем на path  файла
if Length(path)>3 then
    begin
      for I :=Length(path)-1 downto 1 do
      if path[i]='\' then
      begin
      j:=i+1;
      break;
      end;
      result:=Copy(path,j,Length(path)) ;
end
else
result:=path[1]+path[3];
end;

function TaskRun.ReadAndRunMsiOrProc (numberProc,NumberTask:integer;namePC,TYPETASK,UserName,PassWd:string):boolean; // функция запуска установки программы или запуска процесса
var
FDQueryRun: TFDQuery;
TransactionRun    : TFDTransaction;
FSource,ProgramInstall,PathDelete,FDest,KeyNewInstallProgram,FileToRun:string;
DeleteAfterInstall,BeforeInstallCopy,InstallAllUsers,PathCreate,CancelCopyFF:boolean;
OwnerForm:Tform;
NumCount,TypeOperation,step:integer;
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
step:=0;
FDQueryRun.sql.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC='''+TYPETASK+''' AND ID_PROC='''+inttostr(numberProc)+'''';
FDQueryRun.Open;
if TYPETASK='msi' then
 Begin
 step:=1;
 if FDQueryRun.FieldByName('FILEORFOLDER').AsString='File' then // если копируем только файл
   begin
   step:=2;
   FSource:= FDQueryRun.FieldByName('PATCH_PROC').AsString+#0+#0;
   ProgramInstall:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString); // строка запуска файла для установки
   PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0; // путь к удаляемому файлу после установки программы
   step:=3;
   end;
 if FDQueryRun.FieldByName('FILEORFOLDER').AsString='Folder' then // если копируем каталог
   begin
   step:=4;
   FSource:= ExtractFilePath(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;
   ProgramInstall:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString);
   PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;  // путь к удаляемому каталогу после установки программы
   step:=5;
   end;
   step:=6;
 if not FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean then // если не копируем файл или каталог перед установкой
 ProgramInstall:= FDQueryRun.FieldByName('PATCH_PROC').AsString; //то файл для установки берем из источника
 step:=7;
 FDest:=FDQueryRun.FieldByName('PATHCREATE').AsString+#0+#0; //
 PathCreate:=true;  /// проверять и создавать каталог при его отсутствии
 CancelCopyFF:=false;
 BeforeInstallCopy:=FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean; // копировать перед установкой
 DeleteAfterInstall:=FDQueryRun.FieldByName('DELETEAFTERINSTALL').AsBoolean;  // удалить после установки
 OwnerForm:=EditTask;
 NumCount:=1; //
 TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
 KeyNewInstallProgram:=FDQueryRun.FieldByName('KEY_MSI').AsString;
 InstallAllUsers:=true; // установка для всех пользователей
 step:=8;
 RunInstallMSI(FSource, // источник
               FDest,   // куда копировать
               namePC,  // имя компа
               UserName,
               PassWd,
               PathCreate, // создавать каталог и проверять
               CancelCopyFF,
               BeforeInstallCopy, // копировать перед установкой
               DeleteAfterInstall, // удалить после установки
               PathDelete,         // Какой каталог или файл удалять после установки
               OwnerForm,       // родительская форма прогресс бара, можно не указывать
               NumCount,
               TypeOperation,     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
               ProgramInstall,   // программа для установки
               KeyNewInstallProgram, // ключи установки программы
               InstallAllUsers, // устанавливать для всех пользователей  )
               NumberTask); // порядковый номер текущей задачи
 End;
  step:=9;
 if TYPETASK='proc' then
 Begin
    NumCount:=1;
    step:=10;
      if FDQueryRun.FieldByName('FILEORFOLDER').AsString='File' then // если копируем только файл
        begin
        step:=11;
        FSource:= FDQueryRun.FieldByName('FILESOURSE_PROC').AsString+#0+#0;
        FileToRun:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString); // строка запуска файла для установки
        PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+ExtractFileName(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+#0+#0; // путь к удаляемому файлу после установки программы
        end;
      if FDQueryRun.FieldByName('FILEORFOLDER').AsString='Folder' then  // если копируем весь корневой каталог
        begin
        step:=12;
        FSource:= ExtractFilePath(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+#0+#0;
        FileToRun:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+GetLastDir(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString);
        PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;  // путь к удаляемому каталогу после установки программы
        end;
    if not FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean then  FileToRun:= FDQueryRun.FieldByName('PATCH_PROC').AsString; //// если не копируем файл или каталог перед установкой то файл для установки берем из источника
    step:=13;
    FDest:=FDQueryRun.FieldByName('PATHCREATE').AsString+#0+#0;
    PathCreate:=true;
    CancelCopyFF:=false;
    BeforeInstallCopy:=FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean;;
    DeleteAfterInstall:=FDQueryRun.FieldByName('DELETEAFTERINSTALL').AsBoolean;
    OwnerForm:=EditTask;
    TypeOperation:=2;
    step:=14;
    MyNewProcess(FSource,  ///что копировать источник файл или каталог
    FDest, // куда копировать
    NamePC,     // компьютеров
    UserName,   // логин
    PassWd,    // пароль
    PathCreate, // проверка и создания директории
        CancelCopyFF,   // отменить или нет операцию копирования
    BeforeInstallCopy,     // копировать или нет перед установкой
    DeleteAfterInstall, // удалить дистрибутив после установки
    PathDelete,         // Какой каталог или файл удалять после установки
    OwnerForm,       // родительская форма прогресс бара, можно не указывать
    NumCount,
    TypeOperation,     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    FileToRun,   // файл для запуска
    NumberTask); // порядковый номер текущей задачи
   step:=15;
 End;

FDQueryRun.SQL.clear;   //очистить
step:=16;
FDQueryRun.Close;  /// закрыть нах после чтения
step:=17;
FDQueryRun.Free;
step:=18;
TransactionRun.Commit;
TransactionRun.Free;
step:=19;
Except
  on E:Exception do
     begin
     if Assigned(FDQueryRun) then FDQueryRun.Free;
     if Assigned(TransactionRun) then
     begin
     TransactionRun.Commit;
     TransactionRun.Free;
     end;
     Log_write(2,'TASK',' шаг '+inttostr(step)+' - Ошибка запуска установки программы(msi) или запуска процесса- '+e.Message);
     end;
   end;

end;

function TaskRun.createQuery:boolean;
begin
try
ConnectionThReadTask:=TFDConnection.Create(nil);
ConnectionThReadTask.DriverName:='FB';
ConnectionThReadTask.Params.database:=databaseName; // расположение базы данных
ConnectionThReadTask.Params.Add('server='+databaseServer);
ConnectionThReadTask.Params.Add('port='+databasePort);
ConnectionThReadTask.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThReadTask.Params.Add('CharacterSet=UTF8');
ConnectionThReadTask.Params.add('sqlDialect=3');
ConnectionThReadTask.Params.DriverID:=databaseDriverID;
ConnectionThReadTask.Params.UserName:=databaseUserName;
ConnectionThReadTask.Params.Password:=databasePassword;
ConnectionThReadTask.LoginPrompt:= false;  /// отображение диалога user password
ConnectionThReadTask.Connected:=true;
TransactionWriteTT:= TFDTransaction.Create(nil);
TransactionWriteTT.Connection:=ConnectionThReadTask;
TransactionWriteTT.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;  xiUnspecified
TransactionWriteTT.Options.AutoCommit:=false;
TransactionWriteTT.Options.AutoStart:=false;
TransactionWriteTT.Options.AutoStop:=false;
FDQueryWrite:= TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWriteTT;
FDQueryWrite.Connection:=ConnectionThReadTask;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTT;
FDQueryWriteTaskTable.Connection:=ConnectionThReadTask;
result:=true;
except on E: Exception do
begin
result:=false;
Log_write(3,'TASK',' Error Connect DataBase - '+e.Message);
end;
end;
end;

function TaskRun.CloseConnect:boolean;
begin
  try
  if Assigned(FDQueryWrite) then FDQueryWrite.Free;
  if Assigned(FDQueryWriteTaskTable) then  FDQueryWriteTaskTable.Free;
   if Assigned(TransactionWriteTT) then
    begin
    TransactionWriteTT.Rollback;
    TransactionWriteTT.Free;
    end;
    if Assigned(ConnectionThReadTask) then
    begin
    ConnectionThReadTask.Free;
    end;
  except on E: Exception do
  begin
   Log_write(3,'TASK',' Error Disconnect DataBase - '+e.Message);
   result:=false;
  end;
  end;
end;

function TaskRun.StrInDescription(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=br;
end;

function TaskRun.StrCopySourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=copy(br,1,pos(' -> ',br)-1);
end;

function TaskRun.StrCopyDestination(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(' -> ',br)+3);
result:=br;
end;

function TaskRun.StrDeleteSourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
 if AnsiPos(':',br)=2 then // если в пути назначения найдено двоеточие то запрашиваем правильность ввода пути, для копирования в сеть после диска ставится $
 begin
 delete(br,2,1); // удаляем символ :
 insert('$',br,2); // вставляем на его место $
 end;
 result:=br;
end;

function TaskRun.ConnectWMI (NamePC,User,Passwd:string):Boolean; // проверка доступности WMI
var
  FSWbemLocator   : OLEVariant;
  FWMIService   : OLEVariant;
begin
 Result:=False;
 try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Passwd,'','',128);
    Result:=True;
    FWMIService:=Unassigned;
    VariantClear(FWMIService);
    FSWbemLocator:=Unassigned;
    VariantClear(FSWbemLocator);
 except
  Log_write(1,'TASK',NamePC+' - RPС сервер не доступен');
  FWMIService:=Unassigned;
  VariantClear(FWMIService);
  FSWbemLocator:=Unassigned;
  VariantClear(FSWbemLocator);
  result:=false;
 end;
end;
                  // имя компа, пользователь, пароль, номер задачи, время ожидания компьютера в сети в минутах
function TaskRun.waitPing(namePC,User,passwd:string;NumTask,WaitTime:integer):boolean; // функция ожидания появления компьютера в сети
var
step,timeout:integer;
begin
try
timeout:=0;
while timeout<WaitTime*60 do   // ожидать пока таймаут меньше заданного времени (в секундах) ожидания
 Begin
 step:=1;
 if SelectedPing(namePC,pingtimeout,NumTask) and (ConnectWMI(namePC,User,Passwd)) then // если пингуется и доступен WMI
   begin
   step:=2;
   timeout:=WaitTime*60;
   result:=true;
   break;
   end
 else
   begin
   step:=3;
   Sleep(5000);
   timeout:=timeout+20; // прибавляем 15 секунд т.к на ожидание уходт 5 сек + ping (5сек) +доступность WMI (~10 сек)
   result:=false;
   end;
 step:=4;
 End;
except
  on E: Exception do
  begin
   result:=false;
   Log_write(1,'TASK',' шаг - '+inttostr(step)+' - Error WaitPing- '+e.Message);
  end;
end;
end;

function TaskRun.MinPreOper(namePC:string;PreOper,min:integer):integer; //сравниваем время которое прошло с момента выполнения предыдущей операции и время ожидания
var
br:string;
mydate:TdateTime;
gap,i,z:integer;
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThReadTask;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=ConnectionThReadTask;
TransactionWrite.StartTransaction;
FDQuery.SQL.Text:='SELECT RESULTTASK'+inttostr(PreOper)+' FROM '+RunNewTask+' WHERE PC_NAME='''+namePC+'''';
FDQuery.Open;
br:=FDQuery.FieldByName('RESULTTASK'+inttostr(PreOper)).AsString;
FDQuery.Close;
TransactionWrite.commit;
FDQuery.Free;
TransactionWrite.Free;
gap:=0;
z:=0;
Delete(br,length(br),length(br));  // удаляем поседний символ ')'
for I := length(br) downto 1 do   // ищем с конца строки символ '('
begin
  if br[i]='(' then
  begin
   z:=i;
   break;
  end;
end;
delete(br,1,z); // удаляем с начала строки до символа '('
 if trystrtodatetime(br,mydate) then // если br это дата
 begin
 if mydate<now then // если время выполнения предыдущего задания меньше чем текущее (а вот хз, может бля такое быть)
 gap:=MinutesBetween(now,mydate); // то высчитываем разницу в минутах
 end;
 if gap>=min then result:=0 // если разница в минутах больше запрошеного таймаута то ждать не стоит, выполняем действие
 else result:=min-gap; // иначе ждем разницу между запрошеным временем и уже прошедшим прошедшим

 except on E: Exception do // если ошибка то ждем ровно столько сколько запросили
 result:=min;
 end;
end;

procedure TaskRun.Execute;
var
countTask:integer;
CountPC:integer;
i,step,res:integer;
StartStop:boolean;
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
waitafter:boolean;
TimeOutAfter,TimeBeforeOff,TimeOutBefore:integer;
UserPC,Passwd:string;
begin
try
UserPC:='';
Passwd:='';
RunNewTask:=NameRunTASK;
createQuery; // создаем компоненты для работы с БД
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThReadTask;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=ConnectionThReadTask;

TransactionWrite.StartTransaction; /// читаем логин и пароль для текущего задания
FDQuery.SQL.Text:='SELECT USER_NAME,PASS_USER,SAVE_PASS,DESCRIPTION_TASK FROM TABLE_TASK WHERE NAME_TABLE='''+RunNewTask+'''';
FDQuery.open;
if FDQuery.FieldByName('USER_NAME').Value<>null then UserPC:=(FDQuery.FieldByName('USER_NAME').AsString);
if FDQuery.FieldByName('PASS_USER').Value<>null then Passwd:=(FDQuery.FieldByName('PASS_USER').AsString);
if FDQuery.FieldByName('DESCRIPTION_TASK').Value<>null then DesTask:=(FDQuery.FieldByName('DESCRIPTION_TASK').AsString);
if (FDQuery.FieldByName('SAVE_PASS').Value<>null) then
 begin
   if not (FDQuery.FieldByName('SAVE_PASS').AsBoolean) then // если не сохранять пароль после запуска задачи, то чистим поля
   UpdateTableTask(RunNewTask,'','','Подготовка',3,true,true);
 end;
TransactionWrite.Commit;
FDQuery.close;

TransactionWrite.StartTransaction;
FDQuery.SQL.Clear;
FDQuery.SQL.Text:='SELECT * FROM '+RunNewTask;
FDQuery.Open;
FDQuery.Last;/// переход к последней записи
CountPC:=FDQuery.RecordCount;     // количество записей  соответствует количеству компов
countTask:=FDQuery.FieldByName('TASK_QUANT').AsInteger; // количество поставленых заданий в текущей задаче
FDQuery.First;// переход к первой записи
StartStop:=true;
UpdateTableTask(RunNewTask,'','','Выполняется',1,true,true); /// статус задания, тип операции обновления таблицы, статус задачи, признак того что поток для таблицы запущен
Log_write(0,'TASK',' Запуск выполнения задачи '+DesTask);
For i := 0 to countTask-1 do // цикл по заданиям
BEGIN
if not StartStop then break; // если задача остановлена то выходим из цикла
 step:=6;
 FDQuery.First;  // Переход к первой строке, для выполнения новой задачи начиная с 1го компа
 while not FDQuery.Eof  do /// начинаем с первой строки пока не последняя запись
 BEGIN
  step:=7;
  if not StartStopTask(RunNewTask) then // если задача остановлена то переменной присваиваем false и выходим из цикла
   begin
   StartStop:=false;
   break;
   end;
  IF (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='WOL') THEN  // включаем компьютер WOL
    begin
    UpdateTableTask(RunNewTask,               // имя таблицы
    FDQuery.FieldByName('PC_NAME').AsString, //имя компа
    'Включение компьютера (WOL)',  // имя задания
    '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
    WakeOnLan(FDQuery.FieldByName('PC_NAME').AsString, // имя компа+
    StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString),
    i);  // порядковый номер выполняемой задачи
    end
  ELSE
  BEGIN
  // если предыдущая операция был перезагрузка или включение компа то ожидаем его появления в сети
  if (waitafter)then
  begin
  if i>0 then // Если это не первая операция в цикле, то вычистяем сколько времени прошло с момента выполнения последней операции
    begin       // в колонке результат выполнения предыдущей операции, есть время выполнения этой операции
    TimeOutAfter:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),TimeOutAfter); //вычисляем сколько времени ждать
    end;
  if TimeOutAfter>0 then // если вычисленное время больше 0 то ожидаем ровно столько времени
    begin
    UpdateTableTask(RunNewTask,               // имя таблицы
    FDQuery.FieldByName('PC_NAME').AsString, //имя компа
    'Ожидание подключения, компьютер не доступен',  // имя задания
    '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
     waitPing(FDQuery.FieldByName('PC_NAME').AsString, //имя компа
     UserPC,Passwd, // логин и пароль
     i,//номер задания
     TimeOutAfter); // таймаут в минутах ожидания компьютера
     end;
  end;

    if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='ShDown')   // если выключение
     then
      begin
        if (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger>0)and (i>0) then //если необходимо ждать перед выключением
        begin
         // в колонке результат выполнения предыдущей операции есть время выполнения этой операции
         TimeBeforeOff:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger); //вычисляем сколько времени ждать
         if TimeBeforeOff>0 then // если вычисленное время больше 0 то ожидаем ровно столько времени сколько осталось
         begin
         UpdateTableTask(RunNewTask,               // имя таблицы
         FDQuery.FieldByName('PC_NAME').AsString, //имя компа
         'Таймаут ожидания перед выключением компьютера',  // описание задания
         '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
         sleep(TimeBeforeOff*1000*60);// спим перед выключением
         end;
        end;
      UpdateTableTask(RunNewTask,               // имя таблицы
         FDQuery.FieldByName('PC_NAME').AsString, //имя компа
         'Выключение компьютера',  // описание задания
         '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
      if SelectedPing(FDQuery.FieldByName('PC_NAME').AsString,pingtimeout,i) then // если комп доступен то выключаем
       ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //пользователь и пароль
      4, // завершение работы
      i);   // порядковый номер выполняемой задачи
      end; //закончили выключать ПК

   if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='TimeOut')  then // таймаут ожидания до выполнения следующего задания
   begin
     if (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger>0) then //если таймаут больше ноля
        begin
        finditem('OK','Таймаут '+FDQuery.FieldByName('NumTask'+inttostr(i)).AsString+' мин.',FDQuery.FieldByName('PC_NAME').AsString,i);
         // в колонке результат выполнения предыдущей операции есть время выполнения этой операции
         TimeOutBefore:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger); //вычисляем сколько времени ждать
         if TimeOutBefore>0 then // если вычисленное время больше 0 то ожидаем ровно столько времени сколько осталось
         begin
         UpdateTableTask(RunNewTask,               // имя таблицы
         FDQuery.FieldByName('PC_NAME').AsString, //имя компа
         'Таймаут ожидания ('+inttostr(TimeOutBefore)+' мин.)',  // описание задания
         '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
         sleep(TimeOutBefore*1000*60);// таймаут перед следующей операцией
         end;
        end;
   end;

  ////////////////////////////////////////////////////////////////////
  if(SelectedPing(FDQuery.FieldByName('PC_NAME').AsString,pingtimeout,i))                     //если компьютер доступен
  and (FDQuery.FieldByName('STATUSTASK'+inttostr(i)).AsString<>'OK') then //  статус задания не OK(не выполнено)
  Begin
  UpdateTableTask(RunNewTask,               // имя таблицы
  FDQuery.FieldByName('PC_NAME').AsString, //имя компа
  FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString,  // имя задания
  '',0,true,true); /// статус задания, тип операции обновления таблицы, статус задачи  тут не влияет на выполнение т.к. операция 0. Признак того что поток запущен
   ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
   if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='msi') or
   (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='proc') then
    begin
    step:=8;
    ReadAndRunMsiOrProc
    (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger, // номер процесса в таблице  START_PROC_MSI
    i,                                           // порядковый номер выполняемой задачи
    FDQuery.FieldByName('PC_NAME').AsString, // имя компа
    FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString, // ну и тип, что это процесс или msi
    UserPC,Passwd); //пользователь и пароль
    end;
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='DelMSI' then // удаление программы MSI
     begin
     DeleteProgramMSI(FDQuery.FieldByName('PC_NAME').AsString, // имя компа+
     UserPC,Passwd,  // пользователь и пароль
     StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //выделяем имя программы для удаления из описания задачи
     STRtoB(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), // точное совпадение имени программы или нет
     i);  // порядковый номер выполняемой задачи
     end;
    ///////////////////////////////////////////////////////////////////////////////////////////////////////
    if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='KillProc' then // завершение процесса
     begin
     KillProcess(FDQuery.FieldByName('PC_NAME').AsString, // имя компа+
     UserPC,Passwd,  // пользователь и пароль
     StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //выделяем имя процесса для завершения  из описания задачи
     STRtoB(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), // точное совпадение имени процесса или нет
     i);  // порядковый номер выполняемой задачи
     end;

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Service' then // если управление службами
      begin
      ControlService(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //пользователь и пароль
      StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //выделяем имя службы  из описания задачи
      FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger,  // тип операции, 1 -запуск, 2- остановка, 3- удаление
      i);   // порядковый номер выполняемой задачи

      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Reset' then // если управление питанием
      begin
      ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //пользователь и пароль
      2, // перезагрузка
      i);   // порядковый номер выполняемой задачи

      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Logout' then // если управление питанием
      begin
      ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //пользователь и пароль
      0, // завершение сеанса
      i);   // порядковый номер выполняемой задачи
      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Activate' then // если активация office или windows
      begin
        case FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger of  // по типу операции //Product - Office и  Operating System
        1: InstallKeyWinOffice(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // активация Windows
        ,'Operating System',FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i);
        2: InstallKeyWinOffice(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // активация Office для windows 8/8.1/10...
        ,'Office',FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i);
        3: InstallKeyOfficeWin7(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // активация Office для windows 7
        ,FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i)
        end;
      end;
      //////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CopyFF' then // если функция копирования
      begin
        CopyDeleteFF(FDQuery.FieldByName('PC_NAME').AsString, //Имя компа
        UserPC,Passwd, //логин и пароль
        StrCopySourse(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, //что копируем
        StrCopyDestination(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, // куда копируем
        EditTask,
        2, // операция копирования
        false,  //возможность отмены
        strtobool(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), //создавать или нет каталог при его отсутствии
        i); // номер задачи
      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='DelFF' then // если удаляем
      begin
      CopyDeleteFF(FDQuery.FieldByName('PC_NAME').AsString, //Имя компа
        UserPC,Passwd, //логин и пароль
        '', //что копируем  необходимости нет
        StrDeleteSourse(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, // что удалять
        EditTask,
        3, // операция удаления
        false,  //возможность отмены
        false, //создавать или нет каталог при его отсутствии
        i); // номер задачи
      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////////
      if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sSubKey') //sSubKey работа с разделом реестра (создание или удаление)  NumTask определяет удаление или создание раздела реестра
      or (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sNameKey') //sNameKey работа по удалению ключа реестра
      or (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CreateKey') // создание ключа реестра, импортирование ключа из таблицы сохраненных ключей
      then
       begin
       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CreateKey' then
       finditem('RUN','Запущена операция создания ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);

       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sNameKey' then
       finditem('RUN','Запущена операция удаления ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);

       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sSubKey' then
       begin
        if FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger=0 then  finditem('RUN','Запущена операция удаления раздела реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        if FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger=1 then  finditem('RUN','Запущена операция создания раздела реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
       end;
       res:=RunTaskRegEdit.RunEditRegeditWindows(FDQuery.FieldByName('PC_NAME').AsString, //Имя компа
                              UserPC,Passwd, //логин и пароль
                              FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString, //тип выполняемого задания
                              FDQuery.FieldByName('NumTask'+inttostr(i)).AsString, // в NUMTASK записан номер импортируемого ключа в реестр/ создание или удаление раздела реестра
                              FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString // В NAMETASk передается путь к создаваемому или удаляемому разделу реестра/ удаляемому ключу реестра
                              );
        case res of
        0:finditem('OK','Операция успешно выполнена',FDQuery.FieldByName('PC_NAME').AsString,i);
        666: finditem('WARNING','Сервер RPC недоступен',FDQuery.FieldByName('PC_NAME').AsString,i);
        631: finditem('WARNING','Ошибка создания раздела реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        632: finditem('WARNING','Ошибка удаления раздела реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        633: finditem('WARNING','Ошибка удаления ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        637: finditem('WARNING','Ошибка создания ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        639: finditem('WARNING','Ошибка создания ключа DWORD реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        640: finditem('WARNING','Ошибка создания ключа QWORD реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        641: finditem('WARNING','Ошибка создания мультистрокового ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        642: finditem('WARNING','Ошибка создания расширенного ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        643: finditem('WARNING','Ошибка запуска импорта ключей на машины назначения',FDQuery.FieldByName('PC_NAME').AsString,i);
        644: finditem('WARNING','Ошибка импорта ключей из таблицы',FDQuery.FieldByName('PC_NAME').AsString,i);
        646: finditem('WARNING','Ошибка создания бинарного ключа реестра',FDQuery.FieldByName('PC_NAME').AsString,i);
        888: finditem('WARNING','Общая ошибка запуска функций работы с реестром',FDQuery.FieldByName('PC_NAME').AsString,i);
        else finditem('WARNING',SysErrorMessage(res),FDQuery.FieldByName('PC_NAME').AsString,i);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='RestoreWin' then // если работа с точками восстановления
      begin
        res:=RunTaskOsher.RestoreOrNewPoint
        (FDQuery.FieldByName('PC_NAME').AsString, //Имя компа
        UserPC,Passwd, //логин и пароль
        FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString, // дата или описание, или дата и описание при восстановлении по дате и описанию
        strtoint(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString));   // что делаем , 1-новая точка восстановления, 2 - запустить восстановление до даты, - 3 в восстановление по дате и описанию
        case res of
        0: finditem('OK','Операция успешно выполнена',FDQuery.FieldByName('PC_NAME').AsString,i);
        666: finditem('WARNING','Сервер RPC недоступен',FDQuery.FieldByName('PC_NAME').AsString,i);
        777:finditem('WARNING','Точка восстановления не найдена',FDQuery.FieldByName('PC_NAME').AsString,i);
        888:finditem('WARNING','Общая ошибка при работе с точками восстановления',FDQuery.FieldByName('PC_NAME').AsString,i);
        999:finditem('WARNING','Ошибка загрузки точек восстановления',FDQuery.FieldByName('PC_NAME').AsString,i);
        1111:finditem('WARNING','Ошибка включения восстановления на диске',FDQuery.FieldByName('PC_NAME').AsString,i);
        2222:finditem('WARNING','Ошибка создания точки восстановления',FDQuery.FieldByName('PC_NAME').AsString,i);//
        else  finditem('WARNING',SysErrorMessage(res),FDQuery.FieldByName('PC_NAME').AsString,i);
        end;
     end;


  End; //если комп доступен и задание не выполнено
 END; //иначе , когда операции кроме WakeOnLan

 FDQuery.Next;  // переходим на следующую строку
 END;// цикл по строкам (компам) в таблице

 if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Reset') or
 (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='WOL')  then
 begin
 waitafter:=true; // переменная, ожидать после перезагрузки или включения
 TimeOutAfter:=FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger;
 end
  else waitafter:=false; //значит операция была не перезагрузка и не включение ПК
 step:=9;

END; /// цикл по задачам

if StartStop then UpdateTableTask(RunNewTask,'','','Завершена',2,false,false) /// если задача не останавливалась принудительно то она завершилась
else UpdateTableTask(RunNewTask,'','','Остановлена',1,false,false);          // иначе ее остановили принудительно

FDQuery.SQL.clear;   //очистить
FDQuery.Close;  /// закрыть нах после чтения
FDQuery.Free;
FDQueryWrite.Close;
FDQueryWrite.Free;
FDQueryWriteTaskTable.Close;
FDQueryWriteTaskTable.Free;
TransactionWrite.Commit;
TransactionWrite.Free;
TransactionWriteTT.Free;
ConnectionThReadTask.Close;
ConnectionThReadTask.Free;
Log_write(0,'TASK',' Выполнение задачи '+DesTask+ ' завершено');
Except
  on E:Exception do
     begin
     Log_write(3,'TASK',' шаг - '+inttostr(step)+' - Fatal error- '+e.Message);
     if Assigned(TransactionWriteTT) then UpdateTableTask(RunNewTask,'','','Ошибка выполнения задачи',1,false,false);
     if Assigned(FDQueryWrite) then FDQueryWrite.Free;
     if Assigned(FDQuery) then  FDQuery.Free;
     if Assigned(FDQueryWriteTaskTable) then  FDQueryWriteTaskTable.Free;
     if Assigned(TransactionWrite) then
     begin
      TransactionWrite.Rollback;
      TransactionWrite.Free;
     end;
     if Assigned(TransactionWriteTT) then
     begin
      TransactionWriteTT.Rollback;
      TransactionWriteTT.Free;
     end;
     if Assigned(ConnectionThReadTask) then
     begin
      ConnectionThReadTask.Free;
     end;

     end;
   end;
end;

end.

