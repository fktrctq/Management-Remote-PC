unit MyDM;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Phys.FBDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.FB, Data.DB, FireDAC.Comp.Client, FireDAC.Phys.IBBase,VCL.Forms,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt,
  FireDAC.Comp.DataSet,System.StrUtils, FireDAC.Comp.ScriptCommands,
  FireDAC.Stan.Util, FireDAC.Comp.Script, FireDAC.VCLUI.Script, FireDAC.Comp.UI,
  FireDAC.ConsoleUI.Script;

type
  TDataM = class(TDataModule)
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    ConnectionDB: TFDConnection;
    FDTransactionSelectRead: TFDTransaction;
    FDTransactionUpdate: TFDTransaction;
    FDQueryReadSoft: TFDQuery;
    DataSourceSoft: TDataSource;
    FDSelectReadSoft: TFDQuery;
    FDQueryLoadListPC: TFDQuery;
    FDTransactionReadListPC: TFDTransaction;
    FDReadSoftSelectPC: TFDQuery;
    FDQueryWriteEXPT_PC: TFDQuery;
    FDTransactionWriteEXPT: TFDTransaction;
    FDTransactionReadsoft: TFDTransaction;
    FDTransactionReadSoftSelectPC: TFDTransaction;
    FDQueryReadPropertiesPC: TFDQuery;
    FDTransactionReadPropPC: TFDTransaction;
    FDQueryFindViolations: TFDQuery;
    FDTransactionFindViolat: TFDTransaction;
    FDTransactionQuery3: TFDTransaction;
    FDQueryRead3: TFDQuery;
    DataSource2: TDataSource;
    DataSource1: TDataSource;
    DataSourceSelectSort: TDataSource;
    FDTransactionSort: TFDTransaction;
    FDQuerySelectSort: TFDQuery;
    FDQuerySort: TFDQuery;
    DataSourceSort: TDataSource;
    FDTransactionSelectSort: TFDTransaction;
    FDQueryRead2: TFDQuery;
    DataSourceViolation: TDataSource;
    FDScriptTempTabl: TFDScript;
    FDTransactionTempTabl: TFDTransaction;
    FDQueryTempTabl: TFDQuery;
    FDGUIxScriptDialog1: TFDGUIxScriptDialog;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    function EXPTCurrentPC(s:string;ex:integer):boolean;
    function TableExists(TableName: string): Boolean; /// существует ли таблица
    function CleanDeleteTask(s:string):boolean; // очистка базы от помеченых на удаление задач.
  private
    { Private declarations }
  public
    function СreateTablForRegEdit:boolean; // таблица для создания ключей реестра
    function СreateTablForMyProcess:boolean; // таблица для хранения ключей и строк запуска сторонних приложений
    function СreateTablForDriverPrint(TName:string):boolean;
    function СreateTablForTASK(TName,TaskType,NumTask,ResultTask,NameTask,StatusTask:string):boolean;
    Function connectDataBase:boolean; // установка соединения с базой
  end;

var
  DataM: TDataM;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}
 uses
uMain;
{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}
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


function TDataM.CleanDeleteTask(s:string):boolean;
var
FDQueryRead:TFDQuery;
FDQueryDeleteTask:TFDQuery;
TransactionRead: TFDTransaction;
TransactionDeleteTask: TFDTransaction;
ConnectionThread: TFDConnection;
NameTable:string;
del:boolean;
begin
try
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
ConnectionThRead.LoginPrompt:= false;  /// отображение диалога user password
ConnectionThRead.Connected:=true;
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=ConnectionThRead;
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRead.Options.AutoCommit:=false;
TransactionRead.Options.AutoStart:=false;
TransactionRead.Options.AutoStop:=false;
TransactionRead.Options.ReadOnly:=true;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=ConnectionThRead;


TransactionDeleteTask:= TFDTransaction.Create(nil);
TransactionDeleteTask.Connection:=ConnectionThRead;
TransactionDeleteTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionDeleteTask.Options.AutoCommit:=false;
TransactionDeleteTask.Options.AutoStart:=false;
TransactionDeleteTask.Options.AutoStop:=false;
FDQueryDeleteTask:= TFDQuery.Create(nil);
FDQueryDeleteTask.Transaction:=TransactionDeleteTask;
FDQueryDeleteTask.Connection:=ConnectionThRead;

TransactionRead.StartTransaction; // стартуем
FDQueryRead.SQL.Text:='SELECT * FROM TABLE_TASK WHERE STATUS_TASK=''delete''';
FDQueryRead.Open;
 while not FDQueryRead.Eof do
 begin
 try
   del:=false;
   NameTable:=FDQueryRead.FieldByName('NAME_TABLE').AsString; // описание задачи
   TransactionDeleteTask.StartTransaction; // стартуем транзакцию для удаления
   FDQueryDeleteTask.SQL.clear;
   FDQueryDeleteTask.SQL.Text:='DROP TABLE '+nametable;
   FDQueryDeleteTask.ExecSQL;
   del:=true;
  except on E: Exception do del:=false; end;
  if del then
  begin
  FDQueryDeleteTask.SQL.clear;
  FDQueryDeleteTask.SQL.Text:='DELETE FROM TABLE_TASK WHERE NAME_TABLE='''+nametable+'''';
  FDQueryDeleteTask.ExecSQL;
  end;
  TransactionDeleteTask.commit; //сохраняем все после удаления

 FDQueryRead.Next;
 end;
FDQueryRead.SQL.clear;
FDQueryRead.Close;
FDQueryRead.Free;
TransactionRead.Commit;
TransactionRead.Free;
FDQueryDeleteTask.Close;
FDQueryDeleteTask.Free;
TransactionDeleteTask.Free;
result:=true;
except on E: Exception do
begin
if assigned(FDQueryRead) then FDQueryRead.Free;
if assigned(TransactionRead) then
 begin
 TransactionRead.Rollback;
 TransactionRead.Free;
 end;
if assigned(TransactionDeleteTask) then
 begin
 TransactionDeleteTask.Rollback;
 TransactionDeleteTask.Free;
 end;
Log_write('TASK','Ошибка очистки задач :'+E.Message);
result:=false;
end;
end;
ConnectionThRead.Connected:=false;
ConnectionThRead.Free;
end;


function TDataM.TableExists(TableName: string): Boolean; /// существует ли таблица
var
Tables: TStrings;
begin
     Tables := TStringList.Create;
     try
      ConnectionDB.GetTableNames('','','',Tables,[osMy],[tkTable],true);
      Result := Tables.IndexOf(TableName) <> -1;
      //frmDomainInfo.Memo1.Lines:=Tables;
     finally
       Tables.Free;
     end;
end;




function TDataM.EXPTCurrentPC(s:string;ex:integer):boolean; //// включени и исключение компьютера из списка сканирования
var                                                        //// 12 - исключить из списка сканирования, т.е. выключить, 0- включить в список сканирования
i:integer;
x:string;
begin
try
x:='';
FDTransactionWriteEXPT.StartTransaction;
FDQueryWriteEXPT_PC.SQL.Clear;
FDQueryWriteEXPT_PC.SQL.Text:='update or insert into MAIN_PC (PC_NAME,EXPT_PC) VALUES'
+' ('''+s+''','''+inttostr(ex)+''') MATCHING (PC_NAME)';
FDQueryWriteEXPT_PC.ExecSQL;
FDQueryWriteEXPT_PC.SQL.Clear;
FDQueryWriteEXPT_PC.Close;
FDTransactionWriteEXPT.Commit;
result:=true;
Except
   on E: Exception do
      begin
      result:=false;
      Log_write('Data',s+': Ошибка EXPT :'+E.Message);
      end;
  end;
end;

Function TDataM.connectDataBase:boolean;
begin
ConnectionDB.Params.Clear;     ///чистим параметры
ConnectionDB.Params.database:=databaseName; // расположение базы данных
ConnectionDB.Params.Add('server='+databaseServer);
ConnectionDB.Params.Add('port='+databasePort);
ConnectionDB.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionDB.Params.Add('CharacterSet=UTF8');
ConnectionDB.Params.add('sqlDialect=3');
ConnectionDB.Params.DriverID:=databaseDriverID;
ConnectionDB.Params.UserName:=databaseUserName;
ConnectionDB.Params.Password:=databasePassword;
ConnectionDB.Connected:=databaseconnected;
ConnectionDB.LoginPrompt:= not databaseconnected;  /// отображение диалога user password
//FDTransactionRead.StartTransaction; /// если стартовать транзакцию то хуй увидишь что удалил, только при переподлючении бузы
result:=true;
try
if ConnectionDB.Connected then
begin
frmDomainInfo.readdb; /// если соединение установленно то считываем таблицу с конфигурацией
frmDomainInfo.readdbSoft; /// если соединение установленно то считываем таблицу с софтом
frmDomainInfo.readinfoforpcDB; //если соединение установленно то читаем пользователей из базы данных  в ListView8
СreateTablForMyProcess; // создание таблицы если ее нет, для хранения ключей и строк запуска сторонних приложений
end;
Except
   on E: Exception do
      begin
      Log_write('Data','Ошибка Create :'+E.Message);
      result:=false;
      end;
  end;
end;


procedure TDataM.DataModuleCreate(Sender: TObject);
begin
connectDataBase; // установка соединения
end;

procedure TDataM.DataModuleDestroy(Sender: TObject);
begin
    FDTransactionWriteEXPT.Options.DisconnectAction:=xdCommit;
    FDTransactionReadsoft.Options.DisconnectAction:=xdCommit;
    FDTransactionReadSoftSelectPC.Options.DisconnectAction:=xdCommit;
    FDTransactionSelectRead.Options.DisconnectAction:=xdCommit;
    FDTransactionReadListPC.Options.DisconnectAction:=xdCommit;
end;

function TDataM.СreateTablForRegEdit:boolean; // таблица для создания ключей реестра
begin
try
if ConnectionDB.Connected then /// если соединение установленно
begin
if TableExists('REGEDIT_KEY') then // если таблица существует то нах ее создавать
begin
 result:=true;
 exit;
end;
FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
FDScriptTempTabl.SQLScripts[0].SQL.Clear;
FDScriptTempTabl.SQLScripts[0].Name:='RegEditK';
///////////////////////////////////////////////////////////////// генератор
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_REGEDIT_KEY START WITH 0 INCREMENT BY 1;');
///////////////////////////////////////////////////////////////////////////////////таблица
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE REGEDIT_KEY (');
FDScriptTempTabl.SQLScripts[0].SQL.Add('ID_KEY  INTEGER NOT NULL,');  //
FDScriptTempTabl.SQLScripts[0].SQL.Add('Description_key  VARCHAR(1000),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('NamePC  VARCHAR(250),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('hDefKey  VARCHAR(50),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sSubKeyName     VARCHAR(3000),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sValue     VARCHAR(8190),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('TypeKey  VARCHAR(50),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sValueName     VARCHAR(200) );');
///////////////////////////////////////////////////////////////////////// индекс
FDScriptTempTabl.SQLScripts[0].SQL.Add('ALTER TABLE REGEDIT_KEY ADD CONSTRAINT PK_REGEDIT_KEY PRIMARY KEY (ID_KEY);');

///////////////////////////////////////////////////////////////тригеры
FDScriptTempTabl.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER REGEDIT_KEY_BI0 FOR REGEDIT_KEY ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('AS begin ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('if (new.id_key is null) then ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('new.id_key = gen_id(GEN_FOR_REGEDIT_KEY,1);');
FDScriptTempTabl.SQLScripts[0].SQL.Add('end');
FDScriptTempTabl.SQLScripts[0].SQL.Add('^');
FDScriptTempTabl.SQLScripts[0].SQL.Add('SET TERM ; ^');
//FDScriptTempTabl.ValidateAll;
if FDScriptTempTabl.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  end;
end;
except
  on E:Exception do
    begin
    result:=false;
    frmDomainInfo.memo1.Lines.Add('Ошибка создания  таблицы REGEDIT_KEY "'+E.Message+'"');
    exit;
    end;
  end;
end;

function TDataM.СreateTablForMyProcess:boolean; // таблица для хранения ключей и строк запуска сторонних приложений
begin
try
if ConnectionDB.Connected then /// если соединение установленно
begin
if TableExists('MY_PROCESS') then // если таблица существует то нах ее создавать
begin
 result:=true;
 exit;
end;
FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
FDScriptTempTabl.SQLScripts[0].SQL.Clear;
FDScriptTempTabl.SQLScripts[0].Name:='MYPROCESS';
///////////////////////////////////////////////////////////////// генератор
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MY_PROCESS_ID START WITH 0 INCREMENT BY 1;');
///////////////////////////////////////////////////////////////////////////////////таблица
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE MY_PROCESS (');
FDScriptTempTabl.SQLScripts[0].SQL.Add('"KEY"         INTEGER NOT NULL,');  //
FDScriptTempTabl.SQLScripts[0].SQL.Add('PATH_FILE     VARCHAR(300),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('ARG_FILE      VARCHAR(300),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('NAME_PROCESS  VARCHAR(150),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('LOG_PASS      BOOLEAN);');
///////////////////////////////////////////////////////////////////////// индекс
FDScriptTempTabl.SQLScripts[0].SQL.Add('ALTER TABLE MY_PROCESS ADD PRIMARY KEY ("KEY");');

///////////////////////////////////////////////////////////////тригеры
FDScriptTempTabl.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MY_PROCESS_BI FOR MY_PROCESS ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('AS begin ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('if (new."KEY" is null) then ');
FDScriptTempTabl.SQLScripts[0].SQL.Add('new."KEY" = gen_id(gen_my_process_id,1);');
FDScriptTempTabl.SQLScripts[0].SQL.Add('end');
FDScriptTempTabl.SQLScripts[0].SQL.Add('^');
FDScriptTempTabl.SQLScripts[0].SQL.Add('SET TERM ; ^');

if FDScriptTempTabl.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  end;
end;
except
  on E:Exception do
    begin
    result:=false;
    frmDomainInfo.memo1.Lines.Add('Ошибка создания  таблицы MY_PROCESS "'+E.Message+'"');
    exit;
    end;
  end;
end;


function TDataM.СreateTablForDriverPrint(TName:string):boolean;
begin
try
if datam.ConnectionDB.Connected then /// если соединение установленно
begin
if TableExists('DRIVER_PRINT') then // если таблица существует то нах ее создавать
begin
 result:=true;
 exit;
end;
dataM.FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Clear;
dataM.FDScriptTempTabl.SQLScripts[0].Name:='DrvPrint';  //;
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE DRIVER_PRINT (');   //GLOBAL TEMPORARY - временная таблица
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('ID_DRV  INTEGER,'); //ID_EVENT  INTEGER
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_PATCH     VARCHAR(1500) ,');  /// путь к inf файлу
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DLL     VARCHAR(1500) ,');   // Путь к DLL файлу драйвера
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DLL_TRUE     VARCHAR(2) ,');  // Указан или нет DLL файл при установке
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('SECTION_INF    VARCHAR(200) ,'); /// раздел inf файла или архитектура
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_CLASS     VARCHAR(200) ,'); /// класс драйвера
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DESCRIPT     VARCHAR(500) ,'); // описание данного типа установки
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_NAME     VARCHAR(200));'); //  ON COMMIT PRESERVE ROWS
//dataM.FDScriptTempTabl.ValidateAll;
if dataM.FDScriptTempTabl.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  end;
end;
except
  on E:Exception do
    begin
    result:=false;
    frmDomainInfo.memo1.Lines.Add('Ошибка создания  таблицы DRIVER_PRINT "'+E.Message+'"');
    exit;
    end;
  end;
end;

function TDataM.СreateTablForTASK(TName,TaskType,NumTask,ResultTask,NameTask,StatusTask:string):boolean; /// функция создания новой таблицы для задачи
var
FDQueryAdd:TFDScript;
TransactionAdd: TFDTransaction;
begin
try
if datam.ConnectionDB.Connected then /// если соединение установленно
Begin
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionAdd.Options.AutoCommit:=False;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
FDQueryAdd:= TFDScript.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=ConnectionDB;
TransactionAdd.StartTransaction;
if TableExists(Tname) and(TaskType<>'') then // если таблица существует и поля не пустые то нах ее создавать ,
  begin                    // просто обновляем и создаем новые столбцы
  //FDQueryAdd.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
  FDQueryAdd.SQLScripts.Add;
  FDQueryAdd.SQLScripts[0].SQL.Clear;
  FDQueryAdd.SQLScripts[0].Name:='TableTask';  //;
  FDQueryAdd.SQLScripts[0].SQL.Add('ALTER TABLE '+TName+' ADD '+TaskType+' VARCHAR(20),');///
  FDQueryAdd.SQLScripts[0].SQL.Add('ADD '+NumTask+'     INTEGER,');
  FDQueryAdd.SQLScripts[0].SQL.Add('ADD '+ResultTask+'     VARCHAR(250),');
  FDQueryAdd.SQLScripts[0].SQL.Add('ADD '+NameTask+'     VARCHAR(250),');
  FDQueryAdd.SQLScripts[0].SQL.Add('ADD '+StatusTask+'     VARCHAR(10);');
  FDQueryAdd.ValidateAll;
  end
else
  begin
  //FDQueryAdd.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
  FDQueryAdd.SQLScripts.Add;
  FDQueryAdd.SQLScripts[0].SQL.Clear;
  FDQueryAdd.SQLScripts[0].Name:='TableTask';  //;
  FDQueryAdd.SQLScripts[0].SQL.Add('CREATE TABLE '+TName+' (');
  FDQueryAdd.SQLScripts[0].SQL.Add('ID_PC  INTEGER,'); //ID
  FDQueryAdd.SQLScripts[0].SQL.Add('TASK_QUANT  INTEGER,'); /// общее количество задач
  FDQueryAdd.SQLScripts[0].SQL.Add('TASK_COMPLETED  INTEGER,'); // количество выполненых задач
  FDQueryAdd.SQLScripts[0].SQL.Add('PC_NAME     VARCHAR(80));'); //  ON COMMIT PRESERVE ROWS
  FDQueryAdd.ValidateAll;
  end;
if FDQueryAdd.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  end;
//TransactionAdd.Connection.Commit;  // сохраняем то что внесли в базу
TransactionAdd.Commit;
FDQueryAdd.Free;
TransactionAdd.Free;
End;
except
  on E:Exception do
    begin
    if Assigned(FDQueryAdd) then FDQueryAdd.Free;
      if Assigned(TransactionAdd) then
      begin
      TransactionAdd.Rollback;
      TransactionAdd.Free;
      end;
    result:=false;
    Log_write('CreateTask',TName+': Ошибка  :'+E.Message);
    frmDomainInfo.memo1.Lines.Add('Достигнут предел количества заданий ');
    exit;
    end;
  end;
end;

end.
