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
    function TableExists(TableName: string): Boolean; /// ���������� �� �������
    function CleanDeleteTask(s:string):boolean; // ������� ���� �� ��������� �� �������� �����.
  private
    { Private declarations }
  public
    function �reateTablForRegEdit:boolean; // ������� ��� �������� ������ �������
    function �reateTablForMyProcess:boolean; // ������� ��� �������� ������ � ����� ������� ��������� ����������
    function �reateTablForDriverPrint(TName:string):boolean;
    function �reateTablForTASK(TName,TaskType,NumTask,ResultTask,NameTask,StatusTask:string):boolean;
    Function connectDataBase:boolean; // ��������� ���������� � �����
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
var f:TStringList;        /// ������� ������ � ��� ����
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
ConnectionThRead.Params.database:=databaseName; // ������������ ���� ������
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.LoginPrompt:= false;  /// ����������� ������� user password
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

TransactionRead.StartTransaction; // ��������
FDQueryRead.SQL.Text:='SELECT * FROM TABLE_TASK WHERE STATUS_TASK=''delete''';
FDQueryRead.Open;
 while not FDQueryRead.Eof do
 begin
 try
   del:=false;
   NameTable:=FDQueryRead.FieldByName('NAME_TABLE').AsString; // �������� ������
   TransactionDeleteTask.StartTransaction; // �������� ���������� ��� ��������
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
  TransactionDeleteTask.commit; //��������� ��� ����� ��������

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
Log_write('TASK','������ ������� ����� :'+E.Message);
result:=false;
end;
end;
ConnectionThRead.Connected:=false;
ConnectionThRead.Free;
end;


function TDataM.TableExists(TableName: string): Boolean; /// ���������� �� �������
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




function TDataM.EXPTCurrentPC(s:string;ex:integer):boolean; //// �������� � ���������� ���������� �� ������ ������������
var                                                        //// 12 - ��������� �� ������ ������������, �.�. ���������, 0- �������� � ������ ������������
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
      Log_write('Data',s+': ������ EXPT :'+E.Message);
      end;
  end;
end;

Function TDataM.connectDataBase:boolean;
begin
ConnectionDB.Params.Clear;     ///������ ���������
ConnectionDB.Params.database:=databaseName; // ������������ ���� ������
ConnectionDB.Params.Add('server='+databaseServer);
ConnectionDB.Params.Add('port='+databasePort);
ConnectionDB.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionDB.Params.Add('CharacterSet=UTF8');
ConnectionDB.Params.add('sqlDialect=3');
ConnectionDB.Params.DriverID:=databaseDriverID;
ConnectionDB.Params.UserName:=databaseUserName;
ConnectionDB.Params.Password:=databasePassword;
ConnectionDB.Connected:=databaseconnected;
ConnectionDB.LoginPrompt:= not databaseconnected;  /// ����������� ������� user password
//FDTransactionRead.StartTransaction; /// ���� ���������� ���������� �� ��� ������� ��� ������, ������ ��� �������������� ����
result:=true;
try
if ConnectionDB.Connected then
begin
frmDomainInfo.readdb; /// ���� ���������� ������������ �� ��������� ������� � �������������
frmDomainInfo.readdbSoft; /// ���� ���������� ������������ �� ��������� ������� � ������
frmDomainInfo.readinfoforpcDB; //���� ���������� ������������ �� ������ ������������� �� ���� ������  � ListView8
�reateTablForMyProcess; // �������� ������� ���� �� ���, ��� �������� ������ � ����� ������� ��������� ����������
end;
Except
   on E: Exception do
      begin
      Log_write('Data','������ Create :'+E.Message);
      result:=false;
      end;
  end;
end;


procedure TDataM.DataModuleCreate(Sender: TObject);
begin
connectDataBase; // ��������� ����������
end;

procedure TDataM.DataModuleDestroy(Sender: TObject);
begin
    FDTransactionWriteEXPT.Options.DisconnectAction:=xdCommit;
    FDTransactionReadsoft.Options.DisconnectAction:=xdCommit;
    FDTransactionReadSoftSelectPC.Options.DisconnectAction:=xdCommit;
    FDTransactionSelectRead.Options.DisconnectAction:=xdCommit;
    FDTransactionReadListPC.Options.DisconnectAction:=xdCommit;
end;

function TDataM.�reateTablForRegEdit:boolean; // ������� ��� �������� ������ �������
begin
try
if ConnectionDB.Connected then /// ���� ���������� ������������
begin
if TableExists('REGEDIT_KEY') then // ���� ������� ���������� �� ��� �� ���������
begin
 result:=true;
 exit;
end;
FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
FDScriptTempTabl.SQLScripts[0].SQL.Clear;
FDScriptTempTabl.SQLScripts[0].Name:='RegEditK';
///////////////////////////////////////////////////////////////// ���������
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_REGEDIT_KEY START WITH 0 INCREMENT BY 1;');
///////////////////////////////////////////////////////////////////////////////////�������
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE REGEDIT_KEY (');
FDScriptTempTabl.SQLScripts[0].SQL.Add('ID_KEY  INTEGER NOT NULL,');  //
FDScriptTempTabl.SQLScripts[0].SQL.Add('Description_key  VARCHAR(1000),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('NamePC  VARCHAR(250),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('hDefKey  VARCHAR(50),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sSubKeyName     VARCHAR(3000),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sValue     VARCHAR(8190),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('TypeKey  VARCHAR(50),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('sValueName     VARCHAR(200) );');
///////////////////////////////////////////////////////////////////////// ������
FDScriptTempTabl.SQLScripts[0].SQL.Add('ALTER TABLE REGEDIT_KEY ADD CONSTRAINT PK_REGEDIT_KEY PRIMARY KEY (ID_KEY);');

///////////////////////////////////////////////////////////////�������
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
    frmDomainInfo.memo1.Lines.Add('������ ��������  ������� REGEDIT_KEY "'+E.Message+'"');
    exit;
    end;
  end;
end;

function TDataM.�reateTablForMyProcess:boolean; // ������� ��� �������� ������ � ����� ������� ��������� ����������
begin
try
if ConnectionDB.Connected then /// ���� ���������� ������������
begin
if TableExists('MY_PROCESS') then // ���� ������� ���������� �� ��� �� ���������
begin
 result:=true;
 exit;
end;
FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
FDScriptTempTabl.SQLScripts[0].SQL.Clear;
FDScriptTempTabl.SQLScripts[0].Name:='MYPROCESS';
///////////////////////////////////////////////////////////////// ���������
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MY_PROCESS_ID START WITH 0 INCREMENT BY 1;');
///////////////////////////////////////////////////////////////////////////////////�������
FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE MY_PROCESS (');
FDScriptTempTabl.SQLScripts[0].SQL.Add('"KEY"         INTEGER NOT NULL,');  //
FDScriptTempTabl.SQLScripts[0].SQL.Add('PATH_FILE     VARCHAR(300),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('ARG_FILE      VARCHAR(300),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('NAME_PROCESS  VARCHAR(150),'); //
FDScriptTempTabl.SQLScripts[0].SQL.Add('LOG_PASS      BOOLEAN);');
///////////////////////////////////////////////////////////////////////// ������
FDScriptTempTabl.SQLScripts[0].SQL.Add('ALTER TABLE MY_PROCESS ADD PRIMARY KEY ("KEY");');

///////////////////////////////////////////////////////////////�������
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
    frmDomainInfo.memo1.Lines.Add('������ ��������  ������� MY_PROCESS "'+E.Message+'"');
    exit;
    end;
  end;
end;


function TDataM.�reateTablForDriverPrint(TName:string):boolean;
begin
try
if datam.ConnectionDB.Connected then /// ���� ���������� ������������
begin
if TableExists('DRIVER_PRINT') then // ���� ������� ���������� �� ��� �� ���������
begin
 result:=true;
 exit;
end;
dataM.FDScriptTempTabl.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Clear;
dataM.FDScriptTempTabl.SQLScripts[0].Name:='DrvPrint';  //;
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('CREATE TABLE DRIVER_PRINT (');   //GLOBAL TEMPORARY - ��������� �������
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('ID_DRV  INTEGER,'); //ID_EVENT  INTEGER
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_PATCH     VARCHAR(1500) ,');  /// ���� � inf �����
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DLL     VARCHAR(1500) ,');   // ���� � DLL ����� ��������
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DLL_TRUE     VARCHAR(2) ,');  // ������ ��� ��� DLL ���� ��� ���������
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('SECTION_INF    VARCHAR(200) ,'); /// ������ inf ����� ��� �����������
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_CLASS     VARCHAR(200) ,'); /// ����� ��������
dataM.FDScriptTempTabl.SQLScripts[0].SQL.Add('DRV_DESCRIPT     VARCHAR(500) ,'); // �������� ������� ���� ���������
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
    frmDomainInfo.memo1.Lines.Add('������ ��������  ������� DRIVER_PRINT "'+E.Message+'"');
    exit;
    end;
  end;
end;

function TDataM.�reateTablForTASK(TName,TaskType,NumTask,ResultTask,NameTask,StatusTask:string):boolean; /// ������� �������� ����� ������� ��� ������
var
FDQueryAdd:TFDScript;
TransactionAdd: TFDTransaction;
begin
try
if datam.ConnectionDB.Connected then /// ���� ���������� ������������
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
if TableExists(Tname) and(TaskType<>'') then // ���� ������� ���������� � ���� �� ������ �� ��� �� ��������� ,
  begin                    // ������ ��������� � ������� ����� �������
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
  FDQueryAdd.SQLScripts[0].SQL.Add('TASK_QUANT  INTEGER,'); /// ����� ���������� �����
  FDQueryAdd.SQLScripts[0].SQL.Add('TASK_COMPLETED  INTEGER,'); // ���������� ���������� �����
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
//TransactionAdd.Connection.Commit;  // ��������� �� ��� ������ � ����
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
    Log_write('CreateTask',TName+': ������  :'+E.Message);
    frmDomainInfo.memo1.Lines.Add('��������� ������ ���������� ������� ');
    exit;
    end;
  end;
end;

end.
