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
    function ReadAndRunMsiOrProc (numberProc,NumberTask:integer;namePC,TYPETASK,UserName,PassWd:string):boolean; // ������� ������� ��������� ��������� ��� ������� ��������
    Function InstallKeyOfficeWin7(StringKey,namePC,User,Passwd:string;NumberTask:integer):boolean;  // ��������� ����� office � windows 7
    function ActivateOfficeWin7(namePC,User,Passwd:string;NumberTask:integer):boolean;  // ��������� ����� � windows 7
    Function InstallKeyWinOffice(StringKey,Product,namePC,User,Passwd:string;NumberTask:integer):boolean; // ��������� ����� Windows ��� ���� ������ , Office ��� ����� Windows 7
    function ActivateProduct (Product,namePC,User,Passwd:string;NumberTask:integer):boolean; // ��������� Windows ��� ���� ������ , Office ��� ����� Windows 7
    function ControlService(namePC,User,Passwd,NameService:string;Oper,NumberTask:integer):boolean; // ���������� ��������
    function DescriptionLicStatus(num:string):string;
    function RefreshLicinseWin(namePC,user,passwd:string;NumberTask:integer):boolean;  // ���������� ������� �������� windows
    function refreshOfficeLic(NamePC,User,Passwd:string):boolean;     /// ���������� ������� �������� ����
    function MyNewProcess(FSource :String;   ///��� ���������� �������� ���� ��� �������
    FDest:string; // ���� ����������
    NamePC:string;     // ������ �����������
    UserName:string;   // �����
    PassWd:string;     // ������
    PathCreate:boolean;  // �������� � �������� ����������
        CancelCopyFF:boolean;   // �������� ��� ��� �������� �����������
    BeforeInstallCopy:boolean;     // ���������� ��� ��� ����� ����������
    DeleteAfterInstall:boolean; // ������� ����������� ����� ���������
    PathDelete:string;         // ����� ������� ��� ���� ������� ����� ���������
    OwnerForm:Tform;        // ������������ ����� �������� ����, ����� �� ���������
    NumCount:integer;
    TypeOperation:integer;     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
    FileToRun:String;   // ���� ��� �������
    NumberTask:integer):boolean;
    function RunInstallMSI( FSource :String;   /// �������� ����
    FDest:string;      //���� ����������
    NamePC:string;     // ������ �����������
    UserName:string;   // �����
    PassWd:string;     // ������
    PathCreate:boolean;  // �������� � �������� ����������
    CancelCopyFF:boolean;   // �������� ��� ��� �������� �����������
    BeforeInstallCopy:boolean;     // ���������� ��� ��� ����� ����������
    DeleteAfterInstall:boolean; // ������� ����������� ����� ���������
    PathDelete:string;         // ����� ������� ��� ���� ������� ����� ���������
    OwnerForm:Tform;        // ������������ ����� �������� ����, ����� �� ���������
    NumCount:integer;
    TypeOperation:integer;     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
    ProgramInstall:String;   // ��������� ��� ���������
    KeyNewInstallProgram:string; // ����� ��������� ���������
    InstallAllUsers:boolean; // ������������� ��� ���� �������������
    NumberTask:integer):boolean;  // ����� ������� � ������
    function WakeOnLan(NamePC,BRaddress:string;NumTASK:integer):boolean;
    function KillProcess(NamePC,UserName,Passwd,NameProcess:string;equallyName:boolean;
NumTASK:integer):boolean;
    function DeleteProgramMSI(NamePC,UserName,Passwd,NameProgram:string;
equallyName:boolean;NumTASK:integer):boolean;    ///// �������� ��������� msi
    function ShotDownResetCloseSession(namePC,Myuser,MyPasswd:string;myShutdown,NumTask:integer):boolean; // ������������/���������� ������/���������� ������
    function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // ����������� � ������
     function CopyDeleteFF(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // ����������� ��� �������
     function FindAddcreateDir(path,NamePC:string;NumTask:integer):boolean;// �������� � �������� ����������
     function Log_write(Level:integer;fname, text:string):string;
     function finditem(StatusTask,ResultTask,namePC:string;NumTask:integer):boolean;
     function UpdateTableTask(NameTable,CurrentPC,CurrentTask,StatusTask:string;Operation:integer;StartStop,StartThread:boolean):boolean;
     function StartStopTask(nametable:string):boolean; /// ����� ������ ������, ���������� ��� ���������� ����������
     function SelectedPing(HostName:string;Timeout,NumTask:integer):boolean;
     function waitPing(namePC,User,passwd:string;NumTask,WaitTime:integer):boolean; // ������� �������� ��������� ���������� � ����
     function MinPreOper(namePC:string;PreOper,min:integer):integer; //���������� ����� ������� ������ � ������� ���������� ���������� �������� � ����� ��������
     function ConnectWMI (NamePC,User,Passwd:string):Boolean; // �������� ����������� WMI
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
  RunNewTask,DesTask:string; // ���� ��� �������  � �������� �������
  FDQueryWrite,FDQueryWriteTaskTable: TFDQuery;
  TransactionWriteTT  : TFDTransaction;
  ConnectionThreadTask: TFDConnection;
  const
  wbemFlagForwardOnly = $00000020;

implementation
uses umain,TaskEdit,RunTaskRegEdit,RunTaskOsher,PingForTask;
//////////////////////////////////////////////
function bToSTR(b:boolean):string;  //true - 1 ����� false - 0
begin
  if b then result:='1'
  else result:='0';
end;

function STRtoB(b:string):boolean;  //true - 1 ����� false - 0
begin
  if b='0' then result:=false
  else result:=true;
end;
///////////////////////////////////////////////////////////
function TaskRun.Log_write(Level:integer;fname, text:string):string;
var f:TStringList;        /// ������� ������ � ��� ����
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


 ////////////////////////////////////////////////////  ������� ������ ���������� � �������
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
     Log_write(2,'TASK',' ������ ������ ����������� �� ���� '+inttostr(step)+' :' +e.Message);
     result:=false;
     end;
   end;
end;
////////////////////////////////////////////////////////////////////////// ������� ���������� ������� � ������� ����������� ������ � ���� �������
 function TaskRun.UpdateTableTask(NameTable,CurrentPC,CurrentTask,StatusTask:string;Operation:integer;StartStop,StartThread:boolean):boolean;
var
 //Operation : 0- ���������� ������� ������� ��������� � ��� ����������� ������
 //            1 - STATUS_TASK - ����������� ��� ����������� �� ���������  + ������ ��������� ������ �� ���������� START_STOP
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
   FDQueryWriteTaskTable.ParamByName('c').asstring:=booltostr(StartThread); ///������� ���� ��� ����� �����������  ��� ����������
   end;
1:begin
  step:=4;
  FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET STATUS_TASK=:a, START_STOP=:b, WORKS_THREAD=:c  WHERE NAME_TABLE='''+NameTable+'''';
  FDQueryWriteTaskTable.ParamByName('a').asstring:=StatusTask;
  FDQueryWriteTaskTable.ParamByName('b').asstring:=booltostr(StartStop);
  FDQueryWriteTaskTable.ParamByName('c').asstring:=booltostr(StartThread); // ������� ���� ��� ����� ����������� ��� ����������
  end;
2:begin
  step:=5;
  FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET STATUS_TASK=:a, START_STOP=:b, PC_RUN=:c, CURRENT_TASK=:d, WORKS_THREAD=:e   WHERE NAME_TABLE='''+NameTable+'''';
  FDQueryWriteTaskTable.ParamByName('a').asstring:=StatusTask;
  FDQueryWriteTaskTable.ParamByName('b').asstring:=booltostr(StartStop);
  FDQueryWriteTaskTable.ParamByName('c').asstring:='';
  FDQueryWriteTaskTable.ParamByName('d').asstring:='';
  FDQueryWriteTaskTable.ParamByName('e').asstring:=booltostr(StartThread); // ������� ���� ��� ����� ����������� ��� ����������
  end;
3:begin  // ����� ������ � �������
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
     Log_write(2,'TASK',' ������ UPDATE TABLE_TASK �� ���� '+inttostr(step)+' :' +e.Message);
     result:=false;
     end;
   end;
end;
///////////////////////////////////////////////////////////////////////////
function TaskRun.StartStopTask(nametable:string):boolean; /// ����� ������ ������, ���������� ��� ���������� ����������
var
FDQueryRead:TFDQuery;
begin
try
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionWriteTT;
FDQueryRead.Connection:=ConnectionThReadTask;
TransactionWriteTT.StartTransaction; // ��������
FDQueryRead.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE NAME_TABLE='''+nametable+'''';
FDQueryRead.open;
result:=strtobool(FDQueryRead.FieldByName('START_STOP').AsString);
TransactionWriteTT.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryRead.SQL.clear;   //��������
FDQueryRead.Close;  /// ������� ���
FDQueryRead.Free;
 except on E: Exception do
     begin
     result:=true;
     TransactionWriteTT.Rollback;
     if Assigned(FDQueryRead) then FDQueryRead.Free;
     Log_write(1,'TASK',' ������ READ StartStopTask TABLE_TASK :' +e.Message);
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
  if not avalible then finditem('NO','���� �� ��������',HostName,NumTask);;
resIP:='';
Result:=avalible;
end;

////////////////////////////////////////////////////
 function TaskRun.FindAddcreateDir(path,NamePC:string;NumTask:integer):boolean;// �������� � �������� ����������
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //���� ��� ��������
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// ������� ���� ���� �� �����
     begin
     finditem('','�������� ���������� ' +ExtractFileDir(path)+' : ��������� ������ ',NamePC,NumTask);
     result:=false;
     end
     else result:=true; // ���������� �������
    end
    else result:=true; // ���������� ����
  except on E: Exception do
     begin
     finditem('','�������� ���������� ' +ExtractFileDir(path)+' :' +e.Message,NamePC,NumTask);
     result:=false;
     end;
   end;
end;


////////////////////////////////////////////////////////////////
function TaskRun.CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // �����������
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// ����� ������� ������� � vista , �������� ����� � �������� ��� ������ �����������
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //����� ������� ������������, ���� ���� �������� ���������� ������� ������������� ����, ������� �������� ������ � ����� ����� ��������������� ������
      lpszProgressTitle :=nil;            // ��������� �� ��������� ����������� ���� ���������
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // �� ���� �������� ���� �������� �����������
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // ���� ��������
      finditem('','�������� �������� �����������',CurentPC,NumTask);
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // ��� �������  ���� �������� ��������
      pTo := pchar('');   // ���� �������� �� ������������
      //finditem('','�������� �������� �������� ������������',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// ����������� ������������ �� ��������� �����
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // ������� ������� �� ���� � ����
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do finditem('',' : ������ LogonUser - '+e.Message,CurentPC,NumTask)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC,NumTask); //��������� � ������� ������� ���� ��� ���.
     rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/����������� /��������
     if rescopy=0 then
       begin
        if TypeOperation=2 then finditem('','�������� ����������� ������� ���������',CurentPC,NumTask);
        //if TypeOperation=3 then  finditem('','�������� �������� ������������ ������� ���������',CurentPC,NumTask); //finditem('�������� �������� ������� ���������',CurentPC,1);
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then finditem('','������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);
        //if TypeOperation=3 then finditem('','������� �������� ������������ ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);//finditem('������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        result:=false;
       end;
     CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
       begin
       finditem('','������ ������� �����������: '+E.Message,CurentPC,NumTask);
       end;
      if TypeOperation=3 then
       begin
       finditem('','������ ������� ��������: '+E.Message,CurentPC,NumTask);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then finditem('','����� ������ ������� ����������� ����� ���������� ���������  : '+E.Message,CurentPC,NumTask);
     if TypeOperation=3 then finditem('','����� ������ ������� �������� ������������ ����� ��������� ���������  : '+E.Message,CurentPC,NumTask);
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////

function TaskRun.ShotDownResetCloseSession(namePC,Myuser,MyPasswd:string;myShutdown,NumTask:integer):boolean; // ������������/���������� ������/���������� ������

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
                    Operation:='���������� ������';
                    if ResShutdown<>0 then
                      begin
                      ResShutdown:=FWbemObject.Win32Shutdown(4);
                      Operation:='�������������� ���������� ������';
                      end;
                  end;
                  1:
                  begin
                   ResShutdown:=FWbemObject.Win32Shutdown(4);
                   Operation:='�������������� ���������� ������';
                  end;
                   2:
                   begin
                     ResShutdown:=FWbemObject.Win32Shutdown(2);
                     Operation:='������������';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(6);
                        Operation:='�������������� ������������';
                        end;
                   end;
                   3:
                   begin
                   ResShutdown:=FWbemObject.Win32Shutdown(6);
                   Operation:='�������������� ������������';
                   end;
                   4:
                   begin
                     ResShutdown:=FWbemObject.Win32Shutdown(1);
                     Operation:='���������� ������';
                     if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(5);
                        Operation:='�������������� ���������� ������';
                        end;
                    end;
                    5:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(5);
                      Operation:='�������������� ���������� ������';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(12);
                        Operation:='��������� ���������� ������';
                        end;
                    end;
                    6:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(12);
                      Operation:='��������� ���������� ������';
                    end;
                 end;
////////////////////////////////////////////////////////////////////////////////////////////////
          if ResShutdown=0 then
            begin
             result:=true;
             finditem('OK',Operation+' :�������� ������� ���������',namePC,NumTask);
            end
          else
            begin
            result:=false;
            finditem('WARNING',Operation+': ������ - '+SysErrorMessage(ResShutdown),namePC,NumTask);
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
             finditem('ERROR','����� ������ ���������� �������� : '+E.Message,namePC,NumTask);
             result:=false;
             OleUnInitialize;
             end;
      end;
end;
end;
/////////////////////////////////////////////////////////////////
function TaskRun.DeleteProgramMSI(NamePC,UserName,Passwd,NameProgram:string;
equallyName:boolean;NumTASK:integer):boolean;    ///// �������� ��������� msi
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
finditem('RUN','����� ��������� ��� �������� :'+NameProgram,NamePC,NumTASK) ;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', UserName, Passwd,'','',128);
if not equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+NameProgram+'%"','WQL',wbemFlagForwardOnly); // ��� ��������� �� ����������
if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+NameProgram+'"','WQL',wbemFlagForwardOnly);          // ������ ���������� ����� ���������
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
  begin
  finditem('RUN','�������� ������� �������� ��������� :'+NameProgram,NamePC,NumTASK) ;
  ResDelProgram:=FWbemObject.Uninstall(); //// ��������� ���������
  FWbemObject:=Unassigned;
  YesProg:=true;
  end;
if YesProg then
begin
  if (ResDelProgram=0)  then // 0- �������� ���������, 3010 - ����� ��������� ������� ������������???
  begin
  finditem('OK','�������� ��������� '+NameProgram+' :'+SysErrorMessage(ResDelProgram),NamePC,NumTASK);
  result:=true;
  end
  else
  begin
  finditem('WARNING','��� �������� ��������� '+NameProgram+' �������� �������� : ('+inttostr(ResDelProgram)+') '+SysErrorMessage(ResDelProgram),NamePC,NumTASK);
  result:=false;
  end;
end;
if not YesProg then
  begin
  finditem('OK','��������� '+NameProgram+' �� �������',NamePC,NumTASK);
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
    finditem('ERROR','��� �������� ��������� '+NameProgram+' ��������� ������. - "'+E.Message+'"',NamePC,NumTASK);
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
var ///////////////////////////////////// ���������� �������� �� ��������� ������� � ����� ������
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
finditem('RUN','���������� �������� '+NameProcess,NamePC,NumTASK);
ErrorProcKill:=FWbemObject.terminate();
FWbemObject:=Unassigned;
YesProc:=true;
end;
if (YesProc=true) and (ErrorProcKill=0) then
begin
finditem('OK','���������� �������� '+NameProcess+' ������ �������',NamePC,NumTASK);
end;
if (YesProc=true)and (ErrorProcKill<>0) then
begin
finditem('WARNING','��� ���������� �������� '+NameProcess+' �������� ��������: '+SysErrorMessage(ErrorProcKill),NamePC,NumTASK);
end;
if (YesProc=false)then
finditem('OK','������� '+NameProcess+' �� ������',NamePC,NumTASK);
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
  finditem('ERROR','����� ������ ���������� �������� '+NameProcess+' : "'+E.Message+'"',NamePC,NumTASK);
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
Sleep(1500); // ��������� �� 1,5 �������, ����� ���������� ��

FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionWriteTT;
FDQueryRead.Connection:=ConnectionThReadTask;
TransactionWriteTT.StartTransaction;
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT ANSWER_MAC,PC_NAME FROM MAIN_PC WHERE PC_NAME='''+NamePC+'''';
FDQueryRead.Open;
if FDQueryRead.FieldByName('ANSWER_MAC').Value<>null then // ����� MAC ���
aMacAddress:=FDQueryRead.FieldByName('ANSWER_MAC').asString;
TransactionWriteTT.Commit;
FDQueryRead.Close;
FDQueryRead.Free;

if aMacAddress='' then
begin
finditem('WARNING','MAC ����� ���������� '+NamePC+' �� ������, �������� ������������ ����',NamePC,NumTASK);
result:=false;
exit;
end;
//finditem('RUN','��������� ����� WOL �� MAC ����� '+aMacAddress,NamePC,NumTASK);
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
      lUDPClient.Host := BRaddress;   ///// ������ � ��������� IP ������ (���������� ���������� ������� � ����� ����� �����, ��������)
      lUDPClient.Port := 9;
      lUDPClient.SendBuffer(lUDPClient.Host,lUDPClient.Port ,tidbytes(lBuffer));
      finditem('OK','WOL ����� ��������� �� MAC ����� '+RealMAC,NamePC,NumTASK);
      result:=true;
    finally
      lUDPClient.Free;
    end;
  except
   on E: Exception do
   begin
    finditem('ERROR','������ �������� WOL �� MAC ����� '+RealMAC+' :'+E.Message,NamePC,NumTASK);
    result:=false;
   end;
  end;
end;



////////////////////////////////////////////////////
function TaskRun.RunInstallMSI( FSource :String;   /// �������� ����
    FDest:string;      //���� ����������
    NamePC:string;     // ������ �����������
    UserName:string;   // �����
    PassWd:string;     // ������
    PathCreate:boolean;  // �������� � �������� ����������
    CancelCopyFF:boolean;   // �������� ��� ��� �������� �����������
    BeforeInstallCopy:boolean;     // ���������� ��� ��� ����� ����������
    DeleteAfterInstall:boolean; // ������� ����������� ����� ���������
    PathDelete:string;         // ����� ������� ��� ���� ������� ����� ���������
    OwnerForm:Tform;        // ������������ ����� �������� ����, ����� �� ���������
    NumCount:integer;
    TypeOperation:integer;     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
    ProgramInstall:String;   // ��������� ��� ���������
    KeyNewInstallProgram:string; // ����� ��������� ���������
    InstallAllUsers:boolean; // ������������� ��� ���� �������������
    NumberTask:integer):boolean;  // ����� ������� � ������
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
  if BeforeInstallCopy then // ���� ���������� ����������� ����� ����������
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // ��� �����
  UserName,//�����
  PassWd,  // ������
  FSource, // ��� ����������
  FDest,   // ���� ����������
  OwnerForm,// ������������ �����
  TypeOperation, //��� ��������
  CancelCopyFF,
  PathCreate,
  NumberTask) ; // ���������� ����� ������������ �������
  end;
  try
      finditem('RUN','�������� ������� ��������� ��������� '+ExtractFileName(ProgramInstall),ListPC[i],NumberTask);
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(ListPC[i], 'root\CIMV2', UserName, PassWd,'','',128);
      FWbemObject   := FWMIService.Get('Win32_Product');
      Resinstall:=FWbemObject.install(ProgramInstall,KeyNewInstallProgram,InstallAllUsers); //���� ���������, ����� ���������, ��� ���� �������������

       if (Resinstall=0) or (Resinstall=3010) or (Resinstall=1638) then  // 0- �������� ���������, 3010 - ����� ��������� ������� ������������, 1638 - ��� ����������� ������ ������ ����� ��������. ����������� ��������� ����������.
        Begin
        finditem('OK',leftstr('��������� ��������� ' +ExtractFileName(ProgramInstall)
        +' : '+SysErrorMessage(Resinstall),499),ListPC[i],NumberTask);

        End
       else
        Begin
        finditem('WARNING',leftstr('��� ��������� ��������� '
        +ExtractFileName(ProgramInstall) +' �������� ������ : '+' ('+inttostr(Resinstall)+') '
        +SysErrorMessage(Resinstall),499),ListPC[i],NumberTask);

        End;
          {frmDomainInfo.memo1.Lines.Add('��������� ��������� '
          +ExtractFileName(PointForInstallMSI.ProgramInstall)+' �� '
          +ListPC[i]+' : '+SysErrorMessage(Resinstall));}
        FWbemObject:=Unassigned;
        VariantClear(FWbemObject);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
        if BeforeInstallCopy and  //���� ������� ����������� ����� ����������
         DeleteAfterInstall and    // ���� ���������� ������� ����������� ����� ���������
         CopyOrNo then // ���� �������� ����������� ����� ���������� ������ �������
          begin        // �������
          CopyFFSelectPC(ListPC[i], // ��� �����
          UserName,//�����
          PassWd,  // ������
          '', // ��� ���������� ����� �� ��������� �.�. �������
          PathDelete,   // ��� �������
          OwnerForm,// ������������ �����
          3, //��� ��������  (3)- �������, FO_MOVE
          CancelCopyFF,
          false,  // �� ��������� ������� �������� �.�. ������� ����� �����������
          NumberTask) ; // ���������� ����� ������������ �������
          end;
      OleUnInitialize;
        except
          on E:Exception do
           Begin
           finditem('ERROR','������ ��������� ��������� '
           +ExtractFileName(ProgramInstall) +' : '+E.Message,ListPC[i],NumberTask);
           VariantClear(FWbemObject);
           VariantClear(FWMIService);
           VariantClear(FSWbemLocator);
           OleUnInitialize;
           exit;
           End;
       end; // except
  end; //ping
  End; // ����
ListPC.Free;
END;
///////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////
//������ ��������
function TaskRun.MyNewProcess(FSource :String;   ///��� ���������� �������� ���� ��� �������
    FDest:string; // ���� ����������
    NamePC:string;     // ������ �����������
    UserName:string;   // �����
    PassWd:string;     // ������
    PathCreate:boolean;  // �������� � �������� ����������
        CancelCopyFF:boolean;   // �������� ��� ��� �������� �����������
    BeforeInstallCopy:boolean;     // ���������� ��� ��� ����� ����������
    DeleteAfterInstall:boolean; // ������� ����������� ����� ���������
    PathDelete:string;         // ����� ������� ��� ���� ������� ����� ���������
    OwnerForm:Tform;        // ������������ ����� �������� ����, ����� �� ���������
    NumCount:integer;
    TypeOperation:integer;     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
    FileToRun:String;   // ���� ��� �������
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
  if BeforeInstallCopy then // ���� ���������� ����������� ����� ����������
  begin
  CopyOrNo:=CopyFFSelectPC(NamePC, // ��� �����
  UserName,//�����
  PassWd,  // ������
  FSource, // ��� ����������
  FDest,   // ���� ����������
  OwnerForm,// ������������ �����
  TypeOperation, //��� ��������
  CancelCopyFF,
  PathCreate,
  NumberTask); // ���������� ����� ������������ �������                                                                                  end;
  end;
  OleInitialize(nil);
  finditem('RUN','�������� ������� :'+FileToRun,NamePC,NumberTask);
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
    finditem('OK','������ �������� '+FileToRun+' :  '+SysErrorMessage(MyError),NamePC,NumberTask);
    end
  else
    begin
    finditem('WARNING','��� ������� �������� '+FileToRun+' �������� ������ : ('+inttostr(MyError)+') '+SysErrorMessage(MyError),NamePC,NumberTask);
    end;
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  VariantClear(FWbemObject);
  VariantClear(objConfig);
  VariantClear(objProcess);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  if BeforeInstallCopy and  //���� ������� ����������� ����� ����������
  DeleteAfterInstall and    // ���� ���������� ������� ����������� ����� ���������
  CopyOrNo then // ���� �������� ����������� ����� ���������� ������ �������
    begin        // �������
    CopyFFSelectPC(NamePC, // ��� �����
    UserName,//�����
    PassWd,  // ������
    '', // ��� ���������� ����� �� ��������� �.�. �������
    PathDelete,   // ��� �������
    OwnerForm,// ������������ �����
    3, //��� ��������  (3)- �������, FO_MOVE
    CancelCopyFF,
    false,  // �� ��������� ������� �������� �.�. ������� ����� �����������
    NumberTask) ; // ���������� ����� ������������ �������
    end;

    except
      on E:Exception do
      begin
      finditem('ERROR','��� ������� �������� '+FileToRun+' �������� �������� : "'+E.Message+'"',NamePC,NumberTask);
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
 ErrorCodeService: array [0..24] of string=('�������� ������� ���������',
 '�������� �� ��������������','� ������������ ��� ������������ �������',
 '������ ���������� ����������, ��������� �� ��� ������� ������ ���������� ������',
 '����������� ����������� ��� �������������� ��� ���������� ��� ������.'
 ,'����������� ��� �������� ���������� �� ����� ���� ��������� ������ ��-�� ��������� ������ '
 ,'������ �� ��������','������ �� �������� �� ������ ������������',
 '����������� ���� ��� ������� ������','���� � ������������ ����� ������ �� ������'
 ,'������ ��� ��������.','���� ������ ��� ���������� ����� ������ �������������',
 '�����������, �� ������� ��������� ��� ������, ���� ������� �� �������.',
 '������ �� ������� ����� ������, ����������� ��� ��������� ������.',
 ' ������ ���� ��������� �� �������.','������ �� ����� ���������� �������� ����������� ��� ������� � �������.'
 ,'��� ������ ��������� �� �������','������ �� ����� ������ ����������',' ��� ������� ������ ����� ����������� �����������'
 ,'������ ����������� ��� ��� �� ������','��� ������ �������� ������������ �������'
 ,'������ �������� ������������ ���������.','������� ������, ��� ������� ����������� ��� ������, �������� ������������ ��� �� ����� ���������� �� ������ ������.'
 ,'������ ���������� � ���� ������ ��������, ��������� �� �������.',
 '� ��������� ����� ������ �������������� � �������');
begin
try
OleInitialize(nil);
case oper of
1:str:='������ ������';
2:str:='��������� ������';
3:str:='�������� ������';
5:str:='����� ���� ������� ������';
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
1:ResServic:=FWbemObject.StartService();  ////// ��������� ������
2:ResServic:=FWbemObject.StopService();  ////// ���������� ������
3: ResServic:=FWbemObject.Delete(); //// ��������� ������
//4://ResServic:=FWbemObject3.create;//('"DbService"', '"Personnel Database"', '"C:\Program Files (x86)\2gis\3.0\2GISUpdateService.exe"',OWN_PROCESS ,1 ,'"Automatic"' ,NOT_INTERACTIVE,'".\LocalSystem"','');  ////// ������� ������
5:ResServic:=FWbemObject.ChangeStartMode(TypeRunService); //// ����� ���� �������
end;
FWbemObject:=Unassigned;
End;

if findService then   // ���� ������ �������
begin
  if ResServic=0 then finditem('OK',str+' :'+NameService+' :  �������� ������� ���������',NamePC,NumberTask)
  else
  begin
  if ResServic>24 then finditem('WARNING',str+' :'+NameService+' :  '+SysErrorMessage(ResServic),NamePC,NumberTask)
  else finditem('WARNING',str+' :'+NameService+' :  '+ErrorCodeService[ResServic],NamePC,NumberTask);
  end;
end
else  // ����� ����� ������ �� ������� �� �����
finditem('OK','C����� :'+NameService+' �� ������� :  �������� ������� ���������',NamePC,NumberTask);

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
finditem('ERROR',str+' :'+NameService+' ��������� ������ : '+E.Message,NamePC,NumberTask);
result:=false;
end;
end;

end;
/////////////////////////////////////////////////////////////////////////////////////////////++
//////////////////////////////////��������� windows � office //////////////////////////////////////
function TaskRun.DescriptionLicStatus(num:string):string;
var   /// �������� ��� ������ � �� ������� �������� �������� ���� ������
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
begin // ����� ���������� ���� ��� � �������
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
////////////////////////////////////////////////////////////////////Product - Office �  Operating System
function TaskRun.ActivateProduct (Product,namePC,User,Passwd:string;
NumberTask:integer):boolean; // ��������� Windows ��� ���� ������ , ��������� Office ����� Windows 7
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
 finditem('RUN','��������� ��������',NamePC,NumberTask);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 step:=1;
 FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
 step:=2;
 FWMIService.Security_.impersonationlevel:=3;
 step:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy                            WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly
 step:=4;
 {LicenseStatus	DWORD	4	��������� ��������������
0 - ���������������
1 - ������������� (������������)
2 � �������� ������ OOB
3 � �������� ������ OOT
4 - NonGenuineGrace
}                                  //'SELECT  ProductKeyID,ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus,LicenseStatusReason FROM SoftwareLicensingProduct WHERE Description LIKE "%'+Product+'%"'+' and PartialProductKey <> null and LicenseStatus<>1 and LicenseStatus<>0'                                                                                                               //
 ObjectSet:= FWMIService.ExecQuery('SELECT  ProductKeyID,ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus,LicenseStatusReason FROM SoftwareLicensingProduct WHERE Description LIKE "%'+Product+'%"'+' and PartialProductKey <> null and LicenseStatus<>1 and LicenseStatus<>0');   ///����� �������� � ������� � �������� �� ������ �������� �������  � ��� ��������(�����)
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
    if ActProduct then // ���� ���� ����������, �� ����� � ������������ �������
      begin
      finditem('OK','�������� ��������� �������� '+NameProdukt,NamePC,NumberTask);
      step:=9;
      end
      else finditem('OK','������� �����������',NamePC,NumberTask);

    VariantClear(FWMIService);
    VariantClear(ObjectSet);
    VariantClear(FWbemObject);
    oEnum:=nil;
    result:=true;
 except
    on E:Exception do
      begin
      finditem('ERROR','('+inttostr(step)+') ������ ��������� �������� ('+NameProdukt+'): '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
      VariantClear(FWMIService);
      VariantClear(ObjectSet);
      VariantClear(FWbemObject);
      oEnum:=nil;
      result:=false;
      end;
    end;
OleUnInitialize;
end;
                                  // Product - Office �  Operating System
Function TaskRun.InstallKeyWinOffice(StringKey,Product,namePC,User,Passwd:string;
NumberTask:integer):boolean; // ��������� ����� Windows ��� ���� ������ , Office ��� ����� Windows 7
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
  finditem('RUN','��������� ����� ���������',NamePC,NumberTask);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy   WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly
  ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM SoftwareLicensingService');
  oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
      FWbemObject.InstallProductKey(StringKey); //// ��������� ����� windows
      FWbemObject:=Unassigned;
      InstKey:=true;
     end;
   VariantClear(ObjectSet);
   VariantClear(FWMIService);
   VariantClear(FWbemObject);
   oEnum:=nil;   
  if instkey then // ���� ���� ����������, �� ����� � ������������ �������
  begin
  finditem('OK','��������� ����� ���������',NamePC,NumberTask);
  RefreshLicinseWin(namePC,User,Passwd,NumberTask); // ���������� ������� ��������
  ActivateProduct(Product,namePC,User,Passwd,NumberTask);  // ��������� ��������
  end
  else finditem('WARNING','���� �������� �� ����������.',NamePC,NumberTask);
  result:=true;
 except
   on E:Exception do
   begin
    finditem('ERROR','������ ��������� ����� ���������: '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
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
function TaskRun.RefreshLicinseWin(namePC,user,passwd:string;NumberTask:integer):boolean;    // ���������� ������� �������� Windows
var
  FWMIService         : OLEVariant;
  ObjectSet,FWbemObject      : OLEVariant;
  oEnum                : IEnumvariant;
  FSWbemLocator : OLEVariant;
  iValue        : LongWord;
Begin
  OleInitialize(nil);
  finditem('RUN','���������� ������ ��������',NamePC,NumberTask);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User, Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
   try
  ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM SoftwareLicensingService');
   oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
      FWbemObject.RefreshLicenseStatus(); /// ���������� ������� �������������� ��������
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
function TaskRun.refreshOfficeLic(NamePC,User,Passwd:string):boolean;     /// ���������� ������� �������� ����
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
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
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

//////////////////////////////////////////////////////////////////////////////��������� ��� office � windows 7
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
 finditem('RUN','��������� ��������',NamePC,NumberTask);
 step:=1;
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Passwd,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
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
    finditem('OK','�������� ��������� ��������',NamePC,NumberTask);
    result:=true;
 except
    on E:Exception do
      begin
      result:=false;
      finditem('ERROR',inttostr(step)+') ������ ��������� �������� : '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
      VariantClear(FWMIService);
      VariantClear(ObjectSet);
      VariantClear(FWbemObject);
      VariantClear(FSWbemLocator);
      oEnum:=nil;
      end;
    end;
end;

function TaskRun.InstallKeyOfficeWin7(StringKey,namePC,User,Passwd:string;NumberTask:integer):boolean; //��������� ����� ��� office � windows 7
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
 finditem('RUN','��������� ����� ���������',NamePC,NumberTask);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', User,Passwd,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
 ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM OfficeSoftwareProtectionService','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do
  begin
  FWbemObject.InstallProductKey(StringKey); //// ��������� ����� office ��� windows 7
  FWbemObject:=Unassigned;
  InstKey:=true;
  end;
  VariantClear(ObjectSet);
  VariantClear(FWMIService);
  VariantClear(FWbemObject);
  VariantClear(FSWbemLocator);
  oEnum:=nil;
  if instkey then // ���� ���� ����������, �� ����� � ������������ �������
  begin 
  finditem('OK','��������� ����� ���������',NamePC,NumberTask);
  refreshOfficeLic(namePC,User,Passwd); // ���������� ������� ��������
  ActivateOfficeWin7(namePC,User,Passwd,NumberTask); // ��������� ��������
  end
  else finditem('WARNING','���� �������� �� ����������.',NamePC,NumberTask);
  result:=true;
 except
   on E:Exception do
   begin
   result:=false;
    finditem('ERROR','������ ��������� ����� ���������: '+DescriptionLicStatus(E.Message),NamePC,NumberTask);
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
TypeOperation:integer;CancelCopyFF,PathCreate:boolean;NumTask:integer):boolean; // ����������� ��� �������
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// ����� ������� ������� � vista , �������� ����� � �������� ��� ������ �����������
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //����� ������� ������������, ���� ���� �������� ���������� ������� ������������� ����, ������� �������� ������ � ����� ����� ��������������� ������
      lpszProgressTitle :=nil;            // ��������� �� ��������� ����������� ���� ���������
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // �� ���� �������� ���� �������� �����������
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // ���� ��������
      finditem('RUN','�������� �������� �����������',CurentPC,NumTask);
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // ��� �������  ���� �������� ��������
      pTo := pchar('');   // ���� �������� �� ������������
      finditem('RUN','�������� �������� ��������',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// ����������� ������������ �� ��������� �����
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // ������� ������� �� ���� � ����
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do finditem('',' : ������ LogonUser - '+e.Message,CurentPC,NumTask)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC,NumTask); //��������� � ������� ������� ���� ��� ���.
     rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/����������� /��������
     if rescopy=0 then
       begin
        if TypeOperation=2 then finditem('OK','�������� ����������� ������� ���������',CurentPC,NumTask);
        if TypeOperation=3 then  finditem('OK','�������� �������� ������� ���������',CurentPC,NumTask); //finditem('�������� �������� ������� ���������',CurentPC,1);
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then finditem('WARNING','������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);
        if TypeOperation=3 then finditem('WARNING','������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,NumTask);//finditem('������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        result:=false;
       end;
     CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
       begin
       finditem('ERROR','������ ������� �����������: '+E.Message,CurentPC,NumTask);
       end;
      if TypeOperation=3 then
       begin
       finditem('ERROR','������ ������� ��������: '+E.Message,CurentPC,NumTask);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then finditem('ERROR','����� ������ ������� �����������: '+E.Message,CurentPC,NumTask);
     if TypeOperation=3 then finditem('ERROR','����� ������ ������� ��������: '+E.Message,CurentPC,NumTask);
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
 delete(s,2,1); // ������� ������ $
 insert(':',s,2); // ��������� �� ��� ����� :
 end;
result:=s;
end;

function GetLastDir(path:String):String;  // ���������� �������� �����
var
i,j:integer;
begin
path:=ExtractFilePath(path); //���� �� ����� ������ �� path  �����
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

function TaskRun.ReadAndRunMsiOrProc (numberProc,NumberTask:integer;namePC,TYPETASK,UserName,PassWd:string):boolean; // ������� ������� ��������� ��������� ��� ������� ��������
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
 if FDQueryRun.FieldByName('FILEORFOLDER').AsString='File' then // ���� �������� ������ ����
   begin
   step:=2;
   FSource:= FDQueryRun.FieldByName('PATCH_PROC').AsString+#0+#0;
   ProgramInstall:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString); // ������ ������� ����� ��� ���������
   PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
   step:=3;
   end;
 if FDQueryRun.FieldByName('FILEORFOLDER').AsString='Folder' then // ���� �������� �������
   begin
   step:=4;
   FSource:= ExtractFilePath(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;
   ProgramInstall:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString);
   PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
   step:=5;
   end;
   step:=6;
 if not FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean then // ���� �� �������� ���� ��� ������� ����� ����������
 ProgramInstall:= FDQueryRun.FieldByName('PATCH_PROC').AsString; //�� ���� ��� ��������� ����� �� ���������
 step:=7;
 FDest:=FDQueryRun.FieldByName('PATHCREATE').AsString+#0+#0; //
 PathCreate:=true;  /// ��������� � ��������� ������� ��� ��� ����������
 CancelCopyFF:=false;
 BeforeInstallCopy:=FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean; // ���������� ����� ����������
 DeleteAfterInstall:=FDQueryRun.FieldByName('DELETEAFTERINSTALL').AsBoolean;  // ������� ����� ���������
 OwnerForm:=EditTask;
 NumCount:=1; //
 TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
 KeyNewInstallProgram:=FDQueryRun.FieldByName('KEY_MSI').AsString;
 InstallAllUsers:=true; // ��������� ��� ���� �������������
 step:=8;
 RunInstallMSI(FSource, // ��������
               FDest,   // ���� ����������
               namePC,  // ��� �����
               UserName,
               PassWd,
               PathCreate, // ��������� ������� � ���������
               CancelCopyFF,
               BeforeInstallCopy, // ���������� ����� ����������
               DeleteAfterInstall, // ������� ����� ���������
               PathDelete,         // ����� ������� ��� ���� ������� ����� ���������
               OwnerForm,       // ������������ ����� �������� ����, ����� �� ���������
               NumCount,
               TypeOperation,     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
               ProgramInstall,   // ��������� ��� ���������
               KeyNewInstallProgram, // ����� ��������� ���������
               InstallAllUsers, // ������������� ��� ���� �������������  )
               NumberTask); // ���������� ����� ������� ������
 End;
  step:=9;
 if TYPETASK='proc' then
 Begin
    NumCount:=1;
    step:=10;
      if FDQueryRun.FieldByName('FILEORFOLDER').AsString='File' then // ���� �������� ������ ����
        begin
        step:=11;
        FSource:= FDQueryRun.FieldByName('FILESOURSE_PROC').AsString+#0+#0;
        FileToRun:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString); // ������ ������� ����� ��� ���������
        PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+ExtractFileName(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
        end;
      if FDQueryRun.FieldByName('FILEORFOLDER').AsString='Folder' then  // ���� �������� ���� �������� �������
        begin
        step:=12;
        FSource:= ExtractFilePath(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+#0+#0;
        FileToRun:=renewPath(FDQueryRun.FieldByName('PATHCREATE').AsString)+GetLastDir(FDQueryRun.FieldByName('FILESOURSE_PROC').AsString)+ExtractFileName(FDQueryRun.FieldByName('PATCH_PROC').AsString);
        PathDelete:=FDQueryRun.FieldByName('PATHCREATE').AsString+GetLastDir(FDQueryRun.FieldByName('PATCH_PROC').AsString)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
        end;
    if not FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean then  FileToRun:= FDQueryRun.FieldByName('PATCH_PROC').AsString; //// ���� �� �������� ���� ��� ������� ����� ���������� �� ���� ��� ��������� ����� �� ���������
    step:=13;
    FDest:=FDQueryRun.FieldByName('PATHCREATE').AsString+#0+#0;
    PathCreate:=true;
    CancelCopyFF:=false;
    BeforeInstallCopy:=FDQueryRun.FieldByName('BEFOREINSTALLCOPY').AsBoolean;;
    DeleteAfterInstall:=FDQueryRun.FieldByName('DELETEAFTERINSTALL').AsBoolean;
    OwnerForm:=EditTask;
    TypeOperation:=2;
    step:=14;
    MyNewProcess(FSource,  ///��� ���������� �������� ���� ��� �������
    FDest, // ���� ����������
    NamePC,     // �����������
    UserName,   // �����
    PassWd,    // ������
    PathCreate, // �������� � �������� ����������
        CancelCopyFF,   // �������� ��� ��� �������� �����������
    BeforeInstallCopy,     // ���������� ��� ��� ����� ����������
    DeleteAfterInstall, // ������� ����������� ����� ���������
    PathDelete,         // ����� ������� ��� ���� ������� ����� ���������
    OwnerForm,       // ������������ ����� �������� ����, ����� �� ���������
    NumCount,
    TypeOperation,     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
    FileToRun,   // ���� ��� �������
    NumberTask); // ���������� ����� ������� ������
   step:=15;
 End;

FDQueryRun.SQL.clear;   //��������
step:=16;
FDQueryRun.Close;  /// ������� ��� ����� ������
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
     Log_write(2,'TASK',' ��� '+inttostr(step)+' - ������ ������� ��������� ���������(msi) ��� ������� ��������- '+e.Message);
     end;
   end;

end;

function TaskRun.createQuery:boolean;
begin
try
ConnectionThReadTask:=TFDConnection.Create(nil);
ConnectionThReadTask.DriverName:='FB';
ConnectionThReadTask.Params.database:=databaseName; // ������������ ���� ������
ConnectionThReadTask.Params.Add('server='+databaseServer);
ConnectionThReadTask.Params.Add('port='+databasePort);
ConnectionThReadTask.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionThReadTask.Params.Add('CharacterSet=UTF8');
ConnectionThReadTask.Params.add('sqlDialect=3');
ConnectionThReadTask.Params.DriverID:=databaseDriverID;
ConnectionThReadTask.Params.UserName:=databaseUserName;
ConnectionThReadTask.Params.Password:=databasePassword;
ConnectionThReadTask.LoginPrompt:= false;  /// ����������� ������� user password
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
 if AnsiPos(':',br)=2 then // ���� � ���� ���������� ������� ��������� �� ����������� ������������ ����� ����, ��� ����������� � ���� ����� ����� �������� $
 begin
 delete(br,2,1); // ������� ������ :
 insert('$',br,2); // ��������� �� ��� ����� $
 end;
 result:=br;
end;

function TaskRun.ConnectWMI (NamePC,User,Passwd:string):Boolean; // �������� ����������� WMI
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
  Log_write(1,'TASK',NamePC+' - RP� ������ �� ��������');
  FWMIService:=Unassigned;
  VariantClear(FWMIService);
  FSWbemLocator:=Unassigned;
  VariantClear(FSWbemLocator);
  result:=false;
 end;
end;
                  // ��� �����, ������������, ������, ����� ������, ����� �������� ���������� � ���� � �������
function TaskRun.waitPing(namePC,User,passwd:string;NumTask,WaitTime:integer):boolean; // ������� �������� ��������� ���������� � ����
var
step,timeout:integer;
begin
try
timeout:=0;
while timeout<WaitTime*60 do   // ������� ���� ������� ������ ��������� ������� (� ��������) ��������
 Begin
 step:=1;
 if SelectedPing(namePC,pingtimeout,NumTask) and (ConnectWMI(namePC,User,Passwd)) then // ���� ��������� � �������� WMI
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
   timeout:=timeout+20; // ���������� 15 ������ �.� �� �������� ����� 5 ��� + ping (5���) +����������� WMI (~10 ���)
   result:=false;
   end;
 step:=4;
 End;
except
  on E: Exception do
  begin
   result:=false;
   Log_write(1,'TASK',' ��� - '+inttostr(step)+' - Error WaitPing- '+e.Message);
  end;
end;
end;

function TaskRun.MinPreOper(namePC:string;PreOper,min:integer):integer; //���������� ����� ������� ������ � ������� ���������� ���������� �������� � ����� ��������
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
Delete(br,length(br),length(br));  // ������� �������� ������ ')'
for I := length(br) downto 1 do   // ���� � ����� ������ ������ '('
begin
  if br[i]='(' then
  begin
   z:=i;
   break;
  end;
end;
delete(br,1,z); // ������� � ������ ������ �� ������� '('
 if trystrtodatetime(br,mydate) then // ���� br ��� ����
 begin
 if mydate<now then // ���� ����� ���������� ����������� ������� ������ ��� ������� (� ��� ��, ����� ��� ����� ����)
 gap:=MinutesBetween(now,mydate); // �� ����������� ������� � �������
 end;
 if gap>=min then result:=0 // ���� ������� � ������� ������ ����������� �������� �� ����� �� �����, ��������� ��������
 else result:=min-gap; // ����� ���� ������� ����� ���������� �������� � ��� ��������� ���������

 except on E: Exception do // ���� ������ �� ���� ����� ������� ������� ���������
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
createQuery; // ������� ���������� ��� ������ � ��
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThReadTask;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=ConnectionThReadTask;

TransactionWrite.StartTransaction; /// ������ ����� � ������ ��� �������� �������
FDQuery.SQL.Text:='SELECT USER_NAME,PASS_USER,SAVE_PASS,DESCRIPTION_TASK FROM TABLE_TASK WHERE NAME_TABLE='''+RunNewTask+'''';
FDQuery.open;
if FDQuery.FieldByName('USER_NAME').Value<>null then UserPC:=(FDQuery.FieldByName('USER_NAME').AsString);
if FDQuery.FieldByName('PASS_USER').Value<>null then Passwd:=(FDQuery.FieldByName('PASS_USER').AsString);
if FDQuery.FieldByName('DESCRIPTION_TASK').Value<>null then DesTask:=(FDQuery.FieldByName('DESCRIPTION_TASK').AsString);
if (FDQuery.FieldByName('SAVE_PASS').Value<>null) then
 begin
   if not (FDQuery.FieldByName('SAVE_PASS').AsBoolean) then // ���� �� ��������� ������ ����� ������� ������, �� ������ ����
   UpdateTableTask(RunNewTask,'','','����������',3,true,true);
 end;
TransactionWrite.Commit;
FDQuery.close;

TransactionWrite.StartTransaction;
FDQuery.SQL.Clear;
FDQuery.SQL.Text:='SELECT * FROM '+RunNewTask;
FDQuery.Open;
FDQuery.Last;/// ������� � ��������� ������
CountPC:=FDQuery.RecordCount;     // ���������� �������  ������������� ���������� ������
countTask:=FDQuery.FieldByName('TASK_QUANT').AsInteger; // ���������� ����������� ������� � ������� ������
FDQuery.First;// ������� � ������ ������
StartStop:=true;
UpdateTableTask(RunNewTask,'','','�����������',1,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������, ������� ���� ��� ����� ��� ������� �������
Log_write(0,'TASK',' ������ ���������� ������ '+DesTask);
For i := 0 to countTask-1 do // ���� �� ��������
BEGIN
if not StartStop then break; // ���� ������ ����������� �� ������� �� �����
 step:=6;
 FDQuery.First;  // ������� � ������ ������, ��� ���������� ����� ������ ������� � 1�� �����
 while not FDQuery.Eof  do /// �������� � ������ ������ ���� �� ��������� ������
 BEGIN
  step:=7;
  if not StartStopTask(RunNewTask) then // ���� ������ ����������� �� ���������� ����������� false � ������� �� �����
   begin
   StartStop:=false;
   break;
   end;
  IF (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='WOL') THEN  // �������� ��������� WOL
    begin
    UpdateTableTask(RunNewTask,               // ��� �������
    FDQuery.FieldByName('PC_NAME').AsString, //��� �����
    '��������� ���������� (WOL)',  // ��� �������
    '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
    WakeOnLan(FDQuery.FieldByName('PC_NAME').AsString, // ��� �����+
    StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString),
    i);  // ���������� ����� ����������� ������
    end
  ELSE
  BEGIN
  // ���� ���������� �������� ��� ������������ ��� ��������� ����� �� ������� ��� ��������� � ����
  if (waitafter)then
  begin
  if i>0 then // ���� ��� �� ������ �������� � �����, �� ��������� ������� ������� ������ � ������� ���������� ��������� ��������
    begin       // � ������� ��������� ���������� ���������� ��������, ���� ����� ���������� ���� ��������
    TimeOutAfter:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),TimeOutAfter); //��������� ������� ������� �����
    end;
  if TimeOutAfter>0 then // ���� ����������� ����� ������ 0 �� ������� ����� ������� �������
    begin
    UpdateTableTask(RunNewTask,               // ��� �������
    FDQuery.FieldByName('PC_NAME').AsString, //��� �����
    '�������� �����������, ��������� �� ��������',  // ��� �������
    '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
     waitPing(FDQuery.FieldByName('PC_NAME').AsString, //��� �����
     UserPC,Passwd, // ����� � ������
     i,//����� �������
     TimeOutAfter); // ������� � ������� �������� ����������
     end;
  end;

    if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='ShDown')   // ���� ����������
     then
      begin
        if (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger>0)and (i>0) then //���� ���������� ����� ����� �����������
        begin
         // � ������� ��������� ���������� ���������� �������� ���� ����� ���������� ���� ��������
         TimeBeforeOff:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger); //��������� ������� ������� �����
         if TimeBeforeOff>0 then // ���� ����������� ����� ������ 0 �� ������� ����� ������� ������� ������� ��������
         begin
         UpdateTableTask(RunNewTask,               // ��� �������
         FDQuery.FieldByName('PC_NAME').AsString, //��� �����
         '������� �������� ����� ����������� ����������',  // �������� �������
         '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
         sleep(TimeBeforeOff*1000*60);// ���� ����� �����������
         end;
        end;
      UpdateTableTask(RunNewTask,               // ��� �������
         FDQuery.FieldByName('PC_NAME').AsString, //��� �����
         '���������� ����������',  // �������� �������
         '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
      if SelectedPing(FDQuery.FieldByName('PC_NAME').AsString,pingtimeout,i) then // ���� ���� �������� �� ���������
       ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //������������ � ������
      4, // ���������� ������
      i);   // ���������� ����� ����������� ������
      end; //��������� ��������� ��

   if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='TimeOut')  then // ������� �������� �� ���������� ���������� �������
   begin
     if (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger>0) then //���� ������� ������ ����
        begin
        finditem('OK','������� '+FDQuery.FieldByName('NumTask'+inttostr(i)).AsString+' ���.',FDQuery.FieldByName('PC_NAME').AsString,i);
         // � ������� ��������� ���������� ���������� �������� ���� ����� ���������� ���� ��������
         TimeOutBefore:=MinPreOper(FDQuery.FieldByName('PC_NAME').AsString,(i-1),FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger); //��������� ������� ������� �����
         if TimeOutBefore>0 then // ���� ����������� ����� ������ 0 �� ������� ����� ������� ������� ������� ��������
         begin
         UpdateTableTask(RunNewTask,               // ��� �������
         FDQuery.FieldByName('PC_NAME').AsString, //��� �����
         '������� �������� ('+inttostr(TimeOutBefore)+' ���.)',  // �������� �������
         '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
         sleep(TimeOutBefore*1000*60);// ������� ����� ��������� ���������
         end;
        end;
   end;

  ////////////////////////////////////////////////////////////////////
  if(SelectedPing(FDQuery.FieldByName('PC_NAME').AsString,pingtimeout,i))                     //���� ��������� ��������
  and (FDQuery.FieldByName('STATUSTASK'+inttostr(i)).AsString<>'OK') then //  ������ ������� �� OK(�� ���������)
  Begin
  UpdateTableTask(RunNewTask,               // ��� �������
  FDQuery.FieldByName('PC_NAME').AsString, //��� �����
  FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString,  // ��� �������
  '',0,true,true); /// ������ �������, ��� �������� ���������� �������, ������ ������  ��� �� ������ �� ���������� �.�. �������� 0. ������� ���� ��� ����� �������
   ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
   if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='msi') or
   (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='proc') then
    begin
    step:=8;
    ReadAndRunMsiOrProc
    (FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger, // ����� �������� � �������  START_PROC_MSI
    i,                                           // ���������� ����� ����������� ������
    FDQuery.FieldByName('PC_NAME').AsString, // ��� �����
    FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString, // �� � ���, ��� ��� ������� ��� msi
    UserPC,Passwd); //������������ � ������
    end;
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='DelMSI' then // �������� ��������� MSI
     begin
     DeleteProgramMSI(FDQuery.FieldByName('PC_NAME').AsString, // ��� �����+
     UserPC,Passwd,  // ������������ � ������
     StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //�������� ��� ��������� ��� �������� �� �������� ������
     STRtoB(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), // ������ ���������� ����� ��������� ��� ���
     i);  // ���������� ����� ����������� ������
     end;
    ///////////////////////////////////////////////////////////////////////////////////////////////////////
    if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='KillProc' then // ���������� ��������
     begin
     KillProcess(FDQuery.FieldByName('PC_NAME').AsString, // ��� �����+
     UserPC,Passwd,  // ������������ � ������
     StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //�������� ��� �������� ��� ����������  �� �������� ������
     STRtoB(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), // ������ ���������� ����� �������� ��� ���
     i);  // ���������� ����� ����������� ������
     end;

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Service' then // ���� ���������� ��������
      begin
      ControlService(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //������������ � ������
      StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString), //�������� ��� ������  �� �������� ������
      FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger,  // ��� ��������, 1 -������, 2- ���������, 3- ��������
      i);   // ���������� ����� ����������� ������

      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Reset' then // ���� ���������� ��������
      begin
      ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //������������ � ������
      2, // ������������
      i);   // ���������� ����� ����������� ������

      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Logout' then // ���� ���������� ��������
      begin
      ShotDownResetCloseSession(FDQuery.FieldByName('PC_NAME').AsString,
      UserPC,Passwd, //������������ � ������
      0, // ���������� ������
      i);   // ���������� ����� ����������� ������
      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Activate' then // ���� ��������� office ��� windows
      begin
        case FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger of  // �� ���� �������� //Product - Office �  Operating System
        1: InstallKeyWinOffice(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // ��������� Windows
        ,'Operating System',FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i);
        2: InstallKeyWinOffice(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // ��������� Office ��� windows 8/8.1/10...
        ,'Office',FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i);
        3: InstallKeyOfficeWin7(StrInDescription(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)  // ��������� Office ��� windows 7
        ,FDQuery.FieldByName('PC_NAME').AsString,UserPC,Passwd,i)
        end;
      end;
      //////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CopyFF' then // ���� ������� �����������
      begin
        CopyDeleteFF(FDQuery.FieldByName('PC_NAME').AsString, //��� �����
        UserPC,Passwd, //����� � ������
        StrCopySourse(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, //��� ��������
        StrCopyDestination(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, // ���� ��������
        EditTask,
        2, // �������� �����������
        false,  //����������� ������
        strtobool(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString), //��������� ��� ��� ������� ��� ��� ����������
        i); // ����� ������
      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='DelFF' then // ���� �������
      begin
      CopyDeleteFF(FDQuery.FieldByName('PC_NAME').AsString, //��� �����
        UserPC,Passwd, //����� � ������
        '', //��� ��������  ������������� ���
        StrDeleteSourse(FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString)+#0+#0, // ��� �������
        EditTask,
        3, // �������� ��������
        false,  //����������� ������
        false, //��������� ��� ��� ������� ��� ��� ����������
        i); // ����� ������
      end;
      ///////////////////////////////////////////////////////////////////////////////////////////////////////
      if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sSubKey') //sSubKey ������ � �������� ������� (�������� ��� ��������)  NumTask ���������� �������� ��� �������� ������� �������
      or (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sNameKey') //sNameKey ������ �� �������� ����� �������
      or (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CreateKey') // �������� ����� �������, �������������� ����� �� ������� ����������� ������
      then
       begin
       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='CreateKey' then
       finditem('RUN','�������� �������� �������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);

       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sNameKey' then
       finditem('RUN','�������� �������� �������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);

       if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='sSubKey' then
       begin
        if FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger=0 then  finditem('RUN','�������� �������� �������� ������� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        if FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger=1 then  finditem('RUN','�������� �������� �������� ������� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
       end;
       res:=RunTaskRegEdit.RunEditRegeditWindows(FDQuery.FieldByName('PC_NAME').AsString, //��� �����
                              UserPC,Passwd, //����� � ������
                              FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString, //��� ������������ �������
                              FDQuery.FieldByName('NumTask'+inttostr(i)).AsString, // � NUMTASK ������� ����� �������������� ����� � ������/ �������� ��� �������� ������� �������
                              FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString // � NAMETASk ���������� ���� � ������������ ��� ���������� ������� �������/ ���������� ����� �������
                              );
        case res of
        0:finditem('OK','�������� ������� ���������',FDQuery.FieldByName('PC_NAME').AsString,i);
        666: finditem('WARNING','������ RPC ����������',FDQuery.FieldByName('PC_NAME').AsString,i);
        631: finditem('WARNING','������ �������� ������� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        632: finditem('WARNING','������ �������� ������� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        633: finditem('WARNING','������ �������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        637: finditem('WARNING','������ �������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        639: finditem('WARNING','������ �������� ����� DWORD �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        640: finditem('WARNING','������ �������� ����� QWORD �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        641: finditem('WARNING','������ �������� ���������������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        642: finditem('WARNING','������ �������� ������������ ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        643: finditem('WARNING','������ ������� ������� ������ �� ������ ����������',FDQuery.FieldByName('PC_NAME').AsString,i);
        644: finditem('WARNING','������ ������� ������ �� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        646: finditem('WARNING','������ �������� ��������� ����� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        888: finditem('WARNING','����� ������ ������� ������� ������ � ��������',FDQuery.FieldByName('PC_NAME').AsString,i);
        else finditem('WARNING',SysErrorMessage(res),FDQuery.FieldByName('PC_NAME').AsString,i);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='RestoreWin' then // ���� ������ � ������� ��������������
      begin
        res:=RunTaskOsher.RestoreOrNewPoint
        (FDQuery.FieldByName('PC_NAME').AsString, //��� �����
        UserPC,Passwd, //����� � ������
        FDQuery.FieldByName('NAMETASK'+inttostr(i)).AsString, // ���� ��� ��������, ��� ���� � �������� ��� �������������� �� ���� � ��������
        strtoint(FDQuery.FieldByName('NumTask'+inttostr(i)).AsString));   // ��� ������ , 1-����� ����� ��������������, 2 - ��������� �������������� �� ����, - 3 � �������������� �� ���� � ��������
        case res of
        0: finditem('OK','�������� ������� ���������',FDQuery.FieldByName('PC_NAME').AsString,i);
        666: finditem('WARNING','������ RPC ����������',FDQuery.FieldByName('PC_NAME').AsString,i);
        777:finditem('WARNING','����� �������������� �� �������',FDQuery.FieldByName('PC_NAME').AsString,i);
        888:finditem('WARNING','����� ������ ��� ������ � ������� ��������������',FDQuery.FieldByName('PC_NAME').AsString,i);
        999:finditem('WARNING','������ �������� ����� ��������������',FDQuery.FieldByName('PC_NAME').AsString,i);
        1111:finditem('WARNING','������ ��������� �������������� �� �����',FDQuery.FieldByName('PC_NAME').AsString,i);
        2222:finditem('WARNING','������ �������� ����� ��������������',FDQuery.FieldByName('PC_NAME').AsString,i);//
        else  finditem('WARNING',SysErrorMessage(res),FDQuery.FieldByName('PC_NAME').AsString,i);
        end;
     end;


  End; //���� ���� �������� � ������� �� ���������
 END; //����� , ����� �������� ����� WakeOnLan

 FDQuery.Next;  // ��������� �� ��������� ������
 END;// ���� �� ������� (������) � �������

 if (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='Reset') or
 (FDQuery.FieldByName('TYPETASK'+inttostr(i)).AsString='WOL')  then
 begin
 waitafter:=true; // ����������, ������� ����� ������������ ��� ���������
 TimeOutAfter:=FDQuery.FieldByName('NumTask'+inttostr(i)).AsInteger;
 end
  else waitafter:=false; //������ �������� ���� �� ������������ � �� ��������� ��
 step:=9;

END; /// ���� �� �������

if StartStop then UpdateTableTask(RunNewTask,'','','���������',2,false,false) /// ���� ������ �� ��������������� ������������� �� ��� �����������
else UpdateTableTask(RunNewTask,'','','�����������',1,false,false);          // ����� �� ���������� �������������

FDQuery.SQL.clear;   //��������
FDQuery.Close;  /// ������� ��� ����� ������
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
Log_write(0,'TASK',' ���������� ������ '+DesTask+ ' ���������');
Except
  on E:Exception do
     begin
     Log_write(3,'TASK',' ��� - '+inttostr(step)+' - Fatal error- '+e.Message);
     if Assigned(TransactionWriteTT) then UpdateTableTask(RunNewTask,'','','������ ���������� ������',1,false,false);
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

