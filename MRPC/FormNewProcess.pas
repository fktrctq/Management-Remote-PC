unit FormNewProcess;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.Dialogs,inifiles,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.StrUtils,System.Variants,ActiveX,ComObj,IdIcmpClient,ShellAPI,
  Vcl.ExtCtrls;
 type
 TStrForNewProcess = record
    FSource :String;   ///��� ���������� �������� ���� ��� �������
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
  end;
type
  TNewProcForm = class(TForm)
    OKBtn: TButton;
    CheckBox1: TCheckBox;
    Label2: TLabel;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    SpeedButton3: TSpeedButton;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    CheckBox4: TCheckBox;
    EditCopyPath: TLabeledEdit;
    ComboBox1: TComboBox;
    CheckBox5: TCheckBox;
    Button1: TButton;
    CheckBox2: TCheckBox;
    EditSource: TLabeledEdit;
    Button2: TButton;
    EditFileRun: TLabeledEdit;
    SpeedButton4: TSpeedButton;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Label1: TLabel;
    SpeedButton5: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    function readproc(item:integer):boolean;
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;



var
  NewProcForm: TNewProcForm;
  ArrayProcess: array of TThread;
  ExitFor:boolean;
  StrForProcess: array [0..2000] of TStrForNewProcess; //������ ���������� ��� ������� ������� ��������� ��������
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;

 ThreadVar
PointForRunNewProcess: ^TStrForNewProcess;

implementation
uses uMain,unit5, SelectedPCNewProcessThread,MyDM,EditProcMSI;



{$R *.dfm}

function TNewProcForm.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''proc'''
+' ORDER BY DESCRIPTION_PROC'; //ID_PROC
FDQueryProc.Open;
ComboBox2.Clear;
ComboBox3.Clear;
while not FDQueryProc.Eof do
begin
 if FDQueryProc.FieldByName('DESCRIPTION_PROC').Value<>null then
 begin
   ComboBox2.Items.Add(FDQueryProc.FieldByName('DESCRIPTION_PROC').AsString);
   ComboBox3.Items.Add(FDQueryProc.FieldByName('ID_PROC').AsString);
 end;
  FDQueryProc.Next;
end;
FDQueryProc.Close;
if (ComboBox2.Items.Count>0) and (item<>-1) then
  begin
  ComboBox2.ItemIndex:=item;
  ComboBox2.OnSelect(ComboBox2);
  end;
end;

function MyNewProcess(param:pointer):boolean;
var
MyError:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObject   : OLEVariant;
objProcess    : OLEVariant;
objConfig     : OLEVariant;
ProcessID,z,i   : Integer;
listPC:TstringList;
CopyOrNo:bool;
const
wbemFlagForwardOnly = $00000020;
HIDDEN_WINDOW       = 1;
function finditem(s,name:string;image:integer):boolean;
var
z:integer;
begin
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
    if frmDomainInfo.ListView8.Items[z].SubItems[0]=name then
    begin
    frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=image;
    frmDomainInfo.ListView8.Items[z].SubItems[1]:=s;
    break;
    end;
end;
////////////////////////////////////////////////////
 function FindAddcreateDir(path,NamePC:string):boolean;// �������� � �������� ����������
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //���� ��� ��������
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// ������� ���� ���� �� �����
     begin
     finditem('�������� ���������� ' +ExtractFileDir(path)+' : ��������� ������ ',NamePC,2);
     result:=false;
     end
     else result:=true; // ���������� �������
    end
    else result:=true; // ���������� ����
  except on E: Exception do
     begin
     frmdomaininfo.Memo1.Lines.Add(NamePC+' : ������ �������� ���������� - '+e.Message);
     finditem('�������� ���������� ' +ExtractFileDir(path)+' :' +e.Message,NamePC,2);
     result:=false;
     end;
   end;
end;
///////////////////////////////////////////////////////////
function ping(s:string):boolean;
var
z:integer;
Myidicmpclient:TIdIcmpClient;
begin
try
result:=false;
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  finditem('�������� �������� �������� �������',s,2);
  end
else
  begin
  result:=true; ///��������
  frmDomaininfo.Memo1.Lines.Add('IP ����� - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  �����='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'��'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;
   except
    begin
    result:=false;
    finditem('���� �� ��������',s,2);
    end;
   end;
if Assigned(MyIdIcmpClient) then freeandnil(MyIdIcmpClient);
end;
////////////////////////////////////////////////////////////////
function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean):boolean; // ����������� � ������ ��� �����  �����������
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
      finditem('�������� �������� �����������',CurentPC,17);
      frmDomaininfo.Memo1.Lines.Add(CurentPC+' - �������� �������� �����������');
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // ��� �������  ���� �������� ��������
      pTo := pchar('');   // ���� �������� �� ������������
      finditem('�������� �������� �������� ',CurentPC,13);
      frmDomaininfo.Memo1.Lines.Add(CurentPC+' - �������� �������� ��������');
     end;
    end;

    try ////////////////////////////////////// ����������� ������������ �� ��������� �����
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // ������� ������� �� ���� � ����
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do frmdomaininfo.Memo1.Lines.Add(CurentPC+' : ������ LogonUser - '+e.Message)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC); //��������� � ������� ������� ���� ��� ���.
     rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/����������� /��������
     if rescopy=0 then
       begin
        if TypeOperation=2 then
        begin
        finditem('�������� ����������� ������� ���������',CurentPC,1);
        frmDomaininfo.Memo1.Lines.Add(CurentPC+' - �������� ����������� ������� ���������');
        end;
        if TypeOperation=3 then
        begin
         finditem('�������� �������� ������� ���������',CurentPC,1);
         frmDomaininfo.Memo1.Lines.Add(CurentPC+' - �������� �������� ������� ���������');
        end;
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then
        begin
         finditem('������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
         frmDomaininfo.Memo1.Lines.Add(CurentPC+' - ������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        if TypeOperation=3 then
        begin
        finditem('������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        frmDomaininfo.Memo1.Lines.Add(CurentPC+'������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        result:=false;
       end;
     CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
      except
      on E: Exception do
      begin
       if TypeOperation=2 then
       begin
       frmdomaininfo.Memo1.Lines.Add(Format('������ ������� ����������� �� ���������  %s :  %s',[CurentPC,E.Message]));
       finditem('������ ������� �����������: '+E.Message,CurentPC,2);
       end;
       if TypeOperation=3 then
       begin
       frmdomaininfo.Memo1.Lines.Add(Format('������ ������� �������� �� ����������  %s :  %s',[CurentPC,E.Message]));
       finditem('������ ������� ��������: '+E.Message,CurentPC,2);
       end;
       result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then frmdomaininfo.Memo1.Lines.Add(Format('����� ������ ������� ����������� :  %s',[E.Message]));
     if TypeOperation=3 then frmdomaininfo.Memo1.Lines.Add(Format('����� ������ ������� �������� :  %s',[E.Message]));
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////

Begin
PointForRunNewProcess:=param;
try
listPC:=TStringList.Create;
listPC.CommaText:=PointForRunNewProcess.NamePC;
for I := 0 to listPC.Count-1 do
  BEGIN
  if ping(listPC[i]) then    ///// ���� ��������� �������� �� �������� �������
  try
  begin
  CopyOrNo:=false;
  if PointForRunNewProcess.BeforeInstallCopy then // ���� ���������� ����������� ����� ����������
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // ��� �����
  PointForRunNewProcess.UserName,//�����
  PointForRunNewProcess.PassWd,  // ������
  PointForRunNewProcess.FSource, // ��� ����������
  PointForRunNewProcess.FDest,   // ���� ����������
  PointForRunNewProcess.OwnerForm,// ������������ �����
  PointForRunNewProcess.TypeOperation, //��� ��������
  PointForRunNewProcess.CancelCopyFF,
  PointForRunNewProcess.PathCreate) ;
  end;
  OleInitialize(nil);
  frmDomainInfo.memo1.Lines.Add('������ �������� �� - '+listPC[i]);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(listPC[i], 'root\CIMV2', PointForRunNewProcess.UserName, PointForRunNewProcess.PassWd,'','',128);
  FWMIService.security_.AuthenticationLevel:=6;
  FWMIService.security_.ImpersonationLevel:=3;
  //FWMIService.security_.Privileges.AddAsString('SeEnableDelegationPrivilege');
  FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
  objConfig     := FWbemObject.SpawnInstance_;
  objConfig.ShowWindow := HIDDEN_WINDOW;
  objProcess    := FWMIService.Get('Win32_Process');
  MyError:=objProcess.Create(PointForRunNewProcess.FileToRun, null, objConfig, (ProcessID));
  if MyError=0 then
    begin
    finditem('������ �������� '+PointForRunNewProcess.FileToRun+' : '+': '+SysErrorMessage(MyError),listPC[i],1);
    end
  else
    begin
    finditem('��� ������� �������� '+PointForRunNewProcess.FileToRun+' �������� �������� : '+SysErrorMessage(MyError),listPC[i],2);
    end;
  frmDomainInfo.memo1.Lines.Add('������ �������� '+PointForRunNewProcess.FileToRun+' �� '+listPC[i]+' : '+SysErrorMessage(MyError));
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  VariantClear(FWbemObject);
  VariantClear(objConfig);
  VariantClear(objProcess);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  if PointForRunNewProcess.BeforeInstallCopy and  //���� ������� ����������� ����� ����������
  PointForRunNewProcess.DeleteAfterInstall and    // ���� ���������� ������� ����������� ����� ���������
  CopyOrNo then // ���� �������� ����������� ����� ���������� ������ �������
    begin        // �������
    CopyFFSelectPC(ListPC[i], // ��� �����
    PointForRunNewProcess.UserName,//�����
    PointForRunNewProcess.PassWd,  // ������
    '', // ��� ���������� ����� �� ��������� �.�. �������
    PointForRunNewProcess.PathDelete,   // ��� �������
    PointForRunNewProcess.OwnerForm,// ������������ �����
    3, //��� ��������  (3)- �������, FO_MOVE
    PointForRunNewProcess.CancelCopyFF,
    false) ;  // �� ��������� ������� �������� �.�. ������� ����� �����������
    end;
  end;
    except
      on E:Exception do
      begin
      finditem('��� ������� �������� '+PointForRunNewProcess.FileToRun+' �������� �������� : "'+E.Message+'"',listPC[i],2);
      frmDomainInfo.memo1.Lines.Add('��� ������� �������� '+PointForRunNewProcess.FileToRun+' �� '+listPC[i]+' �������� ��������. - "'+E.Message+'"');
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      OleUnInitialize;
      end;
    end;
  END;
finally
listPC.Free;
End;
end;

function  DeliteKeyFilePatch(strfile:string):string; // ������� ������������ ������������ � ��� ���� � ������� �������, ������ ���� �� �����
begin
try
if pos(' -',strfile)<>0 then    // ���� ����� ������� ����� �����
begin
  strfile:=copy(strfile,1,pos(' -',strfile)-1);
end;
if pos('/',strfile)<>0 then //���� �������� ����, ������� ����� �������
begin
 strfile:=copy(strfile,1,pos('/',strfile)-1);
 if strfile[length(strfile)]=' ' then strfile:=copy(strfile,1,length(strfile)-1);
end;
result:=strfile;
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('������ ������������ ������ - "'+E.Message+'"');
result:=strfile;
end;
  end;
end;


procedure TNewProcForm.Button1Click(Sender: TObject);
Function renewPath(s:string):string;
begin
if AnsiPos(':',s)=2 then
 begin
 delete(s,2,1); // ������� ������ :
 insert('$',s,2); // ��������� �� ��� ����� $
 end;
result:=s;
end;
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� �������� ��� �����������';
  Options:=[fdoForceShowHidden,fdoPickFolders]; {��������}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditCopyPath.Text:=renewPath(FileName)+'\';
    end;
 finally
 Destroy;
 end;
 end;
 end;





procedure TNewProcForm.Button2Click(Sender: TObject);
begin
Exitfor:=false;
Close;
end;

procedure TNewProcForm.CheckBox2Click(Sender: TObject);
begin
GroupBox2.Enabled:=CheckBox2.Checked;
if CheckBox2.Checked then
begin
EditSource.Text:=DeliteKeyFilePatch(EditFileRun.Text);
GroupBox2.Height:=160;
end
else GroupBox2.Height:=25;
end;

procedure TNewProcForm.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
combobox2.DroppedDown:=true;
end;

procedure TNewProcForm.ComboBox2Select(Sender: TObject);
var
FDQuery:TFDQuery;
begin
ComboBox3.ItemIndex:=ComboBox2.ItemIndex;
try
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionReadProc;
FDQuery.Connection:=DataM.ConnectionDB;
FDQuery.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE DESCRIPTION_PROC='''+ComboBox2.Text+'''';//' ORDER BY ID_PROC';
FDQuery.Open;
EditfileRun.Text:=vartostr(FDQuery.FieldByName('PATCH_PROC').Value);
if FDQuery.FieldByName('BEFOREINSTALLCOPY').Value<>null then CheckBox2.Checked:=FDQuery.FieldByName('BEFOREINSTALLCOPY').AsBoolean;
if FDQuery.FieldByName('FILESOURSE_PROC').Value<>null then  EditSource.Text:=FDQuery.FieldByName('FILESOURSE_PROC').AsString;
if FDQuery.FieldByName('FILEORFOLDER').Value<>null then
begin
if FDQuery.FieldByName('FILEORFOLDER').AsString='File' then combobox1.ItemIndex:=0
else combobox1.ItemIndex:=1;
end;
if FDQuery.FieldByName('PATHCREATE').Value<>null then EditCopyPath.Text:=FDQuery.FieldByName('PATHCREATE').AsString;
if FDQuery.FieldByName('DELETEAFTERINSTALL').Value<>null then  CheckBox5.Checked:=FDQuery.FieldByName('DELETEAFTERINSTALL').AsBoolean;

finally
FDQuery.Close;
FDQuery.Free;
end;
end;


{procedure TPasswordDlg.ComboBox1DropDown(Sender: TObject);
var i,L,mwidth:integer;  /// ������������� ������ ����������� ������
begin
mwidth:=0;
   with ComboBox1 do
   begin
     for i := 0 to Items.Count - 1 do
     if ( Canvas.TextWidth(Items[I]) > mWidth) then
       mWidth :=Canvas.TextWidth(Items[I])+ 100;
   SendMessage(ComboBox1.Handle ,$0160,mWidth,0);
   end
end; }



procedure TNewProcForm.FormCreate(Sender: TObject);
begin
speedButton1.Hint:='�������� ����� ������� � ��������';
speedButton2.Hint:='������� ������� �� ������';
end;

procedure TNewProcForm.FormShow(Sender: TObject);
begin
if not groupPC then NewProcForm.Caption:='����� ������� �� '+frmDomainInfo.ComboBox2.Text
else NewProcForm.Caption:='����� ������� �� ������ ��';
CheckBox1.Checked:=true; /// � ����� ������ ���������� �� ���������� ������� ��� �� ������� ����������
if GroupPC=true then CheckBox1.Visible:=true
else CheckBox1.Visible:=false;
label2.Visible:=false;
GroupBox2.Height:=25;
CheckBox2.Checked:=false;
if not DataM.ConnectionDB.Connected then exit;
readproc(-1);
end;



procedure TNewProcForm.OKBtnClick(Sender: TObject);
var
b,i:integer;
NumPer:integer;
SelectedPCNewProc:Tstringlist;
function tagForCopy(num:integer):integer;
begin
if OKBtn.tag=1999 then num:=0;
inc(num);
OKBtn.tag:=num;
result:=num;
end;

///////////////////////////////////////////////
function GetLastDir(path:String):String;  // ���������� �������� �����
var
i,j:integer;
begin
try
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
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('������ ���������� �������� ����� - "'+E.Message+'"');
result:=path;
end;
  end;
end;
/////////////////////////////////////////////////////
Function renewPath(s:string):string;
begin
try
if AnsiPos('$',s)=2 then
 begin
 delete(s,2,1); // ������� ������ $
 insert(':',s,2); // ��������� �� ��� ����� :
 end;
result:=s;
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('������ renewPath - "'+E.Message+'"');
result:=s;
end;
  end;
end;
///////////////////////////////////////////////////////
begin
try
{if not FileExists(EditFileRun.Text) then
begin
    b:= MessageDlg('�� ���� ����� ����, �������� ��������� ����� �������.'+ #10#13+'���������� ������ ��������?',mtConfirmation,[mbYes,mbCancel], 0);
    if b = mrCancel then exit;
end;}
  if (CheckBox1.Checked) then //// ������ �������� � ����� ������ �� ������ ����������� ��� �� ����� ���������
    begin
    NumPer:=tagForCopy(OKBtn.tag);
    StrForProcess[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
    StrForProcess[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
    if GroupPC=true then // ���� ��������� ��������� �� ��������� ������ �����������
      begin
      SelectedPCNewProc:=TstringList.Create;
      SelectedPCNewProc:=frmDomainInfo.createListpcForCheck('');
      StrForProcess[NumPer].NamePC:=SelectedPCNewProc.CommaText;
      SelectedPCNewProc.Free;
      end
    else StrForProcess[NumPer].NamePC:=frmDomainInfo.ComboBox2.Text; // ����� ��������� �� ������� ���������
    if CheckBox2.Checked then // ���� ���� ����������� �����
    begin
      if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
        begin
        StrForProcess[NumPer].FSource:= EditSource.Text+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+ExtractFileName(EditFileRun.Text); // ������ ������� ����� ��� ���������
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(EditSource.Text)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
        end
      else // ���� �������� ���� �������� �������
        begin
        StrForProcess[NumPer].FSource:= ExtractFilePath(EditSource.Text)+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+GetLastDir(EditSource.Text)+ExtractFileName(EditFileRun.Text);
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(EditFileRun.Text)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
        end;
     end;
    if not CheckBox2.Checked then  StrForProcess[NumPer].FileToRun:= EditFileRun.Text; //// ���� �� �������� ���� ��� ������� ����� ���������� �� ���� ��� ��������� ����� �� ���������
    StrForProcess[NumPer].FDest:=EditCopyPath.Text+#0+#0;
    StrForProcess[NumPer].PathCreate:=CheckBox4.Checked;
    StrForProcess[NumPer].CancelCopyFF:=false;
    StrForProcess[NumPer].BeforeInstallCopy:=CheckBox2.Checked;
    StrForProcess[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
    StrForProcess[NumPer].OwnerForm:=Self;
    StrForProcess[NumPer].NumCount:=NumPer;
    StrForProcess[NumPer].TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
    res[NumPer]:=BeginThread(nil,0,addr(MyNewProcess),Addr(StrForProcess[NumPer]),0,treadID[NumPer]); ///
    CloseHandle(res[NumPer]);
    end;
  if (GroupPC=true)and(CheckBox1.Checked=false) then //// ������ �������� �� ������ � ������ �������
    begin
    exitfor:=true;
    SelectedPCNewProc:=TstringList.Create;
    SelectedPCNewProc:=frmDomainInfo.createListpcForCheck('');
    label2.Visible:=true;
    for I := 0 to SelectedPCNewProc.Count-1 do
      begin
      NumPer:=tagForCopy(OKBtn.tag);
      StrForProcess[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForProcess[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForProcess[NumPer].NamePC:=SelectedPCNewProc[i];
      if CheckBox2.Checked then // ���� ���� ����������� �����
      begin
        if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
        begin
        StrForProcess[NumPer].FSource:= EditSource.Text+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+ExtractFileName(EditFileRun.Text); // ������ ������� ����� ��� ���������
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(EditSource.Text)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
        end
         else // ���� �������� ���� �������� �������
        begin
        StrForProcess[NumPer].FSource:= ExtractFilePath(EditSource.Text)+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+GetLastDir(EditSource.Text)+ExtractFileName(EditFileRun.Text);
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(EditFileRun.Text)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
        end;
      end;
      if not CheckBox2.Checked then // ���� �� �������� ���� ��� ������� ����� ����������
      StrForProcess[NumPer].FileToRun:= EditFileRun.Text; //�� ���� ��� ��������� ����� �� ���������
      StrForProcess[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForProcess[NumPer].PathCreate:=CheckBox4.Checked;
      StrForProcess[NumPer].CancelCopyFF:=false;
      StrForProcess[NumPer].BeforeInstallCopy:=CheckBox2.Checked;
      StrForProcess[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForProcess[NumPer].OwnerForm:=Self;
      StrForProcess[NumPer].NumCount:=NumPer;
      StrForProcess[NumPer].TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
      res[NumPer]:=BeginThread(nil,0,addr(MyNewProcess),Addr(StrForProcess[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      Application.ProcessMessages;
      sleep(500);
      label2.Caption:=inttostr(((100 div SelectedPCNewProc.Count-1)*(i+1)))+'%';
      if not ExitFor then break;
      end;
     label2.Visible:=false;
     SelectedPCNewProc.Free;
    end;

    {{if GroupPC=false then  //// ������ �������� �� ������� ����������
    begin
    NewProcMyPS:=MyPS;   ��� ����� �������� ��� ������� ����� �� ������ ������ ���������� ����� ��������� �����
    MyNewProc:=unit5.NewProcess.Create(true);
    MyNewProc.FreeOnTerminate:=true;
    MyNewProc.Start;
    end;}
close;
except
on E:Exception do frmDomainInfo.memo1.Lines.Add('������ ������������ ������� ������� �������� - "'+E.Message+'"');
end;
end;


procedure TNewProcForm.SpeedButton1Click(Sender: TObject);
var
newOpenDlg:TopenDialog;
Newpatch,Newdescription:string;
z:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
Newpatch:='';
Newdescription:='';
newOpenDlg:=TOpendialog.Create(self);
newOpenDlg.Title:='������� ���� ��� �������';
if newOpenDlg.Execute then
  begin
   NewPatch:=InputBox('���������� ����� �������', '����:', newOpenDlg.FileName);
   if NewPatch=newOpenDlg.FileName then
     begin
     z:=MessageDlg('�� �� ������� ���� �������'+#10#13+'����������???',
     mtConfirmation, mbOKCancel, 0);
     if z=mrCancel then begin newOpenDlg.Free; exit;  end;
     end;
   Newdescription:=InputBox('���������� �������� � �����', '��������:', '����� ���� ��������');
   // ���� �� ������� �������� �� ��������� � �������� ��� ���������� �����
   if (Newdescription='') and (fileexists(NewOpenDlg.FileName)) then Newdescription:=ExtractFileName(NewOpenDlg.FileName);
  end;

if (NewPatch='') or (Newdescription='') then exit;

// ���������� ����� ������� � ����
EditSource.Text:=DeliteKeyFilePatch(NewPatch); // ��������� ���� ���� ��� �����������
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(NewPatch,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'proc'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+leftstr(Newdescription,1000)+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox2.Checked; // ���������� ����� ����� ����������
if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
FDQueryProc.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // ���� �������� ������������ �������
FDQueryProc.params.ParamByName('p5').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p6').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //���������� ���� ����������
FDQueryProc.params.ParamByName('p7').AsBoolean:=CheckBox5.Checked; // ������� ����� ����� ���������
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditSource.Text,1000)+'';    // ���� �� ����� ���������
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// ��������� ������  ��������� �� ����
readproc(ComboBox2.Items.Count);
newOpenDlg.Free;
end;

procedure TNewProcForm.SpeedButton2Click(Sender: TObject); //�������
var
i:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
if ComboBox2.Text='' then exit;

i:=MessageDlg('����� ��������� ��������� ��� ������ �� ������������ � �������.'+#10#13+'���������� ���������� ��������?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ComboBox3.text+'''';
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// ��������� ������  ��������� �� ����
readproc(ComboBox2.ItemIndex-1);
end;



procedure TNewProcForm.SpeedButton3Click(Sender: TObject);
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� ����� ��� �������';
  //Options:=[fdoForceShowHidden,fdoPickFolders]; {��������}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditFileRun.Text:=FileName;
    EditSource.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;



procedure TNewProcForm.SpeedButton4Click(Sender: TObject);  //���������
var
descript:string;
begin

if (EditFileRun.Text='') then
begin
  ShowMessage('������� ���� � ����� � �������� ��������');
  SpeedButton1.Click;
  exit;
end;

descript:=ComboBox2.Text;

if (ComboBox2.Text='')and (EditFileRun.Text<>'') then
 begin
 descript:=InputBox('�������� �������� � �����', '��������:', '����� �������� ��������');
 end;
 if descript='' then exit;

FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8) MATCHING (DESCRIPTION_PROC)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(EditFileRun.Text,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'proc'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+descript+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox2.Checked; // ���������� ���� ��� ������� ����� ����������
if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
FDQueryProc.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // ���� �������� ������������ �������
FDQueryProc.params.ParamByName('p5').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p6').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //���������� ���� ����������
FDQueryProc.params.ParamByName('p7').AsBoolean:=CheckBox5.Checked; // ������� ����� ����� ���������
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditSource.Text,1000)+'';    // ���� �� ����� ���������
FDQueryProc.ExecSQL;
FDQueryProc.close;
readproc(ComboBox2.ItemIndex);
end;

procedure TNewProcForm.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='��������';
FormEditMsiProc.Caption:='�������� ���������';
FormEditMsiProc.ShowModal;
end;

end.

