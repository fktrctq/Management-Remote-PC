unit DlgNewInstall;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Dialogs,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.variants,System.StrUtils,IdIcmpClient,ComObj,ActiveX,ShellAPI;

  type
 TStrForInstallProgram = record
    FSource :String;   /// �������� ����
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
    InstallAllUsers:boolean; // ������������ ��� ���� �������������
  end;
type
  TOKRightDlg1 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    Label1: TLabel;
    Edit1: TEdit;
    CheckBox2: TCheckBox;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    DataSource1: TDataSource;
    SpeedButton3: TSpeedButton;
    Label2: TLabel;
    CheckBox3: TCheckBox;
    ComboBox1: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button1: TButton;
    EditCopyPath: TLabeledEdit;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    SpeedButton4: TSpeedButton;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Label4: TLabel;
    SpeedButton5: TSpeedButton;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    function  readproc(item:integer):boolean;
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton3Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2Select(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1: TOKRightDlg1;
  ExitFor:Boolean;
  StrForInstall: array [0..2000] of TStrForInstallProgram; //������ ���������� ��� ������� ������� ��������� ��������
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;

 ThreadVar
PointForInstallMSI: ^TStrForInstallProgram;

  implementation
uses umain,MyDM,EditProcMSI;

{$R *.dfm}

function RunInstallMSI(param:pointer):boolean;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  Resinstall,b,z,i :integer;
  CopyOrNo:boolean;
  ListPC:TstringList;
////////////////////////////////////////////////////
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
      frmdomaininfo.Memo1.Lines.Add(CurentPC+' :�������� �������� �����������');
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // ��� �������  ���� �������� ��������
      pTo := pchar('');   // ���� �������� �� ������������
      frmdomaininfo.Memo1.Lines.Add(CurentPC+' :�������� �������� �������� ������������');
      //finditem('�������� �������� �������� ������������',CurentPC,13);
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
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :�������� ����������� ������� ���������');
          end;
        if TypeOperation=3 then
          begin
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :�������� �������� ������� ���������'); //finditem('�������� �������� ������� ���������',CurentPC,1);
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :�������� �������� ������� ���������');
          end;
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then
        begin
        finditem('������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        frmdomaininfo.Memo1.Lines.Add(CurentPC+' :������� ����������� ����������� ������� - '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        if TypeOperation=3 then
        begin
        frmdomaininfo.Memo1.Lines.Add(CurentPC+' :������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        //finditem('������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
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
       //finditem('������ ������� ��������: '+E.Message,CurentPC,2);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then frmdomaininfo.Memo1.Lines.Add(Format('����� ������ ������� ����������� ����� ���������� ���������  :  %s',[E.Message]));
     if TypeOperation=3 then frmdomaininfo.Memo1.Lines.Add(Format('����� ������ ������� �������� ����� ��������� ���������  :  %s',[E.Message]));
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////
BEGIN
PointForInstallMSI:=param;
ListPC:=Tstringlist.Create;
ListPC.CommaText:=PointForInstallMSI.NamePC;
for I := 0 to ListPC.Count-1 do
Begin
 if ping(ListPC[i]) then
  begin
  CopyOrNo:=false;
  if PointForInstallMSI.BeforeInstallCopy then // ���� ���������� ����������� ����� ����������
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // ��� �����
  PointForInstallMSI.UserName,//�����
  PointForInstallMSI.PassWd,  // ������
  PointForInstallMSI.FSource, // ��� ����������
  PointForInstallMSI.FDest,   // ���� ����������
  PointForInstallMSI.OwnerForm,// ������������ �����
  PointForInstallMSI.TypeOperation, //��� ��������
  PointForInstallMSI.CancelCopyFF,
  PointForInstallMSI.PathCreate) ;
  end;
      try
          frmDomainInfo.memo1.Lines.Add('������� ������� ��������� ��������� '+ExtractFileName(PointForInstallMSI.ProgramInstall)+' �� '+ListPC[i]+'.');
          finditem('������� ������� ��������� ��������� '+ExtractFileName(PointForInstallMSI.ProgramInstall),ListPC[i],14);
          OleInitialize(nil);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService  := FSWbemLocator.ConnectServer(ListPC[i], 'root\CIMV2', PointForInstallMSI.UserName, PointForInstallMSI.PassWd,'','',128);
          FWbemObject   := FWMIService.Get('Win32_Product');
          Resinstall:=FWbemObject.install(PointForInstallMSI.ProgramInstall,PointForInstallMSI.KeyNewInstallProgram,PointForInstallMSI.InstallAllUsers); //���� ���������, ����� ���������, ��� ���� �������������
          if Resinstall=0 then
            Begin
            finditem('��������� ��������� ' +ExtractFileName(PointForInstallMSI.ProgramInstall)
            +' : '+SysErrorMessage(Resinstall),ListPC[i],1);
            End
          else
             Begin
             finditem('��� ��������� ��������� '
                  +ExtractFileName(PointForInstallMSI.ProgramInstall) +' �������� ������ : '
                  +SysErrorMessage(Resinstall),ListPC[i],2)
             End;
          frmDomainInfo.memo1.Lines.Add('��������� ��������� '
          +ExtractFileName(PointForInstallMSI.ProgramInstall)+' �� '
          +ListPC[i]+' : '+SysErrorMessage(Resinstall));
          FWbemObject:=Unassigned;
          VariantClear(FWbemObject);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
         if PointForInstallMSI.BeforeInstallCopy and  //���� ������� ����������� ����� ����������
         PointForInstallMSI.DeleteAfterInstall and    // ���� ���������� ������� ����������� ����� ���������
         CopyOrNo then // ���� �������� ����������� ����� ���������� ������ �������
          begin        // �������
          CopyFFSelectPC(ListPC[i], // ��� �����
          PointForInstallMSI.UserName,//�����
          PointForInstallMSI.PassWd,  // ������
          '', // ��� ���������� ����� �� ��������� �.�. �������
          PointForInstallMSI.PathDelete,   // ��� �������
          PointForInstallMSI.OwnerForm,// ������������ �����
          3, //��� ��������  (3)- �������, FO_MOVE
          PointForInstallMSI.CancelCopyFF,
          false) ;  // �� ��������� ������� �������� �.�. ������� ����� �����������
          end;

          OleUnInitialize;
        except
          on E:Exception do
           Begin
           finditem('������ ��������� ��������� '
           +ExtractFileName(PointForInstallMSI.ProgramInstall) +' : '+E.Message,ListPC[i],2);
             frmDomainInfo.memo1.Lines.Add('������ ��������� ��������� '+ExtractFileName(PointForInstallMSI.ProgramInstall)+' �� '+ListPC[i]+' : "'+E.Message+'"');
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             VariantClear(FWbemObject);
             VariantClear(FWMIService);
             VariantClear(FSWbemLocator);
             OleUnInitialize;
           End;
       end; // except
  end; //ping
  End; // ����

END;

function  TOKRightDlg1.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''msi'''
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


procedure TOKRightDlg1.Button1Click(Sender: TObject);
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

procedure TOKRightDlg1.CancelBtnClick(Sender: TObject);
begin
if (CheckBox2.Checked=false) then ExitFor:=true;
Close;
end;

procedure TOKRightDlg1.CheckBox3Click(Sender: TObject);
begin
GroupBox1.Enabled:=CheckBox3.Checked;
if CheckBox3.Checked then groupbox1.Height:=141
else  groupbox1.Height:=25;
end;

procedure TOKRightDlg1.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
ComboBox2.DroppedDown:=true;
end;

procedure TOKRightDlg1.ComboBox2Select(Sender: TObject);
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
LabeledEdit1.Text:=vartostr(FDQuery.FieldByName('PATCH_PROC').Value);
Edit1.Text:=vartostr(FDQuery.FieldByName('KEY_MSI').Value);
if FDQuery.FieldByName('KEY_MSI').Value<>null then  Edit1.Text:=FDQuery.FieldByName('KEY_MSI').AsString;
if FDQuery.FieldByName('BEFOREINSTALLCOPY').Value<>null then
begin
CheckBox3.Checked:=FDQuery.FieldByName('BEFOREINSTALLCOPY').AsBoolean;
if FDQuery.FieldByName('DELETEAFTERINSTALL').Value<>null then  CheckBox5.Checked:=FDQuery.FieldByName('DELETEAFTERINSTALL').AsBoolean;
if FDQuery.FieldByName('FILEORFOLDER').Value<>null then
  begin
  if FDQuery.FieldByName('FILEORFOLDER').AsString='File' then combobox1.ItemIndex:=0
  else combobox1.ItemIndex:=1;
  end;
if FDQuery.FieldByName('PATHCREATE').Value<>null then EditCopyPath.Text:=FDQuery.FieldByName('PATHCREATE').AsString;
end;
 finally
FDQuery.Close;
FDQuery.Free;
end;
end;


procedure TOKRightDlg1.FormClose(Sender: TObject; var Action: TCloseAction);
begin   /// ���� ��� �������� ����, ������ ������ � combobox �� ��������� ���, ����� ������
label3.Visible:=false;
end;


procedure TOKRightDlg1.FormShow(Sender: TObject);
begin
label3.Visible:=false;
CheckBox2.Visible:=GroupPC;
CheckBox3.Checked:=false;
GroupBox1.Height:=25;
if not DataM.ConnectionDB.Connected then
begin
 exit;
end;
readproc(-1);
end;

procedure TOKRightDlg1.OKBtnClick(Sender: TObject);
var
i,NumPer:integer;
SelectedPCInstallProg:TstringList;
function tagForCopy(num:integer):integer;
begin
if OKBtn.tag=1999 then num:=0;
inc(num);
OKBtn.tag:=num;
result:=num;
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

Function renewPath(s:string):string;
begin
if AnsiPos('$',s)=2 then
 begin
 delete(s,2,1); // ������� ������ $
 insert(':',s,2); // ��������� �� ��� ����� :
 end;
result:=s;
end;
begin
   if (GroupPC=true)and (CheckBox2.Checked=false) then ///// ��������� ��������� ��������� � ������ �������
      begin
      SelectedPCInstallProg:=TstringList.Create;
      SelectedPCInstallProg:=frmdomaininfo.createListpcForCheck('');
      if SelectedPCInstallProg.Count=0 then
      begin
        ShowMessage('� ������ ��� ���������� �����������, �������� ���������');
        SelectedPCInstallProg.Free;
        Exit;
      end;
      ExitFor:=false;
      label3.Caption:='%';
      label3.Visible:=true;
        for I := 0 to SelectedPCInstallProg.Count-1 do
          begin
          if ExitFor then break;
          NumPer:=tagForCopy(OKBtn.tag);
          if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
          begin
          StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
          StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // ������ ������� ����� ��� ���������
          StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
          end
          else // ���� �������� ���� �������� �������
          begin
          StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
          StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
          StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
          end;
          if not CheckBox3.Checked then // ���� �� �������� ���� ��� ������� ����� ����������
          StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //�� ���� ��� ��������� ����� �� ���������

          StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
          StrForInstall[NumPer].NamePC:=SelectedPCInstallProg[i];
          StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
          StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
          StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
          StrForInstall[NumPer].CancelCopyFF:=false;
          StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
          StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
          StrForInstall[NumPer].OwnerForm:=Self;
          StrForInstall[NumPer].NumCount:=NumPer;
          StrForInstall[NumPer].TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
          StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
          StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
          res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
          CloseHandle(res[NumPer]);
          label3.Caption:=inttostr (((100 div SelectedPCInstallProg.Count-1)*i+1))+'%';
          Application.ProcessMessages;
          sleep(500);
          end;
       if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
      end;

   if GroupPC=false then    /// ���������  ��������� �� ������� ���������
      begin
      if not frmDomainInfo.ping(frmDomainInfo.ComboBox2.Text) then exit;
      if LabeledEdit1.text='' then
      begin
        ShowMessage('��� ���� � ����� ���������');
        exit;
      end;
      NumPer:=tagForCopy(OKBtn.tag);
      if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
        begin
        StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // ������ ������� ����� ��� ���������
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
        end
      else // ���� �������� ���� �������� �������
        begin
        StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
        end;
      if not CheckBox3.Checked then // ���� �� �������� ���� ��� ������� ����� ����������
      StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //�� ���� ��� ��������� ����� �� ���������

      StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForInstall[NumPer].NamePC:=frmDomainInfo.ComboBox2.Text;
      StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
      StrForInstall[NumPer].CancelCopyFF:=false;
      StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
      StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForInstall[NumPer].OwnerForm:=Self;
      StrForInstall[NumPer].NumCount:=NumPer;
      StrForInstall[NumPer].TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
      StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
      StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
      res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      end;
    if (GroupPC=true)and (CheckBox2.Checked)  then  // ��������� ��������� ��������� � ����� ������
      begin
       SelectedPCInstallProg:=TstringList.Create;
      SelectedPCInstallProg:=frmdomaininfo.createListpcForCheck('');
      if SelectedPCInstallProg.Count=0 then
        begin
        ShowMessage('� ������ ��� ���������� �����������, �������� �� ����� ���� ����������');
        SelectedPCInstallProg.Free;
        Exit;
        end;
       NumPer:=tagForCopy(OKBtn.tag);
      if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
        begin
        StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // ������ ������� ����� ��� ���������
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // ���� � ���������� ����� ����� ��������� ���������
        end
      else // ���� �������� ���� �������� �������
        begin
        StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // ���� � ���������� �������� ����� ��������� ���������
        end;
        if not CheckBox3.Checked then // ���� �� �������� ���� ��� ������� ����� ����������
          StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //�� ���� ��� ��������� ����� �� ���������

      StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForInstall[NumPer].NamePC:=SelectedPCInstallProg.CommaText;
      StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
      StrForInstall[NumPer].CancelCopyFF:=false;
      StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
      StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForInstall[NumPer].OwnerForm:=Self;
      StrForInstall[NumPer].NumCount:=NumPer;
      StrForInstall[NumPer].TypeOperation:=2; // //2 - copy ��� ��������  FO_DELETE (3)- �������, FO_MOVE (1) - ����������� ,FO_RENAME (4) - �����������
      StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
      StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
      res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
      end;
if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
OKRightDlg1.Close;
end;

procedure TOKRightDlg1.SpeedButton1Click(Sender: TObject);
var
newOpenDlg:TopenDialog;
keyrunmsi,Newdescription:string;
z:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
CheckBox3.Checked:=false; // ����� ������� "���������� ����� ����������"
keyrunmsi:='';
Newdescription:='';
newOpenDlg:=TOpendialog.Create(self);
newOpenDlg.Title:='������� ���� ��� ���������';
newOpenDlg.Filter:='|*msi';
if newOpenDlg.Execute then
  begin
   keyrunmsi:=InputBox('���������� ����� ���������, �� �����������', '����:',
    '');
   if keyrunmsi<>'' then
     begin
     z:=MessageDlg('�� ������� ���� �������,'+#10#13+' ��� �� �����������! ����������???',
     mtConfirmation, mbOKCancel, 0);
     if z=mrCancel then begin newOpenDlg.Free; exit;  end;
     end;
   Newdescription:=InputBox('���������� �������� � �����', '��������:', '����� ���� ��������');
   // ���� �� ������� �������� �� ��������� � �������� ��� ���������� �����
   if (Newdescription='') and (fileexists(NewOpenDlg.FileName)) then Newdescription:=ExtractFileName(NewOpenDlg.FileName);
  end;

if (Newdescription='') or (NewOpenDlg.FileName='') then exit;
// ���������� ����� ������ � ����
FDQueryProc.SQL.Clear;
//BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE
FDQueryProc.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,KEY_MSI,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(newOpenDlg.FileName,1000)+''; // ���� � �����
FDQueryProc.params.ParamByName('p2').AsString:=''+'msi'+'';                            // ���
FDQueryProc.params.ParamByName('p3').AsString:=''+leftstr(Newdescription,1000)+'';     //��������
FDQueryProc.params.ParamByName('p4').AsString:=''+leftstr(keyrunmsi,250)+'';          // ����� �������
FDQueryProc.params.ParamByName('p5').AsBoolean:=CheckBox3.Checked; // ���������� ����� ����� ����������
FDQueryProc.params.ParamByName('p6').AsBoolean:=CheckBox5.Checked; // ������� ����� ����� ���������
if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
FDQueryProc.params.ParamByName('p7').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // ���� �������� ������������ �������
FDQueryProc.params.ParamByName('p7').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //���������� ���� ����������
FDQueryProc.ExecSQL;
// ��������� ������ �� ����
readproc(ComboBox2.Items.Count); //����� ���������� ������� ��������� ������� ������
///����������  ������
newOpenDlg.Free;
end;

procedure TOKRightDlg1.SpeedButton2Click(Sender: TObject);
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
FDQueryProc.Active:=false;
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''
+ComboBox3.Text+'''';
FDQueryProc.ExecSQL;
// ��������� ������  ��������� �� ����
readproc(ComboBox2.ItemIndex-1);
end;

procedure TOKRightDlg1.SpeedButton3Click(Sender: TObject);
begin
if OKRightDlg1.findComponent('MyDialog')=nil then
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� ����� ��� ���������';
  Filter:='|*.msi';
  if Execute then
    begin
    LabeledEdit1.text:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
end;
end;



procedure TOKRightDlg1.SpeedButton4Click(Sender: TObject); /// ��������� ������� ��������� ��� ��������� ���������
var
descript:string;
begin

if (LabeledEdit1.Text='') then
begin
  ShowMessage('������� ���� � ����� ��������� � �������� ���������');
  SpeedButton1.Click;
  exit;
end;

descript:=ComboBox2.Text;

if (ComboBox2.Text='')and (LabeledEdit1.Text<>'') then
 begin
 descript:=InputBox('�������� �������� � ���������', '��������:', '����� �������� ���������');
 end;
 if descript='' then exit;


FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE,KEY_MSI)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)MATCHING (DESCRIPTION_PROC)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(LabeledEdit1.Text,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'msi'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+descript+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox3.Checked; // ���������� ����� ����� ����������
FDQueryProc.params.ParamByName('p5').AsBoolean:=CheckBox5.Checked; // ������� ����� ����� ���������
if ComboBox1.ItemIndex=0 then // ���� �������� ������ ����
FDQueryProc.params.ParamByName('p6').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // ���� �������� ������������ �������
FDQueryProc.params.ParamByName('p6').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p7').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //���������� ���� ����������
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(Edit1.Text,250)+'';
FDQueryProc.ExecSQL;
//readproc(combobox2.ItemIndex);
end;

procedure TOKRightDlg1.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='��������� MSI';
FormEditMsiProc.Caption:='�������� ��������';
FormEditMsiProc.ShowModal;
end;

end.
