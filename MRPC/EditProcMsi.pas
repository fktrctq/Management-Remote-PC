unit EditProcMsi;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.StdCtrls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Buttons, JvExButtons, JvButtons,System.strUtils,
  Vcl.Menus;

type
  TFormEditMsiProc = class(TForm)
    ListView1: TListView;
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    ButtonPM: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    SpeedButton1: TSpeedButton;
    procedure ListView1DblClick(Sender: TObject);
    procedure ButtonPMClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);

  private
    FormKey:Tform;
    EditDescr:TLabelEdedit;
    EditPath:TLabelEdedit;
    CheckCopy:TcheckBox;
    ComboFF:TcomboBox;
    EditCreatePath:TLabelEdEdit;
    EditPathFileRun:TlabelEdEdit;
    ButtonOk:Tbutton;
    ButtonNo:Tbutton;
    FormMSI:Tform;
    EditDescrM:TLabelEdedit;
    EditPathM:TLabelEdedit;
    CheckCopyM:TcheckBox;
    CheckDelM:TcheckBox;
    ComboFFM:TcomboBox;
    EditCreatePathM:TLabelEdEdit;
    EditPathFileRunM:TlabelEdEdit;
    ButtonOkM:Tbutton;
    ButtonNoM:Tbutton;
    ButtonOpen:TSpeedbutton;
    function CreateEditFormProc(Descr,path,FileSource,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
    Function CreateEditFormMSI(Descr,FileRunMSI,OptionKey,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
    function updateList(ProcMsi:string):boolean;
    procedure ButtonNoClose(Sender: TObject);
    procedure KeyFormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonSaveMSI(Sender: TObject);
    procedure ButtonNewMSI(Sender: TObject);
    procedure ButtonSaveProc(Sender: TObject);
    procedure ButtonNewProc(Sender: TObject);
    procedure ButtonOpenDlgProc(Sender: TObject);
    procedure ButtonOpenDlgMSI(Sender: TObject);
    procedure updateListPM;

  public
    { Public declarations }
  end;

var
  FormEditMsiProc: TFormEditMsiProc;


implementation
uses MyDM,umain;
{$R *.dfm}

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

function yestobool(s:string):boolean;
begin
if s='��' then Result:=true
else Result:=false;
end;


procedure TFormEditMsiProc.updateListPM;
begin
 if ButtonPM.Caption='��������' then  // ���� ������� �������� �� ��������� ��������� � ��������
begin
  updateList('msi');
end
else
begin
 updateList('proc');
end;
end;


procedure TFormEditMsiProc.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).close;
end;

procedure TFormEditMsiProc.ButtonPMClick(Sender: TObject);
begin
if ButtonPM.Caption='��������' then
begin
  ListView1.Column[3].Width:=0; //������� � ����������� MSI
  ListView1.Column[5].Width:=0; // ������� ����� ���������/�������
  ListView1.Column[8].Width:=200; // ���� ��������
  ListView1.Column[2].Caption:='������ ������� � �����������';
  ButtonPM.Caption:='��������� MSI';
  updateList('proc');
  FormEditMsiProc.caption:='�������� ���������';
end
else
begin
 ListView1.Column[3].Width:=100; //������� � ����������� MSI
 ListView1.Column[5].Width:=100; // ������� ����� ���������/�������
 ListView1.Column[8].Width:=0; // ���� ��������
 ListView1.Column[2].Caption:='���� ���������';
 ButtonPM.Caption:='��������';
 updateList('msi');
 FormEditMsiProc.caption:='�������� �������� msi';
end;
end;

procedure TFormEditMsiProc.KeyFormClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;


procedure TFormEditMsiProc.ListView1DblClick(Sender: TObject);
begin
if ListView1.SelCount=1 then
begin
if ButtonPM.Caption='��������� MSI' then
begin
CreateEditFormProc(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[7],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end
else
begin
CreateEditFormMSI(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[2],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end;
end;
end;

procedure TFormEditMsiProc.N1Click(Sender: TObject); // �������� ����� ������
begin
if ListView1.SelCount=1 then // ���� ���� ���������� ������ �� ������ ������ � ����
  begin
   if ButtonPM.Caption='��������� MSI' then // ���� ������ ��������� �� ������� ������ ��������� � ��������
    begin
    CreateEditFormProc(
    ListView1.Selected.SubItems[0],
    ListView1.Selected.SubItems[1],
    ListView1.Selected.SubItems[7],
    ListView1.Selected.SubItems[5],
    ListView1.Selected.SubItems[6],
    yestobool(ListView1.Selected.SubItems[3]),
    yestobool(ListView1.Selected.SubItems[4])
    ,true)
    end
   else 
    begin
    CreateEditFormMSI(
    ListView1.Selected.SubItems[0],
    ListView1.Selected.SubItems[1],
    ListView1.Selected.SubItems[2],
    ListView1.Selected.SubItems[5],
    ListView1.Selected.SubItems[6],
    yestobool(ListView1.Selected.SubItems[3]),
    yestobool(ListView1.Selected.SubItems[4])
    ,true)
    end;
  end
else // ����� ���� ������ �� �������� ��� �������� ������ ��� 1 �� ���� �������� ��������� �������
  begin
  if ButtonPM.Caption='��������� MSI' then
  CreateEditFormProc('','','','','C$\TEMP\',false,false,true)// ����� ��������
  else
  CreateEditFormMSI('','','','','C$\TEMP\',false,false,true); // ����� MSI
  end;
end;

procedure TFormEditMsiProc.Button1Click(Sender: TObject);
begin
if ButtonPM.Caption='��������� MSI' then
CreateEditFormProc('','','','','C$\TEMP\',false,false,true)// ����� ��������
else
CreateEditFormMSI('','','','','C$\TEMP\',false,false,true); // ����� MSI
end;

procedure TFormEditMsiProc.Button2Click(Sender: TObject);
begin
if ListView1.SelCount=1 then
begin
if ButtonPM.Caption='��������� MSI' then
begin
CreateEditFormProc(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[7],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end
else
begin
CreateEditFormMSI(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[2],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end;
end;
end;



procedure TFormEditMsiProc.Button3Click(Sender: TObject);
begin
updateListPM;
end;

procedure TFormEditMsiProc.Button4Click(Sender: TObject); // �������� ������
var
i,z:integer;
QueryDel: TFDQuery;
TransactionDel:TFDtransaction;
begin
if ListView1.SelCount=0 then  exit;
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
i:=MessageDlg('����� ��������� ��������� ��� ������ �� ������������ � �������.'+#10#13+'���������� ���������� ��������?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
try
TransactionDel:= TFDTransaction.Create(nil);
TransactionDel.Connection:=DataM.ConnectionDB;
TransactionDel.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryDel:=TFDQuery.Create(nil);
QueryDel.Transaction:=TransactionDel;
QueryDel.Connection:=DataM.ConnectionDB;
try
for z := ListView1.Items.Count-1 downto 0 do
if ListView1.Items[z].Selected then
begin
TransactionDel.StartTransaction;
QueryDel.SQL.Clear;
QueryDel.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ListView1.Items[z].Caption+''''; // � Caption ��������� ID ��������� ��� MSI
QueryDel.ExecSQL;
end;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('������ ��������. '+e.Message);
end;
finally
QueryDel.Close;
TransactionDel.Commit;
TransactionDel.Free;
QueryDel.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonOpenDlgProc(Sender: TObject);
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� ����� ��� �������';
  if EditPathFileRun.Text<>'' then if FileExists(EditPathFileRun.Text) then InitialDir:=ExtractFileDir(EditPathFileRun.Text) ;
  //Options:=[fdoForceShowHidden,fdoPickFolders]; {��������}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditPath.Text:=FileName;
    //EditPathFileRun.Text:=DeliteKeyFilePatch(EditPath.Text)
    EditPathFileRun.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;

procedure TFormEditMsiProc.ButtonOpenDlgMSI(Sender: TObject);
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� ����� ��� ���������';
  Filter:='|*msi';
  if EditPathM.Text<>'' then if FileExists(EditPathM.Text) then InitialDir:=ExtractFileDir(EditPathM.Text) ;
  if Execute then
    begin
    EditPathM.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;

procedure TFormEditMsiProc.ButtonSaveMSI(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.clear;
QueryWrite.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE,KEY_MSI)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)MATCHING (DESCRIPTION_PROC)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPathM.Text,1000)+''; //���� � �����
QueryWrite.params.ParamByName('p2').AsString:=''+'msi'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescrM.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopyM.Checked; // ���������� ����� ����� ����������
QueryWrite.params.ParamByName('p5').AsBoolean:=CheckDelM.Checked; // ������� ����� ����� ���������
if ComboFFM.ItemIndex=0 then // ���� �������� ������ ����
QueryWrite.params.ParamByName('p6').AsString:=''+'File'+'';
if ComboFFM.ItemIndex=1 then // ���� �������� ������������ �������
QueryWrite.params.ParamByName('p6').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p7').AsString:=''+leftstr(EditCreatePathM.text,1024)+''; //���������� ���� ����������
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRunM.text,250)+'';  // �����/����� �������
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('������ ����������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonNewMSI(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.clear;
QueryWrite.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,KEY_MSI,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPathM.Text,1000)+''; // ���� � �����
QueryWrite.params.ParamByName('p2').AsString:=''+'msi'+'';                            // ���
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescrM.Text,1000)+'';     //��������
QueryWrite.params.ParamByName('p4').AsString:=''+leftstr(EditPathFileRunM.text,250)+'';          // ����� �������
QueryWrite.params.ParamByName('p5').AsBoolean:=CheckCopyM.Checked; // ���������� ����� ����� ����������
QueryWrite.params.ParamByName('p6').AsBoolean:=CheckDelM.Checked; // ������� ����� ����� ���������
if ComboFFM.ItemIndex=0 then // ���� �������� ������ ����
QueryWrite.params.ParamByName('p7').AsString:=''+'File'+'';
if ComboFFM.ItemIndex=1 then // ���� �������� ������������ �������
QueryWrite.params.ParamByName('p7').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditCreatePathM.text,1024)+''; //���������� ���� ����������
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('������ �������� ����� ���������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonSaveProc(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.Clear;
QueryWrite.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8) MATCHING (DESCRIPTION_PROC)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,1000)+'';
QueryWrite.params.ParamByName('p2').AsString:=''+'proc'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopy.Checked; // ���������� ���� ��� ������� ����� ����������
if ComboFF.ItemIndex=0 then // ���� �������� ������ ����
QueryWrite.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboFF.ItemIndex=1 then // ���� �������� ������������ �������
QueryWrite.params.ParamByName('p5').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p6').AsString:=''+leftstr(EditCreatePath.Text,1024)+''; //���������� ���� ����������
QueryWrite.params.ParamByName('p7').AsBoolean:=false; // ������� ����� ����� ���������
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRun.Text,1000)+'';    // ���� �� ����� ���������
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('������ ���������� ��������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonNewProc(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
EditPathFileRun.Text:=DeliteKeyFilePatch(EditPath.Text); // ��������� ���� ���� ��� �����������
QueryWrite.SQL.Clear;
QueryWrite.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,1000)+'';
QueryWrite.params.ParamByName('p2').AsString:=''+'proc'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopy.Checked; // ���������� ����� ����� ����������
if ComboFF.ItemIndex=0 then // ���� �������� ������ ����
QueryWrite.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboFF.ItemIndex=1 then // ���� �������� ������������ �������
QueryWrite.params.ParamByName('p5').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p6').AsString:=''+leftstr(EditCreatePath.Text,1024)+''; //���������� ���� ����������
QueryWrite.params.ParamByName('p7').AsBoolean:=false; // ������� ����� ����� ���������
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRun.Text,1000)+'';    // ���� �� ����� ���������
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('������ �������� ��������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

Function TFormEditMsiProc.CreateEditFormMSI(Descr,FileRunMSI,OptionKey,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
var
step:integer;
begin
try
step:=1;
FormKey:=TForm.Create(FormEditMsiProc);
if NewOrEdit then FormKey.Caption:='����� ���������'
else FormKey.Caption:='������������� *.msi '+descr;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=360;   //285
FormKey.Height:=290;
FormKey.OnClose:=KeyFormClose;
////////////
step:=2;
EditDescrM:=TLabelEdedit.Create(FormKey);
EditDescrM.Parent:=FormKey;
EditDescrM.EditLabel.caption:='��������';
EditDescrM.Text:=Descr;
EditDescrM.Top:=15;
EditDescrM.Left:=5;
EditDescrM.Width:=335;
/// /////////////////////
step:=3;
EditPathM:=TLabelEdedit.Create(FormKey);
EditPathM.Parent:=FormKey;
EditPathM.EditLabel.Caption:='���� �������';
EditPathM.Text:=FileRunMSI;
EditPathM.Top:=55;
EditPathM.Left:=5;
EditPathM.Width:=308;//335
/////////////////////////
ButtonOpen:=TSpeedbutton.Create(FormKey);
ButtonOpen.Parent:=FormKey;
ButtonOpen.Caption:='';
ButtonOpen.Height:=23;
ButtonOpen.Width:=23;
ButtonOpen.Glyph:=SpeedButton1.Glyph;
ButtonOpen.Flat:=true;
ButtonOpen.Top:=54;
ButtonOpen.Left:=315;
ButtonOpen.OnClick:=ButtonOpenDlgMSI;
/////////////////////////
step:=4;
EditPathFileRunM:=TLabeledEdit.Create(FormKey);
EditPathFileRunM.Parent:=FormKey;
EditPathFileRunM.EditLabel.Caption:='�����, �� ����������� (���������� ����� msiexec /help)';
EditPathFileRunM.Text:=OptionKey;
EditPathFileRunM.Top:=95;
EditPathFileRunM.Left:=5;
EditPathFileRunM.Width:=335;
////////////////////////
step:=5;
CheckCopyM:=TcheckBox.Create(FormKey);
CheckCopyM.Parent:=FormKey;
CheckCopyM.Checked:=Fcopy;
CheckCopyM.Top:=115;
CheckCopyM.Left:=5;
CheckCopyM.Width:=335;
CheckCopyM.Caption:='���������� ����������� ����� ��������';
//////////////////
step:=6;
CheckDelM:=TcheckBox.Create(FormKey);
CheckDelM.Parent:=FormKey;
CheckDelM.Checked:=Fdel;
CheckDelM.Top:=130;
CheckDelM.Left:=5;
CheckDelM.Width:=335;
CheckDelM.Caption:='������� ������������� ����� ����� ���������';
///////////////////
step:=7;
ComboFFM:=TComboBox.Create(FormKey);
ComboFFM.parent:=FormKey;
ComboFFM.Style:=csOwnerDrawFixed;
ComboFFM.Top:=150;
ComboFFM.Left:=5;
ComboFFM.Width:=335;
ComboFFM.Items.Add('���������� ������ ���� ��������� (msi)');
ComboFFM.Items.Add('���������� ������������ ������� ����� ���������');
if FF='Folder' then ComboFFM.ItemIndex:=1
else ComboFFM.ItemIndex:=0;
//////////////////
step:=8;
EditCreatePathM:=TLabelEdEdit.Create(FormKey);
EditCreatePathM.parent:=FormKey;
EditCreatePathM.EditLabel.Caption:='������� ����������';
EditCreatePathM.Text:=CreatePath;
EditCreatePathM.Top:=190;
EditCreatePathM.Left:=5;
EditCreatePathM.width:=335;
//////////////////
ButtonOkM:=Tbutton.Create(FormKey);
ButtonOkM.Parent:=FormKey;
ButtonOkM.Caption:='���������';
ButtonOkM.Top:=220;
ButtonOkM.Left:=185;
if NewOrEdit then ButtonOkM.OnClick:=ButtonNewMSI  // ���� ������� �����
else ButtonOkM.OnClick:=ButtonSaveMSI;  // ������ ������������� ������������
ButtonOkM.TabOrder:=3;
//////////////////
ButtonNoM:=Tbutton.Create(FormKey);
ButtonNoM.Parent:=FormKey;
ButtonNoM.Caption:='�������';
ButtonNoM.Top:=220;
ButtonNoM.Left:=265;
ButtonNoM.OnClick:=ButtonNoClose;
ButtonNoM.TabOrder:=4;
result:=true;
FormKey.ShowModal;
except on E: Exception do
begin
  frmDomainInfo.Memo1.Lines.Add('������ �������� ������� �������������� msi. ('+inttostr(step)+') '+e.Message);
  result:=false;
end;
end;
end;

function TFormEditMsiProc.CreateEditFormProc(Descr,path,FileSource,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
begin
try
FormKey:=TForm.Create(FormEditMsiProc);
if NewOrEdit then  FormKey.Caption:='����� �������'
else FormKey.Caption:='������������� ������� '+descr;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=360;   
FormKey.Height:=275;
FormKey.OnClose:=KeyFormClose;
////////////
EditDescr:=TLabelEdedit.Create(FormKey);
EditDescr.Parent:=FormKey;
EditDescr.EditLabel.caption:='��������';
EditDescr.Text:=Descr;
EditDescr.Top:=15;
EditDescr.Left:=5;
EditDescr.Width:=335;
/// /////////////////////
EditPath:=TLabelEdedit.Create(FormKey);
EditPath.Parent:=FormKey;
EditPath.EditLabel.Caption:='������ ������� � �����������';
EditPath.Text:=path;
EditPath.Top:=55;
EditPath.Left:=5;
EditPath.Width:=308;
/////////////////////////
ButtonOpen:=TSpeedbutton.Create(FormKey);
ButtonOpen.Parent:=FormKey;
ButtonOpen.Caption:='';
ButtonOpen.Height:=23;
ButtonOpen.Width:=23;
ButtonOpen.Glyph:=SpeedButton1.Glyph;
ButtonOpen.Flat:=true;
ButtonOpen.Top:=54;
ButtonOpen.Left:=315;
ButtonOpen.OnClick:=ButtonOpenDlgProc;
///*/////////////////////
EditPathFileRun:=TLabeledEdit.Create(FormKey);
EditPathFileRun.Parent:=FormKey;
EditPathFileRun.EditLabel.Caption:='���� ��������';
EditPathFileRun.Text:=FileSource;
EditPathFileRun.Top:=95;
EditPathFileRun.Left:=5;
EditPathFileRun.Width:=335;
////////////////////////
CheckCopy:=TcheckBox.Create(FormKey);
CheckCopy.Parent:=FormKey;
CheckCopy.Checked:=Fcopy;
CheckCopy.Top:=115;
CheckCopy.Left:=5;
CheckCopy.Width:=335;
CheckCopy.Caption:='���������� ���� ����� ��������';
//////////////////
ComboFF:=TComboBox.Create(FormKey);
ComboFF.parent:=FormKey;
ComboFF.Items.Add('���������� ������ ����');
ComboFF.Items.Add('���������� ������������ ������� �����');
if FF='Folder' then ComboFF.ItemIndex:=1
else ComboFF.ItemIndex:=0;
ComboFF.Style:=csOwnerDrawFixed;
ComboFF.Top:=135;
ComboFF.Left:=5;
ComboFF.Width:=335;
//////////////////
EditCreatePath:=TLabelEdEdit.Create(FormKey);
EditCreatePath.parent:=FormKey;
EditCreatePath.EditLabel.Caption:='������� ����������';
EditCreatePath.Text:=CreatePath;
EditCreatePath.Top:=175;
EditCreatePath.Left:=5;
EditCreatePath.width:=335;
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='���������';
ButtonOk.Top:=205;
ButtonOk.Left:=185;
if NewOrEdit then ButtonOk.OnClick:=ButtonNewProc  // ���� ������ ������� �����
else ButtonOk.OnClick:=ButtonSaveProc;  // ���� �������������
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='�������';
ButtonNo.Top:=205;
ButtonNo.Left:=265;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
result:=true;
FormKey.ShowModal;
except on E: Exception do
begin
  frmDomainInfo.Memo1.Lines.Add('������ �������� ������� �������������� ��������. '+e.Message);
  result:=false;
end;
end;
end;


procedure TFormEditMsiProc.FormShow(Sender: TObject);
begin
if ButtonPM.Caption='��������' then
begin
  ListView1.Column[3].Width:=0; //������� � ����������� MSI
  ListView1.Column[5].Width:=0; // ������� ����� ���������/�������
  ListView1.Column[8].Width:=100; // ���� ��������
  ListView1.Column[2].Caption:='������ ������� � �����������';
  updateList('proc');
  ButtonPM.Caption:='��������� MSI';
end
else
begin
 ListView1.Column[3].Width:=100; //������� � ����������� MSI
 ListView1.Column[5].Width:=100; // ������� ����� ���������/�������
 ListView1.Column[8].Width:=0; // ���� ��������
 ListView1.Column[2].Caption:='���� ���������';
 updateList('msi');
 ButtonPM.Caption:='��������';
end;
end;

function TFormEditMsiProc.updateList(ProcMsi:string):boolean;
var
QueryRead: TFDQuery;
TransactionRead:TFDtransaction;
begin
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
TransactionRead.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryRead:=TFDQuery.Create(nil);
QueryRead.Transaction:=TransactionRead;
QueryRead.Connection:=DataM.ConnectionDB;
try
ListView1.Clear;
TransactionRead.StartTransaction;
QueryRead.SQL.clear;
QueryRead.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC='''+ProcMsi+''''
+' ORDER BY DESCRIPTION_PROC'; //ID_PROC
QueryRead.Open;

while not QueryRead.Eof do
begin
 if QueryRead.FieldByName('ID_PROC').Value<>null then
 begin
 with ListView1.Items.Add do
 begin
   Caption:=QueryRead.FieldByName('ID_PROC').AsString;
   SubItems.add(QueryRead.FieldByName('DESCRIPTION_PROC').AsString);
   SubItems.add(QueryRead.FieldByName('PATCH_PROC').AsString);
   SubItems.add(QueryRead.FieldByName('KEY_MSI').AsString);
   if QueryRead.FieldByName('BEFOREINSTALLCOPY').AsBoolean then SubItems.add('��')
   else  SubItems.add('���');
   if QueryRead.FieldByName('DELETEAFTERINSTALL').AsBoolean then SubItems.add('��')
   else  SubItems.add('���');
   SubItems.add(QueryRead.FieldByName('FILEORFOLDER').AsString);
   SubItems.add(QueryRead.FieldByName('PATHCREATE').AsString);
   SubItems.add(QueryRead.FieldByName('FILESOURSE_PROC').AsString);
 end;
 end;
  QueryRead.Next;
end;

finally
QueryRead.Close;
TransactionRead.Commit;
TransactionRead.Free;
QueryRead.Free;
end;

except on E: Exception do
frmDomainInfo.Memo1.Lines.Add('������ ������ ������ '+e.Message);
end;
end;

end.
