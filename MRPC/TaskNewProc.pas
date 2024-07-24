unit TaskNewProc;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client,
  Data.DB, FireDAC.Comp.DataSet, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,System.StrUtils;

type
  TNewProcTask = class(TForm)
    GroupBox1: TGroupBox;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    OKBtn: TButton;
    Button2: TButton;
    EditFileRun: TLabeledEdit;
    ComboBox2: TComboBox;
    GroupBox2: TGroupBox;
    CheckBox4: TCheckBox;
    EditCopyPath: TLabeledEdit;
    ComboBox1: TComboBox;
    CheckBox5: TCheckBox;
    Button1: TButton;
    EditSource: TLabeledEdit;
    CheckBox2: TCheckBox;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    ComboBox3: TComboBox;
    Label1: TLabel;
    SpeedButton5: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    function readproc(item:integer):boolean;
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure OKBtnClick(Sender: TObject);
    procedure ComboBox3Select(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  NewProcTask: TNewProcTask;


implementation
uses MyDM,umain,TaskEdit,EditProcMsi;
{$R *.dfm}


function TNewProcTask.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''proc'''
+' ORDER BY DESCRIPTION_PROC';;
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
if ComboBox2.Items.Count>0 then
begin
ComboBox2.ItemIndex:=item;
ComboBox3.ItemIndex:=item;
ComboBox2.OnSelect(ComboBox2);
end;

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


procedure TNewProcTask.Button1Click(Sender: TObject);
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





procedure TNewProcTask.Button2Click(Sender: TObject);
begin
Close;
end;

procedure TNewProcTask.CheckBox2Click(Sender: TObject);
begin
GroupBox2.Enabled:=CheckBox2.Checked;
if CheckBox2.Checked then
begin
EditSource.Text:=DeliteKeyFilePatch(EditFileRun.Text);
GroupBox2.Height:=160;
end
else GroupBox2.Height:=25;
end;

procedure TNewProcTask.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
combobox2.DroppedDown:=true;
end;

procedure TNewProcTask.ComboBox2Select(Sender: TObject);
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


procedure TNewProcTask.ComboBox3Select(Sender: TObject);
begin
ComboBox2.ItemIndex:=ComboBox3.ItemIndex;
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



procedure TNewProcTask.FormShow(Sender: TObject);
begin
CheckBox2.Checked:=false;
GroupBox2.Height:=25;
if not DataM.ConnectionDB.Connected then exit;
readproc(0);
end;




procedure TNewProcTask.OKBtnClick(Sender: TObject);
var
i,z:integer;
begin
if (ComboBox2.Text='') and (ComboBox3.Text='') then
begin
  ShowMessage('�������� ������� ��� ���������� � ������');
  exit;
end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='������ �������� :'+ComboBox2.text)
    and (EditTask.ListView1.Items[i].SubItems[1]='proc') then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with EditTask.ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=ComboBox3.Text; // ����� �������� ��� msi
  SubItems.Add('������ �������� :'+ComboBox2.Text);    // ��������
  SubItems.Add('proc');
end;
end;

procedure TNewProcTask.SpeedButton1Click(Sender: TObject); // ����� �������
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
CheckBox2.Checked:=false; // ����� ������� "���������� ����� ����������"
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
   Newdescription:=InputBox('���������� �������� � �����', '��������:', '����� �������� ��������');
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

procedure TNewProcTask.SpeedButton2Click(Sender: TObject); //�������
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
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ComboBox3.text+''''; // � ������ ���������� ��������� ID ��������� ��� MSI
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// ��������� ������  ��������� �� ����
readproc(ComboBox2.ItemIndex-1);
end;



procedure TNewProcTask.SpeedButton3Click(Sender: TObject);
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



procedure TNewProcTask.SpeedButton4Click(Sender: TObject);  //���������
var
descript:string;
begin
if (EditFileRun.Text='') then
begin
  ShowMessage('������� ���� � ����� ������� � �������� ��������');
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


procedure TNewProcTask.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='��������';
FormEditMsiProc.Caption:='�������� ���������';
FormEditMsiProc.ShowModal;
end;

end.
