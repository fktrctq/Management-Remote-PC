unit My_Proc;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.StdCtrls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client,System.strUtils, Vcl.Menus,ShellApi;

type
  TEditMyProc = class(TForm)
    ListView1: TListView;
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    PopupMyCommand: TPopupMenu;
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ListView1DblClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure MyCommandClick(Sender: TObject);
    procedure MyCommandClickListPC(Sender: TObject);
    procedure EditProcOpen(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FormKey:Tform;
    EditDescr:TLabelEdedit; //�������� ����� �������
    EditPath:TLabelEdedit;  // ���������� ����� �������
    EditArg:TlabelEdedit; // ��������� ������� �����
    CheckLogin:TcheckBox;    //������������ ������ ����� � ������
    ButtonOk:Tbutton;
    ButtonNo:Tbutton;
     FormGr:Tform;
     EditUser:Tedit;
     EditPass:Tedit;
     ButOk:Tbutton;
     ButClose:Tbutton;
     MyProcLogin,MyProcPass:string;
    function updateList:boolean;
    Function CreateEditForm(Descr,FileRun,OptionKey:string;NewOrEdit:boolean):boolean; // ����� ��� �������� � �������������� ��������
    procedure KeyFormClose(Sender: TObject; var Action: TCloseAction); // �������� ������������ ����� ��� ��������
    procedure ButtonNoClose(Sender: TObject);    //������ �������� �� ������������ �����
    procedure ButtonSave(Sender: TObject); // ���������� ������� ������
    procedure ButtonNew(Sender: TObject); // �������� ����� ������
    procedure DelItem; // �������� ������
    procedure CreateMenuMyProcess;  // ������������ �������� popupmenu
    function RunMyProcess(FilePath,Arg:string):Boolean;
    function creatformUserPass(S:string):boolean;
    procedure ButtonOkPass(Sender: TObject);
  public
    { Public declarations }
  end;

var
  EditMyProc: TEditMyProc;

implementation
uses  MyDM, uMain;
{$R *.dfm}

procedure TEditMyProc.ButtonOkPass(Sender: TObject);
begin
MyProcLogin:=EditUser.Text;
MyProcPass:=EditPass.Text;
FormGr.Close;
end;

function TEditMyProc.creatformUserPass(S:string):boolean; // ������� ����� ��� ����� ������������ � ������
begin
try
FormGr:=TForm.Create(EditMyProc);
FormGr.Name:='FormUserPassProc';
FormGr.Width:=250;
FormGr.Height:=130;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:='������������ � ������ '+s;
FormGr.OnClose:=KeyFormClose;
FormGr.Position:=poMainFormCenter;


EditUser:=TEdit.Create(FormGr);
EditUser.Parent:=FormGr;
EditUser.Name:='UserProc';
EditUser.Text:='';
EditUser.top:=0;
EditUser.Left:=5;
EditUser.Width:=225;
EditUser.TextHint:='��� ������������';
EditUser.TabOrder:=0;

EditPass:=TEdit.Create(FormGr);
EditPass.Parent:=FormGr;
EditPass.Name:='PasswdProc';
EditPass.Text:='';
EditPass.PasswordChar:=#7;
EditPass.TextHint:='������';
EditPass.top:=25;
EditPass.Left:=5;
EditPass.Width:=225;
EditPass.TabOrder:=1;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='AddTaskProc';
ButOk.Caption:='��';
ButOk.Top:=57;
ButOk.Left:=50;
ButOk.Width:=85;
ButOk.OnClick:= ButtonOkPass;
ButOk.TabOrder:=2;


ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGrProc';
ButClose.Caption:='�������';
ButClose.Top:=57;
ButClose.Left:=145;
ButClose.Width:=85;
ButClose.OnClick:=ButtonNoClose;
ButClose.TabOrder:=4;
FormGr.Showmodal;
except on E: Exception do
showmessage('������ �������� ����� '+e.Message);
end;
end;


procedure TEditMyProc.CreateMenuMyProcess;  // ������������ �������� popupmenu
var
QueryRead: TFDQuery;
TransactionRead:TFDtransaction;
MItem:TmenuItem;
begin
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
TransactionRead.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryRead:=TFDQuery.Create(nil);
QueryRead.Transaction:=TransactionRead;
QueryRead.Connection:=DataM.ConnectionDB;
try
TransactionRead.StartTransaction;
QueryRead.SQL.clear;
QueryRead.SQL.Text:='SELECT * FROM MY_PROCESS '
+' ORDER BY KEY'; //key ��� NAME_PROCESS
QueryRead.Open;
PopupMyCommand.Items.Clear; // ������ ���� � �������� �������
frmDomainInfo.PopupForListPC.Items.Items[1].Clear;
while not QueryRead.Eof do
begin
if QueryRead.FieldByName('KEY').Value<>null then
 begin
 /////////////////// ������ ����� ���� ��� ������� ����������
 MItem:=TMenuItem.Create(PopupMyCommand);
 MItem.Caption:=QueryRead.FieldByName('NAME_PROCESS').AsString;
 MItem.Hint:=QueryRead.FieldByName('PATH_FILE').AsString+'//**//'+QueryRead.FieldByName('ARG_FILE').AsString;///
 MItem.OnClick:=MyCommandClick;
 PopupMyCommand.Items.Add(MItem);
 /////////////////////////////// ������ ����� ���� ��� ������� ����������
 MItem:=TMenuItem.Create(frmDomainInfo.PopupForListPC);
 MItem.Caption:=QueryRead.FieldByName('NAME_PROCESS').AsString;
 MItem.Hint:=QueryRead.FieldByName('PATH_FILE').AsString+'//**//'+QueryRead.FieldByName('ARG_FILE').AsString;///
 MItem.OnClick:=MyCommandClickListPC;
 //PopupMyCommand2.Items.Add(MItem);
 frmDomainInfo.PopupForListPC.Items.Items[1].Add(MItem);
 //PopupMyCommand.Items.Add(NewItem(QueryRead.FieldByName('NAME_PROCESS').AsString,0,false,true,MyCommandClick,0,'MyCItem'+QueryRead.FieldByName('KEY').AsString));
 end;
  QueryRead.Next;
end;
/// � ����� ������� ���� �������� ������ ��� �������������� ����
 MItem:=TMenuItem.Create(PopupMyCommand);
 MItem.Caption:='��������/�������������';
 MItem.OnClick:=EditProcOpen;
 PopupMyCommand.Items.Add(MItem);

 MItem:=TMenuItem.Create(frmDomainInfo.PopupForListPC);
 MItem.Caption:='��������/�������������';
 MItem.OnClick:=EditProcOpen;
 frmDomainInfo.PopupForListPC.Items.Items[1].Add(MItem);

 ////////////////////////////////��� ����� ���� �� ������� ��������� ��������� ������� (submenu)  PopupMyCommand2



finally
QueryRead.Close;
TransactionRead.Commit;
TransactionRead.Free;
QueryRead.Free;
end;

except on E: Exception do
begin
showmessage('������ ������ ������ '+e.Message);
end;
end;
end;


function TEditMyProc.RunMyProcess(FilePath,Arg:string):Boolean;
begin
try
 shellAPI.ShellExecute
  (0,  //- ����� ������������� ����
  'Open',  //��������
  PChar(FilePath+#0#0),// ��������� �� ������, ������������ ��� ����� ��� �������� ��� ������, ��� ��� ����� ��� ��������  ������ ������ ����������� ������� ��������.
  PChar(Arg+#0#0), // ��������� �� ������ ���������� ����������� ����� ������ ������ ����������� ������� ��������.
  nil,   //�������� �� ������, ������������ ���������� �� ���������. ������ ������ ����������� ������� ��������.
  SW_SHOWNORMAL);    //���������� ��� ���� ����� ������������ ����� ���������.
  except
   showmessage('������ ��� �������� ����������')
  end;
end;


procedure TEditMyProc.MyCommandClickListPC(Sender: TObject);
///////////////////////////////////////////////////////////////////////////////
function DivideString(MyCommand,NameProcess,NamePC:string; var MyPath,MyArg:string):boolean;
begin
try
MyPath:=copy(MyCommand,1,pos('//**//',MyCommand)-1);
MyArg:=Copy(MyCommand,pos('//**//',MyCommand)+7,Length(MyCommand));

if pos('[NAMEPC]',MyPath)<>0 then // ���� � ������ ����� �����������  [NAMEPC] �� �������� �� �� ��� ���������
MyPath:=StringReplace(MyPath,'[NAMEPC]',NamePC,[rfReplaceAll, rfIgnoreCase]);
if pos('[LOGIN]',MyPath)<>0 then // ���� � ������ ����� �����������  [LOGIN] �� �������� ���������� ���� ��� ����� ������ � ������
  begin
  MyProcLogin:='';MyProcPass:=''; // ������ ����������
  creatformUserPass(NameProcess);  //������� ���� ��� ����� ������ � ������
  MyPath:=StringReplace(MyPath,'[LOGIN]',MyProcLogin,[rfReplaceAll, rfIgnoreCase]); //������ �� �����
  if pos('[PASS]',MyPath)<>0 then // ���� � ������ ����� �����������  [PASS] �� �������� �� �� ������
  MyPath:=StringReplace(MyPath,'[PASS]',MyProcPass,[rfReplaceAll, rfIgnoreCase]);
  end;


if pos('[NAMEPC]',MyArg)<>0 then // ���� � ������ ����� �����������  [NAMEPC] �� �������� �� �� ��� ���������
MyArg:=StringReplace(MyArg,'[NAMEPC]',NamePC,[rfReplaceAll, rfIgnoreCase]);
if pos('[LOGIN]',MyArg)<>0 then // ���� � ������ ����� �����������  [LOGIN] �� �������� ���������� ���� ��� ����� ������ � ������
  begin
  MyProcLogin:='';MyProcPass:=''; // ������ ����������
  creatformUserPass(NameProcess);  //������� ���� ��� ����� ������ � ������
  MyArg:=StringReplace(MyArg,'[LOGIN]',MyProcLogin,[rfReplaceAll, rfIgnoreCase]); //������ �� �����
  if pos('[PASS]',MyArg)<>0 then // ���� � ������ ����� �����������  [PASS] �� �������� �� �� ������
  MyArg:=StringReplace(MyArg,'[PASS]',MyProcPass,[rfReplaceAll, rfIgnoreCase]);
  end;
except on E: Exception do
showmessage('������ ������������ ������ '+e.Message);
end;
end;
////////////////////////////////////////////////////////////////////////////////////
var
FullCommand,ComPach,ComArg,NameProcess:string;
begin
try
if frmDomainInfo.ListView8.SelCount=1 then // ���� �������� ���� ������ � ���� �����������
begin
FullCommand:=(Sender as TMenuItem).Hint;
NameProcess:=(Sender as TMenuItem).Caption;
DivideString(FullCommand,NameProcess,frmDomainInfo.ListView8.Selected.SubItems[0],ComPach,ComArg);
//showmessage(ComPach+'*'+ComArg);
RunMyProcess(ComPach,ComArg);
end;
except on E: Exception do
showmessage('����� ������ ������� ���������� '+e.Message);
end;
end;

procedure TEditMyProc.MyCommandClick(Sender: TObject);
///////////////////////////////////////////////////////////////////////////////
function DivideString(MyCommand,NameProcess:string; var MyPath,MyArg:string):boolean;
begin
try
MyPath:=copy(MyCommand,1,pos('//**//',MyCommand)-1);
MyArg:=Copy(MyCommand,pos('//**//',MyCommand)+7,Length(MyCommand));

if pos('[NAMEPC]',MyPath)<>0 then // ���� � ������ ����� �����������  [NAMEPC] �� �������� �� �� ��� ���������
MyPath:=StringReplace(MyPath,'[NAMEPC]',frmdomaininfo.ComboBox2.Text,[rfReplaceAll, rfIgnoreCase]);
if pos('[LOGIN]',MyPath)<>0 then // ���� � ������ ����� �����������  [LOGIN] �� �������� ���������� ���� ��� ����� ������ � ������
  begin
  MyProcLogin:='';MyProcPass:=''; // ������ ����������
  creatformUserPass(NameProcess);  //������� ���� ��� ����� ������ � ������
  MyPath:=StringReplace(MyPath,'[LOGIN]',MyProcLogin,[rfReplaceAll, rfIgnoreCase]); //������ �� �����
  if pos('[PASS]',MyPath)<>0 then // ���� � ������ ����� �����������  [PASS] �� �������� �� �� ������
  MyPath:=StringReplace(MyPath,'[PASS]',MyProcPass,[rfReplaceAll, rfIgnoreCase]);
  end;


if pos('[NAMEPC]',MyArg)<>0 then // ���� � ������ ����� �����������  [NAMEPC] �� �������� �� �� ��� ���������
MyArg:=StringReplace(MyArg,'[NAMEPC]',frmdomaininfo.ComboBox2.Text,[rfReplaceAll, rfIgnoreCase]);
if pos('[LOGIN]',MyArg)<>0 then // ���� � ������ ����� �����������  [LOGIN] �� �������� ���������� ���� ��� ����� ������ � ������
  begin
  MyProcLogin:='';MyProcPass:=''; // ������ ����������
  creatformUserPass(NameProcess);  //������� ���� ��� ����� ������ � ������
  MyArg:=StringReplace(MyArg,'[LOGIN]',MyProcLogin,[rfReplaceAll, rfIgnoreCase]); //������ �� �����
  if pos('[PASS]',MyArg)<>0 then // ���� � ������ ����� �����������  [PASS] �� �������� �� �� ������
  MyArg:=StringReplace(MyArg,'[PASS]',MyProcPass,[rfReplaceAll, rfIgnoreCase]);
  end;
except on E: Exception do
showmessage('������ ������������ ������ '+e.Message);
end;
end;
////////////////////////////////////////////////////////////////////////////////////
var
FullCommand,ComPach,ComArg,NameProcess:string;
begin
try
FullCommand:=(Sender as TMenuItem).Hint;
NameProcess:=(Sender as TMenuItem).Caption;
DivideString(FullCommand,NameProcess,ComPach,ComArg);
//showmessage(ComPach+'*'+ComArg);
RunMyProcess(ComPach,ComArg);
  except
   showmessage('����� ������ ������� ����������')
  end;
end;

procedure TEditMyProc.EditProcOpen(Sender: TObject);
begin
EditMyProc.showmodal;
end;

procedure TEditMyProc.KeyFormClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;


procedure TEditMyProc.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).close;
end;



function TEditMyProc.updateList:boolean; // �������� ������ ����������
var
QueryRead: TFDQuery;
TransactionRead:TFDtransaction;
begin
try
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
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
QueryRead.SQL.Text:='SELECT * FROM MY_PROCESS '
+' ORDER BY KEY'; //key  NAME_PROCESS
QueryRead.Open;

while not QueryRead.Eof do
begin
if QueryRead.FieldByName('KEY').Value<>null then
 begin
 with ListView1.Items.Add do
  begin
   Caption:=QueryRead.FieldByName('KEY').AsString;
   SubItems.add(QueryRead.FieldByName('NAME_PROCESS').AsString);
   SubItems.add(QueryRead.FieldByName('PATH_FILE').AsString);
   SubItems.add(QueryRead.FieldByName('ARG_FILE').AsString);
  end;
 end;
  QueryRead.Next;
end;

finally
QueryRead.Close;
TransactionRead.Commit;
TransactionRead.Free;
QueryRead.Free;
result:=true;
end;
except on E: Exception do
begin
showmessage('������ ������ ������ '+e.Message);
result:=false;
end;
end;
end;

procedure TEditMyProc.DelItem; // �������� ������
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
QueryDel.SQL.Text:='DELETE from MY_PROCESS WHERE KEY='''+ListView1.Items[z].Caption+''''; // � Caption ��������� ID ���������
QueryDel.ExecSQL;
end;
except on E: Exception do
  Showmessage('������ ��������. '+e.Message);
end;
finally
QueryDel.Close;
TransactionDel.Commit;
TransactionDel.Free;
QueryDel.Free;
end;
updateList;   // ���������� ������
CreateMenuMyProcess; // ������� � �������� ����� ����
end;

procedure TEditMyProc.ButtonSave(Sender: TObject);  // ���������� ��������� ������� ������
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
if EditDescr.Text='' then
begin
showmessage('�� ������� ��� ��������');
exit;
end;
if EditPath.Text='' then
begin
showmessage('�� ������ ���� �������');
exit;
end;
if copy(EditArg.text,1,1)<>' ' then
begin
  if MessageDlg('� ������ ������ � ����������� ����������� ������!!!'+#10#13+' ���������� ���������� ��������???'
      , mtWarning,[mbYes,mbCancel],0) =mrcancel then exit;
end;
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
QueryWrite.SQL.Text:='update or insert into MY_PROCESS (PATH_FILE,ARG_FILE,NAME_PROCESS)'+
' VALUES (:p1,:p2,:p3) MATCHING(NAME_PROCESS)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,300)+''; //���� � ����� �������
QueryWrite.params.ParamByName('p2').AsString:=''+leftstr(EditArg.text,300)+'';      //��������� �������
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,150)+'';  /// ��������
//QueryWrite.params.ParamByName('p4').AsBoolean:=CheckLogin.Checked; // ������������ ������ ����� � ������
QueryWrite.ExecSQL;
except on E: Exception do
  Showmessage('������ ����������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateList; //���������� ������
CreateMenuMyProcess; // ������� � �������� ����� ����
end;

procedure TEditMyProc.ButtonNew(Sender: TObject); // �������� ����� ������
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
i,z:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('��� ����������� � ��');
 exit;
end;
 for I := 0 to ListView1.Items.Count-1 do
   begin
     if EditDescr.Text=ListView1.Items[i].SubItems[0] then
     begin
       ShowMessage('������ � ������ "'+EditDescr.Text+'" ��� ����������');
       exit;
     end;
   end;

if EditDescr.Text='' then
begin
showmessage('�� ������� ��� ��������');
exit;
end;
if EditPath.Text='' then
begin
showmessage('�� ������ ���� �������');
exit;
end;
if copy(EditArg.text,1,1)<>' ' then
begin
  if MessageDlg('� ������ ������ � ����������� ����������� ������!!!'+#10#13+' ���������� ���������� ��������???'
      , mtWarning,[mbYes,mbCancel],0) =mrcancel then exit;
end;

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
QueryWrite.SQL.Text:='INSERT INTO MY_PROCESS (PATH_FILE,ARG_FILE,NAME_PROCESS)'+
' VALUES (:p1,:p2,:p3)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,300)+''; //���� � ����� �������
QueryWrite.params.ParamByName('p2').AsString:=''+leftstr(EditArg.text,300)+'';      //��������� �������
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,150)+'';  /// ��������
//QueryWrite.params.ParamByName('p4').AsBoolean:=CheckLogin.Checked; // ������������ ������ ����� � ������
QueryWrite.ExecSQL;
except on E: Exception do
 ShowMessage('������ �������� ����� ������. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateList;   //���������� ������
CreateMenuMyProcess; // ������� � �������� ����� ����
FormKey.Close;  // ��������� �����
end;



Function TEditMyProc.CreateEditForm(Descr,FileRun,OptionKey:string;NewOrEdit:boolean):boolean; // �������� ����� ��� ��������������
var
step:integer;
begin
try
step:=1;
FormKey:=TForm.Create(EditMyProc);
if NewOrEdit then FormKey.Caption:='����� ����������'
else FormKey.Caption:='������������� ���������� '+descr;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=360;   //285
FormKey.Height:=190;
FormKey.OnClose:=KeyFormClose;
////////////
step:=2;
EditDescr:=TLabelEdedit.Create(FormKey);
EditDescr.Parent:=FormKey;
EditDescr.EditLabel.caption:='��������';
EditDescr.Text:=Descr;
EditDescr.Top:=15;
EditDescr.Left:=5;
EditDescr.Width:=335;
/// /////////////////////
step:=3;
EditPath:=TLabelEdedit.Create(FormKey);
EditPath.Parent:=FormKey;
EditPath.EditLabel.Caption:='���� �������';
EditPath.Text:=FileRun;
EditPath.Top:=55;
EditPath.Left:=5;
EditPath.Width:=335;//335
/////////////////////////
step:=4;
EditArg:=TLabeledEdit.Create(FormKey);
EditArg.Parent:=FormKey;
EditArg.EditLabel.Caption:='��������� � ���������� [NAMEPC] [LOGIN] [PASS]';
EditArg.Text:=OptionKey;
EditArg.Top:=95;
EditArg.Left:=5;
EditArg.Width:=335;
////////////////////////
{CheckLogin:=TcheckBox.Create(FormKey);
CheckLogin.Parent:=FormKey;
CheckLogin.Checked:=Fcopy;
CheckLogin.Top:=115;
CheckLogin.Left:=5;
CheckLogin.Width:=335;
CheckLogin.Caption:='������������ LOGIN � PASS';}
////////////////////////
step:=5;
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='���������';
ButtonOk.Top:=125;
ButtonOk.Left:=185;
if NewOrEdit then ButtonOk.OnClick:=ButtonNew  // ���� ������� �����
else ButtonOk.OnClick:=ButtonSave;  // ������ ������������� ������������
ButtonOk.TabOrder:=3;
//////////////////
step:=6;
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='�������';
ButtonNo.Top:=125;
ButtonNo.Left:=265;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
result:=true;
FormKey.ShowModal;
except on E: Exception do
begin
  Showmessage('������ �������� ������� (�'+inttostr(step)+') '+e.Message);
  result:=false;
end;
end;
end;


procedure TEditMyProc.Button1Click(Sender: TObject); // �������� ����� ����������
begin
if ListView1.SelCount=1 then CreateEditForm(ListView1.Selected.SubItems[0],ListView1.Selected.SubItems[1],ListView1.Selected.SubItems[2],true)
else CreateEditForm('','','',true);

end;

procedure TEditMyProc.Button2Click(Sender: TObject); //�������
begin
DelItem;
end;

procedure TEditMyProc.Button3Click(Sender: TObject); //��������
begin
updateList;
end;

procedure TEditMyProc.Button4Click(Sender: TObject);
begin
if ListView1.SelCount=1 then CreateEditForm(ListView1.Selected.SubItems[0],ListView1.Selected.SubItems[1],ListView1.Selected.SubItems[2],false);
end;

procedure TEditMyProc.FormCreate(Sender: TObject);
begin
CreateMenuMyProcess;
end;

procedure TEditMyProc.FormShow(Sender: TObject);
begin
updateList;
end;

procedure TEditMyProc.ListView1DblClick(Sender: TObject); // ��������������
begin
if ListView1.SelCount=1 then CreateEditForm(ListView1.Selected.SubItems[0],ListView1.Selected.SubItems[1],ListView1.Selected.SubItems[2],false);
end;

end.
