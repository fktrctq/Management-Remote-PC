unit RegEditKeySave;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.ImgList, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Buttons, JvExExtCtrls, JvNetscapeSplitter,
  Inifiles, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.Comp.Client;

type
  TRegKeySave = class(TForm)
    Panel1: TPanel;
    LVSaveKey: TListView;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    MemoLog: TMemo;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    PopupEditKey: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    PopupImport: TPopupMenu;
    N6: TMenuItem;
    N7: TMenuItem;
    ImageList1: TImageList;
    PopupNewKey: TPopupMenu;
    REGSZ1: TMenuItem;
    REGEXPANDSZ1: TMenuItem;
    REGBINARY1: TMenuItem;
    REGDWORD1: TMenuItem;
    REGQWORD1: TMenuItem;
    REGMULTISZ1: TMenuItem;
    N8: TMenuItem;
    REGSZ2: TMenuItem;
    REGEXPANDSZ2: TMenuItem;
    REGBINARY2: TMenuItem;
    REGDWORD2: TMenuItem;
    REGQWORD2: TMenuItem;
    REGMULTISZ2: TMenuItem;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LVSaveKeyDblClick(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure REGSZ2Click(Sender: TObject);
    procedure REGEXPANDSZ2Click(Sender: TObject);
    procedure REGBINARY2Click(Sender: TObject);
    procedure REGDWORD2Click(Sender: TObject);
    procedure REGQWORD2Click(Sender: TObject);
    procedure REGMULTISZ2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
   FormKey:Tform;
   MemoKey:Tmemo;
   EditKey:TlabeledEdit;
   EditNameKey:TlabeledEdit;
   EditRootKey:TlabeledEdit;
   EditsSubKeyName:TlabeledEdit;
   ComboRootKey:TcomboBox;
   ComboTypeKey:TcomboBox;
   ButtonOk,ButtonNo:Tbutton;
   function ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;
   function IntegerInGetRooT(Node:Integer):String; // �������� root ������ � Integer, �������� ��� ���
   function StringInGetRooT(Node:string):integer; // �������� ��� �������, �������� �������� � Integer
   function StrTointTypeKey(TypeKey:string):integer; //�������� ��� ����� � string �������� ��� � integer
   function DescriptionImportKey:string;
   function DeleteSelectSaveKey(id:integer):boolean; // �������� ������ �� ������� REGEDIT_KEY
   function EditFormForKey(TypeKey,sValueName,sValue,KeyRoot,sSubKeyName:string;IDkey:integer):boolean; // �������������� ������������ ����� �������
   function NewKeySave(TypeKeyStr:string):boolean; // ����� ��� �������� ������ ����� � ������ �����������
   procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
   procedure ButtonNoClose(Sender: TObject);
   procedure FormKeyShowMULTI_SZ(sender:Tobject); // �������� ���������������� ���������
   procedure ButtonOkMULTI_SZ(sender:Tobject);  // ��������� ��������������� ��������
   procedure ButtonOkREG_SZ(sender:Tobject);  // ��������� ������ ���� ����������
   function AddFaforiteKey(hDefKey,sSubKeyName,sValueName,sValue:string;idkey:integer):boolean; // ���������� ���������� ������ � ������ �����������
   procedure UpdateListSaveKey; // ������ ������ ����������� ������
   procedure ButtonOkSaveNewKey(sebder:Tobject); // ������ �������� ������ �����
   function CreateSaveNewKey(description,hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):boolean; //������� �������� ������ ����� � ������ ���������

 public
    { Public declarations }
  end;

var
  RegKeySave: TRegKeySave;

implementation
uses
 ShellAPI,
  ActiveX,
  ComObj,
  MyDM,
  RegEditForm,TaskEdit;

{$R *.dfm}
function TRegKeySave.IntegerInGetRooT(Node:Integer):String; // �������� root ������ � Integer, �������� ��� ���
var
val:string;
begin
try
val:='';
if node=HKEY_CLASSES_ROOT then val:='HKEY_CLASSES_ROOT';
if node=HKEY_CURRENT_USER then val:= 'HKEY_CURRENT_USER';
if node=HKEY_LOCAL_MACHINE then val:= 'HKEY_LOCAL_MACHINE';
if node=HKEY_USERS then val:= 'HKEY_USERS';
if node=HKEY_CURRENT_CONFIG then val:= 'HKEY_CURRENT_CONFIG';
result:=val;
except on E:EOleException do
        MemoLog.Lines.Add(Format('������ EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        MemoLog.Lines.Add('������ (1) '+E.Message);  end;
end;
end;

procedure TRegKeySave.LVSaveKeyDblClick(Sender: TObject);  // ������������� ����
begin
if LVSaveKey.SelCount<>1 then exit;
EditFormForKey(LVSaveKey.Selected.SubItems[3], //��� �����
               LVSaveKey.Selected.SubItems[2], // ��� �����
               LVSaveKey.Selected.SubItems[4], // �������� �����
               LVSaveKey.Selected.SubItems[5], // �������� ������
               LVSaveKey.Selected.SubItems[6], // ����
               strtoint(LVSaveKey.Selected.Caption));
end;



procedure TRegKeySave.N4Click(Sender: TObject); //������ ����� � ������ �������� � ������ �����������
var
i:integer;
begin
for I := 0 to LVSaveKey.Items.Count-1 do
 begin
 if LVSaveKey.Items[i].Selected then
 if ImportSaveKeyhDefKey(LVSaveKey.Items[i].SubItems[5], // hDefKey
                      LVSaveKey.Items[i].SubItems[6], // sSubKeyName
                      LVSaveKey.Items[i].SubItems[2], // sValueName
                      LVSaveKey.Items[i].SubItems[4], // sValue
                      LVSaveKey.Items[i].SubItems[3]) // TypeKey
 =0 then MemoLog.Lines.Add('���� ������� '+LVSaveKey.Items[i].SubItems[2]+' ������������ � ������ '+LVSaveKey.Items[i].SubItems[6]+' �� ��������� '+RegEdit.ComboBox2.Text)
 end;
end;

procedure TRegKeySave.N5Click(Sender: TObject); // ������ ����� � ������� ������ �������
var
i:integer;
hDefKey,sSubKeyName:string;
begin
if LVSaveKey.SelCount=0 then exit;
try
hDefKey:=RegEdit.EditRooT.Text;
sSubKeyName:=RegEdit.EditPath.Text;
if StringInGetRooT(hDefKey)=0 then
begin
  showmessage('�� ����� ������� ��� ��������� ������� � �������');
  exit;
end;
if sSubKeyName='' then
begin
  showmessage('�� ����� ������� ��� ������� � �������');
  exit;
end;

for I := 0 to LVSaveKey.Items.Count-1 do
 begin
 if LVSaveKey.Items[i].Selected then
 if ImportSaveKeyhDefKey(hDefKey, // hDefKey
                      sSubKeyName, // sSubKeyName
                      LVSaveKey.Items[i].SubItems[2], // sValueName
                      LVSaveKey.Items[i].SubItems[4], // sValue
                      LVSaveKey.Items[i].SubItems[3]) // TypeKey
 =0 then MemoLog.Lines.Add('���� ������� '+LVSaveKey.Items[i].SubItems[2]+' ������������ � ������ '+sSubKeyName+' �� ��������� '+RegEdit.ComboBox2.Text)
 end;
except on E: Exception do MemoLog.Lines.Add('������ (2)'+E.Message);
end;
 end;



function TRegKeySave.StringInGetRooT(Node:string):integer; // �������� ��� �������, �������� �������� � Integer
var
val:integer;
begin
try
val:=0;
if node='HKEY_CLASSES_ROOT' then val:=HKEY_CLASSES_ROOT;
if node='HKEY_CURRENT_USER' then val:= HKEY_CURRENT_USER;
if node='HKEY_LOCAL_MACHINE' then val:= HKEY_LOCAL_MACHINE;
if node='HKEY_USERS' then val:= HKEY_USERS;
if node='HKEY_CURRENT_CONFIG' then val:= HKEY_CURRENT_CONFIG;
result:=val;
except on E:EOleException do
        MemoLog.Lines.Add(Format('������ EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        MemoLog.Lines.Add('������ (3):'+E.Message);  end;
end;
end;

function TRegKeySave.StrTointTypeKey(TypeKey:string):integer; //�������� ��� ����� � string �������� ��� � integer
begin
if TypeKey='REG_SZ' then result:=1;
if TypeKey='REG_EXPAND_SZ' then result:=2;
if TypeKey='REG_BINARY' then result:=3;
if TypeKey='REG_DWORD' then result:=4;
if TypeKey='REG_MULTI_SZ' then result:=7;
if TypeKey='REG_QWORD' then result:=11;
end;

function TRegKeySave.DeleteSelectSaveKey(id:integer):boolean; // �������� ������ �� ������� REGEDIT_KEY
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
begin
try
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
TransactionAdd.Options.ReadOnly:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='DELETE FROM REGEDIT_KEY WHERE ID_KEY=:a';
FDQueryAdd.ParamByName('a').asInteger:=ID;
FDQueryAdd.ExecSQL;
TransactionAdd.Commit;
finally
FDQueryAdd.Free;
TransactionAdd.Free;
end;
result:=true;
except on E: Exception do begin MemoLog.Lines.Add('������ (4):'+E.Message); result:=false; end;
end;
end;


procedure TRegKeySave.Button2Click(Sender: TObject);
var
i:integer;
begin
if LVSaveKey.SelCount=0 then exit;
i:=MessageDlg('������� ���������� �����?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
 for I := LVSaveKey.Items.Count-1 DownTo 0 do
   begin
     if LVSaveKey.Items[i].Selected then
     if DeleteSelectSaveKey(strtoint(LVSaveKey.Items[i].Caption)) then LVSaveKey.Items.Delete(i);
   end;
end;

procedure TRegKeySave.Button3Click(Sender: TObject);
begin
close;
end;

procedure TRegKeySave.Button4Click(Sender: TObject); // �������� �������� ����� � ������
var
i,z,x:integer;
begin
try
if LVSaveKey.SelCount=0 then exit; // ���� ��� ���������� �� �������
for x := 0 to LVSaveKey.Items.Count-1 do  // ���� �� ���������� ������
if LVSaveKey.Items[x].Selected then       // ���� ���� �������
if EditTask.ListView1.Items.Count<30 then // ���� ������� � ������ ������ ��� 30 �� ��������� � ������
Begin
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='������� ���� ������� :'+LVSaveKey.items[x].SubItems[2] then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:=LVSaveKey.items[x].Caption; // ID ����� � ���� ������
    SubItems.Add('������� ���� ������� :'+LVSaveKey.items[x].SubItems[6]+':'+LVSaveKey.items[x].SubItems[2]);    // ��������
    SubItems.Add('CreateKey'); //�������� ����� �������
  end;
MemoLog.Lines.Add('���� '+LVSaveKey.items[x].SubItems[2]+' �������� � ������');
End;
RegKeySave.Close; //
except on E: Exception do MemoLog.Lines.Add('������ (11): '+E.Message);
end;
end;

function TRegKeySave.DescriptionImportKey:string;
var
Description:string;
begin
while Description='' do
Begin
Description:=DateTimeToStr(now);
if not InputQuery('������� �������� ��� ����������� ������', '��������', Description)
then begin Description:=''; exit; end;
end;
result:=Description;
end;



procedure TRegKeySave.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LVSaveKey.Clear;
MemoLog.Clear;
end;

procedure TRegKeySave.UpdateListSaveKey;
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
function TypeKeyImage(s:string):integer;
var
i:integer;
begin
i:=0;
if s='REG_SZ' then i:=1;
if s='REG_EXPAND_SZ' then i:=2;
if s='REG_BINARY' then i:=3;
if s='REG_DWORD' then i:=4;
if s='REG_MULTI_SZ' then i:=7;
if s='REG_QWORD' then i:=11;
result:=i;
end;
begin
try
try
if not datam.TableExists('REGEDIT_KEY') then exit; // ���� ������� ��� �� ������� �� �������
LVSaveKey.Clear;
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
TransactionAdd.Options.ReadOnly:=true;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='select * from REGEDIT_KEY';
FDQueryAdd.Open;
FDQueryAdd.First; // ������� � ������ ������
while not FDQueryAdd.Eof do
 begin
 with LVSaveKey.Items.Add do // ��������� ������
  begin
  ImageIndex:=TypeKeyImage(FDQueryAdd.FieldByName('TypeKey').AsString);
  Caption:=FDQueryAdd.FieldByName('ID_KEY').AsString;
  SubItems.add(FDQueryAdd.FieldByName('Description_key').AsString);
  SubItems.add(FDQueryAdd.FieldByName('NamePC').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sValueName').AsString);
  SubItems.add(FDQueryAdd.FieldByName('TypeKey').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sValue').AsString);
  SubItems.add(FDQueryAdd.FieldByName('hDefKey').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sSubKeyName').AsString);
  end;
 FDQueryAdd.Next; // ������� � ��������� ������
 end;
TransactionAdd.Commit;
finally
FDQueryAdd.Free;
TransactionAdd.Free;
end;
except on E: Exception do MemoLog.Lines.Add('������ (5): '+E.Message);
end;
end;

procedure TRegKeySave.FormShow(Sender: TObject);
begin
if Datam.�reateTablForRegEdit then UpdateListSaveKey; //���� ������� ���� ��� �������� �� �������� �� ��������� ������ ����������� ������
end;

function TRegKeySave.ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;
 /// ������ ������ � ������ �� ������ �����������
var
i,res,IntTypeKey:integer;
listT:TstringList;
function StringToList(sValue:string):Tstringlist; // �� ������ ������ ������ �����
var
s:string;
begin
result:=TStringList.Create;
result.Text:=sValue;
end;
function StrQDWord(s:string):string; // �� ������ �� ��������� HEX � Integer �������� ������ Int
begin
delete(s,1,pos('(',s));
result:=Copy(s,1,pos(')',s)-1);
end;
begin
try
IntTypeKey:=StrTointTypeKey(TypeKey);
case IntTypeKey of
   1:res:=RegEdit.CreateSetStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   2:res:=RegEdit.CreateSetExpandedStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   3:res:=RegEdit.CreateBinarySaveKey(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue,'HEX');
   4:res:=RegEdit.CreateSetDWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint(StrQDWord(sValue)));
   7:begin
    try
     listT:=TStringList.Create;
     listT:=StringToList(sValue);
     res:=RegEdit.CreateSetMultiStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,StringToList(sValue));
     finally
     listT.Free;
     end;
     end;
   11:res:=RegEdit.CreateSetQWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint64(StrQDWord(sValue)));
end;
if res<>0 then memolog.lines.add('������ ����� '+sValueName+': '+SysErrorMessage(res));
result:=res;
except on E: Exception do begin MemoLog.Lines.Add('������ (6): '+E.Message); result:=res; end;
end;
end;

procedure TRegKeySave.Button1Click(Sender: TObject);
var
i:integer;
begin
for I := 0 to LVSaveKey.Items.Count-1 do
 begin
 if LVSaveKey.Items[i].Selected then
 if ImportSaveKeyhDefKey(LVSaveKey.Items[i].SubItems[5], // hDefKey
                      LVSaveKey.Items[i].SubItems[6], // sSubKeyName
                      LVSaveKey.Items[i].SubItems[2], // sValueName
                      LVSaveKey.Items[i].SubItems[4], // sValue
                      LVSaveKey.Items[i].SubItems[3]) // TypeKey
 =0 then MemoLog.Lines.Add('���� ������� '+LVSaveKey.Items[i].SubItems[2]+' ������������ � ������ '+LVSaveKey.Items[i].SubItems[6]+' �� ���������� '+LVSaveKey.Items[i].SubItems[1])
 end;
end;


procedure TRegKeySave.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;

procedure TRegKeySave.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;

procedure TRegKeySave.FormKeyShowMULTI_SZ(sender:Tobject); // ������� ��������������� ��������
var
resSrtList:TstringList;
i:integer;
function StringToList(sValue:string):Tstringlist; // �� ������ ������ ������ �����
var
s:string;
begin
result:=TStringList.Create;
result.Text:=sValue;
end;
begin
  resSrtList:= Tstringlist.Create;
try
MemoKey.Clear;
resSrtList:=StringToList(LVSaveKey.Selected.SubItems[4]);
for I := 0 to resSrtList.Count-1 do MemoKey.Lines.Add(resSrtList[i]);
finally
 resSrtList.Free;
end;
end;

function TRegKeySave.CreateSaveNewKey(description,hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):boolean;
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
i:integer;
begin
try
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='INSERT INTO REGEDIT_KEY (Description_key, hDefKey ,sSubKeyName, sValueName, sValue, TypeKey, NamePC) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7)';//'UPDATE REGEDIT_KEY (hDefKey ,sSubKeyName, sValueName, sValue) VALUES (:p1,:p2,:p3,:p4) WHERE ID_KEY=:p5';
FDQueryAdd.params.ParamByName('p1').AsString:=''+Description+'';
FDQueryAdd.params.ParamByName('p2').AsString:=''+hDefKey+'';
FDQueryAdd.params.ParamByName('p3').AsString:=''+sSubKeyName+'';
FDQueryAdd.params.ParamByName('p4').AsString:=''+sValueName+'';
FDQueryAdd.params.ParamByName('p5').AsString:=''+sValue+'';
FDQueryAdd.params.ParamByName('p6').AsString:=''+TypeKey+'';
FDQueryAdd.params.ParamByName('p7').AsString:='����� ����';
FDQueryAdd.ExecSQL;
TransactionAdd.Commit;
finally
FDQueryAdd.free;
TransactionAdd.Free;
end;
result:=true;
except on E: Exception do
begin
  MemoLog.Lines.Add('������ �������� ����� (7) '+sValueName+' � ���������'+e.Message);
  result:=false;
end;
end;
end;

procedure TRegKeySave.ButtonOkSaveNewKey(sebder:Tobject); //������ �������� ������ �����
var
des,sValue:string;
i:integer;
q:int64;
begin
des:=DescriptionImportKey; // �������� ��� ����������� ������
if des='' then exit; // ���� �������� ����� �� �������

if ComboTypeKey.Text='REG_MULTI_SZ' then sValue:=MemoKey.Text
else
if ComboTypeKey.Text='REG_DWORD' then
begin
if TryStrToInt(EditKey.Text,i) then
sValue:=inttohex(i,8)+' ('+EditKey.Text+')'
else begin ShowMessage('�� ������ �������� �����'); EditKey.SetFocus; exit; end;
end
else
if ComboTypeKey.Text='REG_QWORD' then
begin
if TryStrToInt64(EditKey.Text,q) then
sValue:=inttohex(q,8)+' ('+EditKey.Text+')'
else begin ShowMessage('�� ������ �������� �����'); EditKey.SetFocus; exit; end;
end
else  sValue:= EditKey.Text;

if CreateSaveNewKey(des,             //��������
                ComboRootKey.Text, // root ������
                EditsSubKeyName.Text, // ���� � �������
                EditNameKey.Text,    // ��� �����
                sValue,
                ComboTypeKey.Text)
then
begin
MemoLog.Lines.Add('���� '+EditNameKey.Text+' ������� ��������');
UpdateListSaveKey; // ���� ���� ������, ��������� ������
FormKey.Close;
end
else
begin
MemoLog.Lines.Add('������ ���������� �����');
end;

end;


function TRegKeySave.AddFaforiteKey(hDefKey,sSubKeyName,sValueName,sValue:string;idkey:integer):boolean; // ���������� ���������� ������ � ������ �����������
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
begin
try
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='UPDATE REGEDIT_KEY SET hDefKey=:p1 ,sSubKeyName=:p2, sValueName=:p3, sValue=:p4 WHERE ID_KEY=:p5';//'UPDATE REGEDIT_KEY (hDefKey ,sSubKeyName, sValueName, sValue) VALUES (:p1,:p2,:p3,:p4) WHERE ID_KEY=:p5';
FDQueryAdd.params.ParamByName('p1').AsString:=''+hDefKey+'';
FDQueryAdd.params.ParamByName('p2').AsString:=''+sSubKeyName+'';
FDQueryAdd.params.ParamByName('p3').AsString:=''+sValueName+'';
FDQueryAdd.params.ParamByName('p4').AsString:=''+sValue+'';
FDQueryAdd.params.ParamByName('p5').AsInteger:=idkey;
FDQueryAdd.ExecSQL;
TransactionAdd.Commit;
finally
FDQueryAdd.free;
TransactionAdd.Free;
end;
result:=true;
except on E: Exception do
begin
  MemoLog.Lines.Add('������ ���������� ����� (8)'+sValueName+' � ���������'+e.Message);
  result:=false;
end;
end;
end;

procedure TRegKeySave.ButtonOkMULTI_SZ(sender:Tobject); // ���������� ��������� � ����������������� �����
Var
i:integer;
value:string;
begin
value:='';
for i :=0 to MemoKey.Lines.Count-1 do value:=value+MemoKey.Lines[i]+#13#10;
EditRootKey.Font.Color:=clBlack;
if StringInGetRooT(EditRootKey.Text)=0 then
begin
  ShowMessage('�� ������ �������� ��������� �������');
  EditRootKey.SetFocus;
  EditRootKey.Font.Color:=clred;
  exit;
end;
EditsSubKeyName.Font.Color:=clBlack;
if (EditsSubKeyName.Text='') or (pos('/',EditsSubKeyName.Text)<>0) then
begin
ShowMessage('�� ������ �������� �������');
  EditsSubKeyName.SetFocus;
  EditsSubKeyName.Font.Color:=clred;
  exit;
end;
EditNameKey.Font.Color:=clBlack;
if (EditNameKey.Text='') then
begin
  ShowMessage('�� ������ �������� ����� �����');
  EditNameKey.SetFocus;
  EditNameKey.Font.Color:=clred;
  exit;
end;
MemoKey.Font.Color:=clBlack;
if (value='') or (value=' ') or (value=#13#10) then
begin
  ShowMessage('�� ������ �������� ��������� �����');
  MemoKey.SetFocus;
  MemoKey.Font.Color:=clred;
  exit;
end;
if AddFaforiteKey(EditRootKey.Text,EditsSubKeyName.Text,EditNameKey.Text,value,FormKey.Tag) then
begin
UpdateListSaveKey;
MemoLog.Lines.add('�������� ���������� ��������� � ����� '+EditNameKey.Text);
FormKey.Close;
end;
end;

procedure TRegKeySave.ButtonOkREG_SZ(sender:Tobject);
var
value:string;
i:integer;
q:int64;
begin
EditRootKey.Font.Color:=clBlack;
if StringInGetRooT(EditRootKey.Text)=0 then
begin
  ShowMessage('�� ������ �������� ��������� �������');
  EditRootKey.SetFocus;
  EditRootKey.Font.Color:=clred;
  exit;
end;
EditsSubKeyName.Font.Color:=clBlack;
if (EditsSubKeyName.Text='') or (pos('/',EditsSubKeyName.Text)<>0) then
begin
ShowMessage('�� ������ �������� �������');
  EditsSubKeyName.SetFocus;
  EditsSubKeyName.Font.Color:=clred;
  exit;
end;
EditNameKey.Font.Color:=clBlack;
if (EditNameKey.Text='') then
begin
  ShowMessage('�� ������ �������� ����� �����');
  EditNameKey.SetFocus;
  EditNameKey.Font.Color:=clred;
  exit;
end;
EditKey.Font.Color:=clBlack;
if (EditKey.Text='') or (EditKey.Text=' ') then
begin
  ShowMessage('�� ������ �������� ��������� �����');
  EditKey.SetFocus;
  EditKey.Font.Color:=clred;
  exit;
end;
value:=EditKey.Text;
if (pos('REG_QWORD',FormKey.Caption)<>0) then
begin
if TryStrToInt64(value,q) then value:=(inttohex(q,8))+' ('+value+')'
else
begin
  ShowMessage('�� ������ �������� ��������� �����');
  exit;
end;
end;

if (pos('REG_DWORD',FormKey.Caption)<>0) then
begin
if TryStrToInt(value,i) then value:=(inttohex(q,8))+' ('+value+')'
else
begin
  ShowMessage('�� ������ �������� ��������� �����');
  exit;
end;
end;
if AddFaforiteKey(EditRootKey.Text,EditsSubKeyName.Text,EditNameKey.Text,value,FormKey.Tag) then
begin
UpdateListSaveKey;
MemoLog.Lines.add('�������� ���������� ��������� � ����� '+EditNameKey.Text);
FormKey.Close;
end;
end;


procedure TRegKeySave.REGBINARY2Click(Sender: TObject);
begin
NewKeySave('REG_BINARY');
end;
procedure TRegKeySave.REGDWORD2Click(Sender: TObject);
begin
NewKeySave('REG_DWORD');
end;
procedure TRegKeySave.REGEXPANDSZ2Click(Sender: TObject);
begin
NewKeySave('REG_EXPAND_SZ');
end;
procedure TRegKeySave.REGMULTISZ2Click(Sender: TObject);
begin
NewKeySave('REG_MULTI_SZ');
end;
procedure TRegKeySave.REGQWORD2Click(Sender: TObject);
begin
NewKeySave('REG_QWORD');
end;
procedure TRegKeySave.REGSZ2Click(Sender: TObject);
begin
NewKeySave('REG_SZ');
end;

function TRegKeySave.NewKeySave(TypeKeyStr:string):boolean;
begin
try
begin
if TypeKeyStr='REG_MULTI_SZ' then
begin
FormKey:=TForm.Create(RegKeySave);
FormKey.Caption:='�������� ���������������� ���������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=380;
FormKey.OnClose:=FormKeyClose;
////////////
ComboTypeKey:=TComboBox.Create(FormKey);
ComboTypeKey.Parent:=FormKey;
ComboTypeKey.Items.Add('REG_SZ');
ComboTypeKey.Items.Add('REG_EXPAND_SZ');
ComboTypeKey.Items.Add('REG_BINARY');
ComboTypeKey.Items.Add('REG_DWORD');
ComboTypeKey.Items.Add('REG_QWORD');
ComboTypeKey.Items.Add('REG_MULTI_SZ');
if TypeKeyStr='REG_MULTI_SZ' then ComboTypeKey.ItemIndex:=5;
ComboTypeKey.Style:=csOwnerDrawFixed;
ComboTypeKey.Top:=5;
ComboTypeKey.Left:=5;
ComboTypeKey.Width:=260;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=28;
ComboRootKey.Left:=5;
ComboRootKey.Width:=260;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������:';
EditsSubKeyName.Top:=67;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='��������:';
EditNameKey.Top:=105;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:='��� ����� �������';
EditNameKey.TabOrder:=1;
////////////////////////
MemoKey:=Tmemo.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Top:=130;
MemoKey.Left:=5;
MemoKey.Width:=260;
MemoKey.Height:=180;
MemoKey.Text:='';
MemoKey.ScrollBars:=ssBoth;
MemoKey.TabOrder:=2; // ��������� ����� � ���� �� ��������� ���������
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=315;
ButtonOk.Left:=110;
ButtonOk.OnClick:=ButtonOkSaveNewKey;
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=315;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
FormKey.ShowModal;
end
else
begin
FormKey:=TForm.Create(RegKeySave);
FormKey.Caption:='�������� ����� �������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=365;
FormKey.Height:=237;
FormKey.OnClose:= FormKeyClose;
///////////////////////////////
ComboTypeKey:=TComboBox.Create(FormKey);
ComboTypeKey.Parent:=FormKey;
ComboTypeKey.Items.Add('REG_SZ');
ComboTypeKey.Items.Add('REG_EXPAND_SZ');
ComboTypeKey.Items.Add('REG_BINARY');
ComboTypeKey.Items.Add('REG_DWORD');
ComboTypeKey.Items.Add('REG_QWORD');
ComboTypeKey.Items.Add('REG_MULTI_SZ');
if TypeKeyStr='REG_SZ' then ComboTypeKey.ItemIndex:=0;
if TypeKeyStr='REG_EXPAND_SZ' then ComboTypeKey.ItemIndex:=1;
if TypeKeyStr='REG_BINARY' then ComboTypeKey.ItemIndex:=2;
if TypeKeyStr='REG_DWORD' then ComboTypeKey.ItemIndex:=3;
if TypeKeyStr='REG_QWORD' then ComboTypeKey.ItemIndex:=4;
ComboTypeKey.Style:=csOwnerDrawFixed;
ComboTypeKey.Top:=5;
ComboTypeKey.Left:=5;
ComboTypeKey.Width:=340;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=28;
ComboRootKey.Left:=5;
ComboRootKey.Width:=340;
////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������:';
EditsSubKeyName.Top:=67;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=340;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='��������:';
EditNameKey.Top:=105;
EditNameKey.Left:=5;
EditNameKey.Width:=340;
EditNameKey.Text:='��� ����� �������';
EditNameKey.TabOrder:=1;
/////////////////////////
EditKey:=TLabeledEdit.Create(FormKey);
EditKey.Parent:=FormKey;
EditKey.EditLabel.Caption:='��������:';
EditKey.Top:=145;
EditKey.Left:=5;
EditKey.Width:=340;
EditKey.Text:='';
EditKey.TabOrder:=2;
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=169;
ButtonOk.Left:=180;
ButtonOk.OnClick:=ButtonOkSaveNewKey;
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=169;
ButtonNo.Left:=270;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
FormKey.ShowModal;
end;
end;
except on E: Exception do memolog.Lines.Add('��������� �������������� ������ (9) : '+e.Message);
end;
end;




function TRegKeySave.EditFormForKey(TypeKey,sValueName,sValue,KeyRoot,sSubKeyName:string;IDkey:integer):boolean;
begin     // ������� ����� ��� ��������� � ��������������� �������� ������ �������
try
begin
if TypeKey='REG_MULTI_SZ' then
begin
FormKey:=TForm.Create(RegKeySave);
FormKey.Caption:='��������� ���������������� ���������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=380;
FormKey.Tag:=IDkey;
FormKey.OnClose:=FormKeyClose;
FormKey.OnShow:=FormKeyShowMULTI_SZ; // ��� �������� ����� ���������� ������ �������� ���������
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='��������:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:=sValueName;
/// /////////////////////
EditRootKey:=TLabeledEdit.Create(FormKey);
EditRootKey.Parent:=FormKey;
EditRootKey.EditLabel.Caption:='root:';
EditRootKey.Top:=60;
EditRootKey.Left:=5;
EditRootKey.Width:=260;
EditRootKey.Text:=KeyRoot;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������:';
EditsSubKeyName.Top:=100;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:=sSubKeyName;
////////////////////////
MemoKey:=Tmemo.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Top:=125;
MemoKey.Left:=5;
MemoKey.Width:=260;
MemoKey.Height:=180;
MemoKey.Text:='';
MemoKey.ScrollBars:=ssBoth;
MemoKey.TabOrder:=0; // ��������� ����� � ���� �� ��������� ���������
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=310;
ButtonOk.Left:=95;
ButtonOk.OnClick:=ButtonOkMULTI_SZ;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=310;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end
else
begin
FormKey:=TForm.Create(RegKeySave);
FormKey.Caption:='��������� ������������ ����� ������� '+TypeKey;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=365;
FormKey.Height:=230;
FormKey.OnClose:= FormKeyClose;
FormKey.Tag:=IDkey;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='��������:';
EditNameKey.Top:=20;
EditNameKey.Left:=5;
EditNameKey.Width:=340;
EditNameKey.Text:=sValueName;
//////////////////////
EditRootKey:=TLabeledEdit.Create(FormKey);
EditRootKey.Parent:=FormKey;
EditRootKey.EditLabel.Caption:='root:';
EditRootKey.Top:=60;
EditRootKey.Left:=5;
EditRootKey.Width:=340;
EditRootKey.Text:=KeyRoot;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������:';
EditsSubKeyName.Top:=100;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=340;
EditsSubKeyName.Text:=sSubKeyName;
////////////////////////
///
EditKey:=TLabeledEdit.Create(FormKey);
EditKey.Parent:=FormKey;
EditKey.EditLabel.Caption:='��������:';
EditKey.Top:=140;
EditKey.Left:=5;
EditKey.Width:=340;
if (TypeKey='REG_DWORD') or (TypeKey='REG_QWORD') then
begin
  delete(sValue,1,pos('(',sValue));
  sValue:=copy(sValue,1,pos(')',sValue)-1);
end;
EditKey.Text:=sValue;    //inttohex(int64(OutParam.uValue),8)+' ('+vartostr(OutParam.uValue)+')'
EditKey.TabOrder:=0; // ��������� ����� � ���� �� ��������� ���������
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=165;
ButtonOk.Left:=170;
ButtonOk.OnClick:=ButtonOkREG_SZ;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=165;
ButtonNo.Left:=270;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end;
end;

except on E: Exception do memolog.Lines.Add('��������� �������������� ������ (10) : '+e.Message);
end;
end;

end.
