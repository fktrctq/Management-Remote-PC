unit ErrorLic;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS,
  FireDAC.Phys.Intf, FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt,
  FireDAC.Comp.Client, Data.DB, FireDAC.Comp.DataSet, Vcl.Menus, Vcl.StdCtrls,
  Vcl.ExtCtrls,ShellApi;

type
  TCodeErrorLic = class(TForm)
    LVLic: TListView;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    procedure N2Click(Sender: TObject);
    procedure LVLicDblClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FormKey:Tform;
    EditKey:TlabeledEdit;
    EditNameKey:TlabeledEdit;
    ButtonOk,ButtonNo:Tbutton;
    Check:TcheckBox;
    function EditErrorCode(err,des,cap:string;act:boolean):boolean;
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonNoClose(Sender: TObject);
    procedure ButtonClOk(Sender: TObject);
    procedure updateEddodlist;
  public
    { Public declarations }
  end;

var
  CodeErrorLic: TCodeErrorLic;

implementation
uses mydm;

{$R *.dfm}

procedure TCodeErrorLic.updateEddodlist;
var
FDQuery:TFDQuery;
Transaction: TFDTransaction;
begin
Transaction:= TFDTransaction.Create(nil);
Transaction.Connection:=DataM.ConnectionDB;
Transaction.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
Transaction.Options.AutoCommit:=false;
Transaction.Options.AutoStart:=false;
Transaction.Options.AutoStop:=false;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=Transaction;
FDQuery.Connection:=DataM.ConnectionDB;
try
LVLic.Clear;
Transaction.StartTransaction; // стартуем
FDQuery.SQL.Text:='SELECT * FROM LIC_ERROR';
FDQuery.open;
FDQuery.First;
while not FDQuery.Eof do
begin
with LVLic.Items.add do
begin                          //
if FDQuery.FieldByName('ACTIVATE').aSboolean then
Caption:='Да' else Caption:='Нет';
SubItems.Add(FDQuery.FieldByName('CODE').AsString);
SubItems.Add(FDQuery.FieldByName('DESCRIPTION').AsString);
SubItems.Add(FDQuery.FieldByName('ID').AsString);
end;
FDQuery.Next;
end;
Transaction.Commit;
FDQuery.Close;
finally
Transaction.Free;
FDQuery.Free;
end;
end;


procedure TCodeErrorLic.LVLicDblClick(Sender: TObject);
var
b:boolean;
begin
if LVLic.SelCount<>1 then exit;
if LVLic.Selected.caption='Да' then b:=true
else b:=false;
EditErrorCode(LVLic.Selected.SubItems[0],LVLic.Selected.SubItems[1],'Редактировать',b);
end;

procedure TCodeErrorLic.N1Click(Sender: TObject);
begin
EditErrorCode('','Описание','Добавить',false);
end;

procedure TCodeErrorLic.N2Click(Sender: TObject);
var
b:boolean;
begin
if LVLic.SelCount<>1 then exit;
if LVLic.Selected.caption='Да' then b:=true
else b:=false;
EditErrorCode(LVLic.Selected.SubItems[0],LVLic.Selected.SubItems[1],'Редактировать',b);
end;

procedure TCodeErrorLic.N3Click(Sender: TObject);
var
FDQuery:TFDQuery;
Transaction: TFDTransaction;
i:integer;
begin
if LVLic.Items.Count=0 then exit;

Transaction:= TFDTransaction.Create(nil);
Transaction.Connection:=DataM.ConnectionDB;
Transaction.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
Transaction.Options.AutoCommit:=false;
Transaction.Options.AutoStart:=false;
Transaction.Options.AutoStop:=false;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=Transaction;
FDQuery.Connection:=DataM.ConnectionDB;
try
for I := LVLic.Items.Count-1 downto 0 do
if LVLic.Items[i].Selected then
begin
Transaction.StartTransaction; // стартуем
FDQuery.SQL.Text:='DELETE FROM  LIC_ERROR WHERE ID=:a';
FDQuery.ParamByName('a').asInteger:=strtoint(LVLic.items[i].subitems[2]);
FDQuery.ExecSQL;
Transaction.Commit;
LVLic.Items[i].Delete;
end;
finally
Transaction.Free;
FDQuery.Free;
end;
updateEddodlist;
end;

procedure TCodeErrorLic.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;

procedure TCodeErrorLic.FormShow(Sender: TObject);
begin
updateEddodlist;
end;

procedure TCodeErrorLic.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;



procedure TCodeErrorLic.ButtonClOk(Sender: TObject);
var
FDQuery:TFDQuery;
Transaction: TFDTransaction;
begin
if (EditNameKey.Text='') and (EditKey.Text='') then exit;
Transaction:= TFDTransaction.Create(nil);
Transaction.Connection:=DataM.ConnectionDB;
Transaction.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
Transaction.Options.AutoCommit:=false;
Transaction.Options.AutoStart:=false;
Transaction.Options.AutoStop:=false;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=Transaction;
FDQuery.Connection:=DataM.ConnectionDB;
try
Transaction.StartTransaction; // стартуем
FDQuery.SQL.Text:='update or insert into LIC_ERROR (CODE,DESCRIPTION,ACTIVATE) VALUES (:p1,:p2,:p3) MATCHING (CODE)';
FDQuery.params.ParamByName('p1').AsString:=EditNameKey.Text;
FDQuery.params.ParamByName('p2').AsString:=EditKey.Text;
FDQuery.params.ParamByName('p3').AsBoolean:=Check.checked;
FDQuery.ExecSQL;
Transaction.Commit;
finally
Transaction.Free;
FDQuery.Free;
end;
FormKey.Close;
updateEddodlist;
end;

function TCodeErrorLic.EditErrorCode(err,des,cap:string;act:boolean):boolean;
begin
try
FormKey:=TForm.Create(CodeErrorLic);
FormKey.Caption:=cap;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=365;
FormKey.Height:=170;
FormKey.OnClose:= FormKeyClose;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Код:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=340;
EditNameKey.Text:=err;
//////////////////////
EditKey:=TLabeledEdit.Create(FormKey);
EditKey.Parent:=FormKey;
EditKey.EditLabel.Caption:='Описание:';
EditKey.Top:=60;
EditKey.Left:=5;
EditKey.Width:=340;
if (des='Unknown') or (des='') then EditKey.Text:=' Выполните Windows+R slui.exe 0x2a "код ошибки, без кавычек"'
else EditKey.Text:=des;
EditKey.TabOrder:=0; // переводим фокус в поле со значением параметра
//////////////////
Check:=TcheckBox.create(FormKey);
Check.parent:=FormKey;
Check.caption:='Необходима прививка...';
Check.top:=85;
Check.left:=5;
Check.width:=200;
Check.checked:=act;
//////////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=100;
ButtonOk.Left:=170;
ButtonOk.OnClick:=ButtonClOk;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=100;
ButtonNo.Left:=270;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
except on E: Exception do showmessage('Ошибка '+e.Message);
end;
end;

end.
