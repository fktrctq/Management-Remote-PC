unit NewUserProfileDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,Vcl.dialogs;

type
  TOKRightDlg123456789101112 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    Label1: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg123456789101112: TOKRightDlg123456789101112;
  NewChangeOwner:Tthread;
  FlagsBool:byte;
implementation
uses umain, ChangeOwnerThread;

{$R *.dfm}

procedure TOKRightDlg123456789101112.CheckBox1Click(Sender: TObject);
begin
FlagsBool:=integer(CheckBox1.Checked);
end;

procedure TOKRightDlg123456789101112.FormShow(Sender: TObject);
begin
Caption:='Смена владельца профиля '+FrmDomainInfo.listview5.Selected.SubItems[1];
FlagsBool:=0;
labeledEdit1.Text:='';
end;

procedure TOKRightDlg123456789101112.OKBtnClick(Sender: TObject);
begin
NewUserName:=labeledEdit1.Text;
FlagsBool:=integer(CheckBox1.Checked);
if labeledEdit1.Text<>'' then
begin
NewChangeOwner:=ChangeOwnerThread.UserChangeOwner.Create(true);
NewChangeOwner.FreeOnTerminate:=true;
NewChangeOwner.Start;
end
else
begin
ShowMessage('Вы не ввели имя нового пользователя!');
exit;
end;
end;

end.
