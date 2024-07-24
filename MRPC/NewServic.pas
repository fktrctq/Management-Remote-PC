unit NewServic;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,Dialogs;

type
  TOKRightDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    ComboBox1: TComboBox;
    Label1: TLabel;
    ComboBox2: TComboBox;
    Label2: TLabel;
    CheckBox1: TCheckBox;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg: TOKRightDlg;
  MyNewService:TThread;
  NewServType :integer;
  NewServStartMode,NewServStartName,NewServStartPassword,NewServName,NewServDispName,NewServPathName:string;
  NewServDesktop:bool;

implementation
uses uMain,unit7;

{$R *.dfm}

procedure TOKRightDlg.CancelBtnClick(Sender: TObject);
begin
Close;
end;

procedure TOKRightDlg.OKBtnClick(Sender: TObject);
begin
NewServName:=LabeledEdit1.text;
NewServDispName:= LabeledEdit2.text;
NewServPathName:= LabeledEdit3.text;
case combobox1.itemindex of
0: NewServType:=1;
1: NewServType:=2;
2: NewServType:=4;
3: NewServType:=8;
4: NewServType:=16;
5: NewServType:=32;
6: NewServType:=256;
end;
NewServStartMode:=combobox2.text;
NewServDesktop:=CheckBox1.Checked;
NewServStartName:= LabeledEdit4.text;
NewServStartPassword:=LabeledEdit5.text;
//showmessage(NewServName+ NewServDispName+ NewServPathName+inttostr(NewServType) +inttostr(1) +NewServStartMode +booltostr(NewServDesktop)+NewServStartName+NewServStartPassword);
MyNewService:=unit7.NewService.Create(false);
Close;
end;

end.
