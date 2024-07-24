unit DlgChkdsk;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg123 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg123: TOKRightDlg123;
  ChkdskHDD:Tthread;
implementation
uses umain,HDDChkdsk;

{$R *.dfm}

procedure TOKRightDlg123.OKBtnClick(Sender: TObject);
begin
FixErrors:=CheckBox1.Checked;
VigorousIndexCheck:=CheckBox2.Checked;
SkipFolderCycle:=CheckBox3.Checked;
ForceDismount:=CheckBox4.Checked;
RecoverBadSectors:=CheckBox5.Checked;
OkToRunAtBootUp:=CheckBox6.Checked;
ChkdskHDD:=HDDChkdsk.Chkdsk.Create(true);
ChkdskHDD.FreeOnTerminate:=true;
ChkdskHDD.Start;
end;

end.
