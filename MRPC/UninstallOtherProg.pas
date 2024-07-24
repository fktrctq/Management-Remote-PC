unit UninstallOtherProg;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg12 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg12: TOKRightDlg12;
  UnProg:TThread;

implementation
uses umain,unit5;

{$R *.dfm}

procedure TOKRightDlg12.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg12.OKBtnClick(Sender: TObject);
begin
MyCommandLine:=LabeledEdit1.Text;
UnProg:=unit5.NewProcess.Create(false);
end;

end.
