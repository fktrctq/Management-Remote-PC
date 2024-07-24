unit PrintRenamePrintDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg1234567891011 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1234567891011: TOKRightDlg1234567891011;
  NewNamePrintThread:TThread;

implementation
uses umain, PrintRenamePrint;
{$R *.dfm}

procedure TOKRightDlg1234567891011.CancelBtnClick(Sender: TObject);
begin
Close;
end;

procedure TOKRightDlg1234567891011.FormShow(Sender: TObject);
begin
OKRightDlg1234567891011.Caption:=SelectedPrint;
end;

procedure TOKRightDlg1234567891011.OKBtnClick(Sender: TObject);
begin
NewNamePrint:=edit1.Text;
NewNamePrintThread:=PrintRenamePrint.PrintRenamePrinter.Create(true);
NewNamePrintThread.FreeOnTerminate:=true;
NewNamePrintThread.Start;
end;

end.
