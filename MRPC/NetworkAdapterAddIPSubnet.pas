unit NetworkAdapterAddIPSubnet;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg1234567891011121314151617 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1234567891011121314151617: TOKRightDlg1234567891011121314151617;

implementation
uses NetworkAdapterAdditionallyDialog;
{$R *.dfm}

procedure TOKRightDlg1234567891011121314151617.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg1234567891011121314151617.OKBtnClick(Sender: TObject);
begin
if (Edit1.Text<>'')and(Edit2.Text<>'') then
  begin
  if transformIP=false then
  begin    //// добавляем ip адрес
  OKRightDlg12345678910111213141516.ListView1.Items.Add;
  OKRightDlg12345678910111213141516.ListView1.Items[OKRightDlg12345678910111213141516.ListView1.Items.Count-1].Caption:=Edit1.Text;
  OKRightDlg12345678910111213141516.ListView1.Items[OKRightDlg12345678910111213141516.ListView1.Items.Count-1].SubItems.add(Edit2.Text);
  close;
  end;
  if transformIP=true then
  begin      ///// изменяем ip адрес
  OKRightDlg12345678910111213141516.ListView1.Items[OKRightDlg12345678910111213141516.Listview1.Selected.Index].Caption:=Edit1.Text;
  OKRightDlg12345678910111213141516.ListView1.Items[OKRightDlg12345678910111213141516.Listview1.Selected.Index].SubItems[0]:=Edit2.Text;
  close;
  end;

  end;
end;

end.
