unit NetworkAdapterAddDnsWins;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg12345678910111213141516171819 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678910111213141516171819: TOKRightDlg12345678910111213141516171819;

implementation
uses NetworkAdapterAdditionallyDialog;

{$R *.dfm}

procedure TOKRightDlg12345678910111213141516171819.OKBtnClick(Sender: TObject);
begin
if TransformDNSWINS=false then
  begin   //////////// добавляем DNS или WINS
    if DNSWINS='DNS' then
        begin
          OKRightDlg12345678910111213141516.listBox1.Items.Add(edit1.Text);
        end;
    if DNSWINS='WINS' then
        begin
          OKRightDlg12345678910111213141516.listBox2.Items.Add(edit1.Text);
        end;
  end;
if TransformDNSWINS=true then
    begin   //////////// изменяем DNS или WINS
     if DNSWINS='DNS' then
        begin
         OKRightDlg12345678910111213141516.listBox1.Items[OKRightDlg12345678910111213141516.ListBox1.ItemIndex]:=Edit1.Text;
        end;
    if DNSWINS='WINS' then
        begin
          OKRightDlg12345678910111213141516.listBox2.Items[OKRightDlg12345678910111213141516.ListBox2.ItemIndex]:=Edit1.Text;
        end;
    end;

end;

end.
