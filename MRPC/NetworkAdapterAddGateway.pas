unit NetworkAdapterAddGateway;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg123456789101112131415161718 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Label1: TLabel;
    Edit1: TEdit;
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg123456789101112131415161718: TOKRightDlg123456789101112131415161718;

implementation
uses networkAdapterAdditionallyDialog;
{$R *.dfm}

procedure TOKRightDlg123456789101112131415161718.OKBtnClick(Sender: TObject);
begin
if TransformGateway=false then
  begin     //// добавляем шлюз
    if edit1.Text<>'' then
      begin
       OKRightDlg12345678910111213141516.ListView2.Items.Add;
       OKRightDlg12345678910111213141516.ListView2.Items[OKRightDlg12345678910111213141516.ListView2.Items.Count-1].Caption:=Edit1.Text;
       close;
      end;
  end;
if TransformGateway=true then
   begin    //// изменяем шлюз
     if edit1.Text<>'' then
       begin
         OKRightDlg12345678910111213141516.ListView2.Items[OKRightDlg12345678910111213141516.Listview2.Selected.Index].Caption:=Edit1.Text;
          close;
       end;
   end;

end;

end.
