unit PropertiesNetworkDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls;

type
  TOKRightDlg1234567891011121314 = class(TForm)
    CancelBtn: TButton;
    ListView1: TListView;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1234567891011121314: TOKRightDlg1234567891011121314;
  Newnetwork:TThread;
implementation
uses umain,PropertiesNetworkThread;
{$R *.dfm}

procedure TOKRightDlg1234567891011121314.FormShow(Sender: TObject);
begin
caption:=frmDomaininfo.listview7.Selected.SubItems[0];
Newnetwork:=PropertiesNetworkThread.PropertiesNetworkInterface.Create(true);
Newnetwork.FreeOnTerminate:=true;
Newnetwork.Start;
end;

end.
