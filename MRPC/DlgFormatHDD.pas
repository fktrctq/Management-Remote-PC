unit DlgFormatHDD;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,Dialogs;

type
  TOKRightDlg12345 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    ComboBox1: TComboBox;
    Label1: TLabel;
    CheckBox1: TCheckBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg12345: TOKRightDlg12345;
    NewHDDFormat  :Tthread;

implementation
uses umain,HDDformat;
{$R *.dfm}
///FormatClusterSize,,FormatFileSystem,FormatLabel,FormatQuickFormat,FormatEnableCompression

procedure TOKRightDlg12345.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg12345.OKBtnClick(Sender: TObject);
var
ButtonSel:byte;
begin
Buttonsel:=MessageDlg('Форматирование',mtError, mbOKCancel, 0);
if buttonsel=mrOk then
begin
FormatFileSystem:=combobox1.Text;
FormatQuickFormat:=CheckBox1.Checked;
FormatClusterSize:=strtoint(LabeledEdit1.Text);
FormatLabel:= LabeledEdit2.Text;
FormatEnableCompression:=false;
NewHDDFormat:=HDDFormat.FormatHDD.Create(false);
end;
if buttonsel=mrCancel then  close;

end;

end.
