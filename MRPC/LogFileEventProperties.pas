unit LogFileEventProperties;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtDlgs;

type
  TPropertiesFileEvent = class(TForm)
    Memo1: TMemo;
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PropertiesFileEvent: TPropertiesFileEvent;

implementation
uses EventWindows,umain;

{$R *.dfm}

procedure TPropertiesFileEvent.FormShow(Sender: TObject);
begin
Memo1.Lines:=LogsFileProperties(PropertiesFileEvent.Caption,groupbox1.Caption,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
end;

procedure TPropertiesFileEvent.SpeedButton1Click(Sender: TObject);
 var          /// backup ����� �����
selectedFile:string;
begin
if PromptForFileName(selectedFile,
'���� ������� (*.evtx)|*.evtx','','������������ ��������� �����', 'C:\',true)
  then
  begin
  selectedFile:=InputBox('������������ ����������� ��������� �����', '����', selectedFile+'.evtx');
  if (selectedFile<>'.evtx') then
    begin
    BackUpLogsFileEvent(PropertiesFileEvent.Caption,groupbox1.Caption,selectedFile,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
    end
    else ShowMessage('���� �� ����� ���� ������');
  end
  else
  begin
    exit;
  end;
end;

procedure TPropertiesFileEvent.SpeedButton2Click(Sender: TObject);
var
i:integer;
begin
i:=MessageBox(Self.Handle, PChar('��������� ������ ����� ��������?'+#10#13+' ������ ��� ���������� ��������.')
      , PChar('������� �������') ,MB_YESNOCANCEL+MB_ICONQUESTION);
 if i=IDCANCEL then exit;
 if i=IDYes then SpeedButton1.Click;
 ClearLogsFileEvent(PropertiesFileEvent.Caption,groupbox1.Caption,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
end;

end.
