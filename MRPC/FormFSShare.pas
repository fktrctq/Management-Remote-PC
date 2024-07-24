unit FormFSShare;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Buttons, Vcl.StdCtrls, Vcl.ExtCtrls,ActiveX,ComObj;

type
  TFormShareFS = class(TForm)
    EditPath: TLabeledEdit;
    EditName: TLabeledEdit;
    EditDescr: TLabeledEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Label1: TLabel;
    CheckBox1: TCheckBox;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormShareFS: TFormShareFS;
implementation
uses umain,shareFS,FolderSecuretyEdit;

{$R *.dfm}

function CheckArrayForEmpty(s:Variant):bool;  /// �������� Variant ������� �� �������
begin
try
 if VarArrayHighBound(s, 1)<>0 then result:=true
 else result:=false;
except
   result:=false;
end;
end;


procedure clearvar;
begin
OleInitialize(nil);
 VarClear(SDACESACLArray);
 VarClear(SDACEDACLArray);
 VariantClear(FWMIServiceSC);
 VariantClear(FSWbemLocatorSC);
 VariantClear(FolderSecurityDescriptor);
 OleUnInitialize;
end;

procedure TFormShareFS.CheckBox1Click(Sender: TObject);
begin
SpeedButton2.Enabled:= not CheckBox1.Checked
end;

procedure TFormShareFS.FormClose(Sender: TObject; var Action: TCloseAction);
begin
clearvar;
end;

procedure TFormShareFS.FormShow(Sender: TObject);
begin
EditPath.Text:='';
Editname.Text:='';
EditDescr.Text:='';
CheckBox1.Checked:=true;
SpeedButton2.Enabled:=false;
end;

procedure TFormShareFS.SpeedButton1Click(Sender: TObject);
begin
if (EditPath.Text='') or (Editname.Text='')or (EditDescr.Text='') then
begin
  ShowMessage('�� ��� ���� ���������');
  exit;
end;
try
//if CheckArrayForEmpty(SDACEDACLArray) then /// ���� ������ � ������� ������� �� ������ �� ��������� �������
if (not  CheckBox1.Checked) and (CheckArrayForEmpty(SDACEDACLArray)) then   //// ���� ������� ������ b ������ ������������� �� ������
begin
if AddShareFS(
frmDomainInfo.ComboBox2.Text /// ��� �����
,frmDomainInfo.LabeledEdit1.text ///������������
,frmDomainInfo.LabeledEdit2.text  /// ������
,EditPath.Text  // ���� �� ����� ������� �����������
,Editname.Text  // ��� �������� �������� �������
,EditDescr.Text  //�������� �������� �������� �������
,''              /// ������ ������
,0               /// ��� �������� �������
,FolderSecurityDescriptor // ���������� ������������
) then frmDomainInfo.LoadNetworkShare;
end
else  /// ����� ������� ���� �� ��������� ��� ��� ������
begin
 if  AddShareNoFS(
  frmDomainInfo.ComboBox2.Text /// ��� �����
,frmDomainInfo.LabeledEdit1.text ///������������
,frmDomainInfo.LabeledEdit2.text  /// ������
,EditPath.Text  // ���� �� ����� ������� �����������
,Editname.Text  // ��� �������� �������� �������
,EditDescr.Text  //�������� �������� �������� �������
  ) then frmDomainInfo.LoadNetworkShare; // �������� ������� ��������
end;
except
on E:Exception do
         begin
         frmDomainInfo.memo1.Lines.Add(MyPS+'������ �������� �������� ������� : "'+E.Message+'"');
         close;
         end;
end;
close;
end;


procedure TFormShareFS.SpeedButton2Click(Sender: TObject);
begin
OleInitialize(nil);
FSWbemLocatorSC := CreateOleObject('WbemScripting.SWbemLocator');
FWMIServiceSC  := FSWbemLocatorSC.ConnectServer(frmDomainInfo.ComboBox2.Text, 'root\CIMV2',frmDomainInfo.LabeledEdit1.text, frmDomainInfo.LabeledEdit2.text);
FWMIServiceSC.Security_.impersonationlevel:=3;
FWMIServiceSC.Security_.authenticationLevel := 6;
FrmSecFolder.ShowModal;
end;

procedure TFormShareFS.SpeedButton3Click(Sender: TObject);
begin
frmDomainInfo.Memo1.Lines.Add('������ ������� - '+inttostr(VarArrayHighBound(SDACEDACLArray, 1)));;
end;

end.
