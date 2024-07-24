unit DlgUnjoinDomainOrWorkgroup;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,System.Variants,ActiveX,ComObj,Dialogs;

type
  TOKRightDlg12345678 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    EditUserAdmin: TLabeledEdit;
    EditPassAdmin: TLabeledEdit;
    RebootAfter: TCheckBox;
    EditNameGroup: TLabeledEdit;
    CheckBox1: TCheckBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    function UnjoinDomainOrWorkgroup(Mypc,User,Passwd,PasswordD,UserD
,NameWorkGroup:string;Reboot:bool):bool;
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678: TOKRightDlg12345678;


implementation
uses umain;
{$R *.dfm}
function RebootAfterunjoin(FWMIService   : OLEVariant):boolean;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown:integer;
begin
try
OleInitialize(nil);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
//FWMIService.Security_.impersonationlevel:=3;
//FWMIService.Security_.authenticationLevel := 6;
//FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   ��������� ��������� ��������� �������
oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then
begin
ResShutdown:=FWbemObject.Win32Shutdown(6);
frmDomainInfo.memo1.Lines.Add('������������ - '+SysErrorMessage(ResShutdown));
end;
FWbemObject:=Unassigned;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
//VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
 except
  on E:Exception do frmDomainInfo.memo1.Lines.Add('������ - "'+E.Message+'"');
 end;
end;


function TOKRightDlg12345678.UnjoinDomainOrWorkgroup(Mypc,User,Passwd,PasswordD,UserD
,NameWorkGroup:string;Reboot:bool):bool;
var
MyError:integer;
FWbemObjectSet,FSWbemLocator,FWMIService,FWbemObject: OLEVariant;
oEnum : IEnumvariant;
iValue : LongWord;
resstr:string;
begin
try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(Mypc, 'root\CIMV2', User, Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
   if NameWorkGroup<>'' then
     begin
     resstr:='';
      frmDomainInfo.memo1.Lines.Add('������� ������� ������ -'+string(FWbemObject.Workgroup));;
     FWbemObject.Workgroup:=NameWorkGroup; // ��� ������
     resstr:=FWbemObject.put_(); /// ��������� ���������� ����� ������
     if FWbemObject.Workgroup<>null then
     frmDomainInfo.memo1.Lines.Add('����� ������� ������ - '+ string(FWbemObject.Workgroup));
     if (resstr<>'')and(reboot) then  rebootAfterUnjoin(FWMIService);
     end
   else
     begin
      MyError:=FWbemObject.UnjoinDomainOrWorkgroup(PasswordD,UserD,0);    /// 0 - ������� �� ������, 4- ����� �� ������ � �������� ������ ����� � ������
      frmDomainInfo.memo1.Lines.Add('������� �� �� ������. '+SysErrorMessage(MyError));
      if FWbemObject.Workgroup<>null then
      frmDomainInfo.memo1.Lines.Add('������� ������ - '+string(FWbemObject.Workgroup)) ;
     end;
     if MyError=0 then
     begin
      if Reboot then rebootAfterUnjoin(FWMIService)//(Mypc,LocalAdm,LocalPass)
     else frmDomainInfo.memo1.Lines.Add('���������� ������������� ���������');
     end;
  frmDomainInfo.memo1.Lines.Add('---------------------------');
 end;

  FWbemObject:=Unassigned;
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;

except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(Mypc+': ������ "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
         end;

end;
end;

procedure TOKRightDlg12345678.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg12345678.CheckBox1Click(Sender: TObject);
var
b:integer;
begin
EditNameGroup.Enabled:= CheckBox1.Checked;
if not EditNameGroup.Enabled then EditNameGroup.Text:='';
if EditNameGroup.Enabled then b:=MessageDlg('��� ����� ������� ������ ���������� '+#10#13
+' ����� ��������� �� ��� ��������� � ������!' ,mtInformation, mbOKCancel, 0);
if b=mrCancel then CheckBox1.Checked:=false;

end;

procedure TOKRightDlg12345678.OKBtnClick(Sender: TObject);
var
LocalAdmin,LocalPass:string;
begin
UnjoinDomainOrWorkgroup(frmDomainInfo.ComboBox2.Text, // ����� ���� ���������������
                        frmDomainInfo.LabeledEdit1.Text,       //  �����
                        frmDomainInfo.LabeledEdit2.Text,       // ������
                        EditUserAdmin.Text,                        //�����
                        EditPassAdmin.Text,                         // ������
                        EditNameGroup.Text,                        // ������� ������
                        RebootAfter.Checked);               // ������������ ����� ��������������


end;

end.
