unit TaskSelectDelMSI;

interface

uses
  SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,dialogs,IdIcmpClient,
  System.Variants,ActiveX,ComObj,CommCtrl;

type
  TSelectDelMSITask = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Button1: TButton;
    ComboBox1: TComboBox;
    CheckBox2: TCheckBox;
    Edit1: TEdit;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    function LoadListProcess(NamePC:string):boolean;
    function LoadMSI(namePC:string):boolean;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SelectDelMSITask: TSelectDelMSITask;

implementation
uses umain,TaskEdit;
{$R *.dfm}


function ping(s:string):boolean;
var
z:integer;
MyIdIcmpClient: TIdIcmpClient;
begin
try
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=3000;
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
      begin
       result:=false;
      end
    else
      begin
      result:=true;
      end;
  if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
  except
  on E: Exception do
    begin
    if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
    result:=false;
    end;
  end;

end;


function TSelectDelMSITask.LoadListProcess(NamePC:string):boolean; /// ��������� ������ ���������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum           : IEnumvariant;
  iValue          : LongWord;

begin
OleInitialize(nil);
ComboBox1.Clear;
 frmDomainInfo.memo1.Lines.Add('�������� ������ ���������..');
 try
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Process WHERE ProcessId<>0','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ��������
 begin
 if FWbemObject.Caption<>null then
 ComboBox1.Items.add(string(FWbemObject.Caption));       // ��� ���
 FWbemObject:=Unassigned;
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 except
 on E:Exception do
 begin
 frmDomainInfo.memo1.Lines.Add('������ �������� ��������� "'+E.Message+'"');
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 end;
 end;
OleUnInitialize;
ComboBox1.text:=('�������� ������ ��������� ���������.');
end;




function TSelectDelMSITask.LoadMSI(namePC:string):boolean;
  var
   FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject       : OLEVariant;
  FOutParams      : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum         : IEnumvariant;
  Dt            : TDateTime;
  iValue        : LongWord;
begin
try
OleInitialize(nil);
ComboBox1.Clear;
ComboBox1.text:=('�������� ���������� �������� ��������...');
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
begin
if FWbemObject.Caption<> null then
combobox1.Items.add(string(FWbemObject.Caption));
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
except
on E:Exception do
begin
showmessage('������ �������� �������� - "'+E.Message+'"');
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
exit;
end;
end;
ComboBox1.text:=('�������� ������ �������� ���������.');
end;



procedure TSelectDelMSITask.Button1Click(Sender: TObject);
var
NamePSSearch:string;
begin
NamePSSearch:='localhost';
if not InputQuery('�������� ������ ����� ������ ��������� �����!!!', '��� ����������', NamePSSearch) then exit;
   if ping(NamePSSearch) then
   begin
    if SelectDelMSITask.Caption='�������� �������� msi' then LoadMSI(NamePSSearch);
    if SelectDelMSITask.Caption='���������� ��������' then LoadListProcess(NamePSSearch);
   end
   else showmessage('��������� �� ��������');
end;

procedure TSelectDelMSITask.CancelBtnClick(
  Sender: TObject);
begin
Close;
end;


procedure TSelectDelMSITask.CheckBox2Click(
  Sender: TObject);
begin
if CheckBox2.Checked=false then
showmessage('����� �����������!!!'+#10#13+'�������� ��� ��������� ��������������� ������� - '+Edit1.Text);

end;

procedure TSelectDelMSITask.ComboBox1Select(
  Sender: TObject);
begin
Edit1.Text:=ComboBox1.Text;
end;

function StrInDescription(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=br;
end;

procedure TSelectDelMSITask.OKBtnClick(
  Sender: TObject);
  var
  i,z:integer;
function bToSTR(b:boolean):string;
begin
  if b then result:='1'
  else result:='0';
end;
begin
if SelectDelMSITask.Caption='�������� �������� msi' then
Begin
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='�������� ��������� :'+Edit1.text) and (EditTask.ListView1.Items[i].SubItems[1]='DelMSI') then
    begin
       z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with EditTask.ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=bToSTR(CheckBox2.Checked); // ������ ���������� ����� ��� ���
  SubItems.Add('�������� ��������� :'+Edit1.text);    // ��� ���������
  SubItems.Add('DelMSI');            // ��� ��������
end;
End;

if SelectDelMSITask.Caption='���������� ��������' then
Begin
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='���������� �������� :'+Edit1.text) and (EditTask.ListView1.Items[i].SubItems[1]='KillProc') then
    begin
       z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with EditTask.ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=bToSTR(CheckBox2.Checked); // ������ ���������� ����� �������� ��� ���
  SubItems.Add('���������� �������� :'+Edit1.text);    // ��� ��������
  SubItems.Add('KillProc');            // ��� ��������
end;
End;
end;


end.
