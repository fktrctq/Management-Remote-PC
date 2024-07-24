unit NetworkAdapterDHCP;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterDHCPEnable = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NetworkConfiguration;


{ NetworkAdapterDHCPEnable }

procedure NetworkAdapterDHCPEnable.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError,MyErrorDNS       : integer;
  MasDNS        :OLEvariant;
  iValue        : LongWord;
const
IpError: array [64..100] of string =('����� �� �������������� �� ���� ���������',
'����������� ����','������������ ����� �������','��������� ������ ��� ��������� ���������� ������� ��� ���������'
,'������������ ������� ��������','������� ����� 5 ������','������������ IP-�����'
,'���������������� IP-����� �����','��������� ������ ��� ������� � ������� ��� ������������� ����������'
,'������������ ��� ������','�������� �����','��������� / ��������� ������ WINS �� ���������'
,'�������� ����','������������ ��������� ����','������ ����������� �����'
,'������������ �������� ������������','���������� ��������� ������ TCP / IP'
,'���������� ��������� ������ DHCP','�� ������� ����������� ������ DHCP'
,'�� ������� ������������ ������� DHCP','IP �� ������� �� ��������',
'IPX �� ������� �� ��������','������ ����� ����� / ����','������������ ��� ������'
,'������������ ����� ����','������������� ����� ����','�������� �� ��������� ������'
,'������ ������','������������ ������','��� ����������','����, ���� ��� ������ �� �������'
,'�� ������� ���������� ������','�� ������� ��������� ������ DNS'
,'��������� �� �������������','�� ��� �������� ������ DHCP ����� ���� �������� / ���������'
,'����������� ������','DHCP �� ������� �� ��������');
function ArrayToVarArray(Arr : Array Of string):OleVariant;overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
//////////////////////////////////////////////////
function ArrayToVarArray(Arr : Array Of Word):OleVariant;overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
////////////////////////////////////////////////////
begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('��������� DHCP:');
  OleInitialize(nil); ///// ����� ������ �����
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
         MyError:=FWbemObject.EnableDHCP();
         if (MyError>63) and (myerror<101) then frmDomainInfo.memo1.Lines.Add('DHCP: '+IPError[MyError])
         else frmDomainInfo.memo1.Lines.Add(SysErrorMessage(MyError));

         if OKRightDlg123456789101112131415.CheckBox4.Checked=true then
            begin
             MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder(MasDNS); //// �������� DNS
             if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-������ -'+IPerror[MyErrorDNS])
             else frmDomainInfo.memo1.Lines.Add('���������� ����������� DNS: '+SysErrorMessage(MyErrorDNS));
            end
         else
            begin
             MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder();
             if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-������ -'+IPerror[MyErrorDNS])
             else frmDomainInfo.memo1.Lines.Add('���������� ������������ DNS: '+SysErrorMessage(MyErrorDNS));
            end;


         frmDomainInfo.memo1.Lines.Add('---------------------------');
       end;
  oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ ��������� DHCP: "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;



end.
