unit NetworkAdapterDNSThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterDNS = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NetworkConfiguration;


{ NetworkAdapterDNS }

procedure NetworkAdapterDNS.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyErrorDNS    : integer;
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
  ////////////////////////////////////////////////
  function ArrayToVarArray(Arr : Array Of string):OleVariant; overload;
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
  frmDomainInfo.memo1.Lines.Add('���������� ������������ DNS-�������:');
  OleInitialize(nil); ///// ����� ������ �����
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
         MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder(MasDNS); //// �������� DNS
         if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-������ -'+IPerror[MyErrorDNS])
         else frmDomainInfo.memo1.Lines.Add('���������� ����������� DNS: '+SysErrorMessage(MyErrorDNS));
         frmDomainInfo.memo1.Lines.Add('---------------------------');
       end;
  varClear(MasDNS);
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
           frmDomainInfo.memo1.Lines.Add('������ ������������ DNS-�������: "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;



end.
