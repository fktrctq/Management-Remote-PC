unit NetworkAdapterStaticTCPIPAll;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterStaticIPALL = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NetworkConfiguration;


{ NetworkAdapterStaticIPALL }

procedure NetworkAdapterStaticIPALL.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyErrorIP,MyErrorGateway,MyErrorDNS,MyErrorWins       : integer;
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

begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('���������� ������������ IP-������');
  OleInitialize(nil); ///// ����� ������ �����
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
           if YesWins<>'' then MyErrorWins:=FWbemObject.setWINSServer(YesWins)
              else MyErrorWins:=FWbemObject.setWINSServer('','');
           if (YesWins<>'') and (YesWins1<>'') then MyErrorWins:=FWbemObject.setWINSServer(YesWins,YesWINS1);
           MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder(MasDNS); //// �������� DNS
           MyErrorGateway:=FWbemObject.SetGateways(MasGateway,MasGatewayMetric);
           MyErrorIP:=FWbemObject.EnableStatic(MasIP,MasSubnet);

         if (MyErrorIP>63) and(MyErrorIP<101) then frmDomainInfo.memo1.Lines.Add('IP-����� -'+IPerror[MyErrorIP])
           else frmDomainInfo.memo1.Lines.Add('IP-����� -'+SysErrorMessage(MyErrorIP));
         if (MyErrorGateway>63) and(MyErrorGateway<101) then frmDomainInfo.memo1.Lines.Add('����-'+IPerror[MyErrorGateway])
           else frmDomainInfo.memo1.Lines.Add('���� -'+SysErrorMessage(MyErrorGateway));
         if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-������ -'+IPerror[MyErrorDNS])
           else frmDomainInfo.memo1.Lines.Add('DNS-������ -'+SysErrorMessage(MyErrorDNS));
         if (MyErrorWINS>63) and(MyErrorWINS<101) then frmDomainInfo.memo1.Lines.Add('WINS-������ -'+IPerror[MyErrorWINS])
           else frmDomainInfo.memo1.Lines.Add('WINS-������ -'+SysErrorMessage(MyErrorWINS));
         frmDomainInfo.memo1.Lines.Add('---------------------------');
       end;
  varClear(MasIP);
  varClear(MasSubnet);
  varClear(MasGateway);
  varClear(MasGatewayMetric);
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
           frmDomainInfo.memo1.Lines.Add('������ ���������� ������������ IP-������ "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;


end.
