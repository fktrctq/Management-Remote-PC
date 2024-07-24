unit NetworkAdapterEnable;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterEnablethread = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;

{ NetworkAdapterEnablethread }

procedure NetworkAdapterEnablethread.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError       : integer;
  iValue        : LongWord;
begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('��������� ��������');
  OleInitialize(nil); ///// ����� ������ �����
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapter WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
         MyError:=FWbemObject.Enable();
         frmDomainInfo.memo1.Lines.Add(SysErrorMessage(MyError));
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
           frmDomainInfo.memo1.Lines.Add(' - ������ ��������� �������� "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;



end.
