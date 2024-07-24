unit AddPrintLAN;



interface
uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,WinSpool,WinApi.windows;

 function AddNetPrint(namePC,MyUser,MyPasswd,PrintName:string):boolean;

implementation
uses umain;

/// ��������� �������� ������ �� ��������� ������
function AddNetPrint(namePC,MyUser,MyPasswd,PrintName:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIServiceDriver  : OLEVariant;
  objPrinter  :OLEVariant;
  MyError       : integer;

begin
try

     OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     //FWMIServiceDriver   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+Myps+'\root\CIMV2');
     FWMIServiceDriver   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
     FWMIServiceDriver.Security_.impersonationlevel:=3;
     FWMIServiceDriver.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
     objPrinter:=FWMIServiceDriver.get('Win32_printer');
     MyError:=objPrinter.AddPrinterConnection(PrintName);
     frmDomainInfo.memo1.Lines.Add(MyPS+' --- ��������� �������� ���������. '+SysErrorMessage(MyError));
     ///////////////////////////////////////////////////////////
  VariantClear(objPrinter);
  VariantClear(FSWbemLocator);
  VariantClear(FWMIServiceDriver);
  OleUnInitialize;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(MyPS+' --- ������ ���������� �������� �������� "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;


end;
end;

end.
