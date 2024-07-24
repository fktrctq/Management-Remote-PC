unit AddPrintDriverThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  AddPrintDriver = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,AddPrintDriverDialog;

{ AddPrintDriver }

procedure AddPrintDriver.Execute;

var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  objDriver    : OLEVariant;
  ParObjDriver : OLEVariant;
  MyError       : integer;
  function GetObject(const objectName: String): IDispatch;
var
  chEaten: Integer;
  BindCtx: IBindCtx;//for access to a bind context
  Moniker: IMoniker;//Enables you to use a moniker object
begin
  OleCheck(CreateBindCtx(0, bindCtx));
  OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));//Converts a string into a moniker that identifies the object named by the string
  OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));//Binds to the specified object
end;
begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Установка драйвера.');
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  //FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy

  //FWMIService   := FSWbemLocator.ConnectServer(Myps, 'root\CIMV2', MyUser, MyPasswd);
  //FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_printer','WQL',wbemFlagForwardOnly);

  objDriver:=FWMIService.get('Win32_PrinterDriver');
  ParObjDriver:=FWMIService.get('Win32_PrinterDriver');
  if FilePath<>'' then ParObjDriver.Infname:=FilePath;
  ParObjDriver.Name:=DriverName;
  if DriverPath<>'' then ParObjDriver.DriverPath:=DriverPath;
  frmDomainInfo.memo1.Lines.Add('Путь - '+FilePath);
  frmDomainInfo.memo1.Lines.Add('Драйвер - '+DriverName);
  frmDomainInfo.memo1.Lines.Add('DLL - '+DriverPath);

  MyError:=objDriver.AddPrinterDriver(ParObjDriver);
  frmDomainInfo.memo1.Lines.Add(MyPS+'---Установка драйвера. '+SysErrorMessage(MyError));
  frmDomainInfo.memo1.Lines.Add('---------------------------');

    VariantClear(objDriver);
    VariantClear(ParObjDriver);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(MyPS+'---Ошибка установки драйвера "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
