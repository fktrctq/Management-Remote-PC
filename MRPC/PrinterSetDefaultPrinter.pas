unit PrinterSetDefaultPrinter;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  DefaultPrinter = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;


{ DefaultPrinter }

procedure DefaultPrinter.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError       : integer;
  iValue        : LongWord;
  step:integer;
 {  function GetObject(const objectName: String): IDispatch;
var
  chEaten: Integer;
  BindCtx: IBindCtx;//for access to a bind context
  Moniker: IMoniker;//Enables you to use a moniker object
begin
  OleCheck(CreateBindCtx(0, bindCtx));
  OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));//Converts a string into a moniker that identifies the object named by the string
  OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));//Binds to the specified object
end; }
begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Использовать принтер по умолчанию');
  OleInitialize(nil);
  step:=1;
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  step:=2;
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  step:=3;
  FWMIService.Security_.impersonationlevel:=3;
  step:=4;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
  //FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  //FWMIService := CreateOleObject('WbemScripting.SWbemLocator');
  //FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate}!\\'+MyPS+'\root\CIMV2');
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  step:=5;
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
     if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
      step:=6;
         frmDomainInfo.memo1.Lines.Add(FWbemObject.DeviceID);
         step:=7;
         MyError:=FWbemObject.SetDefaultPrinter();
         step:=8;
         frmDomainInfo.memo1.Lines.Add('Использовать принтер по умоланию '+SysErrorMessage(MyError));
         frmDomainInfo.memo1.Lines.Add('---------------------------');
      end;
      step:=9;
  If oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка операции "'+E.Message+'"'+' step - '+inttostr(step));
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;


end.
