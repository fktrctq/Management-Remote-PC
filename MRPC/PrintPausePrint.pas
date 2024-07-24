unit PrintPausePrint;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  PrintPause = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;

{ PrintPause }

procedure PrintPause.Execute;
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
  frmDomainInfo.memo1.Lines.Add('Остановить печать.');
  OleInitialize(nil); ///// нахуй незнаю зачем
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
         frmDomainInfo.memo1.Lines.Add(FWbemObject.DeviceID);
         MyError:=FWbemObject.Pause();
         frmDomainInfo.memo1.Lines.Add('Остановить печать '+SysErrorMessage(MyError));
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
           frmDomainInfo.memo1.Lines.Add('Ошибка остановки печати "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
