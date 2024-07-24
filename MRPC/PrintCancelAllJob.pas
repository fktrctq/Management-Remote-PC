unit PrintCancelAllJob;

interface

uses
    System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  PrintCancelAllJobThread = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;

{ PrintCancelAllJobThread }

procedure PrintCancelAllJobThread.Execute;
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
  frmDomainInfo.memo1.Lines.Add('Очистить очередь печати.');
  OleInitialize(nil); ///// нахуй незнаю зачем
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
         frmDomainInfo.memo1.Lines.Add(FWbemObject.DeviceID);
         MyError:=FWbemObject.CancelAllJobs();
         frmDomainInfo.memo1.Lines.Add('Очистка очереди печати. '+SysErrorMessage(MyError));
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
           frmDomainInfo.memo1.Lines.Add('Ошибка очистки очереди печати "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
