unit DefragAnalysis;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;


type
  HDDDefragAnalysis = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,DefragHDD;
var
  FSWbemLocator    : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject,objAnalysis   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
const
//wbemFlagForwardOnly = $00000020;
ReturnVal:array [0..11] of string=('Операция завершена успешно','Доступ запрещен',
'Не поддерживается','Volume Dirty Bit Is Set','Не хватает свободного места',
'Поврежденная таблица основных файлов','Вызов отменен',
'Call Cancellation Request Too Late','Дефрагментация уже запущена','Не удается произвести дефрагментацию','Ошибка дефрагментации','Неизвестная ошибка');


{ HDDDefragAnalysis }

procedure HDDDefragAnalysis.Execute;
var
DefRes:byte;
begin
try
  begin
   OleInitialize(nil); ///// нахуй незнаю зачем
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(Myps, 'root\CIMV2', MyUser, MyPasswd);
      objAnalysis    := FWMIService.Get('Win32_DefragAnalysis');
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+SelectedHDD+'%"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
       DefRes:=FWbemObject.DefragAnalysis(true,objAnalysis);
       //VariantClear(FWbemObject);
      end;
      oEnum:=nil;
      VariantClear(FWbemObject);
      VariantClear(objAnalysis);
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      // FWbemObject:=Unassigned;
       frmDomainInfo.memo1.Lines.Add('Анализ дефрагментации завершен Return value - '+ReturnVal[DefRes]);
   end;
 OleUnInitialize;
except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка анализа дефрагментации "'+E.Message+'"');
          exit;
         end;
      end;

end;

end.
