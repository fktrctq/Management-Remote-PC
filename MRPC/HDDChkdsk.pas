unit HDDChkdsk;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Chkdsk = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
const
ResultHDD:array [0..5] of string=('Команда Chkdsk успешно завершена', 'Для выпонения проверки перезагрузити компьютер','Файловая система не поддерживается','Неизвестная файловая система','Отсутствие носителя в приводе','Неизвестная ошибка');
{ Chkdsk }

procedure Chkdsk.Execute;
var
resChkdsk:integer;
begin
try
  begin
     OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+SelectedHDD+'%"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            frmDomainInfo.memo1.Lines.Add('Chkdsk - запушена проверка диска '+(string(FWbemObject.caption)));
            reschkdsk:=FWbemObject.Chkdsk(FixErrors,VigorousIndexCheck,SkipFolderCycle,ForceDismount,RecoverBadSectors,OkToRunAtBootUp); //// проверка диска
            FWbemObject:=Unassigned;
           end;

      frmDomainInfo.memo1.Lines.Add('Chkdsk - '+ResultHDD[reschkdsk]);
      VariantClear(FWbemObject);
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
  end;
 except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка Chkdsk "'+E.Message+'"');
          exit;
         end;
      end;
end;




end.
