unit DefragForce;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  MyDefragForce = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,DefragHDD;
{var
  FSWbemLocator,FSWbemLocator2    : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject  : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
const
wbemFlagForwardOnly = $00000020;
}
{procedure MyDefragForce.Execute;
var
DefRes:byte;
begin
try
  begin
   OleInitialize(nil); ///// нахуй незнаю зачем
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(Myps, 'root\CIMV2', MyUser, MyPasswd);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+SelectedHDD+'%"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
       DefRes:=FWbemObject.Defrag(true);
       FWbemObject:=Unassigned;
      end;
       frmDomainInfo.memo1.Lines.Add('Дефрагментация завершена Return value - '+inttostr(DefRes));
   end;
 OleUnInitialize;
except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка дефрагментации "'+E.Message+'"');
          exit;
         end;
      end;

end;}
//var
//  FSWbemLocator,FSWbemLocator2    : OLEVariant;
//  FWMIService   : OLEVariant;
 // FWbemObjectSet: OLEVariant;
 // FWbemObject  : OLEVariant;
 // oEnum         : IEnumvariant;
 // iValue        : LongWord;
 // FInParams      :OleVariant;
{ MyDefragForce }  //Дефрагментация с предварительным анализом
const
wbemFlagForwardOnly = $00000020;
ReturnVal:array [0..11] of string=('Операция завершена успешно','Доступ запрещен',
'Не поддерживается','Volume Dirty Bit Is Set','Не хватает свободного места',
'Поврежденная таблица основных файлов','Вызов отменен',
'Call Cancellation Request Too Late','Дефрагментация уже запущена','Не удается произвести дефрагментацию','Ошибка дефрагментации','Неизвестная ошибка');
function  VolumeDefrag(IDHDD :string):boolean;

var
  FSWbemLocator : OleVariant;
  FWMIService   : oleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject   : OleVariant;
  FInParams     : OleVariant;
  FOutParams    : OleVariant;
  FBol:boolean;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  res,res1:integer;
  err:string;
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
  OleInitialize(nil);
  err:='0';
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  err:='1';
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  //FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  err:='2';
  err:='3';
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+IDHDD+'%"','WQL',wbemFlagForwardOnly);
  err:='4';
  FOutParams:=FWMIService.get('Win32_DefragAnalysis');
  //Fbol:=true;
  err:='5';
  res:=0;
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin                 /// функция возвращает неизвестную ошибку КОД -11, wmi не позволяет делать дефрагментацию диска на удаленном компьютере
      err:=err+'6';
      res:=FWbemObject.DefragAnalysis(FBol,FOutParams);
      res1:=FWbemObject.defrag();

      err:=err+'7';
      //FWbemObject:=Unassigned;
      end;
   err:=err+'8';
  //frmDomainInfo.memo1.Lines.Add('результат - '+inttostr(res));
  if (FOutParams.AverageFileSize)<>null then  FrmDomainInfo.memo1.Lines.Add('AverageFileSize - '+vartostr(FOutParams.AverageFileSize));
  frmDomainInfo.memo1.Lines.Add('Шаг '+(err)+' Res - '+inttostr(res)+'/ res1- '+inttostr(res1));
  frmDomainInfo.memo1.Lines.Add('ID HDD '+IDHDD);
  //FWbemObject   := FWMIService.Get(WmiPath,0,nil);
 // FInParams     := FWbemObject.Methods_.Item('Defrag',0).InParameters.SpawnInstance_(0);

  //varValue :=True;
  //FInParams.Properties_.Item('Force', 0).Set_Value(varValue);
  //FOutParams    := FWMIService.ExecMethod('Win32_Volume.DeviceID="\\\\?\\'+WmiPath+'\\"', 'Defrag', FInParams,0 , nil);
  //frmDomainInfo.memo1.Lines.Add(string('Результат дефрагментации  '+ReturnVal[integer(FOutParams.Properties_.Item('ReturnValue',0).Get_Value)]));
 except
   // on E:EOleException do
   //     frmDomainInfo.memo1.Lines.Add(string('EOleException ' +E.Message));
    on E:Exception do
    begin
      frmDomainInfo.memo1.Lines.Add('ID HDD '+IDHDD);
      frmDomainInfo.memo1.Lines.Add('Общая ошибка '+E.Classname+':'+E.Message);
      frmDomainInfo.memo1.Lines.Add('Шаг '+(err)+' Res - '+inttostr(res)+'/ res1- '+inttostr(res1));
    end;
 end;
  VariantClear(FWbemObject);
  VariantClear(FoutParams);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);

  OleUnInitialize;
end;



procedure MyDefragForce.Execute;
begin
VolumeDefrag(SelectedHDD);
end;


end.
