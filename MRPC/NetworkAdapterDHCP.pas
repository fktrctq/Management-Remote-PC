unit NetworkAdapterDHCP;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterDHCPEnable = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NetworkConfiguration;


{ NetworkAdapterDHCPEnable }

procedure NetworkAdapterDHCPEnable.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError,MyErrorDNS       : integer;
  MasDNS        :OLEvariant;
  iValue        : LongWord;
const
IpError: array [64..100] of string =('Метод не поддерживается на этой платформе',
'Неизвестный сбой','Недопустимая маска подсети','Произошла ошибка при обработке экземпляра который был возвращен'
,'Недопустимый входной параметр','Указано более 5 шлюзов','Недопустимый IP-адрес'
,'Недействительный IP-адрес шлюза','Произошла ошибка при доступе к реестру для запрашиваемой информации'
,'Недопустимое имя домена','Неверный адрес','Первичный / вторичный сервер WINS не определен'
,'неверный файл','Недопустимый системный путь','Ошибка копирования файла'
,'Недопустимый параметр безопасности','Невозможно настроить службу TCP / IP'
,'Невозможно настроить службу DHCP','Не удалось возобновить аренду DHCP'
,'Не удалось опубликовать договор DHCP','IP не включен на адаптере',
'IPX не включен на адаптере','Ошибка рамки кадра / сети','Недопустимый тип фрейма'
,'Недопустимый номер сети','Дублированный номер сети','Параметр за пределами границ'
,'Доступ закрыт','Недостаточно памяти','Уже существует','Путь, файл или объект не найдены'
,'Не удалось оповестить сервис','Не удалось уведомить службу DNS'
,'Интерфейс не настраивается','Не все договора аренды DHCP могут быть выпущены / обновлены'
,'Неизвестная ошибка','DHCP не включен на адаптере');
function ArrayToVarArray(Arr : Array Of string):OleVariant;overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
//////////////////////////////////////////////////
function ArrayToVarArray(Arr : Array Of Word):OleVariant;overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
////////////////////////////////////////////////////
begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Включение DHCP:');
  OleInitialize(nil); ///// нахуй незнаю зачем
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
         MyError:=FWbemObject.EnableDHCP();
         if (MyError>63) and (myerror<101) then frmDomainInfo.memo1.Lines.Add('DHCP: '+IPError[MyError])
         else frmDomainInfo.memo1.Lines.Add(SysErrorMessage(MyError));

         if OKRightDlg123456789101112131415.CheckBox4.Checked=true then
            begin
             MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder(MasDNS); //// основной DNS
             if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-сервер -'+IPerror[MyErrorDNS])
             else frmDomainInfo.memo1.Lines.Add('Назначение статических DNS: '+SysErrorMessage(MyErrorDNS));
            end
         else
            begin
             MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder();
             if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-сервер -'+IPerror[MyErrorDNS])
             else frmDomainInfo.memo1.Lines.Add('Назначение динамических DNS: '+SysErrorMessage(MyErrorDNS));
            end;


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
           frmDomainInfo.memo1.Lines.Add('Ошибка включения DHCP: "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;



end.
