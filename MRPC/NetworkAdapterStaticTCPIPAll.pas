unit NetworkAdapterStaticTCPIPAll;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  NetworkAdapterStaticIPALL = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NetworkConfiguration;


{ NetworkAdapterStaticIPALL }

procedure NetworkAdapterStaticIPALL.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyErrorIP,MyErrorGateway,MyErrorDNS,MyErrorWins       : integer;
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

begin
try
begin
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Назначение статического IP-адреса');
  OleInitialize(nil); ///// нахуй незнаю зачем
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
           if YesWins<>'' then MyErrorWins:=FWbemObject.setWINSServer(YesWins)
              else MyErrorWins:=FWbemObject.setWINSServer('','');
           if (YesWins<>'') and (YesWins1<>'') then MyErrorWins:=FWbemObject.setWINSServer(YesWins,YesWINS1);
           MyErrorDNS:=FWbemObject.SetDNSServerSearchOrder(MasDNS); //// основной DNS
           MyErrorGateway:=FWbemObject.SetGateways(MasGateway,MasGatewayMetric);
           MyErrorIP:=FWbemObject.EnableStatic(MasIP,MasSubnet);

         if (MyErrorIP>63) and(MyErrorIP<101) then frmDomainInfo.memo1.Lines.Add('IP-адрес -'+IPerror[MyErrorIP])
           else frmDomainInfo.memo1.Lines.Add('IP-адрес -'+SysErrorMessage(MyErrorIP));
         if (MyErrorGateway>63) and(MyErrorGateway<101) then frmDomainInfo.memo1.Lines.Add('Шлюз-'+IPerror[MyErrorGateway])
           else frmDomainInfo.memo1.Lines.Add('Шлюз -'+SysErrorMessage(MyErrorGateway));
         if (MyErrorDNS>63) and(MyErrorDNS<101) then frmDomainInfo.memo1.Lines.Add('DNS-сервер -'+IPerror[MyErrorDNS])
           else frmDomainInfo.memo1.Lines.Add('DNS-сервер -'+SysErrorMessage(MyErrorDNS));
         if (MyErrorWINS>63) and(MyErrorWINS<101) then frmDomainInfo.memo1.Lines.Add('WINS-сервер -'+IPerror[MyErrorWINS])
           else frmDomainInfo.memo1.Lines.Add('WINS-сервер -'+SysErrorMessage(MyErrorWINS));
         frmDomainInfo.memo1.Lines.Add('---------------------------');
       end;
  varClear(MasIP);
  varClear(MasSubnet);
  varClear(MasGateway);
  varClear(MasGatewayMetric);
  varClear(MasDNS);
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
           frmDomainInfo.memo1.Lines.Add('Ошибка назначения статического IP-адреса "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;


end.
