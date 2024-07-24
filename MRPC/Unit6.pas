unit Unit6;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Service = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses uMain;

{ Service }

procedure Service.Execute;
var
ResServic:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
Const OWN_PROCESS = 16;
 NOT_INTERACTIVE = True;
 ErrorCodeService: array [0..24] of string=('Операция успешно завершена',
 'Операция не поддерживается','У пользователя нет необходимого доступа',
 'Службу невозможно остановить, поскольку от нее зависят другие запущенные службы',
 'Запрошенный контрольный код недействителен или неприемлем для службы.'
 ,'Запрошенный код элемента управления не может быть отправлен службе из-за состояния службы '
 ,'Служба не запущена','Служба не ответила на запрос своевременно',
 'Неизвестный сбой при запуске службы','Путь к исполняемому файлу службы не найден'
 ,'Служба уже запущена.','База данных для добавления новой службы заблокирована',
 'Зависимость, на которую опирается эта служба, была удалена из системы.',
 'Службе не удалось найти службу, необходимую для зависимой службы.',
 ' Служба была изключена из системы.','Служба не имеет правильную проверку подлинности для запуска в системе.'
 ,'Эта служба удаляется из системы','Служба не имеет потока выполнения',' При запуске служба имеет циклические зависимости'
 ,'Служба выполняется под тем же именем','Имя службы содержит недопустимые символы'
 ,'Службе переданы недопустимые параметры.','Учетная запись, под которой выполняется эта служба, является недопустимой или не имеет разрешений на запуск службы.'
 ,'Служба существует в базе данных сервисов, доступных из системы.',
 'В настоящее время служба приостановлена в системе');
begin
      try
      begin
        OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
      begin
      if ActionServic=3 then FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Service','WQL',wbemFlagForwardOnly)
      else FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Service WHERE Name = '+'"'+SelectServic+'"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            frmDomainInfo.memo1.Lines.Add('---------------------------');
            case ActionServic of
            ///////////////////
            1:
            begin
            frmDomainInfo.memo1.Lines.Add('Запуск службы...');
            ResServic:=FWbemObject.StartService();  ////// запустить службу
            FWbemObject:=Unassigned;
            end;
            /////////////////
            2:
            begin
            frmDomainInfo.memo1.Lines.Add('Остановка службы...');
            ResServic:=FWbemObject.StopService();  ////// остановить службу
            FWbemObject:=Unassigned;
            end;
            3:
            begin
            //ResServic:=FWbemObject3.create;//('"DbService"', '"Personnel Database"', '"C:\Program Files (x86)\2gis\3.0\2GISUpdateService.exe"',OWN_PROCESS ,1 ,'"Automatic"' ,NOT_INTERACTIVE,'".\LocalSystem"','');  ////// Создать службу
            //FWbemObject3:=Unassigned;
            end;
            //////////////////
            4:
            begin
            frmDomainInfo.memo1.Lines.Add('Смена типа запуска службы...');
            ResServic:=FWbemObject.ChangeStartMode(TypeRunService); //// смена типа запуска
            FWbemObject:=Unassigned;
            end;
            5:
            begin
            frmDomainInfo.memo1.Lines.Add('Удаление службы...');
            ResServic:=FWbemObject.Delete(); //// Удалениме службы
            FWbemObject:=Unassigned;
            end;
            end;

            end;
         if ResServic>24 then frmDomainInfo.memo1.Lines.Add('Cлужбы - '+SysErrorMessage(ResServic))
          else frmDomainInfo.memo1.Lines.Add('Cлужбы - '+ErrorCodeService[ResServic]);


         frmDomainInfo.memo1.Lines.Add('---------------------------');
          VariantClear(FWbemObject);
          oEnum:=nil;
          VariantClear(FWbemObjectSet);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
         OleUnInitialize;  /// тоже хуйня непонятная но надо
      end;


      end;
      except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка "'+E.Message+'"');
          exit;
         end;
      end;

end;





end.
