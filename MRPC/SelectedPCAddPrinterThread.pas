unit SelectedPCAddPrinterThread;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,
  IdIcmpClient,Vcl.Controls;

type
  SelectedPCAddPrint = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
  var
  MyIdIcmpClient: TIdIcmpClient;

implementation
uses umain,NetworkAddPrint;

function finditem(s,name:string;image:integer):boolean;
var
z:integer;
begin
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
    if frmDomainInfo.ListView8.Items[z].SubItems[0]=name then
    begin
    frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=image;
    frmDomainInfo.ListView8.Items[z].SubItems[1]:=s;
    break;
    end;
end;

 function ping(s:string):boolean;
var
z:integer;
begin
try
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  finditem('Превышен интервал ожидания запроса',s,2);
  end
else
  begin
  result:=true; ///доступен
  frmDomaininfo.Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;

   except
    begin
    result:=false;
    finditem('Узел не доступен',s,2);
    end;
   end;
end;

{ SelectedPCAddPrint }

function InstallPrintOrDriverForListPC(SelectedPCAdd : TstringList):boolean;
var
  FSWbemLocator : OLEVariant;
  FWMIService,FWMIServiceDriver   : OLEVariant;
  objPrinter,objPrintDriver,objParPrintDriver,objNewPort    :OLEVariant;
  MyError,i,z   : integer;
  ResultPut     : string;


begin
//////////////////////////////////////
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
////////////////////////////////////////
for I := 0 to SelectedPCAdd.Count-1 do
  begin
      if (SelectedPCAdd[i]<>'') and (ping(SelectedPCAdd[i])) then
        try
          begin
           OleInitialize(nil);
           FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
           FWMIServiceDriver   := FSWbemLocator.ConnectServer(SelectedPCAdd[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
           FWMIServiceDriver.Security_.impersonationlevel:=3;
           FWMIServiceDriver.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
           if installdriver then
             begin
             finditem('Запущен процесс установки драйвера '+DriverName,SelectedPCAdd[i],16);
             objPrintDriver:=FWMIServiceDriver.get('Win32_PrinterDriver');
             objParPrintDriver:=FWMIServiceDriver.get('Win32_PrinterDriver');
             if FilePath<>'' then objParPrintDriver.Infname:=FilePath;  //// inf файл
             objParPrintDriver.Name:=DriverName;  /// имя драйвера
             if DriverPath<>'' then objParPrintDriver.DriverPath:=DriverPath; // путь к dll
             frmDomainInfo.memo1.Lines.Add('Драйвер - '+DriverName);
             frmDomainInfo.memo1.Lines.Add('Путь - '+FilePath);
             frmDomainInfo.memo1.Lines.Add('DLL - '+DriverPath);
             MyError:=objPrintDriver.AddPrinterDriver(objParPrintDriver); /// установка драйвера
             frmDomainInfo.memo1.Lines.Add('Установка драйвера. '+SysErrorMessage(MyError));
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             end;
            if PrintIP then
              begin
                ///// запускаем если принтер по IP
               /////////////////////////////// порт
               finditem('Запущен процесс установки принтера ',SelectedPCAdd[i],15);
               objNewPort:=FWMIServiceDriver.get ('Win32_TCPIPPrinterPort').SpawnInstance_;
               objNewPort.Name := PortName;
               objNewPort.Protocol := PortProtocol;
               objNewPort.HostAddress := HostIPAddress;
               objNewPort.PortNumber := PortNumber;
               objNewPort.SNMPEnabled := SNMPEnabled;
               resultput:=vartostr(objNewPort.put_());
               frmDomainInfo.memo1.Lines.Add(SelectedPCAdd[i]+' - настройка порта завершена  - '+ResultPut);
               //////////////////////////////// принтер
               objPrinter:=FWMIServiceDriver.get('Win32_printer').SpawnInstance_;
               objPrinter.DriverName := DriverName;
               objPrinter.PortName   := PortName;
               objPrinter.DeviceID   := DriverName;
               objPrinter.Network := PrintNetwork;
               objPrinter.Shared := PrintShare;
               resultput:=vartostr(objPrinter.put_());
               MyError:=0;
               frmDomainInfo.memo1.Lines.Add('Установка принтера завершена.  - '+ResultPut);
               ///////////////////////////////////////////////////////////
              end;
              if printNet then ///// установка сетевого принтера
                begin
                MyError:=objPrinter.AddPrinterConnection(AddNewNetPrint); /// операция добавления принтера
                end;

            if MyError=0 then
              begin
              finditem('Подключение принтера "'+DriverName+'" прошло успешно',SelectedPCAdd[i],1);
              end
              else
              begin
              finditem('При подключении принтера "'+DriverName+'" возникли проблемы. Ошибка'+SysErrorMessage(MyError),SelectedPCAdd[i],2);
              end;

            VariantClear(objPrintDriver);
            VariantClear(objParPrintDriver);
            VariantClear(objNewPort);
            VariantClear(objPrinter);
            VariantClear(FWMIService);
            VariantClear(FSWbemLocator);
            VariantClear(FWMIServiceDriver);
            OleUnInitialize;
          end;
          except
          on E:Exception do
          begin
          finditem('Ошибка подключение принтера "'+DriverName+'". - '+E.Message,SelectedPCAdd[i],2);
          frmDomainInfo.memo1.Lines.Add('Ошибка добавления сетевого принтера "'+E.Message+'"');
          frmDomainInfo.memo1.Lines.Add('---------------------------');
          OleUnInitialize;
          end;
          end;

 end;

if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
if Assigned(SelectedPCAdd) then freeandnil(SelectedPCAdd);
///FreeAndNil(SelectedPCAdd);
end;

procedure SelectedPCAddPrint.Execute;
begin
InstallPrintOrDriverForListPC(SelectedPC);
SelectedPC.Free;

end;

end.
