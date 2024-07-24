unit AddNetworkPrint;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,windows,
  WinSpool;

type
  AddNetPrint = class(TThread)

  private

    { Private declarations }
  protected
    procedure Execute; override;
  end;


implementation
uses umain,NetworkAddPrint;



{ AddNetPrint }

procedure AddNetPrint.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService,FWMIServiceDriver   : OLEVariant;
  objPrinter,objPrintDriver,objParPrintDriver,objNewPort,printname    :OLEVariant;
  MyError       : integer;
  ResultPut     :string;
  pPrinter:PprinterInfo2;

/////////////////////////////////////////////////////////////////
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
/////////////////////////////////////////////////////////////////////
function AddAPrinter(PrinterName, PortName,
DriverName, PrintProcessor: string):     boolean;
var
  pName: PChar;
  Level: DWORD;
  pPrinter: PPrinterInfo2;
begin

  pName := nil;
  Level := 2;
  New(pPrinter);
  pPrinter^.pServerName := PChar(' ');
  pPrinter^.pShareName := nil;
  pPrinter^.pComment := nil;
  pPrinter^.pLocation := nil;
  pPrinter^.pDevMode := nil;
  pPrinter^.pSepFile := nil;
  pPrinter^.pDatatype := nil;
  pPrinter^.pParameters := nil;
  pPrinter^.pSecurityDescriptor := nil;
  pPrinter^.Attributes := 0;
  pPrinter^.Priority := 0;
  pPrinter^.DefaultPriority := 0;
  pPrinter^.StartTime := 0;
  pPrinter^.UntilTime := 0;
  pPrinter^.Status := 0;
  pPrinter^.cJobs := 0;
  pPrinter^.AveragePPM :=0;

  pPrinter^.pPrinterName := PCHAR(PrinterName);
  pPrinter^.pPortName := PCHAR(PortName);
  pPrinter^.pDriverName := PCHAR(DriverName);
  pPrinter^.pPrintProcessor := PCHAR(PrintProcessor);
  frmDomainInfo.memo1.Lines.Add('.........');
  if AddPrinter(pName, Level, pPrinter) <> 0 then
  begin
    Result := true
  end
  else
  begin
     ShowMessage(inttostr(GetlastError));
    Result := false;
  end;
end;
//////////////////////////////////////////////////////////////////////////
begin

try
begin
     OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     //FWMIServiceDriver   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+Myps+'\root\CIMV2');
     FWMIServiceDriver   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
     FWMIServiceDriver.Security_.impersonationlevel:=3;
     FWMIServiceDriver.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
           if installdriver then
             begin
             objPrintDriver:=FWMIServiceDriver.get('Win32_PrinterDriver');
             objParPrintDriver:=FWMIServiceDriver.get('Win32_PrinterDriver');
             if FilePath<>'' then  objParPrintDriver.Infname:=FilePath;  //// inf файл
             objParPrintDriver.Name:=DriverName;  /// имя драйвера
             if DriverPath<>'' then objParPrintDriver.DriverPath:=DriverPath; /// путь к dll
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
               objNewPort:=FWMIServiceDriver.get ('Win32_TCPIPPrinterPort').SpawnInstance_;
               objNewPort.Name := PortName;
               objNewPort.Protocol := PortProtocol;
               objNewPort.HostAddress := HostIPAddress;
               objNewPort.PortNumber := PortNumber;
               objNewPort.SNMPEnabled := SNMPEnabled;
               resultPut:=vartostr(objNewPort.put_());
               frmDomainInfo.memo1.Lines.Add('Настройка порта завершена - '+ResultPut);
               //////////////////////////////// принтер
               objPrinter:=FWMIServiceDriver.get('Win32_printer').SpawnInstance_;
               objPrinter.DriverName := DriverName;
               objPrinter.PortName   := PortName;
               objPrinter.DeviceID   := DriverName;
               objPrinter.Network := PrintNetwork;
               objPrinter.Shared := PrintShare;
               resultPut:=vartostr(objPrinter.put_());
               MyError:=0;
               frmDomainInfo.memo1.Lines.Add(MyPS+' --- Установка принтера завершена. '+resultPut);
               ///////////////////////////////////////////////////////////
              end;
              if printNet then ///// установка общесетевого принтера
                begin
                exit;
                frmDomainInfo.memo1.Lines.Add(MyPS+' --- Добавить сетевой принтер. '+SysErrorMessage(MyError)+'-'+resultPut);
                frmDomainInfo.memo1.Lines.Add('---------------------------');
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
           frmDomainInfo.memo1.Lines.Add(MyPS+' --- Ошибка добавления сетевого принтера "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
