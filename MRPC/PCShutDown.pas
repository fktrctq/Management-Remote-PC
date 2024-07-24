unit PCShutDown;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Shutdown = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses uMain;

{ Shutdown }

procedure Shutdown.Execute;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown:integer;
begin

  try
          frmDomainInfo.memo1.Lines.Add('Завершение работы, сеанса, перезагрузка....');
          OleInitialize(nil);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
          FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
          //FWMIService.Security_.impersonationlevel:=3;
          //FWMIService.Security_.authenticationLevel := 6;
          //FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   Требуется выключить локальную систему
          oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
          if oEnum.Next(1, FWbemObject, iValue) = 0 then
            case myShutdown of
             0:
              begin
              ResShutdown:=FWbemObject.Win32Shutdown(0);
              frmDomainInfo.memo1.Lines.Add('Завершение сеанса. '+SysErrorMessage(ResShutdown));
              frmDomainInfo.memo1.Lines.Add('-----------------------------');
              end;
              1:
              begin
               ResShutdown:=FWbemObject.Win32Shutdown(4);
               frmDomainInfo.memo1.Lines.Add('Принудительное завершение сеанса. '+SysErrorMessage(ResShutdown));
               frmDomainInfo.memo1.Lines.Add('-----------------------------');
              end;
               2:
               begin
                ResShutdown:=FWbemObject.Win32Shutdown(2);
               frmDomainInfo.memo1.Lines.Add('Перезагрузка. '+SysErrorMessage(ResShutdown));
               frmDomainInfo.memo1.Lines.Add('-----------------------------');
               end;
               3:
               begin
               ResShutdown:=FWbemObject.Win32Shutdown(6);
               frmDomainInfo.memo1.Lines.Add('Принудительная перезагрузка. '+SysErrorMessage(ResShutdown));
               frmDomainInfo.memo1.Lines.Add('-----------------------------');
               end;
               4:
               begin
               ResShutdown:=FWbemObject.Win32Shutdown(1);
               frmDomainInfo.memo1.Lines.Add('Завершение работы. '+SysErrorMessage(ResShutdown));
               frmDomainInfo.memo1.Lines.Add('-----------------------------');
               end;
                5:
                begin
                ResShutdown:=FWbemObject.Win32Shutdown(5);
               frmDomainInfo.memo1.Lines.Add('Принудительное завершение работы. '+SysErrorMessage(ResShutdown));
                frmDomainInfo.memo1.Lines.Add('-----------------------------');
                end;
                6:
                begin
                ResShutdown:=FWbemObject.Win32Shutdown(12);
                frmDomainInfo.memo1.Lines.Add('Принудительное завершение работы. '+SysErrorMessage(ResShutdown));
                frmDomainInfo.memo1.Lines.Add('-----------------------------');
                end;
              end;
          FWbemObject:=Unassigned;
          VariantClear(FWbemObject);
          oEnum:=nil;
          VariantClear(FWbemObjectSet);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
          OleUnInitialize;
          except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка - "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;

      end;




end;

end.
