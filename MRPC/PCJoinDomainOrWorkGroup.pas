unit PCJoinDomainOrWorkGroup;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  JoinDomainOrWorkGroup = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
var
NameDomain,PasswordDomain,UserDomain,
JDOWPC,JDOWUSER,JDOWPASS:string;
JDOWOper:integer;
RebootAfterUnjoinDomain:boolean;

implementation
uses umain;


//  objProcess    : OLEVariant;
//  objConfig     : OLEVariant;
//  ProcessID     : Integer;


{ JoinDomainOrWorkGroup }
function RebootAfterunjoin(FWMIService   : OLEVariant):boolean;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown:integer;
begin
try
OleInitialize(nil);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
//FWMIService.Security_.impersonationlevel:=3;
//FWMIService.Security_.authenticationLevel := 6;
//FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   Требуется выключить локальную систему
oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then
begin
ResShutdown:=FWbemObject.Win32Shutdown(6);
frmDomainInfo.memo1.Lines.Add('Перезагрузка - '+SysErrorMessage(ResShutdown));
end;
FWbemObject:=Unassigned;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
//VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
 except
  on E:Exception do frmDomainInfo.memo1.Lines.Add('Ошибка - "'+E.Message+'"');
 end;
end;

procedure JoinDomainOrWorkGroup.Execute;
var
MyError:integer;
 FWbemObjectSet: OLEVariant;
 oEnum: IEnumvariant;
   FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  iValue        : LongWord;
Const JOIN_DOMAIN = 1 ;
Const ACCT_CREATE = 2 ;
Const ACCT_DELETE = 4 ;
Const WIN9X_UPGRADE = 16 ;
Const DOMAIN_JOIN_IF_JOINED = 32 ;
Const JOIN_UNSECURE = 64  ;
Const MACHINE_PASSWORD_PASSED = 128 ;
Const DEFERRED_SPN_SET = 256;
Const INSTALL_INVOCATION = 262144;

begin
try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(JDOWPC, 'root\CIMV2', JDOWUSER, JDOWPASS,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
     MyError:=FWbemObject.JoinDomainOrWorkgroup(NameDomain,PasswordDomain,UserDomain
     ,null,JDOWOper);
     frmDomainInfo.memo1.Lines.Add('Добавить в домен. ('+inttostr(MyError)+') '+SysErrorMessage(MyError));
     if RebootAfterUnjoinDomain and (MyError=0) then RebootAfterunjoin(FWMIService)
      else if MyError=0 then frmDomainInfo.memo1.Lines.Add('Необходимо перезагрузить компьютер');  //2691 - компьютер уже подсоединен к домену  // 5 отказано в доступе
  end;
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка  "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
