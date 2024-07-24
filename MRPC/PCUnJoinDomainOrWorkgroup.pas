unit PCUnJoinDomainOrWorkgroup;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  UnjoinDomainOrWorkgroup = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,DlgUnjoinDomainOrWorkgroup;


var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
//  objProcess    : OLEVariant;
//  objConfig     : OLEVariant;
//  ProcessID     : Integer;

{ PCUnjoinDomainOrWorkgroup }

procedure UnjoinDomainOrWorkgroup.Execute;
var
MyError:integer;
FWbemObjectSet: OLEVariant;
oEnum : IEnumvariant;
iValue        : LongWord;
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
begin
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  //FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
     MyError:=FWbemObject.UnjoinDomainOrWorkgroup(PasswordDomain,UserDomain,0);    /// 0 - выводим из домена, 4- вывод из домена и удаление учетки компа в домене
     frmDomainInfo.memo1.Lines.Add('Вывести ПК из домена. '+SysErrorMessage(MyError));
     if MyError=0 then
     begin
    // if RebootAfterUnjoinDomain then frmDomainInfo.rebuutSelectPC
    //  else frmDomainInfo.memo1.Lines.Add('Необходимо перезагрузить компьютер');
      end;
     frmDomainInfo.memo1.Lines.Add('---------------------------');
  end;
  FWbemObject:=Unassigned;
  VariantClear(FWbemObject);
  if oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
