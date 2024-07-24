unit RenamePC;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  PCRename = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;


procedure PCRename.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  ResRename:integer;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  iValue        : LongWord;
begin
try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  //FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
    frmDomainInfo.memo1.Lines.Add('Переименование компьютера....');
    ResRename:=FWbemObject.Rename(NewNamePC,PasswordDomain,UserDomain);
    frmDomainInfo.memo1.Lines.Add(MyPS+'---Переименование компьютера. '+SysErrorMessage(ResRename));
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    FWbemObject:=Unassigned;
  end;
  VariantClear(FWbemObject);
  if oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(MyPS+'---Ошибка переименования компьютера "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;
end;
end;

end.
