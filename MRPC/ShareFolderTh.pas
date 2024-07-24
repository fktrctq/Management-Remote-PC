unit ShareFolderTh;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,WinSpool,WinApi.windows;




type
  ShareFolderThread = class(TThread)
  private
    function AddNetPrint(namePC,MyUser,MyPasswd,Path,Name,TypeFS,
    MaxAllowed,Descrpt,Pass,Acces:string):boolean;
  protected
    procedure Execute; override;
  end;

implementation
uses umain;

function AddNetPrint(namePC,MyUser,MyPasswd,Path,Name,TypeFS,
MaxAllowed,Descrpt,Pass,Acces:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService  : OLEVariant;
  objShare  :OLEVariant;
  MyError       : integer;

begin
try

     OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     //FWMIServiceDriver   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+Myps+'\root\CIMV2');
     FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
     FWMIService.Security_.impersonationlevel:=3;
     FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
     objShare:=FWMIService.get('Win32_share');
     MyError:=objShare.create(Path,Name,TypeFS,MaxAllowed,Descrpt,Pass,Acces);
     frmDomainInfo.memo1.Lines.Add(MyPS+'Создание сетевого ресурса : '+SysErrorMessage(MyError));
     ///////////////////////////////////////////////////////////
  VariantClear(objShare);
  VariantClear(FSWbemLocator);
  VariantClear(FWMIService);
  OleUnInitialize;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(MyPS+'Ошибка Создание сетевого ресурса "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;


end;
end;

procedure ShareFolderThread.Execute;
begin

end;

end.
