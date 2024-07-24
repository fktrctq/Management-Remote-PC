unit ChangeOwnerThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  UserChangeOwner = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,NewUserProfileDialog;


{ UserChangeOwner }

procedure UserChangeOwner.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError       : integer;
  NewUserSID  :string;
  iValue        : LongWord;
begin
try
begin
  NewUserSID:='Unknown';
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('������ ��������� (������������) �������.');
  OleInitialize(nil); ///// ����� ������ �����
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', MyUser, MyPasswd);
  //////////////////////////////////////////////////////////////////////////////////
  begin       /// ��������� SID ������ ������������
  frmDomainInfo.memo1.Lines.Add('�������� SID ������������ - '+NewUserName);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,SID FROM Win32_UserAccount WHERE Name = '+'"'+NewUserName+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
         NewUserSID:=vartostr(FWbemObject.SID);
         frmDomainInfo.memo1.Lines.Add('SID - '+NewUserSID);
      end;
      oEnum:=nil;
      VariantClear(FWbemObject);
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
  end;
  //////////////////////////////////////////////////////////////////////////////////
  if NewUserSID='Unknown' then
   begin
   frmDomainInfo.memo1.Lines.Add('SID ������������ �� ���������, �������� �� ����� ������� ��� ������������!');
   exit;
   end;
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserProfile WHERE SID = '+'"'+UserSID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
         frmDomainInfo.memo1.Lines.Add('������ ������� ������������.....');
         MyError:=FWbemObject.ChangeOwner(NewUserSID,FlagsBool);
         frmDomainInfo.memo1.Lines.Add('������ ��������� �������. '+SysErrorMessage(MyError));
         frmDomainInfo.memo1.Lines.Add('---------------------------');
      end;
  If oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ ������ ��������� �������. "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;


end.
