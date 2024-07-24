unit ShareFS;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,WinSpool,WinApi.windows;

  function AddShareFS(namePC,MyUser,MyPasswd,Path,Name,
  Descrpt,pass:string;TypeFS:integer;Access:Olevariant):boolean;   //// ������� ������� ������ � ������������ ������������
  function DeleteShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;   /// ������� ������� ������
  function GetAccessMaskShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;  /// ������ ����
  function AddShareNoFS(namePC,MyUser,MyPasswd,Path,Name,Descrpt:string):boolean; /// ������� ���������� ���� ��� ����������� ������������
  implementation

uses umain;

function AddShareNoFS(namePC,MyUser,MyPasswd,Path,Name,Descrpt:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService  : OLEVariant;
  objShare  :OLEVariant;
  MyError       : integer;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
objShare:=FWMIService.get('Win32_share');
MyError:=objShare.create(Path,Name,0,true,Descrpt);
frmDomainInfo.memo1.Lines.Add(MyPS+'�������� �������� ������� '+Name+' :'+SysErrorMessage(MyError));
VariantClear(objShare);
VariantClear(FSWbemLocator);
VariantClear(FWMIService);
OleUnInitialize;
result:=true;
except
on E:Exception do
         begin
          result:=false;
          frmDomainInfo.memo1.Lines.Add(MyPS+'������ �������� �������� ������� '+name+' :"'+E.Message+'"');
          OleUnInitialize;
          exit;
         end;
end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////
function AddShareFS(namePC,MyUser,MyPasswd,Path,Name,                   /// ������� ���������� ���� � �������������� ����������� ������������
  Descrpt,pass:string;TypeFS:integer;Access:Olevariant):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService  : OLEVariant;
  objShare  :OLEVariant;
  MyError       : integer;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
objShare:=FWMIService.get('Win32_share');
MyError:=objShare.create(Path,Name,TypeFS,true,Descrpt,pass,Access);
frmDomainInfo.memo1.Lines.Add(MyPS+'�������� �������� ������� '+Name+' :'+SysErrorMessage(MyError));
VariantClear(objShare);
VariantClear(FSWbemLocator);
VariantClear(FWMIService);
OleUnInitialize;
result:=true;
except
on E:Exception do
         begin
           result:=false;
           frmDomainInfo.memo1.Lines.Add(MyPS+'������ �������� �������� ������� '+name+' :"'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;
end;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////

function DeleteShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService  : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  objShare  :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  MyError       : integer;
  step:integer;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_share WHERE Name LIKE "'+ShareName+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, objShare, iValue) = 0 then     //// ������������
begin
MyError:=objShare.delete();
frmDomainInfo.memo1.Lines.Add('���������� �������� ������� '+'"'+string(objShare.name)+'" : '+SysErrorMessage(MyError));
objShare:=Unassigned;
result:=true;
end;
 oEnum:=nil;
 VariantClear(objShare);
  VariantClear (FWbemObjectSet);
  VariantClear(FSWbemLocator);
  VariantClear(FWMIService);
  OleUnInitialize;
except
on E:Exception do
         begin
          result:=false;
           frmDomainInfo.memo1.Lines.Add(MyPS+'������ ���������� �������� ������� "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;
end;
end;

///////////////////////////////////////////////////////////////////////////////////////////////////////



function GetAccessMaskShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;
var
  FSWbemLocator       : OLEVariant;
  FWMIService  : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  objShare  :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  MyAccessMask       : integer;
  step:integer;
begin
try
     OleInitialize(nil);

     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd);
     FWMIService.Security_.impersonationlevel:=3;
     FWMIService.Security_.authenticationLevel := 6;
     FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_share WHERE Name LIKE "'+ShareName+'"','WQL',wbemFlagForwardOnly);
     oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         if oEnum.Next(1, objShare, iValue) = 0 then     //// ������������
            begin
            MyAccessMask:=objShare.GetAccessMask();
            frmDomainInfo.memo1.Lines.Add('GetAccessMask '+'"'+string(objShare.name)+'" : '+inttostr(MyAccessMask));
             objShare:=Unassigned;
            end;
     ///////////////////////////////////////////////////////////
 oEnum:=nil;
 VariantClear(objShare);
  VariantClear (FWbemObjectSet);
  VariantClear(FSWbemLocator);
  VariantClear(FWMIService);
  OleUnInitialize;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(MyPS+'������ GetAccessMask "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;
end;
end;
end.
