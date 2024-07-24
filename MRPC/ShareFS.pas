unit ShareFS;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,WinSpool,WinApi.windows;

  function AddShareFS(namePC,MyUser,MyPasswd,Path,Name,
  Descrpt,pass:string;TypeFS:integer;Access:Olevariant):boolean;   //// создать сетевой ресурс с дескриптором безопасности
  function DeleteShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;   /// удалить сетевой ресурс
  function GetAccessMaskShareFS(namePC,MyUser,MyPasswd,ShareName:string):boolean;  /// запрос прав
  function AddShareNoFS(namePC,MyUser,MyPasswd,Path,Name,Descrpt:string):boolean; /// функция добавления шары без дескриптора безопасности
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
frmDomainInfo.memo1.Lines.Add(MyPS+'Создание сетевого ресурса '+Name+' :'+SysErrorMessage(MyError));
VariantClear(objShare);
VariantClear(FSWbemLocator);
VariantClear(FWMIService);
OleUnInitialize;
result:=true;
except
on E:Exception do
         begin
          result:=false;
          frmDomainInfo.memo1.Lines.Add(MyPS+'Ошибка создания сетевого ресурса '+name+' :"'+E.Message+'"');
          OleUnInitialize;
          exit;
         end;
end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////
function AddShareFS(namePC,MyUser,MyPasswd,Path,Name,                   /// функция добавления шары с использованием дескриптора безопасности
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
frmDomainInfo.memo1.Lines.Add(MyPS+'Создание сетевого ресурса '+Name+' :'+SysErrorMessage(MyError));
VariantClear(objShare);
VariantClear(FSWbemLocator);
VariantClear(FWMIService);
OleUnInitialize;
result:=true;
except
on E:Exception do
         begin
           result:=false;
           frmDomainInfo.memo1.Lines.Add(MyPS+'Ошибка создания сетевого ресурса '+name+' :"'+E.Message+'"');
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
if oEnum.Next(1, objShare, iValue) = 0 then     //// Пользователь
begin
MyError:=objShare.delete();
frmDomainInfo.memo1.Lines.Add('Отключение сетевого ресурса '+'"'+string(objShare.name)+'" : '+SysErrorMessage(MyError));
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
           frmDomainInfo.memo1.Lines.Add(MyPS+'Ошибка отключения сетевого ресурса "'+E.Message+'"');
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
         if oEnum.Next(1, objShare, iValue) = 0 then     //// Пользователь
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
           frmDomainInfo.memo1.Lines.Add(MyPS+'Ошибка GetAccessMask "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;
end;
end;
end.
