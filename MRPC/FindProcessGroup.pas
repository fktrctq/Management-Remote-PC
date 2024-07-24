unit FindProcessGroup;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  GroupFindProcess = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,FormFindProcess;

var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  Myidicmpclient:TIdIcmpClient;
  iValue        : LongWord;

  {ping}
function ping(s:string):string;
var
z:integer;
begin
try
result:='';
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
result:='Превышен интервал ожидания запроса'
else result:='1'; ///доступен
except
result:='Узел не доступен';
end;
end;

procedure FindProcessForPC;
var ///////////////////////////////////// завершение процесса на группе машин
YesProc:boolean;
i,z:integer;
ListPCFind:Tstringlist;
/////////////////////////////////
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
begin
/////////////////////////////
YesProc:=false;
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
///////////////////////////////
ListPCFind:=TstringList.Create;
ListPCFind:=ListFindProcess;
for I := 0 to ListPCFind.Count-1 do
if (ping(ListPCFind[i])<>'1') then
  begin
  finditem('Узел не доступен',ListPCFind[i],2);
  end
else
begin  ////////////////////////////////////////////// завершение процесса на группе машин
try
YesProc:=false;
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(ListPCFind[i], 'root\CIMV2', MyUser, MyPasswd);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"'+ProcessNameForFind+'%"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     ////
  begin
  YesProc:=true;
  VariantClear(FWbemObject);
  end;
if YesProc=true then
  begin
  finditem('Процесс  найден',ListPCFind[i],1);
  end;
if YesProc=false then
  begin
  finditem('Процесс не найден',ListPCFind[i],2);
  end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
except
on E:Exception do
  begin
  finditem('Ошибка поиска процесса',ListPCFind[i],2);
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  end;
end;
end; /////////////////////////////////////////////// завершение процесса на на группе машин
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
ListPCFind.Free;
ListFindProcess.Free;
Myidicmpclient.Free;
OleUnInitialize;
end;


procedure GroupFindProcess.Execute;
begin
  FindProcessForPC;
end;


end.
