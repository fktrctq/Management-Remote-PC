unit Unit9;

interface

uses
  System.Classes,Registry,WinApi.windows,System.SysUtils
  ,ActiveX,ComObj,Variants,Vcl.Graphics, Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,WinSock,DateUtils,Math;

type
  Lcheck = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;

  end;
const
wbemFlagForwardOnly = $00000020;
function SendARP(DestIp: DWORD; srcIP: DWORD; pMacAddr: pointer; PhyAddrLen: Pointer): DWORD;stdcall; external 'iphlpapi.dll';

implementation

uses SettingsProgramForm,umain,Unit10,MainPopup,TaskEdit;
 var
  //RootPatch: HKEY; //=HKEY_LOCAL_MACHINE;  //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
  nstring:string;
  RootPatch: HKEY;
/////////////////////////////////////// ARP
function MySendARP(const MyIPAddress: String): String;
var
  DestIP: ULONG;
  MacAddr: Array [0..5] of Byte;
  MacAddrLen: ULONG;
  SendArpResult: Cardinal;
begin
  DestIP := inet_addr(PAnsiChar(AnsiString(MyIPAddress)));
  MacAddrLen := Length(MacAddr);
  SendArpResult := SendARP(DestIP, 0, @MacAddr, @MacAddrLen);

  if SendArpResult = NO_ERROR then
    Result := Format('%2.2X-%2.2X-%2.2X-%2.2X-%2.2X-%2.2X',
                     [MacAddr[0], MacAddr[1], MacAddr[2],
                      MacAddr[3], MacAddr[4], MacAddr[5]])
  else
    Result := '';
end;
/////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////
function Log_write(fname, text:string):string;
var f:TStringList;
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
  exit;
end;
end;
/////////////////////////////////////////////////////////////////////////
function ossn(s:string):string;
var
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
begin
try
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',Form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
FWbemObjectSet:= FWMIService.ExecQuery('SELECT SerialNumber '
+' FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
    begin
      try
      if FWbemObject.SerialNumber<>null then result:=trim(string(FWbemObject.SerialNumber));
      except
      result:=MySendARP(activeserver);
      end;
      FWbemObject:=Unassigned;
    end;
 Except
  on E:Exception do
     begin
     VariantClear(FWbemObject);
     VariantClear(FWbemObjectSet);
     VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
     end;
   end;
   VariantClear(FWbemObject);
   oEnum:=nil;
   VariantClear(FWbemObjectSet);
   VariantClear(FWMIService);
   VariantClear(FSWbemLocator);
end;
/////////////////////////////////////////////////////////////////////
function biossn(s:string):string; /// получение данных о материнской плате
var
i:integer;
sn:string;
res:bool;
oEnumSNM               : IEnumvariant;
iValueSNM              : LongWord;
FWbemObjectSetSNM      : OLEVariant;
FWbemObjectSNM         : OLEVariant;
 FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
begin
try
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',Form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
      FWbemObjectSetSNM:= FWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BIOS','WQL',wbemFlagForwardOnly);
       oEnumSNM:= IUnknown(FWbemObjectSetSNM._NewEnum) as IEnumVariant;
         while oEnumSNM.Next(1, FWbemObjectSNM, iValueSNM) = 0 do
            begin
              try
              if FWbemObjectSNM.SerialNumber<>null then
               sn:=trim(string(FWbemObjectSNM.SerialNumber));
              except
               sn:=MySendARP(activeserver);
              end;
          FWbemObjectSNM:=Unassigned;
            end;
         VariantClear(FWbemObjectSNM);
         oEnumSNM:=nil;
         VariantClear(FWbemObjectSetSNM);
        result:=sn;
        VariantClear(FWMIService);
         VariantClear(FSWbemLocator);
     except
      on E:Exception do
        begin
        Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/'+s+' - Ошибка Md-'+e.Message);
         VariantClear(FWbemObjectSNM);
         oEnumSNM:=nil;
         VariantClear(FWbemObjectSetSNM);
         VariantClear(FWMIService);
         VariantClear(FSWbemLocator);
        end;
      end;
end;
///////////////////////////////////////////////////////////////////////
function monterboardsn(s:string):string; ///
var
i:integer;
sn:string;
res:bool;
oEnumSNM               : IEnumvariant;
iValueSNM              : LongWord;
FWbemObjectSetSNM      : OLEVariant;
FWbemObjectSNM         : OLEVariant;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
begin
try
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',Form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
      FWbemObjectSetSNM:= FWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BaseBoard','WQL',wbemFlagForwardOnly);
       oEnumSNM:= IUnknown(FWbemObjectSetSNM._NewEnum) as IEnumVariant;
         while oEnumSNM.Next(1, FWbemObjectSNM, iValueSNM) = 0 do     //// мАТЕРИНСКАЯ ПЛАТА
            begin
              try
              if FWbemObjectSNM.SerialNumber<>null then
               sn:=trim(string(FWbemObjectSNM.SerialNumber));
              except
               sn:=MySendARP(activeserver);
              end;
          FWbemObjectSNM:=Unassigned;
            end;
        result:=sn;
     except
      on E:Exception do
        begin
         VariantClear(FWbemObjectSNM);
         VariantClear(FWbemObjectSetSNM);
         VariantClear(FWMIService);
         VariantClear(FSWbemLocator);
        end;
      end;
         VariantClear(FWbemObjectSNM);
         oEnumSNM:=nil;
         VariantClear(FWbemObjectSetSNM);
         VariantClear(FWMIService);
         VariantClear(FSWbemLocator);
end;
////////////////////////////////////////////////////////////
function snhddPhisical(idHdd:string):string;
var
FWbemObjectSetsn  : OLEVariant;
oEnumsn           : IEnumvariant;
FWbemObjectsn     : OLEVariant;
iValuesn          : LongWord;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
begin
try
idHdd:=copy(idHdd,5,length(idHdd));
Result := '';
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',Form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
FWbemObjectSetsn:= FWMIService.ExecQuery
('SELECT * FROM Win32_PhysicalMedia WHERE Tag LIKE "%'+idHdd+'%"', 'WQL',wbemFlagForwardOnly);
oEnumsn:= IUnknown(FWbemObjectSetsn._NewEnum) as IEnumVariant;
while oEnumsn.Next(1, FWbemObjectsn, iValueSN) = 0 do
  begin
  if FWbemObjectsn.SerialNumber<>null then result:=trim(vartostr(FWbemObjectsn.SerialNumber));
   FWbemObjectsn:=Unassigned;
  end;
oEnumsn:=nil;
VariantClear(FWbemObjectsn);
VariantClear(FWbemObjectSetsn);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
except
on e:Exception do
begin
VariantClear(FWbemObjectsn);
VariantClear(FWbemObjectSetsn);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - ошибка DSN - '+e.Message);
end;
end;
end;
///////////////////////////////////////////////////////////
function ScanHDD(s:string):string;
var
hddsn:string;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
begin
try
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',Form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * From Win32_DiskDrive WHERE MediaType=''Fixed hard disk media''','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// HDD
            begin
            if FWbemObject.DeviceID<>null then
            begin
            hddsn:=hddsn+snhddPhisical(vartostr(FWbemObject.DeviceID))+'++';
            FWbemObject:=Unassigned;
            if Length(hddsn)>=10 then break;
            end;
            end;
       result:=(hddsn);
  Except
  on E:Exception do
    begin
    Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - ошибка HD - '+e.Message);
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    result:=MySendARP(activeserver);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    end;
   end;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
   VariantClear(FWMIService);
   VariantClear(FSWbemLocator);
end;
////////////////////////////////////////////////////////////////////////////////// DaysBetween Возвращает количество полных дней из промежутка времени, заданного двумя значениями TDateTime.
procedure Code(var text: WideString; password: widestring;
decode: boolean);
var
i, PasswordLength: integer;
sign: shortint;
begin
PasswordLength := length(password);
if PasswordLength = 0 then Exit;
if decode then sign := -1
else sign := 1;
for i := 1 to Length(text) do
text[i] := chr(ord(text[i]) + sign *
ord(password[i mod PasswordLength + 1]));
end;

function Crypt(varStr: WideString):WideString;
var
 k: integer;
 s: WideString;
begin
   RandSeed:=100;
   s:=varStr;
   for k:=1 to Length(s) do
    s[k]:=Chr(ord(s[k]) xor (Random(127)+1));
 Crypt:=s;
end;
/////////////////////////////////////////////////////////////
procedure findComponent; ///frmDomainInfo.findElement
var
i:integer;
CurentNamePC:string;
begin
try
frmDomainInfo.Caption:= captionMainForms;
EditTask.Button7.OnClick:= EditTask.SaveTask;
{if frmDomainInfo.ListView8.items.count>30 then
begin
  for I := frmDomainInfo.ListView8.items.count-1 downto 30 do
   frmDomainInfo.ListView8.Items[i].Delete;
end;
frmDomainInfo.ListView8.OnChange:=frmDomainInfo.ListViewAllChange;
///////////////////////////////////////////////////////////
CurentNamePC:=frmDomainInfo.ComboBox2.Text; /// чтобы имя пк в cjmbobox не менялся    , запоминаем, потом обратно вставим

for I :=  frmDomainInfo.ComboBox2.Items.Count-1 downto 0 do  //// очистка combobox
frmDomainInfo.ComboBox2.Items.Delete(i);

for I := 0 to frmDomainInfo.ListView8.items.count-1 do      //// заполнение combobox
begin
frmDomainInfo.ComboBox2.Items.add(frmDomainInfo.ListView8.Items[i].SubItems[0]);
frmDomainInfo.ComboBox2.ItemsEx[i].ImageIndex:=0;
end;
frmDomainInfo.SpeedButton85.OnClick:=frmDomainInfo.ConnectRDPOsherWindows;
frmDomainInfo.RDP1.OnClick:=frmDomainInfo.RDP2Click;
frmDomainInfo.ComboBox2.Text:=CurentNamePC;  /// а вот тут вставляем еэто имя
frmDomainInfo.ComboBox2.OnClick:=frmDomainInfo.ComboBoxAllClick;
frmDomainInfo.ComboBox2.OnSelect:=frmDomainInfo.ComboBox2Select2;
frmDomainInfo.ComboBox2.OnKeyUp:=frmDomainInfo.ComboBox2KeyUp2;
frmDomainInfo.ComboBox1.Tag:=1;
frmDomainInfo.ComboBox1.enabled:=false;
frmDomainInfo.ComboBox8.enabled:=false; }
Except
  on E:Exception do
    begin
    Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - ошибка 2 - '+e.Message);
    end;
   end;
end;
////////////////////////////////////////////////////////////
function selectPro(i:integer):bool;
begin
  case i of
  1: frmDomainInfo.findElement;
  2: findComponent;
  3: frmDomainInfo.findElement;
  else findComponent;
  end;
end;
////////////////////////////////////////////////////////////////
function gandon(v,g:widestring):string;
var
i,z:integer;
a,b:Widestring;
s,y,u,x:string;
begin
try
sleep(2000);
a:=v;
b:=g;
s:='';
a:=trim(a);
b:=trim(b);
if Length(a)>Length(b) then
  begin
  z:=Length(a);
  for I := Length(b) to z do b:=b+'+';
  end
else
  begin
  z:=Length(b);
  for I := Length(a) to z-1 do  a:=a+'*';
  end;

sleep(1000);
for I := 1 to z do
  begin
  s:=s+inttostr((ord(a[i]))+i)+inttostr((ord(b[i]))+i);
 // x:=x+inttostr(Length(inttostr(ord(a[i]))))+inttostr(Length(inttostr(ord(b[i]))))
  x:=x+inttostr((((ord(a[i]))+i))and((ord(b[i]))-i));
  end;
 y:='';
 u:='';
sleep(1000);
for I := 1 to Length(x) do
  begin
  if x[i] in ['0','1','2','3','4','5','6','7','8','9'] then
    begin
      if (i mod 2)=0 then y:=y+X[i]
      else u:=u+X[i];
    end;
  end;
if Length(y)>Length(u) then z:=Length(u)
else z:=Length(y);
a:='';
sleep(1000);
for I := 1 to z do
  begin
  a:=a+inttostr((strtoint(y[i])+i and strtoint(u[i])-i))
  end;
 result:=a;
 if (a=keyboardstring) or (a=peremennaya) then
 begin
  CloseHandle(HM); /// убиваем мьютекс для разрешения копий
  frmDomainInfo.Caption:=captionMainFormsPro;
 end
   else selectPro(RandomRange(1,5));

     except
      on E:Exception do
        begin
         result:=x+'7458912354844';
         Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - ошибка 1 - '+e.Message);
        end;
      end;
 end;


function gandon2 (v,q:widestring):string;
var
i,z:integer;
a,b:widestring;
s,y,x:string;
begin
a:=v;
b:=q;
s:='';
a:=trim(a);
b:=trim(b);
if Length(a)>Length(b) then
  begin
  z:=Length(a);
  for I := Length(b) to z do b:=b+'+';
  end
else
 begin
 z:=Length(b);
 for I := Length(a) to z-1 do  a:=a+'*';
 end;
for I := 1 to z do
  begin
  s:=s+inttostr((ord(a[i]))+i)+inttostr((ord(b[i]))+i);
  x:=x+inttostr(Length(inttostr(ord(a[i]))))+inttostr(Length(inttostr(ord(b[i]))))
  end;
 y:=inttostr(z);
 result:=s+inttostr(z)+inttostr(Length(y));
end;

function gandon3(r:widestring):string;
var
i,z,n,l,p:integer;
a,b:widestring;
s:widestring;
y:widestring;
begin
try
s:=r;
n:=strtoint(copy(s,length(s),1));
z:=(strtoint(copy(s,length(s)-n,n)));
l:= (length(s) div 2) div z;
s:=copy(s,1,length(s)-1-length(inttostr(z)));
for I := 1 to z do
begin
a:=a+chr((strtoint(copy(s,1,2)))-i);
delete(s,1,2);
b:=b+chr((strtoint(copy(s,1,2)))-i);
delete(s,1,2);
end;
y:=a;
code(y,b,true);
b:=Crypt(b);
sleep(1000);
result:=gandon(y,b);
except
      on E:Exception do
        begin
         result:=a+b;
         end;
      end;
end;








function existreg(patch:string):bool;  //// существование пути в реестре
var
regFile:TregInifile;
begin
regFile:=Treginifile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then result:=true
else result :=false;
end;

function regeditwrite(patch,Section,Value:string):boolean; /// CurKey - подсекция в реестре
var      //// Запись в реестр
regFile:TregIniFile;
begin
RegFile:=TregIniFile.Create(KEY_WRITE OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch; //HKEY_LOCAL_MACHINE или HKEY_CURRENT_USER
if regFile.CreateKey(patch) then
RegFile.WriteString(patch,Section,value);
if Assigned(RegFile) then freeandnil(regFile);
end;

function regeditread(patch,Section:string):string;
var      /// чтение из реестра
regFile:TregIniFile;
begin
RegFile:=TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;  //HKEY_LOCAL_MACHINE или HKEY_CURRENT_USER
if regFile.KeyExists(patch) then
begin
result:=(regFile.ReadString(patch,Section,''));
end;
if Assigned(RegFile) then freeandnil(regFile);
end;

function writestring(S,patch,Section:WideString):bool;
begin
code(s,'1234',true);
s:=crypt(s);
regeditwrite(Patch,section,s);  ///patch -  'software\MRPC\MRPC'
end;


function readstring(patch,section:widestring):widestring; /// patch -  'software\MRPC\MRPC'
var
s:widestring;
begin
s:=regeditread(patch,section);
s:=crypt(s);
code(s,'1234',false);
result:=s;
end;

function gandondt(s:string):string;
begin
result:=inttostr(RandomRange(strtoint(s),60));
end;

function bd(s:string):bool;
var
DB: File of byte;
i:integer;
begin
i:=3000;
if  FileExists(extractfilepath(application.ExeName)+'\db.fdb') then
 begin
 AssignFile(db, extractfilepath(application.ExeName)+'\db.fdb');
 Reset(db);
 i:=(Filesize(db))div 1000;
 CloseFile(db)
 end;
///Log_write('Lic',datetostr(date)+'/'+timetostr(time)+'/'+s+' - size-'+inttostr(i));
if i<=2000 then result:=true else result:=false;
end;

procedure rewte;
var
msn,hsn,masn,alsn,i,os:string;
begin
i:=(gandondt(inttostr(randomrange(30,60))));
if not existreg(stringpatch) then
begin
msn:=monterboardsn(activeserver);
writestring(i+msn,stringpatch,'msn');
hsn:=ScanHDD(activeserver);
writestring(i+hsn,stringpatch,'hsn');
masn:=MySendARP(activeserver);
writestring(i+masn,stringpatch,'masn');
alsn:=(msn+hsn+masn);
writestring(i+alsn,stringpatch,'alsn');
writestring(i+datetostr(date),stringpatch,'dt');
os:=ossn(activeserver);
writestring(i+os,stringpatch,'osn');
end;
//////////////////////////////////////////
end;

function writedate(s:string):string;
begin

end;

function findsimvolS3(s:string):string;
begin
if (s='') or (pos('123456789',s)<>0)
 or ((pos(Ansiuppercase('string'),Ansiuppercase(s))<>0)
 or (pos(Ansiuppercase('Default'),Ansiuppercase(s))<>0)
 or (pos(Ansiuppercase('O.E.M'),Ansiuppercase(s))<>0)
 or (pos(Ansiuppercase('System'),Ansiuppercase(s))<>0)
 or (pos(Ansiuppercase('Number'),Ansiuppercase(s))<>0)
 or (pos(Ansiuppercase('Manuf'),Ansiuppercase(s))<>0)
 ) then result:='9512364857592'
 else result:=s;
 end;

function readdate(s:string):bool;
var
//a,b:integer;
y,x,z,u,s1,s2,s3,s4,s5:string;
begin
try
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
y:=readstring(stringpatch,'dt');
z:=copy(y,0,2);
delete(y,1,2);
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
x:=readstring(stringpatch,'dt');
u:=copy(x,0,2);
delete(x,1,2);
//a:=DaysBetween(date,strtodate(y));
//b:=DaysBetween(date,strtodate(x));
//if (a>=strtoint(z)) or (b>=strtoint(u)) then tratataq
//else tratataq;
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
s1:=monterboardsn(activeserver);
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
s2:=ScanHDD(activeserver);
s3:=biossn(activeserver);

  s3:=findsimvolS3(s3);
  if (s3='9512364857592')and(s1=findsimvolS3(s1)) then s3:=s1;
  if s1<>findsimvolS3(s1) then s1:=s3;
  if length(s1)>20 then s1:=Copy(s1,1,20);;
  if (Length(s2)>20) then s2:=Copy(s2,1,20);
  if (s2='') and (s1<>'') then s2:=s1;
  if (s2='') or (Length(s2)<=5) then  s2:=s3;
  if (s2='') or (Length(s2)<=5) then  s2:=ossn(activeserver);
  if length(s2)>20 then s2:=Copy(s2,1,20);
  if length(s3)>20 then s3:=Copy(s3,1,20);
  if s2='' then s2:='12365489523158';
  if s3='' then s3:='85296374145615';

s5:=gandon2(Ansiuppercase(s2),Ansiuppercase(s3));
s4:=gandon3(s5);
result:=true;
 except
      on E:Exception do
        begin
         result:=false;
         Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - Ошибка -'+e.Message);
        end;
      end;
end;



procedure Lcheck.Execute;
begin
try
OleInitialize(nil);
RootPatch:=HKEY_CURRENT_USER;  //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
rewte;
nstring:=readstring(stringpatch,'alsn');
readdate(stringpatch);

 except
      on E:Exception do
        begin
        //frmDomainInfo.tratataq;
        Log_write('Other',datetostr(date)+'/'+timetostr(time)+' - Thread 2-'+e.Message);
        end;
      end;
OleUnInitialize();
end;



end.
