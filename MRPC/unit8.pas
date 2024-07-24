unit unit8;

interface

uses
  System.Classes,Registry,WinApi.windows,System.SysUtils
  ,ActiveX,ComObj,Variants,Vcl.Forms,WinSock,DateUtils,Math;
procedure Callic;

const
wbemFlagForwardOnly = $00000020;
function SendARP(DestIp: DWORD; srcIP: DWORD; pMacAddr: pointer; PhyAddrLen: Pointer): DWORD;stdcall; external 'iphlpapi.dll';

implementation
uses SettingsProgramForm,Unit10;
 var
  //RootPatch: HKEY; //=HKEY_LOCAL_MACHINE;  //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  RootPatch           : HKEY;
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
////////////////////////////////////////////////////////////
function Log_write(fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
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
iValue             : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT SerialNumber '
+' FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Операционная система
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
     Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/'+s+' - ошибка OS - '+e.Message);
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet)
     end;
   end;
   VariantClear(FWbemObject);
   oEnum:=nil;
   VariantClear(FWbemObjectSet);
end;
/////////////////////////////////////////////////////////////////////
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
begin
try
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
     except
      on E:Exception do
        begin
        Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/'+s+' - Ошибка Md-'+e.Message);
         VariantClear(FWbemObjectSNM);
         oEnumSNM:=nil;
         VariantClear(FWbemObjectSetSNM);
        end;
      end;
end;
/////////////////////////////////////////////////////////////////////
function monterboardsn(s:string):string; /// получение данных о материнской плате
var
i:integer;
sn:string;
res:bool;
oEnumSNM               : IEnumvariant;
iValueSNM              : LongWord;
FWbemObjectSetSNM      : OLEVariant;
FWbemObjectSNM         : OLEVariant;
begin
try
      FWbemObjectSetSNM:= FWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BaseBoard','WQL',wbemFlagForwardOnly);
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
     except
      on E:Exception do
        begin
        Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/'+s+' - Ошибка Md-'+e.Message);
         VariantClear(FWbemObjectSNM);
         oEnumSNM:=nil;
         VariantClear(FWbemObjectSetSNM);
        end;
      end;
end;
////////////////////////////////////////////////////////////
function snhddPhisical(idHdd:string):string; ///// функция извлечения serial number
var
FWbemObjectSetsn  : OLEVariant;
oEnumsn           : IEnumvariant;
FWbemObjectsn     : OLEVariant;
iValuesn          : LongWord;
begin
try
idHdd:=copy(idHdd,5,length(idHdd));
Result := '';
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
except
on e:Exception do
begin
Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - ошибка DSN - '+e.Message);
end;
end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////
function Sjesk(s:string):string;
var
hddsn:string;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
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
       result:=hddsn;
  Except
  on E:Exception do
    begin
    Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/'+s+' - Ошибка DD - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    end;
   end;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
end;
//////////////////////////////////////////////////////////////////////////////////
function gandon2 (v,q:widestring):string;
var
i,z:integer;
a,b:widestring;
s,y,x:string;
begin
a:=(v);
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
/////////////////////////////////////////////////////////////////////////////////




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
y,x,z,u,s1,s2,s3,s4:string;
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
//if (a>=strtoint(z)) or (b>=strtoint(u)) then result:=false
//else result:=true;
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
s1:=monterboardsn(activeserver);
RootPatch:=HKEY_CURRENT_USER; //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
s2:=Sjesk(activeserver);
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

Form10.Memo1.Clear;
Form10.Memo1.lines.add(gandon2(Ansiuppercase(s2),Ansiuppercase(s3)));
 except
      on E:Exception do
        begin
         result:=false;
         Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - '+e.Message);
        end;
      end;
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
if i<=2000 then result:=true else result:=false;
end;

procedure rewte;
var
msn,hsn,masn,alsn,i:string;
begin
i:=(gandondt(inttostr(randomrange(30,60))));
if not existreg(stringpatch) then
begin
msn:=monterboardsn(activeserver);
writestring(i+msn,stringpatch,'msn');
hsn:=Sjesk(activeserver);
writestring(i+hsn,stringpatch,'hsn');
masn:=MySendARP(activeserver);
writestring(i+masn,stringpatch,'masn');
alsn:=(msn+hsn+masn);
writestring(i+alsn,stringpatch,'alsn');
writestring(i+datetostr(date),stringpatch,'dt');
end;
//////////////////////////////////////////
end;

procedure CalLic;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(activeserver, 'root\CIMV2',form10.Edit4.Text,Form10.Edit5.Text); ///WbemUser, WbemPassword
readdate(activeserver);
RootPatch:=HKEY_CURRENT_USER;  //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER
rewte;
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
 except
      on E:Exception do
        begin
         rewte;
         Log_write('Other',datetostr(date)+'/'+timetostr(time)+'/ - '+e.Message);
        end;
      end;
OleUnInitialize();
end;
end.
