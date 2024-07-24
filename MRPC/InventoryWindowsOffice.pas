unit InventoryWindowsOffice;

interface

uses
  System.Classes,ActiveX,ComObj,Variants,FireDAC.Stan.Intf, FireDAC.Stan.Option
  , FireDAC.Stan.Param,System.SysUtils,Vcl.Forms,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,FireDAC.UI.Intf
  ,FireDAC.Comp.UI,FireDAC.Phys.FB,inifiles,SqlTimSt,IdIcmpClient,System.StrUtils,
  WinSock,Windows,System.TimeSpan,IdTCPClient,ShellAPI;

type
  InventoryMicrosoft = class(TThread)
  private
  function ScanMicrosoftWO(User,Pass:string;ListPC:TstringList):boolean; /// ����� ��� ������������ ��������� Microsoft
  protected
    procedure Execute; override;
  end;

implementation
uses umain,PingForInventoryMicrosoft;



ThreadVar
ListPCWindowsOffice:tstringList;
UserWO,PassWO:string;


function InventoryMicrosoft.ScanMicrosoftWO(User,Pass:string;ListPC:TstringList):boolean; /// ����� ��� ������������ ��������� Microsoft
var
  CurrentPC,resIP:string;
  i,CoutPCOK:integer;
  WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptWin,CurrentOfficeLic:string;
  OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList;
  sOfficeProduct,sStatOfLic,sKeyOfc,sTypelicOffice,sDescriptLicOffice,nameOS,VerOS:string;
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  ConnectionThread    : TFDConnection;
  TransactionWrite    : TFDTransaction;
  FDQueryWrite        : TFDQuery;
  KMS,ActivateYesNo,DelKeyKms,KMSUnknownStat:boolean; // kms - ������������ ��� ��� �� ����� ��������, ActivateYesNo - ������������ ��� ��� ��� ������� ������� ��������  �� ������� LIC_ERROR
  KMSpath,kmsfolder,KMSkey:string;
  KMSini:Tmeminifile;
  Deadline:integer;
///////////////////////////////////////////////////////////
/// ���� ������� ����������� ������� ��� ������///////////
function Log_write(level:integer;fname, text:string):Boolean;
var f:TStringList;        /// ������� ������ � ��� ����
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
  exit;
end;
end;

function SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
var
avalible:boolean;
begin
resIP:='';
  case PingType of
    1:  avalible:=PingIdIcmp(resIP,HostName,pingtimeout);
    2:  avalible:=PingGetaddrinfo(resIP,HostName,pingtimeout);
    3:  avalible:=PingGetHostByName(resIP,HostName,pingtimeout);
  END;
  if not avalible then frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:=HostName+' - �� ��������';
resIP:='';
Result:=avalible;
end;


//*********************************************************************************
/////////////////////////////////////////////////////////////////////////////////////////


function GetLastDir(path:String):String;  // ���������� �������� �����
var
i,j:integer;
begin
path:=ExtractFilePath(path); //���� �� ����� ������ �� path  �����
if Length(path)>3 then
    begin
      for I :=Length(path)-1 downto 1 do
      if path[i]='\' then
      begin
      j:=i+1;
      break;
      end;
      result:=Copy(path,j,Length(path)) ;
end
else result:=path[1]+path[3];
end;

function MyNewProcess(
    NamePC:string;     // ������ �����������
    UserName:string;   // �����
    PassWd:string;     // ������
    FileToRun:String   // ���� ��� �������
    ):boolean;
var
MyError:integer;
FWbemObject   : OLEVariant;
objProcess    : OLEVariant;
objConfig     : OLEVariant;
ProcessID,z,i   : Integer;
const
wbemFlagForwardOnly = $00000020;
HIDDEN_WINDOW       = 1;
/////////////////////////////////////////////////////////////////////
BEGIN
  try
  OleInitialize(nil);
  FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
  objConfig     := FWbemObject.SpawnInstance_;
  objConfig.ShowWindow := HIDDEN_WINDOW;
  objProcess    := FWMIService.Get('Win32_Process');
  MyError:=objProcess.Create(FileToRun, null, objConfig, (ProcessID));
  if MyError=0 then result:=true
  else
  begin
   result:=false;
   Log_write(0,'Licenses',CurrentPC+' - ������ �������� ��������� - '+SysErrorMessage(MyError));
  end;
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  objConfig:=Unassigned;
  OleUnInitialize;
    except
      on E:Exception do
      begin
      Log_write(1,'Licenses',CurrentPC+' - ������ ������� �������� ��������� - '+e.Message);
      result:=false;
      OleUnInitialize;
      end;
    end;
END;
///////////////////////////////////////////////////////////
 function FindAddcreateDir(path,NamePC:string):boolean;// �������� � �������� ����������
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //���� ��� ��������
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// ������� ���� ���� �� �����
     result:=false
     else result:=true; // ���������� �������
    end
    else result:=true; // ���������� ����������
  except on E: Exception do
     begin
     Log_write(1,'Licenses',CurrentPC+' - ������ �������� �������� - '+e.Message);
     result:=false;
     end;
   end;
end;
////////////////////////////////////////////////////////////
function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean):boolean; // �����������
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// ����� ������� ������� � vista , �������� ����� � �������� ��� ������ �����������
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //����� ������� ������������, ���� ���� �������� ���������� ������� ������������� ����, ������� �������� ������ � ����� ����� ��������������� ������
      lpszProgressTitle :=nil;            // ��������� �� ��������� ����������� ���� ���������
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // �� ���� �������� ���� �������� �����������
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // ���� ��������
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // ��� �������  ���� �������� ��������
      pTo := pchar('');   // ���� �������� �� ������������
      //finditem('','�������� �������� �������� ������������',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// ����������� ������������ �� ��������� �����
     //if not
     (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // ������� ������� �� ���� � ����
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken));
    // then GetLastError();
     except on E: Exception do
      begin
      result:=false;
      Log_write(1,'Licenses',CurrentPC+' - ������ ����������� - '+e.Message);
      end;
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC); //��������� � ������� ������� ���� ��� ���.
     rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/����������� /��������
     if rescopy=0 then result:=true
     else
       begin
       result:=false;
       Log_write(0,'Licenses',CurrentPC+' - ����������� ������: - '+SysErrorMessage(rescopy));
       end;
     CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
      Log_write(1,'Licenses',CurrentPC+' - ������ ������� �����������: '+E.Message);
      if TypeOperation=3 then
      Log_write(1,'Licenses',CurrentPC+' - ������ ������� ��������: '+E.Message);
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then Log_write(2,'Licenses',CurrentPC+' - ����� ������ ������� ����������� ����� �������� ��������: '+E.Message);
     if TypeOperation=3 then Log_write(2,'Licenses',CurrentPC+' - ����� ������ ������� �������� ������������ ����� ������� �������� : '+E.Message);
     result:=false;
    end;
  end;
end;

function runkms(namepc,user,pass:string):boolean;
begin
if CopyFFSelectPC
(namepc, // ��������� �� ������� ��������
user, // ������������
pass, // ������
kmsfolder+#0+#0,        // ��� ��������
'C$\TEMP\'+#0+#0,              // ���� ��������
frmdomaininfo,                       // ����� ��������
2,                            // �������� ����������
false,                       // ����������� ������
true) then                      // ��������� ������� �������� ����������
begin
if MyNewProcess(namepc, // ��������� �� ������� ��������
user, // ������������
pass, // ������
'C:\TEMP\'+GetLastDir(kmspath)+ExtractFileName(kmspath)+' '+kmskey) then //   /s /l - ����� � ������������ � ��������� ������������� �������� �������, ���� ������ ��� ������ ��� � ����� ������ �� ��������� �� �����
Log_write(0,'Licenses',CurrentPC+' - �������� ������� '+'C:\TEMP\'+GetLastDir(kmspath)+ExtractFileName(kmspath)+' '+kmskey);
end;
end;
//*********************************************************************************
function TypeLicinAndOffice(Des:string):string;
begin
try
if pos('OEM_DM',des)<>0 then result:='OEM_DM' else
if pos('OEM_COA_NSLP',des)<>0 then result:='OEM_COA_NSLP' else
if pos('OEM_SLP',des)<>0 then result:='OEM_SLP' else
if pos('OEM_COA_SLP',des)<>0 then result:='OEM_COA_SLP' else
if pos('OEM_NO_NSLP',des)<>0 then result:='OEM_NO_NSLP' else
if pos('RETAIL',des)<>0 then result:='RETAIL' else
if pos('KMSCLIENT',des)<>0 then result:='KMS' else
if pos('RETAIL(Free)',des)<>0 then result:='RETAIL(Free)' else
if pos('TIMEBASED_SUB',des)<>0 then result:='TIMEBASED_SUB' else
if pos('TIMEBASED_EVAL',des)<>0 then result:='TIMEBASED_EVAL' else
if pos('KMS_W7',des)<>0 then result:='KMS' else
if pos('KMS_R2_C',des)<>0 then result:='KMS' else
if pos('RETAIL(MAK)',des)<>0 then result:='RETAIL(MAK)' else
if pos('KMS_WS19',des)<>0 then result:='KMS' else
if pos('KMS_2012-R2',des)<>0 then result:='KMS' else
if pos('KMS_WS12_R2',des)<>0 then result:='KMS' else
if pos('KMS_WS16',des)<>0 then result:='KMS' else
if pos('VIRTUAL_MACHINE_ACTIVATION',des)<>0 then result:='VIRTUAL_MACHINE_ACTIVATION' else
if pos('KMS_W8.1',des)<>0 then result:='KMS' else
if pos('GVLK',des)<>0 then result:='GVLK' else
if pos('OEM',des)<>0 then result:='OEM' else
if pos('MAK',des)<>0 then result:='MAK' else
if pos('RETAIL',des)<>0 then result:='RETAIL' else
if pos('KMS',des)<>0 then result:='KMS'
else result:=des;
except
result:=des;
end;
end;

function CurentOfficeProduct(des:string):string;
begin
try
if pos(AnsiUpperCase('Word'),AnsiUpperCase(des))<>0 then result:='Word' else
if pos(AnsiUpperCase('Excel'),AnsiUpperCase(des))<>0 then result:='Excel' else
if pos(AnsiUpperCase('Outlook'),AnsiUpperCase(des))<>0 then result:='Outlook' else
if pos(AnsiUpperCase('PowerPoint'),AnsiUpperCase(des))<>0 then result:='PowerPoint' else
if pos(AnsiUpperCase('Access'),AnsiUpperCase(des))<>0 then result:='Access' else
if pos(AnsiUpperCase('InfoPath'),AnsiUpperCase(des))<>0 then result:='InfoPath' else
if pos(AnsiUpperCase('Publisher'),AnsiUpperCase(des))<>0 then result:='Publisher' else
if pos(AnsiUpperCase('Visio'),AnsiUpperCase(des))<>0 then result:='Visio' else
if pos(AnsiUpperCase('Project'),AnsiUpperCase(des))<>0 then result:='Project' else
if pos(AnsiUpperCase('Query'),AnsiUpperCase(des))<>0 then result:='Query' else
if pos(AnsiUpperCase('OneNote'),AnsiUpperCase(des))<>0 then result:='OneNote' else
if pos(AnsiUpperCase('Groove'),AnsiUpperCase(des))<>0 then result:='Groove' else
if pos(AnsiUpperCase('SharePoint'),AnsiUpperCase(des))<>0 then result:='SharePoint' else
if pos(AnsiUpperCase('Exchange'),AnsiUpperCase(des))<>0 then result:='Exchange' else
if pos(AnsiUpperCase('Forms'),AnsiUpperCase(des))<>0 then result:='Forms' else
if pos(AnsiUpperCase('Intune'),AnsiUpperCase(des))<>0 then result:='Intune' else
if pos(AnsiUpperCase('SkypeforBusiness'),AnsiUpperCase(des))<>0 then result:='Skype for Business' else
if pos(AnsiUpperCase('Lync'),AnsiUpperCase(des))<>0 then result:='Lync' else
result:='Office';
except
result:='Office';
end;
end;


function DeleteKeyKMS(namePC,user,pass:string):boolean;   ////// �������� ����� ��������
 var
  FSWbemLocator  : OLEVariant;
  FWMIService    : OLEVariant;
  FWbemObjectSet : OLEVariant;
  FWbemObject    : OLEVariant;
  oEnum          : IEnumvariant;
  iValue         : LongWord;
  step:string;
Begin
try
step:='0';
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', user, pass,'','',128);
step:='1';
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM SoftwareLicensingProduct','WQL',wbemFlagForwardOnly);
step:='2';
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
step:='3';
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
begin
step:='4';
if (FWbemObject.Description<>null)and(FWbemObject.LicenseStatus<>null) then
if (pos(AnsiUpperCase('KMS'),ansiuppercase(TypeLicinAndOffice(vartostr(FWbemObject.Description))))<>0)
and (integer(FWbemObject.LicenseStatus)<>0) then
  begin
  step:='5';
  FWbemObject.UninstallProductKey();  /// �������� ����� ��������
  Log_write(0,'Licenses',CurrentPC+' - �������� ����� �������� - '+vartostr(FWbemObject.Description));
  end;
step:='6';
FWbemObject:=Unassigned;
end;
result:=true;
except
on E:Exception do
  begin
  Log_write(1,'Licenses',CurrentPC+' - '+step+' ������ �������� ����� �������� - '+E.Message);
  result:=false;
  end;
end;
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
end;

function DescriptionLicStatus(num:string):string;
var   /// �������� ��� ������ � �� ������� �������� �������� ���� ������
z:integer;  //slui.exe 0x2a 0xC004FE00
he,s,A:string;
begin
try
a:=num;
if pos('OLE error ',num)<>0 then
begin
s:=copy(num,11,length(num));
a:=s;
end;
//https://support.microsoft.com/ru-ru/windows/%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BA%D0%B0-%D0%BF%D0%BE-%D0%BE%D1%88%D0%B8%D0%B1%D0%BA%D0%B0%D0%BC-%D0%B0%D0%BA%D1%82%D0%B8%D0%B2%D0%B0%D1%86%D0%B8%D0%B8-windows-09d8fb64-6768-4815-0c30-159fa7d89d85
if TryStrToInt(a,z) then
begin
he:=IntToHex(z,1);
end
else he:=a;
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='SELECT * FROM LIC_ERROR WHERE CODE='''+he+'''';
FDQueryWrite.Open;
result:=FDQueryWrite.FieldByName('DESCRIPTION').AsString;
ActivateYesNo:=FDQueryWrite.FieldByName('ACTIVATE').AsBoolean; // ������������ ��� ��� ��� ������� ������� ��������
TransactionWrite.Commit;
FDQueryWrite.Close;
if result='' then
begin // ����� ���������� ���� ��� � �������
result:=he;
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into LIC_ERROR (CODE,DESCRIPTION,ACTIVATE) VALUES (:p1,:p2,:p3) MATCHING (CODE,DESCRIPTION)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(he)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+('Unknown')+'';
FDQueryWrite.params.ParamByName('p3').AsBoolean:=false;
ActivateYesNo:=false; // ������������ ��� ��� ��� ������� ������� ��������
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
FDQueryWrite.Close;
end;
except on E: Exception do
begin
 if TransactionWrite.Active then begin  TransactionWrite.Rollback; FDQueryWrite.Close;  end;
 result:='Unknown';
 Log_write(1,'Licenses',CurrentPC+' - DescriptionLicStatus - '+e.Message);
end;
end;
end;

function ReadNameProductMicrosoft(sID:string):string;  // ������ ����� ��������
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='SELECT * FROM MICROSOFT_PRODUCT WHERE PRODUCT_ID='''+sID+'''';
FDQueryWrite.Open;
Result:=(FDQueryWrite.FieldByName('DESCRIPTION').AsString);
TransactionWrite.Commit;
FDQueryWrite.Close;
except on E: Exception do
  begin
  if TransactionWrite.Active then begin  TransactionWrite.Rollback; FDQueryWrite.Close;  end;
  result:='';
  Log_write(1,'Licenses',CurrentPC+' - ReadNameProductMicrosoft - '+e.Message);
  end;
end;
end;
///////////////////////////////////////////////////////////////////////////////////////
function WriteNameIDProductMicrosoft(sID,sName,KeyProduct,TypeLic,PartProdKey,Product:string):boolean;  // ������ ������ ID � ����� �������� � �����
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into MICROSOFT_PRODUCT '+
  ' (PRODUCT_ID,DESCRIPTION,KEY_PRODUCT,TYPE_LIC,PARTIALPRODUCTKEY,PRODUCT)'+
  ' VALUES (:p1,:p2,:p3,:p4,:p5,:p6) MATCHING (DESCRIPTION,KEY_PRODUCT,TYPE_LIC,PRODUCT_ID)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(sID)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+(sName)+'';
FDQueryWrite.params.ParamByName('p3').AsString:=''+(KeyProduct)+'';
FDQueryWrite.params.ParamByName('p4').AsString:=''+(TypeLic)+'';
FDQueryWrite.params.ParamByName('p5').AsString:=''+(PartProdKey)+'';
FDQueryWrite.params.ParamByName('p6').AsString:=''+(Product)+'';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
FDQueryWrite.Close;
result:=true;
except on E: Exception do
begin
 result:=false;
 if TransactionWrite.Active then begin  TransactionWrite.Rollback; FDQueryWrite.Close;  end;
 Log_write(1,'Licenses',CurrentPC+' - WriteNameIDProductMicrosoft - '+e.Message);
end;
end;
end;
//////////////////////////////////////////////////////////////////////////////////////
function keyOEM:string;   ///// Windows 10
var
FWbemObjectSet   : OLEVariant;
FWbemObject      : OLEVariant;
oEnum            : IEnumvariant;
iValue           : LongWord;
yes:boolean;
Begin
try
yes:=false;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT OA3xOriginalProductKey FROM SoftwareLicensingService');
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
begin
if FWbemObject.OA3xOriginalProductKey<>'' then result:=vartostr(FWbemObject.OA3xOriginalProductKey)
else result:='';
yes:=true;
FWbemObject:=Unassigned;
end;
if not yes then result:='';
except
 on E:Exception do
  begin
  Log_write(1,'Licenses',CurrentPC+' - KeyOEM - '+e.Message);
  result:='';
  end;
end;
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
end;
///////////////////////////////////////////////////////////////////////////////////////
Procedure ViewLicensingOffice(var OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList);
  var
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
  i:integer;
  ProdKeyID:string;
const
 LicStatus: array [0..6] of string =('��� ��������','�������� �������',
  '�������� ������','���������� ���������� �� ��������� �������','�������� ������ ��� ����������� ������','��������������','����������� �������� �����');
Begin
try
 i:=0;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,Description,ProductKeyID,LicenseStatus,PartialProductKey,ID,LicenseStatusReason FROM OfficeSoftwareProtectionProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
  begin
  if (FWbemObject.ID<>null)and (FWbemObject.Name<>null)and(integer(FWbemObject.LicenseStatus)<>0) then
    begin
    OfficeProduct.add(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
    if OfficeProduct[i]='' then
    begin
    OfficeProduct[i]:=(vartostr(FWbemObject.Name));
    end;
    if FWbemObject.LicenseStatus<>null then StatOfLic.add(LicStatus[(integer(FWbemObject.LicenseStatus))]) else StatOfLic.add('');
    if FWbemObject.PartialProductKey<>null then KeyOfc.add( vartostr(FWbemObject.PartialProductKey)) else KeyOfc.add('');
    if FWbemObject.LicenseStatusReason<>null then DescriptLicOffice.Add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason))) else DescriptLicOffice.Add('');
    if FWbemObject.Description<>null then  TypelicOffice.Add(TypeLicinAndOffice(vartostr(FWbemObject.Description))) else TypelicOffice.Add('');
    if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
    if CurentOfficeProduct(OfficeProduct[i])='Office' then CurrentOfficeLic:=StatOfLic[i]; // ���� ��� ����� Office �� ���������� ������ ��������
     WriteNameIDProductMicrosoft((vartostr(FWbemObject.ID)),OfficeProduct[i],KeyOfc[i],TypelicOffice[i],ProdKeyID,CurentOfficeProduct(OfficeProduct[i])); // ������ � ����� ������� ���������
    inc(i);
    end;
  FWbemObject:=Unassigned;
  end;
except
 on E:Exception do
  begin
   Log_write(1,'Licenses',CurrentPC+' - ViewLicensingOffice - '+e.Message);
  end;
end;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
oEnum:=nil;
End;

///////////////////////////////////////////////////////////////////////////////////////
procedure LicensingStatus(var WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptLicWin:string;OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice:TstringList);
  var
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
  i:integer;
  OfficeBool:boolean;
  OEMKey,ProdKeyID:string;
const
 LicStatus: array [0..6] of string =('��� ��������','�������� �������',
  '�������� ������','���������� ���������� �� ��������� �������','�������� ������ ��� ����������� ������','��������������','����������� �������� �����');

function Win7:boolean;
begin
if pos('6.1',VerOS)=1 then result:=true
else Result:=false;
end;
function OfficeBoll(s:string):boolean;
begin
if pos('Office',s)<>0 then Result:=true
else Result:=false;
end;
function WindowsBoll(s:string):boolean;
begin
if pos('Windows',s)<>0 then Result:=true
else Result:=false;
end;
Begin
try
i:=0;
OfficeBool:=false;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,Description,LicenseStatus,ProductKeyID,PartialProductKey,ID,LicenseStatusReason FROM SoftwareLicensingProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
  BEGIN
  if integer(FWbemObject.LicenseStatus)<>0 then
  BEGIN
   if WindowsBoll(vartostr(FWbemObject.Name)) then
    Begin
    if (FWbemObject.ID<>null) then
     begin
     WinProduct:=(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
     if WinProduct='' then WinProduct:=nameOS+' ('+VerOS+')';
     if FWbemObject.LicenseStatus<>null then StatWinLic:=LicStatus[(integer(FWbemObject.LicenseStatus))];
     if FWbemObject.PartialProductKey<>null then KeyWin:=vartostr(FWbemObject.PartialProductKey);
     if FWbemObject.LicenseStatusReason<>null then DescriptLicWin:=DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason));
     if FWbemObject.Description<>null then  TypelicWin:=TypeLicinAndOffice(vartostr(FWbemObject.Description));
     if pos('OEM',TypelicWin)<>0 then //���� ��� �������� OEM �� �������� �������� ���� �� UEFI
     begin
      OEMKey:=keyOEM;
      if Length(OEMKey)>5 then KeyWin:=OEMKey;
     end;
     if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
     WriteNameIDProductMicrosoft(vartostr(FWbemObject.ID),WinProduct,KeyWin,TypelicWin,ProdKeyID,'Windows'); // ������ � ����� ������� ���������
     end;
    End;
   if OfficeBoll(vartostr(FWbemObject.Name)) then
   Begin
    if (FWbemObject.ID<>null) then
    begin
     OfficeProduct.add(ReadNameProductMicrosoft(vartostr(FWbemObject.ID)));
      if OfficeProduct[i]='' then
      begin
       OfficeProduct[i]:=(vartostr(FWbemObject.Name));
      end;
    if FWbemObject.LicenseStatus<>null then StatOfLic.add(LicStatus[(integer(FWbemObject.LicenseStatus))]);
    if FWbemObject.PartialProductKey<>null then KeyOfc.add( vartostr(FWbemObject.PartialProductKey));
    if FWbemObject.LicenseStatusReason<>null then DescriptLicOffice.Add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason))) else DescriptLicOffice.Add('');
    if FWbemObject.Description<>null then  TypelicOffice.Add(TypeLicinAndOffice(vartostr(FWbemObject.Description))) else TypelicOffice.Add('');
    if FWbemObject.ProductKeyID<> null then ProdKeyID:=vartostr(FWbemObject.ProductKeyID) else ProdKeyID:='';
    if CurentOfficeProduct(OfficeProduct[i])='Office' then CurrentOfficeLic:=StatOfLic[i]; // ���� ��� ����� Office �� ���������� ������ ��������
     WriteNameIDProductMicrosoft((vartostr(FWbemObject.ID)),OfficeProduct[i],KeyOfc[i],TypelicOffice[i],ProdKeyID,CurentOfficeProduct(OfficeProduct[i])); // ������ � ����� ������� ���������
    inc(i);
    end;
    OfficeBool:=true;
   End;
 END;
 FWbemObject:=Unassigned;
 END;

if not OfficeBool then
begin
OfficeProduct.Clear;
StatOfLic.Clear;
KeyOfc.Clear;
ViewLicensingOffice(OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice);
end;

except
 on E:Exception do
  begin
  Log_write(1,'Licenses',CurrentPC+' - LicensingStatus - '+e.Message);
  end;
end;
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
End;

Procedure osconfig(var nameOS,VerOS:string);
var
FWbemObjectSet                          : OLEVariant;
FWbemObject                             : OLEVariant;
oEnum                                   : IEnumvariant;
iValue        : LongWord;
begin
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,Version FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������������ �������
    begin
      try
      if FWbemObject.Caption<>null then
      begin
       nameOS:=trim(string(FWbemObject.Caption));
      end;
      except
      nameOS:='unknown';
      end;
      try
      if FWbemObject.version<>null then   VerOS:=trim(string(FWbemObject.version));
      except
      VerOS:= 'unknown';
      end;
      FWbemObject:=Unassigned;
    end;
 Except
  on E:Exception do
     begin
     Log_write(1,'Licenses',CurrentPC+' - OSconfig - '+e.Message);
     end;
  end;
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
end;

function ConnectWMI(NamePC,User,Pass:string): Boolean;
begin
 try
    FSWbemLocator:=Unassigned;
    FWMIService:=Unassigned;
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128); ///WbemUser, WbemPassword
    FWMIService.Security_.impersonationlevel:=3;
    FWMIService.Security_.authenticationLevel := 6;
    Result:=True;
 except
   on E:Exception do
    begin
    result:=false;
    FWMIService:=Unassigned;
    FSWbemLocator:=Unassigned;
    Log_write(1,'Licenses',NamePC+' - '+e.Message);
   end;
 end;
end;
//////////////////////////////// ���� ���������� �������� ���� ������ �� �������� ����� ������� ����������� �������
BEGIN
try
Log_write(0,'Licenses',CurrentPC+'������ ��������������');
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // ������������ ���� ������
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.Connected:=true;
ConnectionThRead.LoginPrompt:= false;  /// ����������� ������� user password
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThRead;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWrite.Options.AutoCommit:=false;
TransactionWrite.Options.AutoStart:=false;
TransactionWrite.Options.AutoStop:=false;
FDQueryWrite:=TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=ConnectionThRead;


if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
begin
{KMS:boolean; KMSApath,KMSkey:string; KMSini}
KMSini:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
KMSUnknownStat:=KMSini.ReadBool('Vaccina','UnknownStatus',false); // ������������ ��� ������� "����������", �.�. ���� ����� ������ ���
KMS:=KMSini.ReadBool('Vaccina','Vaccination',false); // ������������ ��� �������� ��������� � �������
Deadline:=4; // �������� ������� ������ �� ������ ��� ���������
if kms then
Begin
 kmsfolder:=KMSini.ReadString('Vaccina','Foder',extractfilepath(application.ExeName)+'\Vaccina');
 DelKeyKms:=KMSini.ReadBool('Vaccina','DeleteKeyKMS',false);
 case KMSini.ReadInteger('Vaccina','Deadline',1) of
   0:
   begin   // � ��������������
   KMSpath:= KMSini.ReadString('Vaccina','Auto',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
   KMSkey:=KMSini.ReadString('Vaccina','AutoKey','/a /q');
   Deadline:=0;
   end;
   1:
   begin  // �� 180 ����
   KMSpath:= KMSini.ReadString('Vaccina','Manual',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
   KMSkey:=KMSini.ReadString('Vaccina','ManualKey','/m /q');
   Deadline:=1;
   end;
   2:
   begin   // �������� KMS
   KMSpath:= KMSini.ReadString('Vaccina','Uninstall',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
   KMSkey:=KMSini.ReadString('Vaccina','UninstallKey','/r /q');
   Deadline:=2;
   end;
 end;
End;
finally
KMSini.free;
end;
end;


try
InventMicrosoft:=true;// ������� ������� �������������� ��������� ����������
OleInitialize(nil);
SolveExitInvMicrosoft:=false; // ������� ���� ��� �������������� ��������
CoutPCOK:=0;
for I := 0 to ListPC.Count-1 do
BEGIN
if not InventMicrosoft then break; // ���� ���������� �������������� �� �������
if not SelectedPing(resIP,ListPC[i],pingtimeout) then frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:=ListPC[i]+' - �� ��������'// ���� ���� �� ��������
 else  // ����� ��������� ����������� wmi  � ���������
  if ConnectWMI(ListPC[i],User,Pass) then
  Begin
  CurrentPC:=ListPC[i];
  nameOS:='';
  VerOS:='';
  osconfig(nameOS,VerOS); // ����������� ������ ��
  WinProduct:='';
  StatWinLic:='����������';
  KeyWin:='';
  TypelicWin:='';
  DescriptWin:='';
  sOfficeProduct:='';
  sStatOfLic:='';
  sKeyOfc:='';
  sTypelicOffice:='';
  sDescriptLicOffice:='';
  CurrentOfficeLic:='����������'; // �.�. ��������� Office ����� � ����� Office ��� �������� ������� ��� � ����, �� � MAINPC ���������� ������ ��� ��������
  OfficeProduct:=TstringList.Create;
  StatOfLic:=TstringList.Create;
  KeyOfc:=TstringList.Create;
  TypelicOffice:=TStringList.Create;
  DescriptLicOffice:=TStringList.Create;
  try
  LicensingStatus(WinProduct,StatWinLic,KeyWin,TypelicWin,DescriptWin,OfficeProduct,StatOfLic,KeyOfc,TypelicOffice,DescriptLicOffice); // �������� ���������� �� ���������

  sOfficeProduct:=OfficeProduct.CommaText;
  sStatOfLic:=StatOfLic.CommaText;
  sKeyOfc:=KeyOfc.CommaText;
  sTypelicOffice:=TypelicOffice.CommaText;
  sDescriptLicOffice:=DescriptLicOffice.CommaText;
  finally
  OfficeProduct.Free;
  StatOfLic.Free;
  KeyOfc.Free;
  TypelicOffice.Free;
  DescriptLicOffice.Free;
  end;
    try
    TransactionWrite.StartTransaction;
    FDQueryWrite.SQL.Clear;                                                              /// ERROR_INV �������� ������ ������� ��������������
    FDQueryWrite.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
    +' MAIN_PC (PC_NAME,STATWINLIC,SSTATOFLIC) VALUES'
    +' ('''+ListPC[i]+''','''+leftstr(STATWINLIC,50)+''','''+leftstr(CurrentOfficeLic,500)+''')'
    +' MATCHING (PC_NAME)';
    FDQueryWrite.ExecSQL;
    TransactionWrite.commit;
    FDQueryWrite.Close;
    except on E: Exception do
      begin
      Log_write(2,'Licenses',ListPC[i]+' - ������ ������ ������ (MAIN_PC) - '+e.Message);
      if TransactionWrite.Active then begin  TransactionWrite.Rollback; FDQueryWrite.Close;  end;
      end;
    end;

    try
    TransactionWrite.StartTransaction;
    FDQueryWrite.SQL.Clear;                                                              /// ERROR_INV �������� ������ ������� ��������������
    FDQueryWrite.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
    +' MICROSOFT_LIC (NAMEPC,WINPRODUCT,STATWINLIC,KEYWIN,TYPE_LIC_WIN,DESCRIPT_LIC_WIN,SOFFICEPRODUCT,SSTATOFLIC,SKEYOFC,TYPE_LIC_OFFICE,DESCRIPT_LIC_OFFICE) VALUES'
    +' ('''+ListPC[i]+''','''+leftstr(WINPRODUCT,250)+''','''+leftstr(STATWINLIC,100)+''','''+leftstr(KEYWIN,50)+''','''+leftstr(TypelicWin,50)+''','''+leftstr(DescriptWin,250)+''','''+leftstr(SOFFICEPRODUCT,500)+''','''+leftstr(SSTATOFLIC,500)+''','''+leftstr(SKEYOFC,200)+''','''+leftstr(sTypelicOffice,150)+''','''+leftstr(sDescriptLicOffice,1500)+''')'
    +' MATCHING (NAMEPC)';
    FDQueryWrite.ExecSQL;
    TransactionWrite.commit;
    FDQueryWrite.Close;
    inc(CoutPCOK);
    except on E: Exception do
      begin
      Log_write(2,'Licenses',ListPC[i]+' - ������ ������ ������ (MICROSOFT_LIC) - '+e.Message);
      if TransactionWrite.Active then begin  TransactionWrite.Rollback; FDQueryWrite.Close;  end;
      end;
    end;


  if (KMS and ActivateYesNo) then // kms - ���� � ���������� ������� ������� ��������, ActivateYesNo - ���� ������ �������������� ������ �� ��������� ������� �������� � ���� Deadline<>2 �� ������� ��������
    begin
    if Deadline<>2 then runkms(ListPC[i],user,pass); // ���� �� �������� ��������, ������ ���� �������
    end;

  if (KMS and not ActivateYesNo) then // ���� ���� ������������ � ������ ��������� �� ������� �� �����
   begin
   if (KMSUnknownStat) and ((StatWinLic='����������')or(CurrentOfficeLic='����������')) then // KMSUnknownStat ���� � ���������� ������� ������������ � ������������ ��� ������� ����������
    if Deadline<>2 then runkms(ListPC[i],user,pass);
   end;

  if KMS and (Deadline=2) then  //kms - ���� � ���������� ������� ������� �������� � Deadline=2 �������� �������
    begin
    if (pos(ansiuppercase('KMS'),ansiuppercase(TypelicWin))<>0) or
    (pos(ansiuppercase('KMS'),ansiuppercase(sTypelicOffice))<>0) then // ���� ��� ����� ������� ������������� ������� ��� office �������� KMS
     begin
     runkms(ListPC[i],user,pass); // ��������� ������� � ������� �������
     if DelKeyKms then DeleteKeyKMS(ListPC[i],user,pass); // ���� � ���������� ������� ������� ���� KMS
     end;
    end;


  FWMIService:=Unassigned;
  FSWbemLocator:=Unassigned;
  frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:=ListPC[i]+' - OK';
  End; // ���� ���������� c WMI �����������
  frmDomainInfo.StatusInvMicrosoft.Panels[2].Text:=inttostr(ListPC.Count)+'/'+IntToStr(i+1)+'/'+inttostr(CoutPCOK)+' - OK';
if not InventMicrosoft then break; // ���� ���������� �������������� �� �������
END;// ����

finally
frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:='�������������� ���������';
Log_write(0,'Licenses',' �������������� ���������');
FDQueryWrite.free;
TransactionWrite.Free;
ListPC.Free;
ConnectionThRead.Close;
ConnectionThRead.Free;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
OleUnInitialize;
SolveExitInvMicrosoft:=true; // ������ ���� ��� ���� ����� ���������
InventMicrosoft:=false;
end;
result:=true;
except on E: Exception do
  begin
  result:=false;
  Log_write(3,'Licenses','����� ������ - '+e.Message);
  frmDomainInfo.StatusInvMicrosoft.Panels[1].Text:='������ ��������������';
  SolveExitInvMicrosoft:=true; // ������ ���� ��� ���� ����� ���������
  InventMicrosoft:=false;
  FWMIService:=Unassigned;
  FSWbemLocator:=Unassigned;
  end;
end;
END;

procedure InventoryMicrosoft.Execute;
begin
ListPCWindowsOffice:=TStringList.Create;
ListPCWindowsOffice.Text:=ListPCWO.Text;
ListPCWO.Free;
UserWO:=MyUser;
PassWO:=MyPasswd;
ScanMicrosoftWO(userWo,PassWO,ListPCWindowsOffice);
ListPCWindowsOffice.Free;
end;

end.
