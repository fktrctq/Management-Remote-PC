unit LoadCurencyList;

interface
uses System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,
WinApi.windows,Vcl.StdCtrls,Vcl.ComCtrls,Vcl.Controls,vcl.forms,System.StrUtils,Vcl.ExtCtrls,FireDAC.Phys.FB,
 FireDAC.Comp.Client,FireDAC.Stan.Option;

type
  TPropProgpamMSI = record
    NamePC       :String;
    User         :String;
    Pass         :String;
    ListViewMsi  :TListview;
  end;

  TPropWinUpdate = record
    NamePC       :String;
    User         :String;
    Pass         :String;
    ListViewMsi  :TListview;
  end;

  TPropProcessLoad = record
    NamePC       :String;
    User         :String;
    Pass         :String;
    TimerScan    :Ttimer;
    ListViewProc :TListview;
  end;

  TPropAutostart =record
    namePC     :string;
    User       :string;
    Pass       :string;
    SystemType :string;
    ListViewS  :TListView;
  end;

function LoadListProgpamMSI(paramMSI:pointer):integer;  /// ����� �������� ������ msi ��������
function LoadListWindowsUpdate(paramUpdate:pointer):integer;  /// ����� �������� ������ ����������
function LoadListProcess(paramProc:pointer):integer; // ����� �������� ������ ���������
function LoadListServicesThread(paramUpdate:pointer):integer; /// ����� �������� ����� �����
function LoadListServices(NamePC,User,Pass:string;ListViewS:TListView):bool; ///������� ������ �����
function LoadListAllProgram(ParamUpdate:pointer):integer; /// �������� ���� ��������
Function LoadListLocalDisk(s,User,Pass:string;ListViewDisk:TlistView):bool; /// �������� ��������� ������
function LoadListSNetworkInterface(S,User,Pass:string;ListViewNI:TListView):bool; //�������� ������� �����������
function LoadListAutoStart(S,User,Pass,SystemType:string;ListViewS:TListView):bool; // �������� ������ ������������
function LoadListProfile(S,User,Pass:string;ListViewS:TListView):bool; // �������� ������ ��������
function GetUserDirectory(NamePC,User,Pass,UserProfile,UsSID:string):string; // ������ ���� � ������� ������������
function LoadDevicePnP(PcName,User,Pass:string;ListViewPnP:TListView):boolean; // �������� ������ �������� PnP
function EnableDisableDevicePnP(PcName,User,Pass,DeviceID:string; EnabOrDis:integer; var reboot:boolean):integer; // ��������� � ����������� ���������� PnP
function ReadAdnWriteAntivirusStaus(NamePC,User,Pass:string;LV:TlistView):bool; // ������ � ������ ������ � ����������� �� ��
function itemLisToMemo(ListStr:TStringList;Des:string;own:Tform):boolean;
var
CallListMSI    :TPropProgpamMSI;
CallListUpdate :TPropWinUpdate;
CallListProcess :TPropProcessLoad;
CallListServise :TPropWinUpdate;
CallListAllProgram :TPropWinUpdate;
CalllistAutoStart:TPropAutostart;
implementation
uses umain,MYDM;

ThreadVar
ReceptListMSI: ^TPropProgpamMSI;
ReceptListUpdate: ^TPropWinUpdate;
ReceptListServ: ^TPropWinUpdate;
ReceptListProg: ^TPropWinUpdate;
ReceptListProcess: ^TPropProcessLoad;
ReceptListAutoStart:^TPropAutostart;

/////////////////////////////////////////////////////////////////////////////
function itemLisToMemo(ListStr:TStringList;Des:string;own:Tform):boolean;
var
i:integer;
DetalForm:Tform;
DetalMemo:Tmemo;
begin
try
DetalForm:=TForm.Create(own);
DetalForm.Caption:=Des;
DetalForm.Height:=250;
DetalForm.Width:=550;
DetalForm.Position:=poOwnerFormCenter;
DetalForm.BorderStyle:=bsDialog;
DetalForm.KeyPreview:=true;
DetalForm.OnClose:=frmDomainInfo.AllFormClose;
DetalForm.OnKeyPress:=frmDomainInfo.CloseFormEsc;
DetalMemo:=Tmemo.Create(DetalForm);
DetalMemo.parent:=DetalForm;
DetalMemo.Align:=alClient;
DetalMemo.ScrollBars:=ssVertical;
DetalMemo.ScrollBars:=ssHorizontal;
DetalMemo.ReadOnly:=true;
DetalMemo.Clear;
for I := 0 to ListStr.Count-1 do DetalMemo.Lines.Add(ListStr[i]);
DetalForm.ShowModal;
except on E: Exception do showmessage('������ '+e.Message);
end;
end;

//////////////////////////////////////////////////////////  ������ � ������ ������� ������������� �����������
function ReadAdnWriteAntivirusStaus(NamePC,User,Pass:string;LV:TlistView):bool;
var
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
z:integer;
AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate:string;
AntySpw,AntySpwStatus,AntySpwUpdate,AntySpwDef,AntySpwDefStatus,AntySpwDefUpdate,nameOS:string;
function statusAntivirus (stateAnt,OnOrUpdate:string):string;
begin
if length(stateAnt)<5 then   // ���� ������ ������ ������� ������ 5�� ��������
begin
  result:='����������� ������ ('+stateAnt+')';
  exit;
end;
if OnOrUpdate='On' then
 begin
 if (copy(stateAnt,2,2)='10') or (copy(stateAnt,2,2)='11') then
  result:='�������'
  else
 if (copy(stateAnt,2,2)='00') or (copy(stateAnt,2,2)='01') then
  result:='��������'
  else result:='�� ����������';
 end;
if OnOrUpdate='Update' then
begin
  if (copy(stateAnt,4,2)='00') then
  result:='Ok'
  else
 if (copy(stateAnt,4,2)='10') then
  result:='���� ��������'
  else result:='�� ����������';
end;
end;
function StatusImage(AntStatus,AntUpdate:string):integer;
begin
try
if (AntStatus='�������')and(AntUpdate='Ok') then result:=18  // ������ ��
else if (AntStatus='�������')and(AntUpdate='���� ��������') then result:=21  // �� ������ ��
else if (AntStatus='��������')and(AntUpdate='Ok') then result:=19  // ���� �� ���
else if (AntStatus='��������')and(AntUpdate='���� ��������') then result:=20 //�������
else  result:=19;  // ���� ����

except on E: Exception do result:=19;
end;
end;
function AntivirusUpdateTable(NamePC,AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate,
AntySpw,AntySpwStatus,AntySpwUpdate:string):boolean;
var
TransactionWrite:TFDTransaction;
FDQueryWrite:TFDQuery;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
FDQueryWrite:=TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into ANTIVIRUSPRODUCT '+
  ' (NAMEPC,WIN_DEF,WIN_DEF_STATUS,WIN_DEF_UPDATE,ANTIVIRUS,ANTIVIRUS_STATUS,ANTIVIRUS_UPDATE,FIREWALL,FIREWALL_STATUS,ANTISPYWARE,ANTISPYWARE_STATUS,ANTISPYWARE_UPDATE)'+
  ' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12) MATCHING (NAMEPC)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+(NamePC)+'';
FDQueryWrite.params.ParamByName('p2').AsString:=''+(DefProd)+'';
FDQueryWrite.params.ParamByName('p3').AsString:=''+(DefStatus)+'';
FDQueryWrite.params.ParamByName('p4').AsString:=''+(DefUpdate)+'';
FDQueryWrite.params.ParamByName('p5').AsString:=''+(AntProd)+'';
FDQueryWrite.params.ParamByName('p6').AsString:=''+(AntStatus)+'';
FDQueryWrite.params.ParamByName('p7').AsString:=''+(AntUpdate)+'';
FDQueryWrite.params.ParamByName('p8').AsString:=''+(FWProd)+'';
FDQueryWrite.params.ParamByName('p9').AsString:=''+(FWStatus)+'';
FDQueryWrite.params.ParamByName('p10').AsString:=''+(AntySpw)+'';
FDQueryWrite.params.ParamByName('p11').AsString:=''+(AntySpwStatus)+'';
FDQueryWrite.params.ParamByName('p12').AsString:=''+(AntySpwUpdate)+'';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
except on E: Exception do
begin
TransactionWrite.Rollback;
frmDomainInfo.Memo1.lines.add(' ������ ������ � �� :'+E.Message);
end;
end;
finally
FDQueryWrite.Free;
TransactionWrite.Free;
end;
end;

begin
try
//////////////////////////////////////////////////
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');  ////IWbemLocator ���   SWbemLocator
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User,Pass,'','',128); ///WbemUser, WbemPassword ,0- ����� ������ �� ��������� 128- ����� 2 ������
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_OperatingSystem','WQL',
wbemFlagReturnImmediately);   //wbemFlagReturnImmediately ��� wbemFlagForwardOnly
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 if oEnum.Next(1, FWbemObject, iValue) = s_ok then     //// ������������ �������
    begin
    nameOS:=trim(string(FWbemObject.Caption));
    FWbemObject := Unassigned;
    end;
FWbemObject := Unassigned;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
oEnum:=nil;

///////////////////////////////////////////////////////////////////////////////////////////
//�������� ������ ����������  AntivirProd,AntivirStat
//https://theroadtodelphi.com/2011/02/18/getting-the-installed-antivirus-antispyware-and-firewall-software-using-delphi-and-the-wmi/
//https://social.msdn.microsoft.com/Forums/pt-BR/6501b87e-dda4-4838-93c3-244daa355d7c/wmisecuritycenter2-productstate?forum=vblanguage
if (pos('Server',nameOS)<>0) or (pos('Windows XP',nameOS)<>0) then
begin
  // ���� ������� �� ��������� � �� ��
showmessage('�� ����������� ������������ �������...');
exit;
end;
BEGIN
try
AntProd:='';AntStatus:='';AntUpdate:='';
FWProd:='';FWStatus:='';
DefProd:='';DefStatus:='';DefUpdate:='';
AntySpw:='';AntySpwStatus:='';AntySpwUpdate:='';
AntySpwDef:='';AntySpwDefStatus:='';AntySpwDefUpdate:='';
FWbemObject := Unassigned;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
oEnum:=nil;
LV.clear;
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\SecurityCenter2',User,Pass,'','',128);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM AntiVirusProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  //'----AntiVirusProduct--------'
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
  if FWbemObject.displayName<>null then
  Begin
  with lv.items.add do // ��������� ListView
  begin
  caption:='��������� '+vartostr(FWbemObject.displayName);
  if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    subitems.add(statusAntivirus(IntToHex(z,2),'On'));
    subitems.add(statusAntivirus(IntToHex(z,2),'Update'));
    imageIndex:=StatusImage(subitems[0],subitems[1]);
    end else begin subitems.add(''); subitems.add(''); end;
  end;
  if pos(ansiuppercase('Defender'),ansiuppercase(vartostr(FWbemObject.displayName)))<>0 then
    begin // ���������� ����������
    DefProd:= vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    DefStatus:=statusAntivirus(IntToHex(z,2),'On');
    DefUpdate:=statusAntivirus(IntToHex(z,2),'Update');
    end
    end
  else
  if (AntProd<>vartostr(FWbemObject.displayName)) and (AntStatus<>'�������') then //
  begin
    AntProd:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    AntStatus:=statusAntivirus(IntToHex(z,2),'On');
    AntUpdate:=statusAntivirus(IntToHex(z,2),'Update');
    end;
  end;
  FWbemObject:=Unassigned;
  End;
  if (AntProd='')and(DefProd<>'') then
  begin
  AntProd:=DefProd;
  AntStatus:=DefStatus;
  AntUpdate:=DefUpdate;
  end;
 //'-------FirewallProduct------'
  VariantClear(FWbemObjectSet);
  VariantClear(FWbemObject);
  oEnum:=nil;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM FirewallProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do
  if FWbemObject.displayName<>null then
  Begin
  with lv.items.add do // ��������� ListView
  begin
  caption:='������� '+vartostr(FWbemObject.displayName);
  if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    subitems.add(statusAntivirus(IntToHex(z,2),'On'));
    subitems.add('');
    imageIndex:=StatusImage(subitems[0],'Ok');
    end else begin subitems.add(''); subitems.add(''); end;
  end;
  if (FWProd<>vartostr(FWbemObject.displayName))and(FWStatus<>'�������') then
    begin // ��������� ����������
    FWProd:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    FWStatus:=statusAntivirus(IntToHex(z,2),'On');
    end;
  FWbemObject:=Unassigned;
  End;
   //-------AntiSpywareProduct--------
  VariantClear(FWbemObjectSet);
  VariantClear(FWbemObject);
  oEnum:=nil;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM AntiSpywareProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
  if FWbemObject.displayName<>null then
  Begin
  with lv.items.add do // ��������� ListView
  begin
  caption:='��������� '+vartostr(FWbemObject.displayName);
  if TryStrToInt(vartostr(FWbemObject.productState),Z) then
    begin
    subitems.add(statusAntivirus(IntToHex(z,2),'On'));
    subitems.add(statusAntivirus(IntToHex(z,2),'Update'));
    imageIndex:=StatusImage(subitems[0],subitems[1]);
    end else begin subitems.add(''); subitems.add(''); end;
  end;
  if pos(ansiuppercase('Defender'),ansiuppercase(vartostr(FWbemObject.displayName)))<>0 then
    begin
    AntySpwDef:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
      begin
      AntySpwDefStatus:=statusAntivirus(IntToHex(z,2),'On');
      AntySpwDefUpdate:=statusAntivirus(IntToHex(z,2),'Update');
      end;
    end
    else
  if (AntySpwStatus<>'�������')or((AntySpwUpdate<>'Ok')and(AntySpwStatus='�������')) then
    begin
    AntySpw:=vartostr(FWbemObject.displayName);
    if TryStrToInt(vartostr(FWbemObject.productState),Z) then
      begin
      AntySpwStatus:=statusAntivirus(IntToHex(z,2),'On');
      AntySpwUpdate:=statusAntivirus(IntToHex(z,2),'Update');
      end;
    end;
  FWbemObject:=Unassigned;
  End;
  if ((AntySpwStatus='��������')or (AntySpwStatus=''))and ((AntySpwDefStatus='�������')) then
  begin
   AntySpwStatus:=AntySpwDefStatus;
   AntySpwStatus:=AntySpwDefStatus;
   AntySpwUpdate:=AntySpwDefUpdate;
  end;
///////���������� ������
AntivirusUpdateTable(NamePC,AntProd,AntStatus,AntUpdate,FWProd,FWStatus,DefProd,DefStatus,DefUpdate,
AntySpw,AntySpwStatus,AntySpwUpdate);
result:=true;
except
 on E:Exception do
  begin
  frmDomainInfo.Memo1.lines.add('������  :'+E.Message);
  end;
end;
END; // ��������� ������������ �����������
//////////////////////////////////////////////////////////////////////////////////////////////
Except
  on E:Exception do
     begin
     result:=false;
     frmDomainInfo.Memo1.lines.add('�����  ������ :'+E.Message);
     end;
   end;
FWbemObject := Unassigned;
oEnum:=nil;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize();
end;
//***********************************************************************
//////////////////////////////////////////////////////////
//�������� ��� ������������ � ����� �� SID
function SidToAcountName(SecureObject,S: string): String;
var
   Sid: PSID;
      peUse: DWORD;
      cchDomain: DWORD;
      cchName: DWORD;
      Name: array of Char;
      Domain: array of Char;
begin
result:='';
Sid := nil;
// First convert String SID to SID
Win32Check(ConvertStringSidToSid(PChar(SecureObject), Sid));
cchName := 0;
cchDomain := 0;
// Get Length             // nil ��� ��������� ����
if (not LookupAccountSid({Pchar(s) ���}nil, Sid, nil, cchName, nil, cchDomain, peUse))
and (GetLastError = ERROR_INSUFFICIENT_BUFFER) then
begin
SetLength(Name, cchName);
SetLength(Domain, cchDomain);
if LookupAccountSid(nil, Sid, @Name[0], cchName, @Domain[0], cchDomain, peUse) then
begin
// note: cast to PChar because LookupAccountSid returns zero terminated string
Result:=PChar(Domain)+'\'+ PChar(Name);
//result:=
end else result:='';
end;
end;

////////////////////////////////////////////////////////////////////////////////////////
function SIDtoUserName (NewSid: string):string;  /// ��� ������������ �� SID
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject         : OLEVariant;
  UserName:string;
  oEnum               : IEnumvariant;
  iValue        : LongWord;
begin
result:='';
    try
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT name FROM Win32_UserAccount WHERE SID ='+'"'+NewSID+'"','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������
            begin
            result:=(FWbemObject.name);
            //frmDomainInfo.memo1.Lines.Add(username);
            end;
     except
     on E:Exception do
       begin
         result:='Unknown';
         frmDomainInfo.memo1.Lines.Add('��� ������������, ������ - "'+E.Message+'"');
         exit;
        end;
       end;
FSWbemLocator:=Unassigned;
end;

function SIDtoUserNameRemotePC (NewSid:string;FWMIService:OLEVariant):string;  /// ��� ������������ �� SID
var
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
begin
result:='';
    try
    OleInitialize(nil);
    FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,Domain FROM Win32_UserAccount WHERE SID ='+'"'+NewSID+'"','WQL',wbemFlagForwardOnly);
    oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    if oEnum.Next(1, FWbemObject, iValue) = 0 then
    begin
    if FWbemObject.name<>null then
    result:=string(FWbemObject.name);
    if FWbemObject.Domain<>null then
    result:=string(FWbemObject.Domain)+'\'+result;
    FWbemObject:=Unassigned;
    end;
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
   except
    begin
    result:='Unknown';
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    end;
   end;
 oleUnInitialize;
end;

function WbemTimeToDateTimeReg(const V : OleVariant): string; //������� ����
var
  Dt,s : string;
  FormatSettings,FormatSettings1: TFormatSettings;
begin
  //Result:=0;
  try
  if VarIsNull(V) then exit;
  Dt:=string(v);
  insert('/',dt,5);
  insert('/',dt,8);
  FormatSettings.ShortDateFormat:='yyyy-mm-dd';
  FormatSettings.DateSeparator:='/';
  FormatSettings1.ShortDateFormat:='dd.mm.yyyy';
  FormatSettings1.DateSeparator:=':';
  result:=datetostr(strtodate(dt,FormatSettings),FormatSettings1)
  except
   result:='01.01.2001';
  end;
  //result:=dt;
end;

function  WbemTimeToDateTime(vDate : OleVariant) : string;
var
  FWbemDateObj  : OleVariant;
  TDT:TdateTime;
begin;
  FWbemDateObj  := CreateOleObject('WbemScripting.SWbemDateTime');
  FWbemDateObj.Value:=vDate;
  TDT:=FWbemDateObj.GetVarDate;
  result:=DateTimeToStr(TDT);
end;

function WbemTimeToDateAndTime(const V : OleVariant): string; //������� ���� � �������
var
  Dt : string;
begin
  if VarIsNull(V) then exit;
  Dt:=string(v);
  insert('/',dt,5);
  insert('/',dt,8);
  insert('-',dt,11);
  insert(':',dt,14);
  insert(':',dt,17);
  delete(dt,20,Length(dt)-19);
  result:=dt;
end;

function DateConvert(s:string):string;  //// ������������ ����  �� ������
var                                     ///// � ������� �������� � ������ �������
FormatSettings,FormatSettings1: TFormatSettings;
begin
FormatSettings.ShortDateFormat:='mm-dd-yyyy';
FormatSettings.DateSeparator:='/';
FormatSettings1.ShortDateFormat:='dd.mm.yyyy';
FormatSettings1.DateSeparator:=':';
result:=datetostr(strtodate(s,FormatSettings),FormatSettings1);
end;

function LoadListProcess(paramProc:pointer):integer; /// ��������� ������ ���������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet,FWbemObjectSetCPU  : OLEVariant;
  FWbemObject,FWbemObjectCPU     : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum,oEnumCPU           : IEnumvariant;
  iValue          : LongWord;
  MyCPU,UTime,PTime         :integer;
  AccountUs,DomainName:OleVariant;
  UserStr:string;
begin
ReceptListProcess:=paramProc;
OleInitialize(nil);
 begin
       MyCPU:=CpuLogPr*CpuNum;
       frmDomainInfo.memo1.Lines.Add('�������� ������ ���������..');
        try
         ReceptListProcess.ListViewProc.Clear;
         ReceptListProcess.ListViewProc.Items.BeginUpdate;
         FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
         FWMIService   := FSWbemLocator.ConnectServer(ReceptListProcess.NamePC, 'root\CIMV2',ReceptListProcess.User, ReceptListProcess.Pass,'','',128);
         FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,ProcessId,CommandLine FROM Win32_Process WHERE ProcessId<>0','WQL',wbemFlagForwardOnly);
         oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         FWbemObjectSetCPU:= FWMIService.ExecQuery('SELECT IDProcess,PercentPrivilegedTime,PercentUserTime,WorkingSet,WorkingSetPeak FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess<>0' ,'WQL',wbemFlagForwardOnly);
         oEnumCPU:= IUnknown(FWbemObjectSetCPU._NewEnum) as IEnumVariant;
            while (oEnum.Next(1, FWbemObject, iValue) = 0) and (oEnumCPU.Next(1, FWbemObjectCPU, iValue) = 0)  do     //// ��������
            begin
                    if FWbemObject.GetOwner(AccountUs,DomainName)=0 then UserStr:=string(AccountUs)
                   else UserStr:='SYSTEM';
                      with ReceptListProcess.ListViewProc.Items.Add do
                         begin
                         Caption := inttostr(ReceptListProcess.ListViewProc.Items.Count-1);          // �
                         SubItems.Add((FWbemObject.Caption));       // ��� ���
                         Subitems.Add(UserStr);
                         Subitems.Add(vartostr(FWbemObject.ProcessId));
                           if (FWbemObject.ProcessId)=(FWbemObjectCPU.IDProcess) then
                              begin
                              if (FWbemObjectCPU.PercentPrivilegedTime<>null)and(FWbemObjectCPU.PercentUserTime<>null) then
                              begin
                              UTime:= (FWbemObjectCPU.PercentPrivilegedTime) div (MyCPU);
                              PTime:= (FWbemObjectCPU.PercentUserTime) div (MyCPU);
                              SumTime:=SumTime+(UTime+Ptime);
                              end;
                              if FWbemObjectCPU.WorkingSet<>null then
                             SumMemory:= SumMemory+round(((FWbemObjectCPU.WorkingSet) / 1024));
                              if FWbemObjectCPU.WorkingSetPeak<>null then
                              SetPeak:=SetPeak+round(((FWbemObjectCPU.WorkingSetPeak) / 1024));
                              Subitems.Add(vartostr(UTime+Ptime));
                              if FWbemObjectCPU.WorkingSet<>null then
                              Subitems.Add(((vartostr(round((FWbemObjectCPU.WorkingSet) / 1024))))+' Kb')
                              else Subitems.Add('0');
                              end
                           else
                             begin
                             Subitems.Add('0');
                             Subitems.Add('0');
                             end;
                           if (FWbemObject.CommandLine)<>Null then Subitems.Add(string(FWbemObject.CommandLine))
                           else Subitems.Add('');
                          end;
             VariantClear(FWbemObject);
             VariantClear(FWbemObjectCPU);
            end;
            ReceptListProcess.ListViewProc.Items.EndUpdate;
            oEnum:=nil;
            oEnumCPU:=nil;
            VariantClear(FWbemObjectSet);
            VariantClear(FWbemObjectSetCPU);
            VariantClear(FSWbemLocator);
            VariantClear(FWMIService);
            except
               on E:Exception do
               begin
               ReceptListProcess.ListViewProc.Items.EndUpdate;
               frmDomainInfo.memo1.Lines.Add('������ �������� ��������� "'+E.Message+'"');
               oEnum:=nil;
               oEnumCPU:=nil;
               VariantClear(FWbemObjectSet);
               VariantClear(FWbemObjectSetCPU);
               VariantClear(FSWbemLocator);
               VariantClear(FWMIService);
               end;
            end;
OleUnInitialize;
frmDomainInfo.memo1.Lines.Add('�������� ������ ��������� ���������.');
frmDomainInfo.Timer1.Enabled:=true;
frmDomainInfo.TabSheet2.Tag:=0;
NetApiBufferFree(paramProc); /// ������� ������
EndThread(result); /// ����� ������
end;
end;

function LoadListProgpamMSI(paramMSI:pointer):integer; /// ������ �������� msi
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
 begin
    try
      ReceptListMSI:=paramMSI;
      ReceptListMSI.ListViewMsi.Clear;
      ReceptListMSI.ListViewMsi.Items.BeginUpdate;
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(ReceptListMSI.NamePC, 'root\CIMV2',ReceptListMSI.User, ReceptListMSI.Pass,'','',128);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT Version,InstallDate,Caption,InstallLocation,Vendor,InstallSource FROM Win32_Product','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
            begin
               InstallLocation:='';
               InstallSource:='';
               MyVendor:='';
               MyVersion:='';
               if (FWbemObject.InstallLocation)<>null then InstallLocation:=(string(FWbemObject.InstallLocation));
               if (FWbemObject.InstallSource)<>null then InstallSource:=(string(FWbemObject.InstallSource));
               if (FWbemObject.Vendor)<>null then MyVendor:=(string(FWbemObject.Vendor));
               if (FWbemObject.Version)<>null then MyVersion:=(string(FWbemObject.Version));
               if FWbemObject.Caption<> null then
                  begin
                   with ReceptListMSI.ListViewMsi.Items.Add do
                     begin
                     Caption := inttostr(ReceptListMSI.ListViewMsi.Items.Count);
                     SubItems.add(string(FWbemObject.Caption));
                     SubItems.add(InstallLocation);
                     SubItems.add(InstallSource);
                     if FWbemObject.InstallDate<>null then SubItems.add((WbemTimeToDateTimeReg(FWbemObject.InstallDate)))
                     else SubItems.add('');
                     SubItems.add(MyVersion);
                     SubItems.add(MyVendor);
                     ImageIndex:=frmDomainInfo.BrendProg(MyVendor);
                     end;
                   end;
               FWbemObject:=Unassigned;
               end;
      ReceptListMSI.ListViewMsi.Items.EndUpdate;
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
      except
      on E:Exception do
        begin
        ReceptListMSI.ListViewMsi.Items.EndUpdate;
        frmDomainInfo.memo1.Lines.Add('��������� - "'+E.Message+'"');
        frmDomainInfo.SpeedButton3.Enabled:=true;
        end;
       end;
frmDomainInfo.memo1.Lines.Add('�������� ������ �������� ���������.');
frmDomainInfo.SpeedButton3.Enabled:=true;
NetApiBufferFree(paramMSI); /// ������� ������
EndThread(result); /// ����� ������
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
function LoadListWindowsUpdate(paramUpdate:pointer):integer; /// ������ ���������� windows
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
 begin
    try
      ReceptListUpdate:=paramUpdate;
      ReceptListUpdate.ListViewMsi.Clear;
      ReceptListUpdate.ListViewMsi.Items.BeginUpdate;
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(ReceptListUpdate.NamePC, 'root\CIMV2',ReceptListUpdate.User, ReceptListUpdate.Pass,'','',128);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,Description,HotFixID,InstalledBy,InstalledOn FROM Win32_QuickFixEngineering','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ���������� �����������
            begin
                  begin
                   with ReceptListUpdate.ListViewMsi.Items.Add do
                     begin
                     Caption := inttostr(ReceptListUpdate.ListViewMsi.Items.Count-1);
                    if FWbemObject.HotFixID<>null then SubItems.Add(vartostr(FWbemObject.HotFixID))   // ok
                    else SubItems.Add('');
                    if FWbemObject.Description<>null then SubItems.Add(vartostr(FWbemObject.Description))   // ok
                    else SubItems.Add('');
                    if FWbemObject.Caption<>null then SubItems.Add(vartostr(FWbemObject.Caption))   // ok
                    else SubItems.Add('');
                    if (FWbemObject.InstalledBy<>null)  then SubItems.Add(vartostr(FWbemObject.InstalledBy))   // ok
                    else SubItems.Add('');
                    if (FWbemObject.InstalledOn<>null)  then SubItems.Add(DateConvert(vartostr(FWbemObject.InstalledOn)) )  // ok
                    else SubItems.Add('');
                     end;
                   end;
               FWbemObject:=Unassigned;
               end;
      ReceptListUpdate.ListViewMsi.Items.EndUpdate;
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
      except
      on E:Exception do
        begin
        ReceptListUpdate.ListViewMsi.Items.EndUpdate;
        frmDomainInfo.TabSheet13.Tag:=0; //// ������� ��������� �������� ������ ����������
        frmDomainInfo.memo1.Lines.Add('������ �������� ������ ���������� Windows "'+E.Message+'"');
        end;
       end;
frmDomainInfo.memo1.Lines.Add('�������� ������ ����������/����������� Windows ���������.');
frmDomainInfo.TabSheet13.Tag:=0; //// ������� ��������� �������� ������ ����������
NetApiBufferFree(paramUpdate); /// ������� ������
EndThread(result); /// ����� ������
end;
/////////////////////////////////////////////////////////////////////////////////////////////////
function LoadListServices(NamePC,User,Pass:string;ListViewS:TListView):bool; /// ������ �����
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
 begin
    try
      ListViewS.Clear;
      frmDomainInfo.memo1.Lines.Add('�������� ������ �����...');
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2',User, Pass,'','',128);
      FWMIService.Security_.impersonationlevel:=3;
      FWMIService.Security_.authenticationLevel := 6;
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,ProcessId,caption,state,StartMode,PathName FROM Win32_Service','WQL',wbemFlagForwardOnly);
         oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
            while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������
                begin
                  with ListViewS.Items.Add do
                     begin
                     Caption := inttostr(frmDomainInfo.listview1.Items.Count-1);          // �
                     SubItems.Add(vartostr(FWbemObject.name));       // ��� ���
                     if (FWbemObject.ProcessId)=0 then Subitems.Add('')
                     else Subitems.Add(vartostr(FWbemObject.ProcessId));    /// ID ��������
                     Subitems.Add(vartostr(FWbemObject.caption));
                     Subitems.Add(vartostr(FWbemObject.state));
                     Subitems.Add(vartostr(FWbemObject.StartMode));
                     Subitems.Add(vartostr(FWbemObject.PathName));
                     end;
                   FWbemObject:=Unassigned;
                   end;
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      VariantClear(FWbemObject);
      OleUnInitialize;
      except
      on E:Exception do
        begin
        VariantClear(FWbemObjectSet);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
        VariantClear(FWbemObject);
        frmDomainInfo.memo1.Lines.Add('������ �������� ������ ����� "'+E.Message+'"');
        end;
       end;
frmDomainInfo.memo1.Lines.Add('�������� ������ ����� ���������.');
end;
//////////////////////////////////////////////////////////////////////////////////////////////
function LoadListServicesThread(paramUpdate:pointer):integer; /// ������ �����
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
 begin
    try
      ReceptListServ:=paramUpdate;
      ReceptListServ.ListViewMsi.Clear;
      ReceptListServ.ListViewMsi.Items.BeginUpdate;
      frmDomainInfo.memo1.Lines.Add('�������� ������ �����...');
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(ReceptListServ.NamePC, 'root\CIMV2',ReceptListServ.User, ReceptListServ.Pass,'','',128);
      FWMIService.Security_.impersonationlevel:=3;
      FWMIService.Security_.authenticationLevel := 6;
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,ProcessId,caption,state,StartMode,PathName FROM Win32_Service','WQL',wbemFlagForwardOnly);
         oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
            while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������
                begin
                  with ReceptListServ.ListViewMsi.Items.Add do
                     begin
                     Caption := inttostr(ReceptListServ.ListViewMsi.Items.Count-1);          // �
                     SubItems.Add(vartostr(FWbemObject.name));       // ��� ���
                     if (FWbemObject.ProcessId)=0 then Subitems.Add('')
                     else Subitems.Add(vartostr(FWbemObject.ProcessId));    /// ID ��������
                     Subitems.Add(vartostr(FWbemObject.caption));
                     Subitems.Add(vartostr(FWbemObject.state));
                     Subitems.Add(vartostr(FWbemObject.StartMode));
                     Subitems.Add(vartostr(FWbemObject.PathName));
                     end;
                   FWbemObject:=Unassigned;
                  end;
      ReceptListServ.ListViewMsi.Items.EndUpdate;
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      VariantClear(FWbemObject);
      except
      on E:Exception do
        begin
        ReceptListServ.ListViewMsi.Items.EndUpdate;
        VariantClear(FWbemObjectSet);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
        VariantClear(FWbemObject);
        frmDomainInfo.memo1.Lines.Add('������ �������� ������ ����� "'+E.Message+'"');
        end;
       end;
OleUnInitialize;
frmDomainInfo.memo1.Lines.Add('�������� ������ ����� ���������.');
frmDomainInfo.TabSheet3.Tag:=0; //// ������� ��������� �������� ������
NetApiBufferFree(paramUpdate); /// ������� ������
EndThread(result); /// ����� ������
end;
//////////////////////////////////////////////////////////////////////////////////////////////
 function LoadListAllProgram(ParamUpdate:pointer):integer; /// ������ ������ ���� ��������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FInParams,FInParams1                    : OLEVariant;
  FOutParams,FOutParams1                  : OLEVariant;
  i:integer;
  nameSoft,InstallLocation,InstallSource,InstallDate,DisplayVersion,UninstallString,Publisher,RegValue:string;
 begin
 try
 ReceptListProg:=ParamUpdate;
 ReceptListProg.ListViewMsi.Clear;
 ReceptListProg.ListViewMsi.Items.BeginUpdate;
 frmDomainInfo.memo1.Lines.Add('�������� ������ ���� ��������... ��������');
 OleInitialize(nil);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(ReceptListProg.NamePC, 'root\default', ReceptListProg.User, ReceptListProg.Pass,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6;
 FWbemObjectSet:= FWMIService.Get('StdRegProv');
 FInParams     := FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
 FInParams.hDefKey:=HKEY_LOCAL_MACHINE;
 FInParams.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall';
 FoutParams    := FWMIService.ExecMethod('StdRegProv', 'EnumKey', FInParams);
 FInParams1     := FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_();
 FInParams1.hDefKey:=HKEY_LOCAL_MACHINE;
 for i:= VarArrayLowBound(FOutParams.sNames,1) to VarArrayHighBound(FOutParams.sNames,1) do
   begin
   FInParams1.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall\'+(string(FOutParams.sNames[i]));
 ////////////////////////////////////////////////////////////////////////////////////
   FInParams1.sValueName :='DisplayName';
   FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
   if FoutParams1.sValue<>null then
   begin
    nameSoft:='';InstallLocation:='';InstallSource:='';InstallDate:='';DisplayVersion:='';UninstallString:='';Publisher:=''; RegValue:='';
    nameSoft:=string(FoutParams1.sValue);
//////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='InstallLocation';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then InstallLocation:=(string(FoutParams1.sValue));
      ///////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='InstallSource';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then InstallSource:=(string(FoutParams1.sValue));
      //////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='InstallDate';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then
    InstallDate:=(WbemTimeToDateTimeReg(FoutParams1.sValue))
    else InstallDate:='01.01.2001';
      ////////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='DisplayVersion';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then DisplayVersion:=(string(FoutParams1.sValue));
      ///////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='UninstallString';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then UninstallString:=(string(FoutParams1.sValue));
      ////////////////////////////////////////////////////////////////////////////////////////////
    FInParams1.sValueName :='Publisher';
    VariantClear(FOutParams1);
    FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
    if FoutParams1.sValue<>null then Publisher:=(string(FoutParams1.sValue));
   /////////////////////////////////////////////////////////////////////////////////////////////////
   if vartostr(FOutParams.sNames[i])<>'' then RegValue:=(string(FOutParams.sNames[i]));
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////
    with ReceptListProg.ListViewMsi.Items.Add do
    begin
    if Publisher='' then ImageIndex:= frmDomainInfo.BrendProg('other')
    else ImageIndex:= frmDomainInfo.BrendProg(Publisher);
    Caption := inttostr(ReceptListProg.ListViewMsi.Items.Count);
    SubItems.add(nameSoft);
    SubItems.Add(InstallLocation);
    SubItems.Add(InstallSource);
    SubItems.Add(InstallDate);
    SubItems.Add(DisplayVersion);
    SubItems.Add(UninstallString);
    SubItems.Add(Publisher);
    SubItems.Add(RegValue);
    end;
   end;//if
VariantClear(FOutParams1);
end; //����

VariantClear(FWbemObjectSet);
variantclear(FInParams);
variantclear(FInParams1);
VariantClear(FoutParams);
VariantClear(FoutParams1);
////////////////////////////////////////////////////////////////////////////////////
///     HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall
///////////////////////////////////////////////////////////////////////////////////
begin
FWbemObjectSet:= FWMIService.Get('StdRegProv');
FInParams     := FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
FInParams.hDefKey:=HKEY_CURRENT_USER;
FInParams.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall';
FoutParams    := FWMIService.ExecMethod('StdRegProv', 'EnumKey', FInParams);
FInParams1     := FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_();
FInParams1.hDefKey:=HKEY_CURRENT_USER;
for i:= VarArrayLowBound(FOutParams.sNames,1) to VarArrayHighBound(FOutParams.sNames,1) do
 begin
 FInParams1.sSubKeyName:='Software\Microsoft\Windows\CurrentVersion\Uninstall\'+(string(FOutParams.sNames[i]));
////////////////////////////////////////////////////////////////////////////////////
 FInParams1.sValueName :='DisplayName';
 FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
 if FoutParams1.sValue<>null then
  begin
  nameSoft:='';InstallLocation:='';InstallSource:='';InstallDate:='';DisplayVersion:='';UninstallString:='';Publisher:=''; RegValue:='';
  nameSoft:=string(FoutParams1.sValue);
//////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='InstallLocation';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then InstallLocation:=(string(FoutParams1.sValue));
      ///////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='InstallSource';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then InstallSource:=(string(FoutParams1.sValue));
      //////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='InstallDate';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then
  InstallDate:=(WbemTimeToDateTimeReg(FoutParams1.sValue))
  else InstallDate:='01.01.2001';
      ////////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='DisplayVersion';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then DisplayVersion:=(string(FoutParams1.sValue));
      ///////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='UninstallString';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then UninstallString:=(string(FoutParams1.sValue));
      ////////////////////////////////////////////////////////////////////////////////////////////
  FInParams1.sValueName :='Publisher';
  VariantClear(FOutParams1);
  FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
  if FoutParams1.sValue<>null then Publisher:=(string(FoutParams1.sValue));
   /////////////////////////////////////////////////////////////////////////////////////////////////
  if vartostr(FOutParams.sNames[i])<>'' then RegValue:=(string(FOutParams.sNames[i]));
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////
    with ReceptListProg.ListViewMsi.Items.Add do
    begin
    if Publisher='' then ImageIndex:= frmDomainInfo.BrendProg('other')
    else ImageIndex:= frmDomainInfo.BrendProg(Publisher);
    Caption := inttostr(ReceptListProg.ListViewMsi.Items.Count);
    SubItems.add(nameSoft);
    SubItems.Add(InstallLocation);
    SubItems.Add(InstallSource);
    SubItems.Add(InstallDate);
    SubItems.Add(DisplayVersion);
    SubItems.Add(UninstallString);
    SubItems.Add(Publisher);
    SubItems.Add(RegValue);
    end;
   end;// if
VariantClear(FOutParams1);
 end; //����
VariantClear(FInParams);
VariantClear(FInParams1);
VariantClear(FoutParams);
VariantClear(FoutParams1);
VariantClear(FWbemObjectSet);
VariantClear(FSWbemLocator);
VariantClear(FWMIService);
end;
ReceptListProg.ListViewMsi.Items.EndUpdate;
frmDomainInfo.memo1.Lines.Add('�������� ������ ���� �������� ���������.');

except
on E:Exception do
 begin
 VariantClear(FInParams);
 VariantClear(FInParams1);
 VariantClear(FoutParams);
 VariantClear(FoutParams1);
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 ReceptListProg.ListViewMsi.Items.EndUpdate;
 frmDomainInfo.memo1.Lines.Add('������ ��������� ������ ������������� �������� "'+E.Message+'"');
 end;
end;
OleUnInitialize;
frmDomainInfo.TabSheet5.Tag:=0; //// ������� ��������� �������� ������
NetApiBufferFree(paramUpdate); /// ������� ������
EndThread(result); /// ����� ������
end;

////////////////////////////////////////////////////////////////////////////////////////////
Function LoadListLocalDisk(s,User,Pass:string;ListViewDisk:TlistView):bool;  /// �������� ������ ������
var
CaptionHDD,LabelHDD,IDDiskVolume:string;
sizeHDD,FreeSpace,FileSys,DriveType,statusHDD:string;
z,x:int64;
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FInParams       : OLEVariant;
  FOutParams      : OLEVariant;
  FWbemObject         : OLEVariant;
  oEnum               : IEnumvariant;
  iValue        : LongWord;
  const
  MyTypeDrv: array [0..6] of string = ('�����������','�� ���������','������� ����','��������� ����','������� ����','CD/DVD','RAM ����');
begin
try
  begin
  frmDomainInfo.memo1.Lines.Add('�������� ������ ������...');
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
                 if (FWbemObject.DriveType<>Null) then
                 begin
                  sizeHDD:='';
                  FreeSpace:='';
                  FileSys:='';
                  LabelHDD:='';
                  CaptionHDD:='';
                  IDDiskVolume:='';
                  if FWbemObject.FileSystem<>Null then FileSys:=string( FWbemObject.FileSystem);
                  if (FWbemObject.DriveType<>Null) and (VarIsNumeric(FWbemObject.DriveType)) then DriveType:=string( FWbemObject.DriveType);
                  if FWbemObject.Caption<>null then
                    begin
                     CaptionHDD:=string(FWbemObject.Caption);
                     if Length(CaptionHDD)>15 then
                       CaptionHDD:='';
                    end;
                  if FWbemObject.DeviceID<>null then
                  begin
                   IDDiskVolume:=string(FWbemObject.DeviceID);
                  end;
                  if FWbemObject.Label<>null then
                    begin
                    LabelHDD:=string(FWbemObject.Label);
                     if pos('��������������� ��������',LabelHDD)<>0 then
                      LabelHDD:='Boot';
                    end;
                  if FWbemObject.Status<>null then StatusHDD:=string(FWbemObject.Status)
                  else StatusHDD:='OK';
                     if FWbemObject.Capacity<>null then
                     begin
                     SizeHDD:=vartostr(FWbemObject.Capacity);
                     z:=((strtoint64(SizeHDD)) div 1024)div 1024;
                     if z<=1024 then sizeHDD:=inttostr(z)+' ��'
                     else sizeHDD:=inttostr(z div 1024)+' ��';
                     FreeSpace:=vartostr(FWbemObject.FreeSpace);
                     x:=((strtoint64(FreeSpace)) div 1024)div 1024;
                     if x<=1024 then FreeSpace:=inttostr(x)+' ��'
                     else FreeSpace:=inttostr(x div 1024)+' ��';
                     end;

                    with ListViewDisk.Items.Add do
                    begin
                    if DriveType<>'' then  ImageIndex:= strtoint(DriveType)
                    else  ImageIndex:=0;
                    caption:=inttostr(ListViewDisk.Items.Count-1);
                    SubItems.Add( LabelHDD+' ('+CaptionHDD+')');
                    SubItems.Add(sizeHDD);
                    SubItems.Add(FreeSpace);
                    SubItems.add(MyTypeDrv[strtoint(DriveType)]);
                    SubItems.Add(StatusHDD);
                    SubItems.Add(IDDiskVolume);
                    SubItems.Add(FileSys);
                    end;
                 VariantClear(FWbemObject);
            end;
            end;
oEnum:=nil;
OleUnInitialize;
frmDomainInfo.memo1.Lines.Add('�������� ������ ������ ���������.');
end;
except
      on E:Exception do
        begin
        frmDomainInfo.memo1.Lines.Add('������ �������� ������ ������ - "'+E.Message+'"');
          exit;
        end;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
end;
///////////////////////////////////////////////////////////////////////////////////////////
///
function LoadListSNetworkInterface(S,User,Pass:string;ListViewNI:TListView):bool; /// ������ ������� �����������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
  const
Statnetwork :array [0..12] of string =('���������','�����������','����������'
,'����������','Hardware Not Present','Hardware Disabled','Hardware Malfunction'
,'��� �����������','��������������','�������� ��������������'
,'������ ��������������','���������������� �����','����������� ������� ������');
 begin
    try
      listviewNI.Clear;
      frmDomainInfo.memo1.Lines.Add('�������� ������ ������� �����������');
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2', User, Pass,'','',128);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT Description,MACAddress,AdapterType,InterfaceIndex,NetConnectionID,NetConnectionStatus,NetEnabled,GUID FROM Win32_NetworkAdapter','WQL',wbemFlagForwardOnly);
         oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
            while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                begin
                  with listviewNI.Items.Add do
                    begin
                    Caption := inttostr(listviewNI.Items.Count-1);
                    if FWbemObject.Description<>null then SubItems.Add(vartostr(FWbemObject.Description))   // ok
                    else SubItems.Add('');
                    if FWbemObject.MACAddress<>null then SubItems.Add(vartostr(FWbemObject.MACAddress))   // ok
                    else SubItems.Add('');
                    if FWbemObject.AdapterType<>null then SubItems.Add(vartostr(FWbemObject.AdapterType))   // ok
                    else SubItems.Add('');
                    if (FWbemObject.NetConnectionStatus<>null)  then
                     begin
                     SubItems.Add(Statnetwork[integer(FWbemObject.NetConnectionStatus)]);
                       case integer(FWbemObject.NetConnectionStatus) of
                         1,2,8..11:ImageIndex:=0;
                         7: ImageIndex:=1;
                         else ImageIndex:=2;
                        end;
                      end
                     else
                      begin
                       SubItems.Add('');
                       ImageIndex:=2;
                       end;
                    if FWbemObject.NetEnabled<>null then SubItems.Add(vartostr(FWbemObject.NetEnabled))   // ok
                    else SubItems.Add('');
                    if FWbemObject.InterfaceIndex<>null then SubItems.Add(vartostr(FWbemObject.InterfaceIndex))   // ok
                    else SubItems.Add('');
                    if FWbemObject.NetConnectionID<>null then SubItems.Add(vartostr(FWbemObject.NetConnectionID))   // ok
                    else SubItems.Add('');
                    if FWbemObject.GUID<>null then SubItems.Add(vartostr(FWbemObject.GUID))   // ok
                    else SubItems.Add('');
                    end;
                  VariantClear(FWbemObject);
                end;
             except
               on E:Exception do
               begin
               VariantClear(FWbemObject);
                oEnum:=nil;
                VariantClear(FWbemObjectSet);
                VariantClear(FWMIService);
                VariantClear(FSWbemLocator);
               frmDomainInfo.memo1.Lines.Add('������ �������� ������ ������� ����������� "'+E.Message+'"');
               end;
            end;
            VariantClear(FWbemObject);
            oEnum:=nil;
            VariantClear(FWbemObjectSet);
            VariantClear(FWMIService);
            VariantClear(FSWbemLocator);
            frmDomainInfo.memo1.Lines.Add('�������� ������ ������� ����������� ���������.');
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
 function LoadListAutoStart(S,User,Pass,SystemType:string;ListViewS:TListView):bool; /// ������ ������������
var
{wbemPrivilegeRestore
17 (0x11)
��������� C ++: SE_RESTORE_NAME ������: SeRestorePrivilege
������� ��� ��������: ������������
��������� ��� �������������� ������ � ���������, ���������� �� ACL, ���������� ��� �����.}
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
  FInParams,FInParams1                    : OLEVariant;
  FOutParams,FOutParams1                  : OLEVariant;
  i:integer;
 begin
 ListViewS.Clear;
 frmDomainInfo.memo1.Lines.Add('�������� ������ ������������');
 OleInitialize(nil);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2',User, Pass,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6;
 FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege',true);  //��������� ��� ���������� ����������� ������ � ���������, ���������� �� ������ ACL, ���������� ��� �����.
 FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',true);//��������� ��� �������������� ������ � ��������� ���������� �� ������ ACL, ���������� ��� �����.
 Begin
 try
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,command,Location,user FROM Win32_StartupCommand','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������������
 begin
   with ListViewS.Items.Add do
     begin
     Caption := inttostr(ListViewS.Items.Count-1);          // �
     SubItems.Add(vartostr(FWbemObject.name));       // ��� ���
     Subitems.Add(vartostr(FWbemObject.command));
     Subitems.Add(vartostr(FWbemObject.Location));
     Subitems.Add(vartostr(FWbemObject.user));
     end;
   FWbemObject:=Unassigned;
   end;
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 except
   on E:Exception do
   begin
   frmDomainInfo.memo1.Lines.Add('������ ��������� ������ ������������ "'+E.Message+'"');
   end;
End;
 /////////////////////////////////////////////////////////////////////////////////////
 try
 begin
 FWbemObjectSet:= FWMIService.Get('StdRegProv');
 FInParams     := FWbemObjectSet.Methods_.Item('EnumValues').InParameters.SpawnInstance_();
 FInParams.hDefKey:=HKEY_LOCAL_MACHINE;
 FInParams.sSubKeyName:='SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
 FoutParams    := FWMIService.ExecMethod('StdRegProv', 'EnumValues', FInParams);
 FInParams1     := FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_();
 FInParams1.hDefKey:=HKEY_LOCAL_MACHINE;
 FInParams1.sSubKeyName:='SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
 for i:= VarArrayLowBound(FOutParams.sNames,1) to VarArrayHighBound(FOutParams.sNames,1) do
   begin
   FInParams1.sValueName :=(string(FOutParams.sNames[i]));
   FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
   if (string(FOutParams.Types[i]))='1' then
     begin
     with ListViewS.Items.Add do
       begin
       Caption := inttostr(ListViewS.Items.Count);
       SubItems.add(string(FOutParams.sNames[i]));
       SubItems.add(string(FoutParams1.sValue));
       SubItems.add('HKLM\'+(string(FInParams1.sSubKeyName)));
       SubItems.add('');
       end;
     end;
   end;
 VariantClear(FOutParams);
 VariantClear(FOutParams1);
 VariantClear(FInParams1);


 if SystemType='64-bit' then
 begin
   FInParams.sSubKeyName:='SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run';
   FoutParams    := FWMIService.ExecMethod('StdRegProv', 'EnumValues', FInParams);
   FInParams1     := FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_();
   FInParams1.hDefKey:=HKEY_LOCAL_MACHINE;
   FInParams1.sSubKeyName:='SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run';
   for i:= VarArrayLowBound(FOutParams.sNames,1) to VarArrayHighBound(FOutParams.sNames,1) do
     begin
     FInParams1.sValueName :=(string(FOutParams.sNames[i]));
     FOutParams1    := FWMIService.ExecMethod('StdRegProv', 'GetStringValue', FInParams1);
     if (string(FOutParams.Types[i]))='1' then
       begin
       with listviewS.Items.Add do
         begin
         Caption := inttostr(listviewS.Items.Count);
         SubItems.add(string(FOutParams.sNames[i]));
         SubItems.add(string(FoutParams1.sValue));
         SubItems.add('HKLM\'+(string(FInParams1.sSubKeyName)));
         SubItems.add('');
         end;
       end;
     end;
   VariantClear(FInParams1);
   VariantClear(FoutParams);
   VariantClear(FInParams);
 end;
 end;
 except
   on E:Exception do
   begin
   frmDomainInfo.Memo1.Lines.Add('������ ��������� ������ ������������ "'+E.Message+'"');
   end;
   end;
 oEnum:=nil;
 VariantClear(FInParams);
 VariantClear(FInParams1);
 VariantClear(FoutParams);
 VariantClear(FoutParams1);
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 frmDomainInfo.memo1.Lines.Add('�������� ������ ������������ ���������.');
 end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
function LoadListProfile(S,User,Pass:string;ListViewS:TListView):bool; /// ������ ��������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
  userOwner,SIDstr,LocalPath,LastUseTime,Loaded,Special,Status       :string;
const
 StatProfile :array[0..8] of string =('Ok','��������� �������'
,'������������ �������','','������������ �������','','',''
,'������� ��������� � �� ������������');
 begin
 try
 ListViewS.Clear;
 frmDomainInfo.memo1.Lines.Add('�������� ������ ��������');
 OleInitialize(nil);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(s, 'root\CIMV2',User, Pass,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT SID,LocalPath,LastUseTime,Loaded,Special,Status FROM Win32_UserProfile','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// �������
 begin
 if (FWbemObject.SID<>null) then
   begin
   SIDstr:='';
   userOwner:='';
   LocalPath:='';
   LastUseTime:='';
   Loaded:='';
   Special :='';
   Status:='';
   SIDstr:=(vartostr(FWbemObject.SID));
   // userOwner:=SIDtoUserNameRemotePC(vartostr(FWbemObject.SID),FWMIService);
   //  if( (userOwner='') or (userOwner='Unknown') ) and(frmDomainInfo.LabeledEdit3.text<>'') then /// ���� �� ��������� ������ ������������ �� ����� � �������� �����
   userOwner:=SidToAcountName(vartostr(FWbemObject.SID),s); /// API �������
   if userOwner='' then userOwner:=SIDtoUserNameRemotePC(vartostr(FWbemObject.SID),FWMIService);  /// ����������� �� �����
   if FWbemObject.LocalPath<>null then LocalPath:=(vartostr(FWbemObject.LocalPath));
   if FWbemObject.LastUseTime<>null then LastUseTime:=(WbemTimeToDateAndTime(FWbemObject.LastUseTime));
   if (FWbemObject.Loaded)<>null then Loaded:=(vartostr(FWbemObject.Loaded));
   if (FWbemObject.Special)<>null then Special :=(vartostr(FWbemObject.Special));
   if (FWbemObject.Status)<>null then Status:=(StatProfile[integer(FWbemObject.Status)]) ;
     with ListViewS.Items.Add do
     begin
        Caption := inttostr(ListViewS.Items.Count-1);
        Subitems.Add(SIDstr);
        Subitems.Add(userOwner);
        Subitems.Add(LocalPath);
        Subitems.Add(LastUseTime);
        Subitems.Add(Loaded);
        Subitems.Add(Special);
        Subitems.Add(Status);
     end;
     FWbemObject:=Unassigned;
   end;
 end;
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 VariantClear(FWbemObject);
 OleUnInitialize;
 except
 on E:Exception do
 begin
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 VariantClear(FWbemObject);
 frmDomainInfo.memo1.Lines.Add('������ �������� ������ �������� "'+E.Message+'"');
 end;
end;
frmDomainInfo.memo1.Lines.Add('�������� ������ �������� ���������.');
end;

function GetUserDirectory(NamePC,User,Pass,UserProfile,UsSID:string):string; /// ������� ���������� ���� � ������� ������������
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
  userOwner,SIDstr,LocalPath:string;
 begin
 try
 OleInitialize(nil);
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2',User, Pass,'','',128);
 if (UsSID='') or(UsSID='Unknown') then  // ���� ��� ������ ��� �� �������� ������ ���� �� ����� ������������
 Begin
   FWbemObjectSet:= FWMIService.ExecQuery('SELECT SID,LocalPath FROM Win32_UserProfile','WQL',wbemFlagForwardOnly);
   oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
   begin
   userOwner:=SidToAcountName(string(FWbemObject.SID),NamePC);
   if userOwner='' then userOwner:= SIDtoUserName(string(FWbemObject.SID));
   if userOwner=UserProfile then
     begin
     if FWbemObject.LocalPath<>null then result:=(vartostr(FWbemObject.LocalPath));
       FWbemObject:=Unassigned;
     break;
     end;
   end;
 End
 else // ����� ���� ��� ������, �� ������� �� SID
 if (UsSID<>'') or (UsSID<>'Unknown')  then
  Begin
   FWbemObjectSet:= FWMIService.ExecQuery('SELECT SID,LocalPath FROM Win32_UserProfile WHERE SID= '+'"'+UsSID+'"','WQL',wbemFlagForwardOnly);
   oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   if oEnum.Next(1, FWbemObject, iValue) = 0 then
   begin
    if FWbemObject.LocalPath<>null then result:=(vartostr(FWbemObject.LocalPath));
    FWbemObject:=Unassigned;
   end;
  End;
 if ((UsSID='') or (UsSID='Unknown') )and(UserProfile='')  then // ���� �� ������ ��� � ��� ������������ �� ���������� ���� ��������� ������������
   begin
     FWbemObjectSet:= FWMIService.ExecQuery('SELECT Loaded,Special,LocalPath FROM Win32_UserProfile WHERE Special=False AND Loaded=True','WQL',wbemFlagForwardOnly);
   oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
   begin
     if FWbemObject.LocalPath<>null then result:=(vartostr(FWbemObject.LocalPath));
       FWbemObject:=Unassigned;
     break;
     end;
   end;
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 VariantClear(FWbemObject);
 OleUnInitialize;
 except
 on E:Exception do
 begin
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 VariantClear(FWbemObject);
 frmDomainInfo.memo1.Lines.Add('������ ������ ���� ������� "'+E.Message+'"');
 end;
end;
frmDomainInfo.memo1.Lines.Add('����� ���� ������� ��������.');
end;


function LoadDevicePnP(PcName,User,Pass:string;ListViewPnP:TListView):boolean; // �������� ������ ��������� PnP
var
 FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
begin
 try
 ListViewPnP.Clear;
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(PcName, 'root\CIMV2',User, Pass,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,Manufacturer,Status,DeviceID,Service FROM Win32_PnPEntity','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// �������
 begin
   with ListViewPnP.Items.Add do
   begin
   Caption := inttostr(ListViewPnP.Items.Count-1);
   SubItems.Add(vartostr(FWbemObject.Caption));   // ok
   SubItems.Add(vartostr(FWbemObject.Manufacturer)); // ok
   SubItems.Add(vartostr(FWbemObject.Status)); ///ok
   SubItems.Add(vartostr(FWbemObject.DeviceID)); //ok
   SubItems.Add(vartostr(FWbemObject.Service));
   end;
 VariantClear(FWbemObject);
 FWbemObject:=Unassigned;
 end;
 result:=true;
 except
 on E:Exception do
 begin
 result:=false;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 frmdomaininfo.memo1.Lines.Add('������ �������� ������ ��������� "'+E.Message+'"');
 end;
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
end;

function EnableDisableDevicePnP(PcName,User,Pass,DeviceID:string; EnabOrDis:integer; var reboot:boolean):integer;   // ��������� ��� ���������� PnP ����������
var
 FSWbemLocator   : OLEVariant;
 FWMIService     : OLEVariant;
 FWbemObjectSet  : OLEVariant;
 FWbemObject     : OLEVariant;
 oEnum           : IEnumvariant;
 iValue          : LongWord;
 res:integer;
 s:string;
begin
try
s:=StringReplace(DeviceID,'\','\\',[rfReplaceAll]);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(PcName, 'root\CIMV2',User, Pass,'','',128);     //'+'"'+SelectedDriver+'"'
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT DeviceID FROM Win32_PnPEntity WHERE DeviceID='+'"'+s+'"','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// �������
 begin
  if EnabOrDis=1 then res:=FWbemObject.Enable(reboot);// ��������
  if EnabOrDis=0 then res:=FWbemObject.Disable(reboot); // ���������
 end;
 result:=res;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
except on E: EOleException do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ : '+e.Message);
  result:=e.ErrorCode;
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator)
  end;
end;
end;

end.
