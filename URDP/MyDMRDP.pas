unit MyDMRDP;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.Phys.IBBase, Data.DB, FireDAC.Comp.Client,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt,
  FireDAC.Comp.DataSet,IniFiles,Vcl.Forms, FireDAC.VCLUI.Script,
  FireDAC.VCLUI.Login, FireDAC.Comp.UI, FireDAC.Comp.ScriptCommands,
  FireDAC.Stan.Util, FireDAC.Comp.Script, FireDAC.VCLUI.Wait;

type
  TDataMod = class(TDataModule)
    FDTransactionUpdate: TFDTransaction;
    ConnectionDB: TFDConnection;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDTransactionRead: TFDTransaction;
    FDQueryRead: TFDQuery;
    FDQueryWrite: TFDQuery;
    FDTransactionWrite: TFDTransaction;
    FDGUIxScriptDialog1: TFDGUIxScriptDialog;
    FDScriptCreateTabl: TFDScript;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    procedure DataModuleCreate(Sender: TObject);
 private
    function CreateTablSetRDP(s:string):boolean;
    function TableExists(TableName: string): Boolean;
    function transactionOnOff(OnOff:boolean):boolean;
  public
   function writesetDB(NamePC:string;
ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):boolean;
  procedure clearRdpTabl;
  Function WriteItemDB(NamePC:string):boolean; ///запись в таблицу данных с указанием только имени компа, остальные характеристики считываются с панели настроек

  end;

var
  DataMod: TDataMod;

implementation
 uses urdp;

function bitcolorDepth(index:integer):integer;
begin
case index of
0:result:=8;
1:result:=15;
2:result:=16;
3:result:=24;
4:result:=32;
end;
end;

function booltoint(z:boolean):integer;
begin
  if z then result:=1
  else result:=0;
end;

function inttobool(z:integer):boolean;
begin
  if z=1 then result:=true
  else result:=false;
end;

Function TDataMod.WriteItemDB(NamePC:string):boolean;
var
i:integer;
begin
with FormRDP do
begin
 writesetDB(NamePC, //- имя компа
 bitcolorDepth(ComboBoxColorDepth.ItemIndex), // количество бит на пиксель
 booltoint(CheckBoxBitmapPeristence.Checked),//BitmapPeristence
 booltoint(CheckboxCachePersistenceActive.Checked), //CachePersistenceActive
 SpinBitmapCacheSize.Value,                //BitmapCacheSize
 SpinEditCache16BppSize.Value,           //VirtualCache16BppSize
 SpinEditCache32BppSize.Value,          //VirtualCache32BppSize
 SpinBitmapVirtualCacheSize.Value,             //VirtualCacheSize
 booltoint(CheckBoxDisableCtrlAltDel.Checked),   //DisableCtrlAltDel
 booltoint(CheckBoxDoubleClickDetect.Checked),   // DoubleClickDetect
 booltoint(CheckBoxEnableWindowsKey.Checked),   // EnableWindowsKey
 SpinMinutesToIdleTimeout.Value,            //MinutesToIdleTimeout
 SpinOverallConnectionTimeout.Value,       // OverallConnectionTimeout
 strtoint(LabeledEditRdpPort.Text),      //RdpPort
 CalcPerformanceFlags(''),               //PerformanceFlags
 ComboBoxNetworkConnectionType.ItemIndex, // NetworkConnectionType
 SpinMaxReconnectAttempts.Value,           // MaxReconnectAttempts
 ComboAudioRedirectionMode.ItemIndex,      // AudioRedirectionMode
 ComboBoxAuthLevel.ItemIndex,             //  AuthenticationLevel
 ComboKey.ItemIndex,                        // KeyboardHookMode
 CheckBoxSmartSizing.Checked,         //milovSmartSizing
 checkboxGrabFocusOnConnect.Checked,            //GrabFocusOnConnect
 true,                                          //BandwidthDetection
 CheckEnableAutoReconnect.Checked,                  //EnableAutoReconnect
 CheckBoxDisk.Checked,                       //RedirectDrives
 CheckBoxPrint.Checked,                       //RedirectPrinters
 CheckBoxMouseMode.Checked,               //RelativeMouseMode
 CheckBoxClipboard.Checked,                   //RedirectClipboard
 CheckRedirectDevices.Checked,              //RedirectDevices
 CheckBoxPort.Checked,                       //RedirectPorts
 CheckBoxConToAdmSrv.Checked,                  //ConnectToAdministerServer
 CheckBoxRecAudio.Checked,                 //AudioCaptureRedirectionMode
 CheckBoxEnSuperPan.Checked,               //EnableSuperPan
 CheckBoxCredSsp.Checked                    //EnableCredSspSupport
 )
 end;
end;

procedure TDataMod.clearRdpTabl;
begin
if ConnectionDB.Connected then
  begin
  FDScriptCreateTabl.SQLScripts[0].Name:='ClearDataBase';
  FDScriptCreateTabl.SQLScripts[0].SQL.Clear;
  FDScriptCreateTabl.SQLScripts[0].SQL.Add('DELETE FROM RDP_SET;');
  FDScriptCreateTabl.ExecuteAll;
  end;
end;

function TDataMod.transactionOnOff(OnOff:boolean):boolean;
begin
FDTransactionWrite.Options.AutoCommit:=OnOff;
FDTransactionWrite.Options.AutoStart:=OnOff;
FDTransactionWrite.Options.AutoStop:=OnOff;
end;

function TDataMod.writesetDB(NamePC:string;
ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):boolean;
begin
try
if not ConnectionDB.Connected then
begin
  Result:=false;
  exit;
end;
transactionOnOff(true);
 FDQueryWrite.SQL.Clear;
 FDQueryWrite.SQL.Text:='update or insert into RDP_SET'
+'(PC_NAME,ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,'
+'VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,'
+'DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,'
+'OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,'
+'MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode,'
+'SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,'
+'RedirectDrives,RedirectPrinters,RelativeMouseMode'
+',RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,'
+'AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport) VALUES'
+' (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12,:p13,:p14,:p15,:p16'
+',:p17,:p18,:p19,:p20,:p21,:p22,:p23,:p24,:p25,:p26,:p27,:p28,:p29,:p30,:p31,:p32,:p33,:p34)'
+' MATCHING (PC_NAME)';
FDQueryWrite.params.ParamByName('p1').AsString:=''+NamePC+'';
FDQueryWrite.params.ParamByName('p2').Asinteger:=ColorDepth;
FDQueryWrite.params.ParamByName('p3').Asinteger:=BitmapPeristence;
FDQueryWrite.params.ParamByName('p4').Asinteger:=CachePersistenceActive;
FDQueryWrite.params.ParamByName('p5').Asinteger:=BitmapCacheSize;
FDQueryWrite.params.ParamByName('p6').Asinteger:= VirtualCache16BppSize;
FDQueryWrite.params.ParamByName('p7').Asinteger:=VirtualCache32BppSize;
FDQueryWrite.params.ParamByName('p8').Asinteger:=VirtualCacheSize;
FDQueryWrite.params.ParamByName('p9').Asinteger:= DisableCtrlAltDel;
FDQueryWrite.params.ParamByName('p10').Asinteger:=DoubleClickDetect;
FDQueryWrite.params.ParamByName('p11').Asinteger:=EnableWindowsKey;
FDQueryWrite.params.ParamByName('p12').Asinteger:=MinutesToIdleTimeout;
FDQueryWrite.params.ParamByName('p13').Asinteger:=OverallConnectionTimeout;
FDQueryWrite.params.ParamByName('p14').Asinteger:=RdpPort;
FDQueryWrite.params.ParamByName('p15').Asinteger:=PerformanceFlags;
FDQueryWrite.params.ParamByName('p16').Asinteger:= NetworkConnectionType;
FDQueryWrite.params.ParamByName('p17').Asinteger:= MaxReconnectAttempts;
FDQueryWrite.params.ParamByName('p18').Asinteger:=AudioRedirectionMode;
FDQueryWrite.params.ParamByName('p19').Asinteger:=AuthenticationLevel;
FDQueryWrite.params.ParamByName('p20').Asinteger:=KeyboardHookMode;
FDQueryWrite.params.ParamByName('p21').Asboolean:= SmartSizing;
FDQueryWrite.params.ParamByName('p22').Asboolean:= GrabFocusOnConnect;
FDQueryWrite.params.ParamByName('p23').Asboolean:=BandwidthDetection;
FDQueryWrite.params.ParamByName('p24').Asboolean:= EnableAutoReconnect;
FDQueryWrite.params.ParamByName('p25').Asboolean:= RedirectDrives;
FDQueryWrite.params.ParamByName('p26').Asboolean:= RedirectPrinters;
FDQueryWrite.params.ParamByName('p27').Asboolean:= RelativeMouseMode;
FDQueryWrite.params.ParamByName('p28').Asboolean:= RedirectClipboard;
FDQueryWrite.params.ParamByName('p29').Asboolean:= RedirectDevices;
FDQueryWrite.params.ParamByName('p30').Asboolean:= RedirectPorts;
FDQueryWrite.params.ParamByName('p31').Asboolean:= ConnectToAdministerServer;
FDQueryWrite.params.ParamByName('p32').Asboolean:= AudioCaptureRedirectionMode;
FDQueryWrite.params.ParamByName('p33').Asboolean:= EnableSuperPan;
FDQueryWrite.params.ParamByName('p34').Asboolean:= EnableCredSspSupport;
FDQueryWrite.ExecSQL;
transactionOnOff(false);
result:=true;
except
  on E:Exception do
    begin
    result:=false;
    FormRDP.memo1.Lines.Add('Ошибка записи в таблицу RDP_SET - "'+E.Message+'"');
    end;
  end;
end;


function TDataMod.TableExists(TableName: string): Boolean; /// существует ли таблица
var
Tables: TStrings;
begin
     Tables := TStringList.Create;
     try
      ConnectionDB.GetTableNames('','','',Tables,[osMy],[tkTable],true);
      Result := Tables.IndexOf(TableName) <> -1;
       //FormRDP.Memo1.Lines:=Tables;
     finally
       Tables.Free;
     end;
end;


{$R *.dfm}
{createRDP(NamePC,Domain,UserName,Passwd:string;AutoConnect:boolean;
ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):bool;}
function TDataMod.CreateTablSetRDP(s:string):boolean;
begin
try
if ConnectionDB.Connected then /// если соединение установленно
begin
FDScriptCreateTabl.SQLScripts[0].SQL.Clear;
FDScriptCreateTabl.SQLScripts[0].Name:='ADDTABRDPSET';  //;
FDScriptCreateTabl.SQLScripts[0].SQL.Add('CREATE GENERATOR GEN_RDP_SET_ID START WITH 0 INCREMENT BY 1;'); ///  создаем генератор
FDScriptCreateTabl.SQLScripts[0].SQL.Add('CREATE TABLE RDP_SET (');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('ID_RDP  INTEGER NOT NULL,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('PC_NAME     VARCHAR(100) ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('ColorDepth     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('BitmapPeristence     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('CachePersistenceActive     integer ,');

FDScriptCreateTabl.SQLScripts[0].SQL.Add('BitmapCacheSize    integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('VirtualCache16BppSize     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('VirtualCache32BppSize     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('VirtualCacheSize     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('DisableCtrlAltDel     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('DoubleClickDetect     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('EnableWindowsKey     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('MinutesToIdleTimeout     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('OverallConnectionTimeout     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RdpPort     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('PerformanceFlags     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('NetworkConnectionType     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('MaxReconnectAttempts     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('AudioRedirectionMode     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('AuthenticationLevel     integer ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('KeyboardHookMode     integer ,');
{DisplayConnectionBar,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport}
FDScriptCreateTabl.SQLScripts[0].SQL.Add('SmartSizing   BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('GrabFocusOnConnect    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('BandwidthDetection    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('EnableAutoReconnect   BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RedirectDrives    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RedirectPrinters    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RelativeMouseMode    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RedirectClipboard    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RedirectDevices    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('RedirectPorts    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('ConnectToAdministerServer    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('AudioCaptureRedirectionMode    BOOLEAN ,');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('EnableSuperPan    BOOLEAN ,'); //EnableCredSspSupport
FDScriptCreateTabl.SQLScripts[0].SQL.Add('EnableCredSspSupport    BOOLEAN);');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('ALTER TABLE RDP_SET ADD CONSTRAINT PK_RDP_SET PRIMARY KEY (ID_RDP);');
//////////////////////////// далее триггер
FDScriptCreateTabl.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER RDP_SET_BI FOR RDP_SET');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('as begin  if (new.id_rdp is null) then');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('new.id_rdp = gen_id(GEN_RDP_SET_ID,1); end ^');
FDScriptCreateTabl.SQLScripts[0].SQL.Add('SET TERM; ^');

if FDScriptCreateTabl.ExecuteAll then
  begin
  result:=true;
  end
else
  begin
  result:=false;
  end;
end;
except
  on E:Exception do
    begin
    result:=false;
    FormRDP.memo1.Lines.Add('Ошибка создания  таблицы RDP_SET "'+E.Message+'"');
    end;
  end;
end;


procedure Code(var text: WideString; password: string;  //// процедура кодирования и декодирования файла
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

procedure TDataMod.DataModuleCreate(Sender: TObject);
var
SetIni:Tinifile;
s:widestring;
begin
try
 if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
    begin
    SetInI:=TiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini');
    ConnectionDB.Params.Clear;
    ConnectionDB.Params.Add('server='+SetIni.ReadString('DB','Server','localhost'));
    ConnectionDB.Params.Add('port='+SetIni.ReadString('DB','Port','3050'));
    ConnectionDB.Params.Add('protocol='+SetIni.ReadString('DB','Protocol','local'));
    ConnectionDB.Params.Add('CharacterSet=UTF8');
    ConnectionDB.Params.add('sqlDialect=3');
    ConnectionDB.Params.database:=SetIni.ReadString('DB','Patch',extractfilepath(application.ExeName)+'DB.FDB');
    ConnectionDB.Params.DriverID:=SetIni.ReadString('DB','DrID','FB');
    ConnectionDB.Params.UserName:=SetIni.ReadString('DB','user','SYSDBA');
    s:= widestring(SetIni.ReadString('DB','pass','masterkey'));              //// расщифровка пароля
    code(s,'1234',true);                                               //// расщифровка пароля
    ConnectionDB.Params.Password:=s;
    ConnectionDB.Connected:=SetIni.readBool('db','Connected',true);
    end
  else FormRDP.Memo1.Lines.add('Файл  с настройками не найден, базу данных подключить не удалось!!!');

if ConnectionDB.Connected then
begin
 if not (TableExists('RDP_SET')) then
 begin
  CreateTablSetRDP(''); /// если таблица не найдена то создаем ее.
  WriteItemDB('SETDEFAULT'); /// запись строки по умолчанию
 end
end
 else  FormRDP.Memo1.Lines.add('Создание таблицы завершилось неудачей, соединение с базой данных не установлено');

 Except
   on E: Exception do
   begin
   FormRDP.Memo1.Lines.add('Ошибка при подключении к БД :'+E.Message);

   end;
   end;
SetInI.Free;
end;

end.
