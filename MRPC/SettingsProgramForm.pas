unit SettingsProgramForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  inifiles, Vcl.Buttons, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.Comp.ScriptCommands, FireDAC.Stan.Util,
  FireDAC.Comp.Script, FireDAC.Comp.Client, FireDAC.Phys.IBBase, Data.DB,
  FireDAC.VCLUI.Script, FireDAC.Comp.UI, FireDAC.VCLUI.Login,Themes,Registry,math,
  Vcl.Imaging.pngimage,ShellApi, Vcl.ImgList, Vcl.Menus, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,JvRichEdit;

type
  TSettingsProgram = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    GroupBox1: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    GroupBox2: TGroupBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox9: TCheckBox;
    CheckBox10: TCheckBox;
    CheckBox11: TCheckBox;
    CheckBox12: TCheckBox;
    LabeledEdit4: TLabeledEdit;
    GroupBox3: TGroupBox;
    LabeledEdit5: TLabeledEdit;
    LabeledEdit6: TLabeledEdit;
    GroupBox4: TGroupBox;
    CheckBox13: TCheckBox;
    TabSheet2: TTabSheet;
    GroupBox5: TGroupBox;
    LabeledEdit7: TLabeledEdit;
    SpeedButton1: TSpeedButton;
    GroupBox6: TGroupBox;
    CheckBox14: TCheckBox;
    CheckBox15: TCheckBox;
    CheckBox16: TCheckBox;
    CheckBox17: TCheckBox;
    CheckBox18: TCheckBox;
    CheckBox19: TCheckBox;
    CheckBox20: TCheckBox;
    CheckBox21: TCheckBox;
    CheckBox22: TCheckBox;
    CheckBox23: TCheckBox;
    CheckBox24: TCheckBox;
    FDScript1: TFDScript;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDGUIxScriptDialog1: TFDGUIxScriptDialog;
    CheckBox25: TCheckBox;
    LabeledEdit8: TLabeledEdit;
    LabeledEdit9: TLabeledEdit;
    FDConnection1: TFDConnection;
    SpeedButton4: TSpeedButton;
    CheckBox26: TCheckBox;
    GroupBox7: TGroupBox;
    CheckBox27: TCheckBox;
    CheckBox28: TCheckBox;
    CheckBox29: TCheckBox;
    CheckBox30: TCheckBox;
    ComboBox1: TComboBox;
    Label1: TLabel;
    CheckBox31: TCheckBox;
    GroupBox8: TGroupBox;
    CheckBox32: TCheckBox;
    CheckBox33: TCheckBox;
    GroupBox11: TGroupBox;
    LabeledEdit11: TLabeledEdit;
    CheckBox34: TCheckBox;
    CheckBox35: TCheckBox;
    CheckBox36: TCheckBox;
    LinkLabel1: TLinkLabel;
    CheckBox37: TCheckBox;
    CheckBox38: TCheckBox;
    TabSheet3: TTabSheet;
    GroupBox9: TGroupBox;
    ComboProtocol: TComboBox;
    Label2: TLabel;
    EditDBserver: TLabeledEdit;
    EditDBPort: TLabeledEdit;
    SpeedButton2: TButton;
    SpeedButton3: TButton;
    Button3: TButton;
    ImageListSettings: TImageList;
    Button4: TButton;
    CleanDataBasePopup: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Button5: TButton;
    Button6: TButton;
    PopupUpdateDB: TPopupMenu;
    N30401: TMenuItem;
    N40411: TMenuItem;
    N21301: TMenuItem;
    WindowsOffice1: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    GroupBox10: TGroupBox;
    ButtonCodeError: TButton;
    CheckBox39: TCheckBox;
    ComboKMS: TComboBox;
    IdHTTP1: TIdHTTP;
    LoadManual: TButton;
    CheckBox40: TCheckBox;
    CheckBox41: TCheckBox;
    CheckBox42: TCheckBox;
    CheckBox43: TCheckBox;
    CheckBox44: TCheckBox;
    CheckBox45: TCheckBox;
    ComboBox2: TComboBox;
    Label3: TLabel;
    LabeledEdit10: TLabeledEdit;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CheckBox26Click(Sender: TObject);
    procedure SpeedButton4MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButton4MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CheckBox30Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Update4to41(Sender: TObject);
    procedure reloadLC(Sender: TObject);
    procedure LinkLabel1Click(Sender: TObject);
    procedure LinkLabel2Click(Sender: TObject);
    procedure CheckBox36Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure ComboProtocolSelect(Sender: TObject);
    procedure cleandatabase;
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure WindowsOffice1Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure ButtonCodeErrorClick(Sender: TObject);
    procedure CheckBox39MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure LoadManualClick(Sender: TObject);
    procedure CheckBox39Click(Sender: TObject);  // ��������� ������� ���� ������
      const
  private
    ProgressB:TprogressBar;
    { Private declarations }
    function regeditread(patch,Section:string):string;
    function regeditwrite(patch,Section,Value:string):boolean;
    function regeditwriteInteger(patch,Section:string;Value:integer):boolean;
    function existreg(patch:string):bool;
    Function eventDcom(r:integer):bool; ////  ��������� � ���������� ����������� ������� Dcom
    function clearSetTable(tabName,genName:string):boolean;  // ������� ��������� ������� � ���� ���� �� ����� ����������
    procedure cleanTableTask; // ������� ���� �����
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
    procedure CreateFormElseErrorloadFile;
    procedure memo1URLClick(Sender: TObject; const URLText: string;  Button: TMouseButton);
    function LoadFile(url,path,Fname:string):boolean;  // ������� �������� �����
    function loadFileRun:boolean; //������� �������� ����� �������� � ����������� ������ ��� KMS
    procedure HttpWork(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
    function ExistFileKMS(ListFile:TstringList):boolean;
  public
    { Public declarations }
  end;

var
  SettingsProgram: TSettingsProgram;
  IniSettings:TMeminiFile;
  sm: TStyleManager;
  ThrLic:TThread;
  TrLic:Tthread;
  RootPatch: HKEY;
  peremennaya:string;
  keyboardstring,activeserver:string;
  reloadL:Ttimer;
const
 stringpatch:string ='software\MRPC\MRPC';
 //RootPatch:=HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER

implementation
uses umain,MyDM,unit8,unit9,ErrorLic;
{$R *.dfm}
///////////////////////////////////////////// ������ � ��������
function Log_write(fname, text:string):string;
var f:TStringList;        /// ������� ������ � ��� ����
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


////////////////////////////////////////////////////////////////////////
procedure Code(var text: WideString; password: string;  //// ��������� ����������� � ������������� �����
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
///////////////////////////////////////////////////////////////////// ������ � ��������
function TSettingsProgram.existreg(patch:string):bool;  //// ������������� ���� � �������
var
regFile:TregInifile;
begin
regFile:=Treginifile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then result:=true
else result :=false;
end;

function TSettingsProgram.regeditwriteInteger(patch,Section:string;Value:integer):boolean; /// CurKey - ��������� � �������
var      //// ������ � ������
regFile:TregIniFile;
begin
RegFile:=TregIniFile.Create(KEY_WRITE OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.CreateKey(patch) then
RegFile.Writeinteger(patch,Section,value);
if Assigned(RegFile) then freeandnil(regFile);
end;

function TSettingsProgram.regeditwrite(patch,Section,Value:string):boolean; /// CurKey - ��������� � �������
var      //// ������ � ������
regFile:TregIniFile;
begin
RegFile:=TregIniFile.Create(KEY_WRITE OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.CreateKey(patch) then
RegFile.WriteString(patch,Section,value);
if Assigned(RegFile) then freeandnil(regFile);
end;

function TSettingsProgram.regeditread(patch,Section:string):string;
var      /// ������ �� �������
regFile:TregIniFile;
begin
result:='';
RegFile:=TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then
begin
result:=(regFile.ReadString(patch,Section,''));
end;
if Assigned(RegFile) then freeandnil(regFile);
end;

///////////////////////////////////////////////////
Function TSettingsProgram.eventDcom(r:integer):bool;
var      //// ������ � ������
RegFile:TRegistry;
i:integer;
begin
try
i:=0;
RootPatch:=HKEY_LOCAL_MACHINE;
RegFile:=TregIniFile.Create(KEY_ALL_ACCESS);
RegFile.RootKey:=RootPatch;
if RegFile.OpenKey('SOFTWARE\Microsoft\Ole',true) then
begin
    case r of
    0:  RegFile.WriteInteger('ActivationFailureLoggingLevel',0); /// ���� ���� �� ���������� ��������� �����������
    2:  RegFile.WriteInteger('ActivationFailureLoggingLevel',2); /// ���� ��� �� ���������� ���������� �����������
    1:
       begin                                                      /// ���� ���� �� ��������� ��������
       if RegFile.ValueExists('ActivationFailureLoggingLevel') then
       i:=regFile.ReadInteger('ActivationFailureLoggingLevel')
       else i:=0;  //// ���� ������ � ������� ��� �� �� ��������� DCOM ����� ���
       end;
    end;
RegFile.CloseKey;
end;
if  (r=0) or (r=2) then result:=true;
if r=1 then
begin
  if i=0 then result:=true
  else result:=false;
end;
if Assigned(RegFile) then freeandnil(regFile);
except
    on E: Exception do
      begin
      result:=false;
      if Assigned(RegFile) then freeandnil(regFile);
      frmDomainInfo.Memo1.Lines.Add('������ ��������� ����������� DCOM. - '+e.Message);
      end;
  end;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TSettingsProgram.Button3Click(Sender: TObject);
begin
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=Open'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=LabeledEdit7.Text;
FDConnection1.Params.UserName:=''+LabeledEdit8.Text+'';
FDConnection1.Params.Password:=''+LabeledEdit9.Text+'';
FDConnection1.Connected:=true;
if FDConnection1.Connected then
begin
 showmessage('���������� �����������');
 FDConnection1.Connected:=false;
end
else showmessage('�� �������� ���������� ����������');


end;




procedure TSettingsProgram.Button6Click(Sender: TObject);
var
FDQuery:TFDQuery;
qTransaction:TFDTransaction;
MicrList:Tstringlist;
DBOpenpatch:TOpenDialog;
i:integer;
begin
try
MicrList:=TStringList.Create;

dbopenPatch:=TOpenDialog.Create(self);
dbopenPatch.Name:='';
dbopenPatch.Title:='���� � �������� ��������';
dbopenPatch.InitialDir:=extractfilepath(application.ExeName);

dbopenPatch.Filter:='|*.txt';
if DBOpenpatch.Execute then  MicrList.LoadFromFile(DBOpenpatch.FileName);
qTransaction:= TFDTransaction.Create(nil);
qTransaction.Connection:=DataM.ConnectionDB;
qTransaction.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
qTransaction.Options.AutoCommit:=false;
qTransaction.Options.AutoStart:=false;
qTransaction.Options.AutoStop:=false;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=qTransaction;
FDQuery.Connection:=DataM.ConnectionDB;
for I := 0 to MicrList.Count-1 do
begin
if (MicrList[i]<>'') then
  begin
  qTransaction.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='update or insert into LIC_ERROR'+
  ' (CODE,DESCRIPTION)'+
  ' VALUES (:p1,:p2) MATCHING (CODE,DESCRIPTION)';
  FDQuery.params.ParamByName('p1').AsString:=''+copy(MicrList[i],1,pos('-',MicrList[i])-1)+'';
  FDQuery.params.ParamByName('p2').AsString:=''+copy(MicrList[i],pos('-',MicrList[i])+1,length(MicrList[i]))+'';
  FDQuery.ExecSQL;
  qTransaction.commit;
  end;
end;

finally
FDQuery.Free;
MicrList.Free;
dbopenPatch.Free;
end;
end;

function DescriptionLicStatus:boolean;
var
z:integer;
he,des:string;
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;

function writeError(act:boolean;he,des:string):boolean;
begin
TransactionRead.StartTransaction;
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='update or insert into LIC_ERROR (ACTIVATE,CODE,DESCRIPTION) VALUES (:p1,:p2,:p3)  MATCHING (CODE)';
FDQueryRead.params.ParamByName('p1').Asboolean:=act;
FDQueryRead.params.ParamByName('p2').AsString:=''+he+'';
FDQueryRead.params.ParamByName('p3').AsString:=''+des+'';
FDQueryRead.ExecSQL;
TransactionRead.Commit;
end;
begin
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=SettingsProgram.FDConnection1;
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=SettingsProgram.FDConnection1;
try
writeError(false,'0','������� �� ���� ���������� ����������� � ������� �����.');
writeError(false,'4004F401', '������ ��������� �������, ��� �� ���������� ���� ����������� �������� ��������.');
writeError(false,'4004F040', '������ �������������� ������������ ����������� ��������, ��� ������� ��� �����������, �� �������� ������ ����������� ����� �� ������������� ��������.');
writeError(true,'C004C001', '������ ��������� ���������, ��� ��������� ���� �������� ����������');
writeError(true,'C004C003', '������ ��������� ���������, ��� ��������� ���� �������� ������������');
writeError(true,'C004C017', '������ ��������� ���������, ��� ��������� ���� �������� ��� ������������ ��� ����� ��������������� ������������.');
writeError(true,'C004B100', '������ ��������� ���������, ��� ��� ������� ���������� �� ������� ��������� ���������');
writeError(true,'C004D307', '��������� ����������� ���������� ����� ������� ������� ���������. ����� ��������� ������� ������ ������� ���������, �������������� ��.');
writeError(false,'C004F00F', '�������� �������������� ������������ ������� �� ����� �����������. ������ M /Board ��� ������ ���������� ���������');
writeError(true,'C004F014', '������ �������������� ������������ ����������� ��������, ��� ���� �������� ����������');
writeError(true,'C004C001', '������ ��������� ���������, ��� ������ ������������ ���� ��������.');
writeError(true,'C004C003', '������ ��������� ���������, ��� ��������� ���� �������� ������������.');
writeError(true,'C004B100', '������ ��������� ���������, ��� ���� ��������� �� ����� ���� �����������.');
writeError(true,'C004C008', '������ ��������� ���������, ��� ��������� ���� �������� ���������� ������������.');
writeError(true,'C004C020', '������ ��������� ������ ��������� � ���������� ������� ��� ���������������������� ����� ��������� (MAK).');
writeError(true,'C004C021', '������ ��������� ������ ��������� � ���������� ������� ���������� ��� ���������������������� ����� ���������.');
writeError(true,'C004F009', '������ �������������� ������������ ����������� ������� ��������� �� ��������� ��������� �������.');
writeError(true,'C004F00F', '������ �������������� ������������ ����������� ������� ��������� � ������ �������������� �������� ������������ �� ������� ����������� ����������.');
writeError(true,'C004F014', '������ �������������� ������������ ����������� ������� ��������� � ������������� ����� ��������.');
writeError(true,'C004F02C', '������ �������������� ������������ ����������� ������� ��������� � �������������� ������� ������ ���������� ���������.');
writeError(true,'C004F035', '������ �������������� ������������ ����������� ������� ��������� � ������������� ��������� ���������� � ������� ����� ������������ ���������. ����������� ���� ������� ����.');
writeError(true,'C004F038', '������ �������������� ������������ ����������� ������� ��������� � ������������� ��������� ����������. �������, ������������ � ������ ������ ���������� ������� (KMS), ����� ������������� ��������. ���������� � ������ ���������� ��������������.');
writeError(true,'C004F039', '������ �������������� ������������ ����������� ������� ��������� � ������������� ��������� ����������. ������ ���������� ������� (KMS) ���������.');
writeError(true,'C004F041', '������ �������������� ������������ ����������� ����������, ��� ������ ���������� ������� (KMS) �� ������������. ���������� ������������ ������ KMS.');
writeError(true,'C04F042', '������ �������������� ������������ ����������� ����������, ��� ��������� ������ ���������� ������� (KMS) �� ����� ���� ������������. ');
writeError(true,'C004F050', '������ �������������� ������������ ����������� ������� ��������� � �������������� ����� ����� ��������.');
writeError(true,'C004F051', '������ �������������� ������������ ����������� ������� ��������� � ���������� ����� ����� ��������.');
writeError(true,'C004F064', '������ �������������� ������������ ����������� ������� ��������� � ���������� ��������� ������� ��� ����������� ������.');
writeError(true,'C004F065', '������ �������������� ������������ ����������� ������� ��������� � ������ ���������� � ������ ����������� ��������� ������� ��� ����������� ������.');
writeError(true,'C004F066', '������ �������������� ������������ ����������� ������� ���������, ��� ������� �������� �������� �� �������.');
writeError(true,'C004F069', '������ �������������� ������������ ����������� ������� ��������� � ������������� ��������� ����������. ������ �������������� ������������ ����������� (KMS) ����������, ��� ������� ������� ��� ������� �������� ������������.');
writeError(false,'C004F06B', '������ �������������� ������������ ����������� ����������, ��� ��� ����������� �� ����������� ����������. ������ ���������� ������� (KMS) � ���� ������ �� ��������������');
writeError(true,'C004F074', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ������ ���������� ������� (KMS) ����������');
writeError(false,'C004F075', '������ �������������� ������������ ����������� ��������, ��� �������� ���������� ���������, ��� ��� ������ ���������������');
writeError(true,'C004F304', '������ �������������� ������������ ����������� ��������, ��� ��������� �������� �� �������.');
writeError(true,'C004F305', '������ �������������� ������������ ����������� �������� � ���, ��� � ������� �� ������� �����������, ����������� ������������ �������.');
writeError(true,'C004F30A', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ���������� �� ������������� �������� ��������.');
writeError(true,'C004F30D', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ���������������� ���������.');
writeError(true,'C004F30E', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. �� ������� ����� ����������, ��������������� ���������.');
writeError(true,'C004F30F', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ���������� �� ������������� ��������, ��������� � �������� ������.');
writeError(true,'C004F310', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ���������� �� ������������� �������������� ����� ������� (TPID), ���������� � �������� ������.');
writeError(true,'C004F311', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ����������� ����� �� ����� ���� ����������� ��� ���������.');
writeError(true,'C004F312', '������ �������������� ������������ ����������� ��������, ��� ��� ������� ���������� �� ������� ��������� ���������. ���������� �� ����� ���� �����������, ��� ��� ��� �������� ���� �������� ��������������.');
writeError(true,'80070005', '��� �������. ��� ���������� �������������� �������� ��������� ���������� �����.');
writeError(false,'8007232A', '������ DNS-�������.');
writeError(false,'8007232B', 'DNS-��� �� ����������.');
writeError(false,'800706BA', '������ RPC ����������.');
writeError(false,'8007251D', '������ �� ������� ��� ������� ������� DNS.');
writeError(false,'C004F056','������ �������������� ������������ ����������� ��������, ��� ������� �� ����� ���� ����������� � ������� ������ ���������� �������.');
writeError(true,'C004F034','������ ������������ ���� �������� ��� ���� �������� ��� ������ ������ Windows.');
writeError(false,'C004F057','�� ������� ���������� �������������� ������� �� ������� ACPI.');
writeError(false,'4004FC05','������ �������������� ������������ ����������� ��������, ��� ���������� ����� ���������� �������� ������..');
writeError(false,'4004F00D','������ ��������� ��������.');
writeError(false,'C004FE00','��������� ��������� ����� ������������� ����� ����������� � ���������� ���������� ������ SL.');
writeError(false,'4004F00C','��� ����� �������� ��� ������������� �������� � ����������� ����������. ��� ���������� ��������� ���������� ����������� � ������������� ����.');
finally
FDQueryRead.Free;
TransactionRead.Free;
end;
end;

{LicenseStatus	DWORD	4	��������� ��������������
0 - ���������������
1 - ������������� (������������)
2 � �������� ������ OOB
3 � �������� ������ OOT
4 - NonGenuineGrace
}

procedure TSettingsProgram.LoadManualClick(Sender: TObject);
begin
CreateFormElseErrorloadFile;
end;

procedure TSettingsProgram.Update4to41(Sender: TObject); //���������� � ������ 4 �� 4.1
///keyboardstring:=LabeledEdit10.text;
//activeserver:=Edit3.text;
//peremennaya:=keyboardstring;
//RootPatch:=HKEY_LOCAL_MACHINE;
//regeditwrite(stringpatch,'isOK',keyboardstring);
//showmessage('�������! ������ ������������� ����������.');
//regeditwrite(stringpatch,'srv',Edit3.text);
var
i:byte;
UpdateMyDB:TOpenDialog;
DBUser,Dbpass:string;
mes:integer;
begin
try
if pos(Ansiuppercase('Free'),Ansiuppercase(frmDomainInfo.Caption))<>0 then
begin
  ShowMessage('������� �������� � ������������������ ������ ���������.');
  exit;
end;
if (inventConf) or (InventSoft)  then
begin
  ShowMessage('����� ����������� ��������� ��������������!');
  exit;
end;
UpdateMyDB:=TOpenDialog.Create(self);
UpdateMyDB.Title:='�������� ���� ������ ��� ����������';
UpdateMyDB.Filter:='|*.FDB';  /// ��������� ������ fdb �����
UpdateMyDB.DefaultExt:='FDB'; /// ���������� ����������� ������
if not UpdateMyDB.Execute then begin UpdateMyDB.Free; exit; end;
if not InputQuery('��� �������������� FireBird', '���:', DBUser) then exit;
if not InputQuery('������ �������������� FireBird', '������:', Dbpass) then exit;
  /// ��������� ��� ���������� � ��
FDGUIxScriptDialog1.Caption:='������� ���������� ���� ������ '+ UpdateMyDB.FileName;
///FDGUIxScriptDialog1.options:=[ssAutoHide]; /// �������� ���� �����������
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=UpdateMyDB.FileName;
FDConnection1.Params.UserName:=''+DBUser+'';
FDConnection1.Params.Password:=''+Dbpass+'';
mes:=MessageDlg(
    '������ - '+EditDBserver.Text+#10#13
    +'�������� - '+ComboProtocol.Text+#10#13
    +'���� - '+EditDBPort.text+#10#13
    +'�� - '+UpdateMyDB.FileName+#10#13
    +'�������� ���� ������?'
    , mtConfirmation,[mbYes,mbCancel],0);
  if i=mrCancel then
  begin
    if not DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=true;
    exit;
  end;
FDConnection1.Connected:=true;
FDScript1.SQLScripts[0].SQL.Clear;
FDScript1.SQLScripts[0].Name:='UpdateTable';///////////update for version 4.1
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE MAIN_PC ');///
FDScript1.SQLScripts[0].SQL.Add('ADD STATWINLIC     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('ADD SSTATOFLIC     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('ADD ANTIVIRUS_PRODUCT  VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('ADD ANTIVIRUS_STATUS   INTEGER;');

////////////////////////////////////////////////////////////////////////////////// ����������
if not DataM.TableExists('REGEDIT_KEY') then //���� ������� ��� �� �������
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_REGEDIT_KEY START WITH 0 INCREMENT BY 1;');

FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_LIC_ERROR_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MICROSOFT_LIC_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MICROSOFT_PRODUCT_ID START WITH 0 INCREMENT BY 1;');
//////////////////////////////////////////////////////////////////////////////////
//������� ������� ��� ���������� ������ �������
if not DataM.TableExists('REGEDIT_KEY') then   //���� ������� ��� �� ������� ������� � �������
begin
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE REGEDIT_KEY (  ');
FDScript1.SQLScripts[0].SQL.Add('ID_KEY           INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION_KEY  VARCHAR(1000),');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('HDEFKEY          VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SSUBKEYNAME      VARCHAR(3000),');
FDScript1.SQLScripts[0].SQL.Add('SVALUE           VARCHAR(8190),');
FDScript1.SQLScripts[0].SQL.Add('TYPEKEY          VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SVALUENAME       VARCHAR(200));');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE REGEDIT_KEY ADD CONSTRAINT PK_REGEDIT_KEY PRIMARY KEY (ID_KEY);');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER REGEDIT_KEY_BI0 FOR REGEDIT_KEY');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('AS begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.id_key is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id_key = gen_id(GEN_FOR_REGEDIT_KEY,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');
end;
//////// ������� ���� �������� Microsoft
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MICROSOFT_PRODUCT (');
FDScript1.SQLScripts[0].SQL.Add('ID                 INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION        VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('PRODUCT_ID         VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('KEY_PRODUCT        VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC           VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('PARTIALPRODUCTKEY  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('PRODUCT            VARCHAR(50));');
///������ ��� �������������� ��������� Microsoft
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MICROSOFT_LIC (  ');
FDScript1.SQLScripts[0].SQL.Add('ID                   INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC               VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('WINPRODUCT           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('STATWINLIC           VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('KEYWIN               VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC_WIN         VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPT_LIC_WIN     VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('SOFFICEPRODUCT       VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('SSTATOFLIC           VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('SKEYOFC              VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC_OFFICE      VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPT_LIC_OFFICE  VARCHAR(1500));');
/// ������� � ���������� �������� ��������� � ��������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE LIC_ERROR (');
FDScript1.SQLScripts[0].SQL.Add('ID           INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('CODE         VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('ACTIVATE     BOOLEAN DEFAULT false NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION  VARCHAR(1000));');
//����� ������� ��� ���������� ������� �� ���������

// ������� �������������� ������������ ���������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE ANTIVIRUSPRODUCT (');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC              VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF             VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF_STATUS      VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF_UPDATE      VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_STATUS    VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_UPDATE    VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('FIREWALL            VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('FIREWALL_STATUS     VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE         VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE_STATUS  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE_UPDATE  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('OS_NAME             VARCHAR(250));');///////////////////////////////////////////////////////////////�������
///////////////////////// ��������
FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER LIC_ERROR_BI FOR LIC_ERROR');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('as ');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id = gen_id(gen_lic_error_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MICROSOFT_LIC_BI FOR MICROSOFT_LIC');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('  if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('    new.id = gen_id(gen_microsoft_lic_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MICROSOFT_PRODUCT_BI FOR MICROSOFT_PRODUCT');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('  if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('    new.id = gen_id(gen_microsoft_product_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');



FDScript1.ValidateAll;

if FDScript1.ExecuteAll then
begin
DescriptionLicStatus; // ��������� ������� �� ��������� ��������������
ShowMessage('���� ������ ��������� �������!!!');
end
else
begin
 FDGUIxScriptDialog1.Options:=[ssCallstack,ssConsole];  /// �������� ���� ����� ��� ���������� �������� ����
 ShowMessage('��� ���������� ���� ��������� ������!');
end;
FDScript1.SQLScripts[0].SQL.Clear;
FDConnection1.Connected:=false;
UpdateMyDB.Free;
  except
  on E: Exception do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ ���������� ���� ������. - '+e.Message);
  FDConnection1.Connected:=false;
  if Assigned(UpdateMyDB) then UpdateMyDB.Free;
  end;
  end;
end;





procedure TSettingsProgram.ButtonCodeErrorClick(Sender: TObject);
begin
CodeErrorLic.ShowModal;
end;

procedure TSettingsProgram.Button1Click(Sender: TObject);
var
s:widestring;
begin
try
IniSettings.WriteBool('View','Services',CheckBox2.checked);
FrmDomaininfo.TabSheet3.TabVisible:= CheckBox2.checked; // ������
IniSettings.WriteBool('View','Autoload',CheckBox3.checked);
FrmDomaininfo.TabSheet1.TabVisible:= CheckBox3.checked;// ������������
IniSettings.WriteBool('View','ProgramMSI',CheckBox4.checked);
FrmDomaininfo.TabSheet4.TabVisible:= CheckBox4.checked; // ��������� MSi
IniSettings.WriteBool('View','AllProgram',CheckBox5.checked);
FrmDomaininfo.TabSheet5.TabVisible:= CheckBox5.checked; /// ������� ��� ���������
 IniSettings.WriteBool('View','Drive',CheckBox6.checked);
FrmDomaininfo.TabSheet6.TabVisible:= CheckBox6.checked; /// �����
IniSettings.WriteBool('View','Printer',CheckBox7.checked);
FrmDomaininfo.TabSheet7.TabVisible:= CheckBox7.checked; /// ��������
IniSettings.WriteBool('View','Profiles',CheckBox8.checked);
FrmDomaininfo.TabSheet8.TabVisible:= CheckBox8.checked; /// �������
IniSettings.WriteBool('View','DriverPnP',CheckBox9.checked);
FrmDomaininfo.TabSheet10.TabVisible:= CheckBox9.checked; /// ��������
IniSettings.WriteBool('View','Network',CheckBox10.checked);
FrmDomaininfo.TabSheet11.TabVisible:= CheckBox10.checked; /// Network Interface
IniSettings.WriteBool('View','RDP',CheckBox11.checked);
FrmDomaininfo.TabSheet9.TabVisible:= CheckBox11.checked; /// RDP
IniSettings.WriteBool('View','uRDM',CheckBox12.checked);
FrmDomaininfo.TabSheet12.TabVisible:= CheckBox12.checked; /// uRDM
IniSettings.WriteBool('View','HotFix',CheckBox13.checked);
FrmDomaininfo.TabSheet13.TabVisible:= CheckBox13.checked; /// HotFix
IniSettings.WriteBool('View','Share',CheckBox37.checked);
FrmDomaininfo.TabSheet22.TabVisible:= CheckBox37.checked; /// ������� �������
IniSettings.WriteBool('View','EventWindows',CheckBox38.checked);
FrmDomaininfo.TabSheet23.TabVisible:= CheckBox38.checked; /// ������� windows
///
IniSettings.WriteString('MRD','port',LabelEdEdit4.Text);  /// ���� �������
IniSettings.WriteString('MRD','LocalPort',LabelEdEdit3.Text);
IniSettings.WriteString('MRD','user',LabelEdEdit1.Text);
IniSettings.Writebool('MRD','scan',checkBox34.checked);
s:= widestring(LabelEdEdit2.Text);
code(s,'1234',false);
IniSettings.WriteString('MRD','passwd',string(s));
s:= widestring(LabelEdEdit6.Text);
code(s,'1234',false);
IniSettings.WriteString('ADM','passwd',string(s));
IniSettings.WriteString('ADM','user',LabelEdEdit5.Text);
IniSettings.WriteString('ADM','DC',LabelEdEdit10.Text);  // ���������� ������
///////////////////////////////////////////////////// ���� ������



if not Fileexists(labelEdEdit7.Text) then
showmessage('���� ���� ������ �� ������,'+#10#13+' �������� �� ������� ���� ������ �� �������.'+#10#13+' ��� ���������� �������� ������������� ����������');
inisettings.WriteString('DB','Patch',labelEdEdit7.Text);
inisettings.WriteString('DB','user',labelEdEdit8.Text);
s:= widestring(labelEdEdit9.Text);  /// �������� ������ ��� ����������� � ��
code(s,'1234',false);
inisettings.WriteString('DB','pass',string(s));
inisettings.WriteBool('db','Connected',CheckBox25.checked);
inisettings.WriteString('db','Protocol',ComboProtocol.text);
inisettings.WriteString('db','Server',EditDBServer.Text);
inisettings.WriteString('db','Port',EditDBPort.Text);

////// ���������� �������� ��������� ������������ ��
IniSettings.WriteBool('InvWarning','MAC',CheckBox23.checked);
IniSettings.WriteBool('InvWarning','Inv�',CheckBox22.checked);
IniSettings.WriteBool('InvWarning','OS',CheckBox19.checked);
IniSettings.WriteBool('InvWarning','Mon',CheckBox21.checked);
IniSettings.WriteBool('InvWarning','Net',CheckBox20.checked);
IniSettings.WriteBool('InvWarning','Video',CheckBox18.checked);
IniSettings.WriteBool('InvWarning','HDD',CheckBox17.checked);
IniSettings.WriteBool('InvWarning','RAM',CheckBox16.checked);
IniSettings.WriteBool('InvWarning','M/B',CheckBox15.checked);
IniSettings.WriteBool('InvWarning','Name',CheckBox24.checked);
IniSettings.WriteBool('InvWarning','Proc',CheckBox14.checked);
IniSettings.WriteBool('InvWarning','SMART.OS',CheckBox31.checked);
IniSettings.WriteBool('InvWarning','Proc',true); /// ������ ��������� ��� ��������� ������������
//// �������� ��� �������������� ��������
IniSettings.WriteBool('InvWarning','SoftVer',CheckBox27.checked);
IniSettings.WriteBool('InvWarning','SoftManufacture',CheckBox28.checked);
IniSettings.WriteBool('InvWarning','SoftName',CheckBox29.checked);
// ��������� �������� ���������
invWarning:='';
if IniSettings.ReadBool('InvWarning','MAC',false) then invWarning:=invWarning+'ANSWER_MAC,';
if IniSettings.ReadBool('InvWarning','Inv�',true) then invWarning:=invWarning+'INV_NUMBER,';
if IniSettings.ReadBool('InvWarning','OS',true) then invWarning:=invWarning+'OS_DATEINSTALL,';
if IniSettings.ReadBool('InvWarning','Mon',true) then invWarning:=invWarning+'MONITOR_NAME,';
if IniSettings.ReadBool('InvWarning','Net',false) then invWarning:=invWarning+'NETWORKINTERFACE,';
if IniSettings.ReadBool('InvWarning','Video',true) then invWarning:=invWarning+'VIDEOCARD,';
if IniSettings.ReadBool('InvWarning','HDD',true) then invWarning:=invWarning+'HDD_NAME,HDD_SN,';
if IniSettings.ReadBool('InvWarning','RAM',true) then invWarning:=invWarning+'SUMM_MEM_SIZE,';
if IniSettings.ReadBool('InvWarning','M/B',true) then invWarning:=invWarning+'MONTERBOARD,MONTERBOARD_SN,';
if IniSettings.ReadBool('InvWarning','Proc',true) then invWarning:=invWarning+'PROCESSOR,';
if IniSettings.ReadBool('InvWarning','Name',true) then invWarning:=invWarning+'PC_NAME,';
if IniSettings.ReadBool('InvWarning','SMART.OS',true) then invWarning:=invWarning+'HDD_SMART_OS,';
delete(invWarning,Length(invWarning),1);// ������� ��������� ������� � ������

InvSoftWarning:='';    // ��������� �������� �������������� ��������
if IniSettings.ReadBool('InvWarning','SoftVer',true) then InvSoftWarning:=InvSoftWarning+'SOFT_VERSION,';
if IniSettings.ReadBool('InvWarning','SoftManufacture',true) then InvSoftWarning:=InvSoftWarning+'MANUFACTURE,';
if IniSettings.ReadBool('InvWarning','SoftName',true) then InvSoftWarning:=InvSoftWarning+'SOFT_NAME,';
delete(InvSoftWarning,Length(InvSoftWarning),1);// ������� ��������� ������� � ������
IniSettings.WriteBool('SMART','read',CheckBox32.checked); /// ������� smart
IniSettings.WriteBool('Tray','ico',CheckBox30.checked); //// ������ � ����
IniSettings.WriteBool('Scan','ScanLan',CheckBox33.checked); //// ������������ ��� ������ ���������
IniSettings.WriteString('Style','ST',Combobox1.Text); /// ����� ����������
IniSettings.WriteString('ADM','user',LabelEdEdit5.Text);
IniSettings.WriteString('Scan','timeout',LabelEdEdit11.Text);
IniSettings.WriteInteger('Scan','type',Combobox2.ItemIndex+1);
IniSettings.WriteBool('Scan','135',CheckBox35.checked);   /// ��������� ������ �� ���� RPC 135
IniSettings.WriteBool('Vaccina','Vaccination',CheckBox39.checked); // ������� ��� ��� ������� �� ���������
IniSettings.WriteInteger('Vaccina','Deadline',comboKMS.ItemIndex); // 0 - �������������, 1-�� 180 ����, 2-������� ��������
IniSettings.WriteBool('Vaccina','DeleteKeyKMS',CheckBox40.checked); // �������  ����� ��������� ��� �������� ��������
IniSettings.WriteBool('Vaccina','UnknownStatus',CheckBox41.checked); // ������� ������� ��� ������� ����������.
IniSettings.WriteBool('ConfLAN','UserGroup',CheckBox42.checked); // ��������� ������ ������������� ��������� ������ ������������ ������
IniSettings.WriteBool('ConfLAN','PCGroup',CheckBox43.checked); // ��������� ������ ���������� ��������� ������ ������������ ������
IniSettings.WriteBool('Scan','Rendering',CheckBox44.checked); // ��� ���������� ������������ �� ����� ������������ ����������� ��������� ������ �� ListView
IniSettings.WriteBool('Scan','StopLockUser',CheckBox45.checked); // ��������� ������������ ��� ���������� ������������

pingtimeout:=IniSettings.ReadInteger('Scan','timeout',3000); /// ��������� ������� �����
PingType:=Combobox2.ItemIndex+1;   // ����� ������� ICMP
if PingType<0 then PingType:=0;
if PingType>3 then PingType:=3;

urdmport:=IniSettings.Readstring('MRD','port','48999');  /// ���� uRDM
uRDMScan:=IniSettings.Readbool('MRD','scan',CheckBox34.checked);  /// ����������� �� ������� uRDM
RPCport:= IniSettings.ReadBool('Scan','135',CheckBox35.checked); /// ��������� 135 ����
FrmDomaininfo.labelEdEdit6.Text:=LabelEdEdit4.Text; /// ���� �������
if strtoint(LabelEdEdit3.Text)<>FrmDomaininfo.ClientSocket1.Port then
    FrmDomaininfo.ClientSocket1.Port:=strtoint(LabelEdEdit3.Text);  /// ��������� ����
FrmDomaininfo.LabelEdEdit4.text:=LabelEdEdit1.Text; //// ������������ uRDM
FrmDomaininfo.LabelEdEdit5.text:=LabelEdEdit2.Text; //// ������ ������������ uRDM
FrmDomaininfo.LabelEdEdit1.text:=LabelEdEdit5.Text; //// ������������ �������������
FrmDomaininfo.LabelEdEdit2.text:=LabelEdEdit6.Text; //// ������ �������������
CurrentDC:=LabeledEdit10.Text; // ������������ ��������� ���������� ������
IniSettings.UpdateFile;
  except
    on E: Exception do
      begin
      frmDomainInfo.Memo1.Lines.Add('������ ����������. - '+e.Message);
      end;
  end;
end;

procedure TSettingsProgram.Button2Click(Sender: TObject);
begin
close;
end;



procedure TSettingsProgram.CheckBox26Click(Sender: TObject);
begin
SpeedButton3.Enabled:=CheckBox26.Checked;
end;



procedure TSettingsProgram.CheckBox30Click(Sender: TObject);
begin
Traymin:=CheckBox30.Checked;///// ������������ � ����
end;

procedure TSettingsProgram.CheckBox36Click(Sender: TObject);
begin
if CheckBox36.Tag=1 then exit;  /// ��������� tag ��� ���� ����� ��� ������ ����� ����������� �������� � checkbox � ������� onClick �� ��������� �������� ������� ����� ������ �� �������
case CheckBox36.Checked of
 true:  eventDcom(0);
 false: eventDcom(2);
end;
ShowMessage('��� ���������� ������� ��������� ������������� ����������');
end;

procedure TSettingsProgram.ComboBox1Change(Sender: TObject);
begin
TstyleManager.TrySetStyle(ComboBox1.Text,false);
Button1.Click;
Button2.Click;
end;

procedure TSettingsProgram.ComboProtocolSelect(Sender: TObject);
begin
if ComboProtocol.Text='local' then EditDBServer.Text:='localhost';

end;

procedure TSettingsProgram.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Assigned(IniSettings) then IniSettings.Free; /// ���������� ��� ��������
end;

procedure TSettingsProgram.reloadLC(Sender: TObject);
begin
TrLic:=unit9.Lcheck.Create(true);
TrLic.FreeOnTerminate := true;  //// ��� �������� ��������������� ������
TrLic.start;
reloadL.Interval:=10000*RandomRange(22,58);
if Assigned(reloadL) then
  begin
  reloadL.Tag:=reloadL.Tag+1;
  if reloadL.tag>RandomRange(5,10) then
    begin
    reloadL.Enabled:=false;
    FreeAndNil(reloadL);
    end;
  end;
end;


procedure TSettingsProgram.FormCreate(Sender: TObject);
var
i,z:integer;
begin
if not Assigned(reloadL) then
begin
reloadL:=Ttimer.Create(frmDomainInfo);
reloadL.Name:='reloadL';
z:=RandomRange(22,58);
reloadL.Interval:=10000*z;
reloadL.OnTimer:=reloadLC;
end;
FDPhysFBDriverLink1.VendorLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
PageControl1.TabIndex:=0;
sm := TStyleManager.Create;
  for i := 0 to Length(sm.StyleNames)-1 do
    ComboBox1.Items.Add(sm.StyleNames[i]);
IniSettings:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
combobox1.Text:=IniSettings.ReadString('Style','ST','Windows');
IniSettings.Free;
RootPatch:=HKEY_LOCAL_MACHINE;
peremennaya:=regeditread(stringpatch,'isOK');
activeserver:=regeditread(stringpatch,'srv');
PageControl1.TabIndex:=0;
end;

procedure TSettingsProgram.FormShow(Sender: TObject);
var
s:widestring;
i:integer;
begin
SpeedButton3.Enabled:=false;
CheckBox26.Checked:=False;
//if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
  begin
  IniSettings:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  CheckBox1.checked:=true;  //// ��������
  CheckBox2.checked:=IniSettings.ReadBool('View','Services',true);
  CheckBox3.checked:=IniSettings.ReadBool('View','Autoload',true);
  CheckBox4.checked:=IniSettings.ReadBool('View','ProgramMSI',true);
  CheckBox5.checked:=IniSettings.ReadBool('View','AllProgram',true);
  CheckBox6.checked:=IniSettings.ReadBool('View','Drive',true);
  CheckBox7.checked:=IniSettings.ReadBool('View','Printer',true);
  CheckBox8.checked:=IniSettings.ReadBool('View','Profiles',true);
  CheckBox9.checked:=IniSettings.ReadBool('View','DriverPnP',true);
  CheckBox10.checked:=IniSettings.ReadBool('View','Network',true);
  CheckBox11.checked:=IniSettings.ReadBool('View','RDP',true);
  CheckBox12.checked:=IniSettings.ReadBool('View','uRDM',true);
  CheckBox13.checked:=IniSettings.ReadBool('View','HotFix',true);
  CheckBox37.checked:=IniSettings.ReadBool('View','Share',true);
  CheckBox38.checked:=IniSettings.ReadBool('View','EventWindows',true);
  CheckBox39.checked:=IniSettings.ReadBool('Vaccina','Vaccination',false);
  ComboKMS.ItemIndex:=IniSettings.ReadInteger('Vaccina','Deadline',1); // 0 - �������������, 1-�� 180 ����, 2-������� ��������
  CheckBox40.checked:=IniSettings.ReadBool('Vaccina','DeleteKeyKMS',false); // ������� ��� ��� ���� �������� ��� �������� ��������
  CheckBox41.checked:=IniSettings.ReadBool('Vaccina','UnknownStatus',false);  // ������� ������� ��� ���������� ������ � ������� ����������
  LabelEdEdit4.text:=IniSettings.ReadString('MRD','port','48999');
  LabelEdEdit3.text:=IniSettings.ReadString('MRD','LocalPort','48998');
  LabelEdEdit1.text:=IniSettings.ReadString('MRD','user','Administrator');
    s:= widestring(IniSettings.ReadString('MRD','passwd',''));  //// ����������� ������
  code(s,'1234',true);                                   //// ����������� ������
  LabelEdEdit2.text:=s;
  LabelEdEdit5.text:=IniSettings.ReadString('ADM','user','');
    s:= widestring(IniSettings.ReadString('ADM','passwd',''));  //// ����������� ������
  code(s,'1234',true);                                   //// ����������� ������
  LabelEdEdit6.text:=s;
  LabeledEdit10.Text:=IniSettings.ReadString('ADM','DC',''); // ���������� ������ ��������� � ����������

  CheckBox24.checked:=IniSettings.ReadBool('InvWarning','Name',true);
  CheckBox23.checked:=IniSettings.ReadBool('InvWarning','MAC',false);
  CheckBox22.checked:=IniSettings.ReadBool('InvWarning','Inv�',true);
  CheckBox19.checked:=IniSettings.ReadBool('InvWarning','OS',true);
  CheckBox21.checked:=IniSettings.ReadBool('InvWarning','Mon',true);
  CheckBox20.checked:=IniSettings.ReadBool('InvWarning','Net',false);
  CheckBox18.checked:=IniSettings.ReadBool('InvWarning','Video',true);
  CheckBox17.checked:=IniSettings.ReadBool('InvWarning','HDD',true);
  CheckBox16.checked:=IniSettings.ReadBool('InvWarning','RAM',true);
  CheckBox15.checked:=IniSettings.ReadBool('InvWarning','M/B',true);
  CheckBox31.checked:=IniSettings.ReadBool('InvWarning','SMART.OS',true);
  CheckBox27.checked:=IniSettings.ReadBool('InvWarning','SoftVer',true);
  CheckBox28.checked:=IniSettings.ReadBool('InvWarning','SoftManufacture',true);
  CheckBox29.checked:=IniSettings.ReadBool('InvWarning','SoftName',true);
  CheckBox42.checked:=IniSettings.ReadBool('ConfLAN','UserGroup',true); // ��������� ������ ������������� ��������� ������ ������������ ������
  CheckBox43.checked:=IniSettings.ReadBool('ConfLAN','PCGroup',true);  // ��������� ������ ���������� ��������� ������ ������������ ������
  IniSettings.WriteBool('ConfLAN','PCGroup',CheckBox43.checked); // ��������� ������ ���������� ��������� ������ ������������ ������
  CheckBox14.checked:=true; /// ������ �������� �������� ��������� ������������
  CheckBox24.checked:=IniSettings.ReadBool('InvWarning','Name',true);
  CheckBox25.checked:=IniSettings.ReadBool('db','Connected',true);
  LabelEdEdit7.Text:=IniSettings.ReadString('DB','Patch',extractfilepath(application.ExeName)+'DB.FDB');
  LabelEdEdit8.text:=IniSettings.ReadString('DB','user','SYSDBA');
  ComboProtocol.text:=IniSettings.ReadString('DB','Protocol','local');
  EditDBServer.Text:= IniSettings.ReadString('DB','Server','localhost');
  EditDBPort.Text:= IniSettings.ReadString('DB','Port','3050');

  s:= widestring(IniSettings.ReadString('DB','pass','masterkey'));  //// ����������� ������
  code(s,'1234',true);                                   //// ����������� ������
  LabelEdEdit9.text:=s;
  CheckBox30.checked:=IniSettings.ReadBool('Tray','ico',false);
  CheckBox32.checked:=IniSettings.ReadBool('SMART','read',false);
  CheckBox33.checked:=IniSettings.ReadBool('Scan','ScanLan',false); /// ������ ������������ ��� ������ ���������
  LabelEdEdit11.text:=IniSettings.ReadString('Scan','timeout','3000');
  ComboBox2.ItemIndex:=(IniSettings.ReadInteger('Scan','type',1))-1;
  CheckBox34.checked:=IniSettings.ReadBool( 'MRD','scan',true);
 CheckBox35.checked:=IniSettings.Readbool('Scan','135',true);
 CheckBox44.checked:=IniSettings.Readbool('Scan','Rendering',false);// ��� ���������� ������������ �� ����� ������������ ����������� ��������� ������ �� ListView
 CheckBox45.checked:=IniSettings.Readbool('Scan','StopLockUser',false); // ��������� ������������ ��� ���������� ������������


  //////////////////////////////////////////
  CheckBox36.Tag:=1; ///////////// ��� ����, �������� ������ � onclick
  CheckBox36.checked:=eventDcom(1); ///////
  CheckBox36.Tag:=0; ////////////
  ////////////////////////////////////////////////
  end ;

end;

procedure TSettingsProgram.LinkLabel1Click(Sender: TObject);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar('https://blogs.msdn.microsoft.com/carlos/2017/04/14/distributedcom-event-logs-2/'), nil, nil, SW_SHOWNORMAL);
  except
  frmDomainInfo.memo1.Lines.Add('������ ��� �������� ����������')
  end;
end;

procedure TSettingsProgram.LinkLabel2Click(Sender: TObject);
begin
try
//shellAPI.ShellExecute(0, 'Open', PChar('mailto:mrpc@mail.ru?subject=������������ ��������&body=��� ID:'+memo1.Text), nil, nil, SW_SHOWNORMAL);
  except
    //memo1.Lines.Add('������ ��� �������� ����������')
  end;
end;

procedure TSettingsProgram.N1Click(Sender: TObject);
begin
cleandatabase; //������� ��
end;

procedure TSettingsProgram.N2Click(Sender: TObject);
begin
DataM.CleanDeleteTask(''); //������� ���� �� ��������������� �����
end;

procedure TSettingsProgram.N3Click(Sender: TObject); // �������� ������� �������������� ��
begin
clearSetTable('MAIN_PC_SOFT','GEN_MAIN_PC_SOFT_ID');
clearSetTable('SOFT_PC','GEN_SOFT_PC_ID');
end;

procedure TSettingsProgram.N4Click(Sender: TObject);
begin
clearSetTable('MAIN_PC','GEN_MAIN_PC_ID');
end;

procedure TSettingsProgram.N5Click(Sender: TObject);
begin
clearSetTable('CONFIG_PC','GEN_CONFIG_PC');
clearSetTable('ALL_HARDWARE','GEN_ALL_HARDWARE_ID');
end;

procedure TSettingsProgram.N6Click(Sender: TObject); //������� ���� �����
begin
cleanTableTask; // �������� ������ � ��������.
clearSetTable('TABLE_TASK','GEN_FOR_TABLE_TASK'); // ������� ������� ��� ����� + ����� ����������
end;

procedure TSettingsProgram.N7Click(Sender: TObject);
begin
clearSetTable('ANTIVIRUSPRODUCT','');
end;

procedure TSettingsProgram.N8Click(Sender: TObject);
begin
clearSetTable('REGEDIT_KEY','GEN_FOR_REGEDIT_KEY');
end;

procedure TSettingsProgram.SpeedButton1Click(Sender: TObject);
var
DBOpenpatch:TOpenDialog;
begin
try
if SettingsProgram.findComponent('DBOpenpatch')=nil then
begin
dbopenPatch:=TOpenDialog.Create(self);
dbopenPatch.Name:='dbopenPatch';
dbopenPatch.Title:='���� ���� ������';
dbopenPatch.InitialDir:=extractfilepath(application.ExeName);
end;

dbopenPatch.Filter:='|*.FDB';
//dbopenPatch.Files.Text:='DB';
if DBOpenpatch.Execute then
  begin
  LabeledEdit7.Text:=DBOpenpatch.FileName;
  showmessage('��������� ��������� � ������������� ����������!');
  end;
if Assigned(dbopenPatch) then freeandnil(dbopenPatch);
 except
    on E: Exception do
      begin
      frmDomainInfo.Memo1.Lines.Add('������ . - '+e.Message);
      end;
  end;
end;

procedure TSettingsProgram.SpeedButton2Click(Sender: TObject);
var
saveMyDB:TsaveDialog;
DBUser,Dbpass:string;
step:integer;
begin
try
if pos(Ansiuppercase('Free'),Ansiuppercase(frmDomainInfo.Caption))<>0 then
begin
  ShowMessage('������� �������� � ������������������ ������ ���������.');
  exit;
end;
saveMyDB:=TsaveDialog.Create(self);
saveMyDB.Title:='���������� �������� ����� ���� ������';
saveMyDB.Filter:='|*.FDB';  /// ��������� ������ fdb �����
saveMyDB.DefaultExt:='FDB'; /// ���������� ����������� ������
if not saveMyDB.Execute then begin saveMyDB.Free; exit; end;
if not InputQuery('��� �������������� FireBird', '���:', DBUser) then exit;
if not InputQuery('������ �������������� FireBird', '������:', Dbpass) then exit;
//DBUser:=InputBox('��� �������������� ������� FireBird (��� ���������)', '�������������:', 'sysdba');
//Dbpass:=InputBox('������ �������������� ������� FireBird (��� ���������)', '������:', 'masterkey');
//FDScript1.Params.Add.Name:='FBuser';
//FDScript1.Params.Add.Name:='FBpass';
//FDScript1.Params.Add.Name:='MyDBName';
//FDScript1.Params[0].Value:= 'sysdba';
//FDScript1.Params[1].Value:= 'masterkey';
//FDScript1.Params[2].Value:= 'N:\NEWDB\DBTest.FDB';
//FDScript1.Params[0].Size:=6;
//FDScript1.Params[1].Size:=9;
//FDScript1.Params[2].Size:=19;
//FDScript1.Params.ParamByName('FBuser').Value:= 'sysdba';
//FDScript1.Params.ParamByName('FBpass').Value:= 'masterkey';
//FDScript1.Params.ParamByName('MyDBName').Value:= 'N:\NEWDB\DBTest.FDB';
//FDScript1.Params.Add.Name:='FBpass';
//FDScript1.Params.Add.AsString:='''masterkey''';
//FDScript1.Params.Add.Name:='MyDBName';
//FDScript1.Params.Add.AsString:='''N:\NEWDB\DBTest.FDB''';
//FDConnection1.LoginDialog.Caption:='������������� FireBird';
//FDConnection1.Params.Add('OpenMode=open');
FDGUIxScriptDialog1.Caption:='������� �������� ���� ������ '+saveMyDB.FileName;
//FDGUIxScriptDialog1.options:=[ssAutoHide];  �������� ���� �����������
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.add('RoleName=rdb$admin'); /// ���������� ����� ��� ����
FDConnection1.Params.database:=saveMyDB.FileName;
FDConnection1.Params.UserName:=''+DBUser+'';
FDConnection1.Params.Password:=''+Dbpass+'';
FDConnection1.Connected:=true;

//frmDomainInfo.Memo1.Lines.Add(FDConnection1.Params.UserName);
//frmDomainInfo.Memo1.Lines.Add(FDConnection1.Params.Password);

FDScript1.ScriptOptions.ClientLib:=extractfilepath(application.ExeName)+'\fbclient.dll';
FDScript1.SQLScripts[0].SQL.Clear;
FDScript1.SQLScripts[0].Name:='CreateDataBase';
// ������� ���� ������
{FDScript1.SQLScripts[0].SQL.Add('SET SQL DIALECT 3;');
FDScript1.SQLScripts[0].SQL.Add('SET NAMES UTF8;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DATABASE '+saveMyDB.FileName);
FDScript1.SQLScripts[0].SQL.Add('USER '+DBUser+' PASSWORD '+Dbpass);
FDScript1.SQLScripts[0].SQL.Add('PAGE_SIZE 16384');
FDScript1.SQLScripts[0].SQL.Add('DEFAULT CHARACTER SET UTF8;');}
//FDScript1.SQLScripts[0].SQL.Add('CONNECT '+saveMyDB.FileName);
//FDScript1.SQLScripts[0].SQL.Add('user '+DBUser+' password '+Dbpass+';');
/// ������� ������
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN "USER" AS VARCHAR(100) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN MONTERBOARD AS VARCHAR(250) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN PROC AS VARCHAR(250) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN MANE_OBJ AS VARCHAR(100) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN MEMORY AS VARCHAR(300) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN VIDEO AS VARCHAR(250) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN HDD AS VARCHAR(1000) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN OS AS VARCHAR(200) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN NETWORK_INTF AS VARCHAR(300) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN SOFT_STR AS VARCHAR(5000) CHARACTER SET WIN1251 COLLATE WIN1251;');
FDScript1.SQLScripts[0].SQL.Add('CREATE DOMAIN ID_ID AS INTEGER;');
/// ������� ����������
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MAIN_PC_ID START WITH 1 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_RUN_PROC_MSI_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_CONFIG_PC START WITH 1 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE GENERATOR GEN_MAIN_PC_SOFT_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE GENERATOR GEN_SOFT_PC_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_ALL_HARDWARE_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_TABLE_TASK START WITH 0 INCREMENT BY 1;');

///����� ������ 4.1
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_REGEDIT_KEY START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_LIC_ERROR_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MICROSOFT_LIC_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MICROSOFT_PRODUCT_ID START WITH 0 INCREMENT BY 1;');
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_MY_PROCESS_ID START WITH 0 INCREMENT BY 1;');

/// �������� �������  MAIN_PC
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MAIN_PC (');
FDScript1.SQLScripts[0].SQL.Add('PC_ID       INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('PC_NAME     VARCHAR(50) ,');
FDScript1.SQLScripts[0].SQL.Add('DATE_INV    DATE,');
FDScript1.SQLScripts[0].SQL.Add('INV_NUMBER  VARCHAR(50) ,');
FDScript1.SQLScripts[0].SQL.Add('RESULT_INV  VARCHAR(100) DEFAULT ''NO'',');
FDScript1.SQLScripts[0].SQL.Add('ERROR_INV   VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('PC_OS       VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('ANSWER_MAC       VARCHAR(30),');
FDScript1.SQLScripts[0].SQL.Add('HDD_SMART_OS       HDD /* HDD = VARCHAR(1000) */, ');
FDScript1.SQLScripts[0].SQL.Add('HDD_MY_SMART       HDD /* HDD = VARCHAR(1000) */,');
 ///////////////////////////////////////////////////////////////////////// ������ ��� ������ 3.0
FDScript1.SQLScripts[0].SQL.Add('CUR_USER_NAME  VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('CUR_DOMAIN     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('OTHER_NAME     VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('CUR_IP_ADRS    VARCHAR(50), ');
FDScript1.SQLScripts[0].SQL.Add('EXPT_PC        VARCHAR(50) DEFAULT 0, ');
FDScript1.SQLScripts[0].SQL.Add('URDM_CLIENT    VARCHAR(50),');
//////////////////////////////////////////////////////���������� ��� ������ 4.1
FDScript1.SQLScripts[0].SQL.Add('STATWINLIC     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('SSTATOFLIC     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_PRODUCT  VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_STATUS   INTEGER );');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE MAIN_PC ADD CONSTRAINT PK_MAIN_PC PRIMARY KEY (PC_ID);');

/// �������� ������� CONFIG_PC
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE CONFIG_PC (');
FDScript1.SQLScripts[0].SQL.Add('PC_ID                    INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('PC_NAME                  "USER" /* "USER" = VARCHAR(100) */,');
FDScript1.SQLScripts[0].SQL.Add('MONTERBOARD              MONTERBOARD /* MONTERBOARD = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('MONTERBOARD_SN           MONTERBOARD /* MONTERBOARD = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('MONTERBOARD_MANUFACTURE  MONTERBOARD /* MONTERBOARD = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add(' PROCESSOR                PROC /* PROC = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('PROCESSOR_CORE           MANE_OBJ /* MANE_OBJ = VARCHAR(100) */,');
FDScript1.SQLScripts[0].SQL.Add(' PROCESSOR_LOGPROC        PROC /* PROC = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('PROCESSOR_SPEED          MANE_OBJ /* MANE_OBJ = VARCHAR(100) */,');
FDScript1.SQLScripts[0].SQL.Add('PROCESSOR_ARCH           MANE_OBJ /* MANE_OBJ = VARCHAR(100) */,');
FDScript1.SQLScripts[0].SQL.Add('PROCESSOR_SOKET          PROC /* PROC = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('COUNT_PROC               INTEGER DEFAULT 0,');
FDScript1.SQLScripts[0].SQL.Add('MEMORY_SIZE              MEMORY /* MEMORY = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('MEMORY_TYPE              MEMORY /* MEMORY = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('SUMM_MEM_SIZE            VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('COUNT_MEM                INTEGER DEFAULT 1,');
FDScript1.SQLScripts[0].SQL.Add('VIDEOCARD                VIDEO /* VIDEO = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('VIDEOCARD_MEM            VIDEO DEFAULT ''0'' /* VIDEO = VARCHAR(250) */,');
FDScript1.SQLScripts[0].SQL.Add('COUNT_VIDEOCARD          INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('HDD_NAME                 HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_TYPE                 HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_SIZE                 HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_SN                   HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_INTERFACETYPE        HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_FIRMWARE             HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_SMART_OS             HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_MY_SMART             HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('HDD_CUR_TEMP             HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('COUNT_HDD                INTEGER DEFAULT 0,');
FDScript1.SQLScripts[0].SQL.Add('OS_NAME                  OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('OS_VER                   OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('OS_KEY                   OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('OS_TYPE                  OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('INV_NUMBER               OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('OS_SP                    OS /* OS = VARCHAR(200) */,');
FDScript1.SQLScripts[0].SQL.Add('USER_NAME                "USER" /* "USER" = VARCHAR(100) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORKINTERFACE         NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_MAC              NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_SPEED            NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_IP               NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_COUNT_IP         VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_MASK             NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_GATEWAY          NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_COUNT_GATEWAY    VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_DHCP             NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_COUNT_DHCP       VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_DNS              NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_COUNT_DNS        VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_WINS             NETWORK_INTF /* NETWORK_INTF = VARCHAR(300) */,');
FDScript1.SQLScripts[0].SQL.Add('NETWORK_COUNT_WINS       VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('COUNT_NETWORK            INTEGER DEFAULT 0,');
FDScript1.SQLScripts[0].SQL.Add('RESULT_INVENT            VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('OS_TIMEINSTALL           TIME,');
FDScript1.SQLScripts[0].SQL.Add('OS_DATEINSTALL           DATE,');
FDScript1.SQLScripts[0].SQL.Add('MONITOR_HW               VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('MONITOR_NAME             VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('MONITOR_MANUF            VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('MONITOR_DPI              VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('MONITOR_COUNT            VARCHAR(100) DEFAULT ''0'',');
FDScript1.SQLScripts[0].SQL.Add('DATE_INV                 DATE,');
FDScript1.SQLScripts[0].SQL.Add('TIME_INV                 TIME,');
FDScript1.SQLScripts[0].SQL.Add('DVDROM_NAME              HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('DVDROM_TYPE              HDD /* HDD = VARCHAR(1000) */,');
FDScript1.SQLScripts[0].SQL.Add('DVDROM_COUNT             VARCHAR(5) DEFAULT ''0'',');
FDScript1.SQLScripts[0].SQL.Add('SOUND_NAME               VARCHAR(1500),');
FDScript1.SQLScripts[0].SQL.Add('SOUND_COUNT              VARCHAR(5) DEFAULT ''0'',');
FDScript1.SQLScripts[0].SQL.Add('ANSWER_MAC               VARCHAR(30),');
FDScript1.SQLScripts[0].SQL.Add('PRINTER_NAME             VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('PRINTER_COUNT            INTEGER ,');
/////////////////////////////////////////////////////////////////////////// ������ ��� ������ 3.0
FDScript1.SQLScripts[0].SQL.Add('BIOS_DESCRIPTION         VARCHAR(500),'); //
FDScript1.SQLScripts[0].SQL.Add('BIOS_COUNT               VARCHAR(3) ,');  //
FDScript1.SQLScripts[0].SQL.Add('BIOSSN                   VARCHAR(200) ,'); //
FDScript1.SQLScripts[0].SQL.Add('SMBIOSBIOSVERSION        VARCHAR(200) ,'); // ver 3.0
FDScript1.SQLScripts[0].SQL.Add('BIOSSTATUS               VARCHAR(50) ,'); //
FDScript1.SQLScripts[0].SQL.Add('BIOSPRIMARY              VARCHAR(10) ,'); //
FDScript1.SQLScripts[0].SQL.Add('BIOSMANUFAC              VARCHAR(300),');//
FDScript1.SQLScripts[0].SQL.Add('BIOSCHARACTERISTICS      VARCHAR(500) );');
/////////////////////////////////////////////////////////////////////////////
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE CONFIG_PC ADD CONSTRAINT PK_CONFIG_PC PRIMARY KEY (PC_ID);');
/// ������� ������� START_PROC_MSI
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE START_PROC_MSI ( ID_PROC INTEGER NOT NULL,  PATCH_PROC VARCHAR(1000),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_PROC VARCHAR(10), DESCRIPTION_PROC VARCHAR(1000),  KEY_MSI VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('BEFOREINSTALLCOPY  BOOLEAN, DELETEAFTERINSTALL  BOOLEAN,TYPEOPERATION INTEGER,');///
FDScript1.SQLScripts[0].SQL.Add('PATHCREATE   VARCHAR(1024),FILEORFOLDER VARCHAR(20), FILESOURSE_PROC  VARCHAR(1000));  ');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE START_PROC_MSI ADD PRIMARY KEY (ID_PROC);');
/// ������� ������� MAIN_PC_SOFT
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MAIN_PC_SOFT (');
FDScript1.SQLScripts[0].SQL.Add('PC_ID       INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('PC_NAME     VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('RESUL_INV   VARCHAR(200) DEFAULT ''NO'',');
FDScript1.SQLScripts[0].SQL.Add('DATE_INV    DATE,');
FDScript1.SQLScripts[0].SQL.Add('TIME_INV    TIME,');
FDScript1.SQLScripts[0].SQL.Add('INST_SOFT   SOFT_STR /* SOFT_STR = VARCHAR(5000) */, ');
FDScript1.SQLScripts[0].SQL.Add('ERROR_INV   VARCHAR(700),');
FDScript1.SQLScripts[0].SQL.Add(' COUNT_SOFT  INTEGER);');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE MAIN_PC_SOFT ADD CONSTRAINT PK_MAIN_PC_SOFT PRIMARY KEY (PC_ID);');
////  ������� ������� SOFT_PC
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE SOFT_PC (');
FDScript1.SQLScripts[0].SQL.Add('ID              INTEGER NOT NULL,');
//FDScript1.SQLScripts[0].SQL.Add('PC_NAME         VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SOFT_NAME       VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('DIRECT_INSTALL  VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('SOURCE_INST     VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('DATE_INSTALL    VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SOFT_VERSION    VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('UNINSTALL_STR   VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('MANUFACTURE     VARCHAR(400),');
FDScript1.SQLScripts[0].SQL.Add('DATEINVENT      DATE );');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE SOFT_PC ADD CONSTRAINT PK_SOFT_PC PRIMARY KEY (ID);');
////  ������� ������� ALL_HARDWARE
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE ALL_HARDWARE (');
FDScript1.SQLScripts[0].SQL.Add('ID  ID_ID NOT NULL /* ID_ID = INTEGER */,');
FDScript1.SQLScripts[0].SQL.Add('NAME_HARDWARE  VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_HARDWARE  VARCHAR(20)');
FDScript1.SQLScripts[0].SQL.Add(');');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE ALL_HARDWARE ADD CONSTRAINT PK_ALL_HARDWARE_1 PRIMARY KEY (ID);');
////////////// ������� ������� TABLE_TASK
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE TABLE_TASK (');
FDScript1.SQLScripts[0].SQL.Add('ID_TABLE          INTEGER NOT NULL,');  //
FDScript1.SQLScripts[0].SQL.Add('NAME_TABLE        VARCHAR(500) NOT NULL,'); //
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION_TASK  VARCHAR(500),'); //
FDScript1.SQLScripts[0].SQL.Add('PC_RUN            VARCHAR(250),'); //
FDScript1.SQLScripts[0].SQL.Add('STATUS_TASK       VARCHAR(50),'); //
FDScript1.SQLScripts[0].SQL.Add('COUNT_PC          INTEGER,');//
FDScript1.SQLScripts[0].SQL.Add('CURRENT_TASK      VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('START_STOP        VARCHAR(20) DEFAULT 0,');
FDScript1.SQLScripts[0].SQL.Add('WORKS_THREAD      VARCHAR(20),');
FDScript1.SQLScripts[0].SQL.Add('USER_NAME         VARCHAR(256),');
FDScript1.SQLScripts[0].SQL.Add('PASS_USER         VARCHAR(256),');
FDScript1.SQLScripts[0].SQL.Add('SAVE_PASS         BOOLEAN DEFAULT false );');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE TABLE_TASK ADD CONSTRAINT PK_TABLE_TASK PRIMARY KEY (ID_TABLE);');
////////////////////////////// ����� ������ 4.1
//������� ������� ��� ���������� ������ �������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE REGEDIT_KEY (  ');
FDScript1.SQLScripts[0].SQL.Add('ID_KEY           INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION_KEY  VARCHAR(1000),');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('HDEFKEY          VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SSUBKEYNAME      VARCHAR(3000),');
FDScript1.SQLScripts[0].SQL.Add('SVALUE           VARCHAR(8190),');
FDScript1.SQLScripts[0].SQL.Add('TYPEKEY          VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('SVALUENAME       VARCHAR(200));');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE REGEDIT_KEY ADD CONSTRAINT PK_REGEDIT_KEY PRIMARY KEY (ID_KEY);');
//////// ������� ���� �������� Microsoft
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MICROSOFT_PRODUCT (');
FDScript1.SQLScripts[0].SQL.Add('ID                 INTEGER NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION        VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('PRODUCT_ID         VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('KEY_PRODUCT        VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC           VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('PARTIALPRODUCTKEY  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('PRODUCT            VARCHAR(50));');
///������ ��� �������������� ��������� Microsoft
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MICROSOFT_LIC (  ');
FDScript1.SQLScripts[0].SQL.Add('ID                   INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC               VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('WINPRODUCT           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('STATWINLIC           VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('KEYWIN               VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC_WIN         VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPT_LIC_WIN     VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('SOFFICEPRODUCT       VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('SSTATOFLIC           VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('SKEYOFC              VARCHAR(200),');
FDScript1.SQLScripts[0].SQL.Add('TYPE_LIC_OFFICE      VARCHAR(150),');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPT_LIC_OFFICE  VARCHAR(1500));');
/// ������� � ���������� �������� ��������� � ��������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE LIC_ERROR (');
FDScript1.SQLScripts[0].SQL.Add('ID           INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('CODE         VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('ACTIVATE     BOOLEAN DEFAULT false NOT NULL,');
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION  VARCHAR(1000));');
//����� ������� ��� ���������� ������� �� ���������
// ������� �������������� ������������ ���������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE ANTIVIRUSPRODUCT (');
FDScript1.SQLScripts[0].SQL.Add('NAMEPC              VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF             VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF_STATUS      VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('WIN_DEF_UPDATE      VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS           VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_STATUS    VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTIVIRUS_UPDATE    VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('FIREWALL            VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('FIREWALL_STATUS     VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE         VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE_STATUS  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('ANTISPYWARE_UPDATE  VARCHAR(250),');
FDScript1.SQLScripts[0].SQL.Add('OS_NAME             VARCHAR(250));');

///////////////////////////////////////////////////////////////// ������� ��� �������� ������ � ����� ������� ��������� ����������
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE MY_PROCESS (');
FDScript1.SQLScripts[0].SQL.Add('"KEY"         INTEGER NOT NULL,');  //
FDScript1.SQLScripts[0].SQL.Add('PATH_FILE     VARCHAR(300),'); //
FDScript1.SQLScripts[0].SQL.Add('ARG_FILE      VARCHAR(300),'); //
FDScript1.SQLScripts[0].SQL.Add('NAME_PROCESS  VARCHAR(150),'); //
FDScript1.SQLScripts[0].SQL.Add('LOG_PASS      BOOLEAN);');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE MY_PROCESS ADD PRIMARY KEY ("KEY");');

//// ��������
FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MAIN_PC_BI FOR MAIN_PC');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 4');
FDScript1.SQLScripts[0].SQL.Add('as begin');
FDScript1.SQLScripts[0].SQL.Add(' if (new.pc_id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.pc_id = gen_id(gen_main_pc_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER CONFIG_PC FOR CONFIG_PC');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('AS begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.pc_id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.pc_id = gen_id(gen_config_pc,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^ ');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER START_PROC_MSI_BI FOR START_PROC_MSI');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.id_proc is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id_proc = gen_id(gen_run_proc_msi_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MAIN_PC_SOFT_BI FOR MAIN_PC_SOFT ');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('as begin ');
FDScript1.SQLScripts[0].SQL.Add('if (new.pc_id is null) then ');
FDScript1.SQLScripts[0].SQL.Add('new.pc_id = gen_id(gen_main_pc_soft_id,1); ');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER SOFT_PC_BI FOR SOFT_PC');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as begin  if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id = gen_id(gen_soft_pc_id,1); end ^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER ALL_HARDWARE_BI FOR ALL_HARDWARE');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as begin if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id = gen_id(gen_all_hardware_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end ^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER TABLE_TASK_BI0 FOR TABLE_TASK ');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('AS begin ');
FDScript1.SQLScripts[0].SQL.Add('if (new.id_table is null) then ');
FDScript1.SQLScripts[0].SQL.Add('new.id_table = gen_id(GEN_FOR_TABLE_TASK,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');
///////////////////////// ����� ������ 4.1
FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER LIC_ERROR_BI FOR LIC_ERROR');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('as ');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id = gen_id(gen_lic_error_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MICROSOFT_LIC_BI FOR MICROSOFT_LIC');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('  if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('    new.id = gen_id(gen_microsoft_lic_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MICROSOFT_PRODUCT_BI FOR MICROSOFT_PRODUCT');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('as');
FDScript1.SQLScripts[0].SQL.Add('begin');
FDScript1.SQLScripts[0].SQL.Add('  if (new.id is null) then');
FDScript1.SQLScripts[0].SQL.Add('    new.id = gen_id(gen_microsoft_product_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER REGEDIT_KEY_BI0 FOR REGEDIT_KEY');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0');
FDScript1.SQLScripts[0].SQL.Add('AS begin');
FDScript1.SQLScripts[0].SQL.Add('if (new.id_key is null) then');
FDScript1.SQLScripts[0].SQL.Add('new.id_key = gen_id(GEN_FOR_REGEDIT_KEY,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER MY_PROCESS_BI FOR MY_PROCESS ');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('AS begin ');
FDScript1.SQLScripts[0].SQL.Add('if (new."KEY" is null) then ');
FDScript1.SQLScripts[0].SQL.Add('new."KEY" = gen_id(gen_my_process_id,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

/// ����� �������  ��� sysdba
{FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON CONFIG_PC TO '+DBUser+' WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON MAIN_PC TO '+DBUser+' WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON START_PROC_MSI TO '+DBUser+' WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_CONFIG_PC TO '+DBUser+' WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_MAIN_PC_ID TO '+DBUser+' WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_RUN_PROC_MSI_ID TO '+DBUser+' WITH GRANT OPTION;');
}
/// ������� ��� �������� ������������ � ���� ��� ����� �� ����
{FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER USER SKRBLOG PASSWORD '+'''skrblog'''+';');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON CONFIG_PC TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON MAIN_PC TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON START_PROC_MSI TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON MAIN_PC_SOFT TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON SOFT_PC TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON ALL_HARDWARE TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON TABLE_TASK TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON ANTIVIRUSPRODUCT TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON LIC_ERROR TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON MICROSOFT_LIC TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON MICROSOFT_PRODUCT TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT ALL ON REGEDIT_KEY TO SKRBLOG WITH GRANT OPTION;');

FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_CONFIG_PC TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_MAIN_PC_ID TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add('GRANT USAGE ON SEQUENCE GEN_RUN_PROC_MSI_ID TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_MAIN_PC_SOFT_ID TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_SOFT_PC_ID TO TRIGGER SOFT_PC_BI;');/// ���������� ���� ��� ������� � �������
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_SOFT_PC_ID TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_ALL_HARDWARE_ID TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_FOR_TABLE_TASK TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE GEN_FOR_REGEDIT_KEY TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE gen_microsoft_product_id TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE gen_microsoft_lic_id TO SKRBLOG WITH GRANT OPTION;');
FDScript1.SQLScripts[0].SQL.Add ('GRANT USAGE ON SEQUENCE gen_lic_error_id TO SKRBLOG WITH GRANT OPTION;');}

FDScript1.ValidateAll;

if FDScript1.ExecuteAll then
begin
DescriptionLicStatus; // ��������� ������� �� ��������� ��������������
ShowMessage('���� ������ ������� �������!!!');
end
else
begin
 FDGUIxScriptDialog1.Options:=[ssCallstack,ssConsole];  /// �������� ���� ����� ��� ���������� �������� ����
ShowMessage('��� �������� ���� ��������� ������!');
end;
saveMyDB.Free;
FDConnection1.Connected:=false;
  except
    on E: Exception do
      begin
      frmDomainInfo.Memo1.Lines.Add('������ �������� ����� ���� ������. - '+e.Message);
      FDConnection1.Connected:=false;
      if  Assigned(saveMyDB) then  saveMyDB.Free;
      end;
  end;

end;

procedure TSettingsProgram.cleanTableTask; // ������� ���� �����
var
i:byte;
FDQueryDeleteTask,FDQDEL:TFDQuery;
TransactionDeleteTask: TFDTransaction;
begin
try
if (inventConf) or (InventSoft)  then
begin
  ShowMessage('����� �������� ��������� ��������������!');
  exit;
end;

i:=MessageDLG('�������� ������??? - '+LabeledEdit7.Text,mtConfirmation, mbOKCancel, 0);
if i=mrOk then
  begin
  /// ��������� ��� ���������� � ��
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=LabeledEdit7.Text;
FDConnection1.Params.UserName:=''+LabeledEdit8.Text+'';
FDConnection1.Params.Password:=''+LabeledEdit9.Text+'';
FDConnection1.Connected:=true;

try
TransactionDeleteTask:= TFDTransaction.Create(nil);
TransactionDeleteTask.Connection:=FDConnection1;
TransactionDeleteTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionDeleteTask.Options.AutoCommit:=false;
TransactionDeleteTask.Options.AutoStart:=false;
TransactionDeleteTask.Options.AutoStop:=false;
FDQueryDeleteTask:= TFDQuery.Create(nil);
FDQueryDeleteTask.Transaction:=TransactionDeleteTask;
FDQueryDeleteTask.Connection:=FDConnection1;
FDQDEL:= TFDQuery.Create(nil);
FDQDEL.Transaction:=TransactionDeleteTask;
FDQDEL.Connection:=FDConnection1;
TransactionDeleteTask.StartTransaction; // ��������
FDQueryDeleteTask.SQL.clear;
FDQueryDeleteTask.SQL.Text:='SELECT NAME_TABLE FROM TABLE_TASK';
FDQueryDeleteTask.Open;
while not FDQueryDeleteTask.Eof do
begin
  FDQDEL.SQL.Clear;
  FDQDEL.SQL.Text:='DROP TABLE '+FDQueryDeleteTask.FieldByName('NAME_TABLE').AsString;
  FDQDEL.ExecSQL;
  FDQueryDeleteTask.Next;
end;
TransactionDeleteTask.Commit;
finally
FDQueryDeleteTask.Close;
FDQueryDeleteTask.Free;
FDQDEL.Close;
FDQDEL.Free;
TransactionDeleteTask.Free;
end;

FDConnection1.Connected:=false;
FDConnection1.Close;
  end;
  except
  on E: Exception do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ ������� �����. - '+e.Message);
  FDConnection1.Connected:=false;
  FDConnection1.Close;
  end;
  end;
if not DataM.ConnectionDB.Connected then Datam.connectDataBase; // ����� �������  ���� ���������� �� ����������� �� ��������������� ���
end;


procedure TSettingsProgram.cleandatabase; // ������� ���� ���� ������
var
i:byte;
FDQueryDeleteTask,FDQDEL:TFDQuery;
TransactionDeleteTask: TFDTransaction;
begin
try
if (inventConf) or (InventSoft)  then
begin
  ShowMessage('����� �������� ��������� ��������������!');
  exit;
end;
  /// ��������� ��� ���������� � ��
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;

i:=MessageDLG('�������� ��??? - '+LabeledEdit7.Text,mtConfirmation, mbOKCancel, 0);
if i=mrOk then
  begin
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=LabeledEdit7.Text;
FDConnection1.Params.UserName:=''+LabeledEdit8.Text+'';
FDConnection1.Params.Password:=''+LabeledEdit9.Text+'';
FDConnection1.Connected:=true;

try
TransactionDeleteTask:= TFDTransaction.Create(nil);
TransactionDeleteTask.Connection:=FDConnection1;
TransactionDeleteTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionDeleteTask.Options.AutoCommit:=false;
TransactionDeleteTask.Options.AutoStart:=false;
TransactionDeleteTask.Options.AutoStop:=false;
FDQueryDeleteTask:= TFDQuery.Create(nil);
FDQueryDeleteTask.Transaction:=TransactionDeleteTask;
FDQueryDeleteTask.Connection:=FDConnection1;
FDQDEL:= TFDQuery.Create(nil);
FDQDEL.Transaction:=TransactionDeleteTask;
FDQDEL.Connection:=FDConnection1;
TransactionDeleteTask.StartTransaction; // ��������
FDQueryDeleteTask.SQL.clear;
FDQueryDeleteTask.SQL.Text:='SELECT NAME_TABLE FROM TABLE_TASK';
FDQueryDeleteTask.Open;
while not FDQueryDeleteTask.Eof do
begin
  FDQDEL.SQL.Clear;
  FDQDEL.SQL.Text:='DROP TABLE '+FDQueryDeleteTask.FieldByName('NAME_TABLE').AsString;
  FDQDEL.ExecSQL;
  FDQueryDeleteTask.Next;
end;
TransactionDeleteTask.Commit;
finally
FDQueryDeleteTask.Close;
FDQueryDeleteTask.Free;
FDQDEL.Close;
FDQDEL.Free;
TransactionDeleteTask.Free;
end;
FDScript1.SQLScripts[0].Name:='ClearDataBase';
FDScript1.SQLScripts[0].SQL.Clear;
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM MAIN_PC;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM CONFIG_PC;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM START_PROC_MSI;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM MAIN_PC_SOFT;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM SOFT_PC;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM ALL_HARDWARE;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM TABLE_TASK;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM ANTIVIRUSPRODUCT;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM MICROSOFT_LIC;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM MICROSOFT_PRODUCT;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM REGEDIT_KEY;');
FDScript1.SQLScripts[0].SQL.Add('DELETE FROM MY_PROCESS;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_MAIN_PC_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_RUN_PROC_MSI_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_CONFIG_PC RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_MAIN_PC_SOFT_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_SOFT_PC_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_ALL_HARDWARE_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_FOR_TABLE_TASK RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_FOR_REGEDIT_KEY RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_MICROSOFT_LIC_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_MICROSOFT_PRODUCT_ID RESTART;');
FDScript1.SQLScripts[0].SQL.Add('ALTER SEQUENCE GEN_MY_PROCESS_ID RESTART;');
FDScript1.ValidateAll;
if FDScript1.ExecuteAll then
begin
FDGUIxScriptDialog1.Options:=[ssAutoHide];
showmessage('������� ���� ����������� �������.');
end
else showmessage('������� ���� ����������� �������.');

FDConnection1.Connected:=false;
FDConnection1.Close;
  end;
  except
  on E: Exception do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ ������� ���� ������. - '+e.Message);
  FDConnection1.Connected:=false;
  FDConnection1.Close;
  end;
  end;
CheckBox26.Checked:=false;
end;

function TSettingsProgram.clearSetTable(tabName,genName:string):boolean; // ������� ��������� ������� � ���� ���� �� ����������
var
FDQueryDeleteTask:TFDQuery;
TransactionDeleteTask: TFDTransaction;
i:integer;
begin

i:=MessageDLG('�������� ������???',mtConfirmation, mbOKCancel, 0);
if i=mrcancel then  exit;
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;
begin
try
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=LabeledEdit7.Text;
FDConnection1.Params.UserName:=''+LabeledEdit8.Text+'';
FDConnection1.Params.Password:=''+LabeledEdit9.Text+'';
FDConnection1.Connected:=true;
TransactionDeleteTask:= TFDTransaction.Create(nil);
TransactionDeleteTask.Connection:=FDConnection1;
TransactionDeleteTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionDeleteTask.Options.AutoCommit:=false;
TransactionDeleteTask.Options.AutoStart:=false;
TransactionDeleteTask.Options.AutoStop:=false;
FDQueryDeleteTask:= TFDQuery.Create(nil);
FDQueryDeleteTask.Transaction:=TransactionDeleteTask;
FDQueryDeleteTask.Connection:=FDConnection1;
TransactionDeleteTask.StartTransaction; // ��������
FDQueryDeleteTask.SQL.clear;
FDQueryDeleteTask.SQL.Text:='DELETE FROM '+TabName+';';
//FDQueryDeleteTask.ParamByName('a').asString:=TabName;
FDQueryDeleteTask.ExecSQL;
TransactionDeleteTask.Commit;
if genName<>'' then
begin
TransactionDeleteTask.StartTransaction; // ��������
FDQueryDeleteTask.SQL.clear;
FDQueryDeleteTask.SQL.Text:='ALTER SEQUENCE '+genName+' RESTART;';
//FDQueryDeleteTask.ParamByName('a').asString:=genName;
FDQueryDeleteTask.ExecSQL;
TransactionDeleteTask.Commit;
end;
finally
FDQueryDeleteTask.Close;
FDQueryDeleteTask.Free;
TransactionDeleteTask.Free;
end;
if not DataM.ConnectionDB.Connected then Datam.connectDataBase; // ����� �������  ���� ���������� �� ����������� �� ��������������� ���
end;
end;

procedure TSettingsProgram.SpeedButton4MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
LabeledEdit9.PasswordChar:=#0;
end;

procedure TSettingsProgram.SpeedButton4MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
LabeledEdit9.PasswordChar:=#7;
end;

procedure TSettingsProgram.SpeedButton5Click(Sender: TObject); // ���������� � ������ 2,1 �� 3,0
var
i:byte;
UpdateMyDB:TOpenDialog;
DBUser,Dbpass:string;
mes:integer;
begin
try
if pos(Ansiuppercase('Free'),Ansiuppercase(frmDomainInfo.Caption))<>0 then
begin
  ShowMessage('������� �������� � ������������������ ������ ���������.');
  exit;
end;
if (inventConf) or (InventSoft)  then
begin
  ShowMessage('����� ����������� ��������� ��������������!');
  exit;
end;
UpdateMyDB:=TOpenDialog.Create(self);
UpdateMyDB.Title:='�������� ���� ������ ��� ����������';
UpdateMyDB.Filter:='|*.FDB';  /// ��������� ������ fdb �����
UpdateMyDB.DefaultExt:='FDB'; /// ���������� ����������� ������
if not UpdateMyDB.Execute then begin UpdateMyDB.Free; exit; end;
if not InputQuery('��� �������������� FireBird', '���:', DBUser) then exit;
if not InputQuery('������ �������������� FireBird', '������:', Dbpass) then exit;
  /// ��������� ��� ���������� � ��
FDGUIxScriptDialog1.Caption:='������� ���������� ���� ������ '+ UpdateMyDB.FileName;
///FDGUIxScriptDialog1.options:=[ssAutoHide]; /// �������� ���� �����������
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=UpdateMyDB.FileName;
FDConnection1.Params.UserName:=''+DBUser+'';
FDConnection1.Params.Password:=''+Dbpass+'';
mes:=MessageDlg(
    '������ - '+EditDBserver.Text+#10#13
    +'�������� - '+ComboProtocol.Text+#10#13
    +'���� - '+EditDBPort.text+#10#13
    +'�� - '+UpdateMyDB.FileName+#10#13
    +'�������� ���� ������?'
    , mtConfirmation,[mbYes,mbCancel],0);
  if i=mrCancel then exit;
FDConnection1.Connected:=true;
FDScript1.SQLScripts[0].SQL.Clear;
FDScript1.SQLScripts[0].Name:='UpdateTableMain';///////////update for version 3.0
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE MAIN_PC ADD CUR_USER_NAME VARCHAR(50),');///
FDScript1.SQLScripts[0].SQL.Add('ADD CUR_DOMAIN     VARCHAR(100),');
FDScript1.SQLScripts[0].SQL.Add('ADD OTHER_NAME     VARCHAR(50),');
FDScript1.SQLScripts[0].SQL.Add('ADD CUR_IP_ADRS    VARCHAR(50), ');
FDScript1.SQLScripts[0].SQL.Add('ADD EXPT_PC        VARCHAR(50) DEFAULT 0, ');
FDScript1.SQLScripts[0].SQL.Add('ADD URDM_CLIENT    VARCHAR(50); ');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE CONFIG_PC ADD BIOS_DESCRIPTION         VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('ADD BIOS_COUNT               VARCHAR(3) ,');  //
FDScript1.SQLScripts[0].SQL.Add('ADD BIOSSN                   VARCHAR(200) ,'); //
FDScript1.SQLScripts[0].SQL.Add('ADD SMBIOSBIOSVERSION        VARCHAR(200) ,'); //
FDScript1.SQLScripts[0].SQL.Add('ADD BIOSSTATUS               VARCHAR(50) ,'); //
FDScript1.SQLScripts[0].SQL.Add('ADD BIOSPRIMARY              VARCHAR(10) ,'); //
FDScript1.SQLScripts[0].SQL.Add('ADD BIOSMANUFAC              VARCHAR(300),');//
FDScript1.SQLScripts[0].SQL.Add('ADD BIOSCHARACTERISTICS      VARCHAR(500);');
FDScript1.ValidateAll;

if FDScript1.ExecuteAll then
begin
ShowMessage('���� ������ ��������� �������!!!');
end
else
begin
 FDGUIxScriptDialog1.Options:=[ssCallstack,ssConsole];  /// �������� ���� ����� ��� ���������� �������� ����
 ShowMessage('��� ���������� ���� ��������� ������!');
end;
FDScript1.SQLScripts[0].SQL.Clear;
FDConnection1.Connected:=false;
UpdateMyDB.Free;
  except
  on E: Exception do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ ���������� ���� ������. - '+e.Message);
  FDConnection1.Connected:=false;
  if Assigned(UpdateMyDB) then UpdateMyDB.Free;
  end;
  end;
end;

procedure TSettingsProgram.WindowsOffice1Click(Sender: TObject);
begin
clearSetTable('MICROSOFT_LIC','GEN_MICROSOFT_LIC_ID');
clearSetTable('MICROSOFT_PRODUCT','GEN_MICROSOFT_PRODUCT_ID');
end;

///////////////////////////////////////////////////////////////

procedure TSettingsProgram.Button4Click(Sender: TObject); // ���������� �� � ������ 3 �� 4�
var
i:byte;
UpdateMyDB:TOpenDialog;
DBUser,Dbpass:string;
mes:integer;
begin
try
if pos(Ansiuppercase('Free'),Ansiuppercase(frmDomainInfo.Caption))<>0 then
begin
  ShowMessage('������� �������� � ������������������ ������ ���������.');
  exit;
end;
if (inventConf) or (InventSoft)  then
begin
  ShowMessage('����� ����������� ��������� ��������������!');
  exit;
end;
UpdateMyDB:=TOpenDialog.Create(self);
UpdateMyDB.Title:='�������� ���� ������ ��� ����������';
UpdateMyDB.Filter:='|*.FDB';  /// ��������� ������ fdb �����
UpdateMyDB.DefaultExt:='FDB'; /// ���������� ����������� ������
if not UpdateMyDB.Execute then begin UpdateMyDB.Free; exit; end;
if not InputQuery('��� �������������� FireBird', '���:', DBUser) then exit;
if not InputQuery('������ �������������� FireBird', '������:', Dbpass) then exit;
  /// ��������� ��� ���������� � ��
FDGUIxScriptDialog1.Caption:='������� ���������� ���� ������ '+ UpdateMyDB.FileName;
///FDGUIxScriptDialog1.options:=[ssAutoHide]; /// �������� ���� �����������
if DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=false;
FDConnection1.Params.Clear; /// ������ ����������  ��� ���� ������
FDConnection1.Params.Add('server='+EditDBserver.Text);
FDConnection1.Params.Add('protocol='+ComboProtocol.Text);
FDConnection1.Params.Add('port='+EditDBPort.text);
FDConnection1.Params.Add('CharacterSet=UTF8');
FDConnection1.Params.add('sqlDialect=3');
FDConnection1.Params.add('OpenMode=OpenOrCreate'); /// ���� ����� ������� ��� �������� � ��
FDConnection1.Params.DriverID:='FB';
FDConnection1.Params.database:=UpdateMyDB.FileName;
FDConnection1.Params.UserName:=''+DBUser+'';
FDConnection1.Params.Password:=''+Dbpass+'';
mes:=MessageDlg(
    '������ - '+EditDBserver.Text+#10#13
    +'�������� - '+ComboProtocol.Text+#10#13
    +'���� - '+EditDBPort.text+#10#13
    +'�� - '+UpdateMyDB.FileName+#10#13
    +'�������� ���� ������?'
    , mtConfirmation,[mbYes,mbCancel],0);
  if i=mrCancel then
  begin
    if not DataM.ConnectionDB.Connected then DataM.ConnectionDB.Connected:=true;
    exit;
  end;
FDConnection1.Connected:=true;
FDScript1.SQLScripts[0].SQL.Clear;
FDScript1.SQLScripts[0].Name:='UpdateTableMain';///////////update for version 4.0
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE START_PROC_MSI ADD BEFOREINSTALLCOPY   BOOLEAN,');///
FDScript1.SQLScripts[0].SQL.Add('ADD DELETEAFTERINSTALL  BOOLEAN,');
FDScript1.SQLScripts[0].SQL.Add('ADD TYPEOPERATION    INTEGER,');
FDScript1.SQLScripts[0].SQL.Add('ADD PATHCREATE   VARCHAR(1024), ');
FDScript1.SQLScripts[0].SQL.Add('ADD FILEORFOLDER VARCHAR(20), ');
FDScript1.SQLScripts[0].SQL.Add('ADD FILESOURSE_PROC     VARCHAR(1000); ');
////////////////////////////////////////////////////////////////////////////////// ���������
FDScript1.SQLScripts[0].SQL.Add('CREATE SEQUENCE GEN_FOR_TABLE_TASK START WITH 0 INCREMENT BY 1;');
//////////////////////////////////////////////////////////////////////////////////
FDScript1.SQLScripts[0].SQL.Add('CREATE TABLE TABLE_TASK (');
FDScript1.SQLScripts[0].SQL.Add('ID_TABLE          INTEGER NOT NULL,');  //
FDScript1.SQLScripts[0].SQL.Add('NAME_TABLE        VARCHAR(500) NOT NULL,'); //
FDScript1.SQLScripts[0].SQL.Add('DESCRIPTION_TASK  VARCHAR(500),'); //
FDScript1.SQLScripts[0].SQL.Add('PC_RUN            VARCHAR(250),'); //
FDScript1.SQLScripts[0].SQL.Add('STATUS_TASK       VARCHAR(50),'); //
FDScript1.SQLScripts[0].SQL.Add('COUNT_PC          INTEGER,');//
FDScript1.SQLScripts[0].SQL.Add('CURRENT_TASK      VARCHAR(500),');
FDScript1.SQLScripts[0].SQL.Add('START_STOP        VARCHAR(20) DEFAULT 0,');
FDScript1.SQLScripts[0].SQL.Add('WORKS_THREAD      VARCHAR(20),');
FDScript1.SQLScripts[0].SQL.Add('USER_NAME         VARCHAR(256),');
FDScript1.SQLScripts[0].SQL.Add('PASS_USER         VARCHAR(256),');
FDScript1.SQLScripts[0].SQL.Add('SAVE_PASS         BOOLEAN DEFAULT false );');
FDScript1.SQLScripts[0].SQL.Add('ALTER TABLE TABLE_TASK ADD CONSTRAINT PK_TABLE_TASK PRIMARY KEY (ID_TABLE);');
///////////////////////////////////////////////////////////////�������
FDScript1.SQLScripts[0].SQL.Add('SET TERM ^ ;');
FDScript1.SQLScripts[0].SQL.Add('CREATE OR ALTER TRIGGER TABLE_TASK_BI0 FOR TABLE_TASK ');
FDScript1.SQLScripts[0].SQL.Add('ACTIVE BEFORE INSERT POSITION 0 ');
FDScript1.SQLScripts[0].SQL.Add('AS begin ');
FDScript1.SQLScripts[0].SQL.Add('if (new.id_table is null) then ');
FDScript1.SQLScripts[0].SQL.Add('new.id_table = gen_id(GEN_FOR_TABLE_TASK,1);');
FDScript1.SQLScripts[0].SQL.Add('end');
FDScript1.SQLScripts[0].SQL.Add('^');
FDScript1.SQLScripts[0].SQL.Add('SET TERM ; ^');

FDScript1.ValidateAll;

if FDScript1.ExecuteAll then
begin
ShowMessage('���� ������ ��������� �������!!!');
end
else
begin
 FDGUIxScriptDialog1.Options:=[ssCallstack,ssConsole];  /// �������� ���� ����� ��� ���������� �������� ����
 ShowMessage('��� ���������� ���� ��������� ������!');
end;
FDScript1.SQLScripts[0].SQL.Clear;
FDConnection1.Connected:=false;
UpdateMyDB.Free;
  except
  on E: Exception do
  begin
  frmDomainInfo.Memo1.Lines.Add('������ ���������� ���� ������. - '+e.Message);
  FDConnection1.Connected:=false;
  if Assigned(UpdateMyDB) then UpdateMyDB.Free;
  end;
  end;
end;



procedure TSettingsProgram.Button5Click(Sender: TObject);
var
FDQuery:TFDQuery;
qTransaction:TFDTransaction;
MicrList:Tstringlist;
DBOpenpatch:TOpenDialog;
i:integer;
begin
try
MicrList:=TStringList.Create;

dbopenPatch:=TOpenDialog.Create(self);
dbopenPatch.Name:='';
dbopenPatch.Title:='���� � ���������� Microsoft';
dbopenPatch.InitialDir:=extractfilepath(application.ExeName);

dbopenPatch.Filter:='|*.txt';
if DBOpenpatch.Execute then  MicrList.LoadFromFile(DBOpenpatch.FileName);
qTransaction:= TFDTransaction.Create(nil);
qTransaction.Connection:=DataM.ConnectionDB;
qTransaction.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
qTransaction.Options.AutoCommit:=false;
qTransaction.Options.AutoStart:=false;
qTransaction.Options.AutoStop:=false;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=qTransaction;
FDQuery.Connection:=DataM.ConnectionDB;
for I := 0 to MicrList.Count-1 do
begin
if  i<MicrList.Count-1 then
if (MicrList[i]<>'') and (MicrList[i+1]<>'') then
  begin
  qTransaction.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='update or insert into MICROSOFT_PRODUCT'+
  ' (PRODUCT_ID,DESCRIPTION)'+
  ' VALUES (:p1,:p2) MATCHING (PRODUCT_ID,DESCRIPTION)';
  FDQuery.params.ParamByName('p1').AsString:=''+(MicrList[i])+'';
  FDQuery.params.ParamByName('p2').AsString:=''+(MicrList[i+1])+'';
  FDQuery.ExecSQL;
  qTransaction.commit;
  end;
end;

finally
FDQuery.Free;
MicrList.Free;
dbopenPatch.Free;
end;
end;


//******************************************************************************************
procedure TSettingsProgram.HttpWork(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
var
  Http: TIdHTTP;
  ContentLength: Int64;
begin
  Http := TIdHTTP(ASender);
  try
  ContentLength := Http.Response.ContentLength;
  if (Pos('chunked', LowerCase(Http.Response.TransferEncoding)) = 0) and
     (ContentLength > 0) then
  begin
    ProgressB.Position := 100*AWorkCount div ContentLength;
  end;
  finally

  end;
end;

procedure TSettingsProgram.memo1URLClick(Sender: TObject; const URLText: string;
  Button: TMouseButton);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar(URLText), nil, nil, SW_SHOWNORMAL);
  except
  showmessage('������ ��� �������� ����������')
  end;
end;

procedure TSettingsProgram.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;
procedure TSettingsProgram.CreateFormElseErrorloadFile;
var
FormKey:Tform;
MemoKey:TJvRichEdit;
begin
FormKey:=TForm.Create(SettingsProgram);
FormKey.Caption:='���������� ����������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=430;
FormKey.Height:=230;
FormKey.OnClose:=FormKeyClose;
////////////
MemoKey:=TJvRichEdit.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Align:=alClient;
MemoKey.ScrollBars:=ssVertical;
MemoKey.ReadOnly:=true;
MemoKey.AutoURLDetect:=true;
MemoKey.OnURLClick:=memo1URLClick;
MemoKey.TabOrder:=0; // ��������� ����� � ���� �� ��������� ���������
MemoKey.Lines.Add('���� �� ������� ��� ��������� ������ �� ���������� ���������'+
' ����������� ����� ��� �� ������� ������ ��������. ��� �������� ������ ���������� ������'+
' https://skrblog.ru/download/kms/?wpdmdl=1478&masterkey=60322427c5fb6'+
' �������� ����� ������, ������ �� ������ skrblog.ru. ���������� ����� � ���������� ���������' +
' ���������, � ������� Vaccina, ���� �������� ��� �� �������� ���, ������� ��������� �������� � ������ � ����������.'+
' ������ ����� �������� �������� �������� ������� �������������,'+
' ��� ����� � ���������� ���������� ����� ����� �� �������:');
MemoKey.Lines.Add('https://github.com/kkkgo/KMS_VL_ALL');
MemoKey.Lines.Add('https://forums.mydigitallife.net/posts/838808/');
MemoKey.Lines.Add('https://pastebin.com/cpdmr6HZ');
MemoKey.Lines.Add('https://textuploader.com/1dav8');
FormKey.ShowModal;
end;



function TSettingsProgram.LoadFile(url,path,Fname:string):boolean;  // ������� �������� �����
var
MS: TMemoryStream;
  MyClass: TComponent;
begin
try
MS := TMemoryStream.Create;
try
IdHTTP1.Get(url, MS);
MS.SaveToFile(path+Fname);
MS.Clear;
result:=true;
finally
 MS.Free;
end;
except on E: Exception do begin result:=false; showmessage('������ ����������: '+e.Message) end;
end;
end;

function TSettingsProgram.ExistFileKMS(ListFile:TstringList):boolean;
var
i:integer;
begin
 for I := 0 to ListFile.Count-1 do
 if not FileExists(ListFile[i])  then
 begin
   result:=false;
   break;
 end
 else result:=true;
end;

function TSettingsProgram.loadFileRun:boolean; //������� �������� ����� �������� � ����������� ������ ��� KMS
var
LoadIni,SetIni:Tmeminifile;
i:integer;
path:string;
ListValue:TstringList;
NameFile:Tstringlist;
begin
try
if not frmDomainInfo.ping('yandex.ru') then
begin
i:=MessageDlg('��� ��������� ����������� � ���������!!!'+#10#13+' ���������� ���������� ��������?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
  begin
  result:=false;
  exit;
  end;
end;
ListValue:=TStringList.Create;
NameFile:=TStringList.Create;
ProgressB:=TProgressBar.Create(SettingsProgram);
ProgressB.Parent:=Groupbox5;
ProgressB.Left:=0;
ProgressB.Align:=alBottom;
ProgressB.Width:=Groupbox5.Width;
ProgressB.Max:=100;
ProgressB.Step:=1;
ProgressB.Position:=0;

try
IdHTTP1.OnWork:= HttpWork;
if Fileexists(extractfilepath(application.ExeName)+'\Vaccination.ini') then // ���� ���� ���� �� �������
DeleteFile(extractfilepath(application.ExeName)+'\Vaccination.ini');

if LoadFile('http://skrblog.ru/wp-content/uploads/download-manager-files/Vaccina/Vaccination.ini',extractfilepath(application.ExeName),'Vaccination.ini') then
Begin
 if Fileexists(extractfilepath(application.ExeName)+'\Vaccination.ini') then
  begin
  LoadIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Vaccination.ini',TEncoding.Unicode);
  try
  LoadIni.ReadSectionValues('FOLDER',ListValue); // ������ ����������� ���������
  for I := 0 to ListValue.Count-1 do createDir(extractfilepath(application.ExeName)+'\'+ListValue.ValueFromIndex[i]);  // ������� ����������� ���������
  ListValue.Clear;
  LoadIni.ReadSectionValues('URL',ListValue); // ������ ����������� URL ��� ��������
  LoadIni.ReadSectionValues('FILES',NameFile); // ������ ���� � ������������ ������
  for I := 0 to ListValue.Count-1 do
    begin
    frmDomainInfo.Memo1.Lines.Add('URL - '+ListValue.ValueFromIndex[i]+'  ������� - '+extractfilepath(application.ExeName)+' ���� - '+NameFile.ValueFromIndex[i]);
    if not LoadFile(ListValue.ValueFromIndex[i],extractfilepath(application.ExeName),NameFile.ValueFromIndex[i]) then // ���������� � ��������� �����  .
     begin
     frmDomainInfo.Memo1.Lines.Add('�� ������� ��������� ���� - '+NameFile.ValueFromIndex[i]);
     CreateFormElseErrorloadFile;
     CheckBox39.Checked:=false;
     break;
     end;
    end;
  ///// ������ ��������� � ���� ��������,   IniSettings
  SetIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  try
  SetIni.WriteString('Vaccina','Foder',extractfilepath(application.ExeName)+'Vaccina');
  SetIni.WriteString('Vaccina','Manual',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Manual',''));
  SetIni.WriteString('Vaccina','Uninstall',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Uninstall',''));
  SetIni.WriteString('Vaccina','Auto',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Auto',''));
  SetIni.WriteString('Vaccina','ManualKey',LoadIni.ReadString('KEY','ManualKey',''));
  SetIni.WriteString('Vaccina','UninstallKey',LoadIni.ReadString('KEY','UninstallKey',''));
  SetIni.WriteString('Vaccina','AutoKey',LoadIni.ReadString('KEY','AutoKey',''));
  SetIni.UpdateFile;
  finally
  SetIni.Free;
  end;
  /////////// ��������� ����������� ��� ��� �����
    ListValue.Clear;
    for I := 0 to NameFile.Count-1 do
    begin
    ListValue.Add(extractfilepath(application.ExeName)+NameFile.ValueFromIndex[i]);
    frmDomainInfo.Memo1.Lines.Add(extractfilepath(application.ExeName)+NameFile.ValueFromIndex[i]);
    end;

    if not ExistFileKMS(ListValue) then // ���� �� ��������� ��� ����������� �����
    Begin
    ShowMessage('�� ������� ��������� ����������� �����');
    CreateFormElseErrorloadFile;
    CheckBox39.Checked:=false;
    End
    else result:=true;
  finally
  LoadIni.Free;
  end;
  end;

End
else
  begin
  ShowMessage('�� ������� ��������� ���� ��������');
  CreateFormElseErrorloadFile;
  result:=false;
  CheckBox39.Checked:=false;
  end;


finally
ProgressB.Free;
ListValue.Free;
NameFile.Free;
if Fileexists(extractfilepath(application.ExeName)+'\Vaccination.ini') then // ���� ���� ���� �� �������
DeleteFile(extractfilepath(application.ExeName)+'\Vaccination.ini');
end;
except on E: Exception do
begin
  result:=false;
  ShowMessage('������ �������� ������ '+e.Message);
  CheckBox39.Checked:=false;
end;
end;
end;

procedure TSettingsProgram.CheckBox39Click(Sender: TObject);
begin
CheckBox40.Enabled:=CheckBox39.Checked;
CheckBox41.Enabled:=CheckBox39.Checked;
end;

procedure TSettingsProgram.CheckBox39MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
i:integer;
begin
if CheckBox39.Checked then
if not FileExists(IniSettings.ReadString('Vaccina','Manual',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd') ) then
begin
i:=MessageDlg('�� ������� ������� � ����� ��� ������.'+#10#13+' ��������� ����������� �����?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
if loadFileRun then
 ShowMessage('���������� ����� � ��������� ���������');
end;

end;

//********************************************************************************************

end.
