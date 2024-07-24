unit PropertyAccount;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
   Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Buttons, Vcl.StdCtrls,Vcl.ExtDlgs,ActiveX,
  ComObj,CommCtrl,
  JvExStdCtrls, JvCheckBox, Vcl.ExtCtrls,ActiveDs_TLB, Vcl.ComCtrls, JvBaseDlg,
  JvObjectPickerDialog, Vcl.ImgList;

type
  TFormUserAccount = class(TForm)
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    SIDLab: TLabel;
    CheckBox3: TCheckBox;
    ButtonRename: TSpeedButton;
    CheckBox1: TJvCheckBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    CheckBox2: TCheckBox;
    SpeedButton4: TSpeedButton;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox10: TCheckBox;
    CheckBox11: TCheckBox;
    CheckBox12: TCheckBox;
    CheckBox13: TCheckBox;
    CheckBox15: TCheckBox;
    ScrollBox1: TScrollBox;
    ScrollBox2: TScrollBox;
    Edit2: TEdit;
    Edit1: TEdit;
    GroupBox2: TGroupBox;
    PageControl1: TPageControl;
    TabSheetAccount: TTabSheet;
    TabSheetGroup: TTabSheet;
    Panel1: TPanel;
    ListViewGroup: TListView;
    SpeedButton3: TSpeedButton;
    SpeedButton5: TSpeedButton;
    JvObjectPickerDialogLocal: TJvObjectPickerDialog;
    JvObjectPickerDialogDomain: TJvObjectPickerDialog;
    ImageList1: TImageList;
    procedure FormShow(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CheckBox1Click(Sender: TObject);
    procedure ButtonRenameClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure CheckBox7Click(Sender: TObject);
    procedure CheckBox13Click(Sender: TObject);
    procedure CheckBox8Click(Sender: TObject);
    procedure CheckBox12Click(Sender: TObject);
    procedure CheckBox11Click(Sender: TObject);
    procedure CheckBox10Click(Sender: TObject);
    procedure CheckBox15Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);

  private
    function AccountPropertiesForRemotePC(NamePC,User,Pass,SID:string):TstringList;
    function ChengeAccount(NamePC,proper,SID,user,pass:string;ProperBol:boolean):string;
    function RenameAccount(NamePC,proper,SID,user,pass:string;ProperSTR:string):string;
    function GroupForLocalAccount(FWMIService: OLEVariant;NamePC,Domain,User:string):TstringList;
    procedure checkedAllChBox;
    procedure checkBoxTag0;
    function passwordchangeWinNT(namePC,CurUser,UserAdmin,AdminPasswd,NewPass:string):integer; /// ����� ������ ���������� ������������
    function passwordchangeDomain(DomainName,CurUser,UserAdmin,AdminPasswd,NewPass:string):bool;  /// ����� ������ ��������� ������������
    function RenameAccountDomain(DomainName,OldUser,NewName,UserAdmin,AdminPasswd:string):bool;
    function FullNameAccountDomain(DomainName,User,NewFullName,UserAdmin,AdminPasswd:string):bool;
    function FlagsAccountDomain(DomainName,User,UserAdmin,AdminPasswd:string;check:bool;operation:integer):bool;
    function AccountDomainPrivelegi(val:integer;user,domain,UserAdmin,AdminPasswd:string):bool; /// �������� �� ������������� ���������� ������������ ������� �� �����
    function checkPrivelegDomainUser(user,domain:string):bool;  /// �������� ��� ���������� ������������ ������
    function AccountDomainAddGroup(StrGroup:TstringList;user,domain,UserAdmin,AdminPasswd:string):bool;
    function AccountDomainDeleteGroup(StrGroup,user,domain,UserAdmin,AdminPasswd:string):bool;

    function AllLocalGroup(NamePC,User,Pass:string):bool;
    function readgroupForLocal(NamePC,Domain:string):TstringList;
    function readgroupForDomain(NamePC,Domain:string):TstringList;
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure FormGroupShow(Sender: TObject); /// ��������� ����������� ����������� �����
    procedure FormGroupClose(Sender: TObject; var Action: TCloseAction);
    procedure ButOkClick(Sender: TObject);
  public
    { Public declarations }

  end;
   function ADsOpenObject(lpszPathName,lpszUserName,lpszPassword : WideString;dwReserved : DWORD; const riid:TGUID; out ppObject): HResult; stdcall; external 'activeds.dll';
   function NetUserChangePassword(Domain: PWideChar; UserName: PWideChar; OldPassword: PWideChar;
   NewPassword: PWideChar): Longint; stdcall; external 'netapi32.dll';

var
  FormUserAccount: TFormUserAccount;
  CurrentAccountName,CurrentDomain:string;
  DomainOrLocalAccount:boolean;
  FormGr:Tform;
  ListGr:Tcombobox;
  ButOk:Tbutton;
const
wbemFlagForwardOnly = $00000020;
ADS_SECURE_AUTHENTICATION = $00000001;

implementation
{$R *.dfm}
uses umain;

  function GetObject(const name: string): IDispatch;
var
  Moniker: IMoniker;
  Eaten: integer;
  BindContext: IBindCtx;
  Dispatch: IDispatch;
begin
  OleCheck(CreateBindCtx(0, BindContext));
  OleCheck(MkParseDisplayName(BindContext, PWideChar(WideString(name)), Eaten, Moniker));
  OleCheck(Moniker.BindToObject(BindContext, nil, IDispatch, Dispatch));
  Result := Dispatch;
end;

procedure TFormUserAccount.FormGroupShow(Sender: TObject);
var
sid:string;
begin
AllLocalGroup(frmDomainInfo.ComboBox2.Text,frmDomainInfo.LabeledEdit1.Text,
frmDomainInfo.LabeledEdit2.Text); // ��������� ����� ���� ������� ��������� �����
if ListGr.Items.Count>0 then
begin
ListGr.ItemIndex:=0;
ListGr.AutoDropDown:=true;
end;

end;

procedure TFormUserAccount.FormGroupClose(Sender: TObject; var Action: TCloseAction);
begin
FormGr.Release; /// ����������� �����
end;

procedure TFormUserAccount.ButOkClick(Sender: TObject);
var
groupList:tstringList;
i:integer;
begin
groupList:=TStringList.Create;
groupList.Add(ListGr.Text);
if AccountDomainAddGroup(groupList,CurrentAccountName,CurrentDomain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text)
then
  begin
     for I := 0 to groupList.count-1 do
     with ListViewGroup.Items.Add do
     begin
     ImageIndex:=0;
     Caption:=(groupList[i]);
     end;
    // ShowMessage('�������� ���������');
   end;
 groupList.Free;
FormGr.Close;
end;

function TFormUserAccount.FormGroupCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormUserAccount);
FormGr.Name:='FormGroup';
FormGr.Caption:='������ ��������� ����� - '+namePC;
FormGr.Width:=323;
FormGr.Height:=97;
FormGr.BorderStyle:=bsDialog;
FormGr.OnShow:=FormGroupShow;
FormGr.OnClose:=FormGroupClose;
FormGr.Position:=poMainFormCenter;
ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:='ListGr';
ListGr.Text:='';
ListGr.Left:=8;
ListGr.Top:=8;
ListGr.Width:=298;
ListGr.DropDownCount:=15;
ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='ButOk';
ButOk.Caption:='��';
ButOk.Top:=35;
ButOk.Left:=231;
ButOk.OnClick:=ButOkClick;
FormGr.Show;
result:=true;
except
  on E:Exception do
    begin
    ShowMessage(e.Message);
    result:=false;
    end;
  end;
end;

/////////////////////////////////////////////
function TFormUserAccount.AllLocalGroup(NamePC,User,Pass:string):bool;
var            ///// ��������� ������
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  nameGr:string;
begin
try
OleInitialize(nil);
ListGr.Clear;
frmDomainInfo.memo1.Lines.Add(NamePC+' - �������� ������ �����, ��������... ');
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,Domain FROM Win32_Group WHERE Domain="'+NamePC+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������������
 begin
  nameGr:='';
  if (FWbemObject.name)<>null then nameGr:=string(FWbemObject.name);
  ListGr.Items.Add(nameGr);
 FWbemObject:=Unassigned;
 end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
result:=true;
except
  on E:Exception do
    begin
    result:=false;
    OleUnInitialize;
    end;
  end;
end;



/////////////////////////////////////////////////////
function TFormUserAccount.readgroupForLocal(NamePC,Domain:string):TstringList;
var          ////// ������ ����� ������������� ��� ���������� ������������
i:integer;
begin
try      ///https://docs.microsoft.com/en-us/windows/win32/api/objsel/ns-objsel-dsop_filter_flags
         ///https://docs.microsoft.com/ru-ru/windows/win32/api/objsel/ns-objsel-dsop_scope_init_info
OleInitialize(nil);
result:=TStringList.Create;
//JvObjectPickerDialogLocal.TargetComputer:=NamePC;
//JvObjectPickerDialog1.Scopes[0].DcName:= Domain;//frmDomainInfo.GetDomainController(CurrentDomainName);
//JvObjectPickerDialog1.Scopes[1].DcName:=NamePC;
if JvObjectPickerDialogLocal.Execute() then
begin
if JvObjectPickerDialogLocal.Selection.Count<>0 then
for I := 0 to JvObjectPickerDialogLocal.Selection.Count-1 do
  begin
  result.add(JvObjectPickerDialogLocal.Selection.Items[i].Name);
  end;
end;
except
    on E:Exception do
    begin
    showmessage('������ : '+E.Message);
    end;
  end;
OleUnInitialize;
end;

function TFormUserAccount.readgroupForDomain(NamePC,Domain:string):TstringList;
var          ////// ������ ����� ������������� ��� ���������� ������������
i:integer;
begin
try      ///https://docs.microsoft.com/en-us/windows/win32/api/objsel/ns-objsel-dsop_filter_flags
         ///https://docs.microsoft.com/ru-ru/windows/win32/api/objsel/ns-objsel-dsop_scope_init_info
OleInitialize(nil);
result:=TStringList.Create;
//JvObjectPickerDialog1.TargetComputer:=NamePC;
JvObjectPickerDialogDomain.Scopes[0].DcName:= frmDomainInfo.GetDomainController(CurrentDomainName); /// ��������� ���������� ������
//JvObjectPickerDialog1.Scopes[1].DcName:=NamePC;
if JvObjectPickerDialogDomain.Execute() then
begin
if JvObjectPickerDialogDomain.Selection.Count<>0 then
for I := 0 to JvObjectPickerDialogDomain.Selection.Count-1 do
  begin
  result.add(JvObjectPickerDialogDomain.Selection.Items[i].Name);
  end;
end;
except
    on E:Exception do
    begin
    showmessage('������ : '+E.Message);
    end;
  end;
OleUnInitialize;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////
function TFormUserAccount.AccountDomainDeleteGroup(StrGroup,user,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// �������� ������������ �� ������
  InterF :  IADsUser;
  GroupUs: IADsGroup;

begin
  try
    OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
    OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
    GroupUs.Remove(InterF.ADsPath);
    GroupUs.SetInfo;
    GroupUs._Release;//???
    InterF._Release; ///???
    result:=true;
 ///////////////////////////////////////////////////////////////////////
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage(E.message);
    end;
end;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////
function TFormUserAccount.AccountDomainAddGroup(StrGroup:TstringList;user,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// ���������� ������������ � ������
  z:integer;
  InterF :  IADsUser;
  GroupUs: IADsGroup;

begin
  try
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
  for z := 0 to StrGroup.Count-1 do
  begin
    OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup[z]+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
    GroupUs.Add(InterF.ADsPath);
    GroupUs.SetInfo;
  end;
    GroupUs._Release;//
    InterF._Release; ///
    result:=true;
 ///////////////////////////////////////////////////////////////////////
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage(E.message);
    end;
end;
end;


function TFormUserAccount.AccountDomainPrivelegi(val:integer;user,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// ������ ������ ��� ��������� ������������
  UsrFlags: OleVariant;
  InterF :  IADsUser;
begin
  try
    OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
   UsrFlags:=InterF.Get('UserFlags');
   if VarIsNumeric(UsrFlags) then
   begin
     if (integer(UsrFlags)or val)=integer(UsrFlags) then result:=true
     else result:=false;
   end
   else result:=false;
   InterF._Release;
  // frmDomainInfo.Memo1.Lines.add('����� - '+ string(InterF.Get('UserFlags')));
 //////////////////////////////////////////////////////// ������� ������� , ��� �������� ������ � ������ ������, �������� � ��������� ������������
   {PUsr := GetObject('WinNT://'+domain+'/'+user+',user');
    if ((integer(Pusr.UserFlags)) or val)=integer(Pusr.UserFlags) then
    result:=true
    else result:=false; }
 ///////////////////////////////////////////////////////////////////////
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage(E.message);
    end;
end;
//VarClear(Pusr);
VarClear(UsrfLAGS);
end;

function TFormUserAccount.checkPrivelegDomainUser(user,domain:string):bool;
begin
//checkbox5.checked:=AccountDomainPrivelegi(8388608,user,domain);  //��������� ����� ������ ��� ��������� �����
checkbox6.checked:=AccountDomainPrivelegi(64,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); /// ��������� ����� ������ �������������
checkbox7.checked:=AccountDomainPrivelegi(65536,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //���� �������� ������  �� ���������
checkbox13.checked:=AccountDomainPrivelegi(2,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); /// �������� ��� ��������� ������� ������
checkbox8.checked:=AccountDomainPrivelegi(128,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); // ������� ������, ��������� ��������� ����������
checkbox12.checked:=AccountDomainPrivelegi(262144,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); //��� �������������� ����� � ���� ����� �����-�����
checkbox11.checked:=AccountDomainPrivelegi(1048576,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); // ������� ������ ����� � �� ����� ���� �������������
checkbox10.checked:=AccountDomainPrivelegi(2097152,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //������������ ���� ���������� Kerberos DES ��� ����
checkbox15.checked:=AccountDomainPrivelegi(4194304,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //��� ��������������� �������� ����������� Kerberos
end;

procedure TFormUserAccount.ButtonRenameClick(Sender: TObject);
var
s,NewFullName:string;
begin
NewFullName:='';
if not InputQuery('������� ����� ������ ��� ������������', '��� ������������:', NewFullName)
 then exit;

if DomainOrLocalAccount then
  begin
  s:=RenameAccount(frmDomainInfo.ComboBox2.Text,'FullName',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,NewFullName);
  if s='' then  showmessage('��������� ������')
  else
  begin
  frmDomainInfo.Memo1.Lines.Add('��������: "������ ������� �����" �����������: '+s);
  LabeledEdit1.Text:=NewFullName;
  end;
  end
else
  begin
  if FullNameAccountDomain(CurrentDomain,CurrentAccountName,
  NewFullName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
  begin
   showmessage('��������� ����� ������ �������');
   LabeledEdit1.Text:=NewFullName;
  end;
  end;
end;

procedure TFormUserAccount.SpeedButton1Click(Sender: TObject);
var
s,NewName:string;
begin
NewName:='';
if not InputQuery('������� ����� ��� ������������', '��� ������������:', NewName)
 then exit;
if NewName='' then
begin
  ShowMessage('��� �� ����� ���� ������');
  exit;
end;

if DomainOrLocalAccount then
  begin
  s:=RenameAccount(frmDomainInfo.ComboBox2.Text,'Rename',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,NewName);
  if s='' then  showmessage('��������� ������')
   else
   begin
   frmDomainInfo.Memo1.Lines.Add('��������: "��������������" �����������: '+s);
   CurrentAccountName:=NewName; /// ���� ����������� � ��� ��, �� ����������� ���������� ����� ���
   LabeledEdit2.Text:=NewName;
   end;
  end
else
  begin
  if RenameAccountDomain(CurrentDomain,CurrentAccountName,NewName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
   begin
   showmessage('������������ ������������');
   CurrentAccountName:=LabeledEdit2.Text; /// ���� ����������� � ��� ��, �� ����������� ���������� ����� ���
   LabeledEdit2.Text:=NewName;
   end;
  end;
end;


function TFormUserAccount.FlagsAccountDomain(DomainName,User,UserAdmin,AdminPasswd:string;check:bool;operation:integer):bool;
var              ////////////////// ��������� ������ ��� ��������� ������������
  PUsr,UsFlags: OleVariant;
  InterF :  IADsUser;
begin
  try
  OleCheck(ADsOpenObject('WinNT://'+DomainName+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
  UsFlags:=InterF.Get('UserFlags');
  if check then
   begin
   InterF.Put('UserFlags',(UsFlags xor operation));
   end
  else
   begin
   InterF.Put('UserFlags',(UsFlags - operation));
   end;
   InterF.SetInfo;

  { PUsr := GetObject('WinNT://'+DomainName+'/'+User+',user');
    if check then
      begin
      Pusr.UserFlags:=(integer(Pusr.UserFlags) xor operation);
      end
    else
      begin
      Pusr.UserFlags:=integer(Pusr.UserFlags)-operation;
      end;
    Pusr.setinfo;}
   result:=true;
   InterF._Release;
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage(E.message);
    end;
end;
//VarClear(Pusr);
VarClear(UsFlags);
end;
////////////////////////////////////////////////////////////////////////////

function TFormUserAccount.passwordchangeWinNT(namePC,CurUser,UserAdmin,AdminPasswd,NewPass:string):integer;
var              ////////////////// ����� ������ ���������� ������������
  Pusr: OleVariant;
  InterF :  IADsUser;
begin
  try
  //////////////////////////////////////////////////////////////// �������� � ������� ��� ����� ��� AD,
 { result:=NetUserChangePassword(PWideChar(WideString(NamePC)),
     PWideChar(WideString(CurUser)),
     PWideChar(WideString(OldPass)),
     PWideChar(WideString(NewPass)));}
 ///////////////////////////////////////////////////////////////////////
  ///////// ����� ������  � ����������� ������ ������ � ������������/////////////////////////////////////
    OleCheck(ADsOpenObject('WinNT://'+NamePC+'/'+CurUser+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
    //ShowMessage(InterF.Description);
    InterF.SetPassword(NewPass);
    /////////////////////////////////////////////////////////////////////////
   ///////////////////////////////////////////  ���� ������ � ����������� ��� ����� ������������ � ������, � ��������� ������������� �������� ������������
    //PUsr := GetObject('WinNT://'+NamePC+'/'+CurUser+',user');
    ///PUsr.SetPassword(NewPass);
   ///////////////////////////////////////////////////////
     InterF._Release;
  except
    on E: EOleException do
    begin
    ShowMessage(E.message);
    end;
end;
//VarClear(Pusr);
end;

function TFormUserAccount.FullNameAccountDomain(DomainName,User,NewFullName,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// ����� ������� ����� ��������� ������������
  PUsr,Nusr: OleVariant;
  InterF :  IADsUser;
begin
  try
   OleCheck(ADsOpenObject('WinNT://'+DomainName+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
   InterF.FullName:=NewFullName;
   InterF.SetInfo;


   {PUsr := GetObject('WinNT://'+DomainName+'/'+User+',user');
   Pusr.FullName:=NewFullName;
   Pusr.setinfo; }
   result:=true;
   InterF._Release;
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage(E.message);
    end;
end;
//VarClear(Pusr);
end;

function TFormUserAccount.RenameAccountDomain(DomainName,OldUser,NewName,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// ����� ����� ��������� ������������
  PUsr,Nusr: OleVariant;
  InterF :  IADsContainer;//IADsUser;
begin
  try
     OleCheck(ADsOpenObject('WinNT://'+DomainName,UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsContainer, InterF));
     InterF.MoveHere('WinNT://'+DomainName+'/'+OldUser,NewName);
     result:=true;
   // PUsr := GetObject('WinNT://'+DomainName);
   // PUsr.MoveHere('WinNT://'+DomainName+'/'+OldUser,NewName);
     InterF._Release;
   except
    on E: EOleException do
    begin
      Result:=false;
      ShowMessage(E.message);
    end;
end;
//VarClear(Pusr);
end;

function TFormUserAccount.passwordchangeDomain(DomainName,CurUser,UserAdmin,AdminPasswd,NewPass:string):bool;
var              ////////////////// ����� ������ ��������� ������������
  Usr: OleVariant;
  InterF :  IADsUser;
begin
  try
     OleCheck(ADsOpenObject('WinNT://'+DomainName+'/'+CurUser+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
    //ShowMessage(InterF.Description);
    InterF.SetPassword(NewPass);
    result:=true;
    //Usr := GetObject('WinNT://'+DomainName+'/'+CurUser+',user');
    //ShowMessage(string(usr.fullName));
   // Usr.SetPassword(NewPass);
   // result:=true;
    InterF._Release;
  except
    on E: EOleException do
    begin
    result:=false;
    ShowMessage(E.message);
    end;
end;
VarClear(usr);
end;

procedure TFormUserAccount.SpeedButton2Click(Sender: TObject);
var
Oldpass,newpass,newpass1:string;

i:integer;
begin
newpass:='';
newpass1:='';
Oldpass:='';

if DomainOrLocalAccount then // ���� ������������ ���������
  begin
   i:=MessageBox(Self.Handle, PChar('�������� ������ ������������?')
        , PChar('������������ - '+CurrentAccountName ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
   if not InputQuery('������� ����� ������ ������������', '������:', newpass) then exit
   else if not InputQuery('��������� ���� ������', '������:', newpass1) then exit;
   if newpass<>newpass1 then
   begin
   ShowMessage('������ �� ���������!'+#10#13+'��������� ��� ���.');
   exit;
   end;
  i:=passwordchangeWinNT(frmDomainInfo.ComboBox2.Text,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,
  frmDomainInfo.LabeledEdit2.Text,newpass); ShowMessage(SysErrorMessage(i));
  end
else    //// ����� �������� ������
  begin
  i:=MessageBox(Self.Handle, PChar('�������� ������ ������������?')
      , PChar('������������ - '+CurrentAccountName ) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then exit;
 if not InputQuery('������� ����� ������ ������������', '������:', newpass) then exit
 else if not InputQuery('��������� ���� ������', '������:', newpass1) then exit;
 if newpass<>newpass1 then
 begin
 ShowMessage('������ �� ���������!'+#10#13+'��������� ��� ���.');
 exit;
 end;
  if passwordchangeDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,
frmDomainInfo.LabeledEdit2.Text,newpass) then ShowMessage('������ ����������');
  end;
end;

procedure TFormUserAccount.SpeedButton3Click(Sender: TObject);
var
groupList:TstringList;
i:integer;
LocalOrDomain:string;
begin

LocalOrDomain:='';
groupList:=TStringList.Create;
if DomainOrLocalAccount then FormGroupCreate(frmDomainInfo.ComboBox2.Text,'','') ///groupList:=readgroupForLocal(frmDomainInfo.ComboBox2.Text,CurrentDomain)
else groupList:=readgroupForDomain(frmDomainInfo.ComboBox2.Text,CurrentDomain);

if groupList.Count<>0 then /// ���� ������ �� ����, �.�. ���� �������� ���� ���� ������, ����� �� ���������� �� ����������
begin
  if AccountDomainAddGroup(groupList,CurrentAccountName,CurrentDomain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text)
  then
    begin
     for I := 0 to groupList.count-1 do
     with ListViewGroup.Items.Add do
     begin
     ImageIndex:=0;
     Caption:=(groupList[i]);
     end;
    end;
  end;


groupList.Free;
end;

procedure TFormUserAccount.SpeedButton4Click(Sender: TObject);
begin
Memo1.Lines.Add(inttostr(strtoint(Edit1.Text)or strtoint(Edit2.Text)))
end;



procedure TFormUserAccount.SpeedButton5Click(Sender: TObject);
var
i:integer;
begin

if (ListViewGroup.Items.Count>0)and (ListViewGroup.SelCount=1) then
begin
i:=MessageBox(Self.Handle, PChar('������� ������?')
      , PChar(ListViewGroup.ItemFocused.Caption) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then exit;

if AccountDomainDeleteGroup(ListViewGroup.ItemFocused.Caption,CurrentAccountName,CurrentDomain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text)
 then ListViewGroup.ItemFocused.Delete;
end;
end;

procedure TFormUserAccount.CheckBox1Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin
if CheckBox1.Tag=1 then exit; // ���� ��������� ��� ���������� �� �������
if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;

if CheckBox1.Checked then v:='���������'
else v:='��������';
i:=MessageBox(Self.Handle, PChar(v+' ������� ������?')
      , PChar(CurrentAccountName) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then
 begin
  CheckBox1.Tag:=1;
  CheckBox1.Checked:=not CheckBox1.Checked; /// ���������� ��������
  CheckBox1.Tag:=0;
  exit;
 end;

 if i=IDYes then
 begin
   if DomainOrLocalAccount then //// ���� ��������� ������
  begin
  s:=ChengeAccount(RemotePC,'Disabled',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox1.Checked);
  if s='' then ShowMessage('������ ��������')else
  frmDomainInfo.Memo1.Lines.Add(v+' ������� ������: '+s);
  end;
{else  /// ���� �������
  begin
  if FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox1.Checked,1) then ShowMessage('�������� "'+v+' ������� ������" ���������');
  end;}
 end;

end;



procedure TFormUserAccount.CheckBox2Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin

if CheckBox2.Tag=1 then exit; // ���� ��������� ��� ���������� �� �������

 if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;
if DomainOrLocalAccount then
  begin
  s:=ChengeAccount(RemotePC,'PasswordExpires',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text, not CheckBox2.Checked);
  if s='' then ShowMessage('������ ��������')else
  frmDomainInfo.Memo1.Lines.Add('���� �������� ������ �� ���������: '+s);
  end;
{else
  begin
  if FlagsAccountDomain(CurrentDomain,CurrentAccountName,not CheckBox2.Checked,2) then ShowMessage('�������� ������� ���������');
  end;}
end;

procedure TFormUserAccount.CheckBox3Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin
if CheckBox3.Tag=1 then exit; // ���� ��������� ��� ���������� �� �������
if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text
else
begin
ShowMessage('������ ��� ��������� �������������');
exit;
end;
 s:=ChengeAccount(RemotePC,'PasswordChangeable',SIDLab.Caption,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text, not CheckBox3.Checked);
if s='' then ShowMessage('������ ��������')else
frmDomainInfo.Memo1.Lines.Add('��������� ����� ������ �������������: '+s)
end;


procedure TFormUserAccount.CheckBox4Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin
{CheckBox3.Enabled:=not CheckBox4.Checked;
CheckBox2.Enabled:=not CheckBox4.Checked;
if CheckBox3.Checked then CheckBox3.Checked:=false;



if CheckBox4.Tag=1 then exit; // ���� ��������� ��� ���������� �� �������

CheckBox2.Tag:=1;
CheckBox2.Checked:=not CheckBox4.Checked;
CheckBox2.Tag:=0;

if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;
if DomainOrLocalAccount then
begin
s:=ChengeAccount(RemotePC,'PasswordExpires',SIDLab.Caption,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox4.Checked);
frmDomainInfo.Memo1.Lines.Add('��������� ����� ������ ��� ��������� ����� � �������: '+s);
end; }

{else
begin
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox4.Checked,2) then ShowMessage('�������� ������� ���������');
end; }
end;

procedure TFormUserAccount.CheckBox5Click(Sender: TObject);
begin   // ��������� ����� ������ ��� ��������� �����
{if CheckBox5.Tag=1 then  exit;

if CheckBox6.Checked then CheckBox6.Checked:=false; /// ������� ������ �� ����� ������
if CheckBox7.Checked then CheckBox7.Checked:=false;  /// ������� ���� �������� ������ �� ���������

FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox5.Checked,8388608);}
end;

procedure TFormUserAccount.CheckBox6Click(Sender: TObject);
var
s:string;
begin  //��������� ����� ������ �������������
if CheckBox6.Tag=1 then exit;
if checkbox6.Checked then s:='���������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox6.Checked,64) then
 frmDomainInfo.Memo1.Lines.add(s+' ����� ������ �������������: OK');
end;

procedure TFormUserAccount.CheckBox7Click(Sender: TObject);
var
s:string;
begin  //���� �������� ������  �� ���������
if CheckBox7.Tag=1 then exit;
if checkbox7.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox7.Checked,65536) then
 frmDomainInfo.Memo1.Lines.add(s+': C��� �������� ������  �� ���������: OK');
end;

procedure TFormUserAccount.CheckBox8Click(Sender: TObject);
var
s:string;
begin   ///������� ������, ��������� ��������� ����������
if CheckBox8.Tag=1 then exit;
if checkbox8.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox8.Checked,128) then
 frmDomainInfo.Memo1.Lines.add(s+': ������� ������, ��������� ��������� ����������: OK');
end;

procedure TFormUserAccount.CheckBox10Click(Sender: TObject);
var
s:string;
begin  ///������������ ���� ���������� Kerberos DES ��� ���� ������� ������
if CheckBox10.Tag=1 then  exit;
if checkbox10.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox10.Checked,2097152) then
 frmDomainInfo.Memo1.Lines.add(s+': ������������ ���� ���������� Kerberos DES ��� ���� ������� ������: OK');
end;

procedure TFormUserAccount.CheckBox11Click(Sender: TObject);
var
s:string;
begin    // ������� ������ ����� � �� ����� ���� �������������
if CheckBox11.Tag=1 then exit;
if checkbox11.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox11.Checked,1048576) then
 frmDomainInfo.Memo1.Lines.add(s+': ������� ������ ����� � �� ����� ���� �������������: OK');
end;

procedure TFormUserAccount.CheckBox12Click(Sender: TObject);
var
s:string;
begin  //��� �������������� ����� � ���� ����� �����-�����
if CheckBox12.Tag=1 then exit;
if checkbox12.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox12.Checked,262144) then
 frmDomainInfo.Memo1.Lines.add(s+': ��� �������������� ����� � ���� ����� �����-�����: OK');
end;

procedure TFormUserAccount.CheckBox13Click(Sender: TObject);
var
s:string;
begin   /// ��������� �������� ������� ������
if CheckBox13.Tag=1 then  exit;
if not checkbox13.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox13.Checked,2) then
frmDomainInfo.Memo1.Lines.add(s+' ������� ������: OK');
end;

procedure TFormUserAccount.CheckBox15Click(Sender: TObject);
var
s:string;
begin //��� ��������������� �������� ����������� Kerberos
if CheckBox15.Tag=1 then exit;
if checkbox15.Checked then s:='��������' else s:='���������';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox15.Checked,4194304) then
frmDomainInfo.Memo1.Lines.add(s+' ��������������� �������� ����������� Kerberos: OK');
end;


function TFormUserAccount.GroupForLocalAccount(FWMIService: OLEVariant;NamePC,Domain,User:string):TstringList;
var
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  groupSTRAll,groupSTR:string;
begin
try
OleInitialize(nil);
result:=tstringlist.Create;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_GroupUser','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������������
 begin
 groupSTR:='';
 if (FWbemObject.GroupComponent<>null)and(FWbemObject.PartComponent<>null) then
  begin
   if pos(Ansiuppercase('Domain="'+Domain+'",Name="'+User+'"'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0 then
     begin
     groupSTR:='';
     groupSTR:=(string(FWbemObject.GroupComponent));
     groupSTR:=StringReplace(groupSTR,'"','',[rfReplaceAll]);
     result.add(copy(groupSTR,pos(',Name=',groupSTR)+6,length(groupSTR)-1));
     end;
  end;
 FWbemObject:=Unassigned;
 end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    result.add('������ ������ ����� ������������ "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

function TFormUserAccount.RenameAccount(NamePC,proper,SID,user,pass:string;ProperSTR:string):string;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  res:integer;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// ������������
  begin
  if proper='Rename' then
    begin
    res:=FWbemObject.rename(ProperSTR);
    result:=SysErrorMessage(res);
    if res=0 then CurrentAccountName:=ProperSTR; /// ���� ��������� �� �� ����������� ���������� ����� ���
    end;
  if proper='FullName' then
    begin
    FWbemObject.FullName:=ProperSTR;
    result:=FWbemObject.put_();
    end;
  FWbemObject:=Unassigned;
  end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    //result.Add('��� - '+step+' / ������: '+E.Message);
    frmDomainInfo.memo1.Lines.Add(NamePC+' - ������ ��������� ������� ������ "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;




function TFormUserAccount.ChengeAccount(NamePC,proper,SID,user,pass:string;ProperBol:boolean):string;
var
FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
//FWMIService.Security_.Privileges.AddAsString('SeCreateTokenPrivilege',true); // ��������� ��� �������� ��������� ������� ������.
//FWMIService.Security_.Privileges.AddAsString('SeAssignPrimaryTokenPrivilege',true);  // ��������� ��� ������ ������ ������ ��������.
//FWMIService.Security_.Privileges.AddAsString('SeLockMemoryPrivilege',true);         //��������� ��� ���������� ������� � ������.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseQuotaPrivilege',true);     //��������� ��������� ����� ������ ��� ��������.
//FWMIService.Security_.Privileges.AddAsString('SeMachineAccountPrivilege',true);   // ��������� ��� ���������� ������� ������� � �����.
//FWMIService.Security_.Privileges.AddAsString('SeTcbPrivilege',true);             //��������� ����������� ��� ����� ������������ �������. ��������� �������� ������ �������� ������������ ����.
//FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);        //��������� ��� ���������� ������� � �������� ������������ NT.
//FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);  //��������� ������� ����� ������������� �� ����� ��� ������ �������, �� ���� ������ �������� ������� (ACE) � ������ �������� ������� �� ���������� (DACL).
//FWMIService.Security_.Privileges.AddAsString('SeLoadDriverPrivilege',true);     //��������� ��� �������� ��� �������� �������� ����������.
//FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);     //��������� ��� ����� ���������� ���������� � ������������������ �������.
//FWMIService.Security_.Privileges.AddAsString('SeSystemtimePrivilege',true);      // ��������� �������� ��������� �����.
//FWMIService.Security_.Privileges.AddAsString('SeProfileSingleProcessPrivilege',true); //��������� ��� ����� ���������� ������� ��� ������ ��������.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseBasePriorityPrivilege',true);  // ��������� ��� ���������� ���������� ������������
//FWMIService.Security_.Privileges.AddAsString('SeCreatePagefilePrivilege',true);       // ��������� ��� �������� ����� ��������.
//FWMIService.Security_.Privileges.AddAsString('SeCreatePermanentPrivilege',true);      //��������� ��� �������� ���������� ����� ��������.
//FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege',true);               //��������� ��� ���������� ����������� ������ � ���������, ���������� �� ������ ACL, ���������� ��� �����.
//FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',true);            //��������� ��� �������������� ������ � ��������� ���������� �� ������ ACL, ���������� ��� �����.
//FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   ��������� ��������� ��������� �������
//FWMIService.Security_.Privileges.AddAsString('SeDebugPrivilege',true);             //��������� ��� ������� � ��������� ������ ��������, �������������� ������ ������� ������.
//FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);             //��������� ��� �������� ������� ������ � ������� ������������ NT. ������ ���������� ������� ������ ����� ��� ����������.
//FWMIService.Security_.Privileges.AddAsString('SeSystemEnvironmentPrivilege',true); // ��������� ��� ��������� ����������������� ����������� ������ ������, ������� ���������� ���� ��� ������ ��� �������� ������ ������������.
//FWMIService.Security_.Privileges.AddAsString('SeChangeNotifyPrivilege',true);     //  ��������� ��� ��������� ����������� �� ���������� ������ ��� ��������� � ������ �������� �������. ��� ���������� �������� �� ��������� ��� ���� �������������.
//FWMIService.Security_.Privileges.AddAsString('SeRemoteShutdownPrivilege',true);   //��������� ��������� ��������� ���������.
//FWMIService.Security_.Privileges.AddAsString(' SeUndockPrivilege',true);           //��������� ����� ������� � ���-�������.
//FWMIService.Security_.Privileges.AddAsString('SeSyncAgentPrivilege',true);         // ��������� ��� ������������� ������ ������ ���������.
//FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);   // ���������, ����� ������� ������ ����������� � ������������� ���� ����������� ��� �������������.
//FWMIService.Security_.Privileges.AddAsString('SeManageVolumePrivilege',true);       //��������� ��� ���������� ����� ��������� ������������.
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// ������������
  begin
  if proper='Disabled' then
    begin
    FWbemObject.disabled:=ProperBol;
    result:=FWbemObject.put_();
    end;
  if proper='PasswordChangeable' then
    begin
    FWbemObject.PasswordChangeable:=ProperBol;
    result:=FWbemObject.put_();
    end;
  if proper='PasswordExpires' then
    begin
    FWbemObject.PasswordExpires:=ProperBol;
    result:=FWbemObject.put_();
    end;
  FWbemObject:=Unassigned;
  end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    //result.Add('��� - '+step+' / ������: '+E.Message);
    frmDomainInfo.memo1.Lines.Add(NamePC+' - ������ ��������� ������� ������ "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;

end;





procedure TFormUserAccount.FormClose(Sender: TObject; var Action: TCloseAction);
begin
checkBoxTag0;
end;

procedure TFormUserAccount.checkBoxTag0;
begin
CheckBox1.Tag:=0;
CheckBox2.Tag:=0;
CheckBox3.Tag:=0;

CheckBox6.Tag:=0;
CheckBox7.Tag:=0;
CheckBox8.Tag:=0;
CheckBox10.Tag:=0;
CheckBox11.Tag:=0;
CheckBox12.Tag:=0;
CheckBox13.Tag:=0;
CheckBox15.Tag:=0;
end;


procedure TFormUserAccount.checkedAllChBox;
begin
CheckBox1.Tag:=1; /// ��� � ����� ��� ����������� ��������� ����� �� ����������� ���������
CheckBox2.Tag:=1;
CheckBox3.Tag:=1; ///

CheckBox6.Tag:=1;
CheckBox7.Tag:=1;
CheckBox8.Tag:=1;
CheckBox10.Tag:=1;
CheckBox11.Tag:=1;
CheckBox12.Tag:=1;
CheckBox13.Tag:=1;
CheckBox15.Tag:=1;


CheckBox1.Checked:=false;
CheckBox2.Checked:=false;
CheckBox3.Checked:=false;
LabeledEdit1.Text:='';
LabeledEdit2.Text:='';

end;

procedure TFormUserAccount.FormShow(Sender: TObject);
var
sid:string;
begin
PageControl1.TabIndex:=0; /// ��� ��������� ������� ������� ������������
sid:=SidLab.caption;
Memo1.clear;
CurrentAccountName:='';
CurrentDomain:='';
checkedAllChBox;

DomainOrLocalAccount:=true; /// ����� � ��������� ������
Memo1.Lines:=AccountPropertiesForRemotePC(frmDomainInfo.ComboBox2.Text,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,SID);
/// ���� ������ � ���������� ����� �� �������� �� ������� �� ��������� ���� ���������� � ������
if (Memo1.Text='') and (frmDomainInfo.LabelEdEdit3.Text<>'') then
 Memo1.Lines:=AccountPropertiesForRemotePC('127.0.0.1',
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,SID);
end;

function TFormUserAccount.AccountPropertiesForRemotePC(NamePC,User,Pass,SID:string):TstringList;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  utc:integer;
  GroupForUs:TstringList;
  sourceSTR:string;
  i:integer;
  step:string;
  const
  SIDType: array [1..9] of string = ('������������','������','�����','��������� (Alias)','TypeWellKnownGroup','SID ��������� ������� ������','�� �������������� SID','SID �� ���������� ����','SID ����������');
  function TypeAccount(id:integer):string;
  begin
  case id of
  256:result:='��������� �������� ������� ������';
  512:result:='������� ������� ������';
  2048:result:='����������� ������� ������';
  4096:result:='������� ������ ������� �������';
  8192:result:='������������� ������� ������ �������';
  else result:='Unknown';
  end;
  end;
  function boltostr(b:Olevariant):string;
  begin
  try
    if b<>null then
      if b then result:='��'
      else result:='���';
   except  result:='Unknown'
  end;
  end;

begin
try
frmDomainInfo.Memo1.Lines.Add('�������� ������ � ������������, ��������... ');
OleInitialize(nil);
CurrentDomain:='';
result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// ������������
begin

if VarIsNumeric(FWbemObject.AccountType) then  result.Add('��� ������� ������: '+TypeAccount(FWbemObject.AccountType));

if (FWbemObject.Caption<>null) then
begin
 result.Add ('������� ������: '+string(FWbemObject.Caption));
end;

if (FWbemObject.Domain<>null) then
begin
CurrentDomain:=string(FWbemObject.Domain);
CurrentDomain:=AnsiLowerCase(Trim(CurrentDomain)); /// �������� � �������� ������, ������� �������
result.Add('�����: '+CurrentDomain);
end;

if (FWbemObject.Name<>null) then
begin
result.Add('���: '+string(FWbemObject.Name));
CurrentAccountName:=string(FWbemObject.Name);
LabeledEdit2.Text:=string(FWbemObject.Name);
end;

if (FWbemObject.FullName<>null) then
begin
 LabeledEdit1.Text:=string(FWbemObject.FullName);
end;

if (FWbemObject.Description<>null) then
begin
 result.Add ('��������: '+string(FWbemObject.Description));
end;

if (FWbemObject.LocalAccount<>null) then
begin
  result.Add('��������� ������� ������: '+boltostr(FWbemObject.LocalAccount));
  DomainOrLocalAccount:=FWbemObject.LocalAccount;
  ScrollBox1.Visible:=not DomainOrLocalAccount; /// ���� ������ �������� �� ���������� ���� ������ �������
  ScrollBox2.Visible:= DomainOrLocalAccount;   /// ���� ������ ��������� �� ������ ������
  end;

if (FWbemObject.SID<>null) then  result.Add ('SID : '+string(FWbemObject.SID));

if VarIsNumeric(FWbemObject.SIDType) then  result.Add ('��� SID : '+string(SidType[integer(FWbemObject.SIDType)]));


if (FWbemObject.Disabled<>null) then
begin                     /////  ��������� ������� ������
if  DomainOrLocalAccount then
  begin
   CheckBox1.Checked:=(FWbemObject.Disabled);
   CheckBox1.tag:=0;
  end;
end;

{if (FWbemObject.Lockout<>null) then
begin
 ///result.Add('��������������� ������� ������: '+boltostr(FWbemObject.Lockout));
  //CheckBox2.Checked:=(FWbemObject.Lockout);
end;}

if (FWbemObject.PasswordChangeable<>null) then
begin      ///// ��������� ����� ������ �������������
if DomainOrLocalAccount then
 begin
 CheckBox3.Checked:=not(FWbemObject.PasswordChangeable);
 CheckBox3.Tag:=0;
 end;
end;


if (FWbemObject.PasswordExpires<>null) then  ///
 begin
  if DomainOrLocalAccount then
  begin
    CheckBox2.Checked:=not(FWbemObject.PasswordExpires);
    CheckBox2.Tag:=0;
    end;
  end;

//////////////// ���� ������������ �������� �� ������� ���������� � ������
if not DomainOrLocalAccount then checkPrivelegDomainUser(CurrentAccountName,CurrentDomain);

//if (FWbemObject.PasswordRequired<>null) then  result.Add('������� ������: '+boltostr(FWbemObject.PasswordRequired));
//step:='16';
if (FWbemObject.status<>null) then  result.Add ('������: '+string(FWbemObject.status));
step:='17';
///////////////////////////////////// ������ � ������ ������������
if DomainOrLocalAccount then
begin
ListViewGroup.Clear;
GroupForUs:=TStringList.Create;
GroupForUs:=GroupForLocalAccount(FWMIService,NamePC,string(FWbemObject.Domain),CurrentAccountName);
for I := 0 to GroupForUs.Count-1 do
with ListViewGroup.Items.Add do
  begin
  Caption:=GroupForUs[i];
  ImageIndex:=0;
  end;
GroupForUs.Free;
end
else
if frmDomainInfo.LabeledEdit3.Text<>'' then  //// ���� � ������
begin
if (FWbemObject.Name<>null)and (FWbemObject.LocalAccount<>null) then
if not (FWbemObject.LocalAccount) then
  begin
  ListViewGroup.Clear;
  GroupForUs:=TStringList.Create;
  GroupForUs:=frmDomainInfo.GetAllUserGroups(string(FWbemObject.Name));
  if GroupForUs.Count<>0 then
    for I := 0 to  GroupForUs.Count-1 do
      with ListViewGroup.Items.Add do
      begin
      Caption:=GroupForUs[i];
      ImageIndex:=0;
      end;
  GroupForUs.Free;
  end;
//////////////////////////////////////////////////
checkBoxTag0; /// ��������� ������ ���������� ��������� tag ��� ���������
end;


FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    //result.Add('��� - '+step+' / ������: '+E.Message);
    memo1.Lines.Add(NamePC+' - ������ ������ ������� "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

end.
