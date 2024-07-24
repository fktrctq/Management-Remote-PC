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
    function passwordchangeWinNT(namePC,CurUser,UserAdmin,AdminPasswd,NewPass:string):integer; /// смена пароля локального пользователя
    function passwordchangeDomain(DomainName,CurUser,UserAdmin,AdminPasswd,NewPass:string):bool;  /// смена пароля доменного пользователя
    function RenameAccountDomain(DomainName,OldUser,NewName,UserAdmin,AdminPasswd:string):bool;
    function FullNameAccountDomain(DomainName,User,NewFullName,UserAdmin,AdminPasswd:string):bool;
    function FlagsAccountDomain(DomainName,User,UserAdmin,AdminPasswd:string;check:bool;operation:integer):bool;
    function AccountDomainPrivelegi(val:integer;user,domain,UserAdmin,AdminPasswd:string):bool; /// проверка на существование привилегии пользователя лдомена по маске
    function checkPrivelegDomainUser(user,domain:string):bool;  /// проверка все привелегий пользователя домена
    function AccountDomainAddGroup(StrGroup:TstringList;user,domain,UserAdmin,AdminPasswd:string):bool;
    function AccountDomainDeleteGroup(StrGroup,user,domain,UserAdmin,AdminPasswd:string):bool;

    function AllLocalGroup(NamePC,User,Pass:string):bool;
    function readgroupForLocal(NamePC,Domain:string):TstringList;
    function readgroupForDomain(NamePC,Domain:string):TstringList;
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure FormGroupShow(Sender: TObject); /// проседура динамически создаваемой формы
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
frmDomainInfo.LabeledEdit2.Text); // заполняем комбо бокс списком локальных групп
if ListGr.Items.Count>0 then
begin
ListGr.ItemIndex:=0;
ListGr.AutoDropDown:=true;
end;

end;

procedure TFormUserAccount.FormGroupClose(Sender: TObject; var Action: TCloseAction);
begin
FormGr.Release; /// уничтожение формы
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
    // ShowMessage('Операция выполнена');
   end;
 groupList.Free;
FormGr.Close;
end;

function TFormUserAccount.FormGroupCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormUserAccount);
FormGr.Name:='FormGroup';
FormGr.Caption:='Список локальных групп - '+namePC;
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
ButOk.Caption:='Ок';
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
var            ///// локальные группы
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
frmDomainInfo.memo1.Lines.Add(NamePC+' - Загрузка списка групп, ожидайте... ');
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,Domain FROM Win32_Group WHERE Domain="'+NamePC+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
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
var          ////// список групп безопастности для добавления пользователю
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
    showmessage('Ошибка : '+E.Message);
    end;
  end;
OleUnInitialize;
end;

function TFormUserAccount.readgroupForDomain(NamePC,Domain:string):TstringList;
var          ////// список групп безопастности для добавления пользователю
i:integer;
begin
try      ///https://docs.microsoft.com/en-us/windows/win32/api/objsel/ns-objsel-dsop_filter_flags
         ///https://docs.microsoft.com/ru-ru/windows/win32/api/objsel/ns-objsel-dsop_scope_init_info
OleInitialize(nil);
result:=TStringList.Create;
//JvObjectPickerDialog1.TargetComputer:=NamePC;
JvObjectPickerDialogDomain.Scopes[0].DcName:= frmDomainInfo.GetDomainController(CurrentDomainName); /// указываем контроллер домена
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
    showmessage('Ошибка : '+E.Message);
    end;
  end;
OleUnInitialize;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////
function TFormUserAccount.AccountDomainDeleteGroup(StrGroup,user,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// удаление пользователя из группы
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
var              ////////////////// добавление пользователя в группы
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
var              ////////////////// чтение флагов для доменного пользователя
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
  // frmDomainInfo.Memo1.Lines.add('флаги - '+ string(InterF.Get('UserFlags')));
 //////////////////////////////////////////////////////// рабочая функция , без указания логина и пароля админа, работает в контексте пользователя
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
//checkbox5.checked:=AccountDomainPrivelegi(8388608,user,domain);  //Требовать смены пароля при следующем входе
checkbox6.checked:=AccountDomainPrivelegi(64,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); /// запретить смену пароля пользователем
checkbox7.checked:=AccountDomainPrivelegi(65536,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //срок действия пароля  не ограничен
checkbox13.checked:=AccountDomainPrivelegi(2,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); /// Включена или отключена учетная запись
checkbox8.checked:=AccountDomainPrivelegi(128,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); // Хранить пароль, используя обратимое шифрование
checkbox12.checked:=AccountDomainPrivelegi(262144,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); //Для интерактивного входа в сеть нужна смарт-карта
checkbox11.checked:=AccountDomainPrivelegi(1048576,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text); // Учетная запись важна и не может быть делегированна
checkbox10.checked:=AccountDomainPrivelegi(2097152,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //Использовать типы шифрования Kerberos DES для этой
checkbox15.checked:=AccountDomainPrivelegi(4194304,user,domain,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);  //Без предварительная проверка подлинности Kerberos
end;

procedure TFormUserAccount.ButtonRenameClick(Sender: TObject);
var
s,NewFullName:string;
begin
NewFullName:='';
if not InputQuery('Введите новое полное имя пользователя', 'Имя пользователя:', NewFullName)
 then exit;

if DomainOrLocalAccount then
  begin
  s:=RenameAccount(frmDomainInfo.ComboBox2.Text,'FullName',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,NewFullName);
  if s='' then  showmessage('Произошла ошибка')
  else
  begin
  frmDomainInfo.Memo1.Lines.Add('Операция: "Замена полного имени" завершилась: '+s);
  LabeledEdit1.Text:=NewFullName;
  end;
  end
else
  begin
  if FullNameAccountDomain(CurrentDomain,CurrentAccountName,
  NewFullName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
  begin
   showmessage('Изменение имени прошло успешно');
   LabeledEdit1.Text:=NewFullName;
  end;
  end;
end;

procedure TFormUserAccount.SpeedButton1Click(Sender: TObject);
var
s,NewName:string;
begin
NewName:='';
if not InputQuery('Введите новое имя пользователя', 'Имя пользователя:', NewName)
 then exit;
if NewName='' then
begin
  ShowMessage('Имя не может буть пустым');
  exit;
end;

if DomainOrLocalAccount then
  begin
  s:=RenameAccount(frmDomainInfo.ComboBox2.Text,'Rename',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,NewName);
  if s='' then  showmessage('Произошла ошибка')
   else
   begin
   frmDomainInfo.Memo1.Lines.Add('Операция: "Переименование" завершилась: '+s);
   CurrentAccountName:=NewName; /// если периемовали и все ок, то присваиваем переменной новое имя
   LabeledEdit2.Text:=NewName;
   end;
  end
else
  begin
  if RenameAccountDomain(CurrentDomain,CurrentAccountName,NewName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
   begin
   showmessage('Пользователь переименован');
   CurrentAccountName:=LabeledEdit2.Text; /// если периемовали и все ок, то присваиваем переменной новое имя
   LabeledEdit2.Text:=NewName;
   end;
  end;
end;


function TFormUserAccount.FlagsAccountDomain(DomainName,User,UserAdmin,AdminPasswd:string;check:bool;operation:integer):bool;
var              ////////////////// установка флагов для доменного пользователя
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
var              ////////////////// смена пароля локального пользователя
  Pusr: OleVariant;
  InterF :  IADsUser;
begin
  try
  //////////////////////////////////////////////////////////////// Отказано в доступе для сетей без AD,
 { result:=NetUserChangePassword(PWideChar(WideString(NamePC)),
     PWideChar(WideString(CurUser)),
     PWideChar(WideString(OldPass)),
     PWideChar(WideString(NewPass)));}
 ///////////////////////////////////////////////////////////////////////
  ///////// смена пароля  с подклчением вводом пароля и пользователя/////////////////////////////////////
    OleCheck(ADsOpenObject('WinNT://'+NamePC+'/'+CurUser+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
    //ShowMessage(InterF.Description);
    InterF.SetPassword(NewPass);
    /////////////////////////////////////////////////////////////////////////
   ///////////////////////////////////////////  сена пароля с подключение без ввода пользователя и пароля, к контексте безопастности текущего пользователя
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
var              ////////////////// смена полного имени доменного пользователя
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
var              ////////////////// смена имени доменного пользователя
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
var              ////////////////// смена пароля доменного пользователя
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

if DomainOrLocalAccount then // Если пользователь локальный
  begin
   i:=MessageBox(Self.Handle, PChar('Изменить пароль пользователя?')
        , PChar('Пользователь - '+CurrentAccountName ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
   if not InputQuery('Введите новый пароль пользователя', 'Пароль:', newpass) then exit
   else if not InputQuery('Повторите ввод пароля', 'Пароль:', newpass1) then exit;
   if newpass<>newpass1 then
   begin
   ShowMessage('Пароли не совпадают!'+#10#13+'Повторите еще раз.');
   exit;
   end;
  i:=passwordchangeWinNT(frmDomainInfo.ComboBox2.Text,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,
  frmDomainInfo.LabeledEdit2.Text,newpass); ShowMessage(SysErrorMessage(i));
  end
else    //// иначе доменная учетка
  begin
  i:=MessageBox(Self.Handle, PChar('Изменить пароль пользователя?')
      , PChar('Пользователь - '+CurrentAccountName ) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then exit;
 if not InputQuery('Введите новый пароль пользователя', 'Пароль:', newpass) then exit
 else if not InputQuery('Повторите ввод пароля', 'Пароль:', newpass1) then exit;
 if newpass<>newpass1 then
 begin
 ShowMessage('Пароли не совпадают!'+#10#13+'Повторите еще раз.');
 exit;
 end;
  if passwordchangeDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,
frmDomainInfo.LabeledEdit2.Text,newpass) then ShowMessage('Пароль установлен');
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

if groupList.Count<>0 then /// если список не пуст, т.е. если добавили хоть одну группу, можно же отказаться от добавления
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
i:=MessageBox(Self.Handle, PChar('Удалить группу?')
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
if CheckBox1.Tag=1 then exit; // если выставили чек программно то выходим
if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;

if CheckBox1.Checked then v:='Отключить'
else v:='Включить';
i:=MessageBox(Self.Handle, PChar(v+' учетную запись?')
      , PChar(CurrentAccountName) ,MB_YESNO+MB_ICONQUESTION);
 if i=IDNO then
 begin
  CheckBox1.Tag:=1;
  CheckBox1.Checked:=not CheckBox1.Checked; /// возвращаем значение
  CheckBox1.Tag:=0;
  exit;
 end;

 if i=IDYes then
 begin
   if DomainOrLocalAccount then //// если локальная учетка
  begin
  s:=ChengeAccount(RemotePC,'Disabled',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox1.Checked);
  if s='' then ShowMessage('Ошибка операции')else
  frmDomainInfo.Memo1.Lines.Add(v+' учетную запись: '+s);
  end;
{else  /// если доменая
  begin
  if FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox1.Checked,1) then ShowMessage('Операция "'+v+' учетную запись" выполнена');
  end;}
 end;

end;



procedure TFormUserAccount.CheckBox2Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin

if CheckBox2.Tag=1 then exit; // если выставили чек программно то выходим

 if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;
if DomainOrLocalAccount then
  begin
  s:=ChengeAccount(RemotePC,'PasswordExpires',SIDLab.Caption,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text, not CheckBox2.Checked);
  if s='' then ShowMessage('Ошибка операции')else
  frmDomainInfo.Memo1.Lines.Add('Срок действия пароля не ограничен: '+s);
  end;
{else
  begin
  if FlagsAccountDomain(CurrentDomain,CurrentAccountName,not CheckBox2.Checked,2) then ShowMessage('Операция успешно выполнена');
  end;}
end;

procedure TFormUserAccount.CheckBox3Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin
if CheckBox3.Tag=1 then exit; // если выставили чек программно то выходим
if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text
else
begin
ShowMessage('Только для локальных пользователей');
exit;
end;
 s:=ChengeAccount(RemotePC,'PasswordChangeable',SIDLab.Caption,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text, not CheckBox3.Checked);
if s='' then ShowMessage('Ошибка операции')else
frmDomainInfo.Memo1.Lines.Add('Запретить смену пароля пользователем: '+s)
end;


procedure TFormUserAccount.CheckBox4Click(Sender: TObject);
var
RemotePC,s,v:string;
i:integer;
begin
{CheckBox3.Enabled:=not CheckBox4.Checked;
CheckBox2.Enabled:=not CheckBox4.Checked;
if CheckBox3.Checked then CheckBox3.Checked:=false;



if CheckBox4.Tag=1 then exit; // если выставили чек программно то выходим

CheckBox2.Tag:=1;
CheckBox2.Checked:=not CheckBox4.Checked;
CheckBox2.Tag:=0;

if DomainOrLocalAccount then RemotePC:= frmDomainInfo.ComboBox2.Text;
if DomainOrLocalAccount then
begin
s:=ChengeAccount(RemotePC,'PasswordExpires',SIDLab.Caption,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox4.Checked);
frmDomainInfo.Memo1.Lines.Add('Требовать смены пароля при следующем входе в систему: '+s);
end; }

{else
begin
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox4.Checked,2) then ShowMessage('Операция успешно выполнена');
end; }
end;

procedure TFormUserAccount.CheckBox5Click(Sender: TObject);
begin   // Требовать смены пароля при следующем входе
{if CheckBox5.Tag=1 then  exit;

if CheckBox6.Checked then CheckBox6.Checked:=false; /// снимаем запрет на смену пароля
if CheckBox7.Checked then CheckBox7.Checked:=false;  /// снимаем срок действия пароля не ограничен

FlagsAccountDomain(CurrentDomain,CurrentAccountName,CheckBox5.Checked,8388608);}
end;

procedure TFormUserAccount.CheckBox6Click(Sender: TObject);
var
s:string;
begin  //запретить смену пароля пользователем
if CheckBox6.Tag=1 then exit;
if checkbox6.Checked then s:='Запретить' else s:='Разрешить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox6.Checked,64) then
 frmDomainInfo.Memo1.Lines.add(s+' смену пароля пользователем: OK');
end;

procedure TFormUserAccount.CheckBox7Click(Sender: TObject);
var
s:string;
begin  //срок действия пароля  не ограничен
if CheckBox7.Tag=1 then exit;
if checkbox7.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox7.Checked,65536) then
 frmDomainInfo.Memo1.Lines.add(s+': Cрок действия пароля  не ограничен: OK');
end;

procedure TFormUserAccount.CheckBox8Click(Sender: TObject);
var
s:string;
begin   ///Хранить пароль, используя обратимое шифрование
if CheckBox8.Tag=1 then exit;
if checkbox8.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox8.Checked,128) then
 frmDomainInfo.Memo1.Lines.add(s+': Хранить пароль, используя обратимое шифрование: OK');
end;

procedure TFormUserAccount.CheckBox10Click(Sender: TObject);
var
s:string;
begin  ///Использовать типы шифрования Kerberos DES для этой учетной записи
if CheckBox10.Tag=1 then  exit;
if checkbox10.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox10.Checked,2097152) then
 frmDomainInfo.Memo1.Lines.add(s+': Использовать типы шифрования Kerberos DES для этой учетной записи: OK');
end;

procedure TFormUserAccount.CheckBox11Click(Sender: TObject);
var
s:string;
begin    // Учетная запись важна и не может быть делегированна
if CheckBox11.Tag=1 then exit;
if checkbox11.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox11.Checked,1048576) then
 frmDomainInfo.Memo1.Lines.add(s+': Учетная запись важна и не может быть делегированна: OK');
end;

procedure TFormUserAccount.CheckBox12Click(Sender: TObject);
var
s:string;
begin  //Для интерактивного входа в сеть нужна смарт-карта
if CheckBox12.Tag=1 then exit;
if checkbox12.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox12.Checked,262144) then
 frmDomainInfo.Memo1.Lines.add(s+': Для интерактивного входа в сеть нужна смарт-карта: OK');
end;

procedure TFormUserAccount.CheckBox13Click(Sender: TObject);
var
s:string;
begin   /// отключить включить учетную запись
if CheckBox13.Tag=1 then  exit;
if not checkbox13.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox13.Checked,2) then
frmDomainInfo.Memo1.Lines.add(s+' учетную запись: OK');
end;

procedure TFormUserAccount.CheckBox15Click(Sender: TObject);
var
s:string;
begin //Без предварительная проверка подлинности Kerberos
if CheckBox15.Tag=1 then exit;
if checkbox15.Checked then s:='Включить' else s:='Отключить';
if FlagsAccountDomain(CurrentDomain,CurrentAccountName,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,CheckBox15.Checked,4194304) then
frmDomainInfo.Memo1.Lines.add(s+' предварительную проверку подлинности Kerberos: OK');
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
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
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
    result.add('Ошибка чтения групп безопасности "'+E.Message+'"');
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
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
  begin
  if proper='Rename' then
    begin
    res:=FWbemObject.rename(ProperSTR);
    result:=SysErrorMessage(res);
    if res=0 then CurrentAccountName:=ProperSTR; /// если результат ОК то присваиваем переменной новое имя
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
    //result.Add('Шаг - '+step+' / Ошибка: '+E.Message);
    frmDomainInfo.memo1.Lines.Add(NamePC+' - Ошибка изменения учетной записи "'+E.Message+'"');
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
//FWMIService.Security_.Privileges.AddAsString('SeCreateTokenPrivilege',true); // Требуется для создания основного объекта токена.
//FWMIService.Security_.Privileges.AddAsString('SeAssignPrimaryTokenPrivilege',true);  // Требуется для замены токена уровня процесса.
//FWMIService.Security_.Privileges.AddAsString('SeLockMemoryPrivilege',true);         //Требуется для блокировки страниц в памяти.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseQuotaPrivilege',true);     //Требуется настроить квоты памяти для процесса.
//FWMIService.Security_.Privileges.AddAsString('SeMachineAccountPrivilege',true);   // Требуется для добавления рабочих станций в домен.
//FWMIService.Security_.Privileges.AddAsString('SeTcbPrivilege',true);             //Требуется действовать как часть операционной системы. Держатель является частью надежной компьютерной базы.
//FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);        //Требуется для управления аудитом и журналом безопасности NT.
//FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);  //Требуется принять права собственности на файлы или другие объекты, не имея записи контроля доступа (ACE) в списке контроля доступа по усмотрению (DACL).
//FWMIService.Security_.Privileges.AddAsString('SeLoadDriverPrivilege',true);     //Требуется для загрузки или выгрузки драйвера устройства.
//FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);     //Требуется для сбора профильной информации о производительности системы.
//FWMIService.Security_.Privileges.AddAsString('SeSystemtimePrivilege',true);      // Требуется изменить системное время.
//FWMIService.Security_.Privileges.AddAsString('SeProfileSingleProcessPrivilege',true); //Требуется для сбора информации профиля для одного процесса.
//FWMIService.Security_.Privileges.AddAsString('SeIncreaseBasePriorityPrivilege',true);  // Требуется для увеличения приоритета планирования
//FWMIService.Security_.Privileges.AddAsString('SeCreatePagefilePrivilege',true);       // Требуется для создания файла подкачки.
//FWMIService.Security_.Privileges.AddAsString('SeCreatePermanentPrivilege',true);      //Требуется для создания постоянных общих объектов.
//FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege',true);               //Требуется для резервного копирования файлов и каталогов, независимо от списка ACL, указанного для файла.
//FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',true);            //Требуется для восстановления файлов и каталогов независимо от списка ACL, указанного для файла.
//FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   Требуется выключить локальную систему
//FWMIService.Security_.Privileges.AddAsString('SeDebugPrivilege',true);             //Требуется для отладки и настройки памяти процесса, принадлежащего другой учетной записи.
//FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);             //Требуется для создания записей аудита в журнале безопасности NT. Только защищенные серверы должны иметь эту привилегию.
//FWMIService.Security_.Privileges.AddAsString('SeSystemEnvironmentPrivilege',true); // Требуется для изменения энергонезависимой оперативной памяти систем, которые используют этот тип памяти для хранения данных конфигурации.
//FWMIService.Security_.Privileges.AddAsString('SeChangeNotifyPrivilege',true);     //  Требуется для получения уведомлений об изменениях файлов или каталогов и обхода проверок доступа. Эта привилегия включена по умолчанию для всех пользователей.
//FWMIService.Security_.Privileges.AddAsString('SeRemoteShutdownPrivilege',true);   //Требуется выключить удаленный компьютер.
//FWMIService.Security_.Privileges.AddAsString(' SeUndockPrivilege',true);           //Требуется снять ноутбук с док-станции.
//FWMIService.Security_.Privileges.AddAsString('SeSyncAgentPrivilege',true);         // Требуется для синхронизации данных службы каталогов.
//FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);   // Требуется, чтобы учетные записи компьютеров и пользователей были доверенными для делегирования.
//FWMIService.Security_.Privileges.AddAsString('SeManageVolumePrivilege',true);       //Требуется для выполнения задач объемного обслуживания.
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
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
    //result.Add('Шаг - '+step+' / Ошибка: '+E.Message);
    frmDomainInfo.memo1.Lines.Add(NamePC+' - Ошибка изменения учетной записи "'+E.Message+'"');
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
CheckBox1.Tag:=1; /// для т чтобы при программной установке чеков не выполнялись процедуры
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
PageControl1.TabIndex:=0; /// при открытиии активна вкладка пользователь
sid:=SidLab.caption;
Memo1.clear;
CurrentAccountName:='';
CurrentDomain:='';
checkedAllChBox;

DomainOrLocalAccount:=true; /// сброс в локальную учетку
Memo1.Lines:=AccountPropertiesForRemotePC(frmDomainInfo.ComboBox2.Text,
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,SID);
/// если данные с удаленного компа не получили то смотрим на локальном если подключены к домену
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
  SIDType: array [1..9] of string = ('Пользователь','Группа','Домен','Псевдоним (Alias)','TypeWellKnownGroup','SID удаленной учетной записи','Не действительный SID','SID не известного типа','SID компьютера');
  function TypeAccount(id:integer):string;
  begin
  case id of
  256:result:='Временный дубликат учетной записи';
  512:result:='Обычная учетная запись';
  2048:result:='Междоменная учетная запись';
  4096:result:='Учетная запись рабочей станции';
  8192:result:='Доверительная учетная запись сервера';
  else result:='Unknown';
  end;
  end;
  function boltostr(b:Olevariant):string;
  begin
  try
    if b<>null then
      if b then result:='Да'
      else result:='Нет';
   except  result:='Unknown'
  end;
  end;

begin
try
frmDomainInfo.Memo1.Lines.Add('Загрузка данных о пользователе, ожидайте... ');
OleInitialize(nil);
CurrentDomain:='';
result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
begin

if VarIsNumeric(FWbemObject.AccountType) then  result.Add('Тип учетной записи: '+TypeAccount(FWbemObject.AccountType));

if (FWbemObject.Caption<>null) then
begin
 result.Add ('Учетная запись: '+string(FWbemObject.Caption));
end;

if (FWbemObject.Domain<>null) then
begin
CurrentDomain:=string(FWbemObject.Domain);
CurrentDomain:=AnsiLowerCase(Trim(CurrentDomain)); /// приводим к строчным буквам, удаляем пробелы
result.Add('Домен: '+CurrentDomain);
end;

if (FWbemObject.Name<>null) then
begin
result.Add('Имя: '+string(FWbemObject.Name));
CurrentAccountName:=string(FWbemObject.Name);
LabeledEdit2.Text:=string(FWbemObject.Name);
end;

if (FWbemObject.FullName<>null) then
begin
 LabeledEdit1.Text:=string(FWbemObject.FullName);
end;

if (FWbemObject.Description<>null) then
begin
 result.Add ('Описание: '+string(FWbemObject.Description));
end;

if (FWbemObject.LocalAccount<>null) then
begin
  result.Add('Локальная учетная запись: '+boltostr(FWbemObject.LocalAccount));
  DomainOrLocalAccount:=FWbemObject.LocalAccount;
  ScrollBox1.Visible:=not DomainOrLocalAccount; /// если учетка доменная то отображать одну панель свойств
  ScrollBox2.Visible:= DomainOrLocalAccount;   /// если учетка локальная то другая панель
  end;

if (FWbemObject.SID<>null) then  result.Add ('SID : '+string(FWbemObject.SID));

if VarIsNumeric(FWbemObject.SIDType) then  result.Add ('Тип SID : '+string(SidType[integer(FWbemObject.SIDType)]));


if (FWbemObject.Disabled<>null) then
begin                     /////  отключена учетная запись
if  DomainOrLocalAccount then
  begin
   CheckBox1.Checked:=(FWbemObject.Disabled);
   CheckBox1.tag:=0;
  end;
end;

{if (FWbemObject.Lockout<>null) then
begin
 ///result.Add('Заблокированная учетная запись: '+boltostr(FWbemObject.Lockout));
  //CheckBox2.Checked:=(FWbemObject.Lockout);
end;}

if (FWbemObject.PasswordChangeable<>null) then
begin      ///// запретить смену пароля пользователем
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

//////////////// если пользователь доменный то смотрим привелегии в домене
if not DomainOrLocalAccount then checkPrivelegDomainUser(CurrentAccountName,CurrentDomain);

//if (FWbemObject.PasswordRequired<>null) then  result.Add('Требует пароля: '+boltostr(FWbemObject.PasswordRequired));
//step:='16';
if (FWbemObject.status<>null) then  result.Add ('Статус: '+string(FWbemObject.status));
step:='17';
///////////////////////////////////// входит в группы безопасности
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
if frmDomainInfo.LabeledEdit3.Text<>'' then  //// если в домене
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
checkBoxTag0; /// разрешаем менять привелегии установив tag для чекбоксов
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
    //result.Add('Шаг - '+step+' / Ошибка: '+E.Message);
    memo1.Lines.Add(NamePC+' - Ошибка чтения свойств "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

end.
