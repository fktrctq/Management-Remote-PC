unit PropertyGroup;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ImgList, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.ExtCtrls,ActiveX,ComObj, Vcl.StdCtrls,ActiveDs_TLB, Vcl.Menus, JvBaseDlg,
  JvObjectPickerDialog;

type
  TFormPropertyGroup = class(TForm)
    Panel1: TPanel;
    ListViewAccout: TListView;
    SpeedButton1: TSpeedButton;
    ImageListAccount: TImageList;
    LabelGroup: TLabel;
    LabelDomain: TLabel;
    LabelUser: TLabel;
    LabelPasswd: TLabel;
    LabelNamePC: TLabel;
    PopupLocal: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Button1: TButton;
    JvObjectPickerDialogGroup: TJvObjectPickerDialog;
    Button2: TButton;
    C1: TMenuItem;
    SpeedButton2: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    function AccountForLocalGroup(NameGroup,NamePC,Domain,AdmUser,AdmPass:string):TstringList;   //// получение списка пользователей и групп входящих в локальную группу
    function AccountDomainDeleteGroup(StrGroup,user,GrDomain,UsDomain,UserAdmin,AdminPasswd:string):bool;    ///удаление пользователя из группы
    function AccountDomainAddGroup(StrGroup:string;user,domain,UserAdmin,AdminPasswd:string):bool;  /// добавление пользователя в группу
    function GroupDomainAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool; //// добавление группы в группу

    function GroupDomainDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool; //// удаление группы из группы
    function ListGroupDomainAddGroup(StrGroup,domain,LocalGroup,LocalDomain,UserAdmin,AdminPasswd,UserOrGroup:string):bool; /// Добавление списка доменных групп в группу
    function GroupSystemAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool; /// добавляем системную встроеную группу в группу безопастности
    function GroupSystemDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool; /// удаление системных пользователей из групп
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure FormGroupShowAccount(Sender: TObject); //// при загрузке отображает список пользователей
    procedure FormGroupShowGroup(Sender: TObject);   /////// при загрузке отображает список групп
    procedure FormGroupShowSystem(Sender: TObject);  // при загрузке отображает системные группы
    procedure FormGroupClose(Sender: TObject; var Action: TCloseAction);
    procedure ButOkClickAccount(Sender: TObject); //// добавляет пользователя
    procedure ButOkClickGroup(Sender: TObject);  // добавляет группу
    procedure ButOkClickSystem(Sender: TObject);  // системные группы
    function  AllLocalGroup(NamePC,User,Pass:string):TstringList; /// локальные группы
    function LoadAccoutUser (NamePC,User,Pass:string):TstringList; // локальные пользователи
    function AllLocalSystem(NamePC,User,Pass:string):TstringList; /// системные группы
    function readgroupForDomain(NamePC,Domain:string):TstringList; /// спискок из домена
  public
    { Public declarations }
  end;

var
  FormPropertyGroup: TFormPropertyGroup;
  FormGr:Tform;
  ListGr:Tcombobox;
  ButOk:Tbutton;

const
wbemFlagForwardOnly = $00000020;
ADS_SECURE_AUTHENTICATION = $00000001;

implementation
{$R *.dfm}

uses PropertyAccount,umain;


function TFormPropertyGroup.readgroupForDomain(NamePC,Domain:string):TstringList;
var          ////// список групп безопастности для добавления пользователю
i:integer;
recUsGr:string;
begin
try      ///https://docs.microsoft.com/en-us/windows/win32/api/objsel/ns-objsel-dsop_filter_flags
         ///https://docs.microsoft.com/ru-ru/windows/win32/api/objsel/ns-objsel-dsop_scope_init_info
OleInitialize(nil);
result:=TStringList.Create;
//JvObjectPickerDialog1.TargetComputer:=NamePC;
JvObjectPickerDialogGroup.Scopes[0].DcName:= frmDomainInfo.GetDomainController(CurrentDomainName); /// указываем контроллер домена
if JvObjectPickerDialogGroup.Execute() then
begin
if JvObjectPickerDialogGroup.Selection.Count<>0 then
for I := 0 to JvObjectPickerDialogGroup.Selection.Count-1 do
  begin        //// имя выбранного объекта ()                          /// тип выбранного объекта, user/group/ securetygroup -типа того что то
  recUsGr:='';

  if JvObjectPickerDialogGroup.Selection.Items[i].UPN<>'' then
  begin
   recUsGr:=JvObjectPickerDialogGroup.Selection.Items[i].UPN; /// если пользователь то UPN  - т.к. в name отображается полное имя а не логин
   if pos('@',recUsGr)<>0 then
   recUsGr:=copy(recUsGr,1,pos('@',recUsGr)-1); /// нужно только имя без домена
  end
  else recUsGr:=JvObjectPickerDialogGroup.Selection.Items[i].Name;                                                      /// иначе name - т.к. для групп name отображает уникальное название группы
  result.add(recUsGr+'='+JvObjectPickerDialogGroup.Selection.Items[i].ObjectClass);
   ///JvObjectPickerDialogGroup.Selection.Items[i].ObjectClass - сдесь можно узнать, группа это или пользователь
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

function TFormPropertyGroup.AllLocalSystem(NamePC,User,Pass:string):TstringList;
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
Result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,LocalAccount FROM Win32_SystemAccount WHERE LocalAccount=true','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
 begin
  nameGr:='';
  if (FWbemObject.name)<>null then nameGr:=string(FWbemObject.name);
  result.Add(nameGr);
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
    result.Add('ошибка-'+e.Message);
    OleUnInitialize;
    end;
  end;
end;


function TFormPropertyGroup.AllLocalGroup(NamePC,User,Pass:string):TstringList;
var            ///// локальные группы
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  nameGr:string;
function typeGroup(gr:string):integer;
const
Tgroup: Array [0..10] of string=('IIS_IUSRS','Криптографические операторы'
,'Операторы настройки сети','Пользователи DCOM','Пользователи журналов производительности'
,'Опытные пользователи','Пользователи системного монитора','Пользователи удаленного рабочего стола'
,'Репликатор','Читатели журнала событий','');
begin

end;
begin
try
OleInitialize(nil);
Result:=TStringList.Create;
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,LocalAccount FROM Win32_Group WHERE LocalAccount=true','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
 begin
  nameGr:='';
  if (FWbemObject.name)<>null then nameGr:=string(FWbemObject.name);
  result.Add(nameGr);
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
    result.Add('ошибка-'+e.Message);
    OleUnInitialize;
    end;
  end;
end;

function TFormPropertyGroup.LoadAccoutUser (NamePC,User,Pass:string):TstringList;
var                           ///// локальные пользователи
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  oEnum           : IEnumvariant;
  iValue          : LongWord;
  EnDis:integer;
  SID,NameAc:string;
begin
    try
    OleInitialize(nil);
    result:=Tstringlist.Create;
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
    FWMIService.Security_.impersonationlevel:=3;
    FWMIService.Security_.authenticationLevel := 6;
    FWbemObjectSet:= FWMIService.ExecQuery('SELECT name FROM Win32_UserAccount WHERE LocalAccount = true','WQL',wbemFlagForwardOnly);
    oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
        NameAc:='';
        if (FWbemObject.name)<>null then nameAc:=string(FWbemObject.name);
        result.Add(nameAc);
      FWbemObject:=Unassigned;
      end;
    oEnum:=nil;
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    variantclear(FSWbemLocator);

   except
    on E:Exception do
    begin
    result.Add('ошибка-'+E.Message);
    oEnum:=nil;
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    variantclear(FSWbemLocator);
    end;
   end;
 oleUnInitialize;
end;

procedure TFormPropertyGroup.N1Click(Sender: TObject);
begin
FormGroupCreate(LabelNamePC.Caption,'Account','');
end;

procedure TFormPropertyGroup.N2Click(Sender: TObject);
begin
FormGroupCreate(LabelNamePC.Caption,'Group','');
end;



procedure TFormPropertyGroup.FormGroupShowAccount(Sender: TObject);
var                      ///// загрузка списков пользователей  при загрузке формы
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=LoadAccoutUser(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption);  /// заполняем комбобокс списком локальных пользователей
ListGr.Clear;
for I := 0 to listUser.Count-1 do
 ListGr.Items.Add(listUser[i]);
listUser.Free;
if ListGr.Items.Count>0 then
begin
ListGr.ItemIndex:=0;
ListGr.AutoDropDown:=true;
end;

end;

procedure TFormPropertyGroup.FormGroupShowGroup(Sender: TObject);
var                      ///// загрузка списков групп при загрузке формы
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=AllLocalGroup(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption); // заполняем комбо бокс списком локальных групп
ListGr.Clear;
for I := 0 to listUser.Count-1 do
 ListGr.Items.Add(listUser[i]);
listUser.Free;

if ListGr.Items.Count>0 then
begin
ListGr.ItemIndex:=0;
ListGr.AutoDropDown:=true;
end;
end;

procedure TFormPropertyGroup.FormGroupShowSystem(Sender: TObject);
var                      ///// загрузка списков групп при загрузке формы
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=AllLocalSystem(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption); // заполняем комбо бокс списком локальных групп
ListGr.Clear;
for I := 0 to listUser.Count-1 do
 ListGr.Items.Add(listUser[i]);
listUser.Free;

if ListGr.Items.Count>0 then
begin
ListGr.ItemIndex:=0;
ListGr.AutoDropDown:=true;
end;
end;

procedure TFormPropertyGroup.FormGroupClose(Sender: TObject; var Action: TCloseAction);
begin
FormGr.Release; /// уничтожение формы
end;

procedure TFormPropertyGroup.ButOkClickAccount(Sender: TObject);
var
i:integer;
begin
if AccountDomainAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// добавляет пользователя в группу
then //// если операция выполнена успешно то добавляем пользователя в список
  begin
     with ListViewAccout.Items.Add do
     begin
     ImageIndex:=0;
     Caption:=('');
     SubItems.Add(ListGr.Text);
     SubItems.Add(labelDomain.Caption);
     end;
   end;
FormGr.Close;
end;

procedure TFormPropertyGroup.ButOkClickGroup(Sender: TObject);
var
i:integer;
begin
if GroupDomainAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// добавляет группу в группу
then //// если операция выполнена успешно то добавляем группу в список
  begin
     with ListViewAccout.Items.Add do
     begin
     ImageIndex:=1;
     Caption:=('');
     SubItems.Add(ListGr.Text);
     SubItems.Add(labelDomain.Caption);
     end;
   end;
FormGr.Close;
end;

procedure TFormPropertyGroup.ButOkClickSystem(Sender: TObject);
var
i:integer;
begin
if GroupSystemAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// добавляет группу в группу
then //// если операция выполнена успешно то добавляем системную группу в список
  begin
     with ListViewAccout.Items.Add do
     begin
     ImageIndex:=2;
     Caption:=('');
     SubItems.Add(ListGr.Text);
     SubItems.Add(labelDomain.Caption);
     end;
   end;
FormGr.Close;
end;

procedure TFormPropertyGroup.Button2Click(Sender: TObject);
var
i,indexImg:integer;
listGrup:TstringList;
UsOrGr:string;
BEGIN
///listUser.Names[i]
///listUser.ValueFromIndex[i]
listGrup:=TStringList.Create;
listGrup:=readgroupForDomain(LabelNamePC.Caption,CurrentDomainName);  //// запрашиваем списко групп и пользователей
if listGrup.Count<>0 then /// если список не пуст, т.е. если добавили хоть одну группу, можно же отказаться от добавления
Begin
  for I := 0 to listGrup.Count-1 do
  begin
   UsOrGr:='';
   UsOrGr:=listGrup.ValueFromIndex[i];
   if UsOrGr='foreignSecurityPrincipal' then UsOrGr:='user';
   if ListGroupDomainAddGroup
    (listGrup.Names[i],  //// добавляемая группа
    frmDomainInfo.LabeledEdit3.Text,    /// домен группы
    LabelGroup.Caption, /// в какую локальную группу добавляем
    labelNamePC.caption,    //// локальный домен или комп в группу которого добавляем доменную группу
    labeluser.Caption,             //// пользователь
    Labelpasswd.Caption,         //// пароль
    UsOrGr) /// пользователь или группа
    then
       with ListViewAccout.Items.Add do
       begin
       indexImg:=2;                                           /// неизвесно
       if listGrup.ValueFromIndex[i]='user' then indexImg:=0; /// пользователь
       if listGrup.ValueFromIndex[i]='group' then indexImg:=1; /// группа
       if listGrup.ValueFromIndex[i]= 'foreignSecurityPrincipal'then  indexImg:=3; /// foreignsecurityprincipals
       ImageIndex:=indexImg;
       Caption:=('');
       SubItems.Add(listGrup.Names[i]);
       SubItems.add(frmDomainInfo.LabeledEdit3.Text);
       end;
  end;
End;
listGrup.free;
END;

procedure TFormPropertyGroup.C1Click(Sender: TObject);
begin
FormGroupCreate(LabelNamePC.Caption,'System','');
end;

function TFormPropertyGroup.FormGroupCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormUserAccount);
FormGr.Name:='FormGroup';
FormGr.Width:=323;
FormGr.Height:=97;
FormGr.BorderStyle:=bsDialog;

if usr='Account' then  /// если добавляем пользователя
begin
FormGr.OnShow:=FormGroupShowAccount;
FormGr.Caption:='Локальные пользователи - '+namePC;
end;
if usr='Group' then  /// если добавляем группу
begin
FormGr.OnShow:=FormGroupShowGroup;
FormGr.Caption:='Локальные группы - '+namePC;
end;

if usr='System' then  /// если добавляем группу
begin
FormGr.OnShow:=FormGroupShowSystem;
FormGr.Caption:='Системные группы - '+namePC;
end;

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

if usr='Account' then  /// если добавляем пользователя
ButOk.OnClick:=ButOkClickAccount;
if usr='Group' then    /// если добавляем группу
ButOk.OnClick:=ButOkClickGroup;
if usr='System' then    /// если встроеную группы
ButOk.OnClick:=ButOkClickSystem;

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

function TFormPropertyGroup.AccountForLocalGroup(NameGroup,NamePC,Domain,AdmUser,AdmPass:string):TstringList;
var
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  groupOrUs,groupSTR,UsDom,allstring:string;

begin
try
OleInitialize(nil);
result:=tstringlist.Create;
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', AdmUser, Admpass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_GroupUser','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
 begin
 groupSTR:='';
 groupOrUs:='';
 allstring:='';
 if (FWbemObject.GroupComponent<>null)and(FWbemObject.PartComponent<>null) then
  begin
   if pos(Ansiuppercase('Domain="'+Domain+'",Name="'+NameGroup+'"'),Ansiuppercase(string(FWbemObject.GroupComponent)))<>0 then
     begin
     groupSTR:='';
     groupOrUs:='0'; /// Группа или пользователь
     UsDom:='';
     if (pos(Ansiuppercase('Win32_UserAccount'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0)
       then groupOrUs:='0' /// если пользователь то 0
     else /// иначе возможно это группа
     if (pos(Ansiuppercase('Win32_Group'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0)
      then groupOrUs:='1' ///  это группа - 1
     else
     if (pos(Ansiuppercase('Win32_SystemAccount'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0) then
      groupOrUs:='2'
      else groupOrUs:='3'; ///иначе  хуй знает

     UsDom:=(string(FWbemObject.PartComponent)); ///// домен пользователя
     UsDom:=StringReplace(UsDom,'"','',[rfReplaceAll]);
     delete(UsDom,pos(',Name=',UsDom),length(UsDom));
     UsDom:=copy(UsDom,pos('.Domain=',UsDom)+8,length(UsDom));

     groupSTR:=(string(FWbemObject.PartComponent));  //// имя пользователя
     groupSTR:=StringReplace(groupSTR,'"','',[rfReplaceAll]);
     result.add(groupOrUs+copy(groupSTR,pos(',Name=',groupSTR)+6,length(groupSTR)-1)+'='+UsDom);
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
    result.add('Ошибка чтения пользователей группы "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

function TFormPropertyGroup.ListGroupDomainAddGroup(StrGroup,domain,LocalGroup,LocalDomain,UserAdmin,AdminPasswd,UserOrGroup:string):bool;
var              ////////////////// добавление списка групп для группы
  z:integer;
  InterF : IADsGroup;
  GroupUs: IADsGroup;
  UserUs : IADsUser;

begin
  try
  //ShowMessage('Локальный  - '+'WinNT://'+localDomain+'/'+LocalGroup+',Group'+UserAdmin+AdminPasswd);
  //ShowMessage('Домен - '+'WinNT://'+Domain+'/'+StrGroup+','+UserOrGroup+UserAdmin+AdminPasswd);


  if UserOrGroup='user' then
  begin
  OleCheck(ADsOpenObject('WinNT://'+localDomain+'/'+LocalGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, UserUs));
  InterF.Add(UserUs.ADsPath);
  InterF.SetInfo;
  UserUs._Release;
  end;

  if UserOrGroup='group' then
  begin
  OleCheck(ADsOpenObject('WinNT://'+localDomain+'/'+LocalGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
  InterF.Add(GroupUs.ADsPath);
  InterF.SetInfo;
  GroupUs._Release;
  end;

  InterF._Release; ///
  result:=true;
 ///////////////////////////////////////////////////////////////////////
  except
    on E: EOleException do
    begin
      result:=false;
      ShowMessage('"'+StrGroup+'" '+E.message);
    end;
end;
end;

function TFormPropertyGroup.AccountDomainAddGroup(StrGroup:string;user,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// добавление одного пользователя в группы
  z:integer;
  InterF :  IADsUser;
  GroupUs: IADsGroup;

begin
  try
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
  GroupUs.Add(InterF.ADsPath);
  GroupUs.SetInfo;

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

function TFormPropertyGroup.GroupSystemAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// добавление NT AUTHORITY - системные встроеные объекты безопасности в локальную группу
  z:integer;
  InterF : IADsGroup;
  GroupUs: IADsGroup;

begin
  try
  //OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+AddGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
  //GroupUs.Add(InterF.ADsPath);
  GroupUs.Add('WinNT://NT AUTHORITY/'+AddGroup);
  GroupUs.SetInfo;
  GroupUs._Release;//
  //InterF._Release; ///
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

function TFormPropertyGroup.GroupDomainAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// добавление группу в группу
  z:integer;
  InterF :  IADsGroup;
  GroupUs: IADsGroup;

begin
  try
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+AddGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
  OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
  GroupUs.Add(InterF.ADsPath);
  GroupUs.SetInfo;

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


function TFormPropertyGroup.AccountDomainDeleteGroup(StrGroup,user,GrDomain,UsDomain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// удаление пользователя из группы
  InterF :  IADsUser;
  GroupUs: IADsGroup;
begin
  try
    OleCheck(ADsOpenObject('WinNT://'+UsDomain+'/'+User+',user',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsUser, InterF));
    OleCheck(ADsOpenObject('WinNT://'+GrDomain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
    GroupUs.Remove(InterF.ADsPath);
    GroupUs.SetInfo;
    GroupUs._Release;
    InterF._Release;
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

function TFormPropertyGroup.GroupDomainDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// удаление группы из группы
  InterF :  IADsGroup;
  GroupUs:  IADsGroup;
begin
  try
    OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+DelGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
    OleCheck(ADsOpenObject('WinNT://'+Localdomain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
    GroupUs.Remove(InterF.ADsPath);
    GroupUs.SetInfo;
    GroupUs._Release;
    InterF._Release;
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

function TFormPropertyGroup.GroupSystemDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// удаление системных групп и пользоватлей из группы
  //InterF :  IADsGroup;
  GroupUs:  IADsGroup;
begin
  try
   // OleCheck(ADsOpenObject('WinNT://'+Domain+'/'+DelGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, InterF));
    OleCheck(ADsOpenObject('WinNT://'+Localdomain+'/'+StrGroup+',Group',UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsGroup, GroupUs));
    GroupUs.Remove('WinNT://NT AUTHORITY/'+DelGroup);
    GroupUs.SetInfo;
    GroupUs._Release;
    //InterF._Release;
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


procedure TFormPropertyGroup.FormShow(Sender: TObject);
var
listUser:TstringList;
i:integer;
s:string;
begin
FormPropertyGroup.Caption:=LabelGroup.Caption+' - '+labelNamePC.Caption;
ListViewAccout.Clear;
listUser:=TstringList.Create;
listUser:=AccountForLocalGroup(LabelGroup.Caption,labelNamePC.Caption,LabelDomain.Caption,LabelUser.Caption,LabelPasswd.Caption);
for I := 0 to listUser.Count-1 do
with ListViewAccout.Items.Add do
  begin
  Caption:='';
  s:='';
  s:=copy(listUser.Names[i],1,1);
  if (s='1') or (s='0') then ImageIndex:=strtoint(s)
   else ImageIndex:=2;

  SubItems.Add(copy(listUser.Names[i],2,length(listUser.Names[i]))); /// первый символ в названии это индекс группа или пользователь
  SubItems.Add(listUser.ValueFromIndex[i]);   ////// домен пользователя
  end;
listUser.Free;
end;

procedure TFormPropertyGroup.SpeedButton1Click(Sender: TObject);
var
i:integer;
begin
try                              ///// удалить запись, группу или пользователя
if (ListViewAccout.Items.Count>0) and (ListViewAccout.ItemIndex<>-1) then
Begin
 i:=MessageBox(Self.Handle, PChar('Удалить запись?')
        , PChar('Запись - '+ListViewAccout.Items[ListViewAccout.itemindex].subitems[0] ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;

if ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0]<>'' then
begin
 if (ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=0) or (ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=3) then ///// если пользователь  или foreignsecurityprincipals
   begin
   if AccountDomainDeleteGroup
   (LabelGroup.Caption,  /// из какой группы удаляем
   ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0],  /// какого пользователя
   LabelNamePC.Caption,                                               /// домен группы или имя ПК
   ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1],  /// домен учетки которую удаляем
   LabelUser.Caption,                                          /// пользователь
   LabelPasswd.Caption)                                        /// пароль
   then ListViewAccout.Items[ListViewAccout.ItemIndex].delete;
   exit; /// выходим т.к далее идет проверка на группу и будет спрашивать индекс выделенной строки, а ее сука нет..... и не надо
   end;

 if ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=1 then
   begin
    if GroupDomainDeleteGroup(LabelGroup.Caption, /// Из какой группы удалять
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0], /// какую группу удалять
    labeldomain.Caption,                                        /// домен исходной группы, из которой удаляют
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1], /// из какого домена удаляемая группа
    LabelUser.Caption,LabelPasswd.Caption)
   then ListViewAccout.Items[ListViewAccout.ItemIndex].delete;
   exit;
   end;
 if ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=2 then //'\NT AUTHORITY'
   begin
    if GroupSystemDeleteGroup(LabelGroup.Caption, /// Из какой группы удалять
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0], /// какую группу удалять
    labeldomain.Caption,                                        /// домен исходной группы, из которой удаляют
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1], /// из какого домена удаляемая группа
    LabelUser.Caption,LabelPasswd.Caption)
   then ListViewAccout.Items[ListViewAccout.ItemIndex].delete;
   exit;
   end;


end;

End;
except
    on E: EOleException do
    begin
      ShowMessage(E.message);
    end;
end;
end;

procedure TFormPropertyGroup.SpeedButton2Click(Sender: TObject);
var
listUser:tstringList;
s:string;
i:integer;
begin
ListViewAccout.Clear;
listUser:=TstringList.Create;
listUser:=AccountForLocalGroup(LabelGroup.Caption,labelNamePC.Caption,LabelDomain.Caption,LabelUser.Caption,LabelPasswd.Caption);
for I := 0 to listUser.Count-1 do
with ListViewAccout.Items.Add do
  begin
  Caption:='';
  s:='';
  s:=copy(listUser.Names[i],1,1);
  if (s='1') or (s='0') then ImageIndex:=strtoint(s)
   else ImageIndex:=2;
  SubItems.Add(copy(listUser.Names[i],2,length(listUser.Names[i]))); /// первый символ в названии это индекс группа или пользователь
  SubItems.Add(listUser.ValueFromIndex[i]);   ////// домен пользователя
  end;
listUser.Free;
end;

end.
