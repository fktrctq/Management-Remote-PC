unit AccountLocal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ImgList, Vcl.ComCtrls, Vcl.Buttons,
  Vcl.ExtCtrls,ActiveX,ComObj,ActiveDs_TLB;

type
  TFormLocalAccount = class(TForm)
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    ListViewAccount: TListView;
    ImageListAccount: TImageList;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel2: TPanel;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    ListViewGroup: TListView;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
  private
   function LoadAccoutUser(NamePC,User,Pass:string):bool;
   function AddLocalAccount(NamePC,NewName,NewPassword,UserAdmin,AdminPasswd:string):bool;
   function DeleteLocalAccount(NamePC,DelName,UserAdmin,AdminPasswd:string):bool;
   function GroupForLocalAccount(NamePC,domain,User,Pass:string):bool;
   function AddLocalGroup(NamePC,NewName,DescriptionGr,UserAdmin,AdminPasswd:string):bool;
   function DeleteLocalGroup(NamePC,DelName,UserAdmin,AdminPasswd:string):bool;
  public
    { Public declarations }
  end;

var
  FormLocalAccount: TFormLocalAccount;
  DomainForAccount:string;
const
ADS_SECURE_AUTHENTICATION = $00000001;

implementation
uses umain,PropertyAccount,PropertyGroup;
{$R *.dfm}

function TFormLocalAccount.GroupForLocalAccount(NamePC,domain,User,Pass:string):bool;
var            ///// локальные группы
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  nameGr,DescrGr:string;
begin
try
OleInitialize(nil);
frmDomainInfo.memo1.Lines.Add(NamePC+' - Загрузка списка групп, ожидайте... ');
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,Domain,Description,LocalAccount FROM Win32_Group WHERE LocalAccount=true','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
 begin
  nameGr:='';
  DescrGr:='';
  if (FWbemObject.name)<>null then nameGr:=string(FWbemObject.name);
  if (FWbemObject.Description)<>null then DescrGr:=string(FWbemObject.Description);
  with ListViewGroup.Items.Add do
  begin
  Caption:='';
  ImageIndex:=2;
  SubItems.Add(nameGr);
  SubItems.Add(DescrGr);
  end;
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

function TFormLocalAccount.DeleteLocalAccount(NamePC,DelName,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// Удаление пользователя
  InterF :  IADsContainer;
begin
  try
     CoInitialize(nil);
     OleCheck(ADsOpenObject('WinNT://'+NamePC,UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsContainer, InterF));
     InterF.Delete('user',DelName);
     InterF._Release;
     result:=true;
   except
    on E: EOleException do
    begin
      Result:=false;
      ShowMessage(E.message);
    end;
end;
CoUninitialize;
end;


function TFormLocalAccount.DeleteLocalGroup(NamePC,DelName,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// Удаление пользователя
  InterF :  IADsContainer;
begin
  try
     CoInitialize(nil);
     OleCheck(ADsOpenObject('WinNT://'+NamePC,UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsContainer, InterF));
     InterF.Delete('group',DelName);
     InterF._Release;
     result:=true;
   except
    on E: EOleException do
    begin
      Result:=false;
      ShowMessage(E.message);
    end;
end;
CoUninitialize;
end;


function TFormLocalAccount.AddLocalGroup(NamePC,NewName,DescriptionGr,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// создание новой группы
  PUsr,Nusr: OleVariant;
  InterF :  IADsContainer;
  NewGroup : IADsGroup;
  newObject: IADs;
begin
  try
     CoInitialize(nil);
     OleCheck(ADsOpenObject('WinNT://'+NamePC,UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsContainer, InterF));
     newObject:=InterF.Create('group',NewName) as IADs;
     newObject.QueryInterface(IID_IADsGroup,NewGroup);
     NewGroup.Description:=DescriptionGr;
     NewGroup.SetInfo;

     NewGroup._Release;
     InterF._Release;
     newObject._Release;
     result:=true;
   except
    on E: EOleException do
    begin
      Result:=false;
      ShowMessage(E.message);
    end;
end;
CoUninitialize;
end;

function TFormLocalAccount.AddLocalAccount(NamePC,NewName,NewPassword,UserAdmin,AdminPasswd:string):bool;
var              ////////////////// создание нового пользователя
  PUsr,Nusr: OleVariant;
  InterF :  IADsContainer;
  NewUser : IADsUser;
  newObject: IADs;
begin
  try
     CoInitialize(nil);
     OleCheck(ADsOpenObject('WinNT://'+NamePC,UserAdmin,AdminPasswd, ADS_SECURE_AUTHENTICATION, IADsContainer, InterF));
     newObject:=InterF.Create('user',NewName) as IADs;
     newObject.QueryInterface(IID_IADsUser,NewUser);
     NewUser.FullName:=NewName;
     NewUser.SetPassword(NewPassword);
     NewUser.SetInfo;
     NewUser._Release;
     InterF._Release;
     newObject._Release;
     result:=true;
   except
    on E: EOleException do
    begin
      Result:=false;
      ShowMessage(E.message);
    end;
end;
CoUninitialize;
end;



procedure TFormLocalAccount.FormShow(Sender: TObject);
var
GroupForUs:TstringList;
i:integer;
begin
PageControl1.TabIndex:=0;
ListViewAccount.Clear;
LoadAccoutUser(frmDomainInfo.ComboBox2.Text,frmDomainInfo.LabeledEdit1.Text,
frmDomainInfo.LabeledEdit2.Text);

/////////////////////////////////
ListViewGroup.Clear;
GroupForLocalAccount(frmDomainInfo.ComboBox2.Text,DomainForAccount,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
///////////////////////////

end;

function TFormLocalAccount.LoadAccoutUser (NamePC,User,Pass:string):bool;
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
    frmDomainInfo.memo1.Lines.Add(NamePC+' - Загрузка списка пользователей, ожидайте... ');
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, pass,'','',128);
    FWMIService.Security_.impersonationlevel:=3;
    FWMIService.Security_.authenticationLevel := 6;
    FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,SID,LocalAccount,Disabled,Domain FROM Win32_UserAccount WHERE LocalAccount = true','WQL',wbemFlagForwardOnly);
    oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
        EnDis:=0;
        SID:='';
        NameAc:='';
        if (FWbemObject.disabled)<>null then 
        begin
        if (FWbemObject.disabled) then EnDis:=1
        else EnDis:=0; 
        end;
        if (FWbemObject.SID)<>null then SID:=string((FWbemObject.SID));
        if (FWbemObject.name)<>null then nameAc:=string(FWbemObject.name);
        with ListViewAccount.Items.Add do
        begin
          Caption:='';
          ImageIndex:=EnDis;
          SubItems.Add(NameAc);
          SubItems.Add(SID);
        end;
      if (FWbemObject.domain)<>null then  DomainForAccount:=string(FWbemObject.domain);
      FWbemObject:=Unassigned;
      end;
    oEnum:=nil;
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    variantclear(FSWbemLocator);
    result:=true;
   except
    on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add(NamePC+' - Ошибка загрузки списка пользователей "'+E.Message+'"');
    result:=false;
    oEnum:=nil;
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    variantclear(FSWbemLocator);
    end;
   end;
 oleUnInitialize;
end;

procedure TFormLocalAccount.SpeedButton1Click(Sender: TObject);
var
us:string;
i:integer;
begin
if (ListViewAccount.Items.Count<>0) and (ListViewAccount.SelCount<>0) then    //// имя пользователя
  begin

  if ListViewAccount.Items[ListViewAccount.itemindex].subitems[1]='' then /// если нет сида, то возможно пользователя только что создали
  begin
  us:=ListViewAccount.Items[ListViewAccount.itemindex].subitems[0]; /// запоминаем пользователя которого выбиралит
  Speedbutton2.Click; ///  обновляем список пользователей для получения все сидов

  for I := 0 to ListViewAccount.Items.Count-1 do  //проходим в цикле и ищем этого пользователя
  if ListViewAccount.Items[i].SubItems[0]=us then
    begin
    ListViewAccount.itemindex:=i;               /// находим и назначаем индекс
    break;
    end;
  end;

  FormUserAccount.SidLab.caption:='';
  FormUserAccount.SidLab.caption:=ListViewAccount.Items[ListViewAccount.itemindex].subitems[1];
  FormUserAccount.showmodal;
  end;
end;

procedure TFormLocalAccount.SpeedButton2Click(Sender: TObject);
begin
ListViewAccount.Clear;
LoadAccoutUser(frmDomainInfo.ComboBox2.Text,frmDomainInfo.LabeledEdit1.Text,
frmDomainInfo.LabeledEdit2.Text);
end;

procedure TFormLocalAccount.SpeedButton3Click(Sender: TObject);
var
NewUser,NewPass:string;
begin
if not InputQuery('Введите имя нового пользователя', 'Имя: ', NewUser)
 then exit;
if NewUser='' then
begin
  ShowMessage('Имя не может буть пустым');
  exit;
end;

if not InputQuery('Введите пароль для '+NewUser, 'Пароль: ', NewPass)
 then exit;


if  AddLocalAccount(frmDomainInfo.ComboBox2.Text,NewUser,NewPass,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text)
 then
  begin
  with ListViewAccount.Items.Add do
    begin
    ImageIndex:=0;
    Caption:= '';
    SubItems.Add(NewUser);
    SubItems.Add('');
    end;
  end;
end;

procedure TFormLocalAccount.SpeedButton4Click(Sender: TObject);
var
i:integer;
begin
if not ((ListViewAccount.Items.Count<>0) and (ListViewAccount.SelCount<>0)) then exit;

i:=MessageBox(Self.Handle, PChar('Удалить пользователя?')
        , PChar('Пользователь - '+ListViewAccount.Items[ListViewAccount.itemindex].subitems[0] ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;

   //// имя пользователя
  begin
   If   DeleteLocalAccount(frmDomainInfo.ComboBox2.Text,ListViewAccount.Items[ListViewAccount.itemindex].subitems[0],
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
  ListViewAccount.ItemFocused.Delete;
  end;
end;

procedure TFormLocalAccount.SpeedButton5Click(Sender: TObject);
var
NameGroup,GrupDescript:string;
begin

if not InputQuery('Введите имя новой группы', 'Имя: ', NameGroup)
 then exit;
if NameGroup='' then
begin
  ShowMessage('Имя не может буть пустым');
  exit;
end;

if not InputQuery('Введите описание группы '+NameGroup, 'Описание: ', GrupDescript)
 then exit;

if AddLocalGroup(frmDomainInfo.ComboBox2.Text,NameGroup,GrupDescript,
  frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
  begin
  with ListViewGroup.Items.Add do
    begin
    ImageIndex:=0;
    Caption:= '';
    SubItems.Add(NameGroup);
    SubItems.Add(GrupDescript);
    end;
  end;
end;

procedure TFormLocalAccount.SpeedButton6Click(Sender: TObject);
var
i:integer;
begin
if not ((ListViewGroup.Items.Count<>0) and (ListViewGroup.SelCount<>0)) then exit;

i:=MessageBox(Self.Handle, PChar('Удалить группу?')
        , PChar('Группа - '+ListViewGroup.Items[ListViewgroup.itemindex].subitems[0] ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;

if DeleteLocalGroup (frmDomainInfo.ComboBox2.Text,ListViewGroup.Items[ListViewgroup.itemindex].subitems[0],
frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text) then
ListViewGroup.Selected.Delete;
end;

procedure TFormLocalAccount.SpeedButton7Click(Sender: TObject);
begin
ListViewGroup.Clear;
GroupForLocalAccount(frmDomainInfo.ComboBox2.Text,DomainForAccount,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
end;

procedure TFormLocalAccount.SpeedButton8Click(Sender: TObject);
begin
{LabelGroup: TLabel;
    LabelDomain: TLabel;
    LabelUser: TLabel;
    LabelPasswd: TLabel;
    LabelNamePC: TLabel;}
if  ((ListViewGroup.Items.Count=0)) or ((ListViewGroup.SelCount=0)) then exit;

FormPropertyGroup.labelGroup.caption:=ListViewGroup.Items[ListViewgroup.itemindex].subitems[0];
FormPropertyGroup.labelNamePC.caption:=frmDomainInfo.ComboBox2.Text;
FormPropertyGroup.labelDomain.caption:= frmDomainInfo.ComboBox2.text;
FormPropertyGroup.LabelUser.caption:=frmDomainInfo.LabeledEdit1.Text;
FormPropertyGroup.LabelPasswd.caption:=frmDomainInfo.LabeledEdit2.Text;
FormPropertyGroup.show;
end;

end.
