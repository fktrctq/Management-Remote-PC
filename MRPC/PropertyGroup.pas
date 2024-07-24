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
    function AccountForLocalGroup(NameGroup,NamePC,Domain,AdmUser,AdmPass:string):TstringList;   //// ��������� ������ ������������� � ����� �������� � ��������� ������
    function AccountDomainDeleteGroup(StrGroup,user,GrDomain,UsDomain,UserAdmin,AdminPasswd:string):bool;    ///�������� ������������ �� ������
    function AccountDomainAddGroup(StrGroup:string;user,domain,UserAdmin,AdminPasswd:string):bool;  /// ���������� ������������ � ������
    function GroupDomainAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool; //// ���������� ������ � ������

    function GroupDomainDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool; //// �������� ������ �� ������
    function ListGroupDomainAddGroup(StrGroup,domain,LocalGroup,LocalDomain,UserAdmin,AdminPasswd,UserOrGroup:string):bool; /// ���������� ������ �������� ����� � ������
    function GroupSystemAddGroup(StrGroup:string;AddGroup,domain,UserAdmin,AdminPasswd:string):bool; /// ��������� ��������� ��������� ������ � ������ �������������
    function GroupSystemDeleteGroup(StrGroup,DelGroup,Localdomain,Domain,UserAdmin,AdminPasswd:string):bool; /// �������� ��������� ������������� �� �����
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure FormGroupShowAccount(Sender: TObject); //// ��� �������� ���������� ������ �������������
    procedure FormGroupShowGroup(Sender: TObject);   /////// ��� �������� ���������� ������ �����
    procedure FormGroupShowSystem(Sender: TObject);  // ��� �������� ���������� ��������� ������
    procedure FormGroupClose(Sender: TObject; var Action: TCloseAction);
    procedure ButOkClickAccount(Sender: TObject); //// ��������� ������������
    procedure ButOkClickGroup(Sender: TObject);  // ��������� ������
    procedure ButOkClickSystem(Sender: TObject);  // ��������� ������
    function  AllLocalGroup(NamePC,User,Pass:string):TstringList; /// ��������� ������
    function LoadAccoutUser (NamePC,User,Pass:string):TstringList; // ��������� ������������
    function AllLocalSystem(NamePC,User,Pass:string):TstringList; /// ��������� ������
    function readgroupForDomain(NamePC,Domain:string):TstringList; /// ������� �� ������
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
var          ////// ������ ����� ������������� ��� ���������� ������������
i:integer;
recUsGr:string;
begin
try      ///https://docs.microsoft.com/en-us/windows/win32/api/objsel/ns-objsel-dsop_filter_flags
         ///https://docs.microsoft.com/ru-ru/windows/win32/api/objsel/ns-objsel-dsop_scope_init_info
OleInitialize(nil);
result:=TStringList.Create;
//JvObjectPickerDialog1.TargetComputer:=NamePC;
JvObjectPickerDialogGroup.Scopes[0].DcName:= frmDomainInfo.GetDomainController(CurrentDomainName); /// ��������� ���������� ������
if JvObjectPickerDialogGroup.Execute() then
begin
if JvObjectPickerDialogGroup.Selection.Count<>0 then
for I := 0 to JvObjectPickerDialogGroup.Selection.Count-1 do
  begin        //// ��� ���������� ������� ()                          /// ��� ���������� �������, user/group/ securetygroup -���� ���� ��� ��
  recUsGr:='';

  if JvObjectPickerDialogGroup.Selection.Items[i].UPN<>'' then
  begin
   recUsGr:=JvObjectPickerDialogGroup.Selection.Items[i].UPN; /// ���� ������������ �� UPN  - �.�. � name ������������ ������ ��� � �� �����
   if pos('@',recUsGr)<>0 then
   recUsGr:=copy(recUsGr,1,pos('@',recUsGr)-1); /// ����� ������ ��� ��� ������
  end
  else recUsGr:=JvObjectPickerDialogGroup.Selection.Items[i].Name;                                                      /// ����� name - �.�. ��� ����� name ���������� ���������� �������� ������
  result.add(recUsGr+'='+JvObjectPickerDialogGroup.Selection.Items[i].ObjectClass);
   ///JvObjectPickerDialogGroup.Selection.Items[i].ObjectClass - ����� ����� ������, ������ ��� ��� ������������
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

function TFormPropertyGroup.AllLocalSystem(NamePC,User,Pass:string):TstringList;
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
    result.Add('������-'+e.Message);
    OleUnInitialize;
    end;
  end;
end;


function TFormPropertyGroup.AllLocalGroup(NamePC,User,Pass:string):TstringList;
var            ///// ��������� ������
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
  nameGr:string;
function typeGroup(gr:string):integer;
const
Tgroup: Array [0..10] of string=('IIS_IUSRS','����������������� ���������'
,'��������� ��������� ����','������������ DCOM','������������ �������� ������������������'
,'������� ������������','������������ ���������� ��������','������������ ���������� �������� �����'
,'����������','�������� ������� �������','');
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
    result.Add('������-'+e.Message);
    OleUnInitialize;
    end;
  end;
end;

function TFormPropertyGroup.LoadAccoutUser (NamePC,User,Pass:string):TstringList;
var                           ///// ��������� ������������
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
    result.Add('������-'+E.Message);
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
var                      ///// �������� ������� �������������  ��� �������� �����
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=LoadAccoutUser(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption);  /// ��������� ��������� ������� ��������� �������������
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
var                      ///// �������� ������� ����� ��� �������� �����
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=AllLocalGroup(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption); // ��������� ����� ���� ������� ��������� �����
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
var                      ///// �������� ������� ����� ��� �������� �����
listUser:Tstringlist;
i:integer;
begin
listUser:=TStringList.Create;
listUser:=AllLocalSystem(LabelNamePC.Caption,labeluser.Caption,labelPasswd.Caption); // ��������� ����� ���� ������� ��������� �����
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
FormGr.Release; /// ����������� �����
end;

procedure TFormPropertyGroup.ButOkClickAccount(Sender: TObject);
var
i:integer;
begin
if AccountDomainAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// ��������� ������������ � ������
then //// ���� �������� ��������� ������� �� ��������� ������������ � ������
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
if GroupDomainAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// ��������� ������ � ������
then //// ���� �������� ��������� ������� �� ��������� ������ � ������
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
if GroupSystemAddGroup(LabelGroup.Caption,ListGr.Text,labelDomain.Caption,labelUser.Caption,LabelPasswd.Caption) /// ��������� ������ � ������
then //// ���� �������� ��������� ������� �� ��������� ��������� ������ � ������
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
listGrup:=readgroupForDomain(LabelNamePC.Caption,CurrentDomainName);  //// ����������� ������ ����� � �������������
if listGrup.Count<>0 then /// ���� ������ �� ����, �.�. ���� �������� ���� ���� ������, ����� �� ���������� �� ����������
Begin
  for I := 0 to listGrup.Count-1 do
  begin
   UsOrGr:='';
   UsOrGr:=listGrup.ValueFromIndex[i];
   if UsOrGr='foreignSecurityPrincipal' then UsOrGr:='user';
   if ListGroupDomainAddGroup
    (listGrup.Names[i],  //// ����������� ������
    frmDomainInfo.LabeledEdit3.Text,    /// ����� ������
    LabelGroup.Caption, /// � ����� ��������� ������ ���������
    labelNamePC.caption,    //// ��������� ����� ��� ���� � ������ �������� ��������� �������� ������
    labeluser.Caption,             //// ������������
    Labelpasswd.Caption,         //// ������
    UsOrGr) /// ������������ ��� ������
    then
       with ListViewAccout.Items.Add do
       begin
       indexImg:=2;                                           /// ���������
       if listGrup.ValueFromIndex[i]='user' then indexImg:=0; /// ������������
       if listGrup.ValueFromIndex[i]='group' then indexImg:=1; /// ������
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

if usr='Account' then  /// ���� ��������� ������������
begin
FormGr.OnShow:=FormGroupShowAccount;
FormGr.Caption:='��������� ������������ - '+namePC;
end;
if usr='Group' then  /// ���� ��������� ������
begin
FormGr.OnShow:=FormGroupShowGroup;
FormGr.Caption:='��������� ������ - '+namePC;
end;

if usr='System' then  /// ���� ��������� ������
begin
FormGr.OnShow:=FormGroupShowSystem;
FormGr.Caption:='��������� ������ - '+namePC;
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
ButOk.Caption:='��';
ButOk.Top:=35;
ButOk.Left:=231;

if usr='Account' then  /// ���� ��������� ������������
ButOk.OnClick:=ButOkClickAccount;
if usr='Group' then    /// ���� ��������� ������
ButOk.OnClick:=ButOkClickGroup;
if usr='System' then    /// ���� ��������� ������
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
while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ������������
 begin
 groupSTR:='';
 groupOrUs:='';
 allstring:='';
 if (FWbemObject.GroupComponent<>null)and(FWbemObject.PartComponent<>null) then
  begin
   if pos(Ansiuppercase('Domain="'+Domain+'",Name="'+NameGroup+'"'),Ansiuppercase(string(FWbemObject.GroupComponent)))<>0 then
     begin
     groupSTR:='';
     groupOrUs:='0'; /// ������ ��� ������������
     UsDom:='';
     if (pos(Ansiuppercase('Win32_UserAccount'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0)
       then groupOrUs:='0' /// ���� ������������ �� 0
     else /// ����� �������� ��� ������
     if (pos(Ansiuppercase('Win32_Group'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0)
      then groupOrUs:='1' ///  ��� ������ - 1
     else
     if (pos(Ansiuppercase('Win32_SystemAccount'),Ansiuppercase(string(FWbemObject.PartComponent)))<>0) then
      groupOrUs:='2'
      else groupOrUs:='3'; ///�����  ��� �����

     UsDom:=(string(FWbemObject.PartComponent)); ///// ����� ������������
     UsDom:=StringReplace(UsDom,'"','',[rfReplaceAll]);
     delete(UsDom,pos(',Name=',UsDom),length(UsDom));
     UsDom:=copy(UsDom,pos('.Domain=',UsDom)+8,length(UsDom));

     groupSTR:=(string(FWbemObject.PartComponent));  //// ��� ������������
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
    result.add('������ ������ ������������� ������ "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

function TFormPropertyGroup.ListGroupDomainAddGroup(StrGroup,domain,LocalGroup,LocalDomain,UserAdmin,AdminPasswd,UserOrGroup:string):bool;
var              ////////////////// ���������� ������ ����� ��� ������
  z:integer;
  InterF : IADsGroup;
  GroupUs: IADsGroup;
  UserUs : IADsUser;

begin
  try
  //ShowMessage('���������  - '+'WinNT://'+localDomain+'/'+LocalGroup+',Group'+UserAdmin+AdminPasswd);
  //ShowMessage('����� - '+'WinNT://'+Domain+'/'+StrGroup+','+UserOrGroup+UserAdmin+AdminPasswd);


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
var              ////////////////// ���������� ������ ������������ � ������
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
var              ////////////////// ���������� NT AUTHORITY - ��������� ��������� ������� ������������ � ��������� ������
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
var              ////////////////// ���������� ������ � ������
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
var              ////////////////// �������� ������������ �� ������
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
var              ////////////////// �������� ������ �� ������
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
var              ////////////////// �������� ��������� ����� � ������������ �� ������
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

  SubItems.Add(copy(listUser.Names[i],2,length(listUser.Names[i]))); /// ������ ������ � �������� ��� ������ ������ ��� ������������
  SubItems.Add(listUser.ValueFromIndex[i]);   ////// ����� ������������
  end;
listUser.Free;
end;

procedure TFormPropertyGroup.SpeedButton1Click(Sender: TObject);
var
i:integer;
begin
try                              ///// ������� ������, ������ ��� ������������
if (ListViewAccout.Items.Count>0) and (ListViewAccout.ItemIndex<>-1) then
Begin
 i:=MessageBox(Self.Handle, PChar('������� ������?')
        , PChar('������ - '+ListViewAccout.Items[ListViewAccout.itemindex].subitems[0] ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;

if ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0]<>'' then
begin
 if (ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=0) or (ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=3) then ///// ���� ������������  ��� foreignsecurityprincipals
   begin
   if AccountDomainDeleteGroup
   (LabelGroup.Caption,  /// �� ����� ������ �������
   ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0],  /// ������ ������������
   LabelNamePC.Caption,                                               /// ����� ������ ��� ��� ��
   ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1],  /// ����� ������ ������� �������
   LabelUser.Caption,                                          /// ������������
   LabelPasswd.Caption)                                        /// ������
   then ListViewAccout.Items[ListViewAccout.ItemIndex].delete;
   exit; /// ������� �.� ����� ���� �������� �� ������ � ����� ���������� ������ ���������� ������, � �� ���� ���..... � �� ����
   end;

 if ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=1 then
   begin
    if GroupDomainDeleteGroup(LabelGroup.Caption, /// �� ����� ������ �������
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0], /// ����� ������ �������
    labeldomain.Caption,                                        /// ����� �������� ������, �� ������� �������
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1], /// �� ������ ������ ��������� ������
    LabelUser.Caption,LabelPasswd.Caption)
   then ListViewAccout.Items[ListViewAccout.ItemIndex].delete;
   exit;
   end;
 if ListViewAccout.Items[ListViewAccout.ItemIndex].ImageIndex=2 then //'\NT AUTHORITY'
   begin
    if GroupSystemDeleteGroup(LabelGroup.Caption, /// �� ����� ������ �������
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[0], /// ����� ������ �������
    labeldomain.Caption,                                        /// ����� �������� ������, �� ������� �������
    ListViewAccout.Items[ListViewAccout.ItemIndex].SubItems[1], /// �� ������ ������ ��������� ������
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
  SubItems.Add(copy(listUser.Names[i],2,length(listUser.Names[i]))); /// ������ ������ � �������� ��� ������ ������ ��� ������������
  SubItems.Add(listUser.ValueFromIndex[i]);   ////// ����� ������������
  end;
listUser.Free;
end;

end.
