unit PrintSecuretyEdit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.ExtCtrls,ActiveX,ComObj, JvExExtCtrls, JvRadioGroup,
  JvComponentBase, JvDSADialogs, JvBaseDlg, JvWinDialogs, JvObjectPickerDialog,
  JvNetscapeSplitter, Vcl.ImgList, System.ImageList;           ///JwaWindows

type
  TPrintSecuretyEditor = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    ListViewUS: TListView;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button8: TButton;
    JvObjectPickerDialog1: TJvObjectPickerDialog;
    GroupBox4: TGroupBox;
    ImageList1: TImageList;
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ListViewUSClick(Sender: TObject);
    procedure uncheckcheckbox;
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    function addItemSDACEDACLArray(Us:string;Dom:string;USSID:string):bool;
    function listviewuseradd(Us:string;Dom:string;USSID:string):bool;
    function ReadPrivilegesSDACEDACLArray(s:string):bool;
    function DeleteAccountSDACEDACLArray(s:string):bool;
    function CheckPrivilegesSDACEDACLArray(s:string;check:boolean;DenyAllow:byte;ValuePriv:integer):bool;
    Function PrinterSecuretyDescryptor:bool;
    function Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
    function Win32ACEDecompose(MyVarArray:variant):bool;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PrintSecuretyEditor: TPrintSecuretyEditor;
  PrintSecurityDescriptor:OLEVariant; /// дескриптор безопастности принтера
  SDACEDACLArray       :variant;      /// DACL из дескриптора
  SDACESACLArray       :Variant;      /// SACL из дескриптора
  SDTrusteeGroup  :Olevariant;        /// Группа владелец объекта
  SDTrusteeOwner  :Olevariant;        /// Владелей объекта
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObject     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  AceType         : integer;
  iValue        : LongWord;
  selectUsReadPr  :bool;
implementation
{$R *.dfm}
uses umain;

function LocalAccountName(FWMIService:OLEVariant;SID:string):string;
var
  FWbemObjectSet      : OLEVariant;
  FWbemObject         :OLEVariant;
  oEnum               : IEnumvariant;
  iValue              : LongWord;
begin
try
OleInitialize(nil);
result:='';
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SID="'+SID+'"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
if oEnum.Next(1, FWbemObject, iValue) = 0 then     //// Пользователь
if (FWbemObject.Name<>null) then
begin
result:=string(FWbemObject.Name);
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
oEnum:=nil;
OleUnInitialize;
except
  on E:Exception do
    begin
    Showmessage('Ошибка чтения имени пользователя "'+E.Message+'"');
    OleUnInitialize;
    end;
  end;
end;

//////////////////////////////////////////
function TPrintSecuretyEditor.addItemSDACEDACLArray(Us:string;Dom:string;USSID:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
begin
  try
  OleInitialize(nil);
  ACETrustee:=FWMIService.get('WIN32_Trustee');
  Win32ACE:= FWMIService.get('WIN32_ACE');
  VarArrayRedim(SDACEDACLArray,VarArrayHighBound(SDACEDACLArray, 1)+1);
  ACETrustee.name:=Us;
  ACETrustee.domain:=Dom;
  ACETrustee.sidstring:=USSID;
  Win32ACE.AccessMask:=131080; /// привелегия печать
  Win32ACE.AceType:=0; /// при добавлении нового пользователя разрешаем доступ
  Win32ACE.AceFlags:=0;
  Win32ACE.Trustee:=ACETrustee;
  SDACEDACLArray[VarArrayHighBound(SDACEDACLArray, 1)]:=Win32ACE;
  result:=true;
  except
    on E:Exception do
    begin
    result:=false;
    memo1.Lines.Add('Ошибка операции добавления нового пользователя: '+E.Message);
    VariantClear(ACETrustee);
    VariantClear(Win32ACE);
    end;
  end;
 VariantClear(ACETrustee);
 VariantClear(Win32ACE);
 OleUnInitialize;
end;

function TPrintSecuretyEditor.listviewuseradd(Us:string;Dom:string;USSID:string):bool;
var
i:integer;
Fuser:bool;
begin
Fuser:=true;
for I := 0 to ListViewUs.Items.Count-1 do
  begin
    if us=ListViewUs.Items[i].Caption then
    begin
     Fuser:=false;
     Break
    end
    else Fuser:=true;
  end;
if Fuser then
  begin
  with listviewus.items.add do
    begin
    ImageIndex:=1;
      caption:=US;
      subitems.add(Dom);
      subitems.add(USSID);
    end;
  result:=true;
  end
  else result:=false;

end;
//////////////////////////////////////////
procedure TPrintSecuretyEditor.uncheckcheckbox;
begin
selectUsReadPr:=true;
CheckBox1.Checked:=false;
CheckBox2.Checked:=false;
CheckBox3.Checked:=false;
CheckBox4.Checked:=false;
CheckBox5.Checked:=false;
CheckBox6.Checked:=false;
selectUsReadPr:=false;
end;
//////////////////////////////////////////////////////////////
procedure TPrintSecuretyEditor.Button8Click(Sender: TObject);
var
i:integer;
Us,Dom,UsSID:string;
begin
try
OleInitialize(nil);
JvObjectPickerDialog1.TargetComputer:=MyPS;
JvObjectPickerDialog1.Scopes[0].DcName:=frmDomainInfo.GetDomainController(CurrentDomainName);
if JvObjectPickerDialog1.Scopes[0].DcName='' then
JvObjectPickerDialog1.Scopes[0].DcName:='localhost';
if JvObjectPickerDialog1.Execute() then
  begin
  Us:='';
  Dom:='';
  UsSID:='';
  if JvObjectPickerDialog1.Selection.Count<>0 then
  for I := 0 to JvObjectPickerDialog1.Selection.Count-1 do
    begin
    if JvObjectPickerDialog1.Selection.Items[i].UPN<>'' then //// если  Учетной запсиси нет, то используем имя, т.к возможно это локальная учетка или группа
        begin
        us:=JvObjectPickerDialog1.Selection.Items[i].UPN;
        Dom:='';
        UsSID:=frmDomainInfo.GetSID(Us);//AdsPathtoSID(JvObjectPickerDialog1.Selection.Items[i].AdsPath);
        end
      else
        begin
        us:= JvObjectPickerDialog1.Selection.Items[i].Name;
        Dom:='';
        UsSID:=frmDomainInfo.GetSID(Us);
        end;
      listviewuseradd(Us,Dom,UsSID);
      addItemSDACEDACLArray(Us,Dom,UsSID);
      memo1.Lines.Add('Name:'+JvObjectPickerDialog1.Selection.Items[i].Name);
      memo1.Lines.Add('AdsPath:'+JvObjectPickerDialog1.Selection.Items[i].AdsPath);
      memo1.Lines.Add('UPN (User Principal Name):'+JvObjectPickerDialog1.Selection.Items[i].UPN);
      memo1.Lines.Add('ObjectClass:'+JvObjectPickerDialog1.Selection.Items[i].ObjectClass);
     end;
  end;
except
    on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add('Ошибка : '+E.Message);
    end;
  end;
OleUnInitialize;
end;

function TPrintSecuretyEditor.Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
var                   /// функция разложение Win32_Trustee, учетные данные пользователей
ACETrustee    : Olevariant;
SIDforStr     : string;
strUS,strDom,strSID:string;
i:integer;
begin
try
SIDforStr:='';
result:=false;
if not(varisnull(VarTrustee)) then
  begin
  OleInitialize(nil);
  ACETrustee:=FWMIService.get('WIN32_Trustee');
  ACETrustee:=VarTrustee;
  if not(varisnull(ACETrustee.TIME_CREATED)) then memo1.Lines.Add(string(ACETrustee.TIME_CREATED));
  if not(varisnull(ACETrustee.Domain)) then strDom:=(string(ACETrustee.Domain))
  else strDom:='';
  if not(varisnull(ACETrustee.name)) then strUS:=(string(ACETrustee.name))
  else strUS:='Unknown'; ///
  if not(varisnull(ACETrustee.SIDString)) then
  begin
   strSID:=(string(ACETrustee.SIDString));
   if (strUS='Unknown')or (strUS='') then strUS:= LocalAccountName(FWMIService,string(ACETrustee.SIDString));
   if strUS='' then strUS:='Unknown';
  end
  else strSID:='';
  memo1.Lines.add(strUS+'/'+strDom+'/'+strSID);
  if strUS<>'' then listviewuseradd(strUS,strDom,strSID);
    if not(varisnull(ACETrustee.SidLength)) then memo1.Lines.Add('Длина SID в байтах:'+string(ACETrustee.SidLength));
  {if varisarray(ACETrustee.sid) then
  begin
    for I := 0 to VarArrayHighBound(ACETrustee.sid, 1) do
    begin
     SIDforStr:=SIDforStr+string(ACETrustee.sid[i])+'-';
    end;
  memo1.Lines.Add('SID Array:'+SIDforStr);
  end;}

  OleUnInitialize;
  result:=true;
  end;
except
    on E:Exception do
    begin
    result:=false;
    memo1.Lines.Add('Ошибка операции Win32TrusteeDecompose: '+E.Message);
    end;
  end;
VariantClear(ACETrustee);
VariantClear(VarTrustee);
end;
//////////////////////////////////////////////////////////////////////////////////////////
function DeleteItemVarArray(a:variant{p,n:integer}):variant; ///удаление элемента массива
var i,z,x:integer;
    v:array of variant;
begin
  x:=0;
 for i:=0 to VarArrayHighBound(a,1) do
 begin
 if not VarIsEmpty(a[i]) then inc(z);
 end;
 SetLength(v,z);
 for i:=0 to VarArrayHighBound(a,1) do
   begin
     if not VarIsEmpty(a[i]) then
       begin
        v[x]:=a[i];
        inc(x);
       end;
   end;
z:=0;
result:=v;
VarClear(a);
///v:=Unassigned;  ///// вопрос очистки вариант массива
end;
///////////////////////////////////////////////////////////////////////////////////////
function TPrintSecuretyEditor.CheckPrivilegesSDACEDACLArray(s:string;check:boolean;DenyAllow:byte;ValuePriv:integer):bool; //// имя,добавть или удалить привелегию,разрешить или запретить доступ,значение привелегии
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
usDomain,usName,usSIDstring:string;
//usSIDArray:     variant;
usSIDLength: integer;

function additemArrayDACL(NameUS:string):bool;
var
i:integer;
begin
  try
  VarArrayRedim(SDACEDACLArray,VarArrayHighBound(SDACEDACLArray, 1)+1);
  ACETrustee.name:=nameUs;
  ACETrustee.domain:=ListViewUS.Items[ListViewUS.ItemFocused.Index].SubItems[0];
  ACETrustee.sidstring:=ListViewUS.Items[ListViewUS.ItemFocused.Index].SubItems[1];
  Win32ACE.AccessMask:=ValuePriv;
  if ValuePriv=983052 then Win32ACE.AceFlags:=0;
  if ValuePriv=131080 then Win32ACE.AceFlags:=0;
  if ValuePriv=983088 then Win32ACE.AceFlags:=9;
  Win32ACE.AceType:=DenyAllow;
  Win32ACE.Trustee:=ACETrustee;
  SDACEDACLArray[VarArrayHighBound(SDACEDACLArray, 1)]:=Win32ACE;
  result:=true;
  except
    on E:Exception do
    begin
    result:=false;
    memo1.Lines.Add('Ошибка операции добавление привелегии: '+E.Message);
    end;
  end;
end;

begin
try
OleInitialize(nil);
ACETrustee:=FWMIService.get('WIN32_Trustee');
Win32ACE:= FWMIService.get('WIN32_ACE');
case check of
true: begin
      if additemArrayDACL(s) then Memo1.Lines.Add('Привелегия успешно добавлена');
      end;
false: begin
        for i := 0 to VarArrayHighBound(SDACEDACLArray, 1) do
        if not VarIsEmpty(SDACEDACLArray[i]) then
         begin
         Win32ACE:=SDACEDACLArray[i];
         ACETrustee:=Win32ACE.Trustee;
         if (string(ACETrustee.name)=s)and (integer(Win32ACE.AccessMask)=ValuePriv) then
          begin
          SDACEDACLArray[i]:=Unassigned;
          end;
        result:=true;
        end;
        SDACEDACLArray:=DeleteItemVarArray(SDACEDACLArray); /// делать очистку массива от пустых элементов}
        Memo1.Lines.Add('Привелегия успешно удалена');
       end;
end;
 except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка операции CheckPrivilegesSDACEDACLArray: '+E.Message);
    memo1.Lines.Add('---------------------------');
    end;
   end;
 VariantClear(ACETrustee);
 VariantClear(Win32ACE);
 OleUnInitialize;
 end;
//////////////////////////////////////////////////////////////////////////////////////
function TPrintSecuretyEditor.DeleteAccountSDACEDACLArray(s:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
//usSIDArray:     variant;
begin
try
result:=false;
OleInitialize(nil);
ACETrustee:=FWMIService.get('WIN32_Trustee');
Win32ACE:= FWMIService.get('WIN32_ACE');
for i := 0 to VarArrayHighBound(SDACEDACLArray, 1) do
  begin
  if not VarIsEmpty(SDACEDACLArray[i]) then
   begin
      Win32ACE:=SDACEDACLArray[i];
      ACETrustee:=Win32ACE.Trustee;
      if string(ACETrustee.name)=s then
        begin
        SDACEDACLArray[i]:=Unassigned;
        end;
    end;
  result:=true;
  end;
 SDACEDACLArray:=DeleteItemVarArray(SDACEDACLArray); /// делать очистку массива от пустых элементов}
 VariantClear(ACETrustee);
 VariantClear(Win32ACE);
 OleUnInitialize;
 except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка функции DeleteAccountSDACEDACLArray: '+E.Message);
    memo1.Lines.Add('---------------------------');
    VariantClear(ACETrustee);
    VariantClear(Win32ACE);
    OleUnInitialize;
    end;
   end;
 end;
//////////////////////////////////////////////////////////////////////////////////
//разрешить
//печать - маска 131080/ доступ 0/ наследование прав 0
//Управление этим принтером+ печать - маска 983052/доступ 0/ наследование прав 0
//Управение документами - маска 131072/983088/доступ 0 0 /наследование прав 10/9
//Печать+управление этим принтером+управление документами- маска 983052/983088/доступ 0 0 / наследование прав 0 9

//запретить
//печать - маска 131080/ доступ 1/ наследование прав 0
//Управление этим принтером+ печать - маска 983052/доступ 1/ наследование прав 0
//Управение документами - маска 983088/доступ 1 /наследование прав 9
//Печать+управление документами - маска 131080/983088/доступ 1 1 / наследование прав 0 9
//Печать+управление этим принтером+управление документами- маска 983052/983088/доступ 1 1 / наследование прав 0 9

function TPrintSecuretyEditor.ReadPrivilegesSDACEDACLArray(s:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
begin
try
selectUsReadPr:=true;
OleInitialize(nil);
ACETrustee:=FWMIService.get('WIN32_Trustee');
Win32ACE:= FWMIService.get('WIN32_ACE');
for i := 0 to VarArrayHighBound(SDACEDACLArray, 1) do
  begin
   if not VarIsEmpty(SDACEDACLArray[i]) then  //если занчание не  Unassigned
    begin
    Win32ACE:=SDACEDACLArray[i];
    ACETrustee:=Win32ACE.Trustee;
    if ACETrustee.name=null then ACETrustee.name:='Unknown';
    if string(ACETrustee.name)=s then
     begin
        case integer(Win32ACE.AceType) of    /// 0 - разрешено, 1- запрещено , 2 - аудит
        0:
           begin
             case integer(Win32ACE.AccessMask) of
             131080: CheckBox1.Checked:=true; //// печать
             //131072: CheckBox5.Checked:=true; //// Управение документами - маска 131072/983088
             983052: begin CheckBox1.Checked:=true;CheckBox3.Checked:=true; end; //Управление этим принтером+ печать
             983088: CheckBox5.Checked:=true;    // управление документами
             end;
           end;
        1:
          begin
             case integer(Win32ACE.AccessMask) of
             131080: CheckBox2.Checked:=true; //// печать
             983052: begin CheckBox2.Checked:=true;CheckBox4.Checked:=true; end; //Управление этим принтером+ печать
             983088:CheckBox6.Checked:=true;  // управление документами
             end;
          end;
      end;
     // if not(VarIsNull(Win32ACE.TIME_CREATED)) then
     // Memo1.Lines.Add('Время создания:'+string(Win32ACE.TIME_CREATED));
      if not(VarIsNull(Win32ACE.AccessMask)) then
      Memo1.Lines.Add('Маска доступа (сумма прав/AccessMask):'+string(Win32ACE.AccessMask));
      if not(VarIsNull(Win32ACE.AceFlags)) then
      Memo1.Lines.Add('Наследование Прав(AceFlags):'+string(Win32ACE.AceFlags));
      if not(VarIsNull(Win32ACE.AceType)) then
      Memo1.Lines.Add('Доступ 0/1/2:'+string(Win32ACE.AceType));
     result:=true;
     end;
  end;
  end;
except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка операции ReadPrivilegesSDACEDACLArray: '+E.Message);
    memo1.Lines.Add('---------------------------');
    OleUnInitialize;
    end;
  end;
VariantClear(Win32ACE);
VariantClear(ACETrustee);
s:='';
OleUnInitialize;
selectUsReadPr:=false;
end;
/////////////////////////////////////////////////////////////////////////////////
function TPrintSecuretyEditor.Win32ACEDecompose(MyVarArray:variant):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
begin
result:=false;
OleInitialize(nil);
ACETrustee:=FWMIService.get('WIN32_Trustee');
Win32ACE:= FWMIService.get('WIN32_ACE');
for i := 0 to VarArrayHighBound(MyVarArray, 1) do
begin
try
  memo1.Lines.Add('_______________________________');
  Win32ACE:=MyvarArray[i];
  ACETrustee:=Win32ACE.Trustee;
  Win32TrusteeDecompose(ACETrustee);
  //if not(VarIsNull(Win32ACE.TIME_CREATED)) then
  //Memo1.Lines.Add('Время создания:'+string(Win32ACE.TIME_CREATED));
  if not(VarIsNull(Win32ACE.AccessMask)) then
  Memo1.Lines.Add('AccessMask:'+string(Win32ACE.AccessMask));
  if not(VarIsNull(Win32ACE.AceFlags)) then
  Memo1.Lines.Add('AceFlags:'+string(Win32ACE.AceFlags));
  if not(VarIsNull(Win32ACE.AceType)) then
  Memo1.Lines.Add('Доступ 0/1/2:'+string(Win32ACE.AceType));
  if not(VarIsNull(Win32ACE.GuidInheritedObjectType)) then
  Memo1.Lines.Add('Guid родителя:'+string(Win32ACE.GuidInheritedObjectType));
  if not(VarIsNull(Win32ACE.GuidObjectType)) then
  Memo1.Lines.Add('Guid объекта:'+string(Win32ACE.GuidObjectType));
  memo1.Lines.Add('_______________________________');
  VariantClear(ACETrustee);
  VariantClear(Win32ACE);
  Result:=true;
except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка операции Win32ACEDecompose: '+E.Message);
    memo1.Lines.Add('---------------------------');
    OleUnInitialize;
    end;
  end;
end;
VarClear(MyVarArray);
VariantClear(Win32ACE);
VariantClear(ACETrustee);
OleUnInitialize;
end;
//////////////////////////////////////////////////////////////////////////////////////


procedure TPrintSecuretyEditor.FormShow(Sender: TObject);
begin
PrintSecuretyEditor.Width:=366;
PrintSecuretyEditor.Height:=424;
memo1.Clear;
Listviewus.Clear;  //очистка
uncheckcheckbox; /// очистка чекбоксов
PrinterSecuretyDescryptor;  /// чтение дескриптора и вывод пользователей и прав
PrintSecuretyEditor.Caption:='Безопасность для '+SelectedPrint;
GroupBox2.Caption:='Разрешения: ';
end;

procedure TPrintSecuretyEditor.ListViewUSClick(Sender: TObject);
begin
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
begin
uncheckcheckbox; /// очистка чекбоксов
ReadPrivilegesSDACEDACLArray(listviewus.Items[listviewus.ItemFocused.Index].Caption); //// чтение привелегий из массива
GroupBox2.Caption:='Разрешения: '+listviewus.Items[listviewus.ItemFocused.Index].Caption;
end;
end;

Function TPrintSecuretyEditor.PrinterSecuretyDescryptor:bool;
var
  oEnum : IEnumvariant;
  SDescriptorError:integer;
begin
try
begin
memo1.clear;
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
PrintSecurityDescriptor:=FWMIService.get('Win32_SecurityDescriptor');
SDTrusteeGroup:= FWMIService.get('Win32_Trustee');
SDTrusteeOwner:= FWMIService.get('Win32_Trustee');
     if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
      SDescriptorError:=FWbemObject.GetSecurityDescriptor(PrintSecurityDescriptor);
      if PrintSecurityDescriptor.ControlFlags<>null then memo1.Lines.Add('ControlFlags:'+string(PrintSecurityDescriptor.ControlFlags))
      else memo1.Lines.add('ControlFlags is null');

      if VarIsArray(PrintSecurityDescriptor.dacl) then  //// в массива dacl права пользователей
      begin
      SDACEDACLArray:=PrintSecurityDescriptor.dacl;
      if Win32ACEDecompose(SDACEDACLArray)then memo1.Lines.add('Win32ACEDecompose(SDACEDACLArray)');
      // Memo1.Lines.add('Размерность массива SDACEDACLArray - '+ inttostr(VarArrayDimCount(SDACEDACLArray)));
      end
        else Memo1.Lines.Add('PrintSecurityDescriptor.dacl not Array');

      if VarIsArray(PrintSecurityDescriptor.sacl) then ////АУДИТ. Массив попыток доступа к объекту для каждого поьзователя или группы
      begin
      SDACESACLArray:=PrintSecurityDescriptor.sacl;
      if Win32ACEDecompose(SDACESACLArray)then memo1.Lines.add('Win32ACEDecompose(SDACESACLArray)');
      // Memo1.Lines.add('Размерность массива SDACESACLArray - '+ inttostr(VarArrayDimCount(SDACESACLArray)));
      end
        else Memo1.Lines.Add('PrintSecurityDescriptor.sacl not Array');

      SDTrusteeGroup:=PrintSecurityDescriptor.group;
      memo1.Lines.add('Группа владелец объекта');
      if Win32TrusteeDecompose(SDTrusteeGroup) then Memo1.Lines.Add('Выгружены данные SDTrusteeGroup');

      SDTrusteeOwner:=PrintSecurityDescriptor.owner;
      memo1.Lines.add('Владелец объекта');
      if Win32TrusteeDecompose(SDTrusteeOwner) then Memo1.Lines.Add('Выгружены данные SDTrusteeOwner');
      end;
   memo1.Lines.Add('Результат операции PrinterSecureDescrypt: '+  SysErrorMessage(SDescriptorError));

end;
  except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка операции PrinterSecureDescrypt: '+  SysErrorMessage(SDescriptorError));
    memo1.Lines.Add('Ошибка операции PrintSecurityDescriptor: '+E.Message);
    memo1.Lines.Add('---------------------------');
    OleUnInitialize;
    end;
  end;
If oEnum<>nil then oEnum:=nil;
OleUnInitialize;
end;


procedure TPrintSecuretyEditor.Button1Click(Sender: TObject);
begin
Button5.Click;
Button2.click;
end;

procedure TPrintSecuretyEditor.Button2Click(Sender: TObject);   //// закрыть окно
begin
 OleInitialize(nil);
 VariantClear(PrintSecurityDescriptor);
 VariantClear(SDTrusteeOwner);
 VariantClear(SDTrusteeGroup);
 VarClear(SDACESACLArray);
 VarClear(SDACEDACLArray);
 VariantClear(FWbemObject);
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize;
close;
end;

procedure TPrintSecuretyEditor.Button3Click(Sender: TObject);
begin
if DeleteAccountSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption) then
 ListViewUS.Items[ListViewUS.ItemFocused.Index].Delete;
end;

procedure TPrintSecuretyEditor.Button4Click(Sender: TObject);
begin
Listviewus.Clear;  //очистка
uncheckcheckbox; /// очистка чекбоксов
PrinterSecuretyDescryptor;  /// чтение дескриптора и вывод пользователей и прав
end;

procedure TPrintSecuretyEditor.Button5Click(Sender: TObject);
var                                                          //// сохранить дескриптор
SDescriptorError:integer;
SDTrusteeGroup:OleVariant;
begin
try
OleInitialize(nil);
PrintSecurityDescriptor.ControlFlags:=4;
PrintSecurityDescriptor.group:=null;
PrintSecurityDescriptor.Owner:=null;
PrintSecurityDescriptor.dacl:=SDACEDACLArray;
SDescriptorError:=FWbemObject.SetSecurityDescriptor(PrintSecurityDescriptor);
except
on E:Exception do
         begin
           memo1.Lines.Add('результат операции сохранения: '+  SysErrorMessage(SDescriptorError));
           memo1.Lines.Add('Ошибка операции SetSecurityDescriptor: '+E.Message);
           memo1.Lines.Add('---------------------------');
         end;

end;
memo1.Lines.Add('результат операции сохранения: '+  SysErrorMessage(SDescriptorError));
if SysErrorMessage(SDescriptorError)='' then showmessage('Упс.. что-то пошло не так ¯\_(ツ)_/¯');
OleUnInitialize;
end;

procedure TPrintSecuretyEditor.Button6Click(Sender: TObject);
begin
Win32ACEDecompose(SDACEDACLArray);
end;



procedure TPrintSecuretyEditor.CheckBox1Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
    if CheckBox1.Checked then
    begin
     CheckBox2.Checked:=false;
     CheckBox4.Checked:=false;
    end;
  AceType:=0; // разрешить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox1.Checked,AceType,131080);
  end;
end;

procedure TPrintSecuretyEditor.CheckBox2Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox2.Checked then
    begin
    CheckBox1.Checked:=false;
    CheckBox3.Checked:=false;
    end;
  AceType:=1;  //запретить
  if not CheckBox4.Checked then
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox2.Checked,AceType,131080);
  end;
end;

procedure TPrintSecuretyEditor.CheckBox3Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox3.Checked then
    begin
     CheckBox1.Checked:=true;
     CheckBox4.Checked:=false;
     CheckBox2.Checked:=false;
    end;
  AceType:=0;  // разрешить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox3.Checked,AceType,983052);
  if CheckBox3.Checked then CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,false,AceType, 131080);
  end;
end;

procedure TPrintSecuretyEditor.CheckBox4Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox4.Checked then
    begin
    CheckBox2.Checked:=true;
    CheckBox1.Checked:=false;
    CheckBox3.Checked:=false;
    end;
  AceType:=1; //запретить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox4.Checked,AceType, 983052);
  if CheckBox4.Checked then CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,false,AceType, 131080); /// удалить привелегию
  end;
end;

procedure TPrintSecuretyEditor.CheckBox5Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox5.Checked then CheckBox6.Checked:=false;
  AceType:=0;  // разрешить
  //CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox5.Checked,AceType, 131072);
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox5.Checked,AceType, 983088);
  end;
end;

procedure TPrintSecuretyEditor.CheckBox6Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox6.Checked then CheckBox5.Checked:=false;
  AceType:=1; //запретить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox6.Checked,AceType, 983088);
  end;
end;

procedure TPrintSecuretyEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
OleInitialize(nil);
VarClear(SDACEDACLArray);
 VarClear(SDACESACLArray);
 VariantClear(SDTrusteeGroup);
 VariantClear(SDTrusteeOwner);
 VariantClear(PrintSecurityDescriptor);
  VariantClear(FWbemObject);
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize;
end;

procedure TPrintSecuretyEditor.FormHide(Sender: TObject);
begin
OleInitialize(nil);
VarClear(SDACEDACLArray);
 VarClear(SDACESACLArray);
 VariantClear(SDTrusteeGroup);
 VariantClear(SDTrusteeOwner);
 VariantClear(PrintSecurityDescriptor);
  VariantClear(FWbemObject);
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize
end;
procedure TPrintSecuretyEditor.FormResize(Sender: TObject);
begin
PrintSecuretyEditor.Width:=366;
end;

end.
