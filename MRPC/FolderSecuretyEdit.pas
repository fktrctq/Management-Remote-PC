unit FolderSecuretyEdit;

interface

uses
   Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.ExtCtrls,ActiveX,ComObj, JvExExtCtrls, JvRadioGroup,
  JvComponentBase, JvDSADialogs, JvBaseDlg, JvWinDialogs, JvObjectPickerDialog,
  JvNetscapeSplitter, Vcl.ImgList, System.ImageList;           ///JwaWindows

type
  TFrmSecFolder = class(TForm)
    GroupBox1: TGroupBox;
    ListViewUS: TListView;
    Button3: TButton;
    Button8: TButton;
    GroupBox2: TGroupBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    GroupBox3: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    Button5: TButton;
    GroupBox4: TGroupBox;
    Memo1: TMemo;
    JvObjectPickerDialog1: TJvObjectPickerDialog;
    ImageList1: TImageList;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormHide(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ListViewUSClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    procedure uncheckcheckbox;
    function addItemSDACEDACLArray(Us:string;Dom:string;USSID:string):bool;
    function listviewuseradd(Us:string;Dom:string;USSID:string):bool;
    function ReadPrivilegesSDACEDACLArray(s:string):bool;
    function DeleteAccountSDACEDACLArray(s:string):bool;
    function CheckPrivilegesSDACEDACLArray(s:string;check:boolean;DenyAllow:byte;ValuePriv:integer):bool;
    Function PrinterSecuretyDescryptor:bool;
    function Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
    function Win32ACEDecompose(MyVarArray:variant):bool;

  public
    { Public declarations }
  end;

var
  FrmSecFolder: TFrmSecFolder;
  SDACEDACLArray       :variant;      /// DACL из дескриптора
  SDACESACLArray       :Variant;      /// SACL из дескриптора
  AceType         : integer;
  iValue        : LongWord;
  selectUsReadPr  :bool;
  FolderSecurityDescriptor:OLEVariant;
  Win32ACEdacl      : OleVariant;
  Win32ACEsacl      : OleVariant;
  FSWbemLocatorSC   : OLEVariant;
  FWMIServiceSC     : OLEVariant;
implementation
{$R *.dfm}
uses FormFSShare,umain;

function CheckArrayForEmpty(s:Variant):bool;  /// проверка Variant массива на пустоту
begin
try
 if VarArrayHighBound(s, 1)<>0 then result:=true
 else result:=false;
except
   result:=false;
end;
end;

function TFrmSecFolder.addItemSDACEDACLArray(Us:string;Dom:string;USSID:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
step:string;
begin
  try
  step:='0';
  OleInitialize(nil);
  ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
  Win32ACE:= FWMIServiceSC.get('WIN32_ACE');
  step:='1';
  VarArrayRedim(SDACEDACLArray,VarArrayHighBound(SDACEDACLArray, 1)+1);
  step:='1.5';
  ACETrustee.name:=Us;
  ACETrustee.domain:=Dom;
  ACETrustee.sidstring:=USSID;
  step:='2';
  Win32ACE.AccessMask:=1179817; /// привелегия чтение по умолчанию у добавленых пользователей
  Win32ACE.AceType:=0; /// при добавлении нового пользователя разрешаем доступ
  Win32ACE.AceFlags:=0;
  Win32ACE.Trustee:=ACETrustee;
  step:='3';
  SDACEDACLArray[VarArrayHighBound(SDACEDACLArray, 1)]:=Win32ACE;
  result:=true;
  except
    on E:Exception do
    begin
    result:=false;
    memo1.Lines.Add('Ошибка операции добавления нового пользователя: step - '+step+' /'+E.Message);
    VariantClear(ACETrustee);
    VariantClear(Win32ACE);
    end;
  end;
 VariantClear(ACETrustee);
 VariantClear(Win32ACE);
 OleUnInitialize;
end;

function TFrmSecFolder.listviewuseradd(Us:string;Dom:string;USSID:string):bool;
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
procedure TFrmSecFolder.uncheckcheckbox;
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
procedure TFrmSecFolder.Button8Click(Sender: TObject);
var
i:integer;
Us,Dom,UsSID:string;
begin
try
OleInitialize(nil);
JvObjectPickerDialog1.TargetComputer:=MyPS;
JvObjectPickerDialog1.Scopes[0].DcName:=frmDomainInfo.GetDomainController(CurrentDomainName);
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

function TFrmSecFolder.Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
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
  ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
  ACETrustee:=VarTrustee;
  if not(varisnull(ACETrustee.TIME_CREATED)) then memo1.Lines.Add(string(ACETrustee.TIME_CREATED));
  if not(varisnull(ACETrustee.Domain)) then strDom:=(string(ACETrustee.Domain))
  else strDom:='';
  if not(varisnull(ACETrustee.name)) then strUS:=(string(ACETrustee.name))
  else strUS:='Unknown';
  if not(varisnull(ACETrustee.SIDString)) then strSID:=(string(ACETrustee.SIDString))
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
v:=Unassigned;
end;
///////////////////////////////////////////////////////////////////////////////////////

function TFrmSecFolder.CheckPrivilegesSDACEDACLArray(s:string;check:boolean;DenyAllow:byte;ValuePriv:integer):bool; //// имя,добавть или удалить привелегию,разрешить или запретить доступ,значение привелегии
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
ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
Win32ACE:= FWMIServiceSC.get('WIN32_ACE');
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
function TFrmSecFolder.DeleteAccountSDACEDACLArray(s:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
//usSIDArray:     variant;
begin
try
result:=false;
OleInitialize(nil);
ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
Win32ACE:= FWMIServiceSC.get('WIN32_ACE');
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
function TFrmSecFolder.ReadPrivilegesSDACEDACLArray(s:string):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
step:string;
begin
try
step:='1';
selectUsReadPr:=true;
OleInitialize(nil);
step:='2';
ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
Win32ACE:= FWMIServiceSC.get('WIN32_ACE');
step:='3';
memo1.Lines.Add('Длинна массива - '+inttostr(VarArrayHighBound(SDACEDACLArray, 1)));
step:='4';
for i := 0 to VarArrayHighBound(SDACEDACLArray, 1) do
  begin
  step:='5';
   if not VarIsEmpty(SDACEDACLArray[i]) then  //если занчание не  Unassigned
    begin
    step:='6';
    Win32ACE:=SDACEDACLArray[i];
    ACETrustee:=Win32ACE.Trustee;
    if ACETrustee.name=null then ACETrustee.name:='Unknown';
    if string(ACETrustee.name)=s then
     begin
        case integer(Win32ACE.AceType) of    /// 0 - разрешено, 1- запрещено , 2 - аудит
        0:
           begin
             case integer(Win32ACE.AccessMask) of
             2032127: CheckBox1.Checked:=true; //// полный доступ
             1507839: begin  CheckBox3.Checked:=true;  end; //// изменение
             1179817: CheckBox5.Checked:=true;    // чтение
             end;
           end;
        1:
          begin
             case integer(Win32ACE.AccessMask) of
             2032127: CheckBox2.Checked:=true; //// полный доступ
             1507839: begin CheckBox4.Checked:=true; end; //изменение
             1179817:CheckBox6.Checked:=true;  // чтение
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
    memo1.Lines.Add('Ошибка операции ReadPrivilegesSDACEDACLArray: step -  '+step+'/'+E.Message);
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
function TFrmSecFolder.Win32ACEDecompose(MyVarArray:variant):bool;
var
i:integer;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
begin
result:=false;
OleInitialize(nil);
ACETrustee:=FWMIServiceSC.get('WIN32_Trustee');
Win32ACE:= FWMIServiceSC.get('WIN32_ACE');
for i := 0 to VarArrayHighBound(MyVarArray, 1) do
begin
try
  memo1.Lines.Add('_______________________________');
  Win32ACE:=MyvarArray[i];
  ACETrustee:=Win32ACE.Trustee;
  Win32TrusteeDecompose(ACETrustee);
 // if not(VarIsNull(Win32ACE.TIME_CREATED)) then
//  Memo1.Lines.Add('Время создания:'+string(Win32ACE.TIME_CREATED));
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


procedure TFrmSecFolder.FormShow(Sender: TObject);
begin
FrmSecFolder.Height:=421;
memo1.Clear;
Listviewus.Clear;  //очистка
uncheckcheckbox; /// очистка чекбоксов
PrinterSecuretyDescryptor;  /// создание дескриптора и вывод пользователей и прав
GroupBox2.Caption:='Права доступа: ';
end;


procedure TFrmSecFolder.ListViewUSClick(Sender: TObject);
begin
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
begin
uncheckcheckbox; /// очистка чекбоксов
ReadPrivilegesSDACEDACLArray(listviewus.Items[listviewus.ItemFocused.Index].Caption); //// чтение привелегий из массива
GroupBox2.Caption:='Права доступа: '+listviewus.Items[listviewus.ItemFocused.Index].Caption;
end;
end;

Function TFrmSecFolder.PrinterSecuretyDescryptor:bool;
var
STEP:integer;
begin
try
begin
if not CheckArrayForEmpty(SDACEDACLArray) then /// если массив не пустой, т.е. он может уже быть созданным
  begin
  FolderSecurityDescriptor:=FWMIServiceSC.get('Win32_SecurityDescriptor');
  SDACEDACLArray:=VarArrayCreate([0,0],varVariant);
  SDACESACLArray:=VarArrayCreate([0,0],varVariant);
  end
else // иначе если уже редактированли то в массиве есть пользователи, надо их прочитать
  begin
  if Win32ACEDecompose(SDACEDACLArray) then
  memo1.Lines.Add('Загрузка пользователей прошла успешно');
  end;
end;
  except
  on E:Exception do
    begin
    memo1.Lines.Add('Ошибка операции PrinterSecureDescrypt: '+E.Message);
    OleUnInitialize;
    end;
  end;
end;


procedure TFrmSecFolder.Button1Click(Sender: TObject);
begin
Button5.Click;
Button2.click;
end;


procedure TFrmSecFolder.Button2Click(Sender: TObject);   //// закрыть окно
begin
close;
end;


procedure TFrmSecFolder.Button3Click(Sender: TObject);
begin
if DeleteAccountSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption) then
 ListViewUS.Items[ListViewUS.ItemFocused.Index].Delete;
end;


procedure TFrmSecFolder.Button4Click(Sender: TObject);
begin
Listviewus.Clear;  //очистка
uncheckcheckbox; /// очистка чекбоксов
PrinterSecuretyDescryptor;  /// чтение дескриптора и вывод пользователей и прав
end;

procedure TFrmSecFolder.Button5Click(Sender: TObject);
begin
try
OleInitialize(nil);
FolderSecurityDescriptor.ControlFlags:=4;
FolderSecurityDescriptor.group:=null;
FolderSecurityDescriptor.Owner:=null;
FolderSecurityDescriptor.dacl:=SDACEDACLArray;
except
on E:Exception do
 begin
 memo1.Lines.Add('Ошибка операции FolderSecurityDescriptor: '+E.Message);
 memo1.Lines.Add('¯\_(ツ)_/¯');
 end;
end;
OleUnInitialize;
end;




procedure TFrmSecFolder.Button6Click(Sender: TObject);
begin
Win32ACEDecompose(SDACEDACLArray);
end;



procedure TFrmSecFolder.CheckBox1Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
    if CheckBox1.Checked then
    begin
     if CheckBox2.Checked then CheckBox2.Checked:=false;
     if CheckBox4.Checked then CheckBox4.Checked:=false;
     if CheckBox6.Checked then CheckBox6.Checked:=false;
     if not CheckBox3.Checked then CheckBox3.Checked:=true;
     if not CheckBox5.Checked then  CheckBox5.Checked:=true;
    end;
  AceType:=0; // разрешить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox1.Checked,AceType,2032127);
  end;
end;



procedure TFrmSecFolder.CheckBox2Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox2.Checked then
    begin
    if CheckBox5.Checked then CheckBox5.Checked:=false;
    if CheckBox3.Checked then CheckBox3.Checked:=false;
    if CheckBox1.Checked then CheckBox1.Checked:=false;
    if not CheckBox4.Checked then CheckBox4.Checked:=true;
    if not CheckBox6.Checked then CheckBox6.Checked:=true;
    end;
  AceType:=1;  //запретить
 // if not CheckBox4.Checked then
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox2.Checked,AceType,2032127);
  end;
end;



procedure TFrmSecFolder.CheckBox3Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox3.Checked then
    begin
     if CheckBox2.Checked then CheckBox2.Checked:=false;
     if CheckBox4.Checked then CheckBox4.Checked:=false;
     if CheckBox6.Checked then CheckBox6.Checked:=false;
     if not CheckBox5.Checked then CheckBox5.Checked:=true;
    end;
   if not CheckBox3.Checked then
   begin
    if CheckBox1.Checked then CheckBox1.Checked:=false;
   end;
  AceType:=0;  // разрешить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox3.Checked,AceType,1507839); /// изменения
  end;
end;



procedure TFrmSecFolder.CheckBox4Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox4.Checked then
    begin
    if not CheckBox6.Checked then CheckBox6.Checked:=true;
    if CheckBox5.Checked then CheckBox5.Checked:=false;
    if CheckBox3.Checked then CheckBox3.Checked:=false;
    if CheckBox1.Checked then CheckBox1.Checked:=false;
    end;
    if not CheckBox4.Checked then
    begin
    if CheckBox2.Checked then CheckBox2.Checked:=false;
    end;
  AceType:=1; //запретить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox4.Checked,AceType, 1507839); /// изменения
  end;
end;


procedure TFrmSecFolder.CheckBox5Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox5.Checked then
  begin
   if CheckBox2.Checked then CheckBox2.Checked:=false;
   if CheckBox4.Checked then CheckBox4.Checked:=false;
   if CheckBox6.Checked then CheckBox6.Checked:=false;
   end;
  if not CheckBox5.Checked then
  begin
  if CheckBox1.Checked then CheckBox1.Checked:=false;
  if CheckBox3.Checked then CheckBox3.Checked:=false;
  end;

   AceType:=0;  // разрешить
  //CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox5.Checked,AceType, 131072);
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox5.Checked,AceType, 1179817); /// чтение
  end;
end;



procedure TFrmSecFolder.CheckBox6Click(Sender: TObject);
begin
if not selectUsReadPr then
if (ListViewUs.Items.Count<>0) and (ListViewUs.SelCount<>0) then
  begin
  if CheckBox6.Checked then
  begin
  if CheckBox5.Checked then CheckBox5.Checked:=false;
  if CheckBox3.Checked then CheckBox3.Checked:=false;
  if CheckBox1.Checked then CheckBox1.Checked:=false;
  end;
  if not CheckBox6.Checked then
  begin
  if CheckBox2.Checked then CheckBox2.Checked:=false;
  if CheckBox4.Checked then CheckBox4.Checked:=false;
  end;
  AceType:=1; //запретить
  CheckPrivilegesSDACEDACLArray(ListViewUS.Items[ListViewUS.ItemFocused.Index].Caption,CheckBox6.Checked,AceType, 1179817);
  end;
end;




procedure TFrmSecFolder.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
{OleInitialize(nil);
VariantClear(FWbemObjectSetSC);
VariantClear(FWMIServiceSC);
VariantClear(FSWbemLocatorSC);
OleUnInitialize; }
end;

procedure TFrmSecFolder.FormHide(Sender: TObject);
begin
{OleInitialize(nil);
VariantClear(FWbemObjectSetSC);
VariantClear(FWMIServiceSC);
VariantClear(FSWbemLocatorSC);
OleUnInitialize}
end;






procedure TFrmSecFolder.FormResize(Sender: TObject);
begin
FrmSecFolder.Width:=365;
end;

end.
