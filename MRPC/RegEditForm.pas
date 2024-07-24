unit RegEditForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.ImgList, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Buttons, JvExExtCtrls, JvNetscapeSplitter,
  Inifiles, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.Comp.Client;

type
  TRegEdit = class(TForm)
    JvNetscapeSplitter2: TJvNetscapeSplitter;
    Panel1: TPanel;
    Splitter1: TSplitter;
    Panel3: TPanel;
    LVKey: TListView;
    Panel4: TPanel;
    SpeedButton1: TSpeedButton;
    Panel6: TPanel;
    EditPath: TEdit;
    Panel5: TPanel;
    TreeViewFolders: TTreeView;
    Panel2: TPanel;
    ButtonUpdate: TButton;
    ComboBox2: TComboBoxEx;
    EditUser: TEdit;
    EditPass: TEdit;
    MemoLog: TMemo;
    ImageList1: TImageList;
    ImageList2: TImageList;
    ImageListPC: TImageList;
    StatusBar1: TStatusBar;
    PopupKey: TPopupMenu;
    C1: TMenuItem;
    PopupKeyValue: TPopupMenu;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N6: TMenuItem;
    REGBINARY1: TMenuItem;
    REGDWORD1: TMenuItem;
    REGQWORD1: TMenuItem;
    REGMULTISZ1: TMenuItem;
    N7: TMenuItem;
    EditRooT: TComboBox;
    N5: TMenuItem;
    N1: TMenuItem;
    N8: TMenuItem;
    MainMenu1: TMainMenu;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    REGSZ1: TMenuItem;
    REGBINARY2: TMenuItem;
    REGDWORD2: TMenuItem;
    REGQWORD2: TMenuItem;
    REGMULTISZ2: TMenuItem;
    REGEXPANDSZ1: TMenuItem;
    N25: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure TreeViewFoldersChange(Sender: TObject; Node: TTreeNode);
    procedure clearTTreeView;
    procedure ButtonUpdateClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);  // удалить форму при закрытии, редактирования ключей реестра
    procedure FormSaveKeyClose(Sender: TObject); // закрыть форму для сохраненных ключей
    procedure ButtonNoClose(Sender: TObject);
    procedure REGBINARY1Click(Sender: TObject);
    procedure REGDWORD1Click(Sender: TObject);
    procedure REGQWORD1Click(Sender: TObject);
    procedure REGMULTISZ1Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    //////////////////////////процедуры динамической формы/////////////////////////
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonOkREG_SZ(Sender: TObject); // сохраняет изменения в строковом параметре
    procedure ButtonOkREG_Binary(Sender: TObject); // сохраняет изменения в binary параметре
    procedure ButtonOkREG_QWORD(Sender: TObject); // сохраняет изменения в QWORD параметре
    procedure ButtonOkREG_DWORD(Sender: TObject); // сохраняет изменения в DWORD параметре
    procedure ButtonOkREG_EXPAND_SZ(Sender: TObject); // сохраняет изменения в расширенном строковом параметре
    procedure ButtonOkMULTI_SZ(Sender: TObject); // сохраняет изменения мультистрокового параметре
    procedure LVKeyDblClick(Sender: TObject);
    procedure FormKeyShowBinary(Sender: TObject);
    procedure FormKeyShowMULTI_SZ(Sender: TObject); // дествия для REG_MULTI_SZ при показк формы для редактирования
    procedure RadioGrClick(Sender: TObject);
    procedure RadioGrQDClick(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure EditPathKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N5Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N25Click(Sender: TObject);
 private
    FSWbemLocatorRE : OLEVariant;
    FWMIServiceRE   : OLEVariant;
    FormKey:Tform;
    MemoKey:Tmemo;
    EditKey:TlabeledEdit;
    EditNameKey:TlabeledEdit;
    ButtonOk,ButtonNo:Tbutton;
    RadioGr:TradioGroup;
    FormSaveKey:Tform;
    PanelSaveKey:Tpanel;
    LVSaveKey:TlistView;
    ButKeyDel:Tbutton;
    ButKeyClose:Tbutton;
    ButKeyImport:Tbutton;
    FLV:TlistView;
    FLabel:TLabel;
    CopyPasteVariant:Variant;
    CurPCConnectRegEdit:string;
    FavoriteTXT:TstringList;
    function ConnectRemote:boolean;
    function ConnectWMI : Boolean;
    function JumpFolderTreeItem(Node: TTreeNode;hDefKey:integer;sSubKeyName:string):TTreeNode;
    function UpdateFolderTreeItem(Node: TTreeNode;hDefKey:integer;sSubKeyName:string):boolean; // чтение директорий
    function UpdateKeyTreeItem(hDefKey:integer;sSubKeyName:string):boolean; // чтение ключий и их значений
    function GetNodeFullPath(Node: TTreeNode): string;  // директория не считая корня
    function readValueKey(hDefKey,TypesKey:integer;sSubKeyName,sValueName:string):string; // чтение значений ключей
    function GetNodeRooT(Node: TTreeNode): integer;
    function RenameRazdel(hDefKey:integer;sSubKeyName,NewsSubKeyName:string):boolean;
    function CreatNewRazdel(hDefKey:integer;sSubKeyName:string):integer;  // создает раздел в указаном подразделе
    function DeleteRazdel(hDefKey:integer;sSubKeyName:string):integer;  // удаляет раздел в выеленном подразделе
    function DeleteKeyValue(hDefKey:integer;sSubKeyName:string;sValueName:TstringList):integer; // удаляет ключ реестра
     function readBintoHextoInt(HexOrInt:string):boolean; // чтение параметра reg_binary и перевод в HEX или INt
    function newCreateBinary(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // сохранение бинарного ключа
    function read_REG_MULTI_SZ(hDefKey:integer;sSubKeyName,sValueName:string):TstringList; // чтение мультистрокового параметра

    function CreateFormForKey(TypeKey,Description,sValueName,sValue:string):boolean; // создание формы для редактирования ключа реестра
    function CopyKey(hDefKey:integer;sSubKeyName:string;sValueName:TstringList):boolean; // копирование ключей
    function PasteKey(hDefKey:integer;sSubKeyName:string;PasteVariant:Variant):boolean; // вставить  ключи
    function FindsNameValue(sValueName:string;TypeKey:integer):boolean; // поиск ключей перед вставкой в текущей директории
    Procedure LoadFavorite; // Чтение избранного
    function FavoriteAdd(RootKey,patch:string):boolean; // Добавить в избранное
    function Addmainmenu(CapNeme:string;ItemFavorite:integer):boolean; // добавить итем в меню избранное
    procedure OpenFavorite(Sender: TObject); // открыть избранное
    procedure cleanMainMenu;  // очистка основного меню от избранного при закрытии формы
    procedure saveFavorite;   // сохранение списка избранного в файл
    procedure deleteFavorites; // создание формы для удаления избранного
    procedure DeleteSelectFavorive(Sender:Tobject); // Удаление выделенного списка избранного
    function AddFaforiteKey(Description,namePC,hDefKey,sSubKeyName,sValueName,TypeKey,sValue:string):boolean; // Добавляет ключ в избранное
    function CreateFormForSaveKey(s:string):boolean; // создание формы для списка сохраненных ключей
    procedure OnShowFormSaveKey(sender:Tobject); // заполнение списка сохраненных ключей при показе формы
    procedure ButDelSaveKey (sender:Tobject);  // кнопка для удаления выделенных ключей из сохраненного списка
    function DeleteSelectSaveKey(id:integer):boolean; // Удаление ключа из списка сохраненных
    function ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;// импорт ключей в реестр из списка сохраненных ключей
    function StrTointTypeKey(TypeKey:string):integer; //передаем тип ключа в string получаем его в integer
    procedure ButImport(sender:Tobject);  // кнопка импорта
  public
   function CreateSetStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer; // созадние или изменение ключа реестра REG_SZ StringValue
   function CreateSetBinaryValue(hDefKey:integer;sSubKeyName,sValueName,uValue:string):integer;
   function CreateSetDWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:integer):integer;
   function CreateSetQWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:int64):integer;
   function CreateSetMultiStringValue(hDefKey:integer;sSubKeyName,sValueName:string;sValue:TStringList):integer;
   function CreateSetExpandedStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;
   function CreateBinarySaveKey(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // импортирование бинарного ключа в реестре
   function JumpTreeView(RootKey,Jpath:string):boolean; // переход в указанную директорию
   function MemoHexToInt(FullStr:string):string; // HEX числа из текста мемо переводим в INT
   function MemoInttohex(FullStr:string):string;  // INT числа из текста мемо переводим в HEX

  end;

var
  RegEdit: TRegEdit;

implementation
uses
 ShellAPI,
  ActiveX,
  ComObj,
  umain,
  MyDM,
  RegEditKeySave;
const
  wbemFlagForwardOnly = $00000020;




{$R *.dfm}
function IntegerInGetRooT(Node:Integer):String; // передаем root раздел в Integer, получаем его имя
var
val:string;
begin
try
val:='';
if node=HKEY_CLASSES_ROOT then val:='HKEY_CLASSES_ROOT';
if node=HKEY_CURRENT_USER then val:= 'HKEY_CURRENT_USER';
if node=HKEY_LOCAL_MACHINE then val:= 'HKEY_LOCAL_MACHINE';
if node=HKEY_USERS then val:= 'HKEY_USERS';
if node=HKEY_CURRENT_CONFIG then val:= 'HKEY_CURRENT_CONFIG';
result:=val;
except on E:EOleException do
        RegEdit.MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        RegEdit.MemoLog.Lines.Add('Ошибка (2) '+E.Message);  end;
end;
end;

function StringInGetRooT(Node:string):integer; // передаем имя раздела, получаем значение в Integer
var
val:integer;
begin
try
val:=0;
if node='HKEY_CLASSES_ROOT' then val:=HKEY_CLASSES_ROOT;
if node='HKEY_CURRENT_USER' then val:= HKEY_CURRENT_USER;
if node='HKEY_LOCAL_MACHINE' then val:= HKEY_LOCAL_MACHINE;
if node='HKEY_USERS' then val:= HKEY_USERS;
if node='HKEY_CURRENT_CONFIG' then val:= HKEY_CURRENT_CONFIG;
result:=val;
except on E:EOleException do
        RegEdit.MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        RegEdit.MemoLog.Lines.Add('Ошибка (3) '+E.Message);  end;
end;
end;

function TRegEdit.StrTointTypeKey(TypeKey:string):integer; //передаем тип ключа в string получаем его в integer
begin
if TypeKey='REG_SZ' then result:=1;
if TypeKey='REG_EXPAND_SZ' then result:=2;
if TypeKey='REG_BINARY' then result:=3;
if TypeKey='REG_DWORD' then result:=4;
if TypeKey='REG_MULTI_SZ' then result:=7;
if TypeKey='REG_QWORD' then result:=11;
end;

function DescriptionImportKey:string;
var
Description:string;
begin
while Description='' do
Begin
Description:=DateTimeToStr(now);
if not InputQuery('Введите описание для сохраняемых ключей', 'Описание', Description)
then begin Description:=''; exit; end;
end;
result:=Description;
end;

procedure TRegEdit.ButImport(sender:Tobject); // Кнопка импорта из списка сохраненных ключей реестра
var
i:integer;
begin
try
for I := 0 to LVSaveKey.Items.Count-1 do
 begin
 if LVSaveKey.Items[i].Selected then
 if ImportSaveKeyhDefKey(LVSaveKey.Items[i].SubItems[5], // hDefKey
                      LVSaveKey.Items[i].SubItems[6], // sSubKeyName
                      LVSaveKey.Items[i].SubItems[2], // sValueName
                      LVSaveKey.Items[i].SubItems[4], // sValue
                      LVSaveKey.Items[i].SubItems[3]) // TypeKey
  =0 then MemoLog.Lines.Add('Ключ реестра '+LVSaveKey.Items[i].SubItems[2]+' импортирован в раздел '+LVSaveKey.Items[i].SubItems[6]+' на компьютере '+LVSaveKey.Items[i].SubItems[1])
 end;
except on E: Exception do MemoLog.Lines.Add('Ошибка (4) '+E.Message);
end;
 end;

function TRegEdit.ImportSaveKeyhDefKey (hDefKey,sSubKeyName,sValueName,sValue,TypeKey:string):integer;
 /// импорт ключей в реестр из списка сохраненных
var
i,IntTypeKey,res:integer;
listT:TstringList;
function StringToList(sValue:string):Tstringlist; // из строки делаем список строк
var
s:string;
begin
result:=TStringList.Create;
result.Text:=sValue;
end;
function StrQDWord(s:string):string; // из строки со значением HEX и Integer выделяем только Int
begin
delete(s,1,pos('(',s));
result:=Copy(s,1,pos(')',s)-1);
end;
begin
try
IntTypeKey:=StrTointTypeKey(TypeKey);
case IntTypeKey of
   1:res:=CreateSetStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   2:res:=CreateSetExpandedStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue);
   3:res:=CreateBinarySaveKey(StringInGetRooT(hDefKey),sSubKeyName,sValueName,sValue,'HEX');
   4:res:=CreateSetDWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint(StrQDWord(sValue)));
   7:begin
    try
     listT:=TStringList.Create;
     listT:=StringToList(sValue);
     res:=CreateSetMultiStringValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,StringToList(sValue));
     finally
     listT.Free;
     end;
     end;
   11:res:=CreateSetQWORDValue(StringInGetRooT(hDefKey),sSubKeyName,sValueName,strtoint64(StrQDWord(sValue)));
end;
if res<>0 then MemoLog.Lines.Add('Импорт ключа '+sValueName+': '+SysErrorMessage(res));
result:=res;
except on E: Exception do begin MemoLog.Lines.Add('Ошибка (5) '+E.Message); result:=res; end;
end;
end;

procedure TRegEdit.OnShowFormSaveKey(sender:Tobject); // заполняем список сохраненных ключей реестра
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
function TypeKeyImage(s:string):integer;
var
i:integer;
begin
i:=0;
if s='REG_SZ' then i:=1;
if s='REG_EXPAND_SZ' then i:=2;
if s='REG_BINARY' then i:=3;
if s='REG_DWORD' then i:=4;
if s='REG_MULTI_SZ' then i:=7;
if s='REG_QWORD' then i:=11;
result:=i;
end;
begin
try
try
if not datam.TableExists('REGEDIT_KEY') then exit; // если таблица еще не создана то выходим
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
TransactionAdd.Options.ReadOnly:=true;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='select * from REGEDIT_KEY';
FDQueryAdd.Open;
FDQueryAdd.First; // переход к первой строке
while not FDQueryAdd.Eof do
 begin
 with LVSaveKey.Items.Add do // заполняем список
  begin
  ImageIndex:=TypeKeyImage(FDQueryAdd.FieldByName('TypeKey').AsString);
  Caption:=FDQueryAdd.FieldByName('ID_KEY').AsString;
  SubItems.add(FDQueryAdd.FieldByName('Description_key').AsString);
  SubItems.add(FDQueryAdd.FieldByName('NamePC').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sValueName').AsString);
  SubItems.add(FDQueryAdd.FieldByName('TypeKey').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sValue').AsString);
  SubItems.add(FDQueryAdd.FieldByName('hDefKey').AsString);
  SubItems.add(FDQueryAdd.FieldByName('sSubKeyName').AsString);
  end;
 FDQueryAdd.Next; // переход к следующей строке
 end;
TransactionAdd.Commit;
finally
FDQueryAdd.Free;
TransactionAdd.Free;
end;
except on E: Exception do MemoLog.Lines.Add('Ошибка (6) '+E.Message);
end;
end;

function TRegEdit.DeleteSelectSaveKey(id:integer):boolean; // удаление записи из таблицы REGEDIT_KEY
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
begin
try
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
TransactionAdd.Options.ReadOnly:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='DELETE FROM REGEDIT_KEY WHERE ID_KEY=:a';
FDQueryAdd.ParamByName('a').asInteger:=ID;
FDQueryAdd.ExecSQL;
TransactionAdd.Commit;
finally
FDQueryAdd.Free;
TransactionAdd.Free;
end;
result:=true;
except on E: Exception do begin MemoLog.Lines.Add('Ошибка (7) '+E.Message); result:=false; end;
end;
end;

procedure TRegEdit.ButDelSaveKey (sender:Tobject); // удалить выбранные ключи из списка сохраненных
var
i:integer;
begin
 for I := LVSaveKey.Items.Count-1 DownTo 0 do
   begin
     if LVSaveKey.Items[i].Selected then
     if DeleteSelectSaveKey(strtoint(LVSaveKey.Items[i].Caption)) then LVSaveKey.Items.Delete(i);
   end;
end;

function TRegEdit.CreateFormForSaveKey(s:string):boolean; // создание формы для списка сохраненных ключей
begin
try
FormSaveKey:=Tform.Create(self);
FormSaveKey.Caption:='Список сохраненных ключей реестра';
FormSaveKey.position:=poOwnerFormCenter;
FormSaveKey.BorderStyle:=bsSizeable;
FormSaveKey.Width:=900;
FormSaveKey.Height:=550;
FormSaveKey.Position:=poOwnerFormCenter;
FormSaveKey.OnClose:=FormKeyClose;
FormSaveKey.OnShow:=OnShowFormSaveKey;
PanelSaveKey:=TPanel.Create(FormSaveKey);
PanelSaveKey.Parent:=FormSaveKey;
PanelSaveKey.Align:=alTop;
PanelSaveKey.Height:=40;
ButKeyImport:=TButton.Create(FormSaveKey);
ButKeyImport.Parent:=PanelSaveKey;
ButKeyImport.Caption:='Импорт';
ButKeyImport.Top:=5;
ButKeyImport.Left:=10;
ButKeyImport.OnClick:=ButImport;
/////////////////////////////////
ButKeyDel:=TButton.Create(FormSaveKey);
ButKeyDel.Parent:=PanelSaveKey;
ButKeyDel.Caption:='Удалить';
ButKeyDel.Top:=5;
ButKeyDel.Left:=20+ButKeyImport.Width;
ButKeyDel.OnClick:=ButDelSaveKey;
///////////////////////////
ButKeyClose:=TButton.Create(FormSaveKey);
ButKeyClose.Parent:=PanelSaveKey;
ButKeyClose.Caption:='Закрыть';
ButKeyClose.Top:=5;
ButKeyClose.Left:=10+ButKeyDel.Width+ButKeyDel.Left;
ButKeyClose.OnClick:=FormSaveKeyClose;
////////////////////////
LVSaveKey:=TlistView.Create(FormSaveKey);
LVSaveKey.Parent:=FormSaveKey;
LVSaveKey.Align:=alClient;
LVSaveKey.ViewStyle:=vsReport;
LVSaveKey.SmallImages:=ImageList2;
LVSaveKey.RowSelect:=true;
LVSaveKey.MultiSelect:=true;
LVSaveKey.ReadOnly:=true;
with LVSaveKey.Columns.Add do
begin
caption:='';
Width:=20;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Описание';
Width:=150;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Компьютер';
Width:=150;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Имя ключа';
Width:=150;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Тип';
Width:=150;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Значение';
Width:=250;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Корневой раздел';
Width:=150;
end;
with LVSaveKey.Columns.Add do
begin
caption:='Путь';
Width:=150;
end;
FormSaveKey.ShowModal;
except on E: Exception do MemoLog.Lines.Add('Ошибка (8) '+E.Message);
end;
end;

procedure TRegEdit.DeleteSelectFavorive(Sender:Tobject);
var
i,z:integer;
begin
if FLV.Items.Count>0 then
for I := FLV.Items.Count-1 DownTo 0 do
begin
try
  if FLV.Items[i].Selected then
  begin
    FavoriteTXT.Delete(i);
    FLV.items[i].Delete;
    MainMenu1.Items[0].Delete(i+4); // удаление пункта меню
    for z := i+4 to MainMenu1.Items[0].Count-1 do // изменяем tag  для верного доступа к строке в FavoriteTXT
    MainMenu1.Items[0].Items[z].tag:=MainMenu1.Items[0].Items[z].tag-1;
  end;
  except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (9): '+e.Message);
  end;
end;

end;

procedure TRegEdit.deleteFavorites;
var
i:integer;
begin
try
FormKey:=TForm.Create(RegEdit);
FormKey.Caption:='Список избранного';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=420;
FormKey.Height:=345;
FormKey.OnClose:=FormKeyClose;
////////////
FLabel:=TLabel.Create(FormKey);
FLabel.Parent:=FormKey;
FLabel.Top:=5;
FLabel.Left:=5;
FLabel.Width:=260;
FLabel.caption:='Выберите список для удаления';
/// /////////////////////
FLV:=TListView.Create(FormKey);
FLV.Parent:=FormKey;
FLV.Top:=18;
FLV.Left:=3;
FLV.Width:=408;
FLV.Height:=255;
FLV.TabOrder:=0; // переводим фокус в поле со значением параметра
FLV.MultiSelect:=true;
FLV.ReadOnly:=true;
FLV.RowSelect:=true;
FLV.ViewStyle:=vsReport;
with FLV.Columns.Add  do
begin
  Caption:='';
  Width:=410;
end;
for I := 0 to FavoriteTXT.Count-1 do
begin
 with Flv.Items.add do
 Caption:=FavoriteTXT.ValueFromIndex[i]+':'+FavoriteTXT.Names[i];
end;
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Удалить';
ButtonOk.Top:=279;
ButtonOk.Left:=230;
ButtonOk.OnClick:=DeleteSelectFavorive;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Закрыть';
ButtonNo.Top:=279;
ButtonNo.Left:=325;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
except on E: Exception do MemoLog.Lines.Add('Ошибка (10) '+E.Message);
end;
end;



procedure TRegEdit.saveFavorite;
begin
FavoriteTXT.SaveToFile(extractfilepath(application.ExeName)+'\RegFavorite.txt');
end;

procedure TRegEdit.OpenFavorite(Sender: TObject);  // процедура открытия директории из избранного
var
i,step,num:integer;
RootKey,Path:string;
begin
  try
  RootKey:='';
  Path:='';
 if (sender is Tmenuitem) then
  begin
  num:=(Sender as TMenuItem).tag;
  Path:=FavoriteTXT.Names[num];
  RootKey:=FavoriteTXT.ValueFromIndex[num];
  MemoLog.Lines.Add(RootKey+':'+Path);
  EditRooT.Text:=RootKey; // указываем раздел в который заходим
   if (RootKey<>'') and (path<>'') then
   JumpTreeView(RootKey,path);
  end;
  except on E: Exception do memolog.Lines.Add(inttostr(step)+')Произошла непредвиденная ошибка (11): '+e.Message);
  end;
end;

procedure TRegEdit.cleanMainMenu;  // очистка меню перед закрытием формы
var
i:integer;
begin
try
for I := MainMenu1.Items[0].Count-1 DownTo 4 do
 MainMenu1.Items[0].Delete(i);
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (12): '+e.Message);
end;
 end;

function TRegEdit.FavoriteAdd(RootKey,patch:string):boolean;
begin
FavoriteTXT.Add(patch+'='+RootKey);// Добавляем в список (TstringList) избранного
AddMainMenu(patch,FavoriteTXT.Count-1);               // добавляем избранное в меню
end;

function TRegEdit.AddMainMenu(CapNeme:string;ItemFavorite:integer):boolean; // добавляет итем в меню избранном
var
menuitem:Tmenuitem;
begin
  MenuItem:=TMenuItem.Create(MainMenu1);
  menuitem.Caption:=CapNeme;
  menuitem.Tag:=ItemFavorite;
  menuitem.OnClick:=OpenFavorite;
  MainMenu1.Items[0].Add(menuitem);
end;

Procedure TRegEdit.LoadFavorite; // чтение избранного при открытии формы
var
i:integer;
begin
try
FavoriteTXT:=TStringList.Create;
if FileExists(extractfilepath(application.ExeName)+'\RegFavorite.txt') then
FavoriteTXT.LoadFromFile(extractfilepath(application.ExeName)+'\RegFavorite.txt');
if FavoriteTXT.Count>0 then
begin
 for I := 0 to FavoriteTXT.Count-1 do
 begin
  Addmainmenu(FavoriteTXT.Names[i],i);
 end;
end;
except on E: Exception do MemoLog.Lines.Add('Ошибка (13) '+e.Message);
end;
end;

procedure TRegEdit.N10Click(Sender: TObject); // главное меню, добавить в избранное
begin
if (EditRooT.Text<>'')and (EditPath.Text<>'') then
FavoriteAdd(EditRooT.Text,EditPath.Text);
end;

procedure TRegEdit.N11Click(Sender: TObject);
begin
deleteFavorites;
end;

procedure TRegEdit.N13Click(Sender: TObject); // экспорт в excel
begin
 if LVKey.Items.Count<>0 then
 frmDomainInfo.popupListViewSaveAs(LVKey,'Сохранение списка ключей реестра','Ключи реестра');
end;

function TRegEdit.AddFaforiteKey(Description,NamePC,hDefKey,sSubKeyName,sValueName,TypeKey,sValue:string):boolean;
var
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
begin
try
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
FDQueryAdd.SQL.Clear;
FDQueryAdd.SQL.Text:='update or insert into REGEDIT_KEY (Description_key ,NamePC, hDefKey ,sSubKeyName, sValueName, sValue, TypeKey) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7) MATCHING(NamePC, hDefKey ,sSubKeyName, sValueName, sValue, TypeKey)';
FDQueryAdd.params.ParamByName('p1').AsString:=''+Description+'';
FDQueryAdd.params.ParamByName('p2').AsString:=''+NamePC+'';
FDQueryAdd.params.ParamByName('p3').AsString:=''+hDefKey+'';
FDQueryAdd.params.ParamByName('p4').AsString:=''+sSubKeyName+'';
FDQueryAdd.params.ParamByName('p5').AsString:=''+sValueName+'';
FDQueryAdd.params.ParamByName('p6').AsString:=''+sValue+'';
FDQueryAdd.params.ParamByName('p7').AsString:=''+TypeKey+'';
FDQueryAdd.ExecSQL;
TransactionAdd.Commit;
finally
FDQueryAdd.free;
TransactionAdd.Free;
end;
result:=true;
except on E: Exception do
begin
  MemoLog.Lines.Add('Ошибка добавления ключа '+sValueName+' в избранное (14):'+e.Message);
  result:=false;
end;
end;
end;

procedure TRegEdit.N14Click(Sender: TObject);
var
i:integer;
description:string;
function findKey(sValueName:string):boolean;
begin
result:=false;
if sValueName='REG_SZ' then result:=true;
if sValueName='REG_EXPAND_SZ' then result:=true;
if sValueName='REG_BINARY' then result:=true;
if sValueName='REG_DWORD' then result:=true;
if sValueName='REG_MULTI_SZ' then result:=true;
if sValueName='REG_QWORD' then result:=true;
end;
begin
if DataM.СreateTablForRegEdit then // проверка существует ли таблица, если нет то создаем
if LVKey.SelCount<>0 then
begin
description:=DescriptionImportKey; // Описание для сохраняемых ключей
if description='' then exit; // если описание пусто то выходим

 for I := 0 to LVKey.Items.Count-1 do
   begin
     if (LVKey.Items[i].Selected)and(findKey(LVKey.Items[i].SubItems[0])) then
      begin
        if AddFaforiteKey(description,ComboBox2.Text,EditRooT.Text,EditPath.Text,LvKey.items[i].Caption,LvKey.items[i].SubItems[0],LvKey.items[i].SubItems[1]) then
        MemoLog.Lines.Add('Ключ '+LvKey.items[i].Caption+' добавлен в избранное');
      end;
   end;
end;
end;

procedure TRegEdit.N15Click(Sender: TObject); // открвыть список сохраненных ключей
begin
RegKeySave.Button4.Enabled:=false; // отключаем кнопку добавления в задачу
RegKeySave.Button1.Enabled:=true;  // включаем кнопку импорта в реестр
RegKeySave.PopupEditKey.Items[1].Enabled:=true; // включаем пункты меню для импорта
RegKeySave.ShowModal;  // открываем форму
end;

procedure TRegEdit.N17Click(Sender: TObject);
begin
try
// если переменная для копирования и вставки пустая то item вставка не активный
if (not VarIsArray(CopyPasteVariant)) or (VarIsNull(CopyPasteVariant)) then MainMenu1.Items[1].Items[2].Enabled:=false
else MainMenu1.Items[1].Items[2].Enabled:=true;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (15): '+e.Message);
end;
end;

Function TRegEdit.FindsNameValue(sValueName:string;TypeKey:integer):boolean; // поиск ключей перед вставкой в текущей директории
var
i,z:integer;
s:bool;
function findKey(TypeKey:integer):string;
begin
if TypeKey=1 then result:='REG_SZ';
if TypeKey=2 then result:='REG_EXPAND_SZ';
if TypeKey=3 then result:='REG_BINARY';
if TypeKey=4 then result:='REG_DWORD';
if TypeKey=7 then result:='REG_MULTI_SZ';
if TypeKey=11 then result:='REG_QWORD';
end;
begin
try
for I := 0 to LVKey.Items.Count-1 do
begin
s:=true;
 if (sValueName=LVKey.Items[i].Caption) and (findKey(TypeKey)=LVKey.Items[i].SubItems[0]) then
  Begin
  z:=MessageDlg('Заменить значение ключа '+sValueName+'?', mtConfirmation,[mbYes,mbCancel],0);
  if z=IDYes then // заменить ключ
    begin
    s:=true;
    break;
    end
  else // не заменять
    begin
    s:=false;
    break;
    end;
  End;
end;
Result:=s;
except on E: Exception do begin result:=false; memolog.Lines.Add('Произошла непредвиденная ошибка (16): '+e.Message); end;
end;
end;

function TRegEdit.PasteKey(hDefKey:integer;sSubKeyName:string;PasteVariant:Variant):boolean;
var             // вставить скопированные ключи
 i,TypesKey,z:integer;
 FWbemObjectSet: OLEVariant;
 InParam: OleVariant;
 mas:variant;
begin
try
if (not VarIsArray(CopyPasteVariant)) or (VarIsNull(CopyPasteVariant)) then exit; // если переменная не является массивом или равна null то и вставлять нечего
 FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
 for I := VarArrayLowBound(PasteVariant,1) to VarArrayHighBound(PasteVariant,1) do
  BEGIN
   mas:=VarArrayCreate([0,2],varVariant); //  0 - TypeKey(integer), 1 - sValueName(string) , 2 - inParam (variant)
   mas:=PasteVariant[i];
   if (integer(mas[0])<>0) and (FindsNameValue(vartostr(mas[1]),integer(mas[0]))) then
   begin
   case integer(mas[0]) of
   1:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetStringValue').InParameters.SpawnInstance_();  // Получает строковое значение данных именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.sValue:=mas[2];     // значение ключа  строка
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetStringValue', InParam);
    end;
   2:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetExpandedStringValue').InParameters.SpawnInstance_(); //Получает расширенное строковое значение данных именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.sValue:=mas[2];     // значение ключа  строка
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetExpandedStringValue', InParam);
    end;
   3:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetBinaryValue').InParameters.SpawnInstance_(); //Получает значение двоичных данных именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.uValue:=mas[2];     // значение ключа  массив байт
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetBinaryValue', InParam);
    end;
   4:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetDWORDValue').InParameters.SpawnInstance_(); // Получает значение DWORD именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.uValue:=mas[2];     // значение ключа integer
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetDWORDValue', InParam);
    end;
   7:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetMultiStringValue').InParameters.SpawnInstance_(); //Получает несколько строковых значений данных именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.sValue:=mas[2];     // значение ключа массив строк;
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetMultiStringValue', InParam);
    end;
   11:
    begin
    InParam:= FWbemObjectSet.Methods_.Item('SetQWORDValue').InParameters.SpawnInstance_(); // Получает значения данных QWORD именованного значения.
    InParam.hDefKey:=hDefKey;  // root
    InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь куда копируем
    InParam.sValueName:=vartostr(mas[1]); // имя ключа
    InParam.uValue:=mas[2];     // значение ключа int64
    FWMIServiceRE.ExecMethod('StdRegProv', 'SetQWORDValue', InParam);
    end;
   end;
  end;
 VariantClear(InParam);
 VarClear(mas);
END;
result:=true;
VariantClear(FWbemObjectSet);
except on E:EOleException do
        MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        MemoLog.Lines.Add(' Ошибка (17): '+E.Message); result:=false; end;
end;

end;


function TRegEdit.CopyKey(hDefKey:integer;sSubKeyName:string;sValueName:TstringList):boolean;
 var                  // копировать выделенные ключи
 i,step,TypesKey:integer;
 FWbemObjectSet: OLEVariant;
 InParam: OLEVariant;
 outParam:variant;
function findKey(sValueName:string):integer;
begin
if sValueName='REG_SZ' then result:=1;
if sValueName='REG_EXPAND_SZ' then result:=2;
if sValueName='REG_BINARY' then result:=3;
if sValueName='REG_DWORD' then result:=4;
if sValueName='REG_MULTI_SZ' then result:=7;
if sValueName='REG_QWORD' then result:=11;
end;
begin
try
 if (VarIsArray(CopyPasteVariant)) or (not VarIsNull(CopyPasteVariant)) then CopyPasteVariant:=Unassigned; // если переменная является массивом значит надо её почистить перед заполнением
 CopyPasteVariant:=VarArrayCreate([0,sValueName.Count-1],VarVariant);
 FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
 ////////////////////////////////////////////////////////////////////
 for I :=0 to sValueName.Count-1 do
 begin
 TypesKey:=findKey(sValueName.ValueFromIndex[i]); // тип ключа
 if TypesKey<>0 then
  begin
    try
    step:=1;
    case TypesKey of
    1: InParam:= FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_(); // Получает строковое значение данных именованного значения.
    2: InParam:= FWbemObjectSet.Methods_.Item('GetExpandedStringValue').InParameters.SpawnInstance_(); //Получает расширенное строковое значение данных именованного значения.
    3: InParam:= FWbemObjectSet.Methods_.Item('GetBinaryValue').InParameters.SpawnInstance_(); //Получает значение двоичных данных именованного значения.
    4: InParam:= FWbemObjectSet.Methods_.Item('GetDWORDValue').InParameters.SpawnInstance_(); // Получает значение DWORD именованного значения.
    7: InParam:= FWbemObjectSet.Methods_.Item('GetMultiStringValue').InParameters.SpawnInstance_(); //Получает несколько строковых значений данных именованного значения.
    11:InParam:= FWbemObjectSet.Methods_.Item('GetQWORDValue').InParameters.SpawnInstance_(); // Получает значения данных QWORD именованного значения.
    end;
    step:=2;
    InParam.hDefKey:=hDefKey;
    InParam.sSubKeyName:=sSubKeyName; // путь в разделе
    InParam.sValueName:=sValueName.Names[i];
    step:=3;
    case TypesKey of
      1:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetStringValue', InParam); //[out] string sValue
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.sValue]);
      end;
      2:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetExpandedStringValue', InParam); //[out] string sValue
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.sValue]);
      end;
      3:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetBinaryValue', InParam);  //uValue [out] An array of binary bytes.
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.uValue]);
      end;
      4:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetDWORDValue', InParam);  // [out] uint32 uValue
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.uValue]);
      end;
      7:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetMultiStringValue', InParam); //[out] string sValue[]
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.sValue]);
      end;
      11:begin
      outParam:=FWMIServiceRE.ExecMethod('StdRegProv', 'GetQWORDValue', InParam); //[out] uint64 uValue
      CopyPasteVariant[i]:= VarArrayOf([TypesKey,sValueName.Names[i],OutParam.uValue]);
      end;
    end;
    step:=4;
    VariantClear(InParam);
    VarClear(OutParam);
    result:=true;
    except on E:EOleException do
            MemoLog.Lines.Add(Format('EOleException %s %x', [E.Message,E.ErrorCode]));
           on E:Exception do begin MemoLog.Lines.Add('Ошибка копирования ключа (19) '+sValueName.Names[i]+' : '+E.Message); result:=false; end;
    end;
   end; // if TypeKey<>0
 end; // цикл по ключам

 VariantClear(FWbemObjectSet);

 except on E:Exception do
 begin
  MemoLog.Lines.Add('Общая ошибка копирования (20): '+E.Message); result:=false;
 end;
 end;
 ////////////////////////////////////////////////////////////////////
end;


procedure TRegEdit.ButtonUpdateClick(Sender: TObject);
begin
if (frmDomainInfo.ping(ComboBox2.Text)) then
begin
ConnectRemote;  // соединение
end;
end;



procedure TRegEdit.clearTTreeView;
var
Rnode: TTreeNode;
begin
TreeViewFolders.Items.BeginUpdate;
try
CurPCConnectRegEdit:=Combobox2.Text; // в переменной сохраняется имя текущего компа
TreeViewFolders.Items.Clear;
Rnode:=TreeViewFolders.items.Add(Nil,Combobox2.Text);
TreeViewFolders.items.AddChild(Rnode,'HKEY_CLASSES_ROOT');
TreeViewFolders.items.AddChild(Rnode,'HKEY_LOCAL_MACHINE');
TreeViewFolders.items.AddChild(Rnode,'HKEY_CURRENT_USER');
TreeViewFolders.items.AddChild(Rnode,'HKEY_USERS');
TreeViewFolders.items.AddChild(Rnode,'HKEY_CURRENT_CONFIG');
Rnode.Expanded:=true;
finally
TreeViewFolders.Items.EndUpdate;
end;
LVKey.Items.BeginUpdate;
try
LVKey.Items.Clear;
finally
LVKey.Items.EndUpdate;
end;
////сначала все обнуляем
end;

function TRegEdit.ConnectWMI: Boolean;
begin
 Result:=False;
 try
    FSWbemLocatorRE:=Unassigned;
    FWMIServiceRE:=Unassigned;
    FSWbemLocatorRE := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIServiceRE   := FSWbemLocatorRE.ConnectServer(ComboBox2.Text, 'root\DEFAULT', EditUser.Text, EditPass.text,'','',128);
    FWMIServiceRE.Security_.impersonationlevel:=3;
    FWMIServiceRE.Security_.authenticationLevel := 6;
    Result:=True;
 except
    on E:EOleException do
        MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do
        MemoLog.Lines.Add('Ошибка (21): '+E.Message);
 end;
end;



procedure TRegEdit.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var
  i:integer;
begin
try
if (key=VK_RETURN) then
if  (frmDomainInfo.ping(ComboBox2.Text)) then
  Begin
  i:=IDCANCEL;
  i:=MessageDlg('Подключится к реестру компьютера ' +ComboBox2.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
  if i=IDYes then
   begin
   ConnectRemote  // если да то подключаемся к реестру нового компа
   end
  else                             // иначе оставляем как было
   begin
   ComboBox2.DroppedDown:=false;
   ComboBox2.Text:=CurPCConnectRegEdit;
   end;
  End
else ComboBox2.Text:=CurPCConnectRegEdit;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (22): '+e.Message);
end;
end;

procedure TRegEdit.ComboBox2Select(Sender: TObject);
var
i:integer;
begin
if (frmDomainInfo.ping(ComboBox2.Text)) then
begin
  i:=MessageDlg('Подключится к реестру компьютера ' +ComboBox2.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
  if i=IDYes then
   begin
   ConnectRemote  // если да то подключаемся к реестру нового компа
   end
  else
  begin
  ComboBox2.DroppedDown:=false;
  ComboBox2.Text:=CurPCConnectRegEdit;
  end;
end
else ComboBox2.Text:=CurPCConnectRegEdit;

end;

function TRegEdit.ConnectRemote:boolean;
begin
  begin
    StatusBar1.Panels[0].Text:=Format('Подключение к %s',[ComboBox2.Text]);
    try
      if ConnectWMI then // если установлено соединение
      begin
       clearTTreeView;  // чистим и создаем основные разделы реестра
       result:=true;
      end
      else
      begin
      ShowMessage(Format('Невозможно установить соединение с сервером %s ',[ComboBox2.Text]));
      result:=false;
      end;
    finally
      StatusBar1.Panels[0].Text:='';
    end;
  end

end;

function TRegEdit.GetNodeFullPath(Node: TTreeNode): string;
var
path:string;
begin
path:='';
while Node.Parent.Text<>Combobox2.Text do
begin
if node.Parent<>nil then path:=Node.Text+'\'+path;
Node:=Node.Parent;
end;
delete(path,length(path),1);
EditPath.text:=path;
result:=path;
end;

function TRegEdit.GetNodeRooT(Node: TTreeNode):integer; // передаем ноду, получаем имя корневого раздела
begin
try
while node.Parent.Text<>Combobox2.Text do Node:=Node.Parent;
EditRooT.text:=node.Text;
result:=0;
if node.Text='HKEY_CLASSES_ROOT' then result:=HKEY_CLASSES_ROOT;
if node.Text='HKEY_CURRENT_USER' then result:= HKEY_CURRENT_USER;
if node.Text='HKEY_LOCAL_MACHINE' then result:= HKEY_LOCAL_MACHINE;
if node.Text='HKEY_USERS' then result:= HKEY_USERS;
if node.Text='HKEY_CURRENT_CONFIG' then result:= HKEY_CURRENT_CONFIG;
except on E:EOleException do
        MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        MemoLog.Lines.Add('Ошибка (23) :'+E.Message);  end;
end;
end;

procedure TRegEdit.LVKeyDblClick(Sender: TObject);
begin
if LVKey.SelCount<>1 then exit;
CreateFormForKey(LVKey.Selected.SubItems[0],'',LVKey.Selected.Caption,LVKey.Selected.SubItems[1]);
end;





procedure TRegEdit.FormShow(Sender: TObject);
begin
try
LVkey.SetFocus;
OleInitialize(nil);
LoadFavorite; // читаем файл с избранным и добавляем избранное в меню
if ConnectRemote then   // соединение
MemoLog.Lines.Add('Успешое подключение к реестру компьютера '+ComboBox2.Text);
except on E: Exception do  ShowMessage(Format('Ошибка (1): %s ',[E.Message]));
end;
end;

procedure TRegEdit.TreeViewFoldersChange(Sender: TObject; Node: TTreeNode);
var
sSubKeyName:string;
hkeyRoot:integer;
i:integer;
begin
if Node.Text=ComboBox2.Text then exit; // если кликнули на корневой, пустой то выходим
  try
    hkeyRoot:=GetNodeRooT(Node);
    sSubKeyName:='';
    sSubKeyName:=GetNodeFullPath(Node);
    if Assigned(Node) and (Node.Count=0) and not VarIsClear(FWMIServiceRE) then
      UpdateFolderTreeItem(Node,hkeyRoot,sSubKeyName); // обновляем дерево
    if Assigned(Node)and not VarIsClear(FWMIServiceRE) then
      UpdateKeyTreeItem(hkeyRoot,sSubKeyName); //чтение ключей
    except
    on E: Exception do
     MemoLog.lines.add(Format('Ошибка (24) : %s ',[E.Message]));
  end;
end;

function TRegEdit.UpdateFolderTreeItem(Node: TTreeNode;hDefKey:integer;sSubKeyName:string):boolean;
Var                                    //чтение разделов реестра
  lNode         : TTreeNode;
  FWbemObjectSet: OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
begin
    try
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     TreeViewFolders.Items.BeginUpdate;
     StatusBar1.Panels[0].Text:='Чтение разделов реестра '+sSubKeyName;
     InParam:= FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
     InParam.hDefKey:=hDefKey;
     InParam.sSubKeyName:=sSubKeyName; // путь в разделе
     OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'EnumKey', InParam);
     if VarArrayDimCount(OutParam.sNames)<>0 then  // если массив не пустой
     for i:= VarArrayLowBound(OutParam.sNames,1) to VarArrayHighBound(OutParam.sNames,1) do
     begin
     lNode:=TreeViewFolders.Items.AddChild(Node,OutParam.sNames[i]);
     lNode.ImageIndex:=0;
     lNode.SelectedIndex:=0;
     end;
       TreeViewFolders.Items.EndUpdate;
       VariantClear(OutParam);
       VariantClear(InParam);
       VariantClear(FWbemObjectSet);
       StatusBar1.Panels[0].Text:='';
       result:=true;
    except
    on E: Exception do
    begin
     result:=false;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('UpdateFolder Ошибка чтения рееста (25): %s ',[E.Message]));
     TreeViewFolders.Items.EndUpdate;
    end;
  end;
  end;

function TRegEdit.JumpFolderTreeItem(Node: TTreeNode;hDefKey:integer;sSubKeyName:string):TTreeNode;
Var                                    //чтение разделов реестра
  lNode         : TTreeNode;
  FWbemObjectSet: OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
begin
    try
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     TreeViewFolders.Items.BeginUpdate;
     StatusBar1.Panels[0].Text:='Чтение разделов реестра '+sSubKeyName;
     InParam:= FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
     InParam.hDefKey:=hDefKey;
     InParam.sSubKeyName:=sSubKeyName; // путь в разделе
     OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'EnumKey', InParam);
     if VarArrayDimCount(OutParam.sNames)<>0 then  // если массив не пустой
     for i:= VarArrayLowBound(OutParam.sNames,1) to VarArrayHighBound(OutParam.sNames,1) do
     begin
     lNode:=TreeViewFolders.Items.AddChild(Node,OutParam.sNames[i]);
     lNode.ImageIndex:=0;
     lNode.SelectedIndex:=0;
     end;
     result:=node;
       TreeViewFolders.Items.EndUpdate;
       VariantClear(OutParam);
       VariantClear(InParam);
       VariantClear(FWbemObjectSet);
       StatusBar1.Panels[0].Text:='';
    except
    on E: Exception do
    begin
     result:=node;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('JumpFolder Ошибка чтения рееста (26): %s ',[E.Message]));
     TreeViewFolders.Items.EndUpdate;
    end;
  end;
  end;

function ByteToStr(bytes: Variant): string;
const
  BytesHex: array[0..15] of char =
    ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F');
var
  i, len: integer;
  s:string;
begin
result:='';
s:='';
  if VarArrayHighBound(bytes,1)>0 then
  begin
    for i := VarArrayLowBound(bytes,1) to VarArrayHighBound(bytes,1) do
     begin
      s:=s+ BytesHex[integer(bytes[i]) shr 4];
      s:=s+ BytesHex[integer(bytes[i]) and $0F];
      s:=s+ ' ';
     end;
  end;
 result:=s;
end;

function TRegEdit.readValueKey(hDefKey,TypesKey:integer;sSubKeyName,sValueName:string):string;
Var                                       //чтение значений ключей
  FWbemObjectSet,FWMIService     : OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
 begin
    result:='';
    FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
    try
    case TypesKey of
    1: InParam:= FWbemObjectSet.Methods_.Item('GetStringValue').InParameters.SpawnInstance_(); // Получает строковое значение данных именованного значения.
    2: InParam:= FWbemObjectSet.Methods_.Item('GetExpandedStringValue').InParameters.SpawnInstance_(); //Получает расширенное строковое значение данных именованного значения.
    3,8: InParam:= FWbemObjectSet.Methods_.Item('GetBinaryValue').InParameters.SpawnInstance_(); //Получает значение двоичных данных именованного значения.
    4: InParam:= FWbemObjectSet.Methods_.Item('GetDWORDValue').InParameters.SpawnInstance_(); // Получает значение DWORD именованного значения.
    7: InParam:= FWbemObjectSet.Methods_.Item('GetMultiStringValue').InParameters.SpawnInstance_(); //Получает несколько строковых значений данных именованного значения.
    11:InParam:= FWbemObjectSet.Methods_.Item('GetQWORDValue').InParameters.SpawnInstance_(); // Получает значения данных QWORD именованного значения.
    end;
    InParam.hDefKey:=hDefKey;
    InParam.sSubKeyName:=sSubKeyName; // путь в разделе
    InParam.sValueName:=sValueName;
    case TypesKey of
    1:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetStringValue', InParam); //[out] string sValue
    if OutParam.sValue<>null then result:=vartostr(OutParam.sValue)
    else result:='(значение не присвоено)';
     end;
    2:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetExpandedStringValue', InParam); //[out] string sValue
    result:=vartostr(OutParam.sValue);
    end;
    3:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetBinaryValue', InParam);  //uValue [out] An array of binary bytes.
    //for i := VarArrayLowBound(OutParam.uValue,1) to VarArrayHighBound(OutParam.uValue,1) do
    //result:=result+inttostr(OutParam.uValue[i])+' ';
    result:=ByteToStr(OutParam.uValue);
    end;
    4:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetDWORDValue', InParam);  // [out] uint32 uValue
    result:=inttohex(integer(OutParam.uValue),8)+' ('+vartostr(OutParam.uValue)+')';
    end;
    7:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetMultiStringValue', InParam); //[out] string sValue[]
    for i := VarArrayLowBound(OutParam.sValue,1) to VarArrayHighBound(OutParam.sValue,1) do
    result:=result+vartostr(OutParam.sValue[i])+#13#10;
    end;
    8:begin
      OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetBinaryValue', InParam); //[out] string sValue
      if VarArrayDimCount(OutParam.uValue)<>0 then
        begin
        for i := VarArrayLowBound(OutParam.uValue,1) to VarArrayHighBound(OutParam.uValue,1) do
        result:=result+vartostr(OutParam.uValue[i]);
        end;
      end;
    11:begin
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetQWORDValue', InParam); //[out] uint64 uValue
    result:=inttohex(int64(OutParam.uValue),8)+' ('+vartostr(OutParam.uValue)+')';
    end;
    end;
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    except
    on E: Exception do
    begin
    result:='Ошибка чтения значений';
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    Memolog.lines.add(Format(' - Ошибка чтения значений ключей (27) : %s ',[E.Message]));
    end;
  end;
  end;

function TRegEdit.UpdateKeyTreeItem(hDefKey:integer;sSubKeyName:string):boolean;
Var                            //чтение ключей раздела
  FWbemObjectSet: OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
  const
  TypeKey: array [0..11] of string=('REG_NONE','REG_SZ','REG_EXPAND_SZ','REG_BINARY'
  ,'REG_DWORD','REG_DWORD_BIG_ENDIAN','REG_LINK','REG_MULTI_SZ','REG_RESOURCE_LIST'
  ,'REG_FULL_RESOURCE_DESCRIPTION','REG_RESSOURCE_REQUIREMENT_MAP','REG_QWORD');
begin
    FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
    LVkey.clear;
    try
    LVKey.Items.BeginUpdate;
    StatusBar1.Panels[0].Text:='Чтение ключей раздела '+sSubKeyName;
     InParam:= FWbemObjectSet.Methods_.Item('EnumValues').InParameters.SpawnInstance_();
     InParam.hDefKey:=hDefKey;
     InParam.sSubKeyName:=sSubKeyName; // путь в разделе
     OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'EnumValues', InParam);
     if VarArrayDimCount(OutParam.sNames)<>0 then // если массив не пустой
     for i:= VarArrayLowBound(OutParam.sNames,1) to VarArrayHighBound(OutParam.sNames,1) do
     begin
     if OutParam.sNames[i]<>null then
      with LVKey.Items.Add do
       begin
       ImageIndex:=integer(OutParam.Types[i]);
       if OutParam.sNames[i]<>null then caption:=VarToStr(OutParam.sNames[i])
       else Caption:=('По умолчанию');
       SubItems.add(TypeKey[integer(OutParam.Types[i])]);
       //MemoLog.Lines.Add(inttostr(hDefKey)+' - '+vartostr(OutParam.Types[i])+' - '+sSubKeyName+' - '+VarToStr(OutParam.sNames[i]));
       SubItems.add(readValueKey(hDefKey,integer(OutParam.Types[i]),sSubKeyName,VarToStr(OutParam.sNames[i])));
       end;
     end;
     LVKey.Items.EndUpdate;
     VariantClear(OutParam);
     VariantClear(InParam);
     VariantClear(FWbemObjectSet);
     StatusBar1.Panels[0].Text:='';
    except
    on E: Exception do
    begin
    LVKey.Items.EndUpdate;
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
   StatusBar1.Panels[0].Text:='';
    Memolog.lines.add(Format(' Ошибка чтения ключей (28) : %s ',[E.Message]));
    end;
  end;
  end;



  /////////////////////////////////////////////////////////////////////
  ///  /////////////////////////////////////////////////////////////////////
 function GetRootRazdel(sSubKeyName:string):string;
 var       //System\CurrentControlSet\Control\Print\Printers\Fax (перенаправлено 2)
 s:string;
 z,i:integer;
 begin
 for I := length(sSubKeyName) downto 1 do
 if sSubKeyName[i]='\' then
 begin
   z:=i;
   Break;
 end;
 result:=copy(sSubKeyName,1,z-1);
 end;


 function CurNameKey(sSubKeyName:string):string;
 var
 s:string;
 z,i:integer;
 begin
 for I := length(sSubKeyName) downto 1 do
 if sSubKeyName[i]='\' then
 begin
   z:=i;
   Break;
 end; 
 result:=copy(sSubKeyName,z+1,length(sSubKeyName));
 end;

  ///  /////////////////////////////////////////////////////////////////////////


function TRegEdit.RenameRazdel(hDefKey:integer;sSubKeyName,NewsSubKeyName:string):boolean;  // переименование раздела
  Var
  FWbemObjectSet: OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
  CurrentNameKey:string;
  step:integer;
  begin
   try
     FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
     CurrentNameKey:=CurNameKey(sSubKeyName);
     StatusBar1.Panels[0].Text:='Переименование раздела реестра '+CurrentNameKey;
     InParam:= FWbemObjectSet.Methods_.Item('EnumKey').InParameters.SpawnInstance_();
     InParam.hDefKey:=hDefKey;
     InParam.sSubKeyName:=GetRootRazdel(sSubKeyName); //сюда передаем корневой путь для переименовываемого раздела
     OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'EnumKey', InParam);
     if VarArrayDimCount(OutParam.sNames)<>0 then  // если массив не пустой
     for i:= VarArrayLowBound(OutParam.sNames,1) to VarArrayHighBound(OutParam.sNames,1) do
     begin
     if OutParam.sNames[i]=CurrentNameKey then
     begin
     OutParam.sNames[i]:=NewsSubKeyName;
     break;
     end;
     end;
       VariantClear(OutParam);
       VariantClear(InParam);
       VariantClear(FWbemObjectSet);
       StatusBar1.Panels[0].Text:='';
    except
    on E: Exception do
    begin
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка переименования раздела (29) : %s ',[E.Message]));
    end;
  end;
end;

function TRegEdit.JumpTreeView(RootKey,Jpath:string):boolean;
var
z:integer;
RootNode:TTreeNode;
Path:string;
NameNode:TstringList;
begin
try
try
result:=false;
NameNode:=TStringList.Create;
NameNode.Add(RootKey); // добавляем в лист корневой раздел
while jpath<>'' do  // теперь заполняем стринглист
begin
if pos('\',Jpath)<>0 then z:=pos('\',Jpath)-1
else z:=Length(Jpath);
NameNode.Add(copy(Jpath,1,z));
delete(Jpath,1,z+1);
end;

RootNode:=TreeViewFolders.Items.Item[0];
 while NameNode.Count<>0 do
 begin
 for z := 0 to RootNode.Count-1 do
 begin
 if AnsiUpperCase(RootNode.Item[z].Text)=AnsiUpperCase(NameNode[0]) then  // сравниваем значения в ВЕРХНЕМ регистре
   begin
   RootNode:=RootNode.Item[z];
    if RootNode.Count=0 then  // если узел пустой
     begin
     path:=GetNodeFullPath(RootNode);
     RootNode:=JumpFolderTreeItem(RootNode,StringInGetRooT(EditRooT.Text),path);
     end;
   RootNode.Expand(true); // раскрыть узел
   break;
   end;
 end;
 NameNode.Delete(0);
 end;
UpdateKeyTreeItem(GetNodeRooT(RootNode),GetNodeFullPath(RootNode));
TreeViewFolders.Selected:=RootNode;
TreeViewFolders.SetFocus;
result:=true;
finally
NameNode.Free;
end;
except on E: Exception do memolog.Lines.Add('Jump Произошла непредвиденная ошибка (30) : '+e.Message);
end;
end;



procedure TRegEdit.SpeedButton1Click(Sender: TObject);
begin
 if EditRooT.Text='' then
 begin
   ShowMessage('Выберите корневой раздел');
   EditRooT.SetFocus;
   EditRooT.DroppedDown:=true;
   exit;
 end;
 JumpTreeView(EditRooT.Text,EditPath.Text);
end;



procedure TRegEdit.N1Click(Sender: TObject); // копировать ключи
var
sValueNames:TstringList;
i:integer;

function findKey(sValueName:string):boolean;
begin
result:=false;
if sValueName='REG_SZ' then result:=true;
if sValueName='REG_EXPAND_SZ' then result:=true;
if sValueName='REG_BINARY' then result:=true;
if sValueName='REG_DWORD' then result:=true;
if sValueName='REG_MULTI_SZ' then result:=true;
if sValueName='REG_QWORD' then result:=true;
end;

begin
if LVKey.SelCount=0 then exit; // если не выделено то валим нах
sValueNames:=TstringList.create;
try
for I := 0 to LVKey.Items.Count-1 do
  begin
  if (LVKey.Items[i].Selected) and (findKey(LVKey.Items[i].SubItems[0])) then // если ключ выделен и тип соответствует разрешенному для копирования
   begin
   sValueNames.Add(LVKey.Items[i].Caption+'='+LVKey.Items[i].SubItems[0]); // имя ключа и его тип
   end;
  end;
CopyKey(StringInGetRooT(EditRooT.Text),EditPath.Text,sValueNames);
finally
sValueNames.Free;
end;
end;

procedure TRegEdit.N8Click(Sender: TObject); // вставить ключи
begin
try
// если переменная для копирования и вставки пустая то предупреждаем
if (not VarIsArray(CopyPasteVariant)) or (VarIsNull(CopyPasteVariant)) then showmessage('Буфер пустой')
else //иначе производим вставку
begin
 if PasteKey(StringInGetRooT(EditRooT.Text),EditPath.Text,CopyPasteVariant) then
 if (VarIsArray(CopyPasteVariant)) or (not VarIsNull(CopyPasteVariant)) then CopyPasteVariant:=Unassigned; // после вставки производим очистку переменной
 UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text); // обновляем директорию куда произвели вставку
 end;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (18): '+e.Message);
end;
end;

function TRegEdit.CreatNewRazdel(hDefKey:integer;sSubKeyName:string):integer;  // создает раздел в выеленном подразделе
Var
FWbemObjectSet: OLEVariant;
res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание раздела реестра '+sSubKeyName;
  res:=FWbemObjectSet.CreateKey(hDefKey,sSubKeyName);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Создание раздела '+sSubKeyName+': '+SysErrorMessage(res));
  result:=res;
  except
    on E: Exception do
    begin
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка создания раздела реестра (31): %s ',[E.Message]));
    end;
  end;
end;

function TRegEdit.DeleteRazdel(hDefKey:integer;sSubKeyName:string):integer;  // удаляет раздел в выеленном подразделе
Var
  FWbemObjectSet: OLEVariant;
  res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Удаление раздела реестра '+sSubKeyName;
  res:=FWbemObjectSet.DeleteKey(hDefKey,sSubKeyName);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  result:=res;
  except
    on E: Exception do
    begin
     result:=1111;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка удаления раздела реестра (32) : %s ',[E.Message]));
    end;
  end;
end;

procedure TRegEdit.EditPathKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key=VK_RETURN then
begin
   if EditRooT.Text='' then
 begin
   ShowMessage('Выберите корневой раздел');
   EditRooT.SetFocus;
   EditRooT.DroppedDown:=true;
   exit;
 end;
 JumpTreeView(EditRooT.Text,EditPath.Text);
end;
end;

function TRegEdit.DeleteKeyValue(hDefKey:integer;sSubKeyName:string;sValueName:TstringList):integer;  // удаляет ключ в реестре
Var
  FWbemObjectSet: OLEVariant;
  InParam: OLEVariant;
  i:integer;
  step,res:integer;
begin
 try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Удаление ключа реестра';
  for I := 0 to sValueName.Count-1 do // в цикле удаляем ключи
  begin
   res:=FWbemObjectSet.DeleteValue(hDefKey,sSubKeyName,sValueName[i]);
  MemoLog.Lines.Add('Удаление ключа реестра '+sValueName[i]+': '+SysErrorMessage(res));
  end;
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  result:=res;
  except
    on E: Exception do
    begin
     result:=1111;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка удаления ключа реестра (33): %s ',[E.Message]));
    end;
  end;
end;

procedure TRegEdit.C1Click(Sender: TObject);
var
Newvalue:string;
CurentRoot:integer;
CurentPath:string;
lNode:TTreeNode;
begin
try
if not InputQuery('Создание раздела реестра', 'Имя раздела:', Newvalue)
    then exit;
CurentRoot:=GetNodeRooT(TreeViewFolders.Selected);
if CurentRoot=0 then exit;  // если не выбран корневой раздел
CurentPath:=GetNodeFullPath(TreeViewFolders.Selected);;
if CreatNewRazdel(CurentRoot,CurentPath+'\'+Newvalue)=0 then
begin  // если успешно создали раздел то добавляем его в дерево
 lNode:=TreeViewFolders.Items.AddChild(TreeViewFolders.Selected,Newvalue);
 lNode.ImageIndex:=0;
 lNode.SelectedIndex:=0;
end;
except on E: Exception do MemoLog.Lines.Add('Ошибка создания раздела (34) '+E.Message);
end;
end;

procedure TRegEdit.N25Click(Sender: TObject); // редактирование ключа реестра
begin
if LVKey.SelCount<>1 then exit;
CreateFormForKey(LVKey.Selected.SubItems[0],'',LVKey.Selected.Caption,LVKey.Selected.SubItems[1]);
end;

procedure TRegEdit.N2Click(Sender: TObject);
var
CurentRoot,i,res:integer;
CurentPath:string;
lNode:TTreeNode;
begin
try
CurentRoot:=GetNodeRooT(TreeViewFolders.Selected);
if CurentRoot=0 then exit;  // если не выбран корневой раздел
CurentPath:=GetNodeFullPath(TreeViewFolders.Selected);
i:=MessageDlg('Удалить раздел реестра и его содержимое? '+#10#13+CurentPath, mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
res:=DeleteRazdel(CurentRoot,CurentPath);
if res=0 then TreeViewFolders.Selected.Delete  // если успешно удалили раздел
else MemoLog.Lines.Add(SysErrorMessage(res));
except on E:EOleException do
        MemoLog.Lines.Add(Format('Ошибка EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do begin
        MemoLog.Lines.Add('Ошибка (35): '+E.Message);  end;
end;
end;


procedure TRegEdit.N4Click(Sender: TObject);
var
ListDel:Tstringlist;
i:integer;
begin
if LVKey.SelCount=0 then exit;
i:=MessageDlg('Удалить выделенные ключи реестра?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
ListDel:=Tstringlist.Create;
try
for I := 0 to LVKey.Items.Count-1 do
begin
  if LVKey.Items[i].Selected then
  ListDel.Add(LVKey.Items[i].Caption);
end;
if DeleteKeyValue(StringInGetRooT(EditRooT.Text),EditPath.Text,ListDel)=0 then // если удачно создали то обновляем список ключей раздела
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
finally
ListDel.Free;
end;
end;

procedure TRegEdit.N5Click(Sender: TObject);  // обновить список ключей
begin
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;

function TRegEdit.CreateSetStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;
var                   // строковый параметр
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sSubKeyName;
  res:=FWbemObjectSet.SetStringValue(hDefKey,sSubKeyName,sValueName,sValue);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка создания ключа реестра (37): %s ',[E.Message]));
    end;
  end;
  end;

function TRegEdit.CreateSetBinaryValue(hDefKey:integer;sSubKeyName,sValueName,uValue:string):integer;
var                     //новый бинарный ключ
FWbemObjectSet: OLEVariant;
i:integer;
res:integer;
mas:variant;
begin
try
   mas:=VarArrayCreate([0,length(uValue)],varByte);
   for I := 1 to Length(uValue) do
   mas[i-1]:=strtoint(uValue[i]);
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sSubKeyName;
  res:=FWbemObjectSet.SetBinaryValue(hDefKey,sSubKeyName,sValueName,mas);
  VariantClear(FWbemObjectSet);
  VarClear(mas);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
    VariantClear(FWbemObjectSet);
    VarClear(mas);
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(' Ошибка создания ключа реестра (38):  '+E.Message);
    end;
  end;
  end;

function TRegEdit.CreateSetDWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:integer):integer;
var               // новый DWORD (int32) ключ
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sSubKeyName;
  res:=FWbemObjectSet.SetDWORDValue(hDefKey,sSubKeyName,sValueName,uValue);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
     VariantClear(FWbemObjectSet);
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка создания ключа реестра (39): %s ',[E.Message]));
    end;
  end;
  end;

function TRegEdit.CreateSetQWORDValue(hDefKey:integer;sSubKeyName,sValueName:string;uValue:int64):integer;
var               // новый QWORD (int64) ключ
FWbemObjectSet: OLEVariant;
InParam: OLEVariant;
i:integer;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sValueName;
  InParam:= FWbemObjectSet.Methods_.Item('SetQWORDValue').InParameters.SpawnInstance_();
  InParam.hDefKey:=hDefKey;  // root
  InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь для создаваемого раздела
  InParam.sValueName:=sValueName;
  InParam.uValue:=uValue;
  FWMIServiceRE.ExecMethod('StdRegProv', 'SetQWORDValue', InParam);
  //res:=FWbemObjectSet.SetQWORDValue(hDefKey,sSubKeyName,sValueName,uValue); так не понимет входящий параметрт uValue, просит integer хотя должен принимать int64
  VariantClear(InParam);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  res:=0;
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
     //VariantClear(InParam);
     VariantClear(FWbemObjectSet);
     result:=1111;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка создания ключа реестра (40): %s ',[E.Message]));
    end;
  end;
  end;

function TRegEdit.CreateSetMultiStringValue(hDefKey:integer;sSubKeyName,sValueName:string;sValue:TStringList):integer;
var                     //новый Мультистринговый ключ
FWbemObjectSet: OLEVariant;
i:integer;
res:integer;
mas:variant;
begin
try
   mas:=VarArrayCreate([0,sValue.count-1],varOleStr);
   for I := 0 to sValue.count-1 do mas[i]:=sValue[i];
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sSubKeyName;
  res:=FWbemObjectSet.SetMultiStringValue(hDefKey,sSubKeyName,sValueName,mas);
  VariantClear(FWbemObjectSet);
  VarClear(mas);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
     VariantClear(FWbemObjectSet);
     VarClear(mas);
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(' Ошибка создания ключа реестра (41):  '+E.Message);
    end;
  end;
  end;

function TRegEdit.CreateSetExpandedStringValue(hDefKey:integer;sSubKeyName,sValueName,sValue:string):integer;
var                   // расширенный строковый параметр
FWbemObjectSet: OLEVariant;
res:integer;
begin
try
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Создание ключа реестра '+sSubKeyName;
  res:=FWbemObjectSet.SetExpandedStringValue(hDefKey,sSubKeyName,sValueName,sValue);
  VariantClear(FWbemObjectSet);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(res));
  result:=res;
except
    on E: Exception do
    begin
     result:=res;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка создания ключа реестра (42): %s ',[E.Message]));
    end;
  end;
  end;

procedure TRegEdit.FormClose(Sender: TObject; var Action: TCloseAction);
begin
try
VariantClear(FSWbemLocatorRE);
VariantClear(FWMIServiceRE);
SaveFavorite; // сохраняем список избранного в файл
Favoritetxt.Free; // уничтожаем список с избранным
cleanMainMenu; // чистка основного меню избранного
LVkey.Clear;  // чистим
MemoLog.Clear;   // чистим
EditPath.Text:=''; // чистим
EditRooT.Text:=''; // чистим
if (VarIsArray(CopyPasteVariant)) or (not VarIsNull(CopyPasteVariant)) then CopyPasteVariant:=Unassigned; // производим очистку переменной для копирования и вставки клчей
OleUnInitialize;
except on E: Exception do  Memolog.lines.add(Format('Ошибка (43) : %s ',[E.Message]));
end;
end;

procedure TRegEdit.N6Click(Sender: TObject); // создает новый строковый параметр
var
z:integer;
NewName:string;
begin
while NewName='' do
begin
if not InputQuery('Новый параметр REG_SZ', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
if CreateSetStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,'')=0then // если удачно создали то обновляем список ключей раздела
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;

procedure TRegEdit.REGBINARY1Click(Sender: TObject); /// Создает новый бинарный параметр
var
z:integer;
NewName:string;
begin
while NewName='' do
begin
if not InputQuery('Новый параметр REG_BINARY', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
if CreateSetBinaryValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,'0')=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;



procedure TRegEdit.REGDWORD1Click(Sender: TObject);   /// Создает новый числовой (uint32) параметр
var
z:integer;
NewName:string;
begin
while NewName='' do
begin
if not InputQuery('Новый параметр REG_DWORD', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
if CreateSetDWORDValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,0)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;

procedure TRegEdit.REGQWORD1Click(Sender: TObject); /// Создает новый числовой (uint64) параметр
var
z:integer;
NewName:string;
begin
while NewName='' do
begin
if not InputQuery('Новый параметр REG_QWORD', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
if CreateSetQWORDValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,3)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;

procedure TRegEdit.REGMULTISZ1Click(Sender: TObject); // создает новый мультистроковый ключ
var
z:integer;
NewName:string;
ValueS:TstringList;
begin
ValueS := TstringList.Create;
try
while NewName='' do
begin
if not InputQuery('Новый параметр REG_MULTI_SZ', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
ValueS.Add(''); // добавляем пустое значение для передачи в SetMultiStringValue
if CreateSetMultiStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,ValueS)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
finally
ValueS.Free;
end;
end;

procedure TRegEdit.N7Click(Sender: TObject); // создает расширенный строковый параметр
var
z:integer;
NewName:string;
begin
while NewName='' do
begin
if not InputQuery('Новый параметр REG_EXPAND_SZ', 'Имя параметра:', NewName)
    then exit;
if NewName='' then ShowMessage('Вы не указали имя нового параметра');
end;
if CreateSetExpandedStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,NewName,'')=0 then // если удачно создали то обновляем список ключей раздела
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
end;



procedure TRegEdit.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;

procedure TRegEdit.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;

procedure TRegEdit.FormSaveKeyClose(Sender: TObject);
begin
FormSaveKey.Close;
end;

procedure TRegEdit.ButtonOkREG_SZ(Sender: TObject); // сохраняет изменения в строковом параметре
var
i:integer;
begin
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
if CreateSetStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,EditKey.Text)=0 then // если удачно создали то обновляем список ключей раздела
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
end;

procedure TRegEdit.ButtonOkREG_EXPAND_SZ(Sender: TObject); // сохраняет изменения в расширенном строковом параметре
var
i:integer;
begin
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
if CreateSetExpandedStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,EditKey.Text)=0 then // если удачно создали то обновляем список ключей раздела
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
end;

procedure TRegEdit.ButtonOkREG_Binary(Sender: TObject); // сохраняет изменения в binary параметре
var
i:integer;
begin
try
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
if newCreateBinary(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,MemoKey.Text,RadioGr.Caption)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
except
    on E: Exception do  MemoLog.Lines.Add('Ошибка сохранения ключа (44): '+e.Message);
end;
end;

procedure TRegEdit.ButtonOkREG_QWORD(Sender: TObject); // сохраняет изменения в QWORD параметре
var
i:integer;
z:int64;
begin
z:=0;
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
if RadioGr.Caption='INT' then
if not TryStrToInt64(EditKey.Text,z)then
begin
ShowMessage('Не верное значение параметра');
exit;
end;
if RadioGr.Caption='HEX' then
if not TryStrToInt64('$'+EditKey.Text,z)then
begin
ShowMessage('Не верное значение параметра');
exit;
end;
if CreateSetQWORDValue(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,z)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
end;

procedure TRegEdit.ButtonOkREG_DWORD(Sender: TObject); // сохраняет изменения в DWORD параметре
var
i,z:integer;
begin
z:=0;
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
if RadioGr.Caption='INT' then
if not TryStrToInt(EditKey.Text,z)then
begin
ShowMessage('Не верное значение параметра');
exit;
end;
if RadioGr.Caption='HEX' then
if not TryStrToInt('$'+EditKey.Text,z)then
begin
ShowMessage('Не верное значение параметра');
exit;
end;
if CreateSetDWORDValue(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,z)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
end;

procedure TRegEdit.ButtonOkMULTI_SZ(Sender: TObject); // сохраняет изменения мультистрокового параметре
var
i:integer;
ListStr:TstringList;
begin
i:=MessageDlg('Сохранить изменения параметра '+EditNameKey.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
begin
FormKey.Close;
exit;
end;
ListStr :=TstringList.Create;
try
for I := 0 to MemoKey.Lines.Count-1 do ListStr.Add(MemoKey.Lines[i]);
if CreateSetMultiStringValue(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text,ListStr)=0 then
UpdateKeyTreeItem(StringInGetRooT(EditRooT.Text),EditPath.Text);
FormKey.Close;
finally
ListStr.Free;
end;
end;

function TRegEdit.CreateBinarySaveKey(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // импортирование бинарного ключа в реестре
var
FWbemObjectSet: OLEVariant;
//InParam: OLEVariant;
i,z,OutParam:integer;
mas:variant;
res:TstringList;
newstr:string;
function MemoHexToInt(FullStr:string):string;
var
i,z:integer;res:String;newstr:string;
begin
try
res:=''; i:=0;
while FullStr<>'' do
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    inc(i);
    newstr:=copy(FullStr,1,2);
    if TryStrToInt('$'+newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res:=res+inttostr(z)+' ';
    delete(FullStr,1,2);
    end;
  end;
result:=res;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (45): '+e.Message);
end;
end;
begin
try
res:=TStringList.Create;
i:=0;
if HexOrInt='HEX' then // если цисла были 16ти ричные то получаем чистый integer
begin
FullStr:=MemoHexToInt(FullStr);
end;
if FullStr[length(FullStr)]<>' ' then // если в конце строки нет пробела то добавляем его
Fullstr:=Fullstr+' ';
while FullStr<>'' do /// при условии
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    newstr:=copy(FullStr,1,pos(' ',FullStr)-1);
    if TryStrToInt(newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res.Add(inttostr(z));
    delete(FullStr,1,pos(' ',FullStr));
    end;
  end;
   mas:=VarArrayCreate([0,res.Count-1],varByte);
   for I := 0 to res.Count-1 do
   mas[i]:=strtoint(res[i]);
  res.Free;
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Сохранение ключа реестра '+sSubKeyName;
  //InParam:= FWbemObjectSet.Methods_.Item('SetBinaryValue').InParameters.SpawnInstance_();
  //InParam.hDefKey:=hDefKey;  // root
  //InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь для создаваемого раздела
  //InParam.sValueName:=sValueName;
  //InParam.uValue:=mas;
  //FWMIServiceRE.ExecMethod('StdRegProv', 'SetBinaryValue', InParam);
  //VariantClear(InParam);
  outparam:=FWbemObjectSet.SetBinaryValue(hDefKey,sSubKeyName,sValueName,mas);
  VariantClear(FWbemObjectSet);
  VarClear(mas);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(OutParam));
  result:=outparam;
except
    on E: Exception do
    begin
    //VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    VarClear(mas);
     result:=outparam;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(' Ошибка сохранения ключа реестра (46):  '+E.Message);
    end;
  end;
  end;


function TRegEdit.newCreateBinary(hDefKey:integer;sSubKeyName,sValueName,FullStr,HexOrInt:string):integer; // сохранение отредактированного бинарного ключа в реестре
var
FWbemObjectSet: OLEVariant;
//InParam: OLEVariant;
i,z,Outparam:integer;
mas:variant;
res:TstringList;
newstr:string;
begin
try
res:=TStringList.Create;
i:=0;
if HexOrInt='HEX' then // если цисла были 16ти ричные то получаем чистый integer
begin
FullStr:=MemoHexToInt(FullStr);
end;
while (FullStr[length(FullStr)]=#10) or (FullStr[length(FullStr)]=#13) do delete(FullStr,length(FullStr),1); // если в конце строки символы переноса строки то удаляем их
if FullStr[length(FullStr)]<>' ' then // если в конце строки нет пробела то добавляем его
Fullstr:=Fullstr+' ';
while FullStr<>'' do /// при условии
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    inc(i);
    if i mod 9 <>0 then //newstr:=copy(FullStr,1,2)
     newstr:=copy(FullStr,1,pos(' ',FullStr)-1);
    if i mod 9 <>0 then   // каждая строка из Memo заканчивается переводом каретки, #13#10 а это доп символы. оэтому пробелы сдесь не ставим
    if TryStrToInt(newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res.Add(inttostr(z));
    if i mod 9 =0 then delete(FullStr,1,2) // если строка закончилась то удаляем два символа - #13#10 (признак новой строки)
    else delete(FullStr,1,pos(' ',FullStr));
    end;
  end;
   mas:=VarArrayCreate([0,res.Count-1],varByte);
   for I := 0 to res.Count-1 do
   mas[i]:=strtoint(res[i]);
  res.Free;
  FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
  StatusBar1.Panels[0].Text:='Сохранение ключа реестра '+sSubKeyName;
 // InParam:= FWbemObjectSet.Methods_.Item('SetBinaryValue').InParameters.SpawnInstance_();
 // InParam.hDefKey:=hDefKey;  // root
 // InParam.sSubKeyName:=sSubKeyName; //сюда передаем полный путь для создаваемого раздела
 // InParam.sValueName:=sValueName;
 // InParam.uValue:=mas;
//  FWMIServiceRE.ExecMethod('StdRegProv', 'SetBinaryValue', InParam);
 // VariantClear(InParam);
  outparam:=FWbemObjectSet.SetBinaryValue(hDefKey,sSubKeyName,sValueName,mas);
  VariantClear(FWbemObjectSet);
  VarClear(mas);
  StatusBar1.Panels[0].Text:='';
  MemoLog.Lines.Add('Сохранение ключа реестра '+sValueName+': '+SysErrorMessage(OutParam));
  result:=OutParam;
except
    on E: Exception do
    begin
   // VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    VarClear(mas);
     result:=OutParam;
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(' Ошибка сохранения ключа реестра (47):  '+E.Message);
    end;
  end;
  end;



function TRegEdit.readBintoHextoInt(HexOrInt:string):boolean;
Var  //чтение значений ключей
  FWbemObjectSet,FWMIService     : OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
  step:integer;
 begin
    step:=0;
    FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
    try
    step:=1;
    InParam:= FWbemObjectSet.Methods_.Item('GetBinaryValue').InParameters.SpawnInstance_(); //Получает значение двоичных данных именованного значения.
    step:=2;
    InParam.hDefKey:=StringInGetRooT(EditRooT.Text);
    InParam.sSubKeyName:=EditPath.Text; // путь в разделе
    InParam.sValueName:=EditNameKey.Text; // имя параметра передается в Edit при создании формы
    step:=3;
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetBinaryValue', InParam);  //uValue [out] An array of binary bytes.
    MemoKey.Clear;
    if HexOrInt='HEX' then
    begin
    for i := VarArrayLowBound(OutParam.uValue,1) to VarArrayHighBound(OutParam.uValue,1) do
      begin
      if i mod 8 =0 then MemoKey.Lines.Add('');
      memokey.Lines[MemoKey.Lines.Count-1]:=memokey.Lines[MemoKey.Lines.Count-1]+intToHex(integer(OutParam.uValue[i]),2)+' ';
      end;
    end;
    if HexOrInt='INT' then
    begin
    for i := VarArrayLowBound(OutParam.uValue,1) to VarArrayHighBound(OutParam.uValue,1) do
      begin
      if i mod 8 =0 then MemoKey.Lines.Add('');
      memokey.Lines[MemoKey.Lines.Count-1]:=memokey.Lines[MemoKey.Lines.Count-1]+inttostr(OutParam.uValue[i])+' ';
      end;
    end;
    step:=4;
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    result:=true;
    except
    on E: Exception do
    begin
    result:=false;
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    Memolog.lines.add(Format(inttostr(step)+' - Ошибка чтения значений REG_BINARY (48): %s ',[E.Message]));
    end;
  end;
end;

function TRegEdit.read_REG_MULTI_SZ(hDefKey:integer;sSubKeyName,sValueName:string):TstringList;
Var  //чтение значений ключей
  FWbemObjectSet,FWMIService     : OLEVariant;
  InParam,OutParam   : OLEVariant;
  i:integer;
 begin
    result:=TStringList.Create;
    FWbemObjectSet:= FWMIServiceRE.Get('StdRegProv');
    try
    InParam:= FWbemObjectSet.Methods_.Item('GetMultiStringValue').InParameters.SpawnInstance_(); //Получает значение двоичных данных именованного значения.
    InParam.hDefKey:=hDefKey;
    InParam.sSubKeyName:=sSubKeyName; // путь в разделе
    InParam.sValueName:=sValueName; // имя параметра
    OutParam:= FWMIServiceRE.ExecMethod('StdRegProv', 'GetMultiStringValue', InParam);  //uValue [out] An array of binary bytes.
    for i := VarArrayLowBound(OutParam.sValue,1) to VarArrayHighBound(OutParam.sValue,1) do
    result.add(vartostr(OutParam.sValue[i]));
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    except
    on E: Exception do
    begin
    VariantClear(OutParam);
    VariantClear(InParam);
    VariantClear(FWbemObjectSet);
    Memolog.lines.add(Format('Ошибка чтения значений REG_MULTI_SZ (49): %s ',[E.Message]));
    end;
  end;
end;

procedure TRegEdit.FormKeyShowBinary(Sender: TObject); // дествия для REG_BINARY при показк формы для редактирования
begin
readBintoHextoInt('HEX');
RadioGr.Caption:='HEX';
RadioGr.ItemIndex:=0;
end;

procedure TRegEdit.FormKeyShowMULTI_SZ(Sender: TObject); // дествия для REG_MULTI_SZ при показк формы для редактирования
var
resSrtList:Tstringlist;
i:integer;
begin
resSrtList:= Tstringlist.Create;
try
MemoKey.Clear;
resSrtList:=read_REG_MULTI_SZ(StringInGetRooT(EditRooT.Text),EditPath.Text,EditNameKey.Text);
for I := 0 to resSrtList.Count-1 do
MemoKey.Lines.Add(resSrtList[i]);
finally
 resSrtList.Free;
end;
end;

function TRegEdit.MemoInttohex(FullStr:string):string;
var
i,z:integer;
res:String;
newstr:string;
begin
try
res:='';
i:=0;
while FullStr<>'' do
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    inc(i);
    if i mod 9 =0 then newstr:=copy(FullStr,1,2)
    else
    newstr:=copy(FullStr,1,pos(' ',FullStr)-1);
    if i mod 9 =0 then res:=res+newstr   // каждая строка из Memo заканчивается переводом каретки, #13#10 а это доп символы. оэтому пробелы сдесь не ставим
    else
    if TryStrToInt(newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res:=res+IntToHex(z,2)+' ';
    //RegEdit.MemoLog.Lines.Add(res);
    if i mod 9 =0 then delete(FullStr,1,2)
    else
    delete(FullStr,1,pos(' ',FullStr));
    end;
  end;
result:=res;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (50): '+e.Message);
end;
end;

function TRegEdit.MemoHexToInt(FullStr:string):string;
var
i,z:integer;
res:String;
newstr:string;
begin
try
res:='';
i:=0;
while FullStr<>'' do
  begin
  newstr:='';
  if pos(' ',FullStr)=1 then delete(FullStr,1,1);
  if FullStr<>'' then
    begin
    inc(i);
    newstr:=copy(FullStr,1,2);
    if i mod 9 =0 then res:=res+newstr   // каждая строка из Memo заканчивается переводом каретки, #13#10 а это доп символы. оэтому пробелы сдесь не ставим
    else
    if TryStrToInt('$'+newstr,z) then // если  newstr является числом то заносится в z и соотетственно в результат
    res:=res+inttostr(z)+' ';
   // RegEdit.MemoLog.Lines.Add(res);
    delete(FullStr,1,2);
    end;
  end;
result:=res;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (51): '+e.Message);
end;
end;

procedure TRegEdit.RadioGrClick(Sender: TObject); // изменения формата представления Бинарных данных
Var
ans:string;
begin
ans:='';
try
case RadioGr.ItemIndex of
0: // HEX
  begin
  if RadioGr.Caption='HEX' then exit;
  if RadioGr.Caption='INT' then
  begin
   RadioGr.Caption:='HEX';
   ans:=MemoInttohex(MemoKey.Text);
   MemoKey.Text:=ans;
  end;
  end;
1: // INT
  begin
  if RadioGr.Caption='INT' then exit;
  if RadioGr.Caption='HEX' then
  begin
   RadioGr.Caption:='INT';
   ans:=MemoHexToInt(MemoKey.Text);
   MemoKey.Text:=ans;
  end;
  end;
end;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (52): '+e.Message);
end;
end;

procedure TRegEdit.RadioGrQDClick(Sender: TObject); // изменения формата представления QWORD и DWORD данных
Var
ans:int64;
begin
try
case RadioGr.ItemIndex of
0: // HEX
  begin
  if RadioGr.Caption='HEX' then exit;
  if RadioGr.Caption='INT' then
  begin
   RadioGr.Caption:='HEX';
   if TryStrToInt64(EditKey.Text,ans) then EditKey.Text:=IntToHex(ans,8);
  end;
  end;
1: // INT
  begin
  if RadioGr.Caption='INT' then exit;
  if RadioGr.Caption='HEX' then
  begin
   RadioGr.Caption:='INT';
   if TryStrToInt64('$'+EditKey.Text,ans) then EditKey.Text:=IntToStr(ans);
  end;
  end;
end;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (53): '+e.Message);
end;
end;

//'REG_NONE','REG_SZ','REG_EXPAND_SZ','REG_BINARY','REG_DWORD','REG_DWORD_BIG_ENDIAN','REG_LINK','REG_MULTI_SZ','REG_RESOURCE_LIST','REG_FULL_RESOURCE_DESCRIPTION','REG_RESSOURCE_REQUIREMENT_MAP','REG_QWORD'
function TRegEdit.CreateFormForKey(TypeKey,Description,sValueName,sValue:string):boolean;
begin     // создает форму для просмотра и редактироввания значений ключей реестра
try
if (TypeKey='REG_SZ') or (TypeKey='REG_EXPAND_SZ') then
begin
FormKey:=TForm.Create(RegEdit);
FormKey.Caption:='Изменение строкового параметра';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=365;
FormKey.Height:=155;
FormKey.OnClose:= FormKeyClose;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Параметр:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=340;
EditNameKey.Text:=sValueName;
//////////////////////
EditKey:=TLabeledEdit.Create(FormKey);
EditKey.Parent:=FormKey;
EditKey.EditLabel.Caption:='Значение:';
EditKey.Top:=60;
EditKey.Left:=5;
EditKey.Width:=340;
EditKey.Text:=sValue;
EditKey.TabOrder:=0; // переводим фокус в поле со значением параметра
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=90;
ButtonOk.Left:=170;
if TypeKey='REG_SZ' then ButtonOk.OnClick:=ButtonOkREG_SZ;
if TypeKey='REG_EXPAND_SZ' then ButtonOk.OnClick:=ButtonOkREG_EXPAND_SZ;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=90;
ButtonNo.Left:=270;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end;

if TypeKey='REG_BINARY' then
begin
FormKey:=TForm.Create(RegEdit);
FormKey.Caption:='Изменение двоичного параметра';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=345;
FormKey.OnClose:=FormKeyClose;
FormKey.OnShow:=FormKeyShowBinary;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Параметр:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:=sValueName;
//////////////////////
RadioGr:=TRadioGroup.Create(FormKey);
RadioGr.Parent:=FormKey;
RadioGr.Items.Add('HEX');
RadioGr.Items.Add('INT');
RadioGr.Top:=50;
RadioGr.Left:=5;
RadioGr.Width:=260;
RadioGr.Height:=35;
RadioGr.Columns:=2;
RadioGr.OnClick:=RadioGrClick;
/// /////////////////////
MemoKey:=Tmemo.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Top:=90;
MemoKey.Left:=5;
MemoKey.Width:=260;
MemoKey.Height:=180;
MemoKey.Text:='';
MemoKey.ScrollBars:=ssVertical;
MemoKey.TabOrder:=0; // переводим фокус в поле со значением параметра
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=277;
ButtonOk.Left:=95;
ButtonOk.OnClick:=ButtonOkREG_BINARY;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=277;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end;

if  (TypeKey='REG_DWORD')or (TypeKey='REG_QWORD')then
begin
FormKey:=TForm.Create(RegEdit);
if TypeKey='REG_DWORD' then FormKey.Caption:='Изменение параметра DWORD (32 бита)';
if TypeKey='REG_QWORD' then FormKey.Caption:='Изменение параметра QWORD (64 бита)';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=320;
FormKey.Height:=195;
FormKey.OnClose:= FormKeyClose;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Параметр:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=295;
EditNameKey.Text:=sValueName;
////////////////////////////
RadioGr:=TRadioGroup.Create(FormKey);
RadioGr.Parent:=FormKey;
RadioGr.Items.Add('HEX');
RadioGr.Items.Add('INT');
RadioGr.Caption:='HEX';
RadioGr.ItemIndex:=0;
RadioGr.Top:=45;
RadioGr.Left:=5;
RadioGr.Width:=295;
RadioGr.Height:=35;
RadioGr.Columns:=2;
RadioGr.OnClick:=RadioGrQDClick;
//////////////////////
EditKey:=TLabeledEdit.Create(FormKey);
EditKey.Parent:=FormKey;
EditKey.EditLabel.Caption:='Значение:';
EditKey.Top:=100;
EditKey.Left:=5;
EditKey.Width:=295;
EditKey.Text:=copy (sValue,1,pos(' ',sValue)-1); // т.к. передается строка, значение шестнацетиричное и в скобках десятичное
EditKey.TabOrder:=0; // переводим фокус в поле со значением параметра
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=130;
ButtonOk.Left:=125;
if TypeKey='REG_DWORD' then ButtonOk.OnClick:=ButtonOkREG_DWORD;
if TypeKey='REG_QWORD' then ButtonOk.OnClick:=ButtonOkREG_QWORD;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=130;
ButtonNo.Left:=225;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end;

if TypeKey='REG_MULTI_SZ' then
begin
FormKey:=TForm.Create(RegEdit);
FormKey.Caption:='Изменение мультистрокового параметра';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=300;
FormKey.OnClose:=FormKeyClose;
FormKey.OnShow:=FormKeyShowMULTI_SZ; // при открытии формы происходит чтение значения параметра
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Параметр:';
EditNameKey.Top:=21;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:=sValueName;
/// /////////////////////
MemoKey:=Tmemo.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Top:=50;
MemoKey.Left:=5;
MemoKey.Width:=260;
MemoKey.Height:=180;
MemoKey.Text:='';
MemoKey.ScrollBars:=ssBoth;
MemoKey.TabOrder:=0; // переводим фокус в поле со значением параметра
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=234;
ButtonOk.Left:=95;
ButtonOk.OnClick:=ButtonOkMULTI_SZ;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=234;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
end;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (54): '+e.Message);
end;
end;

end.
