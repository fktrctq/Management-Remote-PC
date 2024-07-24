unit FormForCopyPCList;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.ImgList, Vcl.Menus;

type
  TFormCopyFFFGropPC = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    LVFileFolder: TListView;
    EditPath: TLabeledEdit;
    Button1: TButton;
    CheckBox1: TCheckBox;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    ButtonCopy: TButton;
    LVPC: TListView;
    EditPaswd: TEdit;
    EditUser: TEdit;
    ImageList1: TImageList;
    Splitter1: TSplitter;
    ButtonDelete: TButton;
    Button6: TButton;
    PopupListPC: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure ButtonCopyClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure LVFileFolderKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
  private
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure LoadlistPC(Sender: TObject);
    procedure FormListPCClose(Sender: TObject; var Action: TCloseAction);
    procedure ButOkClickNamePC(Sender: TObject);
  public
    { Public declarations }
  end;

var
  FormCopyFFFGropPC: TFormCopyFFFGropPC;

implementation
uses FormCopyFF,Umain;
{$R *.dfm}

var
  FormGr:Tform;
  ListGr:Tcombobox;
  ButOk:Tbutton;

function Shell_StrSourse(Strs: TListView): string;  // Функция преобразует список TListView в спец-строку для Shell это спец-строка для буфера обмена, где строки разделены знаком #0 и вся спец-строка строка заканчивается #0#0
var
  i: Integer;
begin
  for i := 0 to Strs.items.Count - 1 do
    Result := Result + Strs.items[i].Caption + #0;
  Result := Result + #0;
end;

function Shell_StrDest(Strs: string): string;  // Функция преобразует Stirngs с разделительными символами ',' в спец-строку для Shell это спец-строка для буфера обмена, где строки разделены знаком #0 и вся спец-строка строка заканчивается #0#0
begin
    StringReplace(Strs,',',#0,[rfReplaceAll]);
    Result := Strs +#0+#0;
end;

function Shell_StrDelete(Strs: TListView): string;  // Функция преобразует список TListView в спец-строку для Shell это спец-строка для буфера обмена, где строки разделены знаком #0 и вся спец-строка строка заканчивается #0#0
var
  i: Integer;
  strcaption:string;
  ListDel:TstringList;
begin
ListDel:=TStringList.Create;
try
  for i := 0 to Strs.items.Count - 1 do
  begin
    strcaption:='';
    strcaption:=Strs.items[i].Caption;
    if AnsiPos(':',strcaption)=2 then // если в пути назначения найдено двоеточие то запрашиваем правильность ввода пути, для копирования в сеть после диска ставится $
    begin
    delete(strcaption,2,1); // удаляем символ :
    insert('$',strcaption,2); // вставляем на его место $
    end;
   ListDel.add(strcaption);
  end;
Result := ListDel.CommaText;
finally
  ListDel.Free;
end;
end;

procedure TFormCopyFFFGropPC.Button1Click(Sender: TObject);
function findElement(s:string):string;
begin
  if AnsiPos(':',s)=2 then // если в пути назначения найдено двоеточие то запрашиваем правильность ввода пути, для копирования в сеть после диска ставится $
   begin
   delete(s,2,1); // удаляем символ :
   insert('$',s,2); // вставляем на его место $
   end;
 result:=s;
end;
begin
with TFileOpenDialog.Create(self) do
begin
 try
  Name:='DlgCopyFolder';
  Title:='Выбор каталога назначения';
  Options:=[fdoForceShowHidden,fdoPickFolders]; {каталоги}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then EditPath.Text:=findElement(FileName)+'\';
 finally
 Free;
 end;
end;
end;

procedure TFormCopyFFFGropPC.Button2Click(Sender: TObject);
var
i:integer;
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор списка каталогов для копирования';
  Options:=[fdoAllowMultiSelect,fdoForceShowHidden,fdoPickFolders]; {каталоги}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     for I := 0 to Files.Count-1 do
     with LVFileFolder.Items.Add do
     caption:= Files[i];
    end;
 finally
 Destroy;
 end;
 end;
end;

procedure TFormCopyFFFGropPC.Button3Click(Sender: TObject);
var
i:integer;
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор списка файлов для копирования';
  Options:=[fdoAllowMultiSelect,fdoForceShowHidden];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     for I := 0 to Files.Count-1 do
     with LVFileFolder.Items.Add do
     caption:= Files[i];
    end;
 finally
 Destroy;
 end;
 end;
end;

procedure TFormCopyFFFGropPC.Button4Click(Sender: TObject);
var
i:integer;
begin
if LVFileFolder.SelCount=0 then exit;
for I := LVFileFolder.Items.Count-1 downto 0 do
if LVFileFolder.Items[i].Selected then
 LVFileFolder.Items[i].Delete;
end;

procedure TFormCopyFFFGropPC.Button5Click(Sender: TObject);
begin
FormCopyFFFGropPC.Close;
end;

procedure TFormCopyFFFGropPC.Button6Click(Sender: TObject);
var
newstr:string;
begin
if not InputQuery('Полный путь до файла или каталога', 'Путь:', newstr)
    then exit;
LVFileFolder.Items.Add.Caption:=newstr;
end;

procedure TFormCopyFFFGropPC.ButtonCopyClick(Sender: TObject);
var
DestFile,SourceFile:string;
listPC:TstringList;
i,z,countcopy,typeoperation:integer;
CheckFolder:bool;
begin
try
if LVFileFolder.Items.Count=0 then exit;
if (LVPC.Items.Count=0) then exit;
if (sender as tbutton).Name='ButtonCopy' then //если копируем
begin
 typeoperation:=2;
 CheckFolder:=CheckBox1.Checked;
 SourceFile:=Shell_StrSourse(LVFileFolder);
 DestFile:=Shell_StrDest(EditPath.Text+'\');// // куда копируем  \\компьютер\c$\temp\  заканчивается двойным нулем как и последний файл в списке источников
 if AnsiPos(':',DestFile)=2 then // если в пути назначения найдено двоеточие то запрашиваем правильность ввода пути, для копирования в сеть после диска ставится $
 begin
 delete(DestFile,2,1); // удаляем символ :
 insert('$',DestFile,2); // вставляем на его место $
 if not InputQuery('Проверка каталога назначения (C:\ -> C$\)', 'Каталог назначения:', DestFile) then
   begin
   exit;
   end;
 end;
end;

if (sender as tbutton).Name='ButtonDelete' then   // если удаляем
begin
i:=MessageBox(Self.Handle, PChar('Вы действительно хотите удалить выбранные файлы/каталоги, '+#10#13+' на выбранном списке компьютеров?')
        , PChar('Удаление') ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;

typeoperation:=3;
CheckFolder:=false;
SourceFile:='';
DestFile:=Shell_StrDelete(LVFileFolder);//
end;


countcopy:= (sender as Tbutton).Tag;
inc(countcopy);
(sender as Tbutton).Tag:=countcopy;


listPC:=TstringList.Create;
for I := 0 to LVPC.Items.Count-1 do
 if LVPC.Items[i].Checked then listPC.Add(LVPC.Items[i].Caption);

Form11.CopyFileFolderListPC(
SourceFile, // файл или каталог источник
DestFile, //файл или каталог куда копируем
EditUser.Text, // пользователь для авторизации
EditPaswd.Text, //пароль для авторизации
listpc.CommaText, // список компьютеров для копирования
CheckFolder,     // проверять и создавать каталог
false, /// возможность отмены
FormCopyFFFGropPC, // форма родитель
countcopy, //для выбора переменной (record) из массива
typeoperation);//2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименовать

except
    on E: Exception do
     ShowMessage(Format('Ошибка :  %s',[E.Message]));
  end;
 listPC.Free;
end;

procedure TFormCopyFFFGropPC.LVFileFolderKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key=113  then
  begin
  if (LVFileFolder.Items.Count=0) or (LVFileFolder.SelCount<>1) then exit;
  LVFileFolder.Selected.EditCaption;
  end;
end;




procedure TFormCopyFFFGropPC.LoadlistPC(Sender: TObject);
var                      ///// загрузка списков компьютеров  при загрузке формы
i:integer;
begin
try
for I := 0 to FrmDomaininfo.ComboBox2.Items.Count-1 do
ListGr.Items.Add(FrmDomaininfo.ComboBox2.Items[i]);
Except
  on E: Exception do
  begin
  FrmDomaininfo.memo1.Lines.add('Ошибка загрузки списка компьютеров :'+E.Message);
  end;
end;
end;

procedure TFormCopyFFFGropPC.FormListPCClose(Sender: TObject; var Action: TCloseAction);
begin
FormGr.Release; /// уничтожение формы
end;

procedure TFormCopyFFFGropPC.ButOkClickNamePC(Sender: TObject);
var
i:integer;
itemyes:bool;
p:Tpoint;
begin
itemyes:=false;
if ListGr.Text<>'' then
begin
for I := 0 to LVPC.Items.Count-1 do
if AnsiUpperCase(LVPC.Items[i].Caption)=AnsiUpperCase(ListGr.Text) then /// поиск в списке
begin
 LVPC.Items[i].Selected:=true;
 LVPC.ItemIndex:=i;
 LVPC.ItemFocused:=LVPC.Items[i];
 p := LVPC.Items.Item[i].Position;
 LVPC.Scroll(P.X,P.Y);
 itemyes:=true;
 break;
end;
if not itemyes then
 LVPC.Items.Add.Caption:=ListGr.Text;
end;
FormGr.Close;
end;

function TFormCopyFFFGropPC.FormGroupCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormCopyFFFGropPC);
FormGr.Name:='FormGroup';
FormGr.Width:=323;
FormGr.Height:=97;
FormGr.BorderStyle:=bsDialog;
FormGr.OnShow:=LoadlistPC;
FormGr.Caption:='Список компьютеров';
FormGr.OnClose:=FormListPCClose;
FormGr.Position:=poMainFormCenter;
ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:='ListGr';
ListGr.Text:='';
ListGr.Left:=8;
ListGr.Top:=8;
ListGr.Width:=298;
ListGr.DropDownCount:=15;
ListGr.CharCase:= ecUpperCase;
ListGr.AutoDropDown:=true;
ListGr.TabOrder:=0;
ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='ButOk';
ButOk.Caption:='Ок';
ButOk.Top:=35;
ButOk.Left:=231;
ButOk.OnClick:=ButOkClickNamePC;
ButOk.TabOrder:=1;
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

procedure TFormCopyFFFGropPC.N1Click(Sender: TObject);
begin
 FormGroupCreate('','','');
end;

procedure TFormCopyFFFGropPC.N2Click(Sender: TObject);
var
i:integer;
begin
if LVPC.SelCount=0 then exit;
for I := LVPC.Items.Count-1 downto 0 do
if LVPC.Items[i].Selected then
 LVPC.Items[i].Delete;
end;

end.
