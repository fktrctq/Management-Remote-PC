unit RemoteExplorer;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, StdCtrls, ImgList, Vcl.Buttons, Vcl.Menus,
  JvExExtCtrls, JvNetscapeSplitter,clipbrd;

type
 TPropForPasteFF = record
    SourceFileName :String;
    DestFileName   :String;
    NamePC:string;
    user:string;
    pass:string;
    Recursive      :boolean;
  end;

 type
 TPropForDeleteFF = record
    SourceFileName :String;
    NamePC:string;
    user:string;
    pass:string;
  end;

  type
 TFindFolderFile = record
    FindText :String;
    FullPath:string;
    ListViewForFind:TlistView;
    NamePC:string;
    user:string;
    pass:string;
  end;

  TMRPCExplorer = class(TForm)
    TreeViewFolders: TTreeView;
    ListViewFiles: TListView;
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    Splitter1: TSplitter;
    ImageList1: TImageList;
    ImageList2: TImageList;
    MemoLog: TMemo;
    Panel2: TPanel;
    ButtonUpdate: TButton;
    EditPath: TEdit;
    Panel3: TPanel;
    SpeedButton1: TSpeedButton;
    Panel4: TPanel;
    Panel5: TPanel;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    SpeedButton2: TSpeedButton;
    PopupMenu2: TPopupMenu;
    N6: TMenuItem;
    ComboBox2: TComboBoxEx;
    ImageListPC: TImageList;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    LEditFile: TLabeledEdit;
    Button2: TButton;
    Panel6: TPanel;
    Panel7: TPanel;
    ButtonFind: TSpeedButton;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    EditFind: TEdit;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    JvNetscapeSplitter2: TJvNetscapeSplitter;
    EditUser: TEdit;
    EditPass: TEdit;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    Windows1: TMenuItem;
    procedure TreeViewFoldersChange(Sender: TObject; Node: TTreeNode);
    procedure FormShow(Sender: TObject);
    procedure ButtonUpdateClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ListViewFilesDblClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure EditPathKeyPress(Sender: TObject; var Key: Char);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure ListViewFilesKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ListViewFilesEdited(Sender: TObject; Item: TListItem;
      var S: string);
    procedure N10Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure ButtonFindClick(Sender: TObject);
    procedure FormFindClose(Sender: TObject; var Action: TCloseAction); // закрытие формы для поиска
    procedure ListViewFindDblClick(Sender: TObject);
    procedure EditFindKeyPress(Sender: TObject; var Key: Char);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure ListViewFilesColumnClick(Sender: TObject; Column: TListColumn);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure Windows1Click(Sender: TObject);
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState); // открытие каталога/файла из поиска
  private
    Dt            : OleVariant;
    FSWbemLocatorExp : OLEVariant;
    FWMIServiceExp   : OLEVariant;
    FExtensions   : TStringList; /// списко расширений файлов для иконок
    FileFolderSourceName:TstringList; /// источники для копирования файла или каталога
    function ConnectWMI : boolean;
    procedure UpdateFolderTreeItem(Node:TTreeNode);
    procedure UpdateFileTreeItem(Path: String);  ////для диска открытие списка файлов и каталогов по указанному пути
    procedure UpdateFileTreeItemForUSBFlas(Path: String); ///для флешки открытие списка файлов и каталогов по указанному пути
    function deleteFileWmi(Path,FileName:string):integer;
    function deleteFolderWmi(Path,FolderName:string):integer;
    function PasteFolderWmi(SourceFileName,DestFileName:string;Recursive:bool):integer;
    function PasteFileWmi(SourceFileName,DestFileName:string):integer;
    function RenameFileWmi(OldName,NewName:string):integer; /// переименование файла
    function RenameFolderWmi(OldName,NewName:string):integer; // переименование каталога
    function VarDateToDateTime(const V : OleVariant): TDateTime;
    function GetNodeFullPath(Node:TTreeNode) : string;
    function GetImageIndexExt(const Ext : string) : Integer;
    function errorExplorer(err:integer):string;  // ошибки Explorer
    procedure ConnectRemote;
    function findtypedrive(drive:string):integer; // определение типа диска
    function CopyOtherThread(SourceFileName,DestFileName,FFname:string;ind,TypeFF:integer):bool; // в функцию передаются параметры для копирования в отдельных потоках
    function propertyFileFolder(path:string;typeFF:integer;FormOwn:Tform;indCount:integer):bool;
    function OpenFileOrExplorer(PathFile:string):integer; // открытие файла на машине администратора
  public
     function findFFinListview(FindText:string):bool; // функция поиска каталога/файла в listview
  end;
   function PasteFolderWmiThread(paramFF:pointer):integer;  // копирование каталога в отдельном потоке
   function PasteFileWmiThread(paramFF:pointer):integer;   // копирование файла в отдельном потоке
   function deleteFolderFileWmiThread(paramDel:pointer):integer; // удаление каталогов/файлов в потоке
   function FindFileFolderWmiThread(paramFF:pointer):integer;    // поиск каталогов/файлов в потоке

  var
   MRPCExplorer: TMRPCExplorer;

implementation

Uses
  ShellAPI,
  ActiveX,
  ComObj,Umain,CopyPasteFF,FormCopyFF,FormNewProcess,DlgNewInstall;

const
  wbemFlagForwardOnly = $00000020;

  ThreadVar
PointForFolder: ^TPropForPasteFF;
PointForFile: ^TPropForPasteFF;
PointForDelete: ^TPropForDeleteFF;
PointForFind: ^TFindFolderFile;
  var
RunPasteFolder    :array [0..2000] of TPropForPasteFF;///указатель для вставки каталога в потоке
RunPasteFile      :array [0..2000] of TPropForPasteFF;///указатель для вставки файла в потоке
RunDeleteFF       :array [0..2000] of TPropForDeleteFF;  ///указатель для удаление каталога/файла в потоке
RunFindFF         :array [0..2000] of TFindFolderFile ; ///указатель для поиска каталого/папок в потоке
treadIDF : array [0..2000] of LongWord;
resF :array [0..2000] of integer;
CancelCopyFF:boolean;
FunSort:bool;
FunSortIn:integer;
CurPCConnectRegEdit:string;
{$R *.dfm}

////Функция сортировки в TListView
function CustomSortProc(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
  if ParamSort=0 then
    case FunSort of
    true: Result := CompareText(Item1.Caption,Item2.Caption);
    false: Result := CompareText(Item2.Caption,Item1.Caption);
    end
  else
    if Item1.SubItems.Count>ParamSort-1 then
    begin
      if Item2.SubItems.Count>ParamSort-1 then
          case FunSort of
          true:Result := CompareText(Item1.SubItems[ParamSort-1],Item2.SubItems[ParamSort-1]);
          false:Result := CompareText(Item2.SubItems[ParamSort-1],Item1.SubItems[ParamSort-1]);
          end
      else
        Result := 1;
    end
    else
      Result:=-1;
end;



function SortThirdSubItemAsDate3(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
   Result := 0;
   if (StrToDate( Item1.SubItems[0] )) > (StrToDate( Item2.SubItems[0] )) then
      Result := ParamSort
   else
   if StrToDate(( Item1.SubItems[0] )) < StrToDate(( Item2.SubItems[0] )) then
      Result := -ParamSort;
end;


///////////////////////////////////////////////////////////////////


function TMRPCExplorer.errorExplorer(err:integer):string;
begin
  case err of
  0: result:='Операция успешно выполнена' ;
  2:  result:='Доступ запрещен';
  8:  result:='Неизвестная ошибка' ;
  9:  result:='Неверный объект';
  10:  result:='Объект уже существует';
  11:  result:='Файловая система отличается от NTFS';
  12:  result:='Платформа Windows';
  13: result:='Drive not the same';
  14: result:='Каталог не пустой';
  15: result:='Нарушение совместного использования';
  16: result:='Неверный начальный файл';
  17: result:='Не достаточно прав';
  21: result:='Не верный параметр';
  else result:=SysErrorMessage(err);
  end;
end;

function TMRPCExplorer.propertyFileFolder(path:string;typeFF:integer;FormOwn:Tform;indCount:integer):bool;
var                      /// совйства файлов и каталогов
PrForm:Tform;
PrMemo:Tmemo;
Wql,Drive:string;
FWbemObject:olevariant;
function calcKB(sz:OleVariant):integer;
begin
try
result:=((round(sz))div 1024 );
if result=0 then result:=1;
sz:=Unassigned;
VariantClear(sz);
except
result:=1;
end;
end;
begin
try
PrForm:=Tform.Create(Formown);
PrForm.Name:='FormProp'+inttostr(indCount);
PrForm.Width:=300;
PrForm.Height:=200;
PrForm.Position:=poMainFormCenter;
PrForm.Tag:= indCount;
PrForm.BorderIcons:=[biSystemMenu,biMinimize];
PrForm.BorderStyle:=bsSizeable;
PrForm.FormStyle:=fsNormal;//fsMDIChild;//fsNormal;//fsMDIForm;//fsStayOnTop;
PrForm.PopupMode:=pmAuto;
PrForm.Position:=poOwnerFormCenter;
//PrForm.Caption:='Поиск - '+EditFind.Text;
PrForm.OnClose:=FormFindClose;
PrMemo:=Tmemo.Create(PrForm);
PrMemo.parent:=PrForm;
PrMemo.Name:='Prmemo'+inttostr(indCount);
PrMemo.Clear;
PrMemo.Align:=alClient;
PrMemo.ScrollBars:=ssVertical;
PrMemo.ReadOnly:=true;
PrForm.Show; // открываем форму
////////////////////////////////////////////////////
//Wql     :=Format('SELECT FileName,name FROM CIM_DataFile Where Drive="%s" AND Path="%s" AND FileName LIKE "%s"',[Drive,WmiPath,FindText]);
//FWbemObjectSet:= FWMIService.ExecQuery(wql,'WQL',wbemFlagForwardOnly);
if typeFF<>0 then /// свойства файла
begin
Wql:=path;
drive:=ExtractFileDrive(Path);
FWbemObject   := FWMIServiceExp.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
if FWbemObject.Filename<>null then PrForm.Caption:='Свойства '+string(FWbemObject.Filename);
with Prmemo.lines do
 begin
 if FWbemObject.FileName <>null then add('Имя файла - '+string(FWbemObject.FileName));
 if FWbemObject.FileType <>null then add('Тип файла - '+string(FWbemObject.FileType));
 if FWbemObject.Extension <>null then add('Расширение файла - '+string(FWbemObject.Extension));
  if FWbemObject.name<>null then add('Расположение - '+string(FWbemObject.name))
 else if FWbemObject.Description <>null then add('Расположение - '+string(FWbemObject.Description));
 if not ((findtypedrive(Drive)=2) or (findtypedrive(Drive)=5) or (findtypedrive(Drive)=1)) then // если на локальном диске то смотрим даты, иначе нет
 begin
 if FWbemObject.FileSize <>null then add('Размер - '+ (inttostr(calcKB(FWbemObject.FileSize)))+' Кб');
 if FWbemObject.CreationDate <>null then add('Создан - '+datetostr(VarDateToDateTime(FWbemObject.CreationDate)));
 if FWbemObject.LastModified <>null then add('Изменен - '+datetostr(VarDateToDateTime(FWbemObject.LastModified)));
 if FWbemObject.LastAccessed <>null then add('Открыт - '+datetostr(VarDateToDateTime(FWbemObject.LastAccessed)));
 end;
 if FWbemObject.Archive<>null then add('Архивный - '+string(FWbemObject.Archive));
 if FWbemObject.Compressed then
 begin
 add('Сжатый - '+string(FWbemObject.Compressed));
 if FWbemObject.CompressionMethod <>null then add('Метод сжатия - '+string(FWbemObject.CompressionMethod));
 end;
 if FWbemObject.CSName <>null then add('Компьютер - '+string(FWbemObject.CSName));
 if FWbemObject.Encrypted <>null then if FWbemObject.Encrypted then
 begin
 add('Зашифрован - '+string(FWbemObject.Encrypted));
 if FWbemObject.EncryptionMethod <>null then add('Алгоритм шифрования - '+string(FWbemObject.EncryptionMethod))
 else add('Алгоритм шифрования - неизвестен');
 end;
 if FWbemObject.InUseCount <>null then add('Количество открытых файлов - '+string(FWbemObject.InUseCount));
 if FWbemObject.Manufacturer <>null then add('Организация - '+(string(FWbemObject.Manufacturer)));
 if FWbemObject.Version <>null then add('Версия файла - '+string(FWbemObject.Version));
 if FWbemObject.System <>null then if FWbemObject.System then add('Системный - '+string(FWbemObject.System));
 if FWbemObject.Writeable <>null then add('Запись файла - '+string(FWbemObject.Writeable)); //запись
 if FWbemObject.Readable <>null then add('Чтение файла - '+string(FWbemObject.Readable)); //чтение
 if FWbemObject.Hidden <>null then if FWbemObject.Hidden then add('Скрытый - '+string(FWbemObject.Hidden));
 if FWbemObject.Status <>null then add('Статус - '+string(FWbemObject.Status));
 if FWbemObject.FSName <>null then add('Файловая система - '+string(FWbemObject.FSName));
end;
FWbemObject:=Unassigned;
end;

if typeFF=0  then // свойства каталога
begin
Wql:=path;
drive:=ExtractFileDrive(Path);
FWbemObject   := FWMIServiceExp.Get(Format('CIM_Directory.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
if FWbemObject.Filename<>null then PrForm.Caption:='Свойства '+string(FWbemObject.Filename);
with Prmemo.lines do
 begin
 if FWbemObject.FileName <>null then add('Имя каталога - '+string(FWbemObject.FileName));
 if FWbemObject.FileType <>null then add('Тип каталога - '+string(FWbemObject.FileType));
 if FWbemObject.Extension <>null then add('Расширение - '+string(FWbemObject.Extension));
  if FWbemObject.name<>null then add('Расположение - '+string(FWbemObject.name))
 else if FWbemObject.Description <>null then add('Расположение - '+string(FWbemObject.Description));
 if not ((findtypedrive(Drive)=2) or (findtypedrive(Drive)=5) or (findtypedrive(Drive)=1)) then // если на локальном диске то смотрим даты, иначе нет
 begin
 if FWbemObject.FileSize <>null then add('Размер - '+ (inttostr(calcKB(FWbemObject.FileSize)))+' Кб');
 if FWbemObject.CreationDate <>null then add('Создан - '+datetostr(VarDateToDateTime(FWbemObject.CreationDate)));
 if FWbemObject.LastModified <>null then add('Изменен - '+datetostr(VarDateToDateTime(FWbemObject.LastModified)));
 if FWbemObject.LastAccessed <>null then add('Открыт - '+datetostr(VarDateToDateTime(FWbemObject.LastAccessed)));
 end;
 if FWbemObject.Archive<>null then add('Архивный - '+string(FWbemObject.Archive));
 if FWbemObject.Compressed then
 begin
 add('Сжатый - '+string(FWbemObject.Compressed));
 if FWbemObject.CompressionMethod <>null then add('Метод сжатия - '+string(FWbemObject.CompressionMethod));
 end;
 if FWbemObject.CSName <>null then add('Компьютер - '+string(FWbemObject.CSName));
 if FWbemObject.Encrypted <>null then if FWbemObject.Encrypted then
 begin
 add('Зашифрован - '+string(FWbemObject.Encrypted));
 if FWbemObject.EncryptionMethod <>null then add('Алгоритм шифрования - '+string(FWbemObject.EncryptionMethod))
 else add('Алгоритм шифрования - неизвестен');
 end;
 if FWbemObject.InUseCount <>null then add('Количество открытых файлов - '+string(FWbemObject.InUseCount));
 if FWbemObject.System <>null then if FWbemObject.System then add('Системный - '+string(FWbemObject.System));
 if FWbemObject.Writeable <>null then add('Запись файла - '+string(FWbemObject.Writeable)); //запись
 if FWbemObject.Readable <>null then add('Чтение файла - '+string(FWbemObject.Readable)); //чтение
 if FWbemObject.Hidden <>null then if FWbemObject.Hidden then add('Скрытый - '+string(FWbemObject.Hidden));
 if FWbemObject.Status <>null then add('Статус - '+string(FWbemObject.Status));
 if FWbemObject.FSName <>null then add('Файловая система - '+string(FWbemObject.FSName));
 end;
FWbemObject:=Unassigned;
end;


FWbemObject:=Unassigned;
VariantClear(FWbemObject);
except
 on E: Exception do
 begin
 memolog.lines.add('Ошибка создания формы свойств'+e.Message);
 //if Assigned(PrForm) then if not PrForm.ShowHint then PrForm.Free;
 end;
end;
end;

function TMRPCExplorer.findFFinListview(FindText:string):bool;
var
p:Tpoint;
z:integer;
ListItem: TListItem;
begin
try
if (FindText='') then exit;
if (ListViewFiles.Items.Count=0)then exit;
for z:=0 to ListViewFiles.items.count-1 do
begin
if pos(Ansiuppercase(FindText),Ansiuppercase(ListViewFiles.items[z].Caption))<>0 then
  begin
  ListViewFiles.items[z].Selected:=true;
  ListViewFiles.ItemIndex:=z;
  ListViewFiles.ItemFocused:=ListViewFiles.items[z];
  p := ListViewFiles.Items.Item[z].Position;
  ListViewFiles.Scroll(P.X,P.Y-40);
  break;
  end;
end;
 //if ListViewFiles.ItemFocused<>nil then ListViewFiles.ItemFocused.MakeVisible(false);
 except
 on E: Exception do
 MemoLog.Lines.Add('Ошибка поиска: '+e.Message);
 end;
end;


procedure TMRPCExplorer.Button2Click(Sender: TObject); // запуск копирования файла в потоке
var
htoken:THandle;
drive,Path:string;
ProgressCopy:TProgressBar;
step:string;
idn:integer;
begin
  try
   if not (LogonUserA (PAnsiChar(EditUser.Text), PAnsiChar (ComboBox2.Text),
                               PAnsiChar (EditPass.Text), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
                               LOGON32_PROVIDER_WINNT50, hToken))
    then GetLastError();
  except
        on E: Exception do MemoLog.Lines.Add('Ошибка при LogonUser '+ComboBox2.Text+' - '+e.Message)
   end;
  CancelCopyFF:=false;
 try
 step:='0';
  drive:=copy(ExtractFileDrive(EditPath.text),1,1);
  Path:=copy(EditPath.text,3,length(EditPath.text));
  MemoLog.Lines.Add('Источник - '+LEditFile.Text);
  MemoLog.Lines.Add('Приемник - '+'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\'+ExtractFileName(LEditFile.Text));
  step:='1';
  step:='2';
  idn:=button2.Tag;
  inc(idn);
  button2.Tag:=idn;
  step:='3';
  Form11.CopyFileFolder(LEditFile.Text,'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\'+ExtractFileName(LEditFile.Text),false,MRPCExplorer,idn,2); // 2 -  копировать
  except on E: Exception do
  begin
  MemoLog.Lines.Add('Шаг - '+step+' - Ошибка при CopyFile '+ComboBox2.Text+' - '+e.Message);
  end;
   end;
end;



procedure TMRPCExplorer.ButtonUpdateClick(Sender: TObject);
begin
FWMIServiceExp:=Unassigned;
TreeViewFolders.Items.BeginUpdate;
try
TreeViewFolders.Items.Clear;
finally
TreeViewFolders.Items.EndUpdate;
end;

ListViewFiles.Items.BeginUpdate;
try
ListViewFiles.Items.Clear;
finally
ListViewFiles.Items.EndUpdate;
end;
////сначала все обнуляем
ConnectRemote; ///  и соединяемся снова
end;

procedure TMRPCExplorer.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var
  i:integer;
begin
try
if (key=VK_RETURN) then
if  (frmDomainInfo.ping(ComboBox2.Text)) then
  Begin
   i:=MessageDlg('Подключится к компьютеру ' +ComboBox2.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
  if i=IDYes then
  begin
  ConnectRemote // если да то подключаемся к реестру нового компа
  end
  else
  begin
  ComboBox2.DroppedDown:=false;
  ComboBox2.Text:=CurPCConnectRegEdit;
  end;
  End
  else ComboBox2.Text:=CurPCConnectRegEdit;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (22): '+e.Message);
end;
end;

procedure TMRPCExplorer.ComboBox2Select(Sender: TObject);
var
i:integer;
begin
try
if (frmDomainInfo.ping(ComboBox2.Text)) then
begin
  i:=MessageDlg('Подключится к компьютеру ' +ComboBox2.Text+'?', mtConfirmation,[mbYes,mbCancel],0);
  if i=IDYes then
  begin
  ConnectRemote // если да то подключаемся к реестру нового компа
  end
  else
  begin
  ComboBox2.DroppedDown:=false;
  ComboBox2.Text:=CurPCConnectRegEdit;
  end;
end
else ComboBox2.Text:=CurPCConnectRegEdit;
except on E: Exception do memolog.Lines.Add('Произошла непредвиденная ошибка (22): '+e.Message);
end;
end;

procedure TMRPCExplorer.ConnectRemote;
begin
  begin
    StatusBar1.Panels[0].Text:=Format('Подключение к %s',[ComboBox2.Text]);
    try
      if ConnectWMI then
      begin
        UpdateFolderTreeItem(nil);
      end
      else
      ShowMessage(Format('Невозможно установить соединение с сервером %s ',[ComboBox2.Text]));
    finally
      StatusBar1.Panels[0].Text:='';
    end;
  end

end;

function TMRPCExplorer.ConnectWMI : Boolean;
begin
 Result:=False;
 try
    Dt:=Unassigned;
    FSWbemLocatorExp:=Unassigned;
    FWMIServiceExp:=Unassigned;
    Dt:=CreateOleObject('WbemScripting.SWbemDateTime');
    FSWbemLocatorExp := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIServiceExp   := FSWbemLocatorExp.ConnectServer(ComboBox2.Text, 'root\CIMV2', EditUser.Text, EditPass.text,'','',128);
    FWMIServiceExp.Security_.impersonationlevel:=3;
    FWMIServiceExp.Security_.authenticationLevel := 6;
    Result:=True;
 except
    on E:EOleException do
        MemoLog.Lines.Add(Format('EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do
        MemoLog.Lines.Add(E.Message);
 end;
end;

procedure TMRPCExplorer.EditFindKeyPress(Sender: TObject; var Key: Char);
begin
if EditFind.Text='' then exit;
if key=#13 then ButtonFind.OnClick(ButtonFind);
end;

procedure TMRPCExplorer.EditPathKeyPress(Sender: TObject; var Key: Char); /// переход в директорию по нажатию Enter
begin
if key=#13 then
 if  (not VarIsClear(FWMIServiceExp)) and (Length(EditPath.Text)>1) then
 begin
  UpdateFileTreeItem(EditPath.Text);
 end;
end;

procedure TMRPCExplorer.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FileFolderSourceName.free;
VariantClear(Dt);
VariantClear(FSWbemLocatorExp);
VariantClear(FWMIServiceExp);
end;

procedure TMRPCExplorer.FormCreate(Sender: TObject);
begin
FExtensions:=TStringList.Create; /// списко расширений для отображения иконок на файлах
FExtensions.Add('folder');
FunSortIn:=-1; // для сортировки
FunSort:=true;
end;

procedure TMRPCExplorer.FormShow(Sender: TObject);
begin
TreeViewFolders.Items.Clear;
ListViewFiles.Clear;
MemoLog.Clear;
FileFolderSourceName:=TStringList.Create;
ConnectRemote;
end;

function TMRPCExplorer.GetImageIndexExt(const Ext: string): Integer;
var
  Icon : TIcon;
  FileInfo : SHFILEINFO;
begin
  ZeroMemory(@FileInfo, SizeOf(FileInfo));
  Result:=FExtensions.IndexOf(Ext);
  if Result=-1 then
  begin
    Icon := TIcon.Create;
    try
       if SHGetFileInfo(PChar('*'+Ext), FILE_ATTRIBUTE_NORMAL, FileInfo, SizeOf(FileInfo),
       SHGFI_ICON or SHGFI_SMALLICON or SHGFI_SYSICONINDEX or SHGFI_USEFILEATTRIBUTES ) <> 0 then
        begin
          Icon.Handle := FileInfo.hIcon;
          Result:=ImageList2.AddIcon(Icon);
          FExtensions.Add(Ext);
        end;
    finally
      Icon.Free;
    end;
  end;
end;

function TMRPCExplorer.GetNodeFullPath(Node: TTreeNode): string;
var
drive:string;
path:string;
labelpos:integer;
begin
  Result:=Node.Text;
  if pos(' (',result)<>0 then result:=copy(result,1,pos(' (',result)-1); // удаляем Label из названия диска
   if Node.Parent<>nil then
   repeat
      Node:=Node.Parent;
      Result:=Node.Text+'\'+Result;
   until Node.Parent=nil;
  if pos(' (',result)<>0 then /// если метка (label) диска осталась то ищем ее и удаляем из пути
  begin
   drive:=copy(result,1,pos(' (',result)-1);
   labelpos:=pos(') ',result);
   path:=copy(result,labelpos+2,length(result));
   result:=drive+path;
  end;
 //MemoLog.Lines.add(result);
end;

procedure TMRPCExplorer.ListViewFilesColumnClick(Sender: TObject;
  Column: TListColumn);
begin
FunSortIn:=-FunSortIn;
FunSort:=not FunSort;
 if Column=ListViewFiles.Columns[1] then
  ListViewFiles.CustomSort( @SortThirdSubItemAsDate3, FunSortIn)
  else
 ListViewFiles.CustomSort(@CustomSortProc, Column.Index);
end;

function TMRPCExplorer.OpenFileOrExplorer(PathFile:string):integer;
begin
try
result:=shellAPI.ShellExecute(0, 'Open', PChar(PathFile), nil, nil, SW_SHOWNORMAL);
except on E: Exception do memoLog.Lines.Add('Ошибка при открытии приложения : '+e.Message);
end;
end;

procedure TMRPCExplorer.ListViewFilesDblClick(Sender: TObject);  // Открыть каталог или файл
var
z:integer;
function StrRes(z:integer):string;
begin
Result:='';
if z=0 then result:='Оперативной системе не хватает памяти или ресурсов.';
if z=2 then result:='Указаный файл не найден.'  ;
if z=3 then result:='Указаный путь не найден.';
if z=11 then result:='EXE файл не рабочий (не Win32 .EXE или ошибка в .EXE образе).';
if z=5 then result:='Операционная система отказывает в доступе к указаному файлу.';
if z=27 then result:='Имя ассоциированного файла не полное или не правильное.';
if z=30 then result:='Транзакция динамического обмена данными (DDE transaction) не может быть завершена потому что выпоняются другие DDE транзакции.';
if z=29 then result:='Транзакция динамического обмена данными провалилась.';
if z=28 then result:='Транзакция динамического обмена данными не может быть завершена потому что истекло время ожидания ответа.';
if z=32 then result:='Указаная DLL библиотека не найдена.';
if z=2 then result:='Указаный файл не найден.';
if z=31 then result:='Нет приложений ассоциированных с данным расширением файла.';
if z=8 then result:='Не достаточно памяти для завершения операции.';
if z=3 then result:=' Указаный путь не найден.';
if z=26 then result:='Нарушение совместного доступа.';
end;
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
try
if (ListViewFiles.SelCount=0) or (ListViewFiles.Items.Count=0) then exit;
if (ListViewFiles.Selected.SubItems[1]='Каталог') then
begin // если каталог то открываем
if  (not VarIsClear(FWMIServiceExp)) and (ListViewFiles.SelCount=1) then
UpdateFileTreeItem(EditPath.Text+'\'+ListViewFiles.ItemFocused.Caption);
end
else
begin
if not FileExists('\\'+ComboBox2.Text+'\'+findElement(EditPath.Text)+'\'+ListViewFiles.Selected.Caption) then // проверяем доступность файла
begin
ShowMessage('Сетевой путь не найден '+'\\'+ComboBox2.Text+'\'+findElement(EditPath.Text)+'\'+ListViewFiles.Selected.Caption);
exit;
end;
 z:=MessageDlg('Открыть файл '+ListViewFiles.Selected.Caption+' на Вашем компьютере?', mtConfirmation,[mbYes,mbCancel],0);
if z=IDYes then
begin
z:=OpenFileOrExplorer('\\'+ComboBox2.Text+'\'+findElement(EditPath.Text)+'\'+ListViewFiles.Selected.Caption);
if StrRes(z)<>'' then showmessage(StrRes(z));
end;
end;

except
    on E: Exception do
     ShowMessage(Format('Ошибка открытия каталога : %s ',[E.Message]));
  end;
end;



procedure TMRPCExplorer.ListViewFilesEdited(Sender: TObject; Item: TListItem;
  var S: string);          // окончание редактирования имени   OldNameFF,NewNameFF
var
 OldNameFF,NewNameFF:string; //переменные для переименования
 TypeRenameFF:integer; // тип редактируемого объекта, файл или каталог
begin
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount<>1) then exit;
NewNameFF:=s; // новое имя каталога или файла
OldNameFF:=Item.Caption; // старое имя
TypeRenameFF:=Item.ImageIndex; // каталог или файл
if NewNameFF=OldNameFF then exit;  // если одинаковые
if TypeRenameFF=0 then RenameFolderWmi(EditPath.text+'\'+OldNameFF,EditPath.text+'\'+NewNameFF)
  else RenameFileWmi(EditPath.text+'\'+OldNameFF,EditPath.text+'\'+NewNameFF);
Item.Selected:=false;  // снимаем выделение, т.к. при нажатии ентер происходит переименование и сразу переход в директорию, если это каталог
end;


procedure TMRPCExplorer.ListViewFilesKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
i,res,z:integer;

begin
if VarIsClear(FWMIServiceExp) then exit;
//MemoLog.Lines.Add(inttostr(key));
if (ssCtrl in Shift)and (ssShift in Shift) and (key=67) then
begin
  N2Click(PopupMenu1); //Shift+Ctrl+C копировать WMI
  exit;
end;
if (ssCtrl in Shift)and (ssShift in Shift) and (key=86) then
begin
 N3Click(PopupMenu1); // Shift+Ctrl+V вставить WMI
 exit;
end;

if (ssCtrl in Shift)and (key=67)  then    N13Click(PopupMenu1); //Ctrl+C копировать в буфер обмена
if (ssCtrl in Shift)and (key=86)  then    N11Click(PopupMenu1); //Ctrl+V вставить из буфера обмена
if (ssCtrl in Shift)and (key=88)  then    N14Click(PopupMenu1); //Ctrl+X вырезать в буфера обмена

case key of
116: SpeedButton2Click(SpeedButton2); /// F5 - обновить
113: // переименовать
    begin
    if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount<>1) then exit;
    ListViewFiles.Selected.EditCaption;
    end;
46:  // удалить
    begin
    if ListViewFiles.IsEditing then exit;/// если редактируем имя то выход
      if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;
       z:=MessageBox(Self.Handle, PChar('Удалить выбранные файлы?')
        , PChar('Удаление файлов и каталогов') ,MB_YESNO+MB_ICONQUESTION);
      if z=IDNO then exit;
      for I := ListViewFiles.Items.Count-1 downto 0 do
      begin
      if ListViewFiles.Items[i].Selected then
       begin
       if ListViewFiles.Items[i].ImageIndex=0 then /// если каталог
       begin
         res:=deleteFolderWmi(EditPath.Text,ListViewFiles.items[i].Caption); // удаляем каталог
         if  res=0 then ListViewFiles.items[i].Delete
         else memolog.Lines.Add('Ошибка удаления '+ListViewFiles.items[i].Caption +' :'+errorExplorer(res));
       end
      else
       begin
         res:=deleteFileWmi(EditPath.Text,ListViewFiles.items[i].Caption); // удаляем файл
         if res=0 then ListViewFiles.items[i].Delete
         else memolog.Lines.Add('Ошибка удаления '+ListViewFiles.items[i].Caption +' :'+errorExplorer(res));
       end;
       end;
       end;
       end;
8: //на директорию ниже
    begin
    if  ListViewFiles.IsEditing then exit;// если  редактируем имя
    SpeedButton1Click(SpeedButton1);
    end;
13: // enter на листбоксе
   begin
   if not ListViewFiles.IsEditing then
   ListViewFilesDblClick(ListViewFiles); // если не редактируем то при нажатии переходим в каталог
   end;
end;
end;


procedure TMRPCExplorer.N10Click(Sender: TObject); /// удаление в потоке
var
i,z,ind:integer;
treadID: array [0..2000] of longword;
res : array [0..2000] of integer;
SourseList:TstringList;
begin
ind:=TMenuItem(Sender).tag;
inc(ind);
TMenuItem(Sender).tag:=ind;
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;
z:=MessageBox(Self.Handle, PChar('Удалить выбранные файлы?')
        , PChar('Удаление файлов и каталогов') ,MB_YESNO+MB_ICONQUESTION);
   if z=IDNO then exit;
SourseList:=TStringList.Create;
for I := 0 to ListViewFiles.Items.Count-1 do
   begin
     if ListViewFiles.Items[i].Selected then // если выделенна
     begin
     SourseList.add(EditPath.Text+'\'+
     ListViewFiles.Items[i].Caption+'='+inttostr(ListViewFiles.Items[i].ImageIndex));
     end;
   end;
//MemoLog.Lines.Add('Перед удалением - '+SourseList.CommaText);
RunDeleteFF[ind].SourceFileName:= SourseList.CommaText;
RunDeleteFF[ind].NamePC:=ComboBox2.Text;
RunDeleteFF[ind].user:=EditUser.Text;
RunDeleteFF[ind].pass:=EditPass.Text;
res[ind]:=BeginThread(nil,0,addr(deleteFolderFileWmiThread),Addr(RunDeleteFF[ind]),0,treadID[ind]); ///
CloseHandle(res[ind]);
SourseList.Free;
end;

procedure TMRPCExplorer.N11Click(Sender: TObject); // вставить из буфера обмена
var
drive,path:string;
h: THandle;
  Sourse, sr: string;
  s: TStringList;
  copyorpaste,ind,i,resoper:integer;
    htoken:THandle;
function Shell_Str(Strs: TStrings): string;  // Функция преобразует TStirngs в спец-строку для Shell это спец-строка для буфера обмена, где строки разделены знаком #0 и вся спец-строка строка заканчивается #0#0
var
  i: Integer;
begin
  for i := 0 to Strs.Count - 1 do
    Result := Result + Strs.Strings[i] + #0;
  Result := Result + #0;
end;

function Clipboard_SendType: Integer;
 // 5 - copy //  2 - cut //  Функция определяет на что посланы данные в буфер: на вырезку или копирование   // Эта функция создаваласть специально для функции ClipBoard_DataPaste,
//  чтоб было понятно что делать: копировать или вырезать.
var
  hn: Cardinal;
  ClipFormat:word;
  szBuffer: array [0 .. 60] of Char;
  FormatID: string;
  pMem: Pointer;
  NameFormatLength:integer;
begin
  Result := 0;
  if not OpenClipboard(Application.Handle) then exit;
  try
    ClipFormat := EnumClipboardFormats(0);
    while (ClipFormat <> 0) do
    begin
      NameFormatLength:=GetClipboardFormatName(ClipFormat, szBuffer,(SizeOf(szBuffer)));
      FormatID := strpas(szBuffer);
      if SameText(FormatID, 'Preferred DropEffect') then
      begin
        hn := GetClipboardData(ClipFormat);
        pMem := GlobalLock(hn);
        Move(pMem^, Result, 4); // <- теперь в Result тип операции
        GlobalUnlock(hn);
        Break;
      end;
      ClipFormat := EnumClipboardFormats(ClipFormat);
    end;
 finally
    CloseClipboard;
  end;
end;
begin
try
 try // вход пользовтеля на удаленный ПК
   if not (LogonUserA (PAnsiChar(EditUser.Text), PAnsiChar (ComboBox2.Text),
                               PAnsiChar (EditPass.Text), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
                               LOGON32_PROVIDER_WINNT50, hToken))
    then GetLastError();
  except
        on E: Exception do MemoLog.Lines.Add('Ошибка LogonUser '+ComboBox2.Text+' - '+e.Message)
   end;
 drive:=copy(ExtractFileDrive(EditPath.text),1,1);
 Path:=copy(EditPath.text,3,length(EditPath.text));
 ind:=TMenuItem(Sender).tag;
 inc(ind);
 TMenuItem(Sender).tag:=ind;
 if not Clipboard.HasFormat(CF_HDROP) then
   exit;
 Clipboard.Open;
 h := Clipboard.GetAsHandle(CF_HDROP);
    if h <> 0 then
    begin
      s := TStringList.Create;
      WinExplorer.ClipBoardGetFiles(s, h);
      //for I := 0 to s.Count-1 do   MemoLog.Lines.Add(s[i]);
      Sourse := Shell_Str(s); ///каким то макаром в строку помещаются все пути до  файлов и папок для копирования
     // MemoLog.Lines.Add('1 Sourse - '+Sourse);  /// а вот тут они не отображаются
      s.free;
      sr := Copy(Sourse, 0, Pos(#0, Sourse) - 1); // Path того что скопированно в буфер
      sr := ExtractFilePath(sr); // Родительская папка Data
      copyorpaste:=Clipboard_SendType; // запрашиваем что делать, копировать или переместить
      resoper:=0;
      if copyorpaste=5 then
        resoper:=form11.FunctionCopyFF(Sourse,'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\',false,MRPCExplorer,ind,2); // копируем не в потоке
       //Form11.CopyFileFolder(Sourse,'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\',false,MRPCExplorer,ind,2); // копируем в потоке
      if copyorpaste=2 then
       resoper:=form11.FunctionCopyFF(Sourse,'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\',false,MRPCExplorer,ind,1); // перемещаем не в потоке
      // Form11.CopyFileFolder(Sourse,'\\'+ComboBox2.text+'\'+drive+'$'+Path+'\',false,MRPCExplorer,ind,1); // перемещаем в потоке
      if resoper<>0 then memolog.Lines.add('Операция копирования завершилась ошибкой с кодом - '+inttostr(resoper)+' ('+SysErrorMessage(resoper)+')');
   end;
   finally
    Clipboard.Close;
   end;
 CloseHandle(htoken);
end;

procedure TMRPCExplorer.N13Click(Sender: TObject); // копирование в буфер обмена
var
i:integer;
listCopy:TstringList;
drive,path:string;
htoken:THandle;
begin
  try
   if not (LogonUserA (PAnsiChar(EditUser.Text), PAnsiChar (ComboBox2.Text),
                               PAnsiChar (EditPass.Text), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
                               LOGON32_PROVIDER_WINNT50, hToken))
    then GetLastError();
  except
        on E: Exception do MemoLog.Lines.Add('Ошибка при LogonUser '+ComboBox2.Text+' - '+e.Message)
   end;
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;
listCopy:=TStringList.Create;
drive:=copy(ExtractFileDrive(EditPath.text),1,1);
Path:=copy(EditPath.text,3,length(EditPath.text));
for I := 0 to ListViewFiles.Items.Count-1 do
if ListViewFiles.Items[i].Selected then listCopy.Add('\\'+ComboBox2.text+'\'+drive+'$'+Path+'\'+ListViewFiles.Items[i].Caption);
form11.Clipboard_DataSend(listCopy,5);
MemoLog.Lines.Add('Копирую в буфер обмена - '+listCopy.CommaText);
listCopy.Free;
CloseHandle(htoken);
end;

procedure TMRPCExplorer.N14Click(Sender: TObject); // вырезать в буфер обмена
var
i:integer;
listCopy:TstringList;
drive,path:string;
htoken:THandle;
begin
  try
   if not (LogonUserA (PAnsiChar(EditUser.Text), PAnsiChar (ComboBox2.Text),
                               PAnsiChar (EditPass.Text), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
                               LOGON32_PROVIDER_WINNT50, hToken))
    then GetLastError();
  except
        on E: Exception do MemoLog.Lines.Add('Ошибка при LogonUser '+ComboBox2.Text+' - '+e.Message)
   end;

if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;
listCopy:=TStringList.Create;
drive:=copy(ExtractFileDrive(EditPath.text),1,1);
Path:=copy(EditPath.text,3,length(EditPath.text));
for I := 0 to ListViewFiles.Items.Count-1 do
if ListViewFiles.Items[i].Selected then listCopy.Add('\\'+ComboBox2.text+'\'+drive+'$'+Path+'\'+ListViewFiles.Items[i].Caption);
form11.Clipboard_DataSend(listCopy,2);
MemoLog.Lines.Add('Вырезаю в буфер обмена - '+listCopy.CommaText);
listCopy.Free;
CloseHandle(htoken);
end;

function FindExt(filename:Tlistitem):integer; //поиск расширение msi
begin
 if ExtractFileExt(filename.Caption)='.msi' then
 begin
 result:=1;
 exit;
 end;
 if filename.ImageIndex=0 then result:=3  // каталог
 else result:=2;   //все остальное

end;

procedure TMRPCExplorer.N15Click(Sender: TObject);
begin
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;

if FindExt(ListViewFiles.Selected)=1 then // если 1 файл msi то запустить диалог установки
  begin
  GroupPC:=false;
  NewProgramMyPS:=Combobox2.Text;
  OKRightDlg1.LabeledEdit1.Text:=EditPath.Text+'\'+ListViewFiles.Selected.Caption;
  OKRightDlg1.ShowModal;
  end;
if  FindExt(ListViewFiles.Selected)=2 then// если 2 то запустить диалог нового процесса
  begin
  NewProcForm.EditFileRun.Text:=EditPath.Text+'\'+ListViewFiles.Selected.Caption;
  NewProcMyPS:=ComboBox2.Text;
  GroupPC:=false; /// то что запуск процесса не групповой, при запуске процесса на одном компе. логин и пароль берутся в потоке с главной формы
  NewProcForm.ShowModal;
  end;
 if FindExt(ListViewFiles.Selected)=3 then  //если каталог  то открыть
 begin
  ListViewFilesDblClick(ListViewFiles);
 end;


end;

procedure TMRPCExplorer.N16Click(Sender: TObject); // свойства файла
var
ind:integer;
begin
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount<>1) then exit;

ind:=TMenuItem(sender).tag;
inc(ind);
TMenuItem(sender).tag:=ind; // увеличения тага на еденицу
propertyFileFolder(EditPath.Text+'\'+ListViewFiles.Selected.Caption, /// путь к файлу или каталогу
                  ListViewFiles.Selected.ImageIndex,                 // тип, файл или каталог, опрелделяется поиндексу картинки
                  MRPCExplorer,                                      // родительская форма
                  ind);                                              // индекс, для создания до хуя форм с разными именами
end;

procedure TMRPCExplorer.N17Click(Sender: TObject);
begin
 if ListViewFiles.Items.Count<>0 then
 frmDomainInfo.popupListViewSaveAs(ListViewFiles,'Сохранение списка файлов и каталогов','Файлы и каталоги');
end;

procedure TMRPCExplorer.PopupMenu1Popup(Sender: TObject);
begin
PopupMenu1.Items[0].Caption:='Открыть на '+ComboBox2.Text; // по умолчанию типа открыть

if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then
PopupMenu1.Items[0].Enabled:=False
else PopupMenu1.Items[0].Enabled:=true;

if PopupMenu1.Items[0].Enabled=true then
begin
 if FindExt(ListViewFiles.Selected)=1 then
 begin
  PopupMenu1.Items[0].Caption:='Установить на '+ComboBox2.Text;
  exit;
 end;
 if FindExt(ListViewFiles.Selected)=2 then  PopupMenu1.Items[0].Caption:='Выполнить на '+ComboBox2.Text;
 end;
end;

procedure TMRPCExplorer.N1Click(Sender: TObject); /// удаление
var
i,z,res:integer;
begin
try
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount=0) then exit;
z:=MessageBox(Self.Handle, PChar('Удалить выбранные файлы?')
        , PChar('Удаление файлов и каталогов') ,MB_YESNO+MB_ICONQUESTION);
   if z=IDNO then exit;
for I := ListViewFiles.Items.Count-1 downto 0 do
  begin
  if ListViewFiles.Items[i].Selected then
   begin
   if ListViewFiles.Items[i].ImageIndex=0 then /// если каталог
   begin
     res:=deleteFolderWmi(EditPath.Text,ListViewFiles.items[i].Caption); // удаляем каталог
     if  res=0 then ListViewFiles.items[i].Delete
     else memolog.Lines.Add('Ошибка удаления '+ListViewFiles.items[i].Caption +' :'+errorExplorer(res));
   end
  else
   begin
     res:=deleteFileWmi(EditPath.Text,ListViewFiles.items[i].Caption); // удаляем файл
     if res=0 then ListViewFiles.items[i].Delete
     else memolog.Lines.Add('Ошибка удаления '+ListViewFiles.items[i].Caption +' :'+errorExplorer(res));
   end;
   end;
  end;
except
    on E: Exception do
     MemoLog.Lines.Add(Format('Ошибка удаления  :  %s',[E.Message]));
  end;
end;

procedure TMRPCExplorer.N2Click(Sender: TObject); //// копировать, копирует директории файлов и кталогов в stringlist для последующей вставки
var
i:integer;
begin
try
if (ListViewFiles.items.Count=0) or (ListViewFiles.SelCount=0) then exit;
FileFolderSourceName.Clear; /// очищаем список файлов и каталогов для копирования
 for I := 0 to ListViewFiles.Items.Count-1 do
   begin
     if ListViewFiles.Items[i].Selected then
     begin
     FileFolderSourceName.add(EditPath.Text+'\'+
     ListViewFiles.Items[i].Caption+'='+inttostr(ListViewFiles.Items[i].ImageIndex));
     end;
   end;
//for I := 0 to FileFolderSourceName.Count-1 do
//MemoLog.Lines.Add(FileFolderSourceName.Names[i]+'='+FileFolderSourceName.ValueFromIndex[i]);
except
    on E: Exception do
     MemoLog.Lines.Add(Format('Ошибка копирования : %s : Trace %s',[E.Message, E.StackTrace]));
  end;
end;

function TMRPCExplorer.CopyOtherThread(SourceFileName,DestFileName,FFname:string;ind,TypeFF:integer):bool;
var
treadID : array [0..2000] of LongWord;
res :array [0..2000] of integer;
begin
/////////////////////////// если копия каталога
if TypeFF=0 then  // каталог
begin
//MemoLog.Lines.Add(FFname);
RunPasteFolder[ind].SourceFileName:= SourceFileName;
RunPasteFolder[ind].DestFileName:= DestFileName;
RunPasteFolder[ind].Recursive:=true;
RunPasteFolder[ind].user:=EditUser.Text;
RunPasteFolder[ind].pass:=EditPass.Text;
RunPasteFolder[ind].NamePC:=ComboBox2.Text;
 res[ind]:=BeginThread(nil,0,addr(PasteFolderWmiThread),Addr(RunPasteFolder[ind]),0,treadID[ind]); ///
 CloseHandle(res[ind]);
end;
/// /////////////////////////////
if TypeFF<>0 then  // файл
begin
//MemoLog.Lines.Add(FFname);
RunPasteFile[ind].SourceFileName:= SourceFileName;
RunPasteFile[ind].DestFileName:= DestFileName;
RunPasteFile[ind].Recursive:=true;
RunPasteFile[ind].user:=EditUser.Text;
RunPasteFile[ind].pass:=EditPass.Text;
RunPasteFile[ind].NamePC:=ComboBox2.Text;
res[ind]:=BeginThread(nil,0,addr(PasteFileWmiThread),Addr(RunPasteFile[ind]),0,treadID[ind]); ///
CloseHandle(res[ind]);
end;


end;

procedure TMRPCExplorer.N3Click(Sender: TObject); // вставить
var
i,z,y:integer;
res :integer;
FFname:string; /// имя каталога или файла

AppOrThread:integer;

function foldername(s:string):string; // функция извлечения имени папки из полного пути
var
i:integer;
begin
for I := Length(s)-1 downto 0  do
if s[i]='\' then
  begin
    result:=copy(s,i+1,length(s));
    break;
  end;
end;
function findFF(nameFF:string;TypeFF:integer):bool;  //функция поиска папки или файла в директории куда вставляем
var
i:integer;
begin
result:=false;
for I := 0 to ListViewFiles.Items.Count-1 do
 if (nameFF=ListViewFiles.Items[i].Caption) and (ListViewFiles.Items[i].ImageIndex=TypeFF) then
 begin
   result:=true;
   break;
 end;
end;
function additemsForFF(NameFF:string;TypeFF:integer):bool; ///функция добавление строки (записи) при удачном копировании
var
itm: TListItem;
begin
itm:=ListViewFiles.Items.Add;
itm.ImageIndex:=TypeFF;
itm.Caption:=NameFF;
itm.SubItems.Add('');
if TypeFF=0 then itm.SubItems.add('Каталог')
else itm.SubItems.add('файл');
itm.SubItems.add('');
end;
BEGIN
AppOrThread:=TMenuItem(sender).tag;// 0 - вставить в текущем потоке, 1 - в отдельном потоке
for I := 0 to FileFolderSourceName.Count-1 do
Begin
//////////////////////////////////////////////////////////////////////////////////////
  try
  if FileFolderSourceName.ValueFromIndex[i]='0' then /// если каталог
    BEGIN
    FFname:= foldername(FileFolderSourceName.Names[i]);
    if findFF(FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) then  // если папка существует в текущей директории
      Begin
       z:=MessageBox(Self.Handle, PChar('Каталог '+FFname+' уже существует, создать копию?'), PChar('Копия или пропуск каталога') ,MB_YESNO+MB_ICONQUESTION);
        if z=IDYES then
         Begin
          y:=0;  // переменная для создания отличительной копии,  далее производим поиск и копирование
          while findFF(FFname+'(копия '+inttostr(y)+')',strtoint(FileFolderSourceName.ValueFromIndex[i])) do begin inc(y); end;
          if AppOrThread=1 then // в отдельном потоке
            CopyOtherThread(FileFolderSourceName.Names[i], Editpath.Text+'\'+FFname+'(копия '+inttostr(y)+')',FFname,i,0) /// 0- это каталог
          else // в текущем потоке
            begin
            res:=PasteFolderWmi(FileFolderSourceName.Names[i], //полный путь к каталогу который копируем
            Editpath.Text+'\'+FFname+'(копия '+inttostr(y)+')',                  // путь до нового каталога
            true);                                     // true . false
            if res=0 then additemsForFF(FFname+'(копия '+inttostr(y)+')',strtoint(FileFolderSourceName.ValueFromIndex[i])) // если копирование удачное то добавляем строку
            else  MemoLog.Lines.Add(errorExplorer(res)); // иначе результат ошибки в лог
            end;
          End;
        End
      else ///иначе если такова каталога нет просто производим копирование
        Begin
        if AppOrThread=1 then // в отдельном потоке
            CopyOtherThread(FileFolderSourceName.Names[i], Editpath.Text+'\'+FFname,FFname,i,0) /// 0- это каталог
          else
            begin
            res:=PasteFolderWmi(FileFolderSourceName.Names[i],//полный путь к каталогу который копируем
            Editpath.Text+'\'+FFname,                       // путь до нового каталога
            true);                                         // true . false
            if res=0 then additemsForFF(FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) // если копирование удачное то добавляем строку
            else  MemoLog.Lines.Add(errorExplorer(res));
            end;
         End;
    END; /// Окончание копирования каталогов
///////////////////////////////////////////////////////////////////////////////////////////////////////
   if FileFolderSourceName.ValueFromIndex[i]<>'0' then /// если не равно 0 то это файл
    BEGIN
     FFname:=ExtractFileName(FileFolderSourceName.Names[i]);  /// имя файла
     if findFF(FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) then  // если файл существует в текущей директории
     Begin
     z:=MessageBox(Self.Handle, PChar('Файл '+FFname+' уже существует, создать копию?'), PChar('Копия или пропуск файла') ,MB_YESNO+MB_ICONQUESTION);
     if z=IDYES then
      begin
        y:=0; // переменная для создания отличительной копии, далее производим поиск и копирование
        while findFF('(копия '+inttostr(y)+') '+FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) do begin inc(y); end;
        if AppOrThread=1 then // в отдельном потоке
          CopyOtherThread(FileFolderSourceName.Names[i], Editpath.Text+'\'+'(копия '+inttostr(y)+') '+FFname,FFname,i,1) /// 1- это файл
        else  // в текущем потоке
          begin
           res:=PasteFileWmi(FileFolderSourceName.Names[i],Editpath.Text+'\'+'(копия '+inttostr(y)+') '+FFname);
          if res=0 then additemsForFF('(копия '+inttostr(y)+') '+FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) // если копирование удачное то добавляем строку
          else MemoLog.Lines.Add(errorExplorer(res));
          end;
        end; //z=IDYES
       End // если файл существует
     else // иначе такого фала нет
      Begin
       if AppOrThread=1 then // в отдельном потоке
          CopyOtherThread(FileFolderSourceName.Names[i],Editpath.Text+'\'+FFname,FFname,i,1) /// 1- это файл
        else
          begin /// в текущем потоке
           res:=PasteFileWmi(FileFolderSourceName.Names[i],Editpath.Text+'\'+FFname);
           if res=0 then  additemsForFF(FFname,strtoint(FileFolderSourceName.ValueFromIndex[i])) // если копирование удачное то добавляем строку
           else MemoLog.Lines.Add(errorExplorer(res));
          end;
       End;  // такого файла нет
    END; // если это файл
////////////////////////////////////////////////////////////////////////////////////
except
    on E: Exception do
     Memolog.lines.add(Format('Ошибка при копировании : %s ',[E.Message]));
   end;
End; // цикл
END;

procedure TMRPCExplorer.N4Click(Sender: TObject); // обновить
begin
 if  (not VarIsClear(FWMIServiceExp))then
 begin
  UpdateFileTreeItem(EditPath.Text);
 end;
end;

procedure TMRPCExplorer.N5Click(Sender: TObject); // переименование
var
i,FFType:integer;
oldname,newName:string;
begin
if VarIsClear(FWMIServiceExp) then exit;
if (ListViewFiles.Items.Count=0) or (ListViewFiles.SelCount<>1) then exit;
oldname:=ListViewFiles.Selected.Caption; // старое имя каталога или файла
FFtype:=ListViewFiles.Selected.ImageIndex; // каталог или файлы
repeat
newName:= InputBox('Переименование', 'Новое имя', oldname);
until newName<>'';
if newName=oldname then exit;
if FFType=0 then // если каталог
begin
if RenameFolderWmi(EditPath.text+'\'+oldname,EditPath.text+'\'+NewName)=0 then // если операция удачна то меняем caption
ListViewFiles.Selected.Caption:=NewName;
end
else  // иначе файл
begin
if RenameFileWmi(EditPath.text+'\'+oldname,EditPath.text+'\'+NewName)=0 then // если операция удачна то меняем caption
ListViewFiles.Selected.Caption:=NewName;
end;
end;

procedure TMRPCExplorer.SpeedButton1Click(Sender: TObject); /// на директорию ниже
var
sPath:string;
Drive:string;
function findslesh(str:string):integer;
var
i:integer;
begin
for I := Length(str)-1 downto 0 do
  if str[i]='\' then
  begin
  result:=i;
  break;
  end;
end;
begin
 if  (not VarIsClear(FWMIServiceExp)) and (Length(EditPath.Text)>3) then
 begin
  sPath:=copy(EditPath.Text,1,findslesh(EditPath.Text)-1);
  //Drive   :=ExtractFileDrive(sPath);
  UpdateFileTreeItem(sPath);
 end;
// MemoLog.Lines.Add(sPath);
end;

procedure TMRPCExplorer.SpeedButton2Click(Sender: TObject); // обновить
begin
 if  (not VarIsClear(FWMIServiceExp)) and (EditPath.Text<>'') then
 begin
  UpdateFileTreeItem(EditPath.Text);
 end;
end;

procedure TMRPCExplorer.FormFindClose(Sender: TObject; var Action: TCloseAction);
begin          /// уничтожение формы  после закрытия
(sender as Tform).Release;
end;

procedure TMRPCExplorer.ListViewFindDblClick(Sender: TObject);
var
item:TListItem;
path,NameFile:string;
begin
try
 if not (sender is Tlistview) then exit;
 if ((sender as Tlistview).Items.Count=0) or ((sender as Tlistview).SelCount<>1) then exit;
 item:=(sender as Tlistview).Selected;
 path:=item.SubItems[0]; // путь
 NameFile:=item.Caption; // каталог или найденный файл
 path:=ExtractFileDir(path);
 if MRPCExplorer.Showing then  //если проводник открыт, то можно искать
 begin
 MRPCExplorer.EditPath.Text:=path;
 MRPCExplorer.speedbutton2click(speedbutton2);/// заходим в директорию файла или каталога
 MRPCExplorer.findFFinListview(NameFile); // ищем файл/каталог в этой директории и выделяем его
 end;
 except
    on E: Exception do
     ShowMessage(Format('Ошибка открытия каталога : %s ',[E.Message]));
  end;
end;

procedure TMRPCExplorer.ButtonFindClick(Sender: TObject);
var
FindForm:Tform;
ListViewFind:Tlistview;
CountNewForm:integer;
begin
try
if EditFind.Text='' then exit;
CountNewForm:=ButtonFind.tag;
inc(CountNewForm);
ButtonFind.tag:=CountNewForm;
FindForm:=Tform.Create(MRPCExplorer);
FindForm.Name:='FindForm'+inttostr(CountNewForm);
FindForm.Position:=poMainFormCenter;
//FormForRDP.OnClose:=OtherWinForRDPClientClose;
FindForm.Width:=650;
FindForm.Height:=350;
FindForm.Tag:= CountNewForm;
FindForm.BorderIcons:=[biSystemMenu,biMinimize];
FindForm.BorderStyle:=bsSizeable;
FindForm.FormStyle:=fsNormal;//fsMDIChild;//fsNormal;//fsMDIForm;//fsStayOnTop;
FindForm.PopupMode:=pmAuto;
FindForm.Position:=poOwnerFormCenter;
FindForm.Caption:='Поиск - '+EditFind.Text;
FindForm.Tag:=CountNewForm;
FindForm.OnClose:=FormFindClose;
ListViewFind:=TListView.Create(FindForm);
ListViewFind.Parent:=FindForm;
ListViewFind.Name:='ListFind'+inttostr(CountNewForm);
ListViewFind.Align:=alClient;
ListViewFind.ViewStyle:=vsReport;
ListViewFind.SmallImages:=ImageList2;
ListViewFind.RowSelect:=true;
ListViewFind.HideSelection:=false;
ListViewFind.OnDblClick:=ListViewFindDblClick;
with ListViewFind.Columns.Add do
begin
 Caption:='Имя';
 Width:=220;
end;
with ListViewFind.Columns.Add do
begin
 Caption:='Путь';
 Width:=430;
end;
RunFindFF[CountNewForm].FindText:='%'+EditFind.Text+'%';
RunFindFF[CountNewForm].FullPath:=MRPCExplorer.EditPath.Text;
RunFindFF[CountNewForm].ListViewForFind:=ListViewFind;
RunFindFF[CountNewForm].NamePC:=MRPCExplorer.ComboBox2.Text;
RunFindFF[CountNewForm].user:=MRPCExplorer.EditUser.Text;
RunFindFF[CountNewForm].pass:=MRPCExplorer.EditPass.Text;
resF[CountNewForm]:=BeginThread(nil,0,addr(FindFileFolderWmiThread),Addr(RunFindFF[CountNewForm]),0,treadIDF[CountNewForm]); ///
CloseHandle(resF[CountNewForm]);
FindForm.Show;
 except
    on E: Exception do
     MemoLog.Lines.Add(Format('Ошибка : %s ',[E.Message]));
  end;
end;

Procedure TMRPCExplorer.TreeViewFoldersChange(Sender: TObject; Node: TTreeNode); //изменения в треевьюв
var
path:string;
begin
  try
    if Assigned(Node) and (Node.Count=0) and not VarIsClear(FWMIServiceExp) then
      UpdateFolderTreeItem(Node); // обновляем список дисков и каталогов если список дисков пуст
    if Assigned(Node) and not VarIsClear(FWMIServiceExp) then
    begin                       // обновляем список каталогов и файлов после обновления списка дисков
    path:='';
    path:=GetNodeFullPath(Node);
     if path<>''then
     UpdateFileTreeItem(path);  // иначе считываем больше информации
    end;
  except
    on E: Exception do
     ShowMessage(Format('Ошибка : %s ',[E.Message]));
  end;
end;

function TMRPCExplorer.RenameFolderWmi(OldName,NewName:string):integer;
var                 // переименование каталога
FWbemObject   : OLEVariant;
Drive,WmiPath,OutFile :string;
begin
try
StatusBar1.Panels[0].Text:=Format('Переименование %s -> %s',[OldName,NewName]);
FWbemObject   := FWMIServiceExp.Get(Format('CIM_Directory.Name="%s"',[StringReplace(OldName,'\','\\',[rfReplaceAll])]));
begin
//MemoLog.Lines.Add('Переименование - '+string(FWbemObject.Name));
result:=FWbemObject.Rename(NewName);
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
if result<>0 then  Memolog.lines.add(Format('Переименование %s - %s :',[OldName,NewName])+errorExplorer(result));
StatusBar1.Panels[0].Text:='';
 except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка : %s ',[E.Message]));
    end;
  end;
end;


function TMRPCExplorer.RenameFileWmi(OldName,NewName:string):integer;
var                 // переименование файла
FWbemObject   : OLEVariant;
Drive,WmiPath,OutFile :string;
begin
try
StatusBar1.Panels[0].Text:=Format('Переименование %s -> %s',[OldName,NewName]);
FWbemObject   := FWMIServiceExp.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(OldName,'\','\\',[rfReplaceAll])]));
begin
//MemoLog.Lines.Add('Переименование - '+string(FWbemObject.Name));
result:=FWbemObject.Rename(NewName);
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
if result<>0 then  Memolog.lines.add(Format('Переименование %s - %s :',[OldName,NewName])+errorExplorer(result));
StatusBar1.Panels[0].Text:='';
 except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
end;

function deleteFolderFileWmiThread(paramDel:pointer):integer;
var                 // удаление каталога/файла в потоке
FSWbemLocator,FWMIService   : OLEVariant;
FWbemObject,StartFileName,StopFileName: OLEVariant;
Drive,WmiPath :string;
Wql:string;
i,z:integer;
MylistDel:TstringList;
begin
OleInitialize(nil);
PointForDelete:=paramDel;
MylistDel:=TStringList.Create;
MylistDel.CommaText:=PointForDelete.SourceFileName; // список файлов / каталогов
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(PointForDelete.NamePC, 'root\CIMV2',PointForDelete.User, PointForDelete.Pass,'','',128);
for I := MylistDel.Count-1 downto 0 do
begin
try
if MylistDel.ValueFromIndex[i]<>'0' then /// удаление файла
begin
MRPCExplorer.StatusBar1.Panels[0].Text:=Format('Удаление %s',[MylistDel.Names[i]]);
Wql:=MylistDel.Names[i];
FWbemObject   := FWMIService.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
MRPCExplorer.MemoLog.Lines.Add('удаляю - '+string(FWbemObject.Name));
result:=FWbemObject.Delete();
FWbemObject:=Unassigned;
if (result<>0) then  MRPCExplorer.Memolog.lines.add(Format('Удаление %s :',[MylistDel.Names[i]])+MRPCExplorer.errorExplorer(result));
MRPCExplorer.StatusBar1.Panels[0].Text:='';
end;
if MylistDel.ValueFromIndex[i]='0'  then // удаление каталога
begin
MRPCExplorer.StatusBar1.Panels[0].Text:=Format('Удаление %s',[MylistDel.Names[i]]);
Wql:=MylistDel.Names[i];
FWbemObject   := FWMIService.Get(Format('CIM_Directory.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
MRPCExplorer.MemoLog.Lines.Add('удаляю - '+string(FWbemObject.Name));
result:=FWbemObject.DeleteEx(StopFileName,StartFileName); //Variants.Null
//result:=FWbemObject.Delete();
FWbemObject:=Unassigned;
if (result<>0) and (StopFileName<>null) then  MRPCExplorer.Memolog.lines.add(Format('Удаление %s :',[string(StopFileName)])+MRPCExplorer.errorExplorer(result))
else
if (result<>0) then  MRPCExplorer.Memolog.lines.add(Format('Удаление %s :',[MylistDel.Names[i]])+MRPCExplorer.errorExplorer(result));
MRPCExplorer.StatusBar1.Panels[0].Text:='';
end;
 except
    on E: Exception do
    begin
     MRPCExplorer.Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
end; {for }
StopFileName:=Unassigned;
VariantClear(StopFileName);
StartFileName:=Unassigned;
VariantClear(StartFileName);
FWbemObject:=Unassigned;
VariantClear(FWbemObject);
FWMIService:=Unassigned;
VariantClear(FWMIService);
FSWbemLocator:=Unassigned;
VariantClear(FSWbemLocator);
MylistDel.Free;
OleUnInitialize;
//NetApiBufferFree(paramDel); /// очищаем память
EndThread(0);    // убиваем поток
end;


function TMRPCExplorer.deleteFileWmi(Path,FileName:string):integer;
var                 // удаление файла
FWbemObject   : OLEVariant;
Drive,WmiPath,OutFile :string;
Wql:string;
begin
try
StatusBar1.Panels[0].Text:=Format('Удаление %s\%s',[Path,FileName]);
Wql:=Path+'\'+FileName;
FWbemObject   := FWMIServiceExp.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
MemoLog.Lines.Add('удаляю - '+string(FWbemObject.Name));
result:=FWbemObject.Delete();
FWbemObject:=Unassigned;
VariantClear(FWbemObject);
if result<>0 then  Memolog.lines.add(Format('Удаление %s\%s :',[Path,FileName])+errorExplorer(result));
StatusBar1.Panels[0].Text:='';
 except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
end;

function TMRPCExplorer.deleteFolderWmi(Path,FolderName:string):integer;
var                 // удаление каталога
FWbemObjectSet: OLEVariant;
FWbemObject,StartFileName,StopFileName: OLEVariant;
Drive,WmiPath :string;
Wql:string;
begin
try
//StopFileName:=null;
//StartFileName:=null;
StatusBar1.Panels[0].Text:=Format('Удаление %s\%s',[Path,FolderName]);
Wql:=Path+'\'+FolderName;
FWbemObject   := FWMIServiceExp.Get(Format('CIM_Directory.Name="%s"',[StringReplace(wql,'\','\\',[rfReplaceAll])]));
MemoLog.Lines.Add('удаляю - '+string(FWbemObject.Name));
result:=FWbemObject.DeleteEx(StopFileName,StartFileName); //Variants.Null
//result:=FWbemObject.Delete();
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
VariantClear(FWbemObjectSet);
VariantClear(FWbemObject);
if (result<>0) and (StopFileName<>null) then  Memolog.lines.add(Format('Удаление %s :',[string(StopFileName)])+errorExplorer(result));
if (result<>0) then  Memolog.lines.add(Format('Удаление %s\%s :',[Path,FolderName])+errorExplorer(result));
StatusBar1.Panels[0].Text:='';
 except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
end;
                                      //SourceFileName,DestFileName:string;Recursive:Boolean
function PasteFolderWmiThread(paramFF:pointer):integer;
var                 // вставить каталог в отдельном потоке
FWbemObject,FSWbemLocator,FWMIService   : OLEVariant;
StopFileName  : string;
Drive,WmiPath,OutFile:string;
Recursive:Boolean;
Wql:string;
res:integer;
begin
try
OleInitialize(nil);
PointForFolder:=paramFF;
MRPCExplorer.StatusBar1.Panels[0].Text:=Format('Копирование каталога %s -> %s',[PointForFolder.SourceFileName,PointForFolder.DestFileName]);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(PointForFolder.NamePC, 'root\CIMV2',PointForFolder.User, PointForFolder.Pass,'','',128);
FWbemObject:=FWMIService.Get(Format('CIM_Directory.Name="%s"',[StringReplace(PointForFolder.SourceFileName,'\','\\',[rfReplaceAll])]));
MRPCExplorer.MemoLog.Lines.Add(Format('Копирование каталога : %s -> %s',[PointForFolder.SourceFileName,PointForFolder.DestFileName]));
if PointForFolder.Recursive then
 res:=FWbemObject.CopyEx(PointForFolder.DestFileName,StopFileName, Variants.Null, Recursive)
 else
 Res:=FWbemObject.copy(PointForFolder.DestFileName);
FWbemObject:=Unassigned;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
MRPCExplorer.StatusBar1.Panels[0].Text:='';
VariantClear(FWbemObject);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
if res=0 then
 MRPCExplorer.Memolog.lines.add(Format('Копирование завершено : %s -> %s',[PointForFolder.SourceFileName,PointForFolder.DestFileName]));
if (res<>0) and (StopFileName<>'') then
MRPCExplorer.Memolog.lines.add(Format('Ошибка копирования : %s',[StopFileName]));
except
    on E: Exception do
    begin
    MRPCExplorer.StatusBar1.Panels[0].Text:='';
    MRPCExplorer.Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
OleUnInitialize;
//NetApiBufferFree(paramFF); /// очищаем память
EndThread(0);    // убиваем поток
end;

function PasteFileWmiThread(paramFF:pointer):integer;
var                 // вставить файл в отдельном потоке
FWbemObject,FSWbemLocator,FWMIService   : OLEVariant;
StopFileName  : string;
res:integer;
begin
try
OleInitialize(nil);
PointForFile:=paramFF;
MRPCExplorer.StatusBar1.Panels[0].Text:=Format('Копирование файла %s -> %s',[PointForFile.SourceFileName,PointForFile.DestFileName]);
MRPCExplorer.MemoLog.Lines.Add(Format('Копирование файла : %s -> %s',[PointForFile.SourceFileName,PointForFile.DestFileName]));
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(PointForFile.NamePC, 'root\CIMV2',PointForFile.User, PointForFile.Pass,'','',128);
FWbemObject:=FWMIService.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(PointForFile.SourceFileName,'\','\\',[rfReplaceAll])]));
Res:=FWbemObject.copy(PointForFile.DestFileName);
FWbemObject:=Unassigned;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
MRPCExplorer.StatusBar1.Panels[0].Text:='';
VariantClear(FWbemObject);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
if res=0 then
 MRPCExplorer.Memolog.lines.add(Format('Копирование завершено : %s -> %s',[PointForFile.SourceFileName,PointForFile.DestFileName]));
if (res<>0) and (StopFileName<>'') then
MRPCExplorer.Memolog.lines.add(Format('Ошибка копирования : %s',[StopFileName]));
except
    on E: Exception do
    begin
    MRPCExplorer.Memolog.lines.add(Format('Ошибка : %s ',[E.Message]));
    end;
  end;
MRPCExplorer.StatusBar1.Panels[0].Text:='';
OleUnInitialize;
//NetApiBufferFree(paramFF); /// очищаем память
EndThread(0);    // убиваем поток
end;

function FindFileFolderWmiThread(paramFF:pointer):integer;
var                 // вставить файл в отдельном потоке
FWbemObject,FSWbemLocator,FWMIService, FWbemObjectSet   : OLEVariant;
oEnum               : IEnumvariant;
iValue        : LongWord;
StopFileName  : string;
res:integer;
wql,Drive,WmiPath:string;
FullPath,NamePC,user,Pass,FindText,sValue:string;
ListViewForFind:Tlistview;
begin
try
OleInitialize(nil);
PointForFind:=paramFF;
FullPath:= PointForFind.FullPath;
NamePC:=PointForFind.NamePC;
user:=PointForFind.user;
Pass:=PointForFind.pass;
FindText:=PointForFind.FindText;
Drive   :=ExtractFileDrive(FullPath);
//WmiPath :=IncludeTrailingPathDelimiter(Copy(Path,3,Length(Path)));
WmiPath :=(Copy(FullPath,3,Length(FullPath)))+'\';
WmiPath :=StringReplace(WmiPath,'\','\\',[rfReplaceAll]);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2',user, Pass,'','',128);
//FWbemObject:=FWMIService.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(PointForFile.SourceFileName,'\','\\',[rfReplaceAll])]));
 Wql     :=Format('SELECT FileName,name FROM CIM_Directory Where Drive="%s" AND Path="%s" AND FileName LIKE "%s"',[Drive,WmiPath,FindText]);
FWbemObjectSet:= FWMIService.ExecQuery(wql,'WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
begin
with PointForFind.ListViewForFind.Items.Add do
begin
 Caption:= string(FWbemObject.Filename) ;
 SubItems.add(string(FWbemObject.name));
 ImageIndex:=0; // каталог
end;
FWbemObject:=Unassigned;
end;
oEnum:=nil;
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
/////////////////////////////// далее поиск файлов
 Wql     :=Format('SELECT FileName,name FROM CIM_DataFile Where Drive="%s" AND Path="%s" AND FileName LIKE "%s"',[Drive,WmiPath,FindText]);
FWbemObjectSet:= FWMIService.ExecQuery(wql,'WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
begin
with PointForFind.ListViewForFind.Items.Add do
begin
 sValue:=Trim(FWbemObject.Name);
 Caption:= string(FWbemObject.Filename) ;
 SubItems.add(string(FWbemObject.name));
 ImageIndex:=MRPCExplorer.GetImageIndexExt(ExtractFileExt(sValue));
end;
FWbemObject:=Unassigned;
end;
oEnum:=nil;
/////////////////////////////////
FWbemObject:=Unassigned;
FWbemObjectSet:=Unassigned;
FWMIService:=Unassigned;
FSWbemLocator:=Unassigned;
VariantClear(FWbemObjectSet);
VariantClear(FWbemObject);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
except
    on E: Exception do
    begin
    ShowMessage(Format('Ошибка поиска : %s ',[E.Message]));
    end;
  end;
OleUnInitialize;
//NetApiBufferFree(paramFF); /// очищаем память
EndThread(0);    // убиваем поток
end;


function TMRPCExplorer.PasteFolderWmi(SourceFileName,DestFileName:string;Recursive:bool):integer;
var                 // вставить каталог
FWbemObject: OLEVariant;
StopFileName  : string;
begin
try
OleInitialize(nil);
MRPCExplorer.StatusBar1.Panels[0].Text:=Format('Копирование каталога %s -> %s',[SourceFileName,DestFileName]);
FWbemObject:=FWMIServiceExp.Get(Format('CIM_Directory.Name="%s"',[StringReplace(SourceFileName,'\','\\',[rfReplaceAll])]));
MRPCExplorer.MemoLog.Lines.Add(Format('Копирование каталога : %s -> %s',[SourceFileName,DestFileName]));
if Recursive then
 result:=FWbemObject.CopyEx(DestFileName,StopFileName, Variants.Null, Recursive)
 else
 Result:=FWbemObject.copy(DestFileName);
FWbemObject:=Unassigned;
MRPCExplorer.StatusBar1.Panels[0].Text:='';
VariantClear(FWbemObject);
except
    on E: Exception do
    begin
    MRPCExplorer.StatusBar1.Panels[0].Text:='';
    MRPCExplorer.Memolog.lines.add(Format('Ошибка : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
OleUnInitialize;
end;





function TMRPCExplorer.PasteFileWmi(SourceFileName,DestFileName:string):integer;
var                 // вставить файл
FWbemObject   : OLEVariant;
StopFileName  : OLEVariant;
begin
try
StatusBar1.Panels[0].Text:=Format('Копирование файла %s -> %s',[SourceFileName,DestFileName]);
FWbemObject:=FWMIServiceExp.Get(Format('CIM_DataFile.Name="%s"',[StringReplace(SourceFileName,'\','\\',[rfReplaceAll])]));
Result:=FWbemObject.copy(DestFileName);
FWbemObject:=Unassigned;
MemoLog.Lines.Add('Файл '+string(DestFileName)+' - скопирован');
VariantClear(FWbemObject);
StatusBar1.Panels[0].Text:='';
except
    on E: Exception do
    begin
     StatusBar1.Panels[0].Text:='';
     Memolog.lines.add(Format('Ошибка копирования файла : %s : Trace %s',[E.Message, E.StackTrace]));
    end;
  end;
end;

procedure TMRPCExplorer.UpdateFileTreeItemForUSBFlas(Path: String);
Var
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  sValue        : String;
  Drive         : String;
  WmiPath       : String;
  Item          : TListItem;
  Wql,ForEdit           : String;
  countFF:integer;
begin
  Drive   :=ExtractFileDrive(Path);
  ForEdit:=Path;
  countFF:=0;
  //WmiPath :=IncludeTrailingPathDelimiter(Copy(Path,3,Length(Path)));
  WmiPath :=(Copy(Path,3,Length(Path)))+'\';
  WmiPath :=StringReplace(WmiPath,'\','\\',[rfReplaceAll]);
  ListViewFiles.Items.Clear;
  ListViewFiles.Items.BeginUpdate;
  try
    //////////////////////////////////////////////
    StatusBar1.Panels[0].Text:=Format('Чтение каталогов %s',[Path]);
    Wql     :=Format('SELECT Name FROM CIM_Directory Where Drive="%s" AND Path="%s"',[Drive,WmiPath]);
    //MemoLog.Lines.Add(Wql);
    FWbemObjectSet:= FWMIServiceExp.ExecQuery(Wql,'WQL',wbemFlagForwardOnly);
    oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      if FWbemObject.Name<>null then
      begin
        Item:=ListViewFiles.Items.Add;
        sValue:=Trim(FWbemObject.Name);
        item.Caption:=ExtractFileName(sValue);
        Item.SubItems.Add('');
        Item.SubItems.Add('Каталог');
        Item.SubItems.Add('');
        item.ImageIndex:=GetImageIndexExt('folder');
        FWbemObject:=Unassigned;
        inc(countFF);// считаем количество каталогов
      end;
    end;
    oEnum:=nil;
    FWbemObjectSet:=Unassigned;
    StatusBar1.Panels[1].Text:='Каталогов - '+inttostr(CountFF); // показываем количество каталогов
    /////////////////////////////////////////////////
    if copy(ForEdit,1,length(ForEdit))='\' then delete(ForEdit,length(ForEdit),1);
    EditPath.Text:=StringReplace(ForEdit,'\\','\',[rfReplaceAll]);; //// путь
    //////////////////////////////////////////
    StatusBar1.Panels[0].Text:=Format('Чтение файлов %s',[Path]);
    Wql:=Format('SELECT Name,FileSize,CreationDate,LastModified,FileType FROM CIM_DataFile Where Drive="%s" AND Path="%s"',[Drive,WmiPath]);
    countFF:=0; // сбрасываем для подсчета файлов
    FWbemObjectSet:= FWMIServiceExp.ExecQuery(Wql,'WQL',wbemFlagForwardOnly);
    oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      if FWbemObject.Name<>null then
      begin
        sValue:=FWbemObject.Name;
        Item:=ListViewFiles.Items.Add;
        Item.Caption:=ExtractFileName(sValue);
        Item.SubItems.Add('');
        if FWbemObject.FileType<>null then Item.SubItems.Add(FWbemObject.FileType)
        else Item.SubItems.Add('');
        if FWbemObject.FileSize<>null then Item.SubItems.Add(FormatFloat('#,',FWbemObject.FileSize))
        else Item.SubItems.Add('');
        Item.ImageIndex:=GetImageIndexExt(ExtractFileExt(sValue));
        FWbemObject:=Unassigned;
        inc(countFF); // считаем файлы
      end;
    end;
    StatusBar1.Panels[2].Text:='Файлов - '+inttostr(CountFF); // показываем количество файлов
    ListViewFiles.Items.EndUpdate;
    StatusBar1.Panels[0].Text:='';
    FWbemObjectSet:=Unassigned;
    VariantClear(FWbemObjectSet);
    oEnum:=nil;
    except
    on E: Exception do
     begin
     Memolog.lines.add(Format('Ошибка обновления файлов/каталогов съемных носителей : %s ',[E.Message]));
     ListViewFiles.Items.EndUpdate;
     end;
   end;
end;

function TMRPCExplorer.findtypedrive(drive:string):integer; // Функция определения типа диска по индексу image в node
var
i:integer;
begin
result:=3;
if TreeViewFolders.Items.Count=0 then
begin
result:=3;
exit;
end;
for I := 0 to TreeViewFolders.Items.Count-1 do
//if TreeViewFolders.Items[i].Text=drive then
if pos(drive,TreeViewFolders.Items[i].Text)<>0 then
begin
result:=TreeViewFolders.Items[i].ImageIndex;
//MemoLog.Lines.Add('drive - '+drive);
//MemoLog.Lines.Add('driveImage - '+inttostr(result));
break;
end;
end;

procedure TMRPCExplorer.UpdateFileTreeItem(Path: String);
Var
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  sValue        : String;
  Drive         : String;
  WmiPath       : String;
  Item          : TListItem;
  Wql,ForEdit   : String;
  CountFF:integer;
function calcKB(sz:OleVariant):integer;
begin
try
result:=((round(sz))div 1024 );
if result=0 then result:=1;
sz:=Unassigned;
VariantClear(sz);
except
result:=1;
end;
end;
begin
  Drive   :=ExtractFileDrive(Path);
  if (findtypedrive(Drive)=2) or (findtypedrive(Drive)=5) or (findtypedrive(Drive)=1) then
  begin
  UpdateFileTreeItemForUSBFlas(Path);
  exit;
  end;
  CountFF:=0;
  ForEdit:=Path;
  //WmiPath :=IncludeTrailingPathDelimiter(Copy(Path,3,Length(Path)));
  WmiPath :=(Copy(Path,3,Length(Path)))+'\';
  WmiPath :=StringReplace(WmiPath,'\','\\',[rfReplaceAll]);
  ListViewFiles.Items.Clear;
  ListViewFiles.Items.BeginUpdate;
  try
    //////////////////////////////////////////////
    StatusBar1.Panels[0].Text:=Format('Чтение каталогов %s',[Path]);
    Wql     :=Format('SELECT Name,CreationDate,LastModified FROM CIM_Directory Where Drive="%s" AND Path="%s"',[Drive,WmiPath]);
   // MemoLog.Lines.Add(Wql);
    FWbemObjectSet:= FWMIServiceExp.ExecQuery(Wql,'WQL',wbemFlagForwardOnly);
    oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      if FWbemObject.Name<>null then
      begin
        Item:=ListViewFiles.Items.Add;
        sValue:=Trim(FWbemObject.Name);
        item.Caption:=ExtractFileName(sValue);
        if FWbemObject.LastModified<>null then Item.SubItems.Add(DateToStr(VarDateToDateTime(FWbemObject.LastModified)))
        else if FWbemObject.CreationDate<>null then Item.SubItems.Add(DateToStr(VarDateToDateTime(FWbemObject.CreationDate)))
        else Item.SubItems.Add('');
        Item.SubItems.Add('Каталог');
        Item.SubItems.Add('');
        item.ImageIndex:=GetImageIndexExt('folder');
        FWbemObject:=Unassigned;
        inc(countff);
      end;
    end;
    oEnum:=nil;
    FWbemObjectSet:=Unassigned;
    StatusBar1.Panels[1].Text:='Каталогов - '+inttostr(CountFF);
    /////////////////////////////////////////////////
    if copy(ForEdit,1,length(ForEdit))='\' then delete(ForEdit,length(ForEdit),1);
    EditPath.Text:=StringReplace(ForEdit,'\\','\',[rfReplaceAll]);; //// путь
    //////////////////////////////////////////
    CountFF:=0;
    StatusBar1.Panels[0].Text:=Format('Чтение файлов %s',[Path]);
    Wql:=Format('SELECT Name,FileSize,CreationDate,LastModified,FileType FROM CIM_DataFile Where Drive="%s" AND Path="%s"',[Drive,WmiPath]);
   // MemoLog.Lines.Add(Wql);
    FWbemObjectSet:= FWMIServiceExp.ExecQuery(Wql,'WQL',wbemFlagForwardOnly);
    oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      if FWbemObject.Name<>null then
      begin
        sValue:=FWbemObject.Name;
        Item:=ListViewFiles.Items.Add;
        Item.Caption:=ExtractFileName(sValue);

        if FWbemObject.LastModified<>null then  Item.SubItems.Add(DateToStr(VarDateToDateTime(FWbemObject.LastModified)))
        else if FWbemObject.CreationDate<>null then  Item.SubItems.Add(DateToStr(VarDateToDateTime(FWbemObject.CreationDate)))
         else  Item.SubItems.Add('');

        if FWbemObject.FileType<>null then Item.SubItems.Add(FWbemObject.FileType)
        else Item.SubItems.Add('');
        if FWbemObject.FileSize<>null then Item.SubItems.Add(inttostr(calcKB(FWbemObject.FileSize))+' Кб')
        else Item.SubItems.Add('');
        Item.ImageIndex:=GetImageIndexExt(ExtractFileExt(sValue));
        FWbemObject:=Unassigned;
        inc(countfF);
      end;
    end;
    ListViewFiles.Items.EndUpdate;
    StatusBar1.Panels[0].Text:='';
    FWbemObjectSet:=Unassigned;
    VariantClear(FWbemObjectSet);
    oEnum:=nil;

    StatusBar1.Panels[2].Text:='Файлов - '+inttostr(CountFF);
   except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка обновления файлов/каталогов : %s ',[E.Message]));
     ListViewFiles.Items.EndUpdate;
    end;
   end;
end;

procedure TMRPCExplorer.UpdateFolderTreeItem(Node: TTreeNode);
Var
  lNode         : TTreeNode;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;

  sValue        : string;
  Path          : string;
  Drive         : string;
  WmiPath       : string;
  Wql           : string;
begin
  if Node=nil then  /// если нод чистый значит загружаем диски
  begin
    TreeViewFolders.Items.BeginUpdate;
    TreeViewFolders.Items.Clear;
    try
      FWbemObjectSet:= FWMIServiceExp.ExecQuery('SELECT Name,DriveType,VolumeName FROM Win32_LogicalDisk','WQL',wbemFlagForwardOnly);
      oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
        sValue:=Trim(FWbemObject.Name);
        if pos('\',svalue)<>0 then
         svalue:=copy(svalue,1,length(svalue)-1); // или svalue :=StringReplace(svalue,'\','',[rfReplaceAll]);
        lNode:=TreeViewFolders.Items.Add(nil,sValue);
        if VarIsNumeric(FWbemObject.DriveType) then
         begin
         if FWbemObject.DriveType=0 then lNode.ImageIndex:=1
          else lNode.ImageIndex:=FWbemObject.DriveType;
         end;
        if FWbemObject.VolumeName<>null then
        if FWbemObject.VolumeName<>'' then lNode.Text:= lNode.Text+' ('+FWbemObject.VolumeName+') ';
        lNode.SelectedIndex:=lNode.ImageIndex;
       // lNode.Selected:=true; // прочитает все вложенные папки каждого диска
        FWbemObject:=Unassigned;
      end;
     if TreeViewFolders.Items.Count>1 then TreeViewFolders.Items[0].Selected:=true; ///прочитать содержимое первого диска
    finally
      TreeViewFolders.Items.EndUpdate;
      FWbemObjectSet:=Unassigned;
      oEnum:=nil;
    end;
  end
  else
  begin
      TreeViewFolders.Items.BeginUpdate;
    try
      Path    :=GetNodeFullPath(Node);
      StatusBar1.Panels[0].Text:=Format('Чтение каталогов %s',[Path]);
      Drive   :=ExtractFileDrive(Path);
      WmiPath :=IncludeTrailingPathDelimiter(Copy(Path,3,Length(Path)));
      WmiPath :=StringReplace(WmiPath,'\','\\',[rfReplaceAll]);
      Wql     :=Format('SELECT Name FROM CIM_Directory Where Drive="%s" AND Path="%s"',[Drive,WmiPath]);
     // MemoLog.Lines.Add(Wql);
      FWbemObjectSet:= FWMIServiceExp.ExecQuery(Wql,'WQL',wbemFlagForwardOnly);
      oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
        sValue:=Trim(FWbemObject.Name);
        lNode:=TreeViewFolders.Items.AddChild(Node,ExtractFileName(sValue));
        lNode.ImageIndex:=0;
        lNode.SelectedIndex:=0;
        FWbemObject:=Unassigned;
      end;
       StatusBar1.Panels[0].Text:='';
       FWbemObjectSet:=Unassigned;
       TreeViewFolders.Items.EndUpdate;
       oEnum:=nil;
    except
    on E: Exception do
    begin
     Memolog.lines.add(Format('Ошибка обновления каталогов : %s ',[E.Message]));
     TreeViewFolders.Items.EndUpdate;
    end;
  end;
  end;
end;

function TMRPCExplorer.VarDateToDateTime(const V: OleVariant): TDateTime;
begin
  Result:=0;
  if VarIsNull(V) then exit;
  Dt.Value := V;
  Result:=Dt.GetVarDate;
end;


procedure TMRPCExplorer.Windows1Click(Sender: TObject); // Открыть проводник Windows
var
z:integer;
function StrRes(z:integer):string;
begin
Result:='';
if z=0 then result:='Оперативной системе не хватает памяти или ресурсов.';
if z=2 then result:='Указаный файл не найден.'  ;
if z=3 then result:='Указаный путь не найден.';
if z=11 then result:='EXE файл не рабочий (не Win32 .EXE или ошибка в .EXE образе).';
if z=5 then result:='Операционная система отказывает в доступе к указаному файлу.';
if z=27 then result:='Имя ассоциированного файла не полное или не правильное.';
if z=30 then result:='Транзакция динамического обмена данными (DDE transaction) не может быть завершена потому что выпоняются другие DDE транзакции.';
if z=29 then result:='Транзакция динамического обмена данными провалилась.';
if z=28 then result:='Транзакция динамического обмена данными не может быть завершена потому что истекло время ожидания ответа.';
if z=32 then result:='Указаная DLL библиотека не найдена.';
if z=2 then result:='Указаный файл не найден.';
if z=31 then result:='Нет приложений ассоциированных с данным расширением файла.';
if z=8 then result:='Не достаточно памяти для завершения операции.';
if z=3 then result:=' Указаный путь не найден.';
if z=26 then result:='Нарушение совместного доступа.';
end;
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
try
if not DirectoryExists('\\'+ComboBox2.Text+'\'+findElement(EditPath.Text)) then // проверяем доступность каталога
begin
ShowMessage('Сетевой путь не найден '+'\\'+ComboBox2.Text+'\'+findElement(EditPath.Text));
exit;
end;
z:=OpenFileOrExplorer('\\'+ComboBox2.Text+'\'+findElement(EditPath.Text));
if StrRes(z)<>'' then showmessage(StrRes(z));
except
    on E: Exception do
     ShowMessage(Format('Ошибка открытия Проводника Windows : %s ',[E.Message]));
  end;
end;

end.




