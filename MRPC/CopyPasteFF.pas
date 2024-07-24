unit CopyPasteFF;
interface
uses Forms, Classes, SysUtils, ShellApi, ShlObj, ClipBrd, Windows;

type
  TWinMoveType = (wCopy = 5, wMove = 2); // Копирование или перенос

  WinExplorer = class
    // файлы в буфер
    class procedure FilesToClipboard(const Files: TStrings;  MoveType: TWinMoveType = wCopy); overload;
    class procedure FilesToClipboard(const Folder: string; Files: TStrings; MoveType: TWinMoveType = wCopy); overload;
    class procedure FilesToClipboard(const FilePath: string; MoveType: TWinMoveType = wCopy); overload;
    // выполнить операцию вставки в папку из буфера
    class procedure ClipboardPaste(const Target: string;FormOwner:Tform;Idn:integer);
    // получить список файлов в буфере
    class procedure ClipBoardGetFiles(const Files: TStrings; h: THandle = 0);
    class function DeleteToRecycle(const FileName: string; Wnd: HWND = 0): Boolean; overload;
    class function DeleteToRecycle(const Files: TStrings; Wnd: HWND = 0): Boolean; overload;
    class function DeleteToRecycle(const Folder: string; const Files: TStrings; Wnd: HWND = 0): Boolean; overload;

  end;
implementation
uses FormCopyFF;


// ------------------------------------------------------------------------------ Str
function Shell_Str(Strs: TStrings): string;
{
  Функция преобразует TStirngs в спец-строку для Shell
  это спец-строка для буфера обмена, где строки разделены знаком #0
  и вся спец-строка строка заканчивается #0#0
}
var
  // s: string; //uses SysUtils
  i: Integer;

begin

  // s := StringReplace(Strs.Text, #13#10, #0, [rfReplaceAll]);
  // s := trim(s) + #0#0;
  // Result := s;

  for i := 0 to Strs.Count - 1 do
    Result := Result + Strs.Strings[i] + #0;

  Result := Result + #0;

end;

// ------------------------------------------------------------------------------ Operation
function Shell_DataOperations(const Source, Target: string;
  Operacia, Flags: Integer): Boolean;
{
  Функция для копирования/вырезания/перименования/удаления данныx средсвами API

  uses ShellAPI

  source - Special Shell Str of Data

  operacia:
  FO_COPY
  FO_MOVE
  FO_RENAME
  FO_DELETE

  flags:
  FOF_ALLOWUNDO         - Если возможно, сохраняет информацию для возможности UnDo. Если вы хотите не просто удалить файлы, а переместить их в корзину, должен быть установлен флаг/
  FOF_CONFIRMMOUSE      - Не реализовано.
  FOF_FILESONLY         - Если в поле pFrom установлено *.*, то операция будет производиться только с файлами.
  FOF_MULTIDESTFILES    - Указывает, что для каждого исходного файла в поле pFrom указана своя директория - адресат.
  FOF_NOCONFIRMATION    - Отвечает "yes to all" на все запросы в ходе опеации.
  FOF_NOCONFIRMMKDIR    - Не подтверждает создание нового каталога, если операция требует, чтобы он был создан.
  FOF_RENAMEONCOLLISION - В случае, если уже существует файл с данным именем, создается файл с именем "Copy #N of..."
  FOF_SILENT            - Не показывать диалог с индикатором прогресса.
  FOF_SIMPLEPROGRESS    - Показывать диалог с индикатором прогресса, но не показывать имен файлов.
  FOF_WANTMAPPINGHANDLE - Вносит hNameMappings элемент. Дескриптор должен быть освобожден функцией SHFreeNameMappings
}

var
  SHOS: TSHFileOpStruct;

begin

  FillChar(SHOS, SizeOf(SHOS), #0);

  SHOS.Wnd := 0;
  SHOS.wFunc := Operacia;
  SHOS.pFrom := PCHAR(Source);
  SHOS.pTo := PCHAR(Target);
  SHOS.fFlags := Flags;

  Result := (SHFileOperation(SHOS) = 0) and (not SHOS.fAnyOperationsAborted);

end;

// ============================================================================== Send
procedure Clipboard_DataSend(const DataPaths: TStrings;
  const MoveType: Integer);
{
  Отправляет файлы/папки в буфер обмена на
  5 = копирование(будто вы нажали Ctrl+C)
  2 = вырезку(будто вы нажали Ctrl+X)  DROPEFFECT_MOVE
  чтоб потом можно было вставить(Ctrl+V) эти данные в любом файловом менеджере.

  uses ShlObj, ClipBrd, Windows;
}

var
  DropFiles: PDropFiles;
  hGlobal: THandle;
  iLen: Integer;
  f: Cardinal;
  d: PCardinal;
  DataSpecialList: string;
  // список адресов(FullPaths) файлов/папок которые надо копировать

begin

  Clipboard.Open;

  Clipboard.Clear;

  // преобразовываем в спец-строку
  DataSpecialList := Shell_Str(DataPaths);

  iLen := Length(DataSpecialList) * SizeOf(Char);
  hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT,
    SizeOf(TDropFiles) + iLen);
  Win32Check(hGlobal <> 0);
  DropFiles := GlobalLock(hGlobal);
  DropFiles^.pFiles := SizeOf(TDropFiles);

{$IFDEF UNICODE}
  DropFiles^.fWide := true;
{$ENDIF}
  Move(DataSpecialList[1], (PansiChar(DropFiles) + SizeOf(TDropFiles))^, iLen);
  SetClipboardData(CF_HDROP, hGlobal);
  GlobalUnlock(hGlobal);

  // FOR COPY
  begin
    f := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
    hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT,
      SizeOf(dword));
    d := PCardinal(GlobalLock(hGlobal));
    d^ := MoveType; // 2-Cut, 5-Copy
    SetClipboardData(f, hGlobal);
    GlobalUnlock(hGlobal);
  end;

  Clipboard.Close;

end;

// ============================================================================== Type
function Clipboard_SendType: Integer;
{
  5 - copy
  2 - cut

  Функция определяет на что посланы данные в буфер: на вырезку или копирование

  Эта функция создаваласть специально для функции ClipBoard_DataPaste,
  чтоб было понятно что делать: копировать или вырезать.
}
var
  ClipFormat, hn: Cardinal;
  szBuffer: array [0 .. 511] of Char;
  FormatID: string;
  pMem: Pointer;

begin
  Result := 0;
  if not OpenClipboard(Application.Handle) then
    exit;
  try
    ClipFormat := EnumClipboardFormats(0);
    while (ClipFormat <> 0) do
    begin
      GetClipboardFormatName(ClipFormat, szBuffer, SizeOf(szBuffer));
      FormatID := string(szBuffer);
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

{ WinExplorer }

class procedure WinExplorer.FilesToClipboard(const Files: TStrings;
  MoveType: TWinMoveType);
begin
  Clipboard_DataSend(Files, Integer(MoveType));
end;

class procedure WinExplorer.FilesToClipboard(const Folder: string; Files: TStrings;
  MoveType: TWinMoveType);
var
  i: Integer;
  s: TStringList;
begin
  s := TStringList.Create;
  for i := 0 to Files.Count - 1 do
    s.Add(Folder + Files[i]);
  Clipboard_DataSend(s, Integer(MoveType));
  s.free;
end;

class procedure WinExplorer.ClipBoardGetFiles(const Files: TStrings; h: THandle);
{
  Вы нажали Ctrl+C или Ctrl+X => послали данные в буфер обмена.
  Так вот эта функция возвращает список файлов/папок, которые посланы в буфер.
}
var
  FilePath: array [0 .. MAX_PATH] of Char;
  i, FileCount: Integer;
begin
  Files.Clear;
  if h = 0 then
  begin
    if not Clipboard.HasFormat(CF_HDROP) then
      exit;
    Clipboard.Open;
    try
      h := Clipboard.GetAsHandle(CF_HDROP);
    finally
      Clipboard.Close;
    end;
  end;
  if h = 0 then
    exit;
  FileCount := DragQueryFile(h, $FFFFFFFF, nil, 0);
  for i := 0 to FileCount - 1 do
  begin
    DragQueryFile(h, i, FilePath, SizeOf(FilePath));
    Files.Add(FilePath);
  end;
end;

class procedure WinExplorer.ClipboardPaste(const Target: string;FormOwner:Tform;Idn:integer);
{
  Эта функция вставит из буфера обмена файлы/папки,
  которые копировали/вырезали(Ctrl+C / Ctrl+X) в буфер в каком-либо фаловом менеджере(Проводник, ТоталКоммандер)

  Target - папка в которую будет вставлены данные
  Clipboard_OperationType - подфункция которая определяет что надо сделать: Копировать или Вырезать
}

var
  h: THandle;
  Sourse, sr: string;
  s: TStringList;
begin

  // Если то, что находиться в буфере
  // НЕ является файлами/папками, которые копированы/вырезаны, то выходим
  // CF_HDROP - дескриптор который идентифицирует список файлов.
  // Прикладная программа может извлечь информацию о файлах, передавая дескриптор функции DragQueryFile.
  if not Clipboard.HasFormat(CF_HDROP) then
   exit;
  Clipboard.Open;
  try
    h := Clipboard.GetAsHandle(CF_HDROP);
    if h <> 0 then
    begin
      s := TStringList.Create;
      WinExplorer.ClipBoardGetFiles(s, h);
      Sourse := Shell_Str(s);
      s.free;

      sr := Copy(Sourse, 0, Pos(#0, Sourse) - 1); // Path №1
      sr := ExtractFilePath(sr); // Родительская папка Data

      if IncludeTrailingBackslash(sr) = IncludeTrailingBackslash(Target)

      then // Делаем копию фала: откуда copy туда и paste
      begin
        case Clipboard_SendType of
          5: Form11.CopyFileFolder(Sourse,Target,false,FormOwner,Idn,2); // 2- копия //в потоке
          2: Form11.CopyFileFolder(Sourse,Target,false,FormOwner,Idn,1); // 1- вставить //в потоке
        end;
      end
      else
      begin
        case Clipboard_SendType of
          5: Form11.CopyFileFolder(Sourse,Target,false,FormOwner,Idn,2); // 2- копия  /в потоке
          2: Form11.CopyFileFolder(Sourse,Target,false,FormOwner,Idn,1); // 1- вставить /в потоке
        end;
      end;

    end;

  finally
    Clipboard.Close;
  end;
end;

class function WinExplorer.DeleteToRecycle(const Files: TStrings;
  Wnd: HWND): Boolean;
begin
  result:=WinExplorer.DeleteToRecycle(Shell_Str(Files),wnd);
end;

class procedure WinExplorer.FilesToClipboard(const FilePath: string;
  MoveType: TWinMoveType);
var
  s: TStringList;
begin
  s := TStringList.Create;
  s.Add(FilePath);
  Clipboard_DataSend(s, Integer(MoveType));
  s.free;
end;

class function WinExplorer.DeleteToRecycle(const FileName: string; Wnd: HWND): Boolean;
var
  FileOp: TSHFileOpStruct;
begin
  FillChar(FileOp, SizeOf(FileOp), 0);
  if Wnd = 0 then
    Wnd := Application.Handle;
  FileOp.Wnd := Wnd;
  FileOp.wFunc := FO_DELETE;
  FileOp.pFrom := PChar(FileName);
  FileOp.fFlags := FOF_ALLOWUNDO or FOF_NOERRORUI or FOF_SILENT;
  Result := (SHFileOperation(FileOp) = 0) and (not FileOp.fAnyOperationsAborted);
end;

class function WinExplorer.DeleteToRecycle(const Folder: string;
  const Files: TStrings; Wnd: HWND): Boolean;
var
  i: Integer;
  s: TStringList;
begin
  s := TStringList.Create;
  for i := 0 to Files.Count - 1 do
    s.Add(Folder + Files[i]);
  result:=self.DeleteToRecycle(s,Wnd);
  s.free;
end;

end.
