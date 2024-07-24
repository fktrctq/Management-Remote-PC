unit FormCopyFF;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,ShellAPI,ClipBrd,ShlObj,IdIcmpClient,
  Vcl.ExtCtrls;

const
  netapi32lib = 'netapi32.dll';
  NERR_Success = NO_ERROR;
  wbemFlagForwardOnly = $00000020;
  wbemFlagReturnImmediately=16;

type
  TForm11 = class(TForm)
    TaskDialog1: TTaskDialog;
    procedure CancelCopy(Sender: TObject);
    procedure TaskDialog1VerificationClicked(Sender: TObject);
    procedure FormFFClose(Sender: TObject; var Action: TCloseAction);

  private
  //  procedure SetPosition(APos: Int64);
  public
  function CopyFileFolder(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean; /// для запуска копирования в потоке
  function FunctionCopyFF(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):integer; // копирование не в потоке
  procedure Clipboard_DataSend(const DataPaths: TStrings;const MoveType: Integer);
  function FunctionStartThreadCopyFileEx(FSource,FDest,username,passwd:string;
CancelCopyFF:boolean;CountNewForm:integer;ListPC:TstringList):boolean; /// функция для запуска CopyFileEx копирования в потоке
  function CopyFileFolderListPC(FSource,FDest,userPC,Paswd,StrListPC:string; // функция запуска копирования в потоке для группы компьютеров
  FindCreateFolder,CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
  end;
type
 TstrForCopy = record
    FSource :String;   /// источник файл
    FDest:string;      //файл назначение
    ListPC:string;
    UserName:string;
    PassWd:string;
    PathCreate:bool;  // проверка и создания директории
    CancelCopyFF:boolean;   // отменить или нет операцию копирования
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    ProgBar:TprogressBar;      // прогрусс бар, кудпа выводить статус копирования
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
  end;
type
SHFILEOPSTRUCTA = packed record
    Wnd: HWND;
    wFunc: UINT;
    pFrom: PAnsiChar;
    pTo: PAnsiChar;
    fFlags: FILEOP_FLAGS;
    fAnyOperationsAborted: BOOL;
    hNameMappings: Pointer;
    lpszProgressTitle: PAnsiChar; { используется только при установленном флаге FOF_SIMPLEPROGRESS }
  end;
  function ThreadCopyFileEx(paramCopyFF:pointer):boolean; // копирование в потоке c выводом прогресс бара
  function ThreadCopyFF(paramCopyFF:pointer):boolean;
  //function NetApiBufferFree(Buffer: Pointer): DWORD; stdcall;
  //  external netapi32lib;
var
  Form11: TForm11;
  RunCopyFF         :array [0..2000] of TstrForCopy;///указатель для копирования в потоке
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;
  CopyCancelFF: array [0..2000] of boolean;

ThreadVar
PointForCopyFF: ^TstrForCopy;

implementation
uses umain;

{$R *.dfm}
procedure TForm11.FormFFClose(Sender: TObject; var Action: TCloseAction); // уничтожение формы при закрытии
begin
(sender as Tform).Release;
end;



function finditem(s,name:string;image:integer):boolean;
var
z:integer;
begin
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
    if frmDomainInfo.ListView8.Items[z].SubItems[0]=name then
    begin
    frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=image;
    frmDomainInfo.ListView8.Items[z].SubItems[1]:=s;
    break;
    end;
end;

 function ping(s:string):boolean;
var
z:integer;
Myidicmpclient:TIdIcmpClient;
begin
try
result:=false;
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  finditem('Превышен интервал ожидания запроса',s,2);
  end
else
  begin
  result:=true; ///доступен
  frmDomaininfo.Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;
   except
    begin
    result:=false;
    finditem('Узел не доступен',s,2);
    end;
   end;
if Assigned(MyIdIcmpClient) then freeandnil(MyIdIcmpClient);
end;

function Shell_Str(Strs: TStrings): string;
{
  Функция преобразует TStirngs в спец-строку для Shell
  это спец-строка для буфера обмена, где строки разделены знаком #0
  и вся спец-строка строка заканчивается #0#0
}
var
  i: Integer;
begin
  for i := 0 to Strs.Count - 1 do
    Result := Result + Strs.Strings[i] + #0;
  Result := Result + #0;
end;

procedure TForm11.CancelCopy(Sender: TObject);
var
NumThread:integer;
begin
if (sender as Tbutton).owner is Tform then
begin
  NumThread:=((sender as Tbutton).owner as Tform).tag;
 // RunCopyFF[NumThread].CancelCopyFF:=true;
  //CopyCancelFF[NumThread]:=true;
  //CloseHandle(res[NumThread]);
 //((sender as Tbutton).owner as Tform).Release;
end;
end;

function FindAdncreateDir(path,NamePC:string):bool;// проверка и создание директории
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //если нет каталога
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// создает весь путь до папки
     begin
     finditem('Создание директории ' +ExtractFileDir(path)+' : произошла ошибка ',NamePC,2);
     result:=false;
     end
     else result:=true; // директория создана
    end
    else result:=true; // директория есть
  except on E: Exception do
     begin
     frmdomaininfo.Memo1.Lines.Add(NamePC+' : Ошибка создания директории - '+e.Message);
     finditem('Создание директории ' +ExtractFileDir(path)+' :' +e.Message,NamePC,2);
     result:=false;
     end;
   end;
end;

///////////////////////////////////////////////////////////
function ThreadCopyFileEx(paramCopyFF:pointer):boolean; // копирование в потоке c выводом прогресс бара
 var
 i:integer;
 htoken:THandle;
 step:string;
 listpc:tstringlist;
 function CopyCallBack (  /// функция актуальна для CopyFileEx
 TotalFileSize,
 TotalBytesTransferred,
 StreamSize,
 StreamBytesTransferred: Int64;
 StreamNumber,
 CallBackReasom: DWORD;
 SrcFile,
 DestFile: THandle): DWORD; stdcall;
begin
   PointForCopyFF.ProgBar.Position:=TotalBytesTransferred;
end;
begin
try
listpc:=tstringlist.Create;
PointForCopyFF:=paramCopyFF;
listpc.CommaText:=PointForCopyFF.ListPC;
for I := 0 to ListPC.Count-1 do
if ping(ListPC[i]) then
BEGIN
 try
 ////////////////////////////////////// авторизация пользователя на удаленном компе
   try
   if not (LogonUserA (PAnsiChar(PointForCopyFF.UserName), PAnsiChar (ListPC[i]),  // сначала заходим на комп в сети
   PAnsiChar (PointForCopyFF.PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
   LOGON32_PROVIDER_WINNT50, hToken))
   then GetLastError();
   except on E: Exception do frmdomaininfo.Memo1.Lines.Add(ListPC[i]+' : Ошибка LogonUser - '+e.Message)
   end;
   //////////////////////
  // if FindAdncreateDir(PointForCopyFF.FDest,ListPC[i]) then // если каталог есть или создан. то копируем
  //перед копированием созадть директорию если необходимо
   begin
      frmdomaininfo.Memo1.Lines.Add(listpc[i]+'  копирую- ''\\'+ListPC[i]+'\'+PointForCopyFF.FDest);
      if  CopyFileEx(PChar(PointForCopyFF.FSource), //lpExistingFileName - источник
                    PChar('\\'+ListPC[i]+'\'+PointForCopyFF.FDest),   //pNewFileName - назначение  // куда копируем  \\компьютер\c$\temp\
                    nil,//@CopyCallBack,                 //@CopyCallBack - функция куда передается значение прогресса копирования
                    nil,                           //pData - данные для предыдущей функции
                    @CopyCancelFF[PointForCopyFF.NumCount],   // pbCancel - флаг для остановки копирования
                    0)     // dwCopyFlags -  COPY_FILE_FAIL_IF_EXISTS - Если файл назначения существует, то копирование не происходит. COPY_FILE_RESTARTABLE - Указывает, что если в процессе копирования произойдёт ошибка, то копирование можно будет продолжить позднее, указав для lpExistingFileName, lpNewFileName те же значения что были в предыдущем вызове функции.
      then finditem('Копирование файла ' +ExtractFileName(PointForCopyFF.FSource)+' : Операция завершена успешно',ListPC[i],1)
      else
      begin
      GetLastError(); // иначе узнаем почему не получилось скопировать
      finditem('Копирование файла ' +ExtractFileName(PointForCopyFF.FSource)+' : произошла ошибка ',ListPC[i],2)
      end;
    end; // конец копирования файла
    except
      on E: Exception do
      begin
       finditem('Копирование файла ' +ExtractFileName(PointForCopyFF.FSource)+': Ошибка ThreadCopyFileEx  : '+E.Message,ListPC[i],2);
       frmdomaininfo.Memo1.Lines.Add(Format(step+' : '+ListPC[i]+': Ошибка ThreadCopyFileEx  :  %s',[E.Message]));
       end;
    end;
CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
END;

ListPC.Free;
except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format(step+' : Ошибка ThreadCopyFileEx  :  %s',[E.Message]));
  end;
EndThread(0);    // убиваем поток
end;

function TForm11.FunctionStartThreadCopyFileEx(FSource,FDest,username,passwd:string;
CancelCopyFF:boolean;CountNewForm:integer;ListPC:TstringList):boolean;
///////////////////////////////                             /// функция запуска потока копирования через CopeFileEx
 function FileSize (const AFileName: string): Int64;
 var
   FS: TFileStream;
 begin
   FS := TFileStream.Create(AFileName,fmOpenRead);
   try
     Result := FS.Size
   finally
     FS.Free;
   end;
 end;
begin
try
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // возможность отмены
  RunCopyFF[CountNewForm].ListPC:=ListPC.CommaText;
  RunCopyFF[CountNewForm].UserName:=username;
  RunCopyFF[CountNewForm].PassWd:=passwd;
  //RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// какая форма роительская для прогресс быра, нужда только для уничтожения формы после копирования
 // RunCopyFF[CountNewForm].ProgBar:=FProgressBar;  // в каком процесс баре отображать процесс выполнения, актуальна для CopyFileEx
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFileEx),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format(' - Ошибка FunctionStartThreadCopyFileEx  :  %s',[E.Message]));
  end;
 end;

procedure TForm11.TaskDialog1VerificationClicked(Sender: TObject);
begin

end;

//////////////////////////////////////////////////////////////

function ThreadCopyFF(paramCopyFF:pointer):boolean; // копирование в потоке для одного компа
var
SHFileOpStruct : TSHFileOpStruct;
begin
PointForCopyFF:=paramCopyFF;
///////////////////////////// новая функция начиная с vista , копирует файлы и каталоги
  with SHFileOpStruct do
    begin
      Wnd := PointForCopyFF.OwnerForm.Handle;
      case PointForCopyFF.TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      pFrom := pchar(PointForCopyFF.FSource);
      pTo := pchar(PointForCopyFF.FDest);
      fFlags := 0;//FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR or FOF_RENAMEONCOLLISION or FOF_SIMPLEPROGRESS;
      fAnyOperationsAborted := PointForCopyFF.CancelCopyFF;
      hNameMappings := nil;
      lpszProgressTitle :=nil;
    end;
  SHFileOperation( SHFileOpStruct);//непосредственно функция копирования/перемещения
EndThread(0);    // убиваем поток
end;

///////////////////////////////////////////////////////////////////////
function ThreadCopyFFSelectPC(paramCopyFF:pointer):boolean; // копирование в потоке для групп  компьютеров
var
SHFileOpStruct : TSHFileOpStruct;
ListPC,ListDel:Tstringlist;
rescopy,i,z:integer;
htoken:THandle;
begin
try
PointForCopyFF:=paramCopyFF;
try
ListPC:=tstringlist.Create;
ListPC.CommaText:=PointForCopyFF.ListPC; // список компьютеров
///////////////////////////// новая функция начиная с vista , копирует файлы и каталоги для группы компьютеров
  with SHFileOpStruct do
    begin
      Wnd := PointForCopyFF.OwnerForm.Handle;
      case PointForCopyFF.TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4: wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := PointForCopyFF.CancelCopyFF;
      hNameMappings := nil;                //Когда функция возвращается, этот член содержит дескриптор объекта сопоставления имен, который содержит старые и новые имена переименованных файлов
      lpszProgressTitle :=nil;            // Указатель на заголовок диалогового окна прогресса
    end;
    for I := 0 to ListPC.Count-1 do
    if ping(ListPC[i]) then
    Begin

     try ////////////////////////////////////// авторизация пользователя на удаленном компе
     if not (LogonUserA (PAnsiChar(PointForCopyFF.UserName), PAnsiChar (ListPC[i]),  // сначала заходим на комп в сети
     PAnsiChar (PointForCopyFF.PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do frmdomaininfo.Memo1.Lines.Add(ListPC[i]+' : Ошибка LogonUser - '+e.Message)
     end;
     ////////////////////////////////////////////
     try
     if PointForCopyFF.PathCreate then FindAdncreateDir(PointForCopyFF.FDest,ListPC[i]); //проверяем и создаем каталог если его нет.

     if PointForCopyFF.TypeOperation=2 then //если операция копирования то указываем
      begin
      SHFileOpStruct.pTo := pchar('\\'+ListPC[i]+'\'+PointForCopyFF.FDest);   // куда копируем
      SHFileOpStruct.pFrom := pchar(PointForCopyFF.FSource); //если операция копирования, от куда копируем
      rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения/удаления
      if rescopy=0 then finditem('Операция копирования успешно завершена',ListPC[i],1)
      else finditem('Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',ListPC[i],2);
      end;

      if PointForCopyFF.TypeOperation=3 then //если операция удаления
       begin
       ListDel:=TStringList.Create;
       try
       ListDel.CommaText:=PointForCopyFF.FDest;
       for z := 0 to ListDel.Count-1 do
        begin
        SHFileOpStruct.pTo := pchar('');  // то тут пусто
        SHFileOpStruct.pFrom:= pchar('\\'+ListPC[i]+'\'+ListDel[z]+#0+#0); // тут что удаляем , добавляем символы 0 т.к. передать список с нулями не получается
        rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения/удаления
        if rescopy=0 then finditem('Операция удаления успешно завершена',ListPC[i],1)
        else finditem('Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',ListPC[i],2);
        end;
       finally
       ListDel.Free;
       end;
       end;

      CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
      except
      on E: Exception do
      begin
      case PointForCopyFF.TypeOperation of
       2:begin
       frmdomaininfo.Memo1.Lines.Add(Format('Ошибка оперции копирования на компьютер  %s :  %s',[listpc[i],E.Message]));
       finditem('Ошибка оперции копирования: '+E.Message,ListPC[i],2);
       end;
       3:begin
        frmdomaininfo.Memo1.Lines.Add(Format('Ошибка оперции удаления на компьютере  %s :  %s',[listpc[i],E.Message]));
       finditem('Ошибка оперции удаления: '+E.Message,ListPC[i],2);
       end;
      end;
      end;
      end;
    End; //цикл
finally
  ListPC.Free;
end;
except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format('Ошибка функции копирования/удаления на группе компьютеров  :  %s',[E.Message]));
  end;
EndThread(0);    // убиваем поток
end;
////////////////////////////////////////////////////////////////////////



function TForm11.FunctionCopyFF(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):integer;
var       //функция копирования/перемещения не в потоке
SHFileOpStruct : TSHFileOpStruct;
FStyleY:bool;
begin
  with SHFileOpStruct do
    begin
      FStyleY:=false;
      if OwnerForm.FormStyle=fsStayOnTop then
      begin
      OwnerForm.FormStyle:=fsNormal; // для того чтобы форма не скрывалась
      FStyleY:=true;
      end;
      Wnd :=OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      pFrom := pchar(FSource);
      pTo := pchar(FDest);
      fFlags := 0;//FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR or FOF_RENAMEONCOLLISION or FOF_SIMPLEPROGRESS;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;
      lpszProgressTitle :=nil;
    end;
    result:=SHFileOperation( SHFileOpStruct ); //непосредственно функция копирования/перемещения
  if FStyleY then OwnerForm.FormStyle:=fsStayOnTop;
end;


function TForm11.CopyFileFolder(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
///////////////////////////////                             /// функция запуска вставки в потоке  SHFileOperation
begin
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // возможность отмены
  RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// какую форму делать роительской для прогресс бара винды
  RunCopyFF[CountNewForm].NumCount:=CountNewForm;
  RunCopyFF[CountNewForm].TypeOperation:=TypeOperation;
  //CopyCancelFF[CountNewForm]:=false;
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFF),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 end;

 function TForm11.CopyFileFolderListPC(FSource,FDest,userPC,Paswd,StrListPC:
 string;FindCreateFolder,CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
///////////////////////////////                             /// функция запуска вставки/копировани/удаления/переименования в потоке для группы компьютеров SHFileOperation
begin
  RunCopyFF[CountNewForm].UserName:=userPC;
  RunCopyFF[CountNewForm].PassWd:=Paswd;
  RunCopyFF[CountNewForm].ListPC:=StrListPC;
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // возможность отмены
  RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// какую форму делать роительской для прогресс бара винды
  RunCopyFF[CountNewForm].NumCount:=CountNewForm;
  RunCopyFF[CountNewForm].TypeOperation:=TypeOperation;
  RunCopyFF[CountNewForm].PathCreate:=FindCreateFolder;
  //CopyCancelFF[CountNewForm]:=false;
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFFSelectPC),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 end;


procedure TForm11.Clipboard_DataSend(const DataPaths: TStrings;const MoveType: Integer);
{  Отправляет файлы/папки в буфер обмена на
  5 = копирование(будто вы нажали Ctrl+C)
  2 = вырезку(будто вы нажали Ctrl+X)  DROPEFFECT_MOVE
  чтоб потом можно было вставить(Ctrl+V) эти данные в любом файловом менеджере.
  uses ShlObj, ClipBrd, Windows; }
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
  // преобразовываем строку в спец-строку
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


end.
