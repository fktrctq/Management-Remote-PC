unit KMS;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.ExtCtrls,Inifiles
  ,ShellAPI,ActiveX,ComObj, IdComponent, IdHTTP, JvRichEdit,Wininet,
  JvExStdCtrls, IdBaseComponent, IdTCPConnection, IdTCPClient;

type
  TFormKMS = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Panel1: TPanel;
    LabeledEdit4: TLabeledEdit;
    Button4: TButton;
    Edit2: TEdit;
    Edit4: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
    Button5: TButton;
    IdHTTP1: TIdHTTP;
    StatusBar1: TStatusBar;
    memo1: TJvRichEdit;
    LinkLabel1: TLinkLabel;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure memo1URLClick(Sender: TObject; const URLText: string;
      Button: TMouseButton);
    procedure LinkLabel1Click(Sender: TObject);
  private
     MemIni:Tmeminifile;
     ProgressB:TprogressBar;
     function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
     TypeOperation:integer;CancelCopyFF,PathCreate:boolean):boolean; // функция копирования
     function MyNewProcess( NamePC,UserName,PassWd,FileToRun:String):boolean; // функция запуска
     procedure HttpWork(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
     function ExistFileKMS(ListFile:TstringList):boolean; // проверка наличия файлов
     function LoadFile(url,path,Fname:string):boolean; //Непосредственно загрузка и сохранение файла
     procedure CreateFormElseErrorloadFile; // форма если не получилось загрузить файлы
     procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
     procedure memotext;
     function loadFileRun :boolean; //функция загрузки файла настроек и необходимых файлов для KMS
     procedure RereadFile; //перечитывем файл настроек
    public

  end;


var
  FormKMS: TFormKMS;


implementation
uses umain;
{$R *.dfm}


 function FindAddcreateDir(path,NamePC:string):boolean;// проверка и создание директории
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //если нет каталога
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// создает весь путь до папки
     result:=false
     else result:=true; // директория создана
    end
    else result:=true; // директория существует
  except on E: Exception do
     begin
     ShowMessage(e.Message);
     result:=false;
     end;
   end;
end;

function TFormKMS.CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean):boolean; // копирование
var
SHFileOpStruct : TSHFileOpStruct;
rescopy:integer;
htoken:THandle;
begin
try
///////////////////////////// новая функция начиная с vista , копирует файлы и каталоги для группы компьютеров
  with SHFileOpStruct do
    begin
      Wnd := OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - удалить, FO_MOVE - переместить ,FO_RENAME - переименовать
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;                //Когда функция возвращается, этот член содержит дескриптор объекта сопоставления имен, который содержит старые и новые имена переименованных файлов
      lpszProgressTitle :=nil;            // Указатель на заголовок диалогового окна прогресса
     if TypeOperation=2 then
     begin
      pFrom := pchar(FSource); // от куда копируем если операция копирования
      pTo := pchar('\\'+CurentPC+'\'+FDest);   // куда копируем
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // что удалять  если операция удаление
      pTo := pchar('');   // куда копируем не используется
      //finditem('','Запущена операция удаления дистрибутива',CurentPC,NumTask);
     end;
    end;

    try ////////////////////////////////////// авторизация пользователя на удаленном компе
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // сначала заходим на комп в сети
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do begin result:=false; memo1.Lines.Add(CurentPC+' - Ошибка: '+e.Message) end;
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC); //проверяем и создаем каталог если его нет.
     rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения /удаления
     if rescopy=0 then result:=true
     else begin result:=false;  end;
     memo1.Lines.Add(CurentPC+' - Копирование файлов: '+SysErrorMessage(rescopy));
     CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
      except
      on E: Exception do
      begin
      if TypeOperation=2 then
       Memo1.lines.add(CurentPC+' - Ошибка оперции копирования: '+E.Message);
      if TypeOperation=3 then
       Memo1.lines.add(CurentPC+' - Ошибка оперции удаления: '+E.Message);
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then Memo1.lines.add(CurentPC+' - Общая ошибка функции копирования перед запуском процесса  : '+E.Message);
     if TypeOperation=3 then Memo1.lines.add(CurentPC+' - Общая ошибка функции удаления дистрибутива после запуска процесса  : '+E.Message);
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////
function TFormKMS.MyNewProcess(
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    FileToRun:String   // файл для запуска
    ):boolean;
var
MyError:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObject   : OLEVariant;
objProcess    : OLEVariant;
objConfig     : OLEVariant;
ProcessID,z,i   : Integer;
const
wbemFlagForwardOnly = $00000020;
HIDDEN_WINDOW       = 1;
/////////////////////////////////////////////////////////////////////
BEGIN
  try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', UserName, PassWd,'','',128);
  FWMIService.security_.AuthenticationLevel:=6;
  FWMIService.security_.ImpersonationLevel:=3;
  //FWMIService.security_.Privileges.AddAsString('SeEnableDelegationPrivilege');
  FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
  objConfig     := FWbemObject.SpawnInstance_;
  objConfig.ShowWindow := HIDDEN_WINDOW;
  objProcess    := FWMIService.Get('Win32_Process');
  MyError:=objProcess.Create(FileToRun, null, objConfig, (ProcessID));
  if MyError=0 then result:=true
  else result:=false;
  memo1.Lines.Add(NamePC+' - Запуск процесса '+FileToRun+' : '+syserrormessage(MyError));
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  VariantClear(FWbemObject);
  VariantClear(objConfig);
  VariantClear(objProcess);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
    except
      on E:Exception do
      begin
      Memo1.lines.add(NamePC+' - Ошибка запуска процесса : '+E.Message);
      result:=false;
      OleUnInitialize;
      end;
    end;
END;
///////////////////////////////////////////////////////////



function GetLastDir(path:String):String;  // извлечение корневой папки
var
i,j:integer;
begin
path:=ExtractFilePath(path); //Путь до файла меняем на path  файла
if Length(path)>3 then
    begin
      for I :=Length(path)-1 downto 1 do
      if path[i]='\' then
      begin
      j:=i+1;
      break;
      end;
      result:=Copy(path,j,Length(path)) ;
end
else
result:=path[1]+path[3];
end;


procedure TFormKMS.Button2Click(Sender: TObject); // Активация на 180 дней
{Без присмотра (автоматический выход):
Activate.cmd /u
Без звука (подразумевается без присмотра):
Activate.cmd /s
Тихо и создать простой журнал (обязательный порядок):
Activate.cmd /s /l
Режим отладки (подразумевает автоматический):
Activate.cmd /d
Тихий режим отладки (обязательный порядок):
Activate.cmd /s /d}
var
i:integer;
begin
if not FileExists(labelEdEdit2.text) then
begin
i:=MessageDlg('Не найдены скрипты и файлы для работы.'+#10#13+' Загрузить необходимые файлы?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
if not loadFileRun then exit;
end;
Memo1.Clear;
if CopyFFSelectPC
(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
labelEdEdit1.text+#0+#0,        // что копируем
'C$\TEMP\'+#0+#0,              // куда копируем
FormKMS,                       // форма владелец
2,                            // операция копировать
false,                       // возможность отмены
true) then                      // проверять наличие каталога назначения
begin
if MyNewProcess(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
'C:\TEMP\'+GetLastDir(labelEdEdit2.text)+ExtractFileName(labelEdEdit2.text)+' '+Edit2.Text) then //   /s /l - тихая с логированием и удалением родительского каталога скрипта, если запуск без ключей или с одним ключем то удаляться не будет
memo1.Lines.Add(frmdomaininfo.Combobox2.text+' - Запуск процесса активации запущен.');     // Activate.cmd (Активация на 180 дней)
end;

end;

procedure TFormKMS.Button3Click(Sender: TObject); // активация с автопродлением
var
i:integer;
begin
if not FileExists(labelEdEdit4.text) then
begin
i:=MessageDlg('Не найдены скрипты и файлы для работы.'+#10#13+' Загрузить необходимые файлы?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
if not loadFileRun then exit;
end;
Memo1.Clear;
if CopyFFSelectPC
(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
labelEdEdit1.text+#0+#0,        // что копируем
'C$\TEMP\'+#0+#0,              // куда копируем
FormKMS,                       // форма владелец
2,                            // операция копировать
false,                       // возможность отмены
true) then                      // проверять наличие каталога назначения
begin
if MyNewProcess(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
'C:\TEMP\'+GetLastDir(labelEdEdit4.text)+ExtractFileName(labelEdEdit4.text)+' '+Edit4.Text) then //
memo1.Lines.Add(frmdomaininfo.Combobox2.text+' - Запуск процесса активации с автопродлением запущен.');  //KMS_VL_ALL.cmd (Активация с автоматическим продлением)
end;
end;

procedure TFormKMS.Button4Click(Sender: TObject);  // удалить KMS активацию
var
i:integer;
begin
if not FileExists(labelEdEdit3.text) then
begin
i:=MessageDlg('Не найдены скрипты и файлы для работы.'+#10#13+' Загрузить необходимые файлы?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
if not loadFileRun then exit;
end;
Memo1.Clear;
if CopyFFSelectPC
(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
labelEdEdit1.text+#0+#0,        // что копируем
'C$\TEMP\'+#0+#0,              // куда копируем
FormKMS,                       // форма владелец
2,                            // операция копировать
false,                       // возможность отмены
true) then                      // проверять наличие каталога назначения
begin
if MyNewProcess(frmdomaininfo.Combobox2.text, // компьютер на который копируем
frmdomaininfo.labelEdEdit1.text, // пользователь
frmdomaininfo.labelEdEdit2.text, // пароль
'C:\TEMP\'+GetLastDir(labelEdEdit3.text)+ExtractFileName(labelEdEdit3.text)+' '+Edit3.Text) then //   ключ /u удаление скриптов и ключей после очистки (удаляет родительский каталог скрипта)
memo1.Lines.Add(frmdomaininfo.Combobox2.text+' - Запуск процесса удаления.');    // AutoRenewal-Uninstall.cmd (Удаление KMS )
end;
end;





procedure TFormKMS.Button1Click(Sender: TObject);
begin
if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
begin
MemIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
MemIni.WriteString('Vaccina','Foder',LabeledEdit1.Text);
MemIni.WriteString('Vaccina','Manual',LabeledEdit2.Text);
MemIni.WriteString('Vaccina','Uninstall',LabeledEdit3.Text);
MemIni.WriteString('Vaccina','Auto',LabeledEdit4.Text);
MemIni.WriteString('Vaccina','ManualKey',Edit2.Text);
MemIni.WriteString('Vaccina','UninstallKey',Edit3.Text);
MemIni.WriteString('Vaccina','AutoKey',Edit4.Text);
MemIni.UpdateFile;
MemIni.Free;
end;
end;

procedure TFormKMS.memo1URLClick(Sender: TObject; const URLText: string;
  Button: TMouseButton);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar(URLText), nil, nil, SW_SHOWNORMAL);
  except
  showmessage('Ошибка при открытии приложения')
  end;
end;

procedure TFormKMS.memotext;
begin

Memo1.Clear;
with Memo1.Lines do
begin
Memo1.SelAttributes.Color:=clRed;
memo1.SelAttributes.Style:=[fsBold, fsUnderline];
Add('!*!*!*ВНИМАТЕЛЬНО ПРОЧИТАЙТЕ ДАННЫЙ ТЕКСТ*!*!*!');
Add('!*!Всегда используйте лицензионное программное обеспечение!*!');
end;

Memo1.Lines.Add('Все реализованные Функции представлены в ознакомительных'+
' целях. При выполнении функции представленных в данном окне'+
' приложения, пользователь соглашается с тем что выполняет'+
' их в ознакомительных целях и обязуется удалить их'+
' последствия после изучения действий этих функций.'+
' Разработчик не рекомендует использовать стороние приложения'+
' или иные функции нарушающие лицензионные соглашения'+
' третьих сторон.');
Memo1.Lines.Add('Все загружаемые файлы представлены в ознакомительных целях и'+
' являются открытым проектом команды сторонних разработчиков,'+
' ознакомится с данным проектом можно по адресу:');
Memo1.Lines.Add('!*!*!*!*!*!*!*github.com bought by Microsoft*!*!*!*!*!*!*!');
Memo1.Lines.Add('https://github.com/kkkgo/KMS_VL_ALL');
Memo1.Lines.Add(#10+'P.S. Все представленые функции возможно реализовать через запуск процесса.'+
' Каждая функция осуществляет копирование каталога с необходимыми файлами и запуск сmd/bat файла. Краткая инструкция:'+
' "Новый процесс"->"Добавить новый процесс в избранное"->'+
'"Копировать перед запуском"->"Копировать родительский каталог файла",'+
' не забываем нажимать кнопку "Cохранить в избранном". Создать новую задачу, добавить задание "Процессы -> Запустить процесс".');
end;

procedure TFormKMS.RereadFile; //перечитывем файл настроек
begin
if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
begin
MemIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
LabeledEdit1.Text:=MemIni.ReadString('Vaccina','Foder',extractfilepath(application.ExeName)+'Vaccina');
LabeledEdit2.Text:=MemIni.ReadString('Vaccina','Manual',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
Edit2.Text:=MemIni.ReadString('Vaccina','ManualKey','/m /q');
LabeledEdit3.Text:=MemIni.ReadString('Vaccina','Uninstall',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
Edit3.Text:=MemIni.ReadString('Vaccina','UninstallKey','/r /q');
LabeledEdit4.Text:=MemIni.ReadString('Vaccina','Auto',extractfilepath(application.ExeName)+'Vaccina\Vaccination.cmd');
Edit4.Text:=MemIni.ReadString('Vaccina','AutoKey','/a /q');
MemIni.Free;
end;
end;

procedure TFormKMS.FormShow(Sender: TObject);
var
i:integer;
begin
RereadFile;
Memo1.Clear;
StatusBar1.Panels[0].Text:='';
PageControl1.ActivePageIndex:=0;
memotext;
end;

procedure TFormKMS.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;
procedure TFormKMS.CreateFormElseErrorloadFile;
var
FormKey:Tform;
MemoKey:TJvRichEdit;
begin
FormKey:=TForm.Create(FormKMS);
FormKey.Caption:='Справочная информация';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=430;
FormKey.Height:=230;
FormKey.OnClose:=FormKeyClose;
////////////
MemoKey:=TJvRichEdit.Create(FormKey);
MemoKey.Parent:=FormKey;
MemoKey.Align:=alClient;
MemoKey.ScrollBars:=ssVertical;
MemoKey.ReadOnly:=true;
MemoKey.AutoURLDetect:=true;
MemoKey.OnURLClick:=memo1URLClick;
MemoKey.TabOrder:=0; // переводим фокус в поле со значением параметра
MemoKey.Lines.Add('Если вы читаете это сообщение значит не получилось загрузить'+
' необходимые файлы или вы выбрали ручную загрузку. Для загрузки файлов пройдите по адресу'+
' https://skrblog.ru/download/kms/?wpdmdl=1478&masterkey=60322427c5fb6'+
' скачайте архив файлов, пароль от архива skrblog.ru. Распакуйте файлы в директорию установки' +
' программы, в каталог Vaccina, если такового нет то создайте его, укажите директори каталога и файлов а также аргументы в настройках.'+
' Данные файлы являются открытым проектом команды разработчиков,'+
' все файлы и справочную информацию можно найти по адресам:');
MemoKey.Lines.Add('https://github.com/kkkgo/KMS_VL_ALL');
MemoKey.Lines.Add('https://forums.mydigitallife.net/posts/838808/');
MemoKey.Lines.Add('https://pastebin.com/cpdmr6HZ');
MemoKey.Lines.Add('https://textuploader.com/1dav8');
FormKey.ShowModal;
end;

procedure TFormKMS.HttpWork(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
var
  Http: TIdHTTP;
  ContentLength: Int64;
begin
  Http := TIdHTTP(ASender);
  try
  ContentLength := Http.Response.ContentLength;
  if (Pos('chunked', LowerCase(Http.Response.TransferEncoding)) = 0) and
     (ContentLength > 0) then
  begin
    ProgressB.Position := 100*AWorkCount div ContentLength;
  end;
  finally

  end;
end;

function createDir(path:string):boolean;// проверка и создание директории
begin
   try
    if not DirectoryExists(path) then //если нет каталога
    begin
     if ForceDirectories(path) then/// создает весь путь до папки
     result:=true  // директория создана
     else result:=false;
    end
    else result:=true; // директория есть
  except on E: Exception do
     begin
     result:=false;
     end;
   end;
end;

procedure TFormKMS.LinkLabel1Click(Sender: TObject);
begin
CreateFormElseErrorloadFile;
end;

function TFormKMS.LoadFile(url,path,Fname:string):boolean;  // функция загрузки файла
var
MS: TMemoryStream;
  MyClass: TComponent;
begin
try
MS := TMemoryStream.Create;
try
IdHTTP1.Get(url, MS);
MS.SaveToFile(path+Fname);
MS.Clear;
result:=true;
finally
 MS.Free;
end;
except on E: Exception do begin result:=false; StatusBar1.Panels[0].Text:='Ошибка соединения: '+e.Message end;
end;
end;




procedure TFormKMS.Button5Click(Sender: TObject);
begin       //https://skrblog.ru/download/kms/?wpdmdl=1478&masterkey=60322427c5fb6  - архив
if loadFileRun then
 ShowMessage('Неоходимые файлы и настройки загружены');
end;



function TformKMS.ExistFileKMS(ListFile:TstringList):boolean;
var
i:integer;
begin
 for I := 0 to ListFile.Count-1 do
 if not FileExists(ListFile[i])  then
 begin
   result:=false;
   break;
 end
 else result:=true;
end;

function TFormKMS.loadFileRun:boolean; //функция загрузки файла настроек и необходимых файлов для KMS
var
LoadIni,SetIni:Tmeminifile;
i:integer;
path:string;
ListValue:TstringList;
NameFile:Tstringlist;
begin
try
if not frmDomainInfo.ping('yandex.ru') then
begin
i:=MessageDlg('Нет активного подключения к интернету!!!'+#10#13+' Продолжить выполнение операции?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then
  begin
  result:=false;
  exit;
  end;
end;
ListValue:=TStringList.Create;
NameFile:=TStringList.Create;
ProgressB:=TProgressBar.Create(FormKMS);
ProgressB.Parent:=StatusBar1;
ProgressB.Left:=0;
ProgressB.Width:=StatusBar1.Panels[0].Width;
ProgressB.Max:=100;
ProgressB.Step:=1;
ProgressB.Position:=0;
StatusBar1.Panels[0].Text:='';
try
IdHTTP1.OnWork:= HttpWork;
if Fileexists(extractfilepath(application.ExeName)+'\Vaccina.ini') then // если файл есть то удаляем
DeleteFile(extractfilepath(application.ExeName)+'\Vaccina.ini');

if LoadFile('http://skrblog.ru/wp-content/uploads/download-manager-files/Vaccina/Vaccination.ini',extractfilepath(application.ExeName),'Vaccination.ini') then
Begin
 if Fileexists(extractfilepath(application.ExeName)+'\Vaccination.ini') then
  begin
  LoadIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Vaccination.ini',TEncoding.Unicode);
  try
  LoadIni.ReadSectionValues('FOLDER',ListValue); // чтение необходимых каталогов
  for I := 0 to ListValue.Count-1 do createDir(extractfilepath(application.ExeName)+'\'+ListValue.ValueFromIndex[i]);  // содание необходимых каталогов
  ListValue.Clear;
  LoadIni.ReadSectionValues('URL',ListValue); // чтение необходимых URL для загрузки
  LoadIni.ReadSectionValues('FILES',NameFile); // чтение имен и расположений файлов
  for I := 0 to ListValue.Count-1 do
    begin
   // frmDomainInfo.Memo1.Lines.Add('URL - '+ListValue.ValueFromIndex[i]+'  Каталог - '+extractfilepath(application.ExeName)+' Файл - '+NameFile.ValueFromIndex[i]);
    if not LoadFile(ListValue.ValueFromIndex[i],extractfilepath(application.ExeName),NameFile.ValueFromIndex[i]) then // загруужаем и сохраняем файлы  .
     begin
     frmDomainInfo.Memo1.Lines.Add('Не удалось загрузить файл - '+NameFile.ValueFromIndex[i]);
     CreateFormElseErrorloadFile;
     break;
     end;
    end;
  ///// вносим изменения в файл настроек,
  SetIni:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  try
  SetIni.WriteString('Vaccina','Foder',extractfilepath(application.ExeName)+'Vaccina');
  SetIni.WriteString('Vaccina','Manual',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Manual',''));
  SetIni.WriteString('Vaccina','Uninstall',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Uninstall',''));
  SetIni.WriteString('Vaccina','Auto',extractfilepath(application.ExeName)+LoadIni.ReadString('RUN FILES','Auto',''));
  SetIni.WriteString('Vaccina','ManualKey',LoadIni.ReadString('KEY','ManualKey',''));
  SetIni.WriteString('Vaccina','UninstallKey',LoadIni.ReadString('KEY','UninstallKey',''));
  SetIni.WriteString('Vaccina','AutoKey',LoadIni.ReadString('KEY','AutoKey',''));
  SetIni.UpdateFile;
  finally
  SetIni.Free;
  end;
  RereadFile; // перечитать файл настроек
  /////////// проверяем загрузились или нет файлы
    ListValue.Clear;
    for I := 0 to NameFile.Count-1 do
    begin
    ListValue.Add(extractfilepath(application.ExeName)+NameFile.ValueFromIndex[i]);
    frmDomainInfo.Memo1.Lines.Add(extractfilepath(application.ExeName)+NameFile.ValueFromIndex[i]);
    end;

    if not ExistFileKMS(ListValue) then // если не загрузили все необходимые файлы
    Begin
    ShowMessage('Не удалось загрузить необходимые файлы');
    CreateFormElseErrorloadFile;
    End
    else result:=true;
  finally
  LoadIni.Free;
  end;
  end;

End
else
  begin
  ShowMessage('Не удалось загрузить файл настроек');
  CreateFormElseErrorloadFile;
  result:=false;
  end;


finally
ProgressB.Free;
ListValue.Free;
NameFile.Free;
if Fileexists(extractfilepath(application.ExeName)+'\Vaccination.ini') then // если файл есть то удаляем
DeleteFile(extractfilepath(application.ExeName)+'\Vaccination.ini');
end;
except on E: Exception do
begin
  result:=false;
  ShowMessage('Ошибка загрузки файлов '+e.Message);
end;
end;
end;

end.
