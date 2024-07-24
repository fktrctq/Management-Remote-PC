unit FormNewProcess;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.Dialogs,inifiles,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.StrUtils,System.Variants,ActiveX,ComObj,IdIcmpClient,ShellAPI,
  Vcl.ExtCtrls;
 type
 TStrForNewProcess = record
    FSource :String;   ///что копировать источник файл или каталог
    FDest:string; // куда копировать
    NamePC:string;     // список компьютеров
    UserName:string;   // логин
    PassWd:string;     // пароль
    PathCreate:boolean;  // проверка и создания директории
    CancelCopyFF:boolean;   // отменить или нет операцию копирования
    BeforeInstallCopy:boolean;     // копировать или нет перед установкой
    DeleteAfterInstall:boolean; // удалить дистрибутив после установки
    PathDelete:string;         // Какой каталог или файл удалять после установки
    OwnerForm:Tform;        // родительская форма прогресс бара, можно не указывать
    NumCount:integer;
    TypeOperation:integer;     // тип оперции, копировать, вставить и т.д. только для структуры TSHFileOpStruct
    FileToRun:String;   // файл для запуска
  end;
type
  TNewProcForm = class(TForm)
    OKBtn: TButton;
    CheckBox1: TCheckBox;
    Label2: TLabel;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    SpeedButton3: TSpeedButton;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    CheckBox4: TCheckBox;
    EditCopyPath: TLabeledEdit;
    ComboBox1: TComboBox;
    CheckBox5: TCheckBox;
    Button1: TButton;
    CheckBox2: TCheckBox;
    EditSource: TLabeledEdit;
    Button2: TButton;
    EditFileRun: TLabeledEdit;
    SpeedButton4: TSpeedButton;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Label1: TLabel;
    SpeedButton5: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    function readproc(item:integer):boolean;
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;



var
  NewProcForm: TNewProcForm;
  ArrayProcess: array of TThread;
  ExitFor:boolean;
  StrForProcess: array [0..2000] of TStrForNewProcess; //массив переменных для запуска потоков установки программ
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;

 ThreadVar
PointForRunNewProcess: ^TStrForNewProcess;

implementation
uses uMain,unit5, SelectedPCNewProcessThread,MyDM,EditProcMSI;



{$R *.dfm}

function TNewProcForm.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''proc'''
+' ORDER BY DESCRIPTION_PROC'; //ID_PROC
FDQueryProc.Open;
ComboBox2.Clear;
ComboBox3.Clear;
while not FDQueryProc.Eof do
begin
 if FDQueryProc.FieldByName('DESCRIPTION_PROC').Value<>null then
 begin
   ComboBox2.Items.Add(FDQueryProc.FieldByName('DESCRIPTION_PROC').AsString);
   ComboBox3.Items.Add(FDQueryProc.FieldByName('ID_PROC').AsString);
 end;
  FDQueryProc.Next;
end;
FDQueryProc.Close;
if (ComboBox2.Items.Count>0) and (item<>-1) then
  begin
  ComboBox2.ItemIndex:=item;
  ComboBox2.OnSelect(ComboBox2);
  end;
end;

function MyNewProcess(param:pointer):boolean;
var
MyError:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObject   : OLEVariant;
objProcess    : OLEVariant;
objConfig     : OLEVariant;
ProcessID,z,i   : Integer;
listPC:TstringList;
CopyOrNo:bool;
const
wbemFlagForwardOnly = $00000020;
HIDDEN_WINDOW       = 1;
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
////////////////////////////////////////////////////
 function FindAddcreateDir(path,NamePC:string):boolean;// проверка и создание директории
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
////////////////////////////////////////////////////////////////
function CopyFFSelectPC(CurentPC,UserName,PassWd,FSource,FDest:string;OwnerForm:TForm;
TypeOperation:integer;CancelCopyFF,PathCreate:boolean):boolean; // копирование в потоке для групп  компьютеров
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
      finditem('Запущена операция копирования',CurentPC,17);
      frmDomaininfo.Memo1.Lines.Add(CurentPC+' - Запущена операция копирования');
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // что удалять  если операция удаление
      pTo := pchar('');   // куда копируем не используется
      finditem('Запущена операция удаления ',CurentPC,13);
      frmDomaininfo.Memo1.Lines.Add(CurentPC+' - Запущена операция удаления');
     end;
    end;

    try ////////////////////////////////////// авторизация пользователя на удаленном компе
     if not (LogonUserA (PAnsiChar(UserName), PAnsiChar (CurentPC),  // сначала заходим на комп в сети
     PAnsiChar (PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do frmdomaininfo.Memo1.Lines.Add(CurentPC+' : Ошибка LogonUser - '+e.Message)
     end;
     ////////////////////////////////////////////
     try
     if PathCreate then FindAddcreateDir(FDest,CurentPC); //проверяем и создаем каталог если его нет.
     rescopy:=SHFileOperation(SHFileOpStruct); //непосредственно функция копирования/перемещения /удаления
     if rescopy=0 then
       begin
        if TypeOperation=2 then
        begin
        finditem('Операция копирования успешно завершена',CurentPC,1);
        frmDomaininfo.Memo1.Lines.Add(CurentPC+' - Операция копирования успешно завершена');
        end;
        if TypeOperation=3 then
        begin
         finditem('Операция удаления успешно завершена',CurentPC,1);
         frmDomaininfo.Memo1.Lines.Add(CurentPC+' - Операция удаления успешно завершена');
        end;
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then
        begin
         finditem('Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
         frmDomaininfo.Memo1.Lines.Add(CurentPC+' - Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        if TypeOperation=3 then
        begin
        finditem('Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        frmDomaininfo.Memo1.Lines.Add(CurentPC+'Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        result:=false;
       end;
     CloseHandle(htoken);  //закрываем хендл полученый при LogonUser
      except
      on E: Exception do
      begin
       if TypeOperation=2 then
       begin
       frmdomaininfo.Memo1.Lines.Add(Format('Ошибка оперции копирования на компьютер  %s :  %s',[CurentPC,E.Message]));
       finditem('Ошибка оперции копирования: '+E.Message,CurentPC,2);
       end;
       if TypeOperation=3 then
       begin
       frmdomaininfo.Memo1.Lines.Add(Format('Ошибка оперции удаления на компьютере  %s :  %s',[CurentPC,E.Message]));
       finditem('Ошибка оперции удаления: '+E.Message,CurentPC,2);
       end;
       result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then frmdomaininfo.Memo1.Lines.Add(Format('Общая ошибка функции копирования :  %s',[E.Message]));
     if TypeOperation=3 then frmdomaininfo.Memo1.Lines.Add(Format('Общая ошибка функции удаления :  %s',[E.Message]));
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////

Begin
PointForRunNewProcess:=param;
try
listPC:=TStringList.Create;
listPC.CommaText:=PointForRunNewProcess.NamePC;
for I := 0 to listPC.Count-1 do
  BEGIN
  if ping(listPC[i]) then    ///// если компьютер доступен то запусаем процесс
  try
  begin
  CopyOrNo:=false;
  if PointForRunNewProcess.BeforeInstallCopy then // если необходимо скопировать перед установкой
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // имя компа
  PointForRunNewProcess.UserName,//логин
  PointForRunNewProcess.PassWd,  // пароль
  PointForRunNewProcess.FSource, // что копировать
  PointForRunNewProcess.FDest,   // куда копировать
  PointForRunNewProcess.OwnerForm,// родительская форма
  PointForRunNewProcess.TypeOperation, //тип операции
  PointForRunNewProcess.CancelCopyFF,
  PointForRunNewProcess.PathCreate) ;
  end;
  OleInitialize(nil);
  frmDomainInfo.memo1.Lines.Add('Запуск процесса на - '+listPC[i]);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(listPC[i], 'root\CIMV2', PointForRunNewProcess.UserName, PointForRunNewProcess.PassWd,'','',128);
  FWMIService.security_.AuthenticationLevel:=6;
  FWMIService.security_.ImpersonationLevel:=3;
  //FWMIService.security_.Privileges.AddAsString('SeEnableDelegationPrivilege');
  FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
  objConfig     := FWbemObject.SpawnInstance_;
  objConfig.ShowWindow := HIDDEN_WINDOW;
  objProcess    := FWMIService.Get('Win32_Process');
  MyError:=objProcess.Create(PointForRunNewProcess.FileToRun, null, objConfig, (ProcessID));
  if MyError=0 then
    begin
    finditem('Запуск процесса '+PointForRunNewProcess.FileToRun+' : '+': '+SysErrorMessage(MyError),listPC[i],1);
    end
  else
    begin
    finditem('При запуске процесса '+PointForRunNewProcess.FileToRun+' возникли проблемы : '+SysErrorMessage(MyError),listPC[i],2);
    end;
  frmDomainInfo.memo1.Lines.Add('Запуск процесса '+PointForRunNewProcess.FileToRun+' на '+listPC[i]+' : '+SysErrorMessage(MyError));
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  FWbemObject:=Unassigned;
  objProcess:=Unassigned;
  VariantClear(FWbemObject);
  VariantClear(objConfig);
  VariantClear(objProcess);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  if PointForRunNewProcess.BeforeInstallCopy and  //если указано скопировать перед установкой
  PointForRunNewProcess.DeleteAfterInstall and    // если необходимо удалить дистрибутив после установки
  CopyOrNo then // если операция копирования перед установкой прошла успешно
    begin        // удаляем
    CopyFFSelectPC(ListPC[i], // имя компа
    PointForRunNewProcess.UserName,//логин
    PointForRunNewProcess.PassWd,  // пароль
    '', // что копировать можно не укузывать т.к. удаляем
    PointForRunNewProcess.PathDelete,   // что удалять
    PointForRunNewProcess.OwnerForm,// родительская форма
    3, //тип операции  (3)- удалить, FO_MOVE
    PointForRunNewProcess.CancelCopyFF,
    false) ;  // не проверять наличие каталога т.к. удаляем после копирования
    end;
  end;
    except
      on E:Exception do
      begin
      finditem('При запуске процесса '+PointForRunNewProcess.FileToRun+' возникли проблемы : "'+E.Message+'"',listPC[i],2);
      frmDomainInfo.memo1.Lines.Add('При запуске процесса '+PointForRunNewProcess.FileToRun+' на '+listPC[i]+' возникли проблемы. - "'+E.Message+'"');
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      OleUnInitialize;
      end;
    end;
  END;
finally
listPC.Free;
End;
end;

function  DeliteKeyFilePatch(strfile:string):string; // функция обрабатывает передаваемый в нее путь с ключами запуска, выдает путь до файла
begin
try
if pos(' -',strfile)<>0 then    // ищем ключи запуска через дефис
begin
  strfile:=copy(strfile,1,pos(' -',strfile)-1);
end;
if pos('/',strfile)<>0 then //ищем обратный слеш, признак ключа запуска
begin
 strfile:=copy(strfile,1,pos('/',strfile)-1);
 if strfile[length(strfile)]=' ' then strfile:=copy(strfile,1,length(strfile)-1);
end;
result:=strfile;
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('Ошибка формирования строки - "'+E.Message+'"');
result:=strfile;
end;
  end;
end;


procedure TNewProcForm.Button1Click(Sender: TObject);
Function renewPath(s:string):string;
begin
if AnsiPos(':',s)=2 then
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
  Name:='DlgAddFolder';
  Title:='Выбор каталога для копирования';
  Options:=[fdoForceShowHidden,fdoPickFolders]; {каталоги}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditCopyPath.Text:=renewPath(FileName)+'\';
    end;
 finally
 Destroy;
 end;
 end;
 end;





procedure TNewProcForm.Button2Click(Sender: TObject);
begin
Exitfor:=false;
Close;
end;

procedure TNewProcForm.CheckBox2Click(Sender: TObject);
begin
GroupBox2.Enabled:=CheckBox2.Checked;
if CheckBox2.Checked then
begin
EditSource.Text:=DeliteKeyFilePatch(EditFileRun.Text);
GroupBox2.Height:=160;
end
else GroupBox2.Height:=25;
end;

procedure TNewProcForm.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
combobox2.DroppedDown:=true;
end;

procedure TNewProcForm.ComboBox2Select(Sender: TObject);
var
FDQuery:TFDQuery;
begin
ComboBox3.ItemIndex:=ComboBox2.ItemIndex;
try
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionReadProc;
FDQuery.Connection:=DataM.ConnectionDB;
FDQuery.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE DESCRIPTION_PROC='''+ComboBox2.Text+'''';//' ORDER BY ID_PROC';
FDQuery.Open;
EditfileRun.Text:=vartostr(FDQuery.FieldByName('PATCH_PROC').Value);
if FDQuery.FieldByName('BEFOREINSTALLCOPY').Value<>null then CheckBox2.Checked:=FDQuery.FieldByName('BEFOREINSTALLCOPY').AsBoolean;
if FDQuery.FieldByName('FILESOURSE_PROC').Value<>null then  EditSource.Text:=FDQuery.FieldByName('FILESOURSE_PROC').AsString;
if FDQuery.FieldByName('FILEORFOLDER').Value<>null then
begin
if FDQuery.FieldByName('FILEORFOLDER').AsString='File' then combobox1.ItemIndex:=0
else combobox1.ItemIndex:=1;
end;
if FDQuery.FieldByName('PATHCREATE').Value<>null then EditCopyPath.Text:=FDQuery.FieldByName('PATHCREATE').AsString;
if FDQuery.FieldByName('DELETEAFTERINSTALL').Value<>null then  CheckBox5.Checked:=FDQuery.FieldByName('DELETEAFTERINSTALL').AsBoolean;

finally
FDQuery.Close;
FDQuery.Free;
end;
end;


{procedure TPasswordDlg.ComboBox1DropDown(Sender: TObject);
var i,L,mwidth:integer;  /// устанавливаем ширину выпадающего списка
begin
mwidth:=0;
   with ComboBox1 do
   begin
     for i := 0 to Items.Count - 1 do
     if ( Canvas.TextWidth(Items[I]) > mWidth) then
       mWidth :=Canvas.TextWidth(Items[I])+ 100;
   SendMessage(ComboBox1.Handle ,$0160,mWidth,0);
   end
end; }



procedure TNewProcForm.FormCreate(Sender: TObject);
begin
speedButton1.Hint:='Добавить новый процесс в избанное';
speedButton2.Hint:='Удалить процесс из списка';
end;

procedure TNewProcForm.FormShow(Sender: TObject);
begin
if not groupPC then NewProcForm.Caption:='Новый процесс на '+frmDomainInfo.ComboBox2.Text
else NewProcForm.Caption:='Новый процесс на группе ПК';
CheckBox1.Checked:=true; /// в одном потоке независимо от группового запуска или на текущем компьютере
if GroupPC=true then CheckBox1.Visible:=true
else CheckBox1.Visible:=false;
label2.Visible:=false;
GroupBox2.Height:=25;
CheckBox2.Checked:=false;
if not DataM.ConnectionDB.Connected then exit;
readproc(-1);
end;



procedure TNewProcForm.OKBtnClick(Sender: TObject);
var
b,i:integer;
NumPer:integer;
SelectedPCNewProc:Tstringlist;
function tagForCopy(num:integer):integer;
begin
if OKBtn.tag=1999 then num:=0;
inc(num);
OKBtn.tag:=num;
result:=num;
end;

///////////////////////////////////////////////
function GetLastDir(path:String):String;  // извлечение корневой папки
var
i,j:integer;
begin
try
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
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('Ошибка извлечения корневой папки - "'+E.Message+'"');
result:=path;
end;
  end;
end;
/////////////////////////////////////////////////////
Function renewPath(s:string):string;
begin
try
if AnsiPos('$',s)=2 then
 begin
 delete(s,2,1); // удаляем символ $
 insert(':',s,2); // вставляем на его место :
 end;
result:=s;
except
on E:Exception do
begin
frmDomainInfo.memo1.Lines.Add('Ошибка renewPath - "'+E.Message+'"');
result:=s;
end;
  end;
end;
///////////////////////////////////////////////////////
begin
try
{if not FileExists(EditFileRun.Text) then
begin
    b:= MessageDlg('Не могу найти файл, возможно добавлены ключи запуска.'+ #10#13+'Продолжить запуск процесса?',mtConfirmation,[mbYes,mbCancel], 0);
    if b = mrCancel then exit;
end;}
  if (CheckBox1.Checked) then //// запуск процесса в одном потоке на группе компьютеров или на одном компьютер
    begin
    NumPer:=tagForCopy(OKBtn.tag);
    StrForProcess[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
    StrForProcess[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
    if GroupPC=true then // если групповая установка то добавляем список компьютеров
      begin
      SelectedPCNewProc:=TstringList.Create;
      SelectedPCNewProc:=frmDomainInfo.createListpcForCheck('');
      StrForProcess[NumPer].NamePC:=SelectedPCNewProc.CommaText;
      SelectedPCNewProc.Free;
      end
    else StrForProcess[NumPer].NamePC:=frmDomainInfo.ComboBox2.Text; // иначе установка на текущий компьютер
    if CheckBox2.Checked then // если надо скопировать файлы
    begin
      if ComboBox1.ItemIndex=0 then // если копируем только файл
        begin
        StrForProcess[NumPer].FSource:= EditSource.Text+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+ExtractFileName(EditFileRun.Text); // строка запуска файла для установки
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(EditSource.Text)+#0+#0; // путь к удаляемому файлу после установки программы
        end
      else // если копируем весь корневой каталог
        begin
        StrForProcess[NumPer].FSource:= ExtractFilePath(EditSource.Text)+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+GetLastDir(EditSource.Text)+ExtractFileName(EditFileRun.Text);
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(EditFileRun.Text)+#0+#0;  // путь к удаляемому каталогу после установки программы
        end;
     end;
    if not CheckBox2.Checked then  StrForProcess[NumPer].FileToRun:= EditFileRun.Text; //// если не копируем файл или каталог перед установкой то файл для установки берем из источника
    StrForProcess[NumPer].FDest:=EditCopyPath.Text+#0+#0;
    StrForProcess[NumPer].PathCreate:=CheckBox4.Checked;
    StrForProcess[NumPer].CancelCopyFF:=false;
    StrForProcess[NumPer].BeforeInstallCopy:=CheckBox2.Checked;
    StrForProcess[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
    StrForProcess[NumPer].OwnerForm:=Self;
    StrForProcess[NumPer].NumCount:=NumPer;
    StrForProcess[NumPer].TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
    res[NumPer]:=BeginThread(nil,0,addr(MyNewProcess),Addr(StrForProcess[NumPer]),0,treadID[NumPer]); ///
    CloseHandle(res[NumPer]);
    end;
  if (GroupPC=true)and(CheckBox1.Checked=false) then //// запуск процесса на компах в разных потоках
    begin
    exitfor:=true;
    SelectedPCNewProc:=TstringList.Create;
    SelectedPCNewProc:=frmDomainInfo.createListpcForCheck('');
    label2.Visible:=true;
    for I := 0 to SelectedPCNewProc.Count-1 do
      begin
      NumPer:=tagForCopy(OKBtn.tag);
      StrForProcess[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForProcess[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForProcess[NumPer].NamePC:=SelectedPCNewProc[i];
      if CheckBox2.Checked then // если надо скопировать файлы
      begin
        if ComboBox1.ItemIndex=0 then // если копируем только файл
        begin
        StrForProcess[NumPer].FSource:= EditSource.Text+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+ExtractFileName(EditFileRun.Text); // строка запуска файла для установки
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(EditSource.Text)+#0+#0; // путь к удаляемому файлу после установки программы
        end
         else // если копируем весь корневой каталог
        begin
        StrForProcess[NumPer].FSource:= ExtractFilePath(EditSource.Text)+#0+#0;
        StrForProcess[NumPer].FileToRun:=renewPath(EditCopyPath.Text)+GetLastDir(EditSource.Text)+ExtractFileName(EditFileRun.Text);
        StrForProcess[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(EditFileRun.Text)+#0+#0;  // путь к удаляемому каталогу после установки программы
        end;
      end;
      if not CheckBox2.Checked then // если не копируем файл или каталог перед установкой
      StrForProcess[NumPer].FileToRun:= EditFileRun.Text; //то файл для установки берем из источника
      StrForProcess[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForProcess[NumPer].PathCreate:=CheckBox4.Checked;
      StrForProcess[NumPer].CancelCopyFF:=false;
      StrForProcess[NumPer].BeforeInstallCopy:=CheckBox2.Checked;
      StrForProcess[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForProcess[NumPer].OwnerForm:=Self;
      StrForProcess[NumPer].NumCount:=NumPer;
      StrForProcess[NumPer].TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
      res[NumPer]:=BeginThread(nil,0,addr(MyNewProcess),Addr(StrForProcess[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      Application.ProcessMessages;
      sleep(500);
      label2.Caption:=inttostr(((100 div SelectedPCNewProc.Count-1)*(i+1)))+'%';
      if not ExitFor then break;
      end;
     label2.Visible:=false;
     SelectedPCNewProc.Free;
    end;

    {{if GroupPC=false then  //// запуск процесса на текущем компьютере
    begin
    NewProcMyPS:=MyPS;   имя компа задается при запуске формы из разных частей приложения перед открытием формы
    MyNewProc:=unit5.NewProcess.Create(true);
    MyNewProc.FreeOnTerminate:=true;
    MyNewProc.Start;
    end;}
close;
except
on E:Exception do frmDomainInfo.memo1.Lines.Add('Ошибка формирования команды запуска процесса - "'+E.Message+'"');
end;
end;


procedure TNewProcForm.SpeedButton1Click(Sender: TObject);
var
newOpenDlg:TopenDialog;
Newpatch,Newdescription:string;
z:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
Newpatch:='';
Newdescription:='';
newOpenDlg:=TOpendialog.Create(self);
newOpenDlg.Title:='Укажиет файл для запуска';
if newOpenDlg.Execute then
  begin
   NewPatch:=InputBox('Добавление ключа запуска', 'Файл:', newOpenDlg.FileName);
   if NewPatch=newOpenDlg.FileName then
     begin
     z:=MessageDlg('Вы не указали ключ запуска'+#10#13+'Продолжить???',
     mtConfirmation, mbOKCancel, 0);
     if z=mrCancel then begin newOpenDlg.Free; exit;  end;
     end;
   Newdescription:=InputBox('Добавление описания к файлу', 'описание:', 'здесь ваше описание');
   // если не указали описание то добавляем в описание имя выбранного файла
   if (Newdescription='') and (fileexists(NewOpenDlg.FileName)) then Newdescription:=ExtractFileName(NewOpenDlg.FileName);
  end;

if (NewPatch='') or (Newdescription='') then exit;

// записываем новый процесс в базу
EditSource.Text:=DeliteKeyFilePatch(NewPatch); // заполняем поле файл для копирования
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(NewPatch,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'proc'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+leftstr(Newdescription,1000)+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox2.Checked; // копировать файлы перед установкой
if ComboBox1.ItemIndex=0 then // если копируем только файл
FDQueryProc.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // если копируем родительский каталог
FDQueryProc.params.ParamByName('p5').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p6').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //директория куда копировать
FDQueryProc.params.ParamByName('p7').AsBoolean:=CheckBox5.Checked; // удалить файлы после установки
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditSource.Text,1000)+'';    // путь до файла источника
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// открываем список  процессов из базы
readproc(ComboBox2.Items.Count);
newOpenDlg.Free;
end;

procedure TNewProcForm.SpeedButton2Click(Sender: TObject); //удалить
var
i:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
if ComboBox2.Text='' then exit;

i:=MessageDlg('Перед удалением убедитесь что запись не используется в задачах.'+#10#13+'Продолжить выполнение операции?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ComboBox3.text+'''';
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// открываем список  процессов из базы
readproc(ComboBox2.ItemIndex-1);
end;



procedure TNewProcForm.SpeedButton3Click(Sender: TObject);
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор файла для запуска';
  //Options:=[fdoForceShowHidden,fdoPickFolders]; {каталоги}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditFileRun.Text:=FileName;
    EditSource.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;



procedure TNewProcForm.SpeedButton4Click(Sender: TObject);  //сохранить
var
descript:string;
begin

if (EditFileRun.Text='') then
begin
  ShowMessage('Укажите путь к файлу и описание процесса');
  SpeedButton1.Click;
  exit;
end;

descript:=ComboBox2.Text;

if (ComboBox2.Text='')and (EditFileRun.Text<>'') then
 begin
 descript:=InputBox('Добавьте описания к файлу', 'Описание:', 'здесь описание процесса');
 end;
 if descript='' then exit;

FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8) MATCHING (DESCRIPTION_PROC)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(EditFileRun.Text,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'proc'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+descript+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox2.Checked; // копировать файл или каталог перед установкой
if ComboBox1.ItemIndex=0 then // если копируем только файл
FDQueryProc.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // если копируем родительский каталог
FDQueryProc.params.ParamByName('p5').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p6').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //директория куда копировать
FDQueryProc.params.ParamByName('p7').AsBoolean:=CheckBox5.Checked; // удалить файлы после установки
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditSource.Text,1000)+'';    // путь до файла источника
FDQueryProc.ExecSQL;
FDQueryProc.close;
readproc(ComboBox2.ItemIndex);
end;

procedure TNewProcForm.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='Процессы';
FormEditMsiProc.Caption:='Редактор процессов';
FormEditMsiProc.ShowModal;
end;

end.

