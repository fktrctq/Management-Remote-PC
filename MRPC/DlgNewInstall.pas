unit DlgNewInstall;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Dialogs,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.variants,System.StrUtils,IdIcmpClient,ComObj,ActiveX,ShellAPI;

  type
 TStrForInstallProgram = record
    FSource :String;   /// источник файл
    FDest:string;      //файл назначение
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
    ProgramInstall:String;   // программа для установки
    KeyNewInstallProgram:string; // ключи установки программы
    InstallAllUsers:boolean; // устанавливат для всех пользователей
  end;
type
  TOKRightDlg1 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    Label1: TLabel;
    Edit1: TEdit;
    CheckBox2: TCheckBox;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    DataSource1: TDataSource;
    SpeedButton3: TSpeedButton;
    Label2: TLabel;
    CheckBox3: TCheckBox;
    ComboBox1: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button1: TButton;
    EditCopyPath: TLabeledEdit;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    SpeedButton4: TSpeedButton;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Label4: TLabel;
    SpeedButton5: TSpeedButton;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    function  readproc(item:integer):boolean;
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton3Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2Select(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1: TOKRightDlg1;
  ExitFor:Boolean;
  StrForInstall: array [0..2000] of TStrForInstallProgram; //массив переменных для запуска потоков установки программ
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;

 ThreadVar
PointForInstallMSI: ^TStrForInstallProgram;

  implementation
uses umain,MyDM,EditProcMSI;

{$R *.dfm}

function RunInstallMSI(param:pointer):boolean;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  Resinstall,b,z,i :integer;
  CopyOrNo:boolean;
  ListPC:TstringList;
////////////////////////////////////////////////////
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
      frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Запущена операция копирования');
     end;
     if TypeOperation=3 then
     begin
      pFrom := pchar('\\'+CurentPC+'\'+FDest); // что удалять  если операция удаление
      pTo := pchar('');   // куда копируем не используется
      frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Запущена операция удаления дистрибутива');
      //finditem('Запущена операция удаления дистрибутива',CurentPC,13);
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
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Операция копирования успешно завершена');
          end;
        if TypeOperation=3 then
          begin
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Операция удаления успешно завершена'); //finditem('Операция удаления успешно завершена',CurentPC,1);
          frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Операция удаления успешно завершена');
          end;
        result:=true;
       end
     else
       begin
        if TypeOperation=2 then
        begin
        finditem('Оперция копирования завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
        frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Оперция копирования завершилась ошибкой - '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        end;
        if TypeOperation=3 then
        begin
        frmdomaininfo.Memo1.Lines.Add(CurentPC+' :Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')');
        //finditem('Оперция удаления завершилась ошибкой: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',CurentPC,2);
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
       //finditem('Ошибка оперции удаления: '+E.Message,CurentPC,2);
       end;
      result:=false;
      end;
      end;
except
    on E: Exception do
    begin
     if TypeOperation=2 then frmdomaininfo.Memo1.Lines.Add(Format('Общая ошибка функции копирования перед установкой программы  :  %s',[E.Message]));
     if TypeOperation=3 then frmdomaininfo.Memo1.Lines.Add(Format('Общая ошибка функции удаления после установки программы  :  %s',[E.Message]));
     result:=false;
    end;
  end;
end;
//////////////////////////////////////////////////////////////////////
BEGIN
PointForInstallMSI:=param;
ListPC:=Tstringlist.Create;
ListPC.CommaText:=PointForInstallMSI.NamePC;
for I := 0 to ListPC.Count-1 do
Begin
 if ping(ListPC[i]) then
  begin
  CopyOrNo:=false;
  if PointForInstallMSI.BeforeInstallCopy then // если необходимо скопировать перед установкой
  begin
  CopyOrNo:=CopyFFSelectPC(ListPC[i], // имя компа
  PointForInstallMSI.UserName,//логин
  PointForInstallMSI.PassWd,  // пароль
  PointForInstallMSI.FSource, // что копировать
  PointForInstallMSI.FDest,   // куда копировать
  PointForInstallMSI.OwnerForm,// родительская форма
  PointForInstallMSI.TypeOperation, //тип операции
  PointForInstallMSI.CancelCopyFF,
  PointForInstallMSI.PathCreate) ;
  end;
      try
          frmDomainInfo.memo1.Lines.Add('Запущен процесс установки программы '+ExtractFileName(PointForInstallMSI.ProgramInstall)+' на '+ListPC[i]+'.');
          finditem('Запущен процесс установки программы '+ExtractFileName(PointForInstallMSI.ProgramInstall),ListPC[i],14);
          OleInitialize(nil);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService  := FSWbemLocator.ConnectServer(ListPC[i], 'root\CIMV2', PointForInstallMSI.UserName, PointForInstallMSI.PassWd,'','',128);
          FWbemObject   := FWMIService.Get('Win32_Product');
          Resinstall:=FWbemObject.install(PointForInstallMSI.ProgramInstall,PointForInstallMSI.KeyNewInstallProgram,PointForInstallMSI.InstallAllUsers); //файл программы, ключи установки, для всех пользователей
          if Resinstall=0 then
            Begin
            finditem('Установка программы ' +ExtractFileName(PointForInstallMSI.ProgramInstall)
            +' : '+SysErrorMessage(Resinstall),ListPC[i],1);
            End
          else
             Begin
             finditem('При установке программы '
                  +ExtractFileName(PointForInstallMSI.ProgramInstall) +' возникли ошибки : '
                  +SysErrorMessage(Resinstall),ListPC[i],2)
             End;
          frmDomainInfo.memo1.Lines.Add('Установка программы '
          +ExtractFileName(PointForInstallMSI.ProgramInstall)+' на '
          +ListPC[i]+' : '+SysErrorMessage(Resinstall));
          FWbemObject:=Unassigned;
          VariantClear(FWbemObject);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
         if PointForInstallMSI.BeforeInstallCopy and  //если указано скопировать перед установкой
         PointForInstallMSI.DeleteAfterInstall and    // если необходимо удалить дистрибутив после установки
         CopyOrNo then // если операция копирования перед установкой прошла успешно
          begin        // удаляем
          CopyFFSelectPC(ListPC[i], // имя компа
          PointForInstallMSI.UserName,//логин
          PointForInstallMSI.PassWd,  // пароль
          '', // что копировать можно не укузывать т.к. удаляем
          PointForInstallMSI.PathDelete,   // что удалять
          PointForInstallMSI.OwnerForm,// родительская форма
          3, //тип операции  (3)- удалить, FO_MOVE
          PointForInstallMSI.CancelCopyFF,
          false) ;  // не проверять наличие каталога т.к. удаляем после копирования
          end;

          OleUnInitialize;
        except
          on E:Exception do
           Begin
           finditem('Ошибка установки программы '
           +ExtractFileName(PointForInstallMSI.ProgramInstall) +' : '+E.Message,ListPC[i],2);
             frmDomainInfo.memo1.Lines.Add('Ошибка установки программы '+ExtractFileName(PointForInstallMSI.ProgramInstall)+' на '+ListPC[i]+' : "'+E.Message+'"');
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             VariantClear(FWbemObject);
             VariantClear(FWMIService);
             VariantClear(FSWbemLocator);
             OleUnInitialize;
           End;
       end; // except
  end; //ping
  End; // цикл

END;

function  TOKRightDlg1.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''msi'''
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


procedure TOKRightDlg1.Button1Click(Sender: TObject);
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

procedure TOKRightDlg1.CancelBtnClick(Sender: TObject);
begin
if (CheckBox2.Checked=false) then ExitFor:=true;
Close;
end;

procedure TOKRightDlg1.CheckBox3Click(Sender: TObject);
begin
GroupBox1.Enabled:=CheckBox3.Checked;
if CheckBox3.Checked then groupbox1.Height:=141
else  groupbox1.Height:=25;
end;

procedure TOKRightDlg1.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
ComboBox2.DroppedDown:=true;
end;

procedure TOKRightDlg1.ComboBox2Select(Sender: TObject);
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
LabeledEdit1.Text:=vartostr(FDQuery.FieldByName('PATCH_PROC').Value);
Edit1.Text:=vartostr(FDQuery.FieldByName('KEY_MSI').Value);
if FDQuery.FieldByName('KEY_MSI').Value<>null then  Edit1.Text:=FDQuery.FieldByName('KEY_MSI').AsString;
if FDQuery.FieldByName('BEFOREINSTALLCOPY').Value<>null then
begin
CheckBox3.Checked:=FDQuery.FieldByName('BEFOREINSTALLCOPY').AsBoolean;
if FDQuery.FieldByName('DELETEAFTERINSTALL').Value<>null then  CheckBox5.Checked:=FDQuery.FieldByName('DELETEAFTERINSTALL').AsBoolean;
if FDQuery.FieldByName('FILEORFOLDER').Value<>null then
  begin
  if FDQuery.FieldByName('FILEORFOLDER').AsString='File' then combobox1.ItemIndex:=0
  else combobox1.ItemIndex:=1;
  end;
if FDQuery.FieldByName('PATHCREATE').Value<>null then EditCopyPath.Text:=FDQuery.FieldByName('PATHCREATE').AsString;
end;
 finally
FDQuery.Close;
FDQuery.Free;
end;
end;


procedure TOKRightDlg1.FormClose(Sender: TObject; var Action: TCloseAction);
begin   /// если при закрытии окна, открыт список в combobox то закрываем его, иначе ошибка
label3.Visible:=false;
end;


procedure TOKRightDlg1.FormShow(Sender: TObject);
begin
label3.Visible:=false;
CheckBox2.Visible:=GroupPC;
CheckBox3.Checked:=false;
GroupBox1.Height:=25;
if not DataM.ConnectionDB.Connected then
begin
 exit;
end;
readproc(-1);
end;

procedure TOKRightDlg1.OKBtnClick(Sender: TObject);
var
i,NumPer:integer;
SelectedPCInstallProg:TstringList;
function tagForCopy(num:integer):integer;
begin
if OKBtn.tag=1999 then num:=0;
inc(num);
OKBtn.tag:=num;
result:=num;
end;

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

Function renewPath(s:string):string;
begin
if AnsiPos('$',s)=2 then
 begin
 delete(s,2,1); // удаляем символ $
 insert(':',s,2); // вставляем на его место :
 end;
result:=s;
end;
begin
   if (GroupPC=true)and (CheckBox2.Checked=false) then ///// Групповая установка программы в разных потоках
      begin
      SelectedPCInstallProg:=TstringList.Create;
      SelectedPCInstallProg:=frmdomaininfo.createListpcForCheck('');
      if SelectedPCInstallProg.Count=0 then
      begin
        ShowMessage('В списке нет выделенных компьютеров, операция завершена');
        SelectedPCInstallProg.Free;
        Exit;
      end;
      ExitFor:=false;
      label3.Caption:='%';
      label3.Visible:=true;
        for I := 0 to SelectedPCInstallProg.Count-1 do
          begin
          if ExitFor then break;
          NumPer:=tagForCopy(OKBtn.tag);
          if ComboBox1.ItemIndex=0 then // если копируем только файл
          begin
          StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
          StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // строка запуска файла для установки
          StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // путь к удаляемому файлу после установки программы
          end
          else // если копируем весь корневой каталог
          begin
          StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
          StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
          StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // путь к удаляемому каталогу после установки программы
          end;
          if not CheckBox3.Checked then // если не копируем файл или каталог перед установкой
          StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //то файл для установки берем из источника

          StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
          StrForInstall[NumPer].NamePC:=SelectedPCInstallProg[i];
          StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
          StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
          StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
          StrForInstall[NumPer].CancelCopyFF:=false;
          StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
          StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
          StrForInstall[NumPer].OwnerForm:=Self;
          StrForInstall[NumPer].NumCount:=NumPer;
          StrForInstall[NumPer].TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
          StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
          StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
          res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
          CloseHandle(res[NumPer]);
          label3.Caption:=inttostr (((100 div SelectedPCInstallProg.Count-1)*i+1))+'%';
          Application.ProcessMessages;
          sleep(500);
          end;
       if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
      end;

   if GroupPC=false then    /// установка  программы на текущий компьютер
      begin
      if not frmDomainInfo.ping(frmDomainInfo.ComboBox2.Text) then exit;
      if LabeledEdit1.text='' then
      begin
        ShowMessage('Нет пути к файлу установки');
        exit;
      end;
      NumPer:=tagForCopy(OKBtn.tag);
      if ComboBox1.ItemIndex=0 then // если копируем только файл
        begin
        StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // строка запуска файла для установки
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // путь к удаляемому файлу после установки программы
        end
      else // если копируем весь корневой каталог
        begin
        StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // путь к удаляемому каталогу после установки программы
        end;
      if not CheckBox3.Checked then // если не копируем файл или каталог перед установкой
      StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //то файл для установки берем из источника

      StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForInstall[NumPer].NamePC:=frmDomainInfo.ComboBox2.Text;
      StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
      StrForInstall[NumPer].CancelCopyFF:=false;
      StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
      StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForInstall[NumPer].OwnerForm:=Self;
      StrForInstall[NumPer].NumCount:=NumPer;
      StrForInstall[NumPer].TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
      StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
      StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
      res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      end;
    if (GroupPC=true)and (CheckBox2.Checked)  then  // групповая установка программы в одном потоке
      begin
       SelectedPCInstallProg:=TstringList.Create;
      SelectedPCInstallProg:=frmdomaininfo.createListpcForCheck('');
      if SelectedPCInstallProg.Count=0 then
        begin
        ShowMessage('В списке нет выделенных компьютеров, операция не может быть продолжена');
        SelectedPCInstallProg.Free;
        Exit;
        end;
       NumPer:=tagForCopy(OKBtn.tag);
      if ComboBox1.ItemIndex=0 then // если копируем только файл
        begin
        StrForInstall[NumPer].FSource:= LabeledEdit1.Text+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+ExtractFileName(LabeledEdit1.Text); // строка запуска файла для установки
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+ExtractFileName(LabeledEdit1.Text)+#0+#0; // путь к удаляемому файлу после установки программы
        end
      else // если копируем весь корневой каталог
        begin
        StrForInstall[NumPer].FSource:= ExtractFilePath(LabeledEdit1.Text)+#0+#0;
        StrForInstall[NumPer].ProgramInstall:=renewPath(EditCopyPath.Text)+GetLastDir(LabeledEdit1.Text)+ExtractFileName(LabeledEdit1.Text);
        StrForInstall[NumPer].PathDelete:=EditCopyPath.Text+GetLastDir(LabeledEdit1.Text)+#0+#0;  // путь к удаляемому каталогу после установки программы
        end;
        if not CheckBox3.Checked then // если не копируем файл или каталог перед установкой
          StrForInstall[NumPer].ProgramInstall:= LabeledEdit1.Text; //то файл для установки берем из источника

      StrForInstall[NumPer].FDest:=EditCopyPath.Text+#0+#0;
      StrForInstall[NumPer].NamePC:=SelectedPCInstallProg.CommaText;
      StrForInstall[NumPer].UserName:=frmDomainInfo.LabeledEdit1.Text;
      StrForInstall[NumPer].PassWd:=frmDomainInfo.LabeledEdit2.Text;
      StrForInstall[NumPer].PathCreate:=CheckBox4.Checked;
      StrForInstall[NumPer].CancelCopyFF:=false;
      StrForInstall[NumPer].BeforeInstallCopy:=CheckBox3.Checked;
      StrForInstall[NumPer].DeleteAfterInstall:=CheckBox5.Checked;
      StrForInstall[NumPer].OwnerForm:=Self;
      StrForInstall[NumPer].NumCount:=NumPer;
      StrForInstall[NumPer].TypeOperation:=2; // //2 - copy тип операции  FO_DELETE (3)- удалить, FO_MOVE (1) - переместить ,FO_RENAME (4) - переименова
      StrForInstall[NumPer].KeyNewInstallProgram:=Edit1.Text;
      StrForInstall[NumPer].InstallAllUsers:=CheckBox1.Checked;
      res[NumPer]:=BeginThread(nil,0,addr(RunInstallMSI),Addr(StrForInstall[NumPer]),0,treadID[NumPer]); ///
      CloseHandle(res[NumPer]);
      if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
      end;
if Assigned(SelectedPCInstallProg) then FreeAndNil(SelectedPCInstallProg);
OKRightDlg1.Close;
end;

procedure TOKRightDlg1.SpeedButton1Click(Sender: TObject);
var
newOpenDlg:TopenDialog;
keyrunmsi,Newdescription:string;
z:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
CheckBox3.Checked:=false; // снять галочку "копировать перед установкой"
keyrunmsi:='';
Newdescription:='';
newOpenDlg:=TOpendialog.Create(self);
newOpenDlg.Title:='Укажиет файл для установки';
newOpenDlg.Filter:='|*msi';
if newOpenDlg.Execute then
  begin
   keyrunmsi:=InputBox('Добавление ключа установки, не обязательно', 'ключ:',
    '');
   if keyrunmsi<>'' then
     begin
     z:=MessageDlg('Вы указали ключ запуска,'+#10#13+' это не обязательно! Продолжить???',
     mtConfirmation, mbOKCancel, 0);
     if z=mrCancel then begin newOpenDlg.Free; exit;  end;
     end;
   Newdescription:=InputBox('Добавление описания к файлу', 'описание:', 'здесь ваше описание');
   // если не указали описание то добавляем в описание имя выбранного файла
   if (Newdescription='') and (fileexists(NewOpenDlg.FileName)) then Newdescription:=ExtractFileName(NewOpenDlg.FileName);
  end;

if (Newdescription='') or (NewOpenDlg.FileName='') then exit;
// записываем новую запись в базу
FDQueryProc.SQL.Clear;
//BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE
FDQueryProc.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,KEY_MSI,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(newOpenDlg.FileName,1000)+''; // путь к файлу
FDQueryProc.params.ParamByName('p2').AsString:=''+'msi'+'';                            // тип
FDQueryProc.params.ParamByName('p3').AsString:=''+leftstr(Newdescription,1000)+'';     //Описание
FDQueryProc.params.ParamByName('p4').AsString:=''+leftstr(keyrunmsi,250)+'';          // ключи запуска
FDQueryProc.params.ParamByName('p5').AsBoolean:=CheckBox3.Checked; // копировать файлы перед установкой
FDQueryProc.params.ParamByName('p6').AsBoolean:=CheckBox5.Checked; // удалить файлы после установки
if ComboBox1.ItemIndex=0 then // если копируем только файл
FDQueryProc.params.ParamByName('p7').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // если копируем родительский каталог
FDQueryProc.params.ParamByName('p7').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //директория куда копировать
FDQueryProc.ExecSQL;
// открываем список из базы
readproc(ComboBox2.Items.Count); //после сохранения выбрать последний элемент списка
///уничтожаем  диалог
newOpenDlg.Free;
end;

procedure TOKRightDlg1.SpeedButton2Click(Sender: TObject);
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
FDQueryProc.Active:=false;
FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''
+ComboBox3.Text+'''';
FDQueryProc.ExecSQL;
// открываем список  процессов из базы
readproc(ComboBox2.ItemIndex-1);
end;

procedure TOKRightDlg1.SpeedButton3Click(Sender: TObject);
begin
if OKRightDlg1.findComponent('MyDialog')=nil then
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор файла для установки';
  Filter:='|*.msi';
  if Execute then
    begin
    LabeledEdit1.text:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
end;
end;



procedure TOKRightDlg1.SpeedButton4Click(Sender: TObject); /// сохранить текущие настройки для выбранной программы
var
descript:string;
begin

if (LabeledEdit1.Text='') then
begin
  ShowMessage('Укажите путь к файлу установки и описание программы');
  SpeedButton1.Click;
  exit;
end;

descript:=ComboBox2.Text;

if (ComboBox2.Text='')and (LabeledEdit1.Text<>'') then
 begin
 descript:=InputBox('Добавьте описания к программе', 'Описание:', 'здесь описание программы');
 end;
 if descript='' then exit;


FDQueryProc.SQL.Clear;
FDQueryProc.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE,KEY_MSI)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)MATCHING (DESCRIPTION_PROC)';
FDQueryProc.params.ParamByName('p1').AsString:=''+leftstr(LabeledEdit1.Text,1000)+'';
FDQueryProc.params.ParamByName('p2').AsString:=''+'msi'+'';
FDQueryProc.params.ParamByName('p3').AsString:=''+descript+'';
FDQueryProc.params.ParamByName('p4').AsBoolean:=CheckBox3.Checked; // копировать файлы перед установкой
FDQueryProc.params.ParamByName('p5').AsBoolean:=CheckBox5.Checked; // удалить файлы после установки
if ComboBox1.ItemIndex=0 then // если копируем только файл
FDQueryProc.params.ParamByName('p6').AsString:=''+'File'+'';
if ComboBox1.ItemIndex=1 then // если копируем родительский каталог
FDQueryProc.params.ParamByName('p6').AsString:=''+'Folder'+'';
FDQueryProc.params.ParamByName('p7').AsString:=''+leftstr(EditCopyPath.Text,1024)+''; //директория куда копировать
FDQueryProc.params.ParamByName('p8').AsString:=''+leftstr(Edit1.Text,250)+'';
FDQueryProc.ExecSQL;
//readproc(combobox2.ItemIndex);
end;

procedure TOKRightDlg1.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='Программы MSI';
FormEditMsiProc.Caption:='Редактор программ';
FormEditMsiProc.ShowModal;
end;

end.
