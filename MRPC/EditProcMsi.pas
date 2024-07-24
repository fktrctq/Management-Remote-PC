unit EditProcMsi;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.StdCtrls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Buttons, JvExButtons, JvButtons,System.strUtils,
  Vcl.Menus;

type
  TFormEditMsiProc = class(TForm)
    ListView1: TListView;
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    ButtonPM: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    SpeedButton1: TSpeedButton;
    procedure ListView1DblClick(Sender: TObject);
    procedure ButtonPMClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure N1Click(Sender: TObject);

  private
    FormKey:Tform;
    EditDescr:TLabelEdedit;
    EditPath:TLabelEdedit;
    CheckCopy:TcheckBox;
    ComboFF:TcomboBox;
    EditCreatePath:TLabelEdEdit;
    EditPathFileRun:TlabelEdEdit;
    ButtonOk:Tbutton;
    ButtonNo:Tbutton;
    FormMSI:Tform;
    EditDescrM:TLabelEdedit;
    EditPathM:TLabelEdedit;
    CheckCopyM:TcheckBox;
    CheckDelM:TcheckBox;
    ComboFFM:TcomboBox;
    EditCreatePathM:TLabelEdEdit;
    EditPathFileRunM:TlabelEdEdit;
    ButtonOkM:Tbutton;
    ButtonNoM:Tbutton;
    ButtonOpen:TSpeedbutton;
    function CreateEditFormProc(Descr,path,FileSource,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
    Function CreateEditFormMSI(Descr,FileRunMSI,OptionKey,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
    function updateList(ProcMsi:string):boolean;
    procedure ButtonNoClose(Sender: TObject);
    procedure KeyFormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonSaveMSI(Sender: TObject);
    procedure ButtonNewMSI(Sender: TObject);
    procedure ButtonSaveProc(Sender: TObject);
    procedure ButtonNewProc(Sender: TObject);
    procedure ButtonOpenDlgProc(Sender: TObject);
    procedure ButtonOpenDlgMSI(Sender: TObject);
    procedure updateListPM;

  public
    { Public declarations }
  end;

var
  FormEditMsiProc: TFormEditMsiProc;


implementation
uses MyDM,umain;
{$R *.dfm}

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

function yestobool(s:string):boolean;
begin
if s='Да' then Result:=true
else Result:=false;
end;


procedure TFormEditMsiProc.updateListPM;
begin
 if ButtonPM.Caption='Процессы' then  // если надпись процессы то загружены программы и наоборот
begin
  updateList('msi');
end
else
begin
 updateList('proc');
end;
end;


procedure TFormEditMsiProc.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).close;
end;

procedure TFormEditMsiProc.ButtonPMClick(Sender: TObject);
begin
if ButtonPM.Caption='Процессы' then
begin
  ListView1.Column[3].Width:=0; //колонка с аргументами MSI
  ListView1.Column[5].Width:=0; // Удалить после установки/запуска
  ListView1.Column[8].Width:=200; // файл источник
  ListView1.Column[2].Caption:='Строка запуска с аргументами';
  ButtonPM.Caption:='Программы MSI';
  updateList('proc');
  FormEditMsiProc.caption:='Редактор процессов';
end
else
begin
 ListView1.Column[3].Width:=100; //колонка с аргументами MSI
 ListView1.Column[5].Width:=100; // Удалить после установки/запуска
 ListView1.Column[8].Width:=0; // файл источник
 ListView1.Column[2].Caption:='Файл установки';
 ButtonPM.Caption:='Процессы';
 updateList('msi');
 FormEditMsiProc.caption:='Редактор программ msi';
end;
end;

procedure TFormEditMsiProc.KeyFormClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;


procedure TFormEditMsiProc.ListView1DblClick(Sender: TObject);
begin
if ListView1.SelCount=1 then
begin
if ButtonPM.Caption='Программы MSI' then
begin
CreateEditFormProc(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[7],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end
else
begin
CreateEditFormMSI(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[2],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end;
end;
end;

procedure TFormEditMsiProc.N1Click(Sender: TObject); // создание новой записи
begin
if ListView1.SelCount=1 then // если есть выделеннаЯ строка то вводим данные в поля
  begin
   if ButtonPM.Caption='Программы MSI' then // если кнопка программы то выведен список процессов и наоборот
    begin
    CreateEditFormProc(
    ListView1.Selected.SubItems[0],
    ListView1.Selected.SubItems[1],
    ListView1.Selected.SubItems[7],
    ListView1.Selected.SubItems[5],
    ListView1.Selected.SubItems[6],
    yestobool(ListView1.Selected.SubItems[3]),
    yestobool(ListView1.Selected.SubItems[4])
    ,true)
    end
   else 
    begin
    CreateEditFormMSI(
    ListView1.Selected.SubItems[0],
    ListView1.Selected.SubItems[1],
    ListView1.Selected.SubItems[2],
    ListView1.Selected.SubItems[5],
    ListView1.Selected.SubItems[6],
    yestobool(ListView1.Selected.SubItems[3]),
    yestobool(ListView1.Selected.SubItems[4])
    ,true)
    end;
  end
else // иначе если строка не выделена или выделено больше чем 1 то поля создания оставляем пустыми
  begin
  if ButtonPM.Caption='Программы MSI' then
  CreateEditFormProc('','','','','C$\TEMP\',false,false,true)// форма процесса
  else
  CreateEditFormMSI('','','','','C$\TEMP\',false,false,true); // форма MSI
  end;
end;

procedure TFormEditMsiProc.Button1Click(Sender: TObject);
begin
if ButtonPM.Caption='Программы MSI' then
CreateEditFormProc('','','','','C$\TEMP\',false,false,true)// форма процесса
else
CreateEditFormMSI('','','','','C$\TEMP\',false,false,true); // форма MSI
end;

procedure TFormEditMsiProc.Button2Click(Sender: TObject);
begin
if ListView1.SelCount=1 then
begin
if ButtonPM.Caption='Программы MSI' then
begin
CreateEditFormProc(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[7],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end
else
begin
CreateEditFormMSI(
ListView1.Selected.SubItems[0],
ListView1.Selected.SubItems[1],
ListView1.Selected.SubItems[2],
ListView1.Selected.SubItems[5],
ListView1.Selected.SubItems[6],
yestobool(ListView1.Selected.SubItems[3]),
yestobool(ListView1.Selected.SubItems[4])
,false)
end;
end;
end;



procedure TFormEditMsiProc.Button3Click(Sender: TObject);
begin
updateListPM;
end;

procedure TFormEditMsiProc.Button4Click(Sender: TObject); // удаление записи
var
i,z:integer;
QueryDel: TFDQuery;
TransactionDel:TFDtransaction;
begin
if ListView1.SelCount=0 then  exit;
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
i:=MessageDlg('Перед удалением убедитесь что записи не используется в задачах.'+#10#13+'Продолжить выполнение операции?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
try
TransactionDel:= TFDTransaction.Create(nil);
TransactionDel.Connection:=DataM.ConnectionDB;
TransactionDel.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryDel:=TFDQuery.Create(nil);
QueryDel.Transaction:=TransactionDel;
QueryDel.Connection:=DataM.ConnectionDB;
try
for z := ListView1.Items.Count-1 downto 0 do
if ListView1.Items[z].Selected then
begin
TransactionDel.StartTransaction;
QueryDel.SQL.Clear;
QueryDel.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ListView1.Items[z].Caption+''''; // в Caption выводятся ID процессов или MSI
QueryDel.ExecSQL;
end;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('Ошибка удаления. '+e.Message);
end;
finally
QueryDel.Close;
TransactionDel.Commit;
TransactionDel.Free;
QueryDel.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonOpenDlgProc(Sender: TObject);
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор файла для запуска';
  if EditPathFileRun.Text<>'' then if FileExists(EditPathFileRun.Text) then InitialDir:=ExtractFileDir(EditPathFileRun.Text) ;
  //Options:=[fdoForceShowHidden,fdoPickFolders]; {каталоги}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
    EditPath.Text:=FileName;
    //EditPathFileRun.Text:=DeliteKeyFilePatch(EditPath.Text)
    EditPathFileRun.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;

procedure TFormEditMsiProc.ButtonOpenDlgMSI(Sender: TObject);
begin
 with TOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Выбор файла для установки';
  Filter:='|*msi';
  if EditPathM.Text<>'' then if FileExists(EditPathM.Text) then InitialDir:=ExtractFileDir(EditPathM.Text) ;
  if Execute then
    begin
    EditPathM.Text:=FileName;
    end;
   finally
   Destroy;
   end;
 end;
end;

procedure TFormEditMsiProc.ButtonSaveMSI(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.clear;
QueryWrite.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE,KEY_MSI)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)MATCHING (DESCRIPTION_PROC)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPathM.Text,1000)+''; //путь к файлу
QueryWrite.params.ParamByName('p2').AsString:=''+'msi'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescrM.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopyM.Checked; // копировать файлы перед установкой
QueryWrite.params.ParamByName('p5').AsBoolean:=CheckDelM.Checked; // удалить файлы после установки
if ComboFFM.ItemIndex=0 then // если копируем только файл
QueryWrite.params.ParamByName('p6').AsString:=''+'File'+'';
if ComboFFM.ItemIndex=1 then // если копируем родительский каталог
QueryWrite.params.ParamByName('p6').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p7').AsString:=''+leftstr(EditCreatePathM.text,1024)+''; //директория куда копировать
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRunM.text,250)+'';  // опции/ключи запуска
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('Ошибка сохранения. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonNewMSI(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.clear;
QueryWrite.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,KEY_MSI,BEFOREINSTALLCOPY,DELETEAFTERINSTALL,FILEORFOLDER,PATHCREATE) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPathM.Text,1000)+''; // путь к файлу
QueryWrite.params.ParamByName('p2').AsString:=''+'msi'+'';                            // тип
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescrM.Text,1000)+'';     //Описание
QueryWrite.params.ParamByName('p4').AsString:=''+leftstr(EditPathFileRunM.text,250)+'';          // ключи запуска
QueryWrite.params.ParamByName('p5').AsBoolean:=CheckCopyM.Checked; // копировать файлы перед установкой
QueryWrite.params.ParamByName('p6').AsBoolean:=CheckDelM.Checked; // удалить файлы после установки
if ComboFFM.ItemIndex=0 then // если копируем только файл
QueryWrite.params.ParamByName('p7').AsString:=''+'File'+'';
if ComboFFM.ItemIndex=1 then // если копируем родительский каталог
QueryWrite.params.ParamByName('p7').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditCreatePathM.text,1024)+''; //директория куда копировать
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('Ошибка создания новой программы. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonSaveProc(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
QueryWrite.SQL.Clear;
QueryWrite.SQL.Text:='update or insert into START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8) MATCHING (DESCRIPTION_PROC)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,1000)+'';
QueryWrite.params.ParamByName('p2').AsString:=''+'proc'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopy.Checked; // копировать файл или каталог перед установкой
if ComboFF.ItemIndex=0 then // если копируем только файл
QueryWrite.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboFF.ItemIndex=1 then // если копируем родительский каталог
QueryWrite.params.ParamByName('p5').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p6').AsString:=''+leftstr(EditCreatePath.Text,1024)+''; //директория куда копировать
QueryWrite.params.ParamByName('p7').AsBoolean:=false; // удалить файлы после установки
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRun.Text,1000)+'';    // путь до файла источника
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('Ошибка сохранения процесса. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

procedure TFormEditMsiProc.ButtonNewProc(Sender: TObject);
var
QueryWrite: TFDQuery;
TransactionWrite:TFDtransaction;
begin
try
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryWrite:=TFDQuery.Create(nil);
QueryWrite.Transaction:=TransactionWrite;
QueryWrite.Connection:=DataM.ConnectionDB;
try
TransactionWrite.StartTransaction;
EditPathFileRun.Text:=DeliteKeyFilePatch(EditPath.Text); // заполняем поле файл для копирования
QueryWrite.SQL.Clear;
QueryWrite.SQL.Text:='INSERT INTO START_PROC_MSI (PATCH_PROC,TYPE_PROC'
+',DESCRIPTION_PROC,BEFOREINSTALLCOPY,FILEORFOLDER,PATHCREATE,DELETEAFTERINSTALL,FILESOURSE_PROC) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
QueryWrite.params.ParamByName('p1').AsString:=''+leftstr(EditPath.Text,1000)+'';
QueryWrite.params.ParamByName('p2').AsString:=''+'proc'+'';
QueryWrite.params.ParamByName('p3').AsString:=''+leftstr(EditDescr.Text,1000)+'';
QueryWrite.params.ParamByName('p4').AsBoolean:=CheckCopy.Checked; // копировать файлы перед установкой
if ComboFF.ItemIndex=0 then // если копируем только файл
QueryWrite.params.ParamByName('p5').AsString:=''+'File'+'';
if ComboFF.ItemIndex=1 then // если копируем родительский каталог
QueryWrite.params.ParamByName('p5').AsString:=''+'Folder'+'';
QueryWrite.params.ParamByName('p6').AsString:=''+leftstr(EditCreatePath.Text,1024)+''; //директория куда копировать
QueryWrite.params.ParamByName('p7').AsBoolean:=false; // удалить файлы после установки
QueryWrite.params.ParamByName('p8').AsString:=''+leftstr(EditPathFileRun.Text,1000)+'';    // путь до файла источника
QueryWrite.ExecSQL;
except on E: Exception do
  frmDomainInfo.Memo1.Lines.Add('Ошибка создания процесса. '+e.Message);
end;
finally
QueryWrite.Close;
TransactionWrite.Commit;
TransactionWrite.Free;
QueryWrite.Free;
end;
updateListPM;
end;

Function TFormEditMsiProc.CreateEditFormMSI(Descr,FileRunMSI,OptionKey,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
var
step:integer;
begin
try
step:=1;
FormKey:=TForm.Create(FormEditMsiProc);
if NewOrEdit then FormKey.Caption:='Новая программа'
else FormKey.Caption:='Редактировать *.msi '+descr;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=360;   //285
FormKey.Height:=290;
FormKey.OnClose:=KeyFormClose;
////////////
step:=2;
EditDescrM:=TLabelEdedit.Create(FormKey);
EditDescrM.Parent:=FormKey;
EditDescrM.EditLabel.caption:='Описание';
EditDescrM.Text:=Descr;
EditDescrM.Top:=15;
EditDescrM.Left:=5;
EditDescrM.Width:=335;
/// /////////////////////
step:=3;
EditPathM:=TLabelEdedit.Create(FormKey);
EditPathM.Parent:=FormKey;
EditPathM.EditLabel.Caption:='Файл запуска';
EditPathM.Text:=FileRunMSI;
EditPathM.Top:=55;
EditPathM.Left:=5;
EditPathM.Width:=308;//335
/////////////////////////
ButtonOpen:=TSpeedbutton.Create(FormKey);
ButtonOpen.Parent:=FormKey;
ButtonOpen.Caption:='';
ButtonOpen.Height:=23;
ButtonOpen.Width:=23;
ButtonOpen.Glyph:=SpeedButton1.Glyph;
ButtonOpen.Flat:=true;
ButtonOpen.Top:=54;
ButtonOpen.Left:=315;
ButtonOpen.OnClick:=ButtonOpenDlgMSI;
/////////////////////////
step:=4;
EditPathFileRunM:=TLabeledEdit.Create(FormKey);
EditPathFileRunM.Parent:=FormKey;
EditPathFileRunM.EditLabel.Caption:='Опции, не обязательно (посмотреть опции msiexec /help)';
EditPathFileRunM.Text:=OptionKey;
EditPathFileRunM.Top:=95;
EditPathFileRunM.Left:=5;
EditPathFileRunM.Width:=335;
////////////////////////
step:=5;
CheckCopyM:=TcheckBox.Create(FormKey);
CheckCopyM.Parent:=FormKey;
CheckCopyM.Checked:=Fcopy;
CheckCopyM.Top:=115;
CheckCopyM.Left:=5;
CheckCopyM.Width:=335;
CheckCopyM.Caption:='Копировать дистрибутив перед запуском';
//////////////////
step:=6;
CheckDelM:=TcheckBox.Create(FormKey);
CheckDelM.Parent:=FormKey;
CheckDelM.Checked:=Fdel;
CheckDelM.Top:=130;
CheckDelM.Left:=5;
CheckDelM.Width:=335;
CheckDelM.Caption:='Удалить скопированные файлы после установки';
///////////////////
step:=7;
ComboFFM:=TComboBox.Create(FormKey);
ComboFFM.parent:=FormKey;
ComboFFM.Style:=csOwnerDrawFixed;
ComboFFM.Top:=150;
ComboFFM.Left:=5;
ComboFFM.Width:=335;
ComboFFM.Items.Add('Копировать только файл установки (msi)');
ComboFFM.Items.Add('Копировать родительский каталог файла установки');
if FF='Folder' then ComboFFM.ItemIndex:=1
else ComboFFM.ItemIndex:=0;
//////////////////
step:=8;
EditCreatePathM:=TLabelEdEdit.Create(FormKey);
EditCreatePathM.parent:=FormKey;
EditCreatePathM.EditLabel.Caption:='Каталог назначения';
EditCreatePathM.Text:=CreatePath;
EditCreatePathM.Top:=190;
EditCreatePathM.Left:=5;
EditCreatePathM.width:=335;
//////////////////
ButtonOkM:=Tbutton.Create(FormKey);
ButtonOkM.Parent:=FormKey;
ButtonOkM.Caption:='Сохранить';
ButtonOkM.Top:=220;
ButtonOkM.Left:=185;
if NewOrEdit then ButtonOkM.OnClick:=ButtonNewMSI  // если создать новый
else ButtonOkM.OnClick:=ButtonSaveMSI;  // еиначе редактировать существующий
ButtonOkM.TabOrder:=3;
//////////////////
ButtonNoM:=Tbutton.Create(FormKey);
ButtonNoM.Parent:=FormKey;
ButtonNoM.Caption:='Закрыть';
ButtonNoM.Top:=220;
ButtonNoM.Left:=265;
ButtonNoM.OnClick:=ButtonNoClose;
ButtonNoM.TabOrder:=4;
result:=true;
FormKey.ShowModal;
except on E: Exception do
begin
  frmDomainInfo.Memo1.Lines.Add('Ошибка создания диалога редактирования msi. ('+inttostr(step)+') '+e.Message);
  result:=false;
end;
end;
end;

function TFormEditMsiProc.CreateEditFormProc(Descr,path,FileSource,FF,CreatePath:string;Fcopy,Fdel,NewOrEdit:boolean):boolean;
begin
try
FormKey:=TForm.Create(FormEditMsiProc);
if NewOrEdit then  FormKey.Caption:='Новый процесс'
else FormKey.Caption:='Редактировать процесс '+descr;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=360;   
FormKey.Height:=275;
FormKey.OnClose:=KeyFormClose;
////////////
EditDescr:=TLabelEdedit.Create(FormKey);
EditDescr.Parent:=FormKey;
EditDescr.EditLabel.caption:='Описание';
EditDescr.Text:=Descr;
EditDescr.Top:=15;
EditDescr.Left:=5;
EditDescr.Width:=335;
/// /////////////////////
EditPath:=TLabelEdedit.Create(FormKey);
EditPath.Parent:=FormKey;
EditPath.EditLabel.Caption:='Строка запуска с аргументами';
EditPath.Text:=path;
EditPath.Top:=55;
EditPath.Left:=5;
EditPath.Width:=308;
/////////////////////////
ButtonOpen:=TSpeedbutton.Create(FormKey);
ButtonOpen.Parent:=FormKey;
ButtonOpen.Caption:='';
ButtonOpen.Height:=23;
ButtonOpen.Width:=23;
ButtonOpen.Glyph:=SpeedButton1.Glyph;
ButtonOpen.Flat:=true;
ButtonOpen.Top:=54;
ButtonOpen.Left:=315;
ButtonOpen.OnClick:=ButtonOpenDlgProc;
///*/////////////////////
EditPathFileRun:=TLabeledEdit.Create(FormKey);
EditPathFileRun.Parent:=FormKey;
EditPathFileRun.EditLabel.Caption:='Файл источник';
EditPathFileRun.Text:=FileSource;
EditPathFileRun.Top:=95;
EditPathFileRun.Left:=5;
EditPathFileRun.Width:=335;
////////////////////////
CheckCopy:=TcheckBox.Create(FormKey);
CheckCopy.Parent:=FormKey;
CheckCopy.Checked:=Fcopy;
CheckCopy.Top:=115;
CheckCopy.Left:=5;
CheckCopy.Width:=335;
CheckCopy.Caption:='Копировать файл перед запуском';
//////////////////
ComboFF:=TComboBox.Create(FormKey);
ComboFF.parent:=FormKey;
ComboFF.Items.Add('Копировать только файл');
ComboFF.Items.Add('Копировать родительский каталог файла');
if FF='Folder' then ComboFF.ItemIndex:=1
else ComboFF.ItemIndex:=0;
ComboFF.Style:=csOwnerDrawFixed;
ComboFF.Top:=135;
ComboFF.Left:=5;
ComboFF.Width:=335;
//////////////////
EditCreatePath:=TLabelEdEdit.Create(FormKey);
EditCreatePath.parent:=FormKey;
EditCreatePath.EditLabel.Caption:='Каталог назначения';
EditCreatePath.Text:=CreatePath;
EditCreatePath.Top:=175;
EditCreatePath.Left:=5;
EditCreatePath.width:=335;
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Сохранить';
ButtonOk.Top:=205;
ButtonOk.Left:=185;
if NewOrEdit then ButtonOk.OnClick:=ButtonNewProc  // если нажали создать новый
else ButtonOk.OnClick:=ButtonSaveProc;  // если редактировать
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Закрыть';
ButtonNo.Top:=205;
ButtonNo.Left:=265;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
result:=true;
FormKey.ShowModal;
except on E: Exception do
begin
  frmDomainInfo.Memo1.Lines.Add('Ошибка создания диалога редактирования процесса. '+e.Message);
  result:=false;
end;
end;
end;


procedure TFormEditMsiProc.FormShow(Sender: TObject);
begin
if ButtonPM.Caption='Процессы' then
begin
  ListView1.Column[3].Width:=0; //колонка с аргументами MSI
  ListView1.Column[5].Width:=0; // Удалить после установки/запуска
  ListView1.Column[8].Width:=100; // файл источник
  ListView1.Column[2].Caption:='Строка запуска с аргументами';
  updateList('proc');
  ButtonPM.Caption:='Программы MSI';
end
else
begin
 ListView1.Column[3].Width:=100; //колонка с аргументами MSI
 ListView1.Column[5].Width:=100; // Удалить после установки/запуска
 ListView1.Column[8].Width:=0; // файл источник
 ListView1.Column[2].Caption:='Файл установки';
 updateList('msi');
 ButtonPM.Caption:='Процессы';
end;
end;

function TFormEditMsiProc.updateList(ProcMsi:string):boolean;
var
QueryRead: TFDQuery;
TransactionRead:TFDtransaction;
begin
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
TransactionRead.Options.Isolation:=xiSnapshot; ///xiReadCommitted; //xiSnapshot;
QueryRead:=TFDQuery.Create(nil);
QueryRead.Transaction:=TransactionRead;
QueryRead.Connection:=DataM.ConnectionDB;
try
ListView1.Clear;
TransactionRead.StartTransaction;
QueryRead.SQL.clear;
QueryRead.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC='''+ProcMsi+''''
+' ORDER BY DESCRIPTION_PROC'; //ID_PROC
QueryRead.Open;

while not QueryRead.Eof do
begin
 if QueryRead.FieldByName('ID_PROC').Value<>null then
 begin
 with ListView1.Items.Add do
 begin
   Caption:=QueryRead.FieldByName('ID_PROC').AsString;
   SubItems.add(QueryRead.FieldByName('DESCRIPTION_PROC').AsString);
   SubItems.add(QueryRead.FieldByName('PATCH_PROC').AsString);
   SubItems.add(QueryRead.FieldByName('KEY_MSI').AsString);
   if QueryRead.FieldByName('BEFOREINSTALLCOPY').AsBoolean then SubItems.add('Да')
   else  SubItems.add('Нет');
   if QueryRead.FieldByName('DELETEAFTERINSTALL').AsBoolean then SubItems.add('Да')
   else  SubItems.add('Нет');
   SubItems.add(QueryRead.FieldByName('FILEORFOLDER').AsString);
   SubItems.add(QueryRead.FieldByName('PATHCREATE').AsString);
   SubItems.add(QueryRead.FieldByName('FILESOURSE_PROC').AsString);
 end;
 end;
  QueryRead.Next;
end;

finally
QueryRead.Close;
TransactionRead.Commit;
TransactionRead.Free;
QueryRead.Free;
end;

except on E: Exception do
frmDomainInfo.Memo1.Lines.Add('Ошибка чтения списка '+e.Message);
end;
end;

end.
