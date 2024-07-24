unit TaskEdit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.ImgList, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls, Vcl.Menus,IniFiles,DateUtils,ActiveX,ComObj,CommCtrl;

type
  TEditTask = class(TForm)
    Panel1: TPanel;
    ListView1: TListView;
    Button1: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel2: TPanel;
    ListView2: TListView;
    Button2: TButton;
    Button3: TButton;
    ImageList1: TImageList;
    Button4: TButton;
    Button5: TButton;
    FDQueryProcMSI: TFDQuery;
    FDTransactionReadProcMSI: TFDTransaction;
    PopupMSI: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Button7: TButton;
    TaskName: TLabel;
    Button8: TButton;
    ImageList2: TImageList;
    Button6: TButton;
    Button9: TButton;
    PopupProc: TPopupMenu;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Button10: TButton;
    PopupPower: TPopupMenu;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    WakeOnLan1: TMenuItem;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button11: TButton;
    PopupDetalTask: TPopupMenu;
    N10: TMenuItem;
    N11: TMenuItem;
    PopupDetalTaskPC: TPopupMenu;
    N12: TMenuItem;
    N13: TMenuItem;
    Button12: TButton;
    PopupService: TPopupMenu;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    Button13: TButton;
    Button14: TButton;
    PopupActivation: TPopupMenu;
    PopupFileFolder: TPopupMenu;
    Windows1: TMenuItem;
    Office1: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    Windows71: TMenuItem;
    Win881101: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    Button15: TButton;
    Button16: TButton;
    PopupRegEdit: TPopupMenu;
    N25: TMenuItem;
    N26: TMenuItem;
    N27: TMenuItem;
    N28: TMenuItem;
    N29: TMenuItem;
    SZ1: TMenuItem;
    Button17: TButton;
    PopupRestore: TPopupMenu;
    N30: TMenuItem;
    N31: TMenuItem;
    N32: TMenuItem;
    procedure Button2Click(Sender: TObject);  /// заполняем combobox списком групп
    procedure FormGrClose(Sender: TObject; var Action: TCloseAction);
    procedure FormGRShow(Sender: TObject);
    procedure ButOkClick(Sender: TObject);
    procedure LVPCColumnClick(Sender: TObject;Column: TListColumn);
    procedure LVPCDblClick(Sender: TObject); // добавление компа двойным кликом
    procedure LVPCSelectItem(Sender: TObject; Item: TListItem; // выделение shift
  Selected: Boolean);
    procedure ButCloseFormGRClick(Sender: TObject);
    procedure ListGrSelect(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ListView2ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListGrKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure ListGrKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnShow(Sender: TObject);
    function creatformNewtask(S:string):boolean; // создание формы для избранных процессов и программ
    function creatformUserPass(S:string):boolean; // создание формы для логина и пароля
    function creatformService(S:string):boolean; // Создание формы для служб
    procedure ButtonAddUser(Sender: TObject); // Добавляем пользователя и пароль
    procedure FormMsiOpen(Sender: TObject); // список избранного MSI
    procedure FormProcOpen(Sender: TObject); // список избранного процессы
    procedure AddTaskInListview(Sender: TObject); /// выбор процесса или msi из списка
    procedure SelectComboMsiProc(Sender: TObject);
    procedure SelectComboProcMSI(Sender: TObject);
    procedure AddServiceInListview(Sender: TObject); // Добавляем службу в список задач
    procedure N1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject); // при выборе из списка процессов или msi выбирать его номер во втором combobox
    function CreateNewTableTask(TableName:string):boolean;
    function UpdateTableTABLETASK(NameTable:string;countPC:integer):boolean;     //Обновление таблицы для задачи
    procedure Button7Click(Sender: TObject);     //функция запуска создания новой таблицы+ее заполнение
    procedure SaveTask(Sender: TObject);
    function AddNewItemsTable(DescriptionTable:string;countPC:integer):string;
    function AddOrUpdateListViewreadInTableTask(nameTable:string):boolean; // чтение записи в таблице Table_task и вывод информации о ней в гланый ListView
    procedure Button8Click(Sender: TObject);
    function UpdateTableTask(TableName:string):boolean; // обновление таблицы для задачи
    function TransactionAutoCommit(auto:boolean):boolean;
    procedure ListView1Changing(Sender: TObject; Item: TListItem;
      Change: TItemChange; var AllowChange: Boolean);
    procedure ListView2Changing(Sender: TObject; Item: TListItem;
      Change: TItemChange; var AllowChange: Boolean);
    procedure Button6Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button9Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N3Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure WakeOnLan1Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    Function CreateFormDetail(s:string;z:boolean):boolean;
    procedure FormDetailPCShow(Sender: TObject); // подробная информация о выполненыз задачах на выделенных ПК
    procedure FormDetailTaskShow(Sender: TObject); // Подробная инфа о выделенной задаче
    procedure ButDetailSaveClick(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure ListView2DblClick(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure ListView1DblClick(Sender: TObject);
    function LoadListService(NamePC:string):boolean; /// загрузить список служб
    procedure ButLoadService(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure Windows1Click(Sender: TObject);
    procedure Windows71Click(Sender: TObject);
    procedure Win881101Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
    function  deletelistPC(z:bool):boolean;
    procedure Button15Click(Sender: TObject);
    procedure N27Click(Sender: TObject);
    procedure N28Click(Sender: TObject);
    procedure SZ1Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure N31Click(Sender: TObject);
    procedure N32Click(Sender: TObject);
    //procedure N27Click(Sender: TObject); //удалить весь список компов
 private
    FormKey:Tform;
    EditsSubKeyName,EditNameKey:TlabeledEdit;
    ComboRootKey:TcomboBox;
    ButtonOk,ButtonNo:Tbutton;
    function NewsSubKeyName(CreateDel:string):boolean; //Форма для создания раздела
    function DeleteKeyName:boolean; // создать форму для удаления ключа реестра
    procedure ButtonNoClose(Sender: TObject);
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonOksSubKeyName(Sender: TObject); // добавить в задачу задание для создания раздела
    procedure ButtonOkDelKey(Sender: TObject); // добавить задание удаление ключа
    function DateForTaskRestore(var DateRestore:string):boolean; // дата для восстановления точки
    function DescriptionForRestore(var DesRestore:string):boolean; //описание для новой точки восстановления
  public
    function StatusStartStopTask(DescriptionTask:string):boolean; // проверка запущена или остановлена задача
    function nametableForDescription(Descript:string):string;   // Получаем имя таблицы по описанию
    function ReadResulTask(DescriptionTable:string):bool;  // чтение таблицы и результатов выполнения задачи
    function StatrTask(DescriptionTable:string):boolean;  // запуск задачи
    function AddUserPass(DescriptionTable,User,Passwd:string;SavePass:boolean):boolean; // добавление пользователя и пароля
    function StopTask(DescriptionTask:string):boolean;    // остановка задачи
    function RenameTableForTask(NameTable,NewDescriptionTable:string):boolean; //Меняем описание таблицы для задачи. обновлят запись в таблице TABLE_TASK
    function ThereAnyRunStopTask(statustask:boolean):integer; /// читам статус задачи, и смотрим есть ли запущенные или остановленные

  end;


var
 EditTask: TEditTask;


implementation
uses umain,MyDM,RunTask,TaskNewMSI,TasknewProc,TaskSelectDelMSI,TaskCopyFF,RegEditKeySave;

var
  FormGr:Tform;
  ButOk:Tbutton;
  ButLoad:Tbutton;
  ButClose:Tbutton;
  PanelB:Tpanel;
  PanelG:Tpanel;
  LVPC:Tlistview;
  ListGr:Tcombobox;
  ComboNumMisProc:Tcombobox;
  SortLV:boolean;
  sortInt:integer;
  EditUser:Tedit;
  EditPass:Tedit;
  ChekSave:TcheckBox;
{$R *.dfm}

///////////////////////////////////////////////////////////
function Log_write(fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
exit;
end;

end;
//////////////////////////////////////////////////////////////////////////
function TEditTask.TransactionAutoCommit(auto:boolean):boolean;
begin
//if auto then FDTransactionReadProcMSI.Connection.StartTransaction
//else FDTransactionReadProcMSI.Commit;
//FDTransactionReadProcMSI.Options.AutoCommit:=auto;
//FDTransactionReadProcMSI.Options.AutoStart:=auto;
//FDTransactionReadProcMSI.Options.AutoStop:=auto;
end;
///////////////////////////////////////////////////////////////////////////

function CustomSortProc(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
  if ParamSort=0 then
    Result := CompareText(Item1.Caption,Item2.Caption)
  else
    if Item1.SubItems.Count>ParamSort-1 then
    begin
      if Item2.SubItems.Count>ParamSort-1 then
          case SortLV of
          true:Result := CompareText(Item1.SubItems[ParamSort-1],Item2.SubItems[ParamSort-1]);
         false:Result := CompareText(Item2.SubItems[ParamSort-1],Item1.SubItems[ParamSort-1]);
          end
      else
        Result := 1;
    end
    else
      Result:=-1;
end;

function SortThirdSubItemAsInt(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
   Result := 0;
   if strtoint( Item1.Caption ) > strtoint( Item2.Caption ) then
      Result := ParamSort
   else
   if strtoint( Item1.Caption ) < strtoint( Item2.Caption) then
      Result := -ParamSort;
end;

procedure TEditTask.FormClose(Sender: TObject; var Action: TCloseAction);
var
i:integer;
begin
PageControl1.TabIndex:=0; // переходим на первую вкладку для того чтобы не перерисовывать Listview2 при удалении столбцов
for I := ListView2.Columns.Count-1 downto 2 do   // удаляем столбцы созданые для результатов выполнения заданий , оставляем только 2 столбца № и имя компа
 begin
  ListView2.Columns.Delete(i);
 end;

end;

procedure TEditTask.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
i:integer;
YesPC,YesTask:boolean;
begin
YesPC:=false;
YesTask:=false;
for I := 0 to ListView2.Items.Count-1 do
begin
  if ListView2.Items[i].ImageIndex=4 then
  begin
  YesPC:=true; //есть сохраненные компы
  break;
  end;
end;
for I := 0 to ListView1.Items.Count-1 do
begin
  if ListView1.Items[i].ImageIndex=1 then
  begin
  Yestask:=true; //есть сохраненные задачи
  break;
  end;
end;
canclose:=true; // На всякий случай, вдруг все варианты не подберу.
if (not YesPC)and (not YesTask) then canclose:=true; // Если нет сохраненных задач и компов то закрываем
if (not YesPC)and (YesTask) then // если нет компов но есть сохраненные задачи
begin
ShowMessage('Список компьютеров пуст, добавьте компьютеры и сохраните задачу.');
canclose:=false;
end;
if (YesPC)and (not YesTask) then // если нет задач но есть комы
begin
ShowMessage('Список заданий пуст, добавьте задания и сохраните задачу.');
canclose:=false;
end;

if (YesPC)and (YesTask) then canclose:=true; //если есть сохраненные задачи и компы то заебись

end;

procedure TEditTask.FormGrClose(Sender: TObject; var Action: TCloseAction);
begin
if sender is Tform then
(sender as Tform).Release;
end;

procedure TEditTask.FormGRShow(Sender: TObject);
var
i:integer;
begin
ListGr.Clear;
if frmDomainInfo.ComboBox1.enabled then   // если comboBox не заблокирован
begin                                         // добавляем список групп в комбобокс для последующей работы с AD
for I := 0 to frmDomainInfo.ComboBox1.Items.Count-1 do
 ListGr.Items.Add(frmDomainInfo.ComboBox1.Items[i]);
ListGr.Text:=frmDomainInfo.ComboBox1.Text;
ListGr.OnSelect(ListGr);
end
else  /// иначе, используется или список IP или список компов,
begin /// или программа не зарегистрирована вот и добааляем компы которые есть в списке
   for I := 0 to frmDomainInfo.ListView8.Items.Count-1 do
   begin
     with LVPC.Items.Add do
      Caption:=frmDomainInfo.ListView8.Items[i].SubItems[0];
   end;
ListGr.Text:='Группы безопасности Active Directory доступны в зарегистрированной версии';
ListGr.Enabled:=false;
end;


end;

function statusubtask(nametable:string;CountSubTask:integer):TstringList; // передаем запрос и имя задания из задача, получаем статус этого задания
var
i,ok,no,warning,error,AllPC:integer;
res,MainRes:string;
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
begin
try
result:=TStringList.Create;
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
//TransactionRead.Options.DisconnectAction:=xdCommit; /// или  xdCommit
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRead.Options.AutoCommit:=false;
TransactionRead.Options.AutoStart:=false;
TransactionRead.Options.AutoStop:=false;
TransactionRead.Options.ReadOnly:=true;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=DataM.ConnectionDB;
TransactionRead.StartTransaction;
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT * FROM '+nameTable;
FDQueryRead.Open;
for i := 0 to CountSubTask-1 do
begin
  ok:=0;
  no:=0;
  warning:=0;
  error:=0;
  AllPC:=0;
  res:='';
  mainres:='';
  while not FDQueryRead.Eof do
    begin
    AllPC:=AllPC+1;
    res:=FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).AsString;
    if res='OK' then ok:=ok+1;
    if res='NO' then no:=no+1;
    if res='WARNING' then warning:=warning+1;
    if res='ERROR' then error:=error+1;
    FDQueryRead.Next;
    end;
  if (no<>0)and(ok<>0) then MainRes:='Выполняется';
  if no=0 then MainRes:='Выполнено';
  if ok=0 then MainRes:='Ожидает выполнения';
  if (no+ok+warning+error=AllPC) and (ok<>0) then MainRes:='Выполнено';

                       // сколько ОК   //сколько не выполнено+ошибки   // общее количество ошибок и предупреждений
  result.Add(MainRes+'/'+inttostr(ok)+'/'+inttostr(no+warning+error)+'/'+inttostr(error+warning));
  FDQueryRead.First; // переходим к первой строке для цикла по новому столбцу
end;


TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
except on E: Exception do
 begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
     Log_write('TASK',timetostr(now)+' Ошибка чтения результатов отдельных заданий - '+e.Message);
     frmdomaininfo.Memo1.Lines.Add('Ошибка чтения результатов отдельных заданий - '+e.Message);
  end;
end;

end;


function TEditTask.ReadResulTask(DescriptionTable:string):bool; /// передаем описание таблицы для чтения результатов выполнения
var
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
i,z,step:integer;
nameTable,str:string;
ResStatSubtask:tstringlist;
BEGIN
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
//TransactionRead.Options.DisconnectAction:=xdCommit; /// или  xdCommit
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRead.Options.AutoCommit:=false;
TransactionRead.Options.AutoStart:=false;
TransactionRead.Options.AutoStop:=false;
TransactionRead.Options.ReadOnly:=true;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=DataM.ConnectionDB;
TransactionRead.StartTransaction;
nameTable:=nametableForDescription(DescriptionTable);//тут получаем имя таблицы по опитанию
TaskName.Caption:=nameTable; /// за место описания таблицы записываем ее имя
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT * FROM '+nameTable;
FDQueryRead.Open;
step:=2;
ListView1.Clear;
ListView2.Clear;
step:=3;
ResStatSubtask:=TStringList.Create;
ResStatSubtask:=statusubtask(nameTable,FDQueryRead.FieldByName('TASK_QUANT').AsInteger);
 for I := 0 to FDQueryRead.FieldByName('TASK_QUANT').AsInteger-1 do   // цикл по столбцам таблицы
 begin
   with ListView1.Items.add do
     begin
     ImageIndex:=1;
     caption:=inttostr(FDQueryRead.FieldByName('NUMTASK'+inttostr(i)).AsInteger);
     Subitems.add(string(FDQueryRead.FieldByName('NAMETASK'+inttostr(i)).AsString));
     SubItems.Add(string(FDQueryRead.FieldByName('TYPETASK'+inttostr(i)).AsString));
     str:='';
     str:=ResStatSubtask[i];
     SubItems.Add(copy(str,1,pos('/',str)-1)); // статус задания
     System.delete(str,1,pos('/',str));        // удаляем до OK
     SubItems.Add(copy(str,1,pos('/',str)-1)); // на скольки компах выполнено удачно со статусом OK
     system.delete(str,1,pos('/',str));        // удаляем до NO
     SubItems.Add(copy(str,1,pos('/',str)-1)); // на скольки компах не выполнено (включая ERROR и WARNING)
     system.delete(str,1,pos('/',str));        // Удаляем до ERROR и WARNING
     SubItems.Add(str);                        // на скольки компах выполнено со статусом ERROR и  WARNING
     end;
   with ListView2.Columns.Add do // Добавляем столбцы для вывода результата выполнения в ListWiew с компами
   begin                         // потом при закрытии окна мы их удалим
     Caption:=FDQueryRead.FieldByName('NAMETASK'+inttostr(i)).AsString;
     Width:=350;
   end;
   with ListView2.Columns.Add do // Добавляем столбцы для вывода результата выполнения в ListWiew с компами
   begin                // потом при закрытии окна мы их удалим
     Caption:='Статус';
     Width:=70;
   end;

 end;
 step:=4;

while not FDQueryRead.Eof do /// добавляем имена компов //цикл по строкам таблицы
begin
  with ListView2.Items.Add do
  begin
    ImageIndex:=4;
    Caption:=inttostr(ListView2.Items.Count);
    SubItems.Add(FDQueryRead.FieldByName('PC_NAME').AsString);
    for I := 0 to FDQueryRead.FieldByName('TASK_QUANT').AsInteger-1 do // Добавляем результаты выполнения заданий
    begin
    if FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).Value<>null then
      begin
      if FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).AsString ='NO'then SubItems.Add('Ожидает выполнения')
      else
      SubItems.Add(FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).AsString);
      end
    else SubItems.Add('');
    if FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).Value<>null then
    SubItems.Add(FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).AsString)
    else SubItems.Add('');
    end;

  end;
  FDQueryRead.Next
end;
step:=5;


ResStatSubtask.Free;
TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
except on E: Exception do
  begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
     if Assigned(ResStatSubtask) then ResStatSubtask.Free;

      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
     Log_write('TASK',timetostr(now)+' Ошибка чтения результатов выполнения задачи - '+e.Message);
     frmdomaininfo.Memo1.Lines.Add('Ошибка чтения результатов выполнения задачи - '+e.Message);
  end;
end;
END;

procedure TEditTask.OnShow(Sender: TObject);
begin
PageControl1.TabIndex:=0;
Button7.Enabled:=false;
sortInt:=1;   // для сортировки списков
SortLV:=true; // для сортировки списков
end;

procedure TEditTask.ButOkClick(Sender: TObject);
var
groupList:tstringList;
i,z:integer;
PCinList:boolean;
begin
if LVPC.Items.Count>0 then
begin
  for I := 0 to LVPC.Items.Count-1 do
  Begin
  PCinList:=false;
  if LVPC.Items[i].Checked then
  begin
   for z := 0 to ListView2.Items.Count-1 do // будем искать есть ли этот комп в списке для задачи
    if LVPC.Items[i].Caption=ListView2.Items[z].SubItems[0] then // если комп уже в списке
    begin
    PCinList:=true;
    break;
    end;
    if not PCinList then
     with ListView2.Items.Add do
     begin
      ImageIndex:=0;
      Caption:=inttostr(ListView2.Items.Count);
      SubItems.add(LVPC.Items[i].Caption);
     end;
  end;

  End;
end;
end;

procedure TEditTask.ButCloseFormGRClick(Sender: TObject);
begin
FormGr.close;
end;

procedure TEditTask.LVPCColumnClick(Sender: TObject;
  Column: TListColumn);
  var
  i:integer;
begin
if LVPC.Columns[0].ImageIndex=1 then
  begin
  for I := 0 to LVPC.Items.Count-1 do
   LVPC.Items[i].Checked:=true;
   LVPC.Columns[0].ImageIndex:=2;
  exit;
  end;
if LVPC.Columns[0].ImageIndex=2 then
  begin
  for I := 0 to LVPC.Items.Count-1 do
   LVPC.Items[i].Checked:=false;
   LVPC.Columns[0].ImageIndex:=1;
  exit;
  end;
end;

procedure TEditTask.LVPCDblClick(Sender: TObject);
var
z,i:integer;
PCinList:bool;
begin
PCinList:=false;
if LVPC.SelCount=1 then
  begin
   for z := 0 to ListView2.Items.Count-1 do // будем искать есть ли этот комп в списке для задачи
    if LVPC.Selected.Caption=ListView2.Items[z].SubItems[0] then // если комп уже в списке
    begin
    PCinList:=true;
    break;
    end;
    if not PCinList then
     with ListView2.Items.Add do
     begin
      ImageIndex:=0;
      Caption:=inttostr(ListView2.Items.Count);
      SubItems.add(LVPC.Selected.Caption);
     end;
  end;
end;

procedure TEditTask.LVPCSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
if LVPC.SelCount>1 then
 if(item.Selected)and(item.Checked<>true) then item.Checked:=true;
end;


procedure TEditTask.ListGrSelect(Sender: TObject); /// выбор группы домена из списка
var
i:integer;
ListPC:Tstringlist;
begin
 LVPC.Clear;
 LVPC.Columns[0].ImageIndex:=1; // сбрасываем картинку если выделяли весь список
 ListPC:=TStringList.Create;
 try
 ListPC:=frmDomainInfo.GetAllGroupPC(ListGr.Text);
 for I := 0 to ListPC.Count-1 do
   begin
     with LVPC.Items.add do
     begin
       Caption:=ListPC[i];
     end;
   end;
 finally
   ListPC.free;
 end;
end;


procedure TEditTask.ListView1Changing(Sender: TObject; Item: TListItem;
  Change: TItemChange; var AllowChange: Boolean);
  var
  i:integer;
begin
try
Button7.Enabled:=true; // включить кнопку сохранить
if ListView1.Items.Count>30 then
begin
  for I := ListView1.Items.Count Downto 30 do // поциклу от крайнего до 30го итема
   begin
     ListView1.Items[i].Selected:=true; //выделяем
     Button4.Click;                    // удаляем
   end;
frmdomaininfo.Memo1.Lines.Add('Превышено максимальное количесто заданий в задаче.');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания, превышено максимальное количесто заданий в задаче');
end;
end;

procedure TEditTask.ListView1DblClick(Sender: TObject);
begin
if ListView1.SelCount=1 then  //   если выделен один элемент
if ListView1.ItemFocused.ImageIndex=1 then     //если это задание сохранено
CreateFormDetail(ListView1.ItemFocused.SubItems[0],false); // получаем отчет
end;

procedure TEditTask.ListView2Changing(Sender: TObject; Item: TListItem;
  Change: TItemChange; var AllowChange: Boolean);
begin
Button7.Enabled:=true; // включить кнопку сохранить
end;

procedure TEditTask.ListView2ColumnClick(Sender: TObject; Column: TListColumn);
begin

if Column.Index=1 then
begin
  SortLV:=not SortLV;
   ListView2.CustomSort(@CustomSortProc, Column.Index);
end;

if Column.Index=0 then
begin
  sortInt:=-sortInt;
  ListView2.CustomSort(@SortThirdSubItemAsInt, sortInt);
end;


end;




procedure TEditTask.ListView2DblClick(Sender: TObject);
begin
if listview2.SelCount=0 then exit;
if listview2.Selected.ImageIndex=4 then // если ImageIndex=4 то данные сохранены в таблицу и таблица существует то читаем ее
CreateFormDetail(EditTask.Caption,true);
end;

procedure TEditTask.Button11Click(Sender: TObject);
begin
creatformUserPass('');
end;

procedure TEditTask.Button15Click(Sender: TObject);  /// таймаут
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

WT:=2;
WiteTime:=InputBox('Таймаут ожидания до следующего задания', 'Время (мин):', '1');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('Вы указали не верное значение таймаута, продолжение операции не возможно!');
  exit;
end;
if WT=0 then
begin
  ShowMessage('Вы указали не верное значение таймаута, продолжение операции не возможно!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // таймаут ожидания
  SubItems.Add('Таймаут '+inttostr(WT)+' мин.');    // описание
  SubItems.Add('TimeOut');
end;
end;



function StrDeleteSourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
 if AnsiPos(':',br)=2 then // если в пути назначения найдено двоеточие то запрашиваем правильность ввода пути, для копирования в сеть после диска ставится $
 begin
 delete(br,2,1); // удаляем символ :
 insert('$',br,2); // вставляем на его место $
 end;
 result:=br;
end;


procedure TEditTask.ButDetailSaveClick(Sender: TObject);
begin
if LVPC.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(LVPC,'Сохранение списка результатов','');
end;

procedure TEditTask.FormDetailPCShow(Sender: TObject);
var
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
i,z:integer;
ConnectionThread: TFDConnection;
begin
try
try
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // расположение базы данных
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.LoginPrompt:= false;  /// отображение диалога user password
ConnectionThRead.Connected:=true;

TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThRead;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=ConnectionThRead;
for I := 0 to EditTask.ListView2.Items.Count-1 do
begin
if (EditTask.ListView2.Items[i].Selected) and(EditTask.ListView2.Items[i].ImageIndex=4) //если было сохранено то можно посмотреть подробности выполнения задач на коме
then
  begin
  TransactionWrite.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='select * from '+nametableForDescription(EditTask.Caption)+' where PC_NAME='''+EditTask.ListView2.Items[i].SubItems[0]+'''';
  FDQuery.Open;
  FDQuery.First;
  while not FDQuery.Eof do
    begin
       for z := 0 to FDQuery.FieldByName('TASK_QUANT').AsInteger-1 do
        with LVPC.Items.Add do
        begin
        Caption:=FDQuery.FieldByName('PC_NAME').AsString;
        SubItems.Add(FDQuery.FieldByName('NAMETASK'+inttostr(z)).AsString);
        SubItems.Add(FDQuery.FieldByName('RESULTTASK'+inttostr(z)).AsString);
        SubItems.Add(FDQuery.FieldByName('STATUSTASK'+inttostr(z)).AsString);
        end;
    FDQuery.Next;
    end;

  end;

end;
finally
TransactionWrite.Commit;
ConnectionThRead.Close;
FDQuery.Close;
FDQuery.Free;
TransactionWrite.Free;
ConnectionThRead.Free;
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка '+e.Message)
end;
end;

procedure TEditTask.FormDetailTaskShow(Sender: TObject);
var
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
ConnectionThread: TFDConnection;
begin
try
try
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // расположение базы данных
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.LoginPrompt:= false;  /// отображение диалога user password
ConnectionThRead.Connected:=true;
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=DataM.ConnectionDB;

  TransactionWrite.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='select * from '+nametableForDescription(EditTask.Caption);
  FDQuery.Open;
  FDQuery.First;
  if ListView1.SelCount=1 then
  while not FDQuery.Eof do
    begin
        with LVPC.Items.Add do
        begin
        Caption:=FDQuery.FieldByName('PC_NAME').AsString;
        SubItems.Add(FDQuery.FieldByName('NAMETASK'+inttostr(ListView1.Selected.Index)).AsString);
        SubItems.Add(FDQuery.FieldByName('RESULTTASK'+inttostr(ListView1.Selected.Index)).AsString);
        SubItems.Add(FDQuery.FieldByName('STATUSTASK'+inttostr(ListView1.Selected.Index)).AsString);
        end;
    FDQuery.Next;
    end;

finally
TransactionWrite.Commit;
ConnectionThRead.Close;
FDQuery.Close;
FDQuery.Free;
TransactionWrite.Free;
ConnectionThRead.Free;
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка '+e.Message)
end;
end;

Function TEditTask.CreateFormDetail(s:string;z:boolean):boolean;
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:='FormListGroup';
FormGr.Caption:='Подробно - '+s;
FormGr.Width:=800;
FormGr.Height:=400;
if z then FormGr.OnShow:=FormDetailPCShow
else FormGr.OnShow:=FormDetailTaskShow;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


PanelG:=Tpanel.Create(FormGr);
PanelG.Parent:=FormGr;
PanelG.Name:='PanelOK';
PanelG.Height:=35;
PanelG.Align:=alTop;
PanelG.Caption:='';

ButClose:=TButton.Create(FormGr);
ButClose.Parent:=PanelG;
ButClose.Name:='SaveList';
ButClose.Caption:='Сохранить список';
ButClose.Top:=5;
ButClose.Left:=5;
ButClose.Width:=100;
ButClose.OnClick:=ButDetailSaveClick;

LVPC:=Tlistview.Create(FormGr);
LVPC.Parent:=FormGr;
LVPC.Name:='LVRES';
LVPC.Align:=alClient;
LVPC.ViewStyle:=vsReport;
LVPC.ReadOnly:=true;
LVPC.RowSelect:=true;
LVPC.MultiSelect:=true;
LVPC.GridLines:=true;
lvpc.Checkboxes:=false;

  with LVPC.Columns.Add  do
  begin
    Caption:='Компьютер';
    Width:=100;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='Задание';
    Width:=300;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='Результат';
    Width:=300;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='Статус';
    Width:=50;
  end;
//LVPC.OnColumnClick:= LVPCColumnClick;
FormGr.Showmodal;
end;

procedure TEditTask.Button2Click(Sender: TObject);
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:='FormListGroup';
FormGr.Caption:='Список групп и компьютеров';
FormGr.Width:=500;
FormGr.Height:=300;
FormGr.BorderStyle:=bsDialog;
FormGr.OnShow:=FormGrShow;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;

PanelB:=Tpanel.Create(FormGr);
PanelB.Parent:=FormGr;
PanelB.Name:='PanelBUt';
PanelB.Width:=95;
panelB.Align:=alRight;
PanelB.Caption:='';
PanelB.TabOrder:=2;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=PanelB;
ButOk.Name:='AddPC';
ButOk.Caption:='Добавить';
ButOk.Top:=5;
ButOk.Left:=5;
ButOk.Width:=85;
ButOk.OnClick:=ButOkClick;
ButOk.TabOrder:=0;

ButClose:=TButton.Create(FormGr);
ButClose.Parent:=PanelB;
ButClose.Name:='ClFormGr';
ButClose.Caption:='Закрыть';
ButClose.Top:=35;
ButClose.Left:=5;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=1;

PanelG:=Tpanel.Create(FormGr);
PanelG.Parent:=FormGr;
PanelG.Name:='PanelCombo';
PanelG.Height:=30;
PanelG.Align:=alTop;
PanelG.Caption:='';
PanelG.TabOrder:=0;

ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=PanelG;
ListGr.Name:='ListGr';
ListGr.Text:='';
ListGr.Align:=alClient;
ListGr.DroppedDown:=true;
ListGr.DropDownCount:=15;
ListGr.OnSelect:=ListGrSelect;
ListGr.TabOrder:=0;


LVPC:=Tlistview.Create(FormGr);
LVPC.Parent:=FormGr;
LVPC.Name:='LVPC';
LVPC.Align:=alClient;
LVPC.ViewStyle:=vsReport;
LVPC.ReadOnly:=true;
LVPC.RowSelect:=true;
LVPC.MultiSelect:=true;
LVPC.GridLines:=true;
lvpc.Checkboxes:=true;
LVPC.SmallImages:=ImageList1;
with LVPC.Columns.Add  do
begin
  Caption:='Имя компьютера';
  Width:=250;
  ImageIndex:=1;
end;
LVPC.OnColumnClick:= LVPCColumnClick;
LVPC.OnDblClick:=LVPCDblClick;
LVPC.OnSelectItem:=LVPCSelectItem;
LVPC.TabOrder:=1;
FormGr.Showmodal;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TEditTask.N10Click(Sender: TObject);
begin
if listview1.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(listview1,'Сохранение списка заданий','');
end;

procedure TEditTask.N11Click(Sender: TObject); // Подробно о задании в задаче
begin
if ListView1.SelCount=1 then
if ListView1.Selected.ImageIndex<>0 then // если image<>0 (пауза) значит данное поле сохранено в таблицу и таблица сохранена и существует то читаем ее
CreateFormDetail(ListView1.ItemFocused.SubItems[0],false);
end;

procedure TEditTask.N12Click(Sender: TObject);
begin
if listview2.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(listview2,'Сохранение списка результатов','');
end;

procedure TEditTask.N13Click(Sender: TObject);
begin
if listview2.SelCount=0 then exit;
if listview2.Selected.ImageIndex=4 then // если ImageIndex=4 то данные сохранены в таблицу и таблица существует то читаем ее
CreateFormDetail(EditTask.Caption,true);
end;

procedure TEditTask.N14Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
 creatformService('Запустить службу');
end;

procedure TEditTask.N15Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
creatformService('Остановить службу');
end;

procedure TEditTask.N16Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
creatformService('Удалить службу');
end;

procedure TEditTask.N19Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
TaskCopyDelFF.Caption:='Копировать каталог';
TaskCopyDelFF.ShowModal;
end;

procedure TEditTask.N20Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
TaskCopyDelFF.Caption:='Копировать файл';
TaskCopyDelFF.ShowModal;
end;

procedure TEditTask.N21Click(Sender: TObject);
var
patch:string;
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Укажите каталог для удаления';
  Options:=[fdoPickFolders];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     patch:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
 if (patch='') then exit;
 if not InputQuery('Каталог для удаления ', 'Путь:', patch) then exit;
  if (patch='') then
  begin
   showmessage('Не указан каталог');
   exit;
  end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='Удалить :'+patch then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='0'; //    удаляем каталог
    SubItems.Add('Удалить :'+patch);    // путь
    SubItems.Add('DelFF');
  end;

end;


procedure TEditTask.N22Click(Sender: TObject);
var
patch:string;
z,i:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='Укажите файл для удаления';
  Options:=[fdoForceShowHidden];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     patch:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
 if (patch='') then exit;
 if not InputQuery('Файл для удаления ', 'Путь:', patch) then exit;
  if (patch='') then
  begin
   showmessage('Не указан файл');
   exit;
  end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='Удалить :'+patch then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='1'; //    удаляем файл
    SubItems.Add('Удалить :'+patch);    // путь
    SubItems.Add('DelFF');
  end;

end;



procedure TEditTask.N1Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
 creatformNewtask('msi'); // Создает форму для выббора MSI
end;

procedure TEditTask.N30Click(Sender: TObject); /// создать точку восстановления
var
NameTask,DesTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
try

DesTask:='Management Remote PC';
if not DescriptionForRestore(DesTask) then  //описание точки восстановления
begin
 ShowMessage('Вы не указали описание точки восстановления, операция прервана...');
 exit;
end;

NameTask:='Создать точку восстановления :'+DesTask;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(1); //1- создать точку восстановления, 2 - восстановить точку
  SubItems.Add(NameTask); // описание задания + описание точки восстановления
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;

function TEditTask.DateForTaskRestore(var DateRestore:string):boolean; // функция возвращает в переменную DateRestore введеную дату если она корректна
var
WaitName:boolean;
Rdate:TdateTime;
begin
try
DateRestore:=datetostr(now);
WaitName:=false;
while not WaitName do
Begin
if not InputQuery('Введите дату точки восстановления', 'Дата', DateRestore)
then
begin
DateRestore:='';
result:=false;
break;
end;
if (DateRestore='') then
  begin
  result:=false;
  break;
  end
else
 begin
  if TryStrToDate(DateRestore,Rdate) then
  begin
  result:=true;
  WaitName:=true;
  end
  else WaitName:=false;
 end;
End;
except
  begin
  ShowMessage('Ошибка формата даты');
  result:=false;
  end;
end;
end;

function TEditTask.DescriptionForRestore(var DesRestore:string):boolean;
var
WaitName:boolean;
begin
WaitName:=false;
while not WaitName do
BEGIN
if not InputQuery('Введите описание точки восстановления', 'Описание', DesRestore)
 then
  Begin
  result:=false;
  break;
  End
 else
  Begin
  DesRestore:=trim(DesRestore);
  if (DesRestore='') then
   begin
   WaitName:=false;
   end
  else
   begin
   WaitName:=true;
   result:= true;
   end;
  End;
END;
end;

procedure TEditTask.N31Click(Sender: TObject); // восстановить систему
var
NameTask:string;
Rdate:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
try
NameTask:='Восстановить систему';
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if not DateForTaskRestore(Rdate) then // дата для восстановления системы
begin
  ShowMessage('Вы не указали дату точки восстановления, операция прервана...');
  exit;
end;

i:=MessageDlg('Для восстановления системы потребуется перезагрузка компьютера.'+#10#13+'Продолжить выполнение операции?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
// сдесь открываем диалог для ввода даты, на подобии как ввод имени при создание новой задачи
// дату для определения точки восстановления добавляем к описанию
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(2); // 2 - восстановить точку определенную между датами, 1- создать точку восстановления,
  SubItems.Add(NameTask+' :'+Rdate); // описание задания
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;


procedure TEditTask.N32Click(Sender: TObject);
var
NameTask:string;
Rdate,DesPoint:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
try
NameTask:='Восстановить систему';
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;


// дату для определения точки восстановления добавляем к описанию
if not DateForTaskRestore(Rdate) then // дата для восстановления системы
begin
  ShowMessage('Вы не указали дату точки восстановления, операция прервана...');
  exit;
end;
// описание точки восстановления
DesPoint:='Management Remote PC';
if not DescriptionForRestore(DesPoint) then  //описание точки восстановления
begin
 ShowMessage('Вы не указали описание точки восстановления, операция прервана...');
 exit;
end;

i:=MessageDlg('Для восстановления системы потребуется перезагрузка компьютера.'+#10#13+'Продолжить выполнение операции?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(3); // 1- создать точку восстановления, 2 - восстановить точку по дате, 3- восстановление по дате и описаию
  SubItems.Add(NameTask+' :'+Rdate+' -> '+DesPoint); // описание задания
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;

procedure TEditTask.N3Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
creatformNewtask('proc');   // Создает форму для выббора процесса
end;



procedure TEditTask.N2Click(Sender: TObject); // открывает форму для новой программы MSI
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
NewMSiTask.ShowModal;
end;

procedure TEditTask.N4Click(Sender: TObject);  // открывает форму для нового процесса
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
NewProcTask.ShowModal;
end;


procedure TEditTask.N5Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
if SelectDelMSITask.Caption<>'Удаление программ msi' then
begin
 SelectDelMSITask.Edit1.Text:='';
 SelectDelMSITask.ComboBox1.clear;
end;
SelectDelMSITask.Caption:='Удаление программ msi';
SelectDelMSITask.Edit1.TextHint:='Имя или часть имени программы для удаления.';
SelectDelMSITask.ComboBox1.TextHint:='Загрузите список программ с компьютера в сети';
SelectDelMSITask.ShowModal;
end;

procedure TEditTask.N6Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

 if SelectDelMSITask.Caption<>'Завершение процесса' then
begin
 SelectDelMSITask.Edit1.Text:='';
 SelectDelMSITask.ComboBox1.clear;
end;
SelectDelMSITask.Caption:='Завершение процесса';
SelectDelMSITask.Edit1.TextHint:='Имя процесса для завершения notepad.exe.';
SelectDelMSITask.ComboBox1.TextHint:='Загрузите список процессов с компьютера в сети';
SelectDelMSITask.ShowModal;
end;

procedure TEditTask.N7Click(Sender: TObject);  // завершение сеанса
var
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='Завершение сеанса') then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:='0'; //
  SubItems.Add('Завершение сеанса пользователя');    // описание
  SubItems.Add('Logout');
end;
end;

procedure TEditTask.N8Click(Sender: TObject); // перезагрузка
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='Перезагрузка') then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

WT:=4;
WiteTime:=InputBox('Таймаут ожидания компьютера после перезагрузки', 'Время (мин):', '3');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('Вы указали не верное значение таймаута, продолжение операции не возможно!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // таймаут ожидания
  SubItems.Add('Перезагрузка компьютера');    // описание
  SubItems.Add('Reset');
end;
end;

procedure TEditTask.N9Click(Sender: TObject); // завершение работы
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='Завершение работы') then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
WT:=2;
WiteTime:=InputBox('Таймаут ожидания до выключения компьютера', 'Время (мин):', '1');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('Вы указали не верное значение таймаута, продолжение операции не возможно!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // таймаут ожидания до завершения работы
  SubItems.Add('Завершение работы');    // описание
  SubItems.Add('ShDown');
end;
end;

procedure TEditTask.WakeOnLan1Click(Sender: TObject); // включение компьютера
var
SetInI:TMeMiniFile;
BRaddress,WiteTime:string;
WT:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
BRaddress:=setini.readString('ConfLAN','broadcast','192.168.0.255');
IpBroadCast:=InputBox('Введите Broadcast адрес Вашей сети', 'IP-', BRaddress);
if IpBroadCast='' then
begin
  ShowMessage('Вы не указали Broadcast адрес сети, продолжение операции не возможно!');
  SetInI.Free;
  exit;
end;
SetInI.WriteString('ConfLAN','broadcast',IpBroadCast);
SetInI.UpdateFile;
SetInI.Free;
WT:=4;
WiteTime:=InputBox('Таймаут ожидания компьютера после включения', 'Время (мин):', '3');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('Вы указали не верное значение таймаута, продолжение операции не возможно!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // таймаут ожидания
  SubItems.Add('Включение компьютера. Broadcast :'+IpBroadCast);    // описание, из описания в потоке выделяется broadcast адрес
  SubItems.Add('WOL');
end;
end;



procedure TEditTask.Windows1Click(Sender: TObject);
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

StringKey:='';
try
   if not InputQuery('Формат XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', 'Ключ продукта:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('Вы не ввели ключ активации продукта');
   exit;
  end;

NameTask:='Активация Windows';
 if not InputQuery('Активация Windows', 'Описание задания:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('Вы не ввели описание задания');
   exit;
  end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask+' :'+StringKey then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(1); // активация windows
  SubItems.Add(NameTask+' :'+StringKey); // описание с ключем активации
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;

procedure TEditTask.Win881101Click(Sender: TObject);
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
StringKey:='';
try
   if not InputQuery('Формат XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', 'Ключ продукта:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('Вы не ввели ключ активации продукта');
   exit;
  end;

    NameTask:='Активация Office (Win 8/8.1/10...)';
 if not InputQuery('Активация Office', 'Описание задания:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('Вы не ввели описание задания');
   exit;
  end;

  for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=nametask+' :'+StringKey then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(2); // активация office для windows 8/8.1/10...
  SubItems.Add(nametask+' :'+StringKey); // описание с ключем активации
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;

procedure TEditTask.Windows71Click(Sender: TObject); // активация office для windows 7
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
StringKey:='';
try
   if not InputQuery('Формат XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', 'Ключ продукта:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('Вы не ввели ключ активации продукта');
   exit;
  end;

  NameTask:='Активация Office (Windows 7)';
 if not InputQuery('Активация Office', 'Описание задания:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('Вы не ввели описание задания');
   exit;
  end;

 for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask+' :'+StringKey then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(3); // активация office для windows 7
  SubItems.Add(NameTask+' :'+StringKey); // описание с ключем активации
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('Ошибка добавления задания - '+e.Message);
end;
end;


procedure TEditTask.FormMsiOpen(Sender: TObject); // добавляет список избранных msi программ
begin
try
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''msi''';
FDQueryProcMSI.Open;
ListGr.Clear;
ComboNumMisProc.Clear;
 while not FDQueryProcMSI.Eof do
 begin
 ListGr.Items.Add(string(FDQueryProcMSI.FieldByName('DESCRIPTION_PROC').Value));
 ComboNumMisProc.Items.Add(string(FDQueryProcMSI.FieldByName('ID_PROC').Value));
 FDQueryProcMSI.Next;
 end;
FDQueryProcMSI.Close;
if ListGr.Items.Count>0 then
  begin
  ListGr.ItemIndex:=0;
  ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
  end;
except on E: Exception do
begin
   Log_write('TASK',timetostr(now)+' Ошибка чтения списка программ- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' - Ошибка чтения списка программ- '+e.Message);
end;
end;
end;



procedure TEditTask.FormProcOpen(Sender: TObject);  // добавляет ссписок избранных процессов
begin
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''proc''';
FDQueryProcMSI.Open;
ListGr.Clear;
ComboNumMisProc.Clear;
 while not FDQueryProcMSI.Eof do
 begin
 ListGr.Items.Add(string(FDQueryProcMSI.FieldByName('DESCRIPTION_PROC').Value));
 ComboNumMisProc.Items.Add(string(FDQueryProcMSI.FieldByName('ID_PROC').Value));
 FDQueryProcMSI.Next;
 end;
FDQueryProcMSI.Close;
if ListGr.Items.Count>0 then
begin
 ListGr.ItemIndex:=0;
 ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
end;
end;

procedure TEditTask.AddTaskInListview(Sender: TObject);
var
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]=ListGr.Text) and (ListView1.Items[i].Caption=ComboNumMisProc.Text) then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=ComboNumMisProc.Text; // номер процесса или msi
  if FormGr.Caption='Избранные программы' then
  begin
   SubItems.Add('Установка программы '+ListGr.Text);    // описание
   SubItems.Add('msi');
  end;
  if FormGr.Caption='Избранные процессы' then
  begin
  SubItems.Add('Запуск процесса '+ListGr.Text);    // описание
  SubItems.Add('proc');
  end;
end;
end;

procedure TEditTask.AddServiceInListview(Sender: TObject);
var
i,z,step:integer;
begin
try
if ComboNumMisProc.Text='' then
begin
ShowMessage('Укажите имя службы');
ComboNumMisProc.SetFocus;
exit;
end;

step:=0;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=FormGr.Caption+' :'+ComboNumMisProc.Text then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
 step:=1;
with EditTask.ListView1.Items.Add do
begin
step:=2;
ImageIndex:=0;
  if FormGr.Caption='Запустить службу' then
  begin
   caption:=inttostr(1); // запуск службы
   SubItems.Add('Запустить службу :'+ComboNumMisProc.Text);    // имя службы
   SubItems.Add('Service');
  end;
  if FormGr.Caption='Остановить службу' then
  begin
  caption:=inttostr(2); // остановка службы
  SubItems.Add('Остановить службу :'+ComboNumMisProc.Text);    // имя службы
  SubItems.Add('Service');
  end;
  if FormGr.Caption='Удалить службу' then
  begin
  caption:=inttostr(3); // удаление службы
  SubItems.Add('Удалить службу :'+ComboNumMisProc.Text);    // имя службы
  SubItems.Add('Service');
  end;
end;
step:=3;
except on E: Exception do showmessage(E.Message);
end;
end;


procedure TEditTask.SelectComboMsiProc(Sender: TObject); // при выборе из списка процессов или msi выбирать его номер во втором combobox
begin
ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
end;

procedure TEditTask.SelectComboProcMSI(Sender: TObject); // при выборе из списка процессов или msi выбирать его номер во втором combobox
begin
ListGr.ItemIndex:=ComboNumMisProc.ItemIndex;
end;

procedure TEditTask.ListGrKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if (sender is Tcombobox) then
(sender as TCombobox).DroppedDown:=true;
end;

procedure TEditTask.ListGrKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key=13 then
begin
ButOk.Click;
end;
end;

function TEditTask.creatformNewtask(S:string):boolean; // создает форму для выбора избранных msi или процессов
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:=s+'Form';

FormGr.Width:=350;
FormGr.Height:=100;
FormGr.BorderStyle:=bsDialog;
if s='msi' then
begin
FormGr.OnShow:=FormMsiOpen;
FormGr.Caption:='Избранные программы';
end;
if s='proc' then
begin
 FormGr.OnShow:=FormProcOpen;
 FormGr.Caption:='Избранные процессы';
end;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:=s+'ComboBox';
ListGr.Text:='';
ListGr.top:=5;
ListGr.Left:=5;
ListGr.Width:=330;
ListGr.DropDownCount:=15;
Listgr.OnKeyDown:=ListGrKeyDown;
ListGr.OnSelect:= SelectComboMsiProc;
ListGr.OnKeyUp:=ListGrKeyUp;
ListGr.TabOrder:=0;

ComboNumMisProc:=TComboBox.Create(FormGr);
ComboNumMisProc.Parent:=FormGr;
ComboNumMisProc.Name:=s+'NumCombo';
ComboNumMisProc.Text:='';
ComboNumMisProc.top:=30;
ComboNumMisProc.Left:=5;
ComboNumMisProc.Width:=40;
ComboNumMisProc.OnSelect:=SelectComboProcMSI;
ComboNumMisProc.Style:=csOwnerDrawFixed;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:=s+'AddTask';
ButOk.Caption:='Добавить';
ButOk.Top:=40;
ButOk.Left:=160;
ButOk.Width:=85;
ButOk.OnClick:= AddTaskInListview;
ButOk.TabOrder:=1;


ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='Закрыть';
ButClose.Top:=40;
ButClose.Left:=250;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=2;
FormGr.Showmodal;
end;

procedure TEditTask.ButtonAddUser(Sender: TObject);
begin
AddUserPass(EditTask.Caption,EditUser.Text,EditPass.text,ChekSave.Checked);
FormGr.close;
end;

function TEditTask.creatformUserPass(S:string):boolean; // создает форму для ввода пользователя и пароля
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:=s+'FormUser';
FormGr.Width:=250;
FormGr.Height:=143;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:='Пользователь и пароль';
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


EditUser:=TEdit.Create(FormGr);
EditUser.Parent:=FormGr;
EditUser.Name:=s+'User';
EditUser.Text:='';
EditUser.top:=0;
EditUser.Left:=5;
EditUser.Width:=230;
EditUser.TextHint:='Имя пользователя';
EditUser.TabOrder:=0;

EditPass:=TEdit.Create(FormGr);
EditPass.Parent:=FormGr;
EditPass.Name:=s+'Passwd';
EditPass.Text:='';
EditPass.PasswordChar:=#7;
EditPass.TextHint:='Пароль';
EditPass.top:=25;
EditPass.Left:=5;
EditPass.Width:=230;
EditPass.TabOrder:=1;

ChekSave:=TCheckBox.Create(FormGr);
ChekSave.parent:=FormGr;
ChekSave.Name:=s+'SavePass';
ChekSave.Checked:=false;
ChekSave.Caption:='Сохранить после завершения задачи';
ChekSave.Top:=52;
ChekSave.Left:=5;
ChekSave.Width:=230;
ChekSave.TabOrder:=2;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:=s+'AddTask';
ButOk.Caption:='ОК';
ButOk.Top:=75;
ButOk.Left:=50;
ButOk.Width:=85;
ButOk.OnClick:= ButtonAddUser;
ButOk.TabOrder:=2;


ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='Закрыть';
ButClose.Top:=75;
ButClose.Left:=145;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=4;
FormGr.Showmodal;
end;

////////////////////////////////////////////////////////////////////////
procedure TEditTask.ButLoadService(Sender: TObject);
var
NamePSSearch:string;
begin
NamePSSearch:='localhost';
if not InputQuery('Загрузка списка может занять некоторое время!!!', 'Имя компьютера', NamePSSearch) then exit;
   if frmdomaininfo.ping(NamePSSearch) then
    LoadListService(NamePSSearch);
end;

function TEditTask.LoadListService(NamePC:string):boolean; /// загрузить список служб
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum           : IEnumvariant;
  iValue          : LongWord;

begin
OleInitialize(nil);
ListGr.Clear;
ComboNumMisProc.clear;
 frmDomainInfo.memo1.Lines.Add('Загрузка списка служб...');
 try
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,caption FROM Win32_Service','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// процессы
 begin
 if FWbemObject.Caption<>null then
 begin
 ListGr.Items.add(string(FWbemObject.Caption));       // описание службы
 ComboNumMisProc.Items.add(string(FWbemObject.name));       // имя службы
 end;
 FWbemObject:=Unassigned;
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 except
 on E:Exception do
 begin
 frmDomainInfo.memo1.Lines.Add('Ошибка загрузки служб "'+E.Message+'"');
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 end;
 end;
OleUnInitialize;
ListGr.text:=('Загрузка списка служб завершена.');
ComboNumMisProc.text:=('Загрузка списка служб  завершена.');
end;
//////////////////////////////////////////////////////////////////////
function TEditTask.creatformService(S:string):boolean; // создает форму для выбора служб  / перезаем имя задания -Запуск служы/ Остановка службы /Удаление службы
var
step:integer;
begin
try
step:=0;
FormGr:=TForm.Create(EditTask);
FormGr.Name:='Form';
FormGr.Width:=346;
FormGr.Height:=130;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:=s;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;

step:=1;
ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:='CaptionService';
ListGr.Text:='';
ListGr.top:=5;
ListGr.Left:=3;
ListGr.Width:=328;
ListGr.DropDownCount:=15;
Listgr.OnKeyDown:=ListGrKeyDown;
ListGr.OnSelect:= SelectComboMsiProc;
//ListGr.OnKeyUp:=ListGrKeyUp;
ListGr.TabOrder:=0;
ListGr.TextHint:='Описание службы';
step:=2;
ComboNumMisProc:=TComboBox.Create(FormGr);
ComboNumMisProc.Parent:=FormGr;
ComboNumMisProc.Name:='NameService';
ComboNumMisProc.Text:='';
ComboNumMisProc.DropDownCount:=15;
ComboNumMisProc.OnKeyDown:=ListGrKeyDown;
ComboNumMisProc.top:=33;
ComboNumMisProc.Left:=3;
ComboNumMisProc.Width:=238;
ComboNumMisProc.OnSelect:=SelectComboProcMSI;
ComboNumMisProc.TextHint:='Имя службы';
ListGr.TabOrder:=1;

step:=3;
ButLoad:=TButton.Create(FormGr);
ButLoad.Parent:=FormGr;
ButLoad.Name:='AddService';
ButLoad.Caption:='Загрузить';
ButLoad.Hint:='Загрузить список служб с компьютера в сети';
ButLoad.ShowHint:=true;
ButLoad.Top:=30;
ButLoad.Left:=248;
ButLoad.Width:=82;
ButLoad.OnClick:= ButLoadService;
ButLoad.TabOrder:=2;

step:=4;
ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='AddTask';
ButOk.Caption:='Добавить';
ButOk.Top:=65;
ButOk.Left:=248;
ButOk.Width:=82;
ButOk.OnClick:= AddServiceInListview;
ButOk.TabOrder:=3;

step:=5;
ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='Закрыть';
ButClose.Top:=65;
ButClose.Left:=148;
ButClose.Width:=82;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=4;
FormGr.Showmodal;
except on E: Exception do ShowMessage('Ошибка загрузки списка служб');
end;
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TEditTask.N24Click(Sender: TObject);   // удалить выделенные компы , вызов из popupmenu
begin
Button3.Click;
end;


procedure TEditTask.N27Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
NewsSubKeyName('Создать');
end;

procedure TEditTask.N28Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
NewsSubKeyName('Удалить');
end;

procedure TEditTask.N29Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
DeleteKeyName;
end;

procedure TEditTask.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;
procedure TEditTask.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;

procedure TEditTask.ButtonOksSubKeyName(Sender: TObject); // добавить задание (удалить или создать раздел)
var
i,z,x:integer;
des:string;
begin
if EditsSubKeyName.Text='' then begin showmessage('Не указан раздел'); exit; end;
if EditsSubKeyName.Text[Length(EditsSubKeyName.Text)]='\' then // если в конце строки наден символ '/' то удаляем его
begin
 EditsSubKeyName.Text:=copy(EditsSubKeyName.Text,1,Length(EditsSubKeyName.Text)-1);
end;

if pos('Создать',FormKey.Caption)<>0 then begin x:=1; Des:='Создать'; end;
if pos('Удалить',FormKey.Caption)<>0 then begin x:=0; Des:='Удалить'; end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=Des+' раздел :'+ComboRootKey.Text+':'+EditsSubKeyName.Text then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if length(Des+' раздел :'+ComboRootKey.Text+':'+EditsSubKeyName.Text)>250 then
 begin
   ShowMessage('Длинна команды превышает допустимую величину');
   exit;
 end;
with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:=inttostr(x); // 1 - создать раздел 0 - удалить раздел
    SubItems.Add(Des+' раздел :'+ComboRootKey.Text+':'+EditsSubKeyName.Text);    // описание
    SubItems.Add('sSubKey'); //работа с разделом реестра (создание или удаление)
  end;
FormKey.close;
end;

procedure TEditTask.ButtonOkDelKey(Sender: TObject); // добавить задание удаление ключа
var
i,z,x:integer;
begin
if EditsSubKeyName.Text='' then begin showmessage('Не указан раздел'); exit; end;
if EditNameKey.Text='' then begin showmessage('Не указано имя ключа'); exit; end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='Удалить ключ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if length('Удалить ключ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text)>250 then
 begin
   ShowMessage('Длинна команды превышает допустимую величину');
   exit;
 end;

with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='0'; // значение не влияет на выполнение операции
    SubItems.Add('Удалить ключ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text);    // описание
    SubItems.Add('sNameKey'); //работа с разделом реестра (удаление ключа реестра)
  end;
FormKey.close;
end;


function TEditTask.NewsSubKeyName(CreateDel:string):boolean; // создать или удалить раздел реестра
begin
try
FormKey:=TForm.Create(EditTask);
if CreateDel='Создать' then FormKey.Caption:='Создать раздел реестра';
if CreateDel='Удалить' then FormKey.Caption:='Удалить раздел реестра';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=140;
FormKey.OnClose:=FormKeyClose;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=5;
ComboRootKey.Left:=5;
ComboRootKey.Width:=260;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='Раздел (пример - HARDWARE\MRPC):';
EditsSubKeyName.Top:=45;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=72;
ButtonOk.Left:=110;
ButtonOk.OnClick:=ButtonOksSubKeyName;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=72;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
except on E: Exception do
begin
  Log_write('TASK',timetostr(now)+' Ошибка создания задания (создания/удаления раздела реестра) - '+e.Message);
  frmdomaininfo.Memo1.Lines.Add('Ошибка создания задания (создания/удаления раздела реестра) - '+e.Message);
end;
end;
end;

procedure TEditTask.SZ1Click(Sender: TObject); /// открываем форму сохраненных ключей реестра
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('Невозможно добавить больше 30ти заданий');
 exit;
end;
RegKeySave.Button4.Enabled:=true; // включаем кнопку добавления в задачу
RegKeySave.Button1.Enabled:=false;  // отключаем кнопку импорта в реестр
RegKeySave.PopupEditKey.Items[1].Enabled:=false; // отключаем пункты меню для импорта в реестр
RegKeySave.ShowModal;  // открываем форму
end;


function TEditTask.DeleteKeyName:boolean;
begin
try
FormKey:=TForm.Create(EditTask);
FormKey.Caption:='Удалить ключ реестра';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=180;
FormKey.OnClose:=FormKeyClose;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=5;
ComboRootKey.Left:=5;
ComboRootKey.Width:=260;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='Раздел (пример - HARDWARE\MRPC):';
EditsSubKeyName.Top:=45;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Удаляемый параметр:';
EditNameKey.Top:=83;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:='Имя ключа реестра';
EditNameKey.TabOrder:=1;
////////////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=110;
ButtonOk.Left:=110;
ButtonOk.OnClick:=ButtonOkDelKey;
ButtonOk.TabOrder:=2;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=110;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=3;
FormKey.ShowModal;
except on E: Exception do
begin
  Log_write('TASK',timetostr(now)+' Ошибка создания задания (создания/удаления раздела реестра) - '+e.Message);
  frmdomaininfo.Memo1.Lines.Add('Ошибка создания задания (создания/удаления раздела реестра) - '+e.Message);
end;
end;
end;

procedure TEditTask.Button3Click(Sender: TObject); /// удалить выделенные компьютеры
var
i,step:integer;
NameTabl:string;
TransactionDelPC:TFDTransaction;
FDQueryDelPC:TFDQuery;
begin
try
step:=0;
if ListView2.SelCount=0 then exit;
TransactionDelPC:= TFDTransaction.Create(nil);
TransactionDelPC.Connection:=DataM.ConnectionDB;
TransactionDelPC.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionDelPC.Options.AutoCommit:=false;
TransactionDelPC.Options.AutoStart:=false;
TransactionDelPC.Options.AutoStop:=false;
FDQueryDelPC:= TFDQuery.Create(nil);
FDQueryDelPC.Transaction:=TransactionDelPC;
FDQueryDelPC.Connection:=DataM.ConnectionDB;
TransactionDelPC.StartTransaction; // стартуем транзакцию
NameTabl:=nametableForDescription(EditTask.Caption);
for I := ListView2.Items.Count-1 downto 0 do
begin
step:=1;
  if (ListView2.Items[i].Selected) and (ListView2.Items[i].ImageIndex=4) then
  begin
  FDQueryDelPC.SQL.Clear;
  FDQueryDelPC.SQL.Text:='delete FROM '+NameTabl+' WHERE PC_NAME='''+ListView2.Items[i].SubItems[0]+'''';
  FDQueryDelPC.ExecSQL;
  end;
  step:=2;
  if (ListView2.Items[i].Selected) then ListView2.Items[i].Delete;
end;
step:=3;
TransactionDelPC.Commit;  // сохраняем то что внесли в базу
FDQueryDelPC.Close;
FDQueryDelPC.Free;
TransactionDelPC.Free
except on E: Exception do
begin
   if Assigned(FDQueryDelPC) then FDQueryDelPC.Free;
      if Assigned(TransactionDelPC) then
      begin
      TransactionDelPC.Rollback;
      TransactionDelPC.Free;
      end;
   Log_write('TASK',timetostr(now)+' шаг - '+inttostr(step)+' - Ошибка удаления компьютера- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' шаг - '+inttostr(step)+' - Ошибка удаления компьютера- '+e.Message);
end;
end;
end;


function TEditTask.nametableForDescription(Descript:string):string; /// Имя таблицы по описанию
begin
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM TABLE_TASK WHERE DESCRIPTION_TASK='''+Descript+'''';
FDQueryProcMSI.Open;
result:= FDQueryProcMSI.FieldByName('NAME_TABLE').AsString; // тут получаем имя таблицы по опитанию
FDQueryProcMSI.SQL.clear;
end;

procedure TEditTask.N23Click(Sender: TObject); // удалить задание из popup menu
begin
Button4.Click;
end;

procedure TEditTask.Button4Click(Sender: TObject); /// удалить задание
var
i,z:integer;
step:integer;
nameT:string;
TransactionTaskDel:TFDTransaction;
FDQueryDelTask:TFDQuery;
begin
try
if StatusStartStopTask(EditTask.Caption) then // если задача запущена (true)
 begin
   i:=MessageDlg('Не рекомендуется вносить изменения в запущенную задачу. Продолжить?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit;
 end;
step:=0;
if ListView1.SelCount=0 then exit; /// если не выделили запись то выходим
 step:=1;
nameT:=nametableForDescription(EditTask.Caption);  // имя таблицы
TransactionTaskDel:= TFDTransaction.Create(nil);
TransactionTaskDel.Connection:=DataM.ConnectionDB;
TransactionTaskDel.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionTaskDel.Options.AutoCommit:=false;
TransactionTaskDel.Options.AutoStart:=false;
TransactionTaskDel.Options.AutoStop:=false;
FDQueryDelTask:= TFDQuery.Create(nil);
FDQueryDelTask.Transaction:=TransactionTaskDel;
FDQueryDelTask.Connection:=DataM.ConnectionDB;
TransactionTaskDel.StartTransaction; // стартуем транзакцию
for I := ListView1.Items.Count-1 downto 0 do
begin
 step:=2;
  if (ListView1.Items[i].Selected) and (ListView1.Items[i].ImageIndex=1) then // если удаляем задание уже сохраненное в таблице
  begin
   step:=3;
   FDQueryDelTask.SQL.clear;
   FDQueryDelTask.SQL.Text:='ALTER TABLE '+nameT+' DROP  TypeTask'+inttostr(i)
   +',DROP NumTask'+inttostr(i)+',DROP ResultTask'+inttostr(i)+',DROP NameTask'+inttostr(i)+',DROP StatusTask'+inttostr(i);
   FDQueryDelTask.ExecSQL;
   ListView1.Items[i].Delete;
   step:=4;
   for z := i+1 to ListView1.Items.Count do // переименовываем столбцы от удаленного до последнего
   begin
   step:=5;
     FDQueryDelTask.SQL.clear;                            /// текущий столбец переименовываем в предыдущий
     FDQueryDelTask.SQL.Text:='ALTER TABLE '+nameT
     +' ALTER TypeTask'+inttostr(z)+' TO TypeTask'+inttostr(z-1)
     +',ALTER NumTask'+inttostr(z)+' TO NumTask'+inttostr(z-1)
     +',ALTER ResultTask'+inttostr(z)+' TO ResultTask'+inttostr(z-1)
     +',ALTER NameTask'+inttostr(z)+' TO NameTask'+inttostr(z-1)
     +',ALTER StatusTask'+inttostr(z)+' TO StatusTask'+inttostr(z-1);
     FDQueryDelTask.ExecSQL;
   step:=6;
   end;
  FDQueryDelTask.SQL.clear;                            /// изменяем количество задач
  FDQueryDelTask.SQL.Text:='UPDATE '+nameT+' SET TASK_QUANT='+inttostr(ListView1.Items.Count);
  FDQueryDelTask.ExecSQL;
  end
 else
  if (ListView1.Items[i].Selected) and (ListView1.Items[i].ImageIndex=0) then  // если выделенно и еще не сохранено в таблицу то просто удаляем строку
   ListView1.Items[i].Delete;


end;
step:=7;
TransactionTaskDel.Commit;  // сохраняем то что внесли в базу
FDQueryDelTask.Close;
FDQueryDelTask.Free;
TransactionTaskDel.Free;
except on E: Exception do
  begin
  if Assigned(FDQueryDelTask) then FDQueryDelTask.Free;
      if Assigned(TransactionTaskDel) then
      begin
      TransactionTaskDel.Rollback;
      TransactionTaskDel.Free;
      end;
  Log_write('TASK',timetostr(now)+' шаг - '+inttostr(step)+' - Ошибка удаления задания из задачи- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' шаг - '+inttostr(step)+' - Ошибка удаления задания из задачи- '+e.Message);
  end;
end;
end;


function TEditTask.RenameTableForTask(NameTable,NewDescriptionTable:string):boolean; //Меняем описание таблицы для задачи. обновлят запись в таблице TABLE_TASK
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // стартуем
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET DESCRIPTION_TASK=:a  WHERE NAME_TABLE='''+NameTable+'''';
step:=2;
FDQueryWriteTaskTable.ParamByName('a').asstring:=NewDescriptionTable;
step:=3;
FDQueryWriteTaskTable.ExecSQL;
step:=4;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.Close;
TransactionWriteTaskTable.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryWriteTaskTable.SQL.clear;   //очистить
FDQueryWriteTaskTable.Close;  /// закрыть нах
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
result:=true;
Except
on E:Exception do
     begin
     if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:=false;
     Log_write('TASK',timetostr(now)+' шаг - '+inttostr(step)+' - Ошибка переименования задачи- '+e.Message);
     end;
end;
end;




function TEditTask.UpdateTableTABLETASK(NameTable:string;countPC:integer):boolean; // обновлят запись в таблице TABLE_TASK
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // стартуем

FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='update or insert into TABLE_TASK (NAME_TABLE,COUNT_PC) VALUES (:p1,:p2)  MATCHING (NAME_TABLE)';
step:=2;
FDQueryWriteTaskTable.params.ParamByName('p1').AsString:=''+NameTable+'';
FDQueryWriteTaskTable.params.ParamByName('p2').Asinteger:=countPC; // всего компьютеров
step:=3;
FDQueryWriteTaskTable.ExecSQL;
step:=4;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.Close;
result:=true;
TransactionWriteTaskTable.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryWriteTaskTable.SQL.clear;   //очистить
FDQueryWriteTaskTable.Close;  /// закрыть нах
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
Except
on E:Exception do
     begin
     if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:=false;
     Log_write('TASK',timetostr(now)+' шаг - '+inttostr(step)+' - Ошибка Update TABLE_TASK- '+e.Message);
     end;
end;
end;

function TEditTask.ThereAnyRunStopTask(statustask:boolean):integer; /// читам статус задачи, b смотрим есть ли запущенные или остановленные
var
FDQueryReadStatusTask:TFDQuery;
TransactionReadStatusTask: TFDTransaction;
begin
try
TransactionReadStatusTask:= TFDTransaction.Create(nil);
TransactionReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionReadStatusTask.Options.AutoCommit:=false;
TransactionReadStatusTask.Options.AutoStart:=false;
TransactionReadStatusTask.Options.AutoStop:=false;
TransactionReadStatusTask.Options.ReadOnly:=true;
FDQueryReadStatusTask:= TFDQuery.Create(nil);
FDQueryReadStatusTask.Transaction:=TransactionReadStatusTask;
FDQueryReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.StartTransaction; // стартуем
FDQueryReadStatusTask.SQL.Clear;
FDQueryReadStatusTask.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE START_STOP='''+booltostr(statustask)+'''';
FDQueryReadStatusTask.open;
FDQueryReadStatusTask.Last;
result:=FDQueryReadStatusTask.RecordCount; //результат отправляем количество записей со статусом  statustask
TransactionReadStatusTask.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryReadStatusTask.SQL.clear;   //очистить
FDQueryReadStatusTask.Close;  /// закрыть нах
FDQueryReadStatusTask.Free;
TransactionReadStatusTask.Free;
 except on E: Exception do
     begin
     result:=0;
     if Assigned(FDQueryReadStatusTask) then FDQueryReadStatusTask.Free;
      if Assigned(TransactionReadStatusTask) then
      begin
      TransactionReadStatusTask.Rollback;
      TransactionReadStatusTask.Free;
      end;
     Log_write('TASK',timetostr(now)+' - Ошибка READ StartStopTask TABLE_TASK :' +e.Message);
     end;
   end;
end;

function TEditTask.StatusStartStopTask(DescriptionTask:string):boolean; /// читам статус задачи, остановить или продолжать выполнение
var                                                           //// если задача остановлена то результат присваиваем false
FDQueryReadStatusTask:TFDQuery;
TransactionReadStatusTask: TFDTransaction;
begin
TransactionReadStatusTask:= TFDTransaction.Create(nil);
TransactionReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionReadStatusTask.Options.AutoCommit:=false;
TransactionReadStatusTask.Options.AutoStart:=false;
TransactionReadStatusTask.Options.AutoStop:=false;
TransactionReadStatusTask.Options.ReadOnly:=true;
FDQueryReadStatusTask:= TFDQuery.Create(nil);
FDQueryReadStatusTask.Transaction:=TransactionReadStatusTask;
FDQueryReadStatusTask.Connection:=DataM.ConnectionDB;
try
try
TransactionReadStatusTask.StartTransaction; // стартуем
FDQueryReadStatusTask.SQL.Clear;
FDQueryReadStatusTask.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE DESCRIPTION_TASK='''+DescriptionTask+'''';
FDQueryReadStatusTask.open;
result:=strtobool(FDQueryReadStatusTask.FieldByName('START_STOP').AsString);
//frmDomainInfo.Memo1.Lines.Add('Значение из функции - '+FDQueryReadStatusTask.FieldByName('START_STOP').AsString);
TransactionReadStatusTask.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryReadStatusTask.SQL.clear;   //очистить
FDQueryReadStatusTask.Close;  /// закрыть нах
 except on E: Exception do
     begin
      TransactionReadStatusTask.Rollback;
      FDQueryReadStatusTask.Close;
      Log_write('TASK',timetostr(now)+' - Ошибка READ StartStopTask TABLE_TASK :' +e.Message);
      end;
     end;
finally
FDQueryReadStatusTask.Free;
TransactionReadStatusTask.Free;
end;
end;

function TEditTask.AddNewItemsTable(DescriptionTable:string;countPC:integer):string; // Записывает описание таблицы в таблицу имен заданий и возвращает имя новой таблицы
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // стартуем
FDQueryWriteTaskTable.SQL.Text:='SELECT GEN_ID(GEN_FOR_TABLE_TASK,0) FROM RDB$DATABASE'; // текущее значение генератора, без увеличения на 1
FDQueryWriteTaskTable.Open;
counttable:=FDQueryWriteTaskTable.Fields[0].AsInteger+1; ///вот и значение, если текущее значение 0 то при добавлении строки буте 1, соответсвенно добавляем еденицу для соответствия имени таблицы и ID генератора
FDQueryWriteTaskTable.Close;
TransactionWriteTaskTable.Commit;
TransactionWriteTaskTable.StartTransaction;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='update or insert into TABLE_TASK (NAME_TABLE,DESCRIPTION_TASK,STATUS_TASK,START_STOP,COUNT_PC,CURRENT_TASK,PC_RUN,WORKS_THREAD,USER_NAME,PASS_USER,SAVE_PASS) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11)  MATCHING (NAME_TABLE)';
FDQueryWriteTaskTable.params.ParamByName('p1').AsString:=''+'TASK_'+inttostr(counttable)+'';
FDQueryWriteTaskTable.params.ParamByName('p2').AsString:=''+DescriptionTable+'';
FDQueryWriteTaskTable.params.ParamByName('p3').AsString:='Остановлена'; // статус запуска задачи
FDQueryWriteTaskTable.params.ParamByName('p4').AsString:=booltostr(false); /// признак того что задача должна быть остановлена остановлена/или разрешен запуск
FDQueryWriteTaskTable.params.ParamByName('p5').Asinteger:=countPC; // всего компьютеров
FDQueryWriteTaskTable.params.ParamByName('p6').AsString:='';     // Текщее задание
FDQueryWriteTaskTable.params.ParamByName('p7').AsString:='';     // на каком компе выполняется
FDQueryWriteTaskTable.params.ParamByName('p8').AsString:=booltostr(false); // признак того что Поток остановил выполнение задачи после чтения значения из START_STOP
FDQueryWriteTaskTable.params.ParamByName('p9').AsString:='';   // пользователь
FDQueryWriteTaskTable.params.ParamByName('p10').AsString:='';  // пароль
FDQueryWriteTaskTable.params.ParamByName('p11').AsBoolean:=false; // Не сохранять пароль после выполнения задачи
FDQueryWriteTaskTable.ExecSQL;
TransactionWriteTaskTable.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit
FDQueryWriteTaskTable.SQL.clear;   //очистить
FDQueryWriteTaskTable.Close;  /// закрыть нах
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
result:='TASK_'+inttostr(counttable); // в результате имя таблицы
Except
on E:Exception do
     begin
      if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:='';
     Log_write('TASK',timetostr(now)+' шаг - '+inttostr(step)+' - Ошибка записи имени таблицы- '+e.Message);
     end;
end;
end;

////////////////////////////////////////////////////////////////////
function TeditTask.AddOrUpdateListViewreadInTableTask(nameTable:string):boolean;
var
i:integer;
update:boolean;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin  // чтение записи в таблице Table_task и добавление информации о ней в гланый ListView
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
TransactionWriteTaskTable.Options.ReadOnly:=true;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // стартуем

update:=false;
FDQueryWriteTaskTable.SQL.clear;
FDQueryWriteTaskTable.SQL.Text:='SELECT * FROM TABLE_TASK WHERE NAME_TABLE='''+nameTable+'''';
FDQueryWriteTaskTable.Open;
for I := 0 to frmdomaininfo.TaskListView.Items.Count-1 do // ищем есть ли такая задача в списке чтобы ее обновить или добавить
begin
 if frmdomaininfo.TaskListView.Items[i].Caption=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString then
  begin
  frmdomaininfo.TaskListView.Items[i].Caption:=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString; // описание задачи
  frmdomaininfo.TaskListView.Items[i].SubItems[0]:=FDQueryWriteTaskTable.FieldByName('STATUS_TASK').AsString; // статус задачи (выполняется/остановлена/остановка)
  frmdomaininfo.TaskListView.Items[i].SubItems[1]:=FDQueryWriteTaskTable.FieldByName('CURRENT_TASK').AsString;     // Текущая выполняема задача
  frmdomaininfo.TaskListView.Items[i].SubItems[2]:=FDQueryWriteTaskTable.FieldByName('PC_RUN').AsString;      // на каком компьютере сейчас выполняется
  frmdomaininfo.TaskListView.Items[i].SubItems[3]:=FDQueryWriteTaskTable.FieldByName('COUNT_PC').AsString;    // всего компьютеров в задаче
  update:=true;
  end;
end;
if not update then /// если такой задачи нет значит ее надо добавить в список
with frmdomaininfo.TaskListView.Items.Add do
begin
caption:=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString; // описание задачи
SubItems.add(FDQueryWriteTaskTable.FieldByName('STATUS_TASK').AsString); // статус задачи (выполняется/остановлена/остановка)
SubItems.add(FDQueryWriteTaskTable.FieldByName('CURRENT_TASK').AsString);     // Текущая выполняема задача
SubItems.add(FDQueryWriteTaskTable.FieldByName('PC_RUN').AsString);      // на каком компьютере сейчас выполняется
SubItems.add(FDQueryWriteTaskTable.FieldByName('COUNT_PC').AsString);    // всего компьютеров в задаче
end;
TransactionWriteTaskTable.Commit;  /// сохранить изменения , даже при чтении надо заверщать commit

FDQueryWriteTaskTable.SQL.clear;   //очистить
FDQueryWriteTaskTable.Close;  /// закрыть нах
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;

result:=true;
except on E: Exception do
begin
 if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
 if Assigned(TransactionWriteTaskTable) then
 begin
 TransactionWriteTaskTable.Rollback;
 TransactionWriteTaskTable.Free;
 end;
frmdomaininfo.memo1.Lines.add('Ошибка обновления ListView :'+E.Message);
result:=false;
end;
end;
end;
////////////////////////////////////////////////////////////////////


function TEditTask.createNewTableTask(TableName:string):boolean;     //функция запуска создания новой таблицы+ее заполнение
var
i,z,step,p:integer;
ColName,ParamNum:string;
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
function incP(z:integer):integer;
begin
p:=z+1;
result:=p;
end;
begin
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
step:=0;
dataM.СreateTablForTASK(TableName,'','','','',booltostr(false)); //создаем таблицу //создаем таблицу
step:=1;
for I := 0 to ListView1.Items.Count-1 do /// создаем столбцы для задач
begin
dataM.СreateTablForTASK(TableName,'TypeTask'+inttostr(i),'NumTask'+inttostr(i),'ResultTask'+inttostr(i),'NameTask'+inttostr(i),'StatusTask'+inttostr(i)) ;
end;
step:=2;
ColName:='';
ParamNum:='';
p:=3;
for i := 0 to ListView1.items.Count-1 do  // заполняем строковые переменные данными для sql запроса
Begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
End;
step:=3;
FDQueryAdd.SQL.clear;
FDQueryAdd.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME);';
 step:=4;
 FDQueryAdd.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - общее количество задач
 step:=5;
p:=3;
for i := 0 to ListView1.items.Count-1 do //цикл по задачам заполняем один раз так как эти параметры для каждой строки одинаковые
begin
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - тип задачи msi, proc, питание
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - номер выполняемой задачи
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // результат выполнения задачи  ResultTask
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // Описание задания NameTask
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';      // Текущее состояние задачи  StatusTask
end;
step:=6;
for i := 0 to ListView2.items.Count-1 do /// цикл по компам заполняем и выполняем sql запрос т.к. строки будут отличаться только именем компа и ID
begin
  FDQueryAdd.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryAdd.params.ParamByName('p3').AsString:=ListView2.items[i].subitems[0];           //PC_NAME
  step:=7;
  FDQueryAdd.ExecSQL;
end;
step:=8;
///////////////////////////////////////////////////////////////////
TransactionAdd.Commit;
FDQueryAdd.Close;
FDQueryAdd.Free;
TransactionAdd.Free;
result:=true;
Except
   on E: Exception do
      begin
      if Assigned(TransactionAdd) then
      begin
      TransactionAdd.Rollback;
      TransactionAdd.Free;
      end;
       if Assigned(FDQueryAdd) then
      begin
       FDQueryAdd.Free;
      end;
      frmdomaininfo.memo1.Lines.add('Ошибка создания новой задачи (шаг '+inttostr(step)+') :'+E.Message);
      result:=false;
      end;
   end;

end;

function TEditTask.UpdateTableTask(TableName:string):boolean;     //функция обновления таблицы+ее заполнение
var
i,z,step,p,NewZorPC:integer;
ColName,ParamNum:string;
FDQueryUpdate:TFDQuery;
TransactionUpdate: TFDTransaction;
function incP(z:integer):integer;
begin
p:=z+1;
result:=p;
end;
begin
try
step:=0;
step:=1;
for I := 0 to ListView1.Items.Count-1 do /// создаем столбцы для задач
begin
if ListView1.Items[i].ImageIndex=0 then // если новые задачи были добавлены то создаем новые столбцы
dataM.СreateTablForTASK(TableName,'TypeTask'+inttostr(i),'NumTask'+inttostr(i),'ResultTask'+inttostr(i),'NAMETASK'+inttostr(i),'StatusTASK'+inttostr(i)) ;
end;
step:=2;
ColName:=''; /// сдесь будет строка с именами колонок
ParamNum:=''; // сдесь будет строка с параметрами (:p1....:p.)
NewZorPC:=0;  // подсчет новых заданий или новых компов
////////////////////////////////////////////////////////заполнение таблицы новыми задачами для  (тех кто был в задаче) компов
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
for i := 0 to ListView1.items.Count-1 do /// Проверяем в цикле были ли добавлены новые задачи
if ListView1.Items[i].ImageIndex=0 then NewZorPC:=NewZorPC+1;

if NewZorPC<>0 then // если NewZorPC больше 0 то задачи были добавлены
Begin
TransactionUpdate:= TFDTransaction.Create(nil);
TransactionUpdate.Connection:=DataM.ConnectionDB;
TransactionUpdate.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;  xiUnspecified
TransactionUpdate.Options.AutoCommit:=false;
TransactionUpdate.Options.AutoStart:=false;
TransactionUpdate.Options.AutoStop:=false;
FDQueryUpdate:= TFDQuery.Create(nil);
FDQueryUpdate.Transaction:=TransactionUpdate;
FDQueryUpdate.Connection:=DataM.ConnectionDB;
p:=3;
for i := 0 to ListView1.items.Count-1 do  // заполняем строковые переменные данными для sql запроса
begin
 if ListView1.Items[i].ImageIndex=0 then // если это новая задача
 begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
 end;
end;
step:=3;
FDQueryUpdate.SQL.clear;
FDQueryUpdate.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME)';;
 step:=4;
FDQueryUpdate.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - общее количество задач
 step:=5;
p:=3;
for i := 0 to ListView1.items.Count-1 do //цикл по задачам заполняем один раз так как эти параметры для каждой строки одинаковые
begin
   if ListView1.Items[i].ImageIndex=0 then  // если это новая задача
   begin
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - тип задачи msi, proc, питание
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - номер выполняемой задачи
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // результат выполнения задачи
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // Описание задания
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO'; // текущее состояние задачи  //statusTask
  end;
end;
step:=6;
for i := 0 to ListView2.items.Count-1 do /// цикл по компам заполняем и выполняем sql запрос т.к. строки будут отличаться только именем компа и ID
begin
if ListView2.items[i].ImageIndex=4 then   // обновлять для старых компов, добавлять им новые задачи
  begin
  TransactionUpdate.StartTransaction;
  FDQueryUpdate.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryUpdate.params.ParamByName('p3').AsString:=''+ListView2.items[i].subitems[0]+'';           //PC_NAME
  step:=7;
  FDQueryUpdate.ExecSQL;
  TransactionUpdate.Commit;  /// сохранить изменения
  end;
end;
step:=8;

FDQueryUpdate.SQL.clear;   //очистить
FDQueryUpdate.Close;  /// закрыть нах после чтения
FDQueryUpdate.Free;
TransactionUpdate.Free;
End; /// окончание добавления новых задач для старых компов
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////заполнение таблицы всеми задачами для новых компов
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
NewZorPC:=0;
for i := 0 to ListView2.items.Count-1 do // проходим в цикле и смотрим добавлялисть ли новые компы
if ListView2.items[i].ImageIndex=0 then NewZorPC:=NewZorPC+1;
if NewZorPC<>0 then
Begin
TransactionUpdate:= TFDTransaction.Create(nil);
TransactionUpdate.Connection:=DataM.ConnectionDB;
TransactionUpdate.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionUpdate.Options.AutoCommit:=false;
TransactionUpdate.Options.AutoStart:=false;
TransactionUpdate.Options.AutoStop:=false;
FDQueryUpdate:= TFDQuery.Create(nil);
FDQueryUpdate.Transaction:=TransactionUpdate;
FDQueryUpdate.Connection:=DataM.ConnectionDB;
ColName:='';
ParamNum:='';
p:=3;
for i := 0 to ListView1.items.Count-1 do  // заполняем строковые переменные данными для sql запроса
Begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
End;
step:=9;
FDQueryUpdate.SQL.clear;
FDQueryUpdate.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME)';;
 step:=10;
FDQueryUpdate.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - общее количество задач
step:=11;
p:=3;
for i := 0 to ListView1.items.Count-1 do //цикл по задачам заполняем один раз так как эти параметры для каждой строки одинаковые
begin
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - тип задачи msi, proc, питание
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - номер выполняемой задачи
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // результат выполнения задачи
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // Описание задания
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO'; // Текущее состояние задания StatusTask
end;
step:=12;
for i := 0 to ListView2.items.Count-1 do /// цикл по компам заполняем и выполняем sql запрос т.к. строки будут отличаться только именем компа и ID
begin
 if ListView2.items[i].ImageIndex=0 then  // обновлять только для новых компов, добавляем все задачи
  begin
  TransactionUpdate.StartTransaction; // стартуем
  FDQueryUpdate.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryUpdate.params.ParamByName('p3').AsString:=''+ListView2.items[i].subitems[0]+'';           //PC_NAME
  step:=13;
  FDQueryUpdate.ExecSQL;
  TransactionUpdate.Commit;  /// сохранить изменения
  end;
end;
step:=14;

FDQueryUpdate.SQL.clear;   //очистить
FDQueryUpdate.Close;  /// закрыть нах после чтения
FDQueryUpdate.Free;
TransactionUpdate.Free;
End; /// окончание добавления заданий для новых компов
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
result:=true;
Except
   on E: Exception do
      begin
      if Assigned(FDQueryUpdate) then FDQueryUpdate.Free;
      if Assigned(TransactionUpdate) then
      begin
      TransactionUpdate.Rollback;
      TransactionUpdate.Free;
      end;
      //
      frmdomaininfo.memo1.Lines.add('Ошибка обновления задачи (шаг '+inttostr(step)+') :'+E.Message);
      result:=false;
      end;
   end;

end;


procedure TEditTask.SaveTask(Sender: TObject); /// сохранить задачу
var
NameNewTask:string;
i:integer;
function imageLV (b:boolean):boolean;
var
i:integer;
begin
if b then
  begin
  for I := 0 to ListView1.Items.Count-1 do ListView1.Items[i].ImageIndex:=1;
  for I := 0 to ListView2.Items.Count-1 do ListView2.Items[i].ImageIndex:=4;
  end;
end;
begin
try
if StatusStartStopTask(EditTask.Caption) then // если задача запущена (true)
 begin
    i:=MessageDlg('Не рекомендуется вносить изменения в запущенную задачу. Продолжить?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit
 end;
if (frmDomainInfo.TaskListView.Items.Count>=1) and (TaskName.Caption='') then
begin
ShowMessage('Больше задач только в зарегистрированной версии программы!!!');
exit;
end;
  if ListView2.Items.Count=0 then // если список компов пуст то их неоюходимо добавить
  begin
  ShowMessage('Добавьте в задачу компьютеры');
  PageControl1.TabIndex:=1;
  Button2.SetFocus;
  exit;
  end;

  if ListView1.Items.Count=0 then // если список заданий пуст то их неоюходимо добавить
  begin
  ShowMessage('Для сохранения задачи добавьте задания');
  exit;
  end;

if ListView1.Items.Count>30 then
begin
ShowMessage('Превышено допустимое число заданий в задаче, количество заданий '+#13#10+' не может быть больше 30, удалите лишние задания!!!');
exit
end;


if TaskName.Caption='' then // если takLabel пусто значит создается новая задача
 begin
  NameNewTask:=AddNewItemsTable(EditTask.Caption,ListView2.Items.Count); // Получаем имя нововй таблицы, создаем запись в таблице для таблиц задач
  if NameNewTask='' then
  begin
   ShowMessage('Не удалось сохранить задачу... ошибки в логах');
   exit;
  end;

  if createNewTableTask(NameNewTask) then // Создаем таблицу для задачи
  begin
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then // добавляем запись о таблице в главный lisView для задач
  imageLV(true); // меняем картинки в строках для обозначения того что эти задания и компы уесть в таблице
  TaskName.Caption:=NameNewTask; /// вносим в лабел имя таблицы, вдруг захотим ее сразу отредактировать
  end;
 end
else // иначе редактируем задачу
  begin
  NameNewTask:=TaskName.Caption;       // при редактировании имя таблицы передается в label
  if UpdateTableTask(NameNewTask) then // обновляем таблицу для задачи
  begin
  if UpdateTableTableTask(NameNewTask,ListView2.Items.Count) then // обновляем запись о таблице в TABLE_TASK
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then// сдесь обновляем информацию о измененной таблице в главном ListView для задач
     imageLV(true); // меняем картинки в строках для обозначения того что эти задания и компы уесть в таблице
  end;
  end;
Button7.Enabled:=false;
Button8.Enabled:=true;

except on E: Exception do frmDomainInfo.Memo1.Lines.Add('Ошибка сохранения задачи '+e.Message)
end;
end;

procedure TEditTask.Button7Click(Sender: TObject); /// сохранить задачу
var
NameNewTask:string;
i:integer;
function imageLV (b:boolean):boolean;
var
i:integer;
begin
if b then
  begin
  for I := 0 to ListView1.Items.Count-1 do ListView1.Items[i].ImageIndex:=1;
  for I := 0 to ListView2.Items.Count-1 do ListView2.Items[i].ImageIndex:=4;
  end;
end;
begin
try
if StatusStartStopTask(EditTask.Caption) then // если задача запущена (true)
 begin
    i:=MessageDlg('Не рекомендуется вносить изменения в запущенную задачу. Продолжить?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit
 end;
  if ListView2.Items.Count=0 then // если список компов пуст то их неоюходимо добавить
  begin
  ShowMessage('Добавьте в задачу компьютеры');
  PageControl1.TabIndex:=1;
  Button2.SetFocus;
  exit;
  end;

  if ListView1.Items.Count=0 then // если список заданий пуст то их неоюходимо добавить
  begin
  ShowMessage('Для сохранения задачи добавьте задания');
  exit;
  end;

if ListView1.Items.Count>30 then
begin
ShowMessage('Превышено допустимое число заданий в задаче, количество заданий '+#13#10+' не может быть больше 30, удалите лишние задания!!!');
exit
end;


if TaskName.Caption='' then // если takLabel пусто значит создается новая задача
 begin
  NameNewTask:=AddNewItemsTable(EditTask.Caption,ListView2.Items.Count); // Получаем имя нововй таблицы, создаем запись в таблице для таблиц задач
  if NameNewTask='' then
  begin
   ShowMessage('Не удалось сохранить задачу... ошибки в логах');
   exit;
  end;

  if createNewTableTask(NameNewTask) then // Создаем таблицу для задачи
  begin
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then // добавляем запись о таблице в главный lisView для задач
  imageLV(true); // меняем картинки в строках для обозначения того что эти задания и компы уесть в таблице
  TaskName.Caption:=NameNewTask; /// вносим в лабел имя таблицы, вдруг захотим ее сразу отредактировать
  end;
 end
else // иначе редактируем задачу
  begin
  NameNewTask:=TaskName.Caption;       // при редактировании имя таблицы передается в label
  if UpdateTableTask(NameNewTask) then // обновляем таблицу для задачи
  begin
  if UpdateTableTableTask(NameNewTask,ListView2.Items.Count) then // обновляем запись о таблице в TABLE_TASK
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then// сдесь обновляем информацию о измененной таблице в главном ListView для задач
     imageLV(true); // меняем картинки в строках для обозначения того что эти задания и компы уесть в таблице
  end;
  end;
Button7.Enabled:=false;
Button8.Enabled:=true;

except on E: Exception do frmDomainInfo.Memo1.Lines.Add('Ошибка сохранения задачи '+e.Message)
end;
end;

function TEditTask.AddUserPass(DescriptionTable,User,Passwd:string;SavePass:boolean):boolean; // Добавление пользователя и пароль для задачи
var
FDQueryStarTask:TFDQuery;
TransactionStarTask: TFDTransaction;
begin
try
TransactionStarTask:= TFDTransaction.Create(nil);
TransactionStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionStarTask.Options.AutoCommit:=false;
TransactionStarTask.Options.AutoStart:=false;
TransactionStarTask.Options.AutoStop:=false;
FDQueryStarTask:= TFDQuery.Create(nil);
FDQueryStarTask.Transaction:=TransactionStarTask;
FDQueryStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.StartTransaction; // стартуем
NameRunTASK:=nametableForDescription(DescriptionTable); /// имя таблицы по описанию задачи для передачи в поток
FDQueryStarTask.SQL.Clear;
FDQueryStarTask.SQL.Text:='update TABLE_TASK set USER_NAME=:a ,PASS_USER=:b, SAVE_PASS=:c WHERE NAME_TABLE='''+NameRunTASK+'''';
FDQueryStarTask.ParamByName('a').asstring:=user;    // пользователь
FDQueryStarTask.ParamByName('b').asstring:=Passwd;  // пароль
FDQueryStarTask.ParamByName('c').AsBoolean:=SavePass;// сохранять пароль после запуска задачи или нет
FDQueryStarTask.ExecSQL;
TransactionStarTask.Commit;  /// сохранить изменения
FDQueryStarTask.SQL.clear;   //очистить
FDQueryStarTask.Close;  /// закрыть нах после чтения
FDQueryStarTask.Free;
TransactionStarTask.Free;
result:=true;
except on E: Exception do
begin
if Assigned(FDQueryStarTask) then FDQueryStarTask.Free;
      if Assigned(TransactionStarTask) then
      begin
      TransactionStarTask.Rollback;
      TransactionStarTask.Free;
      end;
 frmDomainInfo.memo1.Lines.add('Ошибка добавления пользователя и пароля '+NameRunTASK+' : '+E.Message);
 result:=false;
end;
end;
end;

function TEditTask.StatrTask(DescriptionTable:string):boolean; // запуск задачи
var
TaskThread:TThread;
FDQueryStarTask:TFDQuery;
TransactionStarTask: TFDTransaction;
begin
try
TransactionStarTask:= TFDTransaction.Create(nil);
TransactionStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionStarTask.Options.AutoCommit:=false;
TransactionStarTask.Options.AutoStart:=false;
TransactionStarTask.Options.AutoStop:=false;
FDQueryStarTask:= TFDQuery.Create(nil);
FDQueryStarTask.Transaction:=TransactionStarTask;
FDQueryStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.StartTransaction; // стартуем
//NameRunTASK:=TaskName.Caption; // При сохранении или изменении задачи, имя таблицы сохраняется в   label
NameRunTASK:=nametableForDescription(DescriptionTable); /// имя таблицы по описанию задачи для передачи в поток
FDQueryStarTask.SQL.Clear;
FDQueryStarTask.SQL.Text:='update TABLE_TASK set STATUS_TASK=:a ,START_STOP=:b WHERE NAME_TABLE='''+NameRunTASK+'''';
FDQueryStarTask.ParamByName('a').asstring:='Запускается';
FDQueryStarTask.ParamByName('b').asstring:=booltostr(true);  // передаем true для разрешения запуска задачи
FDQueryStarTask.ExecSQL;
TransactionStarTask.Commit;  /// сохранить изменения
FDQueryStarTask.SQL.clear;   //очистить
FDQueryStarTask.Close;  /// закрыть нах после чтения
FDQueryStarTask.Free;
TransactionStarTask.Free;

TaskThread:=RunTask.TaskRun.Create(true);
TaskThread.FreeOnTerminate:=true;
TaskThread.Start;
result:=true;
except on E: Exception do
begin
if Assigned(FDQueryStarTask) then FDQueryStarTask.Free;
      if Assigned(TransactionStarTask) then
      begin
      TransactionStarTask.Rollback;
      TransactionStarTask.Free;
      end;
 frmDomainInfo.memo1.Lines.add('Ошибка запуска задачи '+NameRunTASK+' : '+E.Message);
 result:=false;
end;
end;
end;

procedure TEditTask.Button8Click(Sender: TObject); /// запустить задачу
begin
if StatusStartStopTask(EditTask.Caption) then // если задача запущена (true)
 begin
   showmessage('Задача '+EditTask.Caption+' уже запущена');
   exit;
 end;

 if  StatrTask(EditTask.Caption) then
 begin
 frmDomainInfo.Memo1.Lines.Add('Запуск задачи '+EditTask.Caption+' успешно выполнен');
 EditTask.Close;
 end;
end;

function TEditTask.deletelistPC(z:bool):boolean;
var
i:integer;
begin
if z then
begin
 i:=MessageDlg('Вы действительно хотите удалить весь список компьютеров?', mtConfirmation,[mbYes,mbCancel],0);
     if i=IDCancel then   exit;
end;
ListView2.SelectAll; //выделить  весь список компов
Button3.Click; //и удалить
end;

procedure TEditTask.Button9Click(Sender: TObject);
begin
deletelistPC(true);
end;

function TEditTask.StopTask(DescriptionTask:string):boolean; // остановка задачи
var
StopTASK:string;
FDQueryStopTask:TFDQuery;
TransactionStopTask: TFDTransaction;
begin
try
TransactionStopTask:= TFDTransaction.Create(nil);
TransactionStopTask.Connection:=DataM.ConnectionDB;
TransactionStopTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionStopTask.Options.AutoCommit:=false;
TransactionStopTask.Options.AutoStart:=false;
TransactionStopTask.Options.AutoStop:=false;
FDQueryStopTask:= TFDQuery.Create(nil);
FDQueryStopTask.Transaction:=TransactionStopTask;
FDQueryStopTask.Connection:=DataM.ConnectionDB;
TransactionStopTask.StartTransaction; // стартуем
StopTASK:=nametableForDescription(DescriptionTask);
FDQueryStopTask.SQL.Clear;
FDQueryStopTask.SQL.Text:='update TABLE_TASK set STATUS_TASK=:a ,START_STOP=:b WHERE NAME_TABLE='''+StopTASK+'''';
FDQueryStopTask.ParamByName('a').asstring:='Остановка задачи';
FDQueryStopTask.ParamByName('b').asstring:=booltostr(false); // устанавливаем статус задачи в false для того чтобы остановить её
FDQueryStopTask.ExecSQL;
TransactionStopTask.Commit;  /// сохранить изменения
FDQueryStopTask.SQL.clear;   //очистить
FDQueryStopTask.Close;  /// закрыть нах
FDQueryStopTask.Free;
TransactionStopTask.Free;
result:=true;
except on E: Exception do
begin
 if Assigned(FDQueryStopTask) then FDQueryStopTask.Free;
      if Assigned(TransactionStopTask) then
      begin
      TransactionStopTask.Rollback;
      TransactionStopTask.Free;
      end;
 frmdomaininfo.memo1.Lines.add('Ошибка остановки задачи '+StopTASK+':'+E.Message);
 result:=false;
 end;
end
end;



procedure TEditTask.Button6Click(Sender: TObject); // остановить задачу
begin
if not StatusStartStopTask(EditTask.Caption) then // если задача остановлена (false)
 begin
   frmDomainInfo.Memo1.Lines.Add('Задача '+EditTask.Caption+' уже остановлена');
   exit;
 end;
if StopTask(EditTask.Caption) then
frmDomainInfo.Memo1.Lines.Add('Операция, остановка задачи '+EditTask.Caption+' успешно выполнена');
end;

end.
