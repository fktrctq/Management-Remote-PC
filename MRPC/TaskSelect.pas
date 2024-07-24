unit TaskSelect;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Imaging.pngimage, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ImgList,FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls, Vcl.Menus;

type
  TSelectTask = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ListView1: TListView;
    CancelB: TButton;
    OkB: TButton;
    StaticText1: TStaticText;
    Image1: TImage;
    ImageList2: TImageList;
    Button1: TButton;
    Button2: TButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CancelBClick(Sender: TObject);
    procedure OkBClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
  private
    { Private declarations }
    function readTaskTables(s:string):boolean;
    function AddOrUpdateListPcForTask(AddOrUpdate,group:boolean):boolean; // Добавление или замена списка компов в задаче
  public
    { Public declarations }
  end;

var
  SelectTask: TSelectTask;

implementation
uses MyDM,TaskEdit,umain;
{$R *.dfm}

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

function TSelectTask.readTaskTables(s:string):boolean;
var
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
begin
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
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
FDQueryRead.SQL.Text:='SELECT * FROM TABLE_TASK WHERE STATUS_TASK<>''delete''';
FDQueryRead.Open;
ListView1.Clear;
 while not FDQueryRead.Eof do
 begin
   with ListView1.Items.Add do
   begin
   if strtobool(FDQueryRead.FieldByName('START_STOP').AsString) then
   ImageIndex:=0 else ImageIndex:=1;
   caption:=FDQueryRead.FieldByName('DESCRIPTION_TASK').AsString; // описание задачи
   if (not FDQueryRead.FieldByName('WORKS_THREAD').AsBoolean)and  // если поток остановлен
    (FDQueryRead.FieldByName('STATUS_TASK').AsString='Остановка задачи') then SubItems.add('Остановлена') //если поток утановил заначение  false то задача остановлена или завершена
   else SubItems.add(FDQueryRead.FieldByName('STATUS_TASK').AsString); // статус задачи (выполняется/остановлена/остановка)
   SubItems.add(FDQueryRead.FieldByName('CURRENT_TASK').AsString);     // Текущее задание
   SubItems.add(FDQueryRead.FieldByName('PC_RUN').AsString);      // на каком компьютере сейчас выполняется
   SubItems.add(FDQueryRead.FieldByName('COUNT_PC').AsString);    // всего компьютеров в задаче
   end;
 FDQueryRead.Next;
 end;
TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
Result:=true;
except on E: Exception do
 begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
      Log_write('TASK',timetostr(now)+' Ошибка чтения задач : '+e.Message);
  result:=false;
  end;
end;
end;

procedure TSelectTask.Button1Click(Sender: TObject);
begin                          // обновить
readTaskTables('');
end;



procedure TSelectTask.CancelBClick(Sender: TObject);
begin
close;
end;

procedure TSelectTask.FormShow(Sender: TObject);
begin                    //Загрузка списка при открытии формы
readTaskTables('');
end;

function NewNameTask(s:string):string;
var
nametask:string;
SelectedPC:Tstringlist;
i:integer;
WaitName:boolean;
begin
WaitName:=false;
while not WaitName do
Begin
if not InputQuery('Введите имя новой Задачи', 'Имя новой задачи', nametask)
        then break;
     if (nametask='') then break;
if SelectTask.ListView1.Items.Count>0 then
begin
for I := 0 to SelectTask.ListView1.Items.Count-1 do
begin
  if nametask=SelectTask.ListView1.Items[i].Caption then
  begin
    ShowMessage('Задача с таким именем уже существует');
    WaitName:=false;
    nametask:='';
    Break;
  end
  else WaitName:=true;
end;
end else WaitName:=true;
End;
result:=nametask;
end;

procedure TSelectTask.N1Click(Sender: TObject);  // новая задача
var
nametask:string;
WaitName,PCinList:boolean;
i:integer;
SelectedPC:TstringList;
begin
readTaskTables(''); // обновляем список задач
nametask:=NewNameTask('');
 if nametask='' then
 begin
 exit;
 end;

EditTask.ListView2.Clear; // очистка списка компов
EditTask.ListView1.Clear; // очистка списка заданий
 try
  SelectedPC:=TstringList.Create;
  SelectedPC:=frmdomaininfo.createListpcForCheck('');
  for I := 0 to SelectedPC.count-1 do  //заполнение списка компов
  begin
    with EditTask.ListView2.Items.Add do
    begin
    ImageIndex:=0;
    Caption := inttostr(EditTask.ListView2.Items.Count);
    SubItems.Add(SelectedPC[i]);
    end;
  end;
  finally
  SelectedPC.Free;
  end;

 Edittask.Caption:=nametask;
 EditTask.TaskName.Caption:=''; // очищаем label т.к. задача новая
 EditTask.Button8.Enabled:=false; //кнопка запуска не активна , активируется после сохранения задачи
 Edittask.Showmodal;
end;

procedure TSelectTask.N2Click(Sender: TObject); // редактировать задачу
begin
if ListView1.SelCount=1 then
begin
  EditTask.Caption:= ListView1.Selected.Caption; // имя задачи в заголовок
  EditTask.TaskName.Caption:=ListView1.Selected.Caption; // имя задачи в label т.к. происходит редактирование задачи
  //EditTask.ReadTableSelectedTask(EditTask.TaskName.Caption); // запускаем чтение таблицы и в label меняем описание таблицы на ее имя
  EditTask.ReadResulTask(EditTask.TaskName.Caption); // новая функция чтения таблицы
  EditTask.Button8.Enabled:=true;
  if Edittask.WindowState=wsMinimized then // если форма свернута то восстанавливаем ее
   Edittask.WindowState:=wsNormal
   else                                        // иначе показываем
   Edittask.Show;
end;
end;

procedure TSelectTask.N3Click(Sender: TObject);  // запуск задачи
begin
if ListView1.SelCount=1 then
begin
if EditTask.StatusStartStopTask(ListView1.Selected.Caption) then // если задача запущена (true)
 begin
   ShowMessage('Задача '+ListView1.Selected.Caption+' уже запущена');
   exit;
 end;
if  EditTask.StatrTask(ListView1.Selected.Caption) then
 frmdomaininfo.Memo1.Lines.Add('Запуск задачи '+ListView1.Selected.Caption+' успешно выполнен');
end;
readTaskTables(''); // обновить список задач
end;

procedure TSelectTask.N4Click(Sender: TObject); // остановка задачи
begin
if ListView1.SelCount=1 then
begin
if not EditTask.StatusStartStopTask(ListView1.Selected.Caption) then // если задача остановлена (false)
 begin
   ShowMessage('Задача уже остановлена');
   exit;
 end;
if edittask.StopTask(ListView1.Selected.Caption) then
frmDomainInfo.Memo1.Lines.Add('Операция, остановка задачи '+ListView1.Selected.Caption+' успешно выполнена');
end;
readTaskTables(''); // обновить список задач
end;

procedure TSelectTask.N5Click(Sender: TObject); // обновить список задач
begin
readTaskTables('');
end;

procedure TSelectTask.OkBClick(Sender: TObject);
begin
if SelectTask.tag=1 then
AddOrUpdateListPcForTask(true,true); // Заменяем список компов в задаче
if SelectTask.tag=0 then
AddOrUpdateListPcForTask(true,false); // заменяем список компов текущим компом
end;

procedure TSelectTask.Button2Click(Sender: TObject);
begin
if SelectTask.tag=1 then
AddOrUpdateListPcForTask(false,true); // добавляем список компов к задаче
if SelectTask.tag=0 then
AddOrUpdateListPcForTask(false,false); // добавляем в список текущий компьютер
end;


function TSelectTask.AddOrUpdateListPcForTask(AddOrUpdate,Group:boolean):boolean; // AddOrUpdate - обновить или добавить, Group - для группы компов или для одного компа
var                 //если AddOrUpdate то обновляем список компов в задаче
SelectedPC:tstringlist;
i,z:integer;
PCinList:boolean;
begin
if ListView1.SelCount=1 then
begin
  if EditTask.StatusStartStopTask(ListView1.Selected.Caption) then
   begin
     showmessage('Задача '+ListView1.Selected.Caption+' уже запущена'+#10#13+' остановите ее для внесения изменений.');
     exit;
   end;
  EditTask.Caption:= ListView1.Selected.Caption; // описание задачи в заголовок
  EditTask.TaskName.Caption:=ListView1.Selected.Caption; // имя задачи в label т.к. происходит редактирование задачи
  EditTask.ReadResulTask(EditTask.TaskName.Caption); // новая функция чтения таблицы
  EditTask.Button8.Enabled:=true; //
 if AddOrUpdate then EditTask.deletelistPC(false); // удаляем весь список компов
  try
  if Group then // если задача для списка компов
  begin
    SelectedPC:=TstringList.Create;
    SelectedPC:=frmdomaininfo.createListpcForCheck('');
    if SelectedPC.Count=0 then
    begin
    if frmdomaininfo.listview8.SelCount=1 then
     SelectedPC.Add(frmdomaininfo.listview8.Selected.SubItems[0])
     else
     begin
      ShowMessage('Не выделен список компьютеров!');
      SelectedPC.Free;
      exit;
     end;
    end;
    for I := 0 to SelectedPC.count-1 do  //заполнение списка компов
    begin
    PCinList:=false;
     for z := 0 to EditTask.ListView2.Items.Count-1 do // будем искать есть ли этот комп в списке для задачи
      if SelectedPC[i]=EditTask.ListView2.Items[z].SubItems[0] then // если комп уже в списке
      begin
      PCinList:=true;
      break;
      end;
      if not PCinList then // если компа нет в списке то добавить
      with EditTask.ListView2.Items.Add do
      begin
      ImageIndex:=0;
      Caption := inttostr(EditTask.ListView2.Items.Count);
      SubItems.Add(SelectedPC[i]);
      end;
    end;
  end;
  finally
  SelectedPC.Free;
  end;
  if not Group then  // если задача для текущего одного компа
  begin
   PCinList:=false;
     for z := 0 to EditTask.ListView2.Items.Count-1 do // будем искать есть ли этот комп в списке для задачи
      if frmdomaininfo.combobox2.text=EditTask.ListView2.Items[z].SubItems[0] then // если комп уже в списке
      begin
      PCinList:=true;
      break;
      end;
      if not PCinList then // если компа нет в списке то добавить
      with EditTask.ListView2.Items.Add do
      begin
      ImageIndex:=0;
      Caption := inttostr(EditTask.ListView2.Items.Count);
      SubItems.Add(frmdomaininfo.combobox2.text);
      end;
  end;

  EditTask.Button7.Click; // сохраняем все изменения
  SelectTask.Close; // Закрываю нах это окно

  EditTask.ShowModal;     // показываем окно задачи

end;

end;


end.

