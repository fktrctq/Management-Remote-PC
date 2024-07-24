unit TaskNewMSI;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Dialogs,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.variants,System.StrUtils,IdIcmpClient,ComObj,ActiveX,ShellAPI;

type
  TNewMSITask = class(TForm)
    GroupBox1: TGroupBox;
    ComboBox1: TComboBox;
    Button1: TButton;
    EditCopyPath: TLabeledEdit;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    GroupBox2: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    CancelBtn: TButton;
    CheckBox1: TCheckBox;
    LabeledEdit1: TLabeledEdit;
    OKBtn: TButton;
    Edit1: TEdit;
    CheckBox3: TCheckBox;
    FDQueryProc: TFDQuery;
    FDTransactionReadProc: TFDTransaction;
    Label3: TLabel;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    SpeedButton5: TSpeedButton;
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    function  readproc(item:integer):boolean;
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox3Select(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  NewMSITask: TNewMSITask;
  //ExitFor:Boolean;


implementation
uses MyDM,TaskEdit,EditProcMSI;
{$R *.dfm}



function TNewMSITask.readproc(item:integer):boolean;
begin
FDQueryproc.SQL.clear;
FDQueryProc.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''msi'''
+' ORDER BY DESCRIPTION_PROC';;
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
if ComboBox2.Items.Count>0 then
begin
ComboBox2.ItemIndex:=item;
ComboBox3.ItemIndex:=item;
ComboBox2.OnSelect(ComboBox2);
end;

end;


procedure TNewMSITask.Button1Click(Sender: TObject);
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

procedure TNewMSITask.CancelBtnClick(Sender: TObject);
begin
Close;
end;

procedure TNewMSITask.CheckBox3Click(Sender: TObject);
begin
GroupBox1.Enabled:=CheckBox3.Checked;
if CheckBox3.Checked then groupbox1.Height:=146
else  groupbox1.Height:=25;
end;

procedure TNewMSITask.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
combobox2.DroppedDown:=true;
end;

procedure TNewMSITask.ComboBox2Select(Sender: TObject);
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

procedure TNewMSITask.ComboBox3Select(Sender: TObject);
begin
ComboBox2.ItemIndex:=ComboBox3.ItemIndex;
end;



procedure TNewMSITask.FormShow(Sender: TObject);
begin
CheckBox3.Checked:=false;
GroupBox1.Height:=25;
if not DataM.ConnectionDB.Connected then exit;
readproc(0);
end;

procedure TNewMSITask.OKBtnClick(Sender: TObject);
var
i,z:integer;
begin
if (ComboBox2.Text='') and (ComboBox3.Text='') then
begin
  ShowMessage('Выберите программу для добавления в задачу');
  exit;
end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='Установка программы :'+ComboBox2.text)
    and (EditTask.ListView1.Items[i].SubItems[1]='msi') then
    begin
      z:=MessageDlg('Это задание уже добавлено в задачу!!!'+#10#13+' Вы уверены в своих действиях???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with EditTask.ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=ComboBox3.Text; // номер процесса или msi
  SubItems.Add('Установка программы :'+ComboBox2.Text);    // описание
  SubItems.Add('msi');
end;

end;

procedure TNewMSITask.SpeedButton1Click(Sender: TObject);
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
   Newdescription:=InputBox('Добавление описания к файлу', 'описание:', 'здесь описание программы');
   // если не указали описание то добавляем в описание имя выбранного файла
   if (Newdescription='') and (fileexists(NewOpenDlg.FileName)) then Newdescription:=ExtractFileName(NewOpenDlg.FileName);
  end;
// записываем новую запись в базу
if (Newdescription='') or (NewOpenDlg.FileName='') then exit;

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
FDQueryProc.Close;

// открываем список из базы
readproc(ComboBox2.Items.Count);
///уничтожаем  диалог
newOpenDlg.Free;
end;

procedure TNewMSITask.SpeedButton2Click(Sender: TObject); // удаляем
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
FDQueryProc.SQL.Text:='DELETE from START_PROC_MSI WHERE ID_PROC='''+ComboBox3.Text+'''';
FDQueryProc.ExecSQL;
FDQueryProc.Close;
// открываем список  процессов из базы
readproc(ComboBox2.ItemIndex-1);
end;

procedure TNewMSITask.SpeedButton3Click(Sender: TObject);
begin
//if TNewMSITask.findComponent('MyDialog')=nil then
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



procedure TNewMSITask.SpeedButton4Click(Sender: TObject); /// сохранить текущие настройки для выбранной программы
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
 descript:=InputBox('Добавьте описания к программе', 'Описание:', 'сдесь описание программы');
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
readproc(ComboBox2.ItemIndex);
end;

procedure TNewMSITask.SpeedButton5Click(Sender: TObject);
begin
FormEditMsiProc.ButtonPM.Caption:='Программы MSI';
FormEditMsiProc.Caption:='Редактор программ';
FormEditMsiProc.ShowModal;
end;

end.
