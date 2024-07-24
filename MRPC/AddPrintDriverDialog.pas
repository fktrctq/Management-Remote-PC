unit AddPrintDriverDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Dialogs,Inifiles,
  Vcl.DBCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,
  FireDAC.Comp.DataSet,System.variants,System.StrUtils;

type
  TOKRightDlg12345678910 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Edit1: TEdit;
    SpeedButton1: TSpeedButton;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Edit2: TEdit;
    Label5: TLabel;
    LabeledEdit1: TLabeledEdit;
    Button1: TButton;
    CheckBox1: TCheckBox;
    SpeedButton3: TSpeedButton;
    SpeedButton2: TSpeedButton;
    FDTransactionDrvPrint: TFDTransaction;
    Label6: TLabel;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    Panel1: TPanel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ComboBox3Select(Sender: TObject);
  private
    procedure readDRVPrint;
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678910: TOKRightDlg12345678910;
  drvini:Tinifile;
  prefix:string;
  FilePath,DriverName,DriverPath:string;
  AddNewDriverPrinter:TThread;

  SelectedPCInstalDrv:TstringList;

implementation
uses umain,AddprintDriverThread,SelectedPCInstallDriverPrint,MyDM;

{$R *.dfm}
procedure TOKRightDlg12345678910.readDRVPrint;
var
FDQuery:TFDQuery;
begin
try
ComboBox3.Clear;
ComboBox4.Clear;
DataM.СreateTablForDriverPrint('DRIVER_PRINT');
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionDrvPrint;
FDQuery.Connection:=DataM.ConnectionDB;
FDQuery.SQL.Text:='SELECT * FROM DRIVER_PRINT ORDER BY DRV_DESCRIPT';
FDQuery.Open;
while not FDQuery.Eof do
begin
 if FDQuery.FieldByName('DRV_DESCRIPT').Value<>null then
 begin
   ComboBox3.Items.Add(FDQuery.FieldByName('DRV_DESCRIPT').AsString);
   ComboBox4.Items.Add(FDQuery.FieldByName('DRV_PATCH').AsString);
 end;
 FDQuery.Next;
end;
finally
FDQuery.Close;
FDQuery.Free;
end;
end;
 ///////////////////////////////////////////////
procedure RemoveDuplicates(const stringList : TStringList) ;
var                         /// удаление повторяющихся строк в stringlist
  buffer: TStringList;
  cnt: Integer;
begin
  stringList.Sort;
  buffer := TStringList.Create;
  try
    buffer.Sorted := True;
    buffer.Duplicates := dupIgnore;
    buffer.BeginUpdate;
    for cnt := 0 to stringList.Count - 1 do
      buffer.Add(stringList[cnt]) ;
    buffer.EndUpdate;
    stringList.Assign(buffer) ;
  finally
    FreeandNil(buffer) ;
  end;
end;
//////////////////////////////////////////////
procedure TOKRightDlg12345678910.Button1Click(Sender: TObject);
var
MyDialog: TOpenDialog;
begin
MyDialog:=TOpenDialog.Create(Self);

  MyDialog.Name:='MyDialog';
  MyDialog.Filter:='|*.dll';
  MyDialog.Title:='Путь к dll файлу драйвера';
  if DirectoryExists(ExtractFileDir(LabeledEdit1.Text)) then
   MyDialog.InitialDir:=(ExtractFileDir(LabeledEdit1.Text));
  if MyDialog.Execute then
    begin
      LabeledEdit1.Text:=MyDialog.FileName;
    end;

if Assigned(MyDialog) then  FreeAndNil(MyDialog);
end;


procedure TOKRightDlg12345678910.CancelBtnClick(Sender: TObject);
begin
drvini.Free;
CLOSE;
end;

procedure TOKRightDlg12345678910.CheckBox1Click(Sender: TObject);
begin
LabeledEdit1.Enabled:=CheckBox1.Checked;
Button1.Enabled:= CheckBox1.Checked;
end;


procedure TOKRightDlg12345678910.ComboBox1Change(Sender: TObject);
var
namesection:Tstringlist;
i:integer;
begin
try
if (not Assigned(drvini))and FileExists(edit1.Text) then drvini:=TiniFile.Create(edit1.Text);  ///если ини файла нет и файл в едите доступен
namesection:=tstringlist.Create;
if combobox1.Text<>'Other' then
     begin
      if drvini.SectionExists(prefix+'.'+combobox1.Text) then
        begin
        combobox2.Clear;
        //drvini.ReadSection(prefix+'.'+combobox1.Text,combobox2.Items);
        drvini.ReadSection(prefix+'.'+combobox1.Text,namesection);
        RemoveDuplicates(namesection);
          for I := 0 to namesection.Count-1 do
            begin
            combobox2.Items.Add(StringReplace(namesection[i],'"','',[rfReplaceAll]));
            end;
        end;
      end
  else
      begin
      combobox2.Clear;
        drvini.ReadSection('Strings',namesection);
        RemoveDuplicates(namesection);
          for I := 0 to namesection.Count-1 do
            begin
            combobox2.Items.Add(StringReplace(drvini.ReadString('Strings',namesection[i],''),'"','',[rfReplaceAll]));
            end;
      end;
namesection.Free;
except
    on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add('Ошибка "'+E.Message+'"');
    end;
   end;
end;

procedure TOKRightDlg12345678910.ComboBox3Select(Sender: TObject);
var
FDQuery:TFDQuery;
begin
ComboBox4.ItemIndex:=ComboBox3.ItemIndex;
try
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionDrvPrint;
FDQuery.Connection:=DataM.ConnectionDB;
FDQuery.SQL.Text:='SELECT * FROM DRIVER_PRINT WHERE DRV_DESCRIPT='''+ComboBox3.Text+''''+' AND DRV_PATCH='''+ComboBox4.Text+'''';
FDQuery.Open;

Edit1.Text:=vartostr(FDQuery.FieldByName('DRV_PATCH').Value);
/// после ввода в Edit1 надо типа создать ини файл а после заполнять остальные элементы
LabeledEdit1.Text:=vartostr(FDQuery.FieldByName('DRV_DLL').Value);
if vartostr(FDQuery.FieldByName('DRV_DLL_TRUE').Value)='1' then
CheckBox1.Checked:=true else  CheckBox1.Checked:=false;
ComboBox1.Text:= vartostr(FDQuery.FieldByName('SECTION_INF').Value);
Edit2.Text:=vartostr(FDQuery.FieldByName('DRV_CLASS').Value);
ComboBox2.Text:= vartostr(FDQuery.FieldByName('DRV_NAME').Value);

finally
FDQuery.Close;
FDQuery.Free;
end;
end;



procedure TOKRightDlg12345678910.FormCreate(Sender: TObject);
begin
CheckBox1.Checked:=false;
end;

procedure TOKRightDlg12345678910.FormShow(Sender: TObject);
begin
MyPS:=frmDomainInfo.ComboBox2.Text;
MyUser:=frmDomainInfo.LabeledEdit1.Text;
MyPasswd:=frmDomainInfo.LabeledEdit1.Text;

if not DataM.ConnectionDB.Connected then  // если нет соединения с БД то выходим
begin
 ComboBox3.Enabled:=DataM.ConnectionDB.Connected;
 exit;
end;
ComboBox3.Enabled:=DataM.ConnectionDB.Connected;
readDRVPrint;/// чтение списка драйверов из БД
end;

procedure TOKRightDlg12345678910.OKBtnClick(Sender: TObject);
var
i:integer;
begin

if fileexists(Edit1.Text)<>true then
begin
i:=MessageBox(Self.Handle, PChar('Файл Inf не доступен!'+#10#13+'Продолжить выполнение операции?')
, PChar('Установка драйвера принтера...') ,MB_YESNO+MB_ICONQUESTION);
if i=IDNO then exit;
end;

if combobox2.Text='' then
begin
i:=MessageBox(Self.Handle, PChar('Вы не указали драйвер принтера!'+#10#13+'Продолжить выполнение операции?')
, PChar('Установка драйвера принтера...') ,MB_YESNO+MB_ICONQUESTION);
if i=IDNO then exit;
end;

if (not  FileExists(LabelEdEdit1.text)) and CheckBox1.Checked then
begin
i:=MessageBox(Self.Handle, PChar('Не доступен dll файл дравера!'+#10#13+'Продолжить выполнение операции?')
, PChar(LabelEdEdit1.text) ,MB_YESNO+MB_ICONQUESTION);
if i=IDNO then exit;
end;

if GroupPC=false then
begin
FilePath:=Edit1.Text;
DriverName:=Combobox2.Text;
if CheckBox1.Checked then DriverPath:=LabeledEdit1.Text /// если указали длл файл
else DriverPath:='';
AddNewDriverPrinter:=AddPrintDriverThread.AddPrintDriver.Create(true);
AddNewDriverPrinter.FreeOnTerminate:=true;
AddNewDriverPrinter.Start;
end;
if GroupPC=true then
begin
SelectedPCInstalDrv:=TstringList.Create;
SelectedPCInstalDrv:=frmDomainInfo.createListpcForCheck(''); /// получаем список компьютеров
if SelectedPCInstalDrv.Count=0 then
begin
  ShowMessage('Не выбран список компьютеров');
  exit;
  SelectedPCInstalDrv.free;
end;
FilePath:=Edit1.Text;
DriverName:=Combobox2.Text;
if CheckBox1.Checked then DriverPath:=LabeledEdit1.Text /// если указали длл файл
else DriverPath:='';
AddNewDriverPrinter:=SelectedPCInstallDriverPrint.SelectedPCInstallDriver.Create(true);
AddNewDriverPrinter.FreeOnTerminate:=true;
AddNewDriverPrinter.Start;
end;

if Assigned(drvini) then freeandnil(drvini);
close;
end;


procedure TOKRightDlg12345678910.SpeedButton1Click(Sender: TObject);
var
namesection:Tstringlist;
stringkey:string;
i:integer;
MyDialog: TOpenDialog;
begin
MyDialog:=TOpenDialog.Create(Self);
MyDialog.Name:='MyDialog';
MyDialog.Filter:='|*.inf';
MyDialog.Title:='Укажите путь к inf файлу драйвера';
if FileExists(Edit1.Text) then MyDialog.InitialDir:=(ExtractFileDir(Edit1.Text));
combobox1.Clear;
combobox2.Clear;
combobox3.Text:='';
combobox4.Text:='';
namesection:=Tstringlist.Create;
if MyDialog.Execute  then
begin
edit1.Text:=MyDialog.FileName;
if ExtractFilePath(LabeledEdit1.Text)<>ExtractFilePath(edit1.Text) then
LabeledEdit1.Text:=ExtractFilePath(MyDialog.FileName)+
copy(ExtractFileName(MyDialog.FileName),1,length(ExtractFileName(MyDialog.FileName))-4)+'.dll'
end
else
begin
if Assigned(MyDialog) then  FreeAndNil(MyDialog);
namesection.Free;
exit;
end;

if fileexists(edit1.Text) then
try
    begin
    drvini:=TiniFile.Create(edit1.Text);
    Edit2.Text:=drvini.ReadString('Version','Class','NONE');
    drvini.ReadSection('Manufacturer',namesection); /// считываем имя секции
    stringkey:=drvini.ReadString('Manufacturer',namesection[0],'');
    stringkey:=StringReplace(stringkey,' ','',[rfReplaceAll]);
    insert(',',stringkey,Length(stringkey)+1);
    prefix:=copy(stringkey,1,(pos(',',stringkey)-1)); /// префикс перед архитектурой
    while pos(',',stringkey)<>0 do            /// добавляем список архитектур
      begin
      delete(stringkey,1,(pos(',',stringkey)));
      //delete(stringkey,1,(pos(' ',stringkey)));  //// изменить эту херь
      combobox1.Items.Add(copy(stringkey,1,(pos(',',stringkey)-1)));
      end;
    combobox1.Items.delete(combobox1.Items.Count-1);
    if drvini.SectionExists('Strings') then  combobox1.Items.Add('Other');
    combobox1.ItemIndex:=0;
    label5.Caption:=(prefix);
    end;
    ////////////////////////////////////////////
    namesection.Clear;
    drvini.ReadSection(prefix+'.'+combobox1.Text,namesection);
    RemoveDuplicates(namesection);
     for I := 0 to namesection.Count-1 do
      begin
      combobox2.Items.Add(StringReplace(namesection[i],'"','',[rfReplaceAll]));
      end;
except
  showmessage('Тут драйверов нет');
end;
if Assigned(MyDialog) then  FreeAndNil(MyDialog);
namesection.Free;
end;

procedure TOKRightDlg12345678910.SpeedButton2Click(Sender: TObject); // удаление записи
var
FDQuery:TFDQuery;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionDrvPrint;
FDQuery.Connection:=DataM.ConnectionDB;
try
FDQuery.SQL.Clear;
FDQuery.SQL.Text:='DELETE FROM DRIVER_PRINT WHERE DRV_DESCRIPT='''
+ComboBox3.text+''' and DRV_PATCH='''+ComboBox4.text+'''';
FDQuery.ExecSQL;
// открываем список  процессов из базы
readDRVPrint;/// чтение списка драйверов из БД
finally
FDQuery.Close;
FDQuery.Free;
end;
end;

procedure TOKRightDlg12345678910.SpeedButton3Click(Sender: TObject);
var
Newdescription:string;
i:integer;
FDQuery:TFDQuery;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;

i:=MessageBox(Self.Handle, PChar('Сохранить текущие настроки драйвера?')
        , PChar('Драйвер - '+Combobox2.Text ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;
Newdescription:=ComboBox3.Text;
if ComboBox3.Text='' then
begin
Newdescription:=InputBox('Добавление описания к драйверу', 'Описание:', 'Сдесь ваше описание');
if (Newdescription='') or (Newdescription='Сдесь ваше описание') then Newdescription:=Combobox2.Text;
end;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=FDTransactionDrvPrint;
FDQuery.Connection:=DataM.ConnectionDB;
try
FDQuery.SQL.Text:='update or insert into DRIVER_PRINT'+
' (DRV_PATCH,DRV_DLL,DRV_DLL_TRUE,SECTION_INF,DRV_CLASS,DRV_NAME,DRV_DESCRIPT)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7) MATCHING (DRV_DESCRIPT,DRV_PATCH)';
FDQuery.params.ParamByName('p1').AsString:=''+leftstr(Edit1.Text,1500)+'';  //  DRV_PATCH
FDQuery.params.ParamByName('p2').AsString:=''+leftstr(LabeledEdit1.Text,1500)+''; //DRV_DLL
if CheckBox1.Checked then FDQuery.params.ParamByName('p3').AsString:=''+leftstr('1',1)+''  //  DRV_DLL_TRUE
else FDQuery.params.ParamByName('p3').AsString:=''+leftstr('0',1)+'';                       //  DRV_DLL_TRUE
FDQuery.params.ParamByName('p4').AsString:=''+leftstr(Combobox1.Text,200)+'';     //SECTION_INF
FDQuery.params.ParamByName('p5').AsString:=''+leftstr(Edit2.Text,200)+'';    //DRV_CLASS
FDQuery.params.ParamByName('p6').AsString:=''+leftstr(Combobox2.Text,200)+'';    //DRV_NAME
FDQuery.params.ParamByName('p7').AsString:=''+leftstr(Newdescription,500)+''; //DRV_DESCRIPT
FDQuery.ExecSQL;
readDRVPrint;/// чтение списка драйверов из БД
finally
FDQuery.Close;
FDQuery.Free;
end;
end;
end.
