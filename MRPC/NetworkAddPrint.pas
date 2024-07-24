unit NetworkAddPrint;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Dialogs,Inifiles,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls,System.variants,System.StrUtils;

type
  TOKRightDlg123456789 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    SpeedButton1: TSpeedButton;
    Edit1: TEdit;
    Edit2: TEdit;
    ComboBox1: TComboBox;
    Label1: TLabel;
    Label4: TLabel;
    Label3: TLabel;
    ComboBox2: TComboBox;
    Label5: TLabel;
    Label2: TLabel;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    CheckBox1: TCheckBox;
    GroupBox3: TGroupBox;
    CheckBox2: TCheckBox;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    LabeledEdit6: TLabeledEdit;
    CheckBox5: TCheckBox;
    LabeledEdit1: TLabeledEdit;
    Button1: TButton;
    CheckBox6: TCheckBox;
    FDQueryDrvPrint: TFDQuery;
    FDTransactionDrvPrint: TFDTransaction;
    DataSourceDrvPrint: TDataSource;
    DBLookupComboBox1: TDBLookupComboBox;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Label6: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure ComboBox2Select(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LabeledEdit2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CheckBox5Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure DBLookupComboBox1MouseEnter(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormResize(Sender: TObject);
    procedure DBLookupComboBox1Click(Sender: TObject);
  private
    procedure readDRVPrint;
  public
    { Public declarations }
  end;

var
  OKRightDlg123456789: TOKRightDlg123456789;
  NewAddNetPrint:TThread;
  drvini:Tinifile;
  prefix:string;
  FilePath,DriverName,DriverPath,PortName,HostIPAddress,PortNumber:string;
  AddNewDriverPrinter:TThread;

  PortProtocol:integer;
  SNMPEnabled,PrintNetwork,PrintShare,installDriver,PrintIP,printNet:boolean;
  DriverNamePrintNet,PortNamePrintNet,PrintNameNet:string;

implementation
uses umain,AddNetworkPrint,SelectedPCAddPrinterThread,MyDM;
{$R *.dfm}
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


procedure TOKRightDlg123456789.Button1Click(Sender: TObject);
var
MyDialog: TOpenDialog;
begin
 MyDialog:=TOpenDialog.Create(Self);
  begin
  MyDialog.Name:='MyDialog';
  MyDialog.Filter:='|*.dll';
  MyDialog.Title:='Путь к dll файлу драйвера';
  if DirectoryExists(ExtractFileDir(LabeledEdit1.Text)) then
   MyDialog.InitialDir:=(ExtractFileDir(LabeledEdit1.Text));
  if MyDialog.Execute then
    begin
      LabeledEdit1.Text:=MyDialog.FileName;
    end;

  end;
  if Assigned(MyDialog) then  FreeAndNil(MyDialog);
end;

procedure TOKRightDlg123456789.CheckBox1Click(Sender: TObject);
begin
groupBox2.Enabled:=CheckBox1.Checked;
groupBox2.Enabled:=CheckBox1.Checked;
LabeledEdit6.Enabled:=not(CheckBox1.Checked);
installDriver:=CheckBox1.Checked;
end;

procedure TOKRightDlg123456789.CheckBox2Click(Sender: TObject);
begin
groupBox3.Enabled:=CheckBox2.Checked;

PrintIP:=CheckBox2.Checked;
end;

procedure TOKRightDlg123456789.CheckBox3Click(Sender: TObject);
begin
PrintNetwork:=CheckBox3.Checked;
end;

procedure TOKRightDlg123456789.CheckBox4Click(Sender: TObject);
begin
PrintShare:=CheckBox4.Checked;
end;

procedure TOKRightDlg123456789.CheckBox5Click(Sender: TObject);
begin
SNMPEnabled:=CheckBox5.Checked;
end;



procedure TOKRightDlg123456789.CheckBox6Click(Sender: TObject);
begin
LabeledEdit1.Enabled:=CheckBox6.Checked;
Button1.Enabled:= CheckBox6.Checked;
end;



procedure TOKRightDlg123456789.ComboBox2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
LabeledEdit6.Text:=ComboBox2.Text;
end;

procedure TOKRightDlg123456789.ComboBox2Select(Sender: TObject);
begin
LabeledEdit6.Text:=ComboBox2.Text;
end;

procedure TOKRightDlg123456789.DBLookupComboBox1Click(Sender: TObject);
var
i:integer;
begin
if not DataM.ConnectionDB.Connected then exit;
if DBLookupComboBox1.Text='' then exit;

i:=MessageBox(Self.Handle, PChar('Выбрать профиль драйвера '+ DBLookupComboBox1.Text+'?')
, PChar('Избранное - '+DBLookupComboBox1.Text) ,MB_YESNO+MB_ICONQUESTION);
if i=IDNO then exit;

ComboBox1.Clear;
ComboBox2.Clear;

Edit1.Text:=vartostr(FDQueryDrvPrint.FieldByName('DRV_PATCH').Value);
/// после ввода в Edit1 надо типа создать ини файл а после заполнять остальные элементы
LabeledEdit1.Text:=vartostr(FDQueryDrvPrint.FieldByName('DRV_DLL').Value);
if vartostr(FDQueryDrvPrint.FieldByName('DRV_DLL_TRUE').Value)='1' then
CheckBox6.Checked:=true else  CheckBox6.Checked:=false;
ComboBox1.Text:= vartostr(FDQueryDrvPrint.FieldByName('SECTION_INF').Value);
Edit2.Text:=vartostr(FDQueryDrvPrint.FieldByName('DRV_CLASS').Value);
ComboBox2.Text:= vartostr(FDQueryDrvPrint.FieldByName('DRV_NAME').Value);
LabeledEdit6.Text:=ComboBox2.Text;

end;

procedure TOKRightDlg123456789.DBLookupComboBox1MouseEnter(Sender: TObject);
begin
//if not DataM.ConnectionDB.Connected then exit;
//DBLookupComboBox1.Hint:=vartostr(FDQueryDrvPrint.FieldByName('DRV_DESCRIPT').Value);
end;

procedure TOKRightDlg123456789.readDRVPrint;
begin
DataM.СreateTablForDriverPrint('DRIVER_PRINT');
FDQueryDrvPrint.SQL.clear;
FDQueryDrvPrint.SQL.Text:='SELECT * FROM DRIVER_PRINT';
FDQueryDrvPrint.Open;
end;

procedure TOKRightDlg123456789.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
DBLookupComboBox1.CloseUp(true);
FDQueryDrvPrint.SQL.clear;
FDQueryDrvPrint.Close;
end;

procedure TOKRightDlg123456789.FormResize(Sender: TObject);
begin
Width:=453;
Height:=490;
end;

procedure TOKRightDlg123456789.FormShow(Sender: TObject);
begin
//CheckBox1.Checked:=false;
CheckBox2.Checked:=false;
CheckBox6.Checked:=false;
ComboBox2.ItemIndex:=-1;
MyPS:=frmDomainInfo.combobox2.Text;
MyUser:=frmDomainInfo.LabeledEdit1.Text;
MyPasswd:=frmDomainInfo.LabeledEdit2.Text;

if not DataM.ConnectionDB.Connected then  // если нет соединения с БД то выходим
begin
 DBLookupComboBox1.Enabled:=DataM.ConnectionDB.Connected;
 exit;
end;
DBLookupComboBox1.Enabled:=DataM.ConnectionDB.Connected;
readDRVPrint;/// чтение списка драйверов из БД
end;

procedure TOKRightDlg123456789.LabeledEdit2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
LabeledEdit5.Text:='IP_'+LabeledEdit2.Text;
end;

procedure TOKRightDlg123456789.OKBtnClick(Sender: TObject);
var
b,z,i:integer;

begin
installdriver:=CheckBox1.Checked;
PrintIP:=CheckBox2.Checked;


if GroupPC=false then ///// естановка сетевого принтера на одну машину
  BEGIN
    if CheckBox1.Checked then   //// если необходим драйвер
    Begin
       if (Edit1.Text='')or(Combobox2.Text='') then
         begin
           showmessage('Не указано имя драйвера или путь к Inf файлу драйвера!');
           exit;
         end;
       if (not  FileExists(LabelEdEdit1.text))and (CheckBox6.Checked) then
         begin
         i:=MessageBox(Self.Handle, PChar('Не доступен dll файл дравера!'+#10#13+'Продолжить выполнение операции?')
          , PChar(LabelEdEdit1.text) ,MB_YESNO+MB_ICONQUESTION);
         if i=IDNO then exit;
         end;

       FilePath:=Edit1.Text; //путь к inf файлу
       DriverName:=Combobox2.Text;  // имя драйвера
       if CheckBox6.Checked then DriverPath:=LabeledEdit1.Text /// если указали длл файл
       else DriverPath:='';
    End;
    if CheckBox2.Checked then    /// принтер сетевой по IP
    Begin
      if (LabeledEdit2.text='')or(LabeledEdit3.text='')
        or(LabeledEdit5.text='') then
        begin
          showmessage('Заполнены не все поля!');
          exit;
        end;
      PrintNetwork:=CheckBox3.Checked;
      printShare:=CheckBox4.Checked;
      SNMPEnabled:=CheckBox5.Checked;
      HostIPAddress:=LabeledEdit2.text;
      PortNumber:=LabeledEdit3.text;
      PortProtocol:=strtoint(LabeledEdit4.text);
      PortName:=LabeledEdit5.text;
      DriverName:=LabeledEdit6.text;
    End;
  NewAddNetPrint:= AddNetworkPrint.AddNetPrint.Create(true);
  NewAddNetPrint.FreeOnTerminate:=true;
  NewAddNetPrint.Start;
  END;


if GroupPC=true then    //// установка сетевого принтера на группу машин
  BEGIN
    if CheckBox1.Checked then   //// если необходим драйвер
     Begin
       if (Edit1.Text='')or(Combobox2.Text='') then
         begin
         showmessage('Не указано имя драйвера или путь к Inf файлу драйвера!');
         exit;
         end;
       if (not  FileExists(LabelEdEdit1.text))and (CheckBox6.Checked) then
         begin
         i:=MessageBox(Self.Handle, PChar('Не доступен dll файл дравера!'+#10#13+'Продолжить выполнение операции?')
          , PChar(LabelEdEdit1.text) ,MB_YESNO+MB_ICONQUESTION);
         if i=IDNO then exit;
         end;
       FilePath:=Edit1.Text; //путь к inf файлу
       DriverName:=Combobox2.Text;
       if CheckBox6.Checked then DriverPath:=LabeledEdit1.Text /// если указали длл файл
       else DriverPath:='';
      End;
    if CheckBox2.Checked then    /// принтер сетевой по IP
      Begin
      PrintNetwork:=CheckBox3.Checked;
      printShare:=CheckBox4.Checked;
      SNMPEnabled:=CheckBox5.Checked;
      HostIPAddress:=LabeledEdit2.text;
      PortNumber:=LabeledEdit3.text;
      PortProtocol:=strtoint(LabeledEdit4.text);
      PortName:=LabeledEdit5.text;
      DriverName:=LabeledEdit6.text;
      End;
  NewAddNetPrint:= SelectedPCAddPrinterThread.SelectedPCAddPrint.Create(true);
  NewAddNetPrint.FreeOnTerminate:=true;
  NewAddNetPrint.Start;
 END;
  if Assigned(drvini) then freeandnil(drvini);
  close;
end;


procedure TOKRightDlg123456789.ComboBox1Change(Sender: TObject);
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


procedure TOKRightDlg123456789.SpeedButton1Click(Sender: TObject);
var
namesection:Tstringlist;
stringkey:string;
i:integer;
MyDialog: TOpenDialog;
begin


MyDialog:=TOpenDialog.Create(Self);
MyDialog.Name:='MyDialog';
MyDialog.Filter:='|*.inf';
MyDialog.Title:='Путь к inf файлу драйвера';
if FileExists(Edit1.Text) then MyDialog.InitialDir:=(ExtractFileDir(Edit1.Text));

combobox1.Clear;
combobox2.Clear;
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
    //label2.Caption:=ExtractFileDir(MyDialog.FileName);
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
  showmessage('Драйвер не обнаружен.');
end;

if Assigned(namesection) then freeandnil(namesection);
if Assigned(MyDialog) then  FreeAndNil(MyDialog);
//namesection.clear;


end;


procedure TOKRightDlg123456789.SpeedButton2Click(Sender: TObject);
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;
FDQueryDrvPrint.Active:=false;
FDQueryDrvPrint.SQL.Clear;
FDQueryDrvPrint.SQL.Text:='DELETE FROM DRIVER_PRINT WHERE DRV_DESCRIPT='''
+VarToStr(DBLookupComboBox1.KeyValue)+'''';
FDQueryDrvPrint.ExecSQL;
// открываем список  процессов из базы
readDRVPrint;/// чтение списка драйверов из БД
end;

procedure TOKRightDlg123456789.SpeedButton3Click(Sender: TObject);
var
Newdescription:string;
i:integer;
begin
if not DataM.ConnectionDB.Connected then
begin
 ShowMessage('Нет подключения к БД');
 exit;
end;

i:=MessageBox(Self.Handle, PChar('Сохранить текущие настроки драйвера?')
        , PChar('Драйвер - '+Combobox2.Text ) ,MB_YESNO+MB_ICONQUESTION);
   if i=IDNO then exit;


Newdescription:=InputBox('Добавление описания к драйверу', 'Описание:', 'Сдесь ваше описание');
if Newdescription='' then Newdescription:=Combobox2.Text;

FDQueryDrvPrint.SQL.Clear;
FDQueryDrvPrint.SQL.Text:='INSERT INTO DRIVER_PRINT'+
' (DRV_PATCH,DRV_DLL,DRV_DLL_TRUE,SECTION_INF,DRV_CLASS,DRV_NAME,DRV_DESCRIPT)'+
' VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7)';
FDQueryDrvPrint.params.ParamByName('p1').AsString:=''+leftstr(Edit1.Text,1500)+'';  //  DRV_PATCH
FDQueryDrvPrint.params.ParamByName('p2').AsString:=''+leftstr(LabeledEdit1.Text,1500)+''; //DRV_DLL
if CheckBox6.Checked then FDQueryDrvPrint.params.ParamByName('p3').AsString:=''+leftstr('1',1)+''  //  DRV_DLL_TRUE
else FDQueryDrvPrint.params.ParamByName('p3').AsString:=''+leftstr('0',1)+'';                       //  DRV_DLL_TRUE
FDQueryDrvPrint.params.ParamByName('p4').AsString:=''+leftstr(Combobox1.Text,200)+'';     //SECTION_INF
FDQueryDrvPrint.params.ParamByName('p5').AsString:=''+leftstr(Edit2.Text,200)+'';    //DRV_CLASS
FDQueryDrvPrint.params.ParamByName('p6').AsString:=''+leftstr(Combobox2.Text,200)+'';    //DRV_NAME
FDQueryDrvPrint.params.ParamByName('p7').AsString:=''+leftstr(Newdescription,500)+''; //DRV_DESCRIPT
FDQueryDrvPrint.ExecSQL;

readDRVPrint;/// чтение списка драйверов из БД
///
end;

end.
