unit LoadListPC;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,Inifiles,
  Vcl.ExtCtrls,Math, Vcl.Menus;

type
  TFormLoadList = class(TForm)
    GroupBox1: TGroupBox;
    ListView1: TListView;
    ComboBox1: TComboBox;
    GroupBox2: TGroupBox;
    Button2: TButton;
    Button3: TButton;
    Button1: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    GroupAD: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4MouseLeave(Sender: TObject);
    procedure Button5MouseLeave(Sender: TObject);
    procedure Button3MouseLeave(Sender: TObject);
    procedure ListView1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure ListView1ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView1Change(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListView1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure ListView1EndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure ListView1StartDrag(Sender: TObject; var DragObject: TDragObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure GroupADSelect(Sender: TObject);
    procedure GroupADKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    function FormGroupCreate (namePC,Usr,Pass:string):bool;
    procedure FormGroupShowAccount(Sender: TObject);
    procedure FormGroupClose(Sender: TObject; var Action: TCloseAction);
    procedure ButOkClickNamePC(Sender: TObject);
    procedure ButOkClickLoadListPC(Sender: TObject);
    procedure ButCancelClickLoadListPC(Sender: TObject);
    function FormPassCreate (namePC,Usr,Pass:string):bool;
  public
    function loadsetForNewRdp(NamePC,Domain,User,Passw:string;AUTConnect,CredSsp:bool;AutLevel:integer):boolean;
    function OpenListRDP(user,pass,dom:string;AutConn,CredSSP:boolean;AutLevel:integer):bool;
  end;

var
  FormLoadList: TFormLoadList;
  ListIni:Tinifile;
  FormGr:Tform;
  ListGr:Tcombobox;
  ButOk:Tbutton;
  ButCan:Tbutton;
  LEditDomain:TLabeledEdit;
  LEditlogin:TLabeledEdit;
  LEditPassw:TLabeledEdit;
  CheckCredSSP:TcheckBox;
  CheckAutoConn:TcheckBox;
  AutLevel:TCombobox;
implementation

{$R *.dfm}
uses urdp,MyDMRDP,unit1,ADWork;

 function MutCheck(s:string): boolean;
begin
  HM:= OpenMutex(MUTEX_ALL_ACCESS, false, pchar(s));  //MForMRPC
  Result:= (HM <> 0);
  if HM = 0 then HM:= CreateMutex(nil, false, pchar(s)); //MForMRPC
end;

function booltoint(z:boolean):integer;
begin
  if z then result:=1
  else result:=0;
end;

function inttobool(z:integer):bool;
begin
  if z=1 then result:=true
  else result:=false;
end;

{createRDP(NamePC,Domain,UserName,Passwd:string;AutoConnect:boolean;
ColorDepth,BitmapPeristence,CachePersistenceActive,BitmapCacheSize,VirtualCache16BppSize,
VirtualCache32BppSize,VirtualCacheSize,DisableCtrlAltDel,
DoubleClickDetect,EnableWindowsKey,MinutesToIdleTimeout,
OverallConnectionTimeout,RdpPort,PerformanceFlags,NetworkConnectionType,
MaxReconnectAttempts,AudioRedirectionMode,AuthenticationLevel,KeyboardHookMode:integer;
SmartSizing,GrabFocusOnConnect,BandwidthDetection,EnableAutoReconnect,
RedirectDrives,RedirectPrinters,RelativeMouseMode
,RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdministerServer,
AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport:boolean):bool;}

function TFormLoadList.loadsetForNewRdp(NamePC,Domain,User,Passw:string;AUTConnect,CredSsp:bool;AutLevel:integer):boolean;
begin
try
FormRDP.createRDP(NamePC, //- имя компа
Domain,               // - домен
User,               //пользователь
Passw,               // пароль
 AUTConnect,                            /// автоконнект
DataMod.FDQueryRead.FieldByName('ColorDepth').AsInteger,
DataMod.FDQueryRead.FieldByName('BitmapPeristence').AsInteger,
DataMod.FDQueryRead.FieldByName('CachePersistenceActive').Asinteger,
DataMod.FDQueryRead.FieldByName('BitmapCacheSize').AsInteger,
DataMod.FDQueryRead.FieldByName('VirtualCache16BppSize').AsInteger,
DataMod.FDQueryRead.FieldByName('VirtualCache32BppSize').AsInteger,
DataMod.FDQueryRead.FieldByName('VirtualCacheSize').AsInteger,
DataMod.FDQueryRead.FieldByName('DisableCtrlAltDel').AsInteger,
DataMod.FDQueryRead.FieldByName('DoubleClickDetect').AsInteger,
DataMod.FDQueryRead.FieldByName('EnableWindowsKey').AsInteger,
DataMod.FDQueryRead.FieldByName('MinutesToIdleTimeout').AsInteger,
DataMod.FDQueryRead.FieldByName('OverallConnectionTimeout').AsInteger,
DataMod.FDQueryRead.FieldByName('RdpPort').AsInteger,
0, //PerformanceFlags /// отключить все графические приблуды
DataMod.FDQueryRead.FieldByName('NetworkConnectionType').AsInteger,
DataMod.FDQueryRead.FieldByName('MaxReconnectAttempts').AsInteger,
DataMod.FDQueryRead.FieldByName('AudioRedirectionMode').AsInteger,
AutLevel,//DataMod.FDQueryRead.FieldByName('AuthenticationLevel').AsInteger,
DataMod.FDQueryRead.FieldByName('KeyboardHookMode').AsInteger,

DataMod.FDQueryRead.FieldByName('SmartSizing').AsBoolean,
DataMod.FDQueryRead.FieldByName('GrabFocusOnConnect').AsBoolean,
true, //BandwidthDetection автоматическое определение качества сети
DataMod.FDQueryRead.FieldByName('EnableAutoReconnect').AsBoolean,
DataMod.FDQueryRead.FieldByName('RedirectDrives').AsBoolean,
DataMod.FDQueryRead.FieldByName('RedirectPrinters').AsBoolean,
DataMod.FDQueryRead.FieldByName('RelativeMouseMode').AsBoolean,
DataMod.FDQueryRead.FieldByName('RedirectClipboard').AsBoolean,
DataMod.FDQueryRead.FieldByName('RedirectDevices').AsBoolean,
DataMod.FDQueryRead.FieldByName('RedirectPorts').AsBoolean,
DataMod.FDQueryRead.FieldByName('ConnectToAdministerServer').AsBoolean,
DataMod.FDQueryRead.FieldByName('AudioCaptureRedirectionMode').AsBoolean,
DataMod.FDQueryRead.FieldByName('EnableSuperPan').AsBoolean,
CredSSp     // отключение или включение EnableCredSspSupport. Отключить для автоматического подключения без запроса
);
result:=true;
Except
  on E: Exception do
  begin
  result:=false;
  FormRDP.memo1.Lines.add(NamePC+' : Ошибка подключения списка компьютеров :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.N1Click(Sender: TObject);
var
i,z:integer;
SourItem,DestItem:string;
begin
try
if ListView1.SelCount<>1 then exit;
SourItem:=ListView1.ItemFocused.Caption;
i:=ListView1.ItemIndex;
if i<>0 then
begin
DestItem:=ListView1.Items[i-1].Caption;
ListView1.Items[i-1].Caption:=SourItem;
ListView1.Items[i].Caption:=DestItem;
end;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.N2Click(Sender: TObject);
var
i,z:integer;
SourItem,DestItem:string;
begin
try
if ListView1.SelCount<>1 then exit;
SourItem:=ListView1.ItemFocused.Caption;
i:=ListView1.ItemIndex;
if i<>ListView1.Items.Count-1 then
begin
DestItem:=ListView1.Items[i+1].Caption;
ListView1.Items[i+1].Caption:=SourItem;
ListView1.Items[i].Caption:=DestItem;
end;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.N3Click(Sender: TObject);
begin
if ListView1.SelCount<>1 then exit;
ListView1.Selected.Delete;
end;

procedure TFormLoadList.N4Click(Sender: TObject);
begin
if ListView1.SelCount<>1 then exit;
ListView1.Selected.EditCaption;
end;

procedure TFormLoadList.ButOkClickNamePC(Sender: TObject);
var
i:integer;
itemyes:bool;
p:Tpoint;
begin
itemyes:=false;
if ListGr.Text<>'' then
begin
for I := 0 to ListView1.Items.Count-1 do
if AnsiUpperCase(ListView1.Items[i].Caption)=AnsiUpperCase(ListGr.Text) then /// поиск в списке
begin
 ListView1.Items[i].Selected:=true;
 ListView1.ItemIndex:=i;
 ListView1.ItemFocused:=ListView1.Items[i];
 p := ListView1.Items.Item[i].Position;
 ListView1.Scroll(P.X,P.Y);
 itemyes:=true;
 break;
end;
if not itemyes then
 ListView1.Items.Add.Caption:=ListGr.Text;
end;
FormGr.Close;
end;

procedure TFormLoadList.ButOkClickLoadListPC(Sender: TObject);
var        //// кнопка OK на форме ввода домена логина и пароля
i:integer;
begin
OpenListRDP(LEditlogin.text,LEditPassw.text,LeditDomain.text
,CheckAutoConn.checked,CheckCredSSP.checked,AutLevel.ItemIndex);
FormGr.Close;
end;

procedure TFormLoadList.ButCancelClickLoadListPC(Sender: TObject);
begin        /// закрытие формы
FormGr.Close;
end;


procedure TFormLoadList.FormGroupClose(Sender: TObject; var Action: TCloseAction);
begin
FormGr.Release; /// уничтожение формы
end;

procedure TFormLoadList.CheckBox1Click(Sender: TObject);
begin
{if (sender as TCheckBox).Name='ckCredSSP' then /// это для одного чекбокса
begin
LEditDomain.Enabled:=not(CheckCredSSP.Checked);
LEditlogin.Enabled:=not(CheckCredSSP.Checked);
LEditPassw.Enabled:=not(CheckCredSSP.Checked);
if CheckCredSSP.Checked then AutLevel.ItemIndex:=2; /// предупреждать
CheckAutoConn.Checked:=not CheckCredSSP.Checked;
end;
if (sender as TCheckBox).Name='ckAutoConn' then // это для второго чекбокса
begin
 CheckCredSSP.Checked:= not CheckAutoConn.Checked;
 LEditDomain.Enabled:= CheckAutoConn.Checked;
 LEditlogin.Enabled:= CheckAutoConn.Checked;
 LEditPassw.Enabled:= CheckAutoConn.Checked;
 if CheckAutoConn.Checked then AutLevel.ItemIndex:=0   // подключаться без предупреждения
end; }
end;

procedure TFormLoadList.FormGroupShowAccount(Sender: TObject);
var                      ///// загрузка списков пользователей  при загрузке формы
i:integer;
begin
try
for I := 0 to FormRDP.ComboKomp.Items.Count-1 do
ListGr.Items.Add(FormRDP.ComboKomp.Items[i]);
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка загрузки списка компьютеров :'+E.Message);
  end;
end;
end;

function TFormLoadList.FormPassCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormLoadList);
FormGr.Name:='FormPass';
FormGr.Width:=225;
FormGr.Height:=235;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:='Учетные данные';
FormGr.OnClose:=FormGroupClose;
FormGr.Position:=poMainFormCenter;

AutLevel:=TCombobox.Create(FormGr);
AutLevel.Parent:=FormGr;
AutLevel.Name:='AutLevel';
AutLevel.Items.Add('Подключаться без предупреждения');
AutLevel.Items.Add('Не соединять');
AutLevel.Items.Add('Предупреждать');
AutLevel.ItemIndex:=2;
AutLevel.Left:=5;
AutLevel.Top:=1;
AutLevel.Width:=200;
AutLevel.TabOrder:=0;

LEditDomain:=TLabeledEdit.Create(FormGr);
LEditDomain.Parent:=FormGr;
LEditDomain.Name:='EditDomain';
LEditDomain.Text:=FormRDP.LabeledEdit3.Text;
LEditDomain.Left:=5;
LEditDomain.Top:=40;
LEditDomain.Width:=200;
LEditDomain.EditLabel.Caption:='Домен или localhost';
LEditDomain.TabOrder:=0;

LEditlogin:=TLabeledEdit.Create(FormGr);
LEditlogin.Parent:=FormGr;
LEditlogin.Name:='Editlogin';
LEditlogin.Text:=FormRDP.LabeledEdit1.Text;
LEditlogin.Left:=5;
LEditlogin.Top:=74;
LEditlogin.Width:=200;
LEditlogin.EditLabel.Caption:='Пользователь';
LEditlogin.TabOrder:=1;

LEditPassw:=TLabeledEdit.Create(FormGr);
LEditPassw.Parent:=FormGr;
LEditPassw.Name:='EditPassw';
LEditPassw.Text:=FormRDP.LabeledEdit2.Text;
LEditPassw.Left:=5;
LEditPassw.Top:=110;
LEditPassw.Width:=200;
LEditPassw.EditLabel.Caption:='Пароль';
LEditPassw.TabOrder:=2;
LEditPassw.PasswordChar:=#7;

CheckCredSSP:=TcheckBox.create(FormGr);
CheckCredSSP.Parent:=FormGr;
CheckCredSSP.Name:='ckCredSSP';
CheckCredSSP.Left:=5;
CheckCredSSP.Top:=133;
CheckCredSSP.Width:=200;
CheckCredSSP.caption:='Проверка подлинности CredSSP';
CheckCredSSP.taborder:=3;
CheckCredSSP.OnClick:=CheckBox1Click;
CheckCredSSP.Checked:=true;

CheckAutoConn:=TcheckBox.create(FormGr);
CheckAutoConn.Parent:=FormGr;
CheckAutoConn.Name:='ckAutoConn';
CheckAutoConn.Left:=5;
CheckAutoConn.Top:=150;
CheckAutoConn.Width:=200;
CheckAutoConn.caption:='Автоматически подключать';
CheckAutoConn.taborder:=3;
CheckAutoConn.OnClick:= CheckBox1Click;
CheckAutoConn.Checked:=true; //// устанавливаем чекбокс на автосоединение

ButCan:=TButton.Create(FormGr);
ButCan.Parent:=FormGr;
ButCan.Name:='ButCan';
ButCan.Caption:='Отмена';
ButCan.Top:=174;
ButCan.Left:=50;
ButCan.OnClick:=ButCancelClickLoadListPC;
ButCan.TabOrder:=5;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='ButOk';
ButOk.Caption:='Ок';
ButOk.Top:=174;
ButOk.Left:=130;
ButOk.OnClick:=ButOkClickLoadListPC;
ButOk.TabOrder:=4;



FormGr.Show;
result:=true;
except
  on E:Exception do
    begin
    ShowMessage(e.Message);
    result:=false;
    end;
  end;
end;


function TFormLoadList.FormGroupCreate (namePC,Usr,Pass:string):bool;
begin
try
FormGr:=TForm.Create(FormLoadList);
FormGr.Name:='FormGroup';
FormGr.Width:=323;
FormGr.Height:=97;
FormGr.BorderStyle:=bsDialog;
FormGr.OnShow:=FormGroupShowAccount;
FormGr.Caption:='Список компьютеров';
FormGr.OnClose:=FormGroupClose;
FormGr.Position:=poMainFormCenter;
ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:='ListGr';
ListGr.Text:='';
ListGr.Left:=8;
ListGr.Top:=8;
ListGr.Width:=298;
ListGr.DropDownCount:=15;
ListGr.CharCase:= ecUpperCase;
ListGr.AutoDropDown:=true;
ListGr.TabOrder:=0;
ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='ButOk';
ButOk.Caption:='Ок';
ButOk.Top:=35;
ButOk.Left:=231;
ButOk.OnClick:=ButOkClickNamePC;
ButOk.TabOrder:=1;
FormGr.Show;
result:=true;
except
  on E:Exception do
    begin
    ShowMessage(e.Message);
    result:=false;
    end;
  end;
end;


function TFormLoadList.OpenListRDP(user,pass,dom:string;AutConn,CredSSP:boolean;AutLevel:integer):bool;
var            /// подключение списка компьютеров
i:integer;
def:string;
BEGIN
try
def:='SETDEFAULT';

for I := 0 to ListView1.Items.Count-1  do
if (ListView1.Items[i].Checked) and (FormRDP.PageControl1.PageCount<CountTab) then /// если отмечены чекбоксы
  Begin
    DataMod.FDQueryRead.SQL.Clear;
    DataMod.FDQueryRead.SQL.Text:='SELECT * FROM RDP_SET WHERE PC_NAME='''+ListView1.Items[i].Caption+'''';
    DataMod.FDQueryRead.Open;
 if DataMod.FDQueryRead.FieldByName('PC_NAME').Value<>null then
   begin
    loadsetForNewRdp(ListView1.Items[i].Caption,dom,user,pass,AutConn,CredSSP,AutLevel);
    DataMod.FDQueryRead.Close;
   end
 else
   begin
   DataMod.FDQueryRead.Close;
   DataMod.FDQueryRead.SQL.Clear;
   DataMod.FDQueryRead.SQL.Text:='SELECT * FROM RDP_SET WHERE PC_NAME='''+def+'''';
   DataMod.FDQueryRead.Open;
   if DataMod.FDQueryRead.FieldByName('PC_NAME').Value<>null then
    begin
    loadsetForNewRdp(ListView1.Items[i].Caption,dom,user,pass,AutConn,CredSSP,AutLevel);
    DataMod.FDQueryRead.Close;
    end
   else
    begin
     ShowMessage('Не найдены настройки по умолчанию. '
     +#10#13+' Установите настройки по умолчанию и повторите попытку.');
     DataMod.FDQueryRead.Close;
     break;
    end;
   end;
  End;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка подключения списка компьютеров :'+E.Message);
  end;
end;
END;

procedure TFormLoadList.Timer1Timer(Sender: TObject);
var
TrLic:Tthread;
begin
TrLic:=unit1.Lcheck.Create(true);
TrLic.FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
TrLic.Priority:=tpLower;
TrLic.start;
if sender is Ttimer then
(sender as Ttimer).Enabled:=false;
end;

procedure TFormLoadList.Button1Click(Sender: TObject);
var
i:integer;
check:bool;
begin
check:=false;
for I := 0 to ListView1.Items.Count-1 do
if ListView1.Items[i].Checked then
begin
check:=true;
break;
end;
if not check then
begin
  ShowMessage('Не выделен список компьютеров.');
  exit;
end;
FormPassCreate('','',''); /// Создаем форму для подлючения
close;
end;

procedure TFormLoadList.Button2Click(Sender: TObject);
begin //// удаление компа из списка
try
if ListView1.SelCount=1 then
begin
if ListIni.ValueExists(ComboBox1.Text,listview1.Selected.Caption)then
begin
ListIni.DeleteKey(ComboBox1.Text,listview1.Selected.Caption);
end;
listview1.Selected.Delete;
end;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка удаления компьютера :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.Button3Click(Sender: TObject);
var
namepc:string;
begin
FormGroupCreate('','','');
{if not InputQuery('Новый компьютер в список '+ComboBox1.Text, 'Имя ПК', namepc)
    then exit;
if namepc='' then Exit;
ListView1.Items.Add.Caption:=namepc;}
end;

procedure TFormLoadList.Button3MouseLeave(Sender: TObject);
begin
Button3.Hint:='Добавить новый компьютер в список - '+ComboBox1.Text;
end;

procedure TFormLoadList.Button4Click(Sender: TObject);
begin  /// удаление списка
try
if listini.SectionExists(ComboBox1.Text) then
begin
ListIni.EraseSection(ComboBox1.Text);
end;
ListView1.Clear;
ComboBox1.Items.Delete(ComboBox1.ItemIndex);
if ComboBox1.Items.Count>0 then Combobox1.ItemIndex:=0
else ComboBox1.Text:='';
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка удаления списка :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.Button4MouseLeave(Sender: TObject);
begin
Button4.Hint:='Удалить список компьютеров - '+ComboBox1.Text;
end;

procedure TFormLoadList.Button5Click(Sender: TObject);
var      //// сохранить список
i:integer;
begin
try
if listini.SectionExists(ComboBox1.Text) then
begin
ListIni.EraseSection(ComboBox1.Text);
end;

for I :=0 to listview1.Items.Count-1 do
 begin
 ListIni.WriteString(ComboBox1.Text,Listview1.Items[i].Caption,inttostr(i));
 end;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка сохранения списка компьютеров :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.Button5MouseLeave(Sender: TObject);
begin
Button5.Hint:='Сохранить список компьютеров - '+ComboBox1.Text
end;

procedure TFormLoadList.Button6Click(Sender: TObject); ///// добавить новый список
var
s:string;
begin
try
if not InputQuery('Новый список компьютеров', 'Имя списка:', s)
    then exit;
ComboBox1.Items.Add(s);
ComboBox1.Text:=s;
ListView1.Clear; /// оишаем список
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка добавления списка компьютеров :'+E.Message);
  end;
end;
end;



procedure TFormLoadList.Button7Click(Sender: TObject);
begin
Close;
end;

procedure TFormLoadList.ComboBox1Select(Sender: TObject);
var
listValue:Tstringlist;
i:integer;
begin
try
ListView1.Clear;
listValue:=TStringList.Create;
Listini.ReadSection(ComboBox1.Items[ComboBox1.ItemIndex],listValue);
for I := 0 to listValue.Count-1 do
with ListView1.Items.Add do
begin
  caption:=listValue[i];
end;
listValue.Free;
ListView1.Column[0].ImageIndex:=1;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка чтения списка компьютеров :'+E.Message);
  end;
end;
end;

procedure TFormLoadList.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 ListIni.Free;
end;

procedure TFormLoadList.FormCreate(Sender: TObject);
var
TimerOne:Ttimer;
begin
TimerOne:=Ttimer.Create(FormLoadList);
TimerOne.Name:='Timr';
TimerOne.OnTimer:=Timer1Timer;
TimerOne.Interval:=TimerOne.Interval*(randomrange(100,180)) ;
TimerOne.Enabled:=true;

GroupAD.clear;
GetCurrentComputerName; // информация о имени компа и где он находится (т.е. домен)
if GetDomainController(CurrentDomainName)<>'' then // передаем имя домена и узнаем контроллер домена
GroupAD.items:= EnumAllGroups;
//EnumAllGroups; //gjkexftv список групп домена
end;

Procedure TFormLoadList.FormShow(Sender: TObject);
var
listSection:Tstringlist;
i:integer;
begin
try
ListView1.Clear;
ComboBox1.Clear;
ListIni:=TIniFile.Create(ExtractFileDir(Application.ExeName)+'\RDPlist.dat');
listSection:=TStringList.Create;
ListIni.ReadSections(listSection);
for I := 0 to listSection.Count-1 do ComboBox1.Items.Add(listSection[i]);
listSection.Free;
if ComboBox1.Items.Count<>0 then
begin
 combobox1.ItemIndex:=0;
 ComboBox1Select(ComboBox1); /// выбор списка
end;
ListView1.Column[0].ImageIndex:=1;
Except
  on E: Exception do
  begin
  FormRDP.memo1.Lines.add('Ошибка при загрузке формы:'+E.Message);
  end;
end;
end;

procedure TFormLoadList.GroupADKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
GroupAD.DroppedDown:=true;
end;

procedure TFormLoadList.GroupADSelect(Sender: TObject);
var
ListPC:TstringList;
i:integer;
begin
if GroupAD.items.Count=0 then exit;

if GroupAD.Items[GroupAD.ItemIndex]<>'' then
  begin
  ListPC:=TstringList.create;
  ListPC:=GetAllGroupUsers(GroupAD.Items[GroupAD.ItemIndex]);
  ListView1.Clear;
  for I := 0 to ListPC.Count-1 do
    ListView1.Items.Add.Caption:= ListPC[i];
  ListPC.Free;
  end;

end;

procedure TFormLoadList.ListView1Change(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
//Item.ImageIndex:=0;
end;

procedure TFormLoadList.ListView1ColumnClick(Sender: TObject;
  Column: TListColumn);
  var
  i:integer;
begin

if ListView1.Columns[0].ImageIndex=1 then
begin
for I := 0 to ListView1.Items.Count-1 do
 ListView1.Items[i].Checked:=true;
ListView1.Columns[0].ImageIndex:=2;
exit;
end;

if ListView1.Columns[0].ImageIndex=2 then
begin
for I := 0 to ListView1.Items.Count-1 do
 ListView1.Items[i].Checked:=false;
ListView1.Columns[0].ImageIndex:=1;
exit;
end;


end;

procedure TFormLoadList.ListView1DragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
 with Sender as TListView do
    Accept := (Sender = ListView1) and (GetItemAt(X, Y) <> Selected);
end;

procedure TFormLoadList.ListView1EndDrag(Sender, Target: TObject; X,
  Y: Integer);
var
  DropItem, CurrentItem: TListItem;
  i, n: Integer;
 // st, st1, st2, st3: string;
begin
  if Sender = Target then
    with TListView(Sender) do
    begin
      DropItem := GetItemAt(X, Y);
      if DropItem = nil then
        Exit;
      CurrentItem := TListItem.Create(Items);
      CurrentItem.Assign(Selected);
      Selected.Delete;
      n := DropItem.Index;
      AddItem('', nil);
      for i := Items.Count - 1 downto DropItem.Index + 1 do
      begin
        Items.item[i].Assign(Items.item[i - 1]);
      end;
      Items.item[n].Assign(CurrentItem);
    end;

end;

procedure TFormLoadList.ListView1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key=113 then
begin
if ListView1.SelCount=1 then
ListView1.Selected.EditCaption;
end;

end;

type // Нужно подключить CommCtrl к списку Uses
  TLVDragControlObjectEx = class(TDragControlObjectEx)
  protected
    function GetDragImages: TDragImageList; override;
  end;

function TLVDragControlObjectEx.GetDragImages: TDragImageList;
var
  Bmp: TBitmap;
  R: TRect;
begin
  Bmp := TBitmap.Create;
  Bmp.Canvas.Brush.Color := clSkyBlue;
  R := TListView(Control).Selected.DisplayRect(drBounds);
  Bmp.SetSize(R.Right - R.Left, R.Bottom - R.Top);
  Bmp.Canvas.Font := TListView(Control).Font;
  Bmp.Canvas.TextOut(0, 0, TListView(Control).Selected.Caption);

  Result := TDragImageList.Create(Control);
  Result.Width := Bmp.Width;
  Result.Height := Bmp.Height;
 // ImageList.EndDrag;
  Result.SetDragImage(Result.Add(Bmp, nil), 0, 0);
  Bmp.Free;
end;

procedure TFormLoadList.ListView1StartDrag(Sender: TObject;
  var DragObject: TDragObject);
begin
 DragObject := TLVDragControlObjectEx.Create(ListView1);
 DragObject.AlwaysShowDragImages := True;
end;

end.
