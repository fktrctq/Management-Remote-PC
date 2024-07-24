unit RestoreWindows;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ExtCtrls, ActiveX,ComObj,
  Vcl.StdCtrls;

type

  TPropRestore = record
    NamePC       :String;
    User         :String;
    Pass         :String;
    Description  :string;
    RestorePointType,EventType:integer;
  end;

  TPropRunRestore = record
    NamePC       :String;
    User         :String;
    Pass         :String;
    NumPoint     :integer;
    Reboot       :boolean;
  end;

  TFormRestoreWin = class(TForm)
    Panel1: TPanel;
    ListView1: TListView;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button5: TButton;
    Button6: TButton;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    FSWbemLocatorRes   : OLEVariant;
    FWMIServiceRes     : OLEVariant;
    FormKey:Tform;
    EditNameKey:TlabeledEdit;
    ComboType:TcomboBox;
    LV:TlistView;

    ButtonOk,ButtonNo,ButtonOK2:Tbutton;
    function LoadPointsRestoreWin(ListView:TListView):boolean;
    Function CreateNewPointRestore(Description:string;RestorePointType,EventType:integer):integer; // создаие точки восстановления
    Function RestoreSelectedPoint(SequenceNumber:integer):integer; // запуск восстановления
    function ConnectWMI(NamePC,User,Pass:string):boolean;
    procedure createformNewPoint; // создание формы для создания точки
    Function EnableDisableRestore(descript:string):boolean; // создание формы для вкоючения или отключения восстановления
    procedure LoadDiskDrive(sender:Tobject); // загрузка списка дисков
    procedure ButtonDisableRestore(sender:Tobject);  // отключить мониторинг на выбранном диске
    procedure ButtonEnableRestore(sender:Tobject);  // Включить мониторинг на выбранном диске
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonNoClose(Sender: TObject);
    procedure ButtonAddPoint(Sender: TObject);
    function LoadConfig:boolean;
  public
    { Public declarations }
  end;

var
  FormRestoreWin: TFormRestoreWin;
  ParamForRestore:TPropRestore;
  ParamForRunRestorePoint:TPropRunRestore;

implementation
uses umain;

ThreadVar
  ParamRestore: ^TPropRestore; // создание точки восстановления
  ParamRunRestore:^TPropRunRestore; // запуск восстановления из точки
{$R *.dfm}

function RebootPC(namePC,user,pass:string):boolean;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown:integer;
begin
try
   OleInitialize(nil);
   FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
   FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128);
   FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
   oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   if oEnum.Next(1, FWbemObject, iValue) = 0 then
   begin
   ResShutdown:=FWbemObject.Win32Shutdown(2);
   if ResShutdown<>0 then ResShutdown:=FWbemObject.Win32Shutdown(6);
   frmDomainInfo.memo1.Lines.Add(namePC+': Перезагрузка компьютера - '+SysErrorMessage(ResShutdown));
   result:=true;
   end;
  FWbemObject:=Unassigned;
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
except
  on E:Exception do
  begin
  result:=false;
  frmDomainInfo.memo1.Lines.Add(NamePC+': Ошибка перезагрузки - '+E.Message);
  end;
end;

end;

procedure TFormRestoreWin.Button1Click(Sender: TObject);
begin
createformNewPoint;
end;

procedure TFormRestoreWin.Button2Click(Sender: TObject); //обновить
begin
LoadPointsRestoreWin(ListView1);
//LoadConfig;
end;

procedure TFormRestoreWin.Button3Click(Sender: TObject);
begin
EnableDisableRestore('Настройка защиты на дисках');
end;


Function TFormRestoreWin.RestoreSelectedPoint(SequenceNumber:integer):integer; // запуск восстановления системы.
var                     // вроде как надо запустить перезагрузку после этой операции
 FWbemObjectSet  : OLEVariant;
 res:integer;
 begin
try
 FWbemObjectSet:= FWMIServiceRes.Get('SystemRestore');
 res:=FWbemObjectSet.Restore(SequenceNumber);
 result:=res;
 except
 on e:EInOutError do result:=e.ErrorCode;
 end;
 VariantClear(FWbemObjectSet);
 end;

Function RestoreSelectedPointThread(ParamRR:pointer):integer; // запуск восстановления системы.
var                     // вроде как надо запустить перезагрузку после этой операции
 FWbemObjectSet  : OLEVariant;
 FSWbemLocator   : OLEVariant;
 FWMIService     : OLEVariant;
 res:integer;
 NamePC:String;
 begin
try
OleInitialize(nil);
 ParamRunRestore:=ParamRR;
 NamePC:=string(ParamRunRestore.NamePC);
 frmDomainInfo.Memo1.Lines.add(NamePC+': Запуск восстановления системы');
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\DEFAULT', ParamRunRestore.User, ParamRunRestore.Pass,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6;
 FWbemObjectSet:= FWMIService.Get('SystemRestore');
 res:=FWbemObjectSet.Restore(ParamRunRestore.NumPoint);
 if ParamRunRestore.Reboot then
   begin
   frmDomainInfo.Memo1.Lines.add(NamePC+':Запуск восстановление системы - '+SysErrorMessage(res));
   if res=0 then RebootPC(NamePC,ParamRunRestore.User,ParamRunRestore.pass);
   end
 else
   begin
   frmDomainInfo.Memo1.Lines.add(NamePC+':Запуск восстановление системы - '+SysErrorMessage(res));
   frmDomainInfo.Memo1.Lines.add(NamePC+':Для продолжения требуется перезагрузка системы');
   end;
 except
 on E:Exception do frmDomainInfo.Memo1.Lines.add(NamePC+': Ошибка запуска восстановления системы '+e.Message);
 end;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize;
 NetApiBufferFree(ParamRR); /// очищаем память
 EndThread(result); /// конец потока
 end;


procedure TFormRestoreWin.Button5Click(Sender: TObject);
var
i,z,res:integer;
NewTread:integer;
treadID :LongWord;
rb:boolean;
begin
rb:=false;
if listview1.SelCount<>1 then
begin
 ShowMessage('Выберите точку для восстановления');
 exit;
end;

{if TryStrToInt(listview1.Selected.Caption,i) then
begin
z:=MessageDlg('Запустить процесс восстановления до выбранной контрольной точки? '+#10#13+Listview1.Selected.SubItems[0]+' - '+Listview1.Selected.SubItems[1], mtConfirmation,[mbYes,mbCancel],0);
if z=IDCancel then   exit;
res:=RestoreSelectedPoint(i);
if res=0 then
begin
z:=MessageDlg('Для продолжения восстановления необходимо перезагрузить компьютер.', mtConfirmation,[mbYes,mbCancel],0);
if z=IDYes then  frmDomainInfo.rebuutSelectPC;
end
else ShowMessage(SysErrorMessage(res));
end; }

if TryStrToInt(listview1.Selected.Caption,i) then
begin
z:=MessageDlg('Запустить процесс восстановления до выбранной контрольной точки? '+#10#13+Listview1.Selected.SubItems[0]+' - '+Listview1.Selected.SubItems[1], mtConfirmation,[mbYes,mbCancel],0);
if z=IDCancel then   exit;
z:=MessageDlg('Для восстановления требуется перезагрузка компьютера', mtConfirmation,[mbYes,mbCancel],0);
if z=IDYes then  rb:=true;
ParamForRunRestorePoint.NamePC:=frmDomainInfo.ComboBox2.Text;
ParamForRunRestorePoint.User:=frmDomainInfo.LabeledEdit1.text;
ParamForRunRestorePoint.Pass:=frmDomainInfo.LabeledEdit2.text;
ParamForRunRestorePoint.NumPoint:=i;
ParamForRunRestorePoint.Reboot:=rb;
NewTread:=BeginThread(nil,0,addr(RestoreSelectedPointThread),Addr(ParamForRunRestorePoint),0,treadID); /// поток для заполнения списка событий
CloseHandle(NewTread);
end;

end;

procedure TFormRestoreWin.Button6Click(Sender: TObject); // Проверка статуса последнего восстановления
var
 FWbemObjectSet  : OLEVariant;
 res:integer;
begin
try
 FWbemObjectSet:= FWMIServiceRes.Get('SystemRestore');
 res:=FWbemObjectSet.GetLastRestoreStatus();
 case res of
 0:showmessage('Последнее восстановление не удалось.');
 1:showmessage('Последнее восстановление выполнено успешно.');
 2:showmessage('Последнее восстановление было прервано.');
 else ShowMessage(SysErrorMessage(res));
 end;
except
 on e:Exception do Showmessage(e.Message);
 end;
 VariantClear(FWbemObjectSet);
end;

function TFormRestoreWin.ConnectWMI(NamePC,User,Pass:string):boolean;
begin
 try
    FSWbemLocatorRes := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIServiceRes   := FSWbemLocatorRes.ConnectServer(NamePC, 'root\DEFAULT', User, Pass,'','',128);
    FWMIServiceRes.Security_.impersonationlevel:=3;
    FWMIServiceRes.Security_.authenticationLevel := 6;
    Result:=true;
 except
   on E:Exception do
     begin
     frmDomainInfo.Memo1.Lines.Add('Ошибка '+e.Message);
     result:=false;
     end;
 end;
end;

function DescriptionPoint:string;
var
Description:string;
begin
while Description='' do
Begin
Description:=DateTimeToStr(now);
if not InputQuery('Описание для точки восстановления', 'Описание:', Description)
then begin Description:=''; exit; end;
end;
result:=Description;
end;

function  WbemTimeToDateTime(vDate : OleVariant) : string;  // перевод даты и времени
var
  FWbemDateObj  : OleVariant;
  TDT:TdateTime;
begin;
  FWbemDateObj  := CreateOleObject('WbemScripting.SWbemDateTime');
  try
  FWbemDateObj.Value:=vDate;
  TDT:=FWbemDateObj.GetVarDate;
  result:=DateTimeToStr(TDT);
  finally
  VariantClear(FWbemDateObj);
  end;
end;

procedure TFormRestoreWin.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;
procedure TFormRestoreWin.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;

procedure TFormRestoreWin.createformNewPoint;  // создание формы для новвой точки восстановления
begin
FormKey:=TForm.Create(FormRestoreWin);
FormKey.Caption:='Создание точки восстановления';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=365;
FormKey.Height:=135;
FormKey.OnClose:= FormKeyClose;
///////////////////////////////
ComboType:=TComboBox.Create(FormKey);
ComboType.Parent:=FormKey;
ComboType.Items.Add('Установлено приложение');
ComboType.Items.Add('Приложение было удалено');
ComboType.Items.Add('Установлен драйвер устройства');
ComboType.Items.Add('Обновление приложения');
ComboType.ItemIndex:=0;
ComboType.Style:=csOwnerDrawFixed;
ComboType.Top:=5;
ComboType.Left:=5;
ComboType.Width:=340;
////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='Описание точки восстановления';
EditNameKey.Top:=45;
EditNameKey.Left:=5;
EditNameKey.Width:=340;
EditNameKey.Text:='Системная точка';
EditNameKey.TabOrder:=0;
//////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Ок';
ButtonOk.Top:=70;
ButtonOk.Left:=170;
ButtonOk.OnClick:=ButtonAddPoint;
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Отмена';
ButtonNo.Top:=70;
ButtonNo.Left:=270;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
FormKey.ShowModal;
end;

procedure TFormRestoreWin.ButtonEnableRestore(sender:Tobject);   // включение мониторинга диска
var
 FWbemObjectSet  : OLEVariant;
 res:integer;
 begin
try
if LV.SelCount<>1 then
begin
  ShowMessage('Выберите диск');
  exit;
end;
 FWbemObjectSet:= FWMIServiceRes.Get('SystemRestore');
 res:=FWbemObjectSet.Enable(LV.Selected.Caption);
 ShowMessage(SysErrorMessage(res));
except
on e:Exception do showmessage(e.message);
end;
VariantClear(FWbemObjectSet);
end;


procedure TFormRestoreWin.ButtonDisableRestore(sender:Tobject);  // выклчение мониторинга диска
var
 FWbemObjectSet  : OLEVariant;
 res:integer;
 begin
try
if LV.SelCount<>1 then
begin
  ShowMessage('Выберите диск');
  exit;
end;
 FWbemObjectSet:= FWMIServiceRes.Get('SystemRestore');
 res:=FWbemObjectSet.Disable(LV.Selected.Caption);
 ShowMessage(SysErrorMessage(res));
except
on e:Exception do showmessage(e.message);
end;
VariantClear(FWbemObjectSet);
end;



procedure TFormRestoreWin.LoadDiskDrive(sender:Tobject); // загрузка списка дисков
var
FSWbemLocator   : OLEVariant;
FWMIService     : OLEVariant;
FWbemObject     : OLEVariant;
oEnum           : IEnumvariant;
FWbemObjectSet  : OLEVariant;
iValue        : LongWord;
sizeHDD,FreeSpace,CapDrive:string;
z,x:int64;
const
MyTypeDrv: array [0..6] of string = ('Неизвестный','Не системный','Съемный диск','Локальный диск','Сетевой диск','CD/DVD','RAM диск');
begin
try
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption,Capacity,FreeSpace,DriveType FROM Win32_Volume','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
begin
if (FWbemObject.Caption<>Null) then  CapDrive:=vartostr(FWbemObject.Caption);
if (CapDrive<>'') and (integer(FWbemObject.DriveType)=3) then
  begin
  SizeHDD:=vartostr(FWbemObject.Capacity);
  z:=((strtoint64(SizeHDD)) div 1024)div 1024;
  if z<=1024 then sizeHDD:=inttostr(z)+' МБ'
  else sizeHDD:=inttostr(z div 1024)+' ГБ';
  FreeSpace:=vartostr(FWbemObject.FreeSpace);
  x:=((strtoint64(FreeSpace)) div 1024)div 1024;
  if x<=1024 then FreeSpace:=inttostr(x)+' МБ'
  else FreeSpace:=inttostr(x div 1024)+' ГБ';

  if (CapDrive<>'')and (Length(CapDrive)<4) then
  with lv.Items.Add do
  begin
    Caption:= CapDrive;
    SubItems.Add(SizeHDD);
    SubItems.Add(FreeSpace);
  end;

  end;

VariantClear(FWbemObject);
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator)
except on E: Exception do
  begin
  ShowMessage(e.Message);
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator)
  end;
end;
end;

Function TFormRestoreWin.EnableDisableRestore(descript:string):boolean;
begin
FormKey:=TForm.Create(FormRestoreWin);
FormKey.Caption:=descript;
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=300;
FormKey.Height:=230;
FormKey.OnClose:= FormKeyClose;
FormKey.OnShow:=LoadDiskDrive;
///////////////////////////////
LV:=TListView.Create(FormKey);
LV.Parent:=FormKey;
LV.Top:=5;
LV.Left:=5;
LV.Width:=290;
LV.Height:=150;
lv.MultiSelect:=false;
lv.GridLines:=true;
lv.HideSelection:=false;
lv.ReadOnly:=true;
lv.RowSelect:=true;
lv.ViewStyle:=vsReport;
with lv.Columns.Add do
begin
  Caption:='Диск';
  Width:=75;
end;
with lv.Columns.Add do
begin
  Caption:='Объем';
  Width:=95;
end;
with lv.Columns.Add do
begin
  Caption:='Свободно';
  Width:=95;
end;
//////////////////
ButtonOk2:=Tbutton.Create(FormKey);
ButtonOk2.Parent:=FormKey;
ButtonOk2.Caption:='Отключить';
ButtonOk2.Top:=160;
ButtonOk2.Left:=15;
ButtonOk2.OnClick:=ButtonDisableRestore;
ButtonOk2.TabOrder:=4;
/// ///////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='Включить';
ButtonOk.Top:=160;
ButtonOk.Left:=105;
ButtonOk.OnClick:=ButtonEnableRestore;
ButtonOk.TabOrder:=3;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='Закрыть';
ButtonNo.Top:=160;
ButtonNo.Left:=200;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=4;
FormKey.ShowModal;
end;

Function TFormRestoreWin.CreateNewPointRestore(Description:string;RestorePointType,EventType:integer):integer;
var
 FWbemObjectSet  : OLEVariant;
 res:integer;
 begin
try
 FWbemObjectSet:= FWMIServiceRes.Get('SystemRestore');
 res:=FWbemObjectSet.CreateRestorePoint(Description,RestorePointType,EventType);
 result:=res;
 except
 on e:EInOutError do result:=e.ErrorCode;
 end;
 VariantClear(FWbemObjectSet);
 end;

 Function CreateNewPointRestoreThread(ParamRP:pointer):integer;
var
 FWbemObjectSet  : OLEVariant;
 FSWbemLocator   : OLEVariant;
 FWMIService     : OLEVariant;
 res:integer;
 NamePC:String;
 begin
try
 OleInitialize(nil);
 ParamRestore:=ParamRP;
 NamePC:=string(ParamRestore.NamePC);
 frmDomainInfo.Memo1.Lines.add(NamePC+': Запуск создания точки восстановления');
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\DEFAULT', ParamRestore.User, ParamRestore.Pass,'','',128);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6;
 FWbemObjectSet:= FWMIService.Get('SystemRestore');
 res:=FWbemObjectSet.CreateRestorePoint(ParamRestore.Description,ParamRestore.RestorePointType,ParamRestore.EventType);
 frmDomainInfo.Memo1.Lines.add(NamePC+':Запуск создания точки восстановления - '+SysErrorMessage(res));
 except
 on E:Exception do frmDomainInfo.Memo1.Lines.add(NamePC+': Ошибка запуска создания точки восстановления '+e.Message);
 end;
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize;
 NetApiBufferFree(ParamRP); /// очищаем память
 EndThread(result); /// конец потока
 end;

procedure TFormRestoreWin.ButtonAddPoint(Sender: TObject); // создать точку
var
pointType,NewTread:integer;
res:integer;
treadID :LongWord;
begin
{APPLICATION_INSTALL 0  Установлено приложение.
APPLICATION_UNINSTALL 1  Приложение было удалено.
DEVICE_DRIVER_INSTALL  10 Установлен драйвер устройства.
MODIFY_SETTINGS  12  В приложение были добавлены или удалены функции.
CANCELLED_OPERATION 13  Приложению необходимо удалить созданную точку восстановления. Например, приложение будет использовать этот флаг, когда пользователь отменяет установку.}
{BEGIN_NESTED_SYSTEM_CHANGE 102
BEGIN_SYSTEM_CHANGE 100
END_NESTED_SYSTEM_CHANGE 103
END_SYSTEM_CHANGE 101}
if ComboType.Text='Установлено приложение' then pointType:=0;
if ComboType.Text='Приложение было удалено' then pointType:=1;
if ComboType.Text='Установлен драйвер устройства' then pointType:=10;
if ComboType.Text='Обновление приложения' then pointType:=12;

//res:=CreateNewPointRestore(EditNameKey.text,pointType,100);  // функция запуска создания точки восстановления
//ShowMessage(SysErrorMessage(res));
ParamForRestore.NamePC:=frmDomainInfo.ComboBox2.Text;
ParamForRestore.User:=frmDomainInfo.LabeledEdit1.text;
ParamForRestore.Pass:=frmDomainInfo.LabeledEdit2.text;
ParamForRestore.Description:=EditNameKey.text;
ParamForRestore.RestorePointType:=pointType;
ParamForRestore.EventType:=100;
NewTread:=BeginThread(nil,0,addr(CreateNewPointRestoreThread),Addr(ParamForRestore),0,treadID); /// поток для заполнения списка событий

CloseHandle(NewTread);
FormKey.Close;
end;

procedure TFormRestoreWin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VariantClear(FWMIServiceRes);
VariantClear(FSWbemLocatorRes);
end;

procedure TFormRestoreWin.FormShow(Sender: TObject);
begin
if ConnectWMI(frmDomainInfo.ComboBox2.Text,
              frmDomainInfo.LabeledEdit1.Text,
              frmDomainInfo.LabeledEdit2.Text) then
              begin
               LoadPointsRestoreWin(ListView1);// загруза списка точек восстановления
               //LoadRestoreConfig; // Загрузка доп.конфиг
              Panel1.Enabled:=true;
              end
              else
              begin
               ShowMessage('RPC сервер не доступен');
               Panel1.Enabled:=false;
              end;


end;

function TFormRestoreWin.LoadConfig:boolean;
var
 FWbemObjectSet  : OLEVariant;
 FWbemObject     : OLEVariant;
 oEnum           : IEnumvariant;
 iValue          : LongWord;
 res:integer;
begin
 try
 FWbemObjectSet:= FWMIServiceRes.ExecQuery('SELECT * FROM SystemRestoreConfig','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do
 begin
 if FWbemObject.DiskPercent<>null then frmDomainInfo.Memo1.Lines.Add('Процент диска - '+string(FWbemObject.DiskPercent));
 if FWbemObject.RPGlobalInterval<>null then frmDomainInfo.Memo1.Lines.Add('Интервал созадния точек - '+string(FWbemObject.RPGlobalInterval));
 if FWbemObject.RPLifeInterval<>null then frmDomainInfo.Memo1.Lines.Add('Время хранения точек - '+string(FWbemObject.RPLifeInterval));
 if FWbemObject.RPSessionInterval<>null then frmDomainInfo.Memo1.Lines.Add('Интервал создания сеансовых точек - '+string(FWbemObject.RPSessionInterval));
 VariantClear(FWbemObject);
 FWbemObject:=Unassigned;
 end;
 result:=true;
 except
 on E:Exception do
 begin
 frmDomainInfo.Memo1.Lines.Add('Ошибка '+e.Message);
 result:=false;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 end;
 on e:EInOutError do frmDomainInfo.Memo1.Lines.Add(inttostr(e.ErrorCode))
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
end;



function TFormRestoreWin.LoadPointsRestoreWin(ListView:TListView):boolean;
var
 FWbemObjectSet  : OLEVariant;
 FWbemObject     : OLEVariant;
 oEnum           : IEnumvariant;
 iValue          : LongWord;
 res:integer;
 function ResPointType(idType:integer):string;
 begin
 result:='';
 if idType=0 then result:='Установлено приложение';
 if idType=1 then result:='Удалено приложение';
 if idType=7 then result:='Системная точка';
 if idType=6 then result:='Точка отмены восстановления';
 if idType=10 then result:='Установлен драйвер устройства';
 if idType=12 then result:='В приложение были добавлены или удалены функции';
 if idType=13 then result:='Приложению необходимо удалить созданную точку восстановления.';
 if idType=18 then result:='Критическое обновление';
 end;
 const
 EvType: array [100..103] of string=('Начались системные изменения.','Смена системы закончилась.','Начались системные изменения.'
 +'Последующий вложенный вызов не создает новую точку восстановления.'
 +'  Последующие вызовы должны использовать END_NESTED_SYSTEM_CHANGE, а не END_SYSTEM_CHANGE.','Смена системы закончилась.');
begin
 try
 ListView.Clear;
 FWbemObjectSet:= FWMIServiceRes.ExecQuery('SELECT * FROM SystemRestore','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// профили
 begin
   if FWbemObject.SequenceNumber<>null then
   with ListView.Items.Add do
   begin
   Caption := vartostr(FWbemObject.SequenceNumber);
   if FWbemObject.CreationTime<>null then SubItems.Add(WbemTimeToDateTime(FWbemObject.CreationTime))   // ok
   else SubItems.Add('');
   if FWbemObject.Description<> null then SubItems.Add(vartostr(FWbemObject.Description))
    else SubItems.Add('');
   if FWbemObject.RestorePointType<>null then SubItems.Add(ResPointType(integer(FWbemObject.RestorePointType))+' ('+vartostr(FWbemObject.RestorePointType)+')') ///ok
    else SubItems.Add('');
   end;
 VariantClear(FWbemObject);
 FWbemObject:=Unassigned;
 end;
 result:=true;
 except
 on E:Exception do
 begin
 frmDomainInfo.Memo1.Lines.Add('Ошибка '+e.Message);
 result:=false;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 end;
 on e:EInOutError do frmDomainInfo.Memo1.Lines.Add(inttostr(e.ErrorCode))
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
end;

end.
