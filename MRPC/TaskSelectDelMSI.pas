unit TaskSelectDelMSI;

interface

uses
  SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,dialogs,IdIcmpClient,
  System.Variants,ActiveX,ComObj,CommCtrl;

type
  TSelectDelMSITask = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Button1: TButton;
    ComboBox1: TComboBox;
    CheckBox2: TCheckBox;
    Edit1: TEdit;
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    function LoadListProcess(NamePC:string):boolean;
    function LoadMSI(namePC:string):boolean;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SelectDelMSITask: TSelectDelMSITask;

implementation
uses umain,TaskEdit;
{$R *.dfm}


function ping(s:string):boolean;
var
z:integer;
MyIdIcmpClient: TIdIcmpClient;
begin
try
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=3000;
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
      begin
       result:=false;
      end
    else
      begin
      result:=true;
      end;
  if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
  except
  on E: Exception do
    begin
    if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
    result:=false;
    end;
  end;

end;


function TSelectDelMSITask.LoadListProcess(NamePC:string):boolean; /// загрузить список процессов
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
ComboBox1.Clear;
 frmDomainInfo.memo1.Lines.Add('Загрузка списка процессов..');
 try
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Process WHERE ProcessId<>0','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// процессы
 begin
 if FWbemObject.Caption<>null then
 ComboBox1.Items.add(string(FWbemObject.Caption));       // Его имя
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
 frmDomainInfo.memo1.Lines.Add('Ошибка загрузки процессов "'+E.Message+'"');
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 end;
 end;
OleUnInitialize;
ComboBox1.text:=('Загрузка списка процессов завершена.');
end;




function TSelectDelMSITask.LoadMSI(namePC:string):boolean;
  var
   FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject       : OLEVariant;
  FOutParams      : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum         : IEnumvariant;
  Dt            : TDateTime;
  iValue        : LongWord;
begin
try
OleInitialize(nil);
ComboBox1.Clear;
ComboBox1.text:=('Ожидайте завершения загрузки программ...');
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while oEnum.Next(1, FWbemObject, iValue) = 0 do
begin
if FWbemObject.Caption<> null then
combobox1.Items.add(string(FWbemObject.Caption));
FWbemObject:=Unassigned;
end;
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
except
on E:Exception do
begin
showmessage('Ошибка загрузки программ - "'+E.Message+'"');
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
exit;
end;
end;
ComboBox1.text:=('Загрузка списка программ завершена.');
end;



procedure TSelectDelMSITask.Button1Click(Sender: TObject);
var
NamePSSearch:string;
begin
NamePSSearch:='localhost';
if not InputQuery('Загрузка списка может занять несколько минут!!!', 'Имя компьютера', NamePSSearch) then exit;
   if ping(NamePSSearch) then
   begin
    if SelectDelMSITask.Caption='Удаление программ msi' then LoadMSI(NamePSSearch);
    if SelectDelMSITask.Caption='Завершение процесса' then LoadListProcess(NamePSSearch);
   end
   else showmessage('Компьютер не доступен');
end;

procedure TSelectDelMSITask.CancelBtnClick(
  Sender: TObject);
begin
Close;
end;


procedure TSelectDelMSITask.CheckBox2Click(
  Sender: TObject);
begin
if CheckBox2.Checked=false then
showmessage('Будте внимательны!!!'+#10#13+'Удалятся все программы соответствующие шаблону - '+Edit1.Text);

end;

procedure TSelectDelMSITask.ComboBox1Select(
  Sender: TObject);
begin
Edit1.Text:=ComboBox1.Text;
end;

function StrInDescription(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=br;
end;

procedure TSelectDelMSITask.OKBtnClick(
  Sender: TObject);
  var
  i,z:integer;
function bToSTR(b:boolean):string;
begin
  if b then result:='1'
  else result:='0';
end;
begin
if SelectDelMSITask.Caption='Удаление программ msi' then
Begin
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='Удаление программы :'+Edit1.text) and (EditTask.ListView1.Items[i].SubItems[1]='DelMSI') then
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
  caption:=bToSTR(CheckBox2.Checked); // Точное совпадение имени или нет
  SubItems.Add('Удаление программы :'+Edit1.text);    // имя программы
  SubItems.Add('DelMSI');            // Имя операции
end;
End;

if SelectDelMSITask.Caption='Завершение процесса' then
Begin
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if (EditTask.ListView1.Items[i].SubItems[0]='Завершение процесса :'+Edit1.text) and (EditTask.ListView1.Items[i].SubItems[1]='KillProc') then
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
  caption:=bToSTR(CheckBox2.Checked); // Точное совпадение имени процесса или нет
  SubItems.Add('Завершение процесса :'+Edit1.text);    // имя процесса
  SubItems.Add('KillProc');            // Имя операции
end;
End;
end;


end.
