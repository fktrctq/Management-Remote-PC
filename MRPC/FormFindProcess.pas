unit FormFindProcess;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,ActiveX,ComObj;

type
  TForm12 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    LabeledEdit1: TLabeledEdit;
    Button3: TButton;
    ComboBox1: TComboBox;
    procedure Button3Click(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    function loadprocess(namePC,Login,Pass:string):bool;
  public
    { Public declarations }
  end;

var
  Form12: TForm12;
  ListFindProcess:TstringList;
  ProcessNameForFind:string;


const
wbemFlagForwardOnly = $00000020;

implementation

uses umain,FindProcessGroup;
{$R *.dfm}

procedure TForm12.Button1Click(Sender: TObject);
var
FindProcessThread:TThread;
begin
try
if LabeledEdit1.text='' then
begin
  showmessage('Введите имя процесса');
  exit;
 end;
ProcessNameForFind:=LabeledEdit1.text;
ListFindProcess:=TstringList.create;
ListFindProcess:=frmDomaininfo.createListpcForCheck('');
if ListFindProcess.Count<1 then
begin
  showmessage('Список компьютеров пуст');
  exit;
end;
FindProcessThread:=FindProcessGroup.GroupFindProcess.create(true);
FindProcessThread.FreeOnTerminate:=true;
FindProcessThread.start;
close;
except
on E:Exception do
begin
ShowMessage('Ошибка запуска поиска процесса: '+e.Message);
end;
end;
end;

procedure TForm12.Button2Click(Sender: TObject);
begin
close;
end;

procedure TForm12.Button3Click(Sender: TObject);
var
PCname:string;
begin
ComboBox1.Clear;
 if not InputQuery('Загрузка может занять определенное время', 'Имя компьютера:', PCname)
    then exit;
  if (PCname='') then PCname:='localhost';
  if frmDomaininfo.ping(PCname) then
  begin
  Combobox1.text:='Загрузка списка процессов';
  loadprocess(PCname,frmDomaininfo.labeledEdit1.text,frmDomaininfo.labeledEdit2.text);
  Combobox1.text:='Загрузка завершена';
  end
  else showmessage(PCname+' не доступен');
end;

procedure TForm12.ComboBox1Select(Sender: TObject);
begin
labelededit1.text:=Combobox1.text;
end;

function TForm12.loadprocess(namePC,Login,Pass:string):bool;
var
 FSWbemLocator          : OLEVariant;
  FWMIService           : OLEVariant;
  FWbemObjectSet        : OLEVariant;
  FWbemObject           : OLEVariant;
  oEnum                 : IEnumvariant;
  iValue                : LongWord;
begin
   try
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', Login, Pass,'','',128);
FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Process WHERE ProcessId<>0','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
while (oEnum.Next(1, FWbemObject, iValue) = 0) do
begin
if FWbemObject.Caption<>null then
ComboBox1.Items.Add (string(FWbemObject.Caption));       // Его имя
VariantClear(FWbemObject);
end;
except
on E:Exception do
begin
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
end;
end;
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
end;
end.
