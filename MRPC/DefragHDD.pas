unit DefragHDD;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,System.Variants,ActiveX,ComObj;

type
  TOKRightDlg1234 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    LabeledEdit2: TLabeledEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
  private
    function  VolumeRead(IDHDD :string):boolean;
    function  VolumeWrite(IDHDD :string):string;
  public
    { Public declarations }
  end;

var
  OKRightDlg1234: TOKRightDlg1234;
  ForceHDDDefrag,HDDDefraganalysis:TThread;

implementation
uses DefragForce,umain,DefragAnalysis;

{$R *.dfm}
function  TOKRightDlg1234.VolumeWrite(IDHDD :string):string;
var
  FSWbemLocator : OleVariant;
  FWMIService   : oleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject   : OleVariant;
  FInParams     : OleVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;

begin
try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+IDHDD+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
      if LabeledEdit2.text='' then FWbemObject.label:=null
     else FWbemObject.label:=LabeledEdit2.text;
     if LabeledEdit1.text='' then  FWbemObject.DriveLetter:=null
     else FWbemObject.DriveLetter:=LabeledEdit1.text;
     FWbemObject.IndexingEnabled:=CheckBox1.Checked;
     result:=FWbemObject.put_();
     FWbemObject:=Unassigned;
      end;

 frmDomainInfo.memo1.Lines.Add('Результат операции - '+result);
 except
    on E:Exception do
    begin
      frmDomainInfo.memo1.Lines.Add('Общая ошибка '+E.Message);
    end;
 end;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

function  TOKRightDlg1234.VolumeRead(IDHDD :string):boolean;
var
  FSWbemLocator : OleVariant;
  FWMIService   : oleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject   : OleVariant;
  FInParams     : OleVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;

begin
try
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+IDHDD+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
   if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin                 /// функция возвращает неизвестную ошибку КОД -11, wmi не позволяет делать дефрагментацию диска на удаленном компьютере
      if FWbemObject.label<>null then
      begin
      LabeledEdit2.text:=string(FWbemObject.label);
      caption:='Диск -'+string(FWbemObject.label);
      end
      else
      begin
      LabeledEdit2.text:='';
      caption:='Диск -';
      end;
      if FWbemObject.DriveLetter<>null then  LabeledEdit1.text:=string(FWbemObject.DriveLetter)
      else LabeledEdit1.text:='';
      if not VarIsNull(FWbemObject.IndexingEnabled) then  CheckBox1.Checked:=(FWbemObject.IndexingEnabled)
      else CheckBox1.Checked:=false;
      FWbemObject:=Unassigned;
      end;
 except
    on E:Exception do
    begin
      frmDomainInfo.memo1.Lines.Add('Общая ошибка '+E.Classname+':'+E.Message);
    end;
 end;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TOKRightDlg1234.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg1234.FormShow(Sender: TObject);
begin
 VolumeRead(SelectedHDD);
end;

procedure TOKRightDlg1234.OKBtnClick(Sender: TObject);
begin
VolumeWrite(SelectedHDD);
end;

procedure TOKRightDlg1234.SpeedButton1Click(Sender: TObject);
begin
MyPS:=frmDomainInfo.Combobox2.Text;
MyUser:=frmDomainInfo.LabeledEdit1.Text;
MyPasswd:=frmDomainInfo.LabeledEdit2.Text;
ForceHDDDefrag:=DefragForce.MyDefragForce.Create(true);
ForceHDDDefrag.FreeOnTerminate:=true;
ForceHDDDefrag.Start;
close;
end;

procedure TOKRightDlg1234.SpeedButton2Click(Sender: TObject);
begin
MyPS:=frmDomainInfo.Combobox2.Text;
MyUser:=frmDomainInfo.LabeledEdit1.Text;
MyPasswd:=frmDomainInfo.LabeledEdit2.Text;
HDDDefraganalysis:=DefragAnalysis.HDDDefragAnalysis.Create(true);
HDDDefraganalysis.FreeOnTerminate:=true;
HDDDefraganalysis.Start;
close;
end;

end.
