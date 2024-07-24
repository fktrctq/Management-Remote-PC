unit SelectedPCDeleteProgramDialog;

interface

uses SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,dialogs,IdIcmpClient,
  System.Variants,ActiveX,ComObj,CommCtrl;

type
  TOKRightDlg12345678910111213141516171819202122 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    Label1: TLabel;
    Button1: TButton;
    ComboBox1: TComboBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    Label3: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678910111213141516171819202122: TOKRightDlg12345678910111213141516171819202122;

  MyIdIcmpClient: TIdIcmpClient;
  SelectedDeleteProgram:TThread;
  equallyName,GroupPCDeleteProg:boolean;
  GroupPCSelectProg:string;
  SelectedPCDeleteProg:TstringList;

implementation
uses umain,SelectedPCDeleteProgramThread;
{$R *.dfm}
function ping(s:string):boolean;
var
z:integer;
begin
try
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
  except
  on E: Exception do
    begin
      result:=false;
    end;
  end;

end;





procedure TOKRightDlg12345678910111213141516171819202122.Button1Click(
  Sender: TObject);
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
  NamePSSearch  :string;
begin
if InputQuery('Загрузка программ может занять несколько минут!!!', 'Имя компьютера', NamePSSearch) then
  begin
      //////////////////////////////////////
    Myidicmpclient:=TIdIcmpClient.Create;
    Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
    Myidicmpclient.PacketSize:=32;
    Myidicmpclient.Port:=0;
    Myidicmpclient.Protocol:=1;
    Myidicmpclient.ReceiveTimeout:=3000;
    ////////////////////////////////////////
      if ping(NamePSSearch) then
          begin
            try
              OleInitialize(nil);
              ComboBox1.Clear;
              ComboBox1.text:=('Ожидайте завершения загрузки программ...');
              FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
              FWMIService   := FSWbemLocator.ConnectServer(NamePSSearch, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
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
        end
       else showmessage('Компьютер не доступен');
  end;
if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
end;

procedure TOKRightDlg12345678910111213141516171819202122.CancelBtnClick(
  Sender: TObject);
begin
GroupPCDeleteProg:=true;
Close;
end;

procedure TOKRightDlg12345678910111213141516171819202122.CheckBox2Click(
  Sender: TObject);
begin
if CheckBox2.Checked=false then
showmessage('Будте внимательны!!!'+#10#13+'Удалятся все программы соответствующие шаблону - '+LabelEdEdit1.Text);

end;

procedure TOKRightDlg12345678910111213141516171819202122.ComboBox1Select(
  Sender: TObject);
begin
LabeledEdit1.Text:=ComboBox1.Text;
end;

procedure TOKRightDlg12345678910111213141516171819202122.FormShow(
  Sender: TObject);
begin
label3.Visible:=false;
end;

procedure TOKRightDlg12345678910111213141516171819202122.OKBtnClick(
  Sender: TObject);
  var
  i,z:integer;
begin
if LabelEdEdit1.Text='' then
begin
  showmessage('Вы не указали имя программы');
end;
i:=MessageDlg('Вы действительно хотите удалить эту программу???',mtConfirmation,[mbYes,mbCancel], 0);
if i=mrCancel then exit;

SelectedPCDeleteProg:=TStringList.Create;
SelectedPCDeleteProg:=frmDomainInfo.createListpcForCheck(''); /// созданем список компьютеров
if SelectedPCDeleteProg.Count=0 then
begin
  ShowMessage('Не выбран список компьютеров.');
  SelectedPCDeleteProg.Free;
  exit;
end;

equallyName:=CheckBox2.Checked;
GroupPCSelectProg:=LabelEdEdit1.Text;  /// имя программы для удаления
GroupPCDeleteProg:=CheckBox1.Checked;  /// признак запуска в одном или разных потоках

if CheckBox1.Checked then
begin
  SelectedDeleteProgram:=SelectedPCDeleteProgramThread.SelectedPCDeleteProgram.Create(true);
  SelectedDeleteProgram.FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
  SelectedDeleteProgram.start;
  end
 else
  begin
  label3.Caption:='%';
  label3.Visible:=true;
  for z := 0 to SelectedPCDeleteProg.Count-1 do
  begin
  label3.Caption:=inttostr (((100 div SelectedPCDeleteProg.Count-1)*z+1))+'%';
  Application.ProcessMessages;
  Sleep(500);
  if GroupPCDeleteProg then break;
  CurrentPCDiffThread:=SelectedPCDeleteProg[z];
  SelectedDeleteProgram:=SelectedPCDeleteProgramThread.SelectedPCDeleteProgram.Create(true);
  SelectedDeleteProgram.FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
  SelectedDeleteProgram.start;
  end;
end;

   if Assigned(SelectedPCDeleteProg) then freeandnil(SelectedPCDeleteProg);
end;

end.
