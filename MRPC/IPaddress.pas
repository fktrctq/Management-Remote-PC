unit IPaddress;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, IdBaseComponent,
  IdNetworkCalculator,Vcl.dialogs,Inifiles;

type
  TOKRightDlg1234567891011121314151617181920 = class(TForm)
    CancelBtn: TButton;
    IdNetworkCalculator1: TIdNetworkCalculator;
    Edit1: TEdit;
    Edit2: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    CheckBox1: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBox1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg1234567891011121314151617181920: TOKRightDlg1234567891011121314151617181920;

implementation
uses umain,MyDM;
{$R *.dfm}

procedure TOKRightDlg1234567891011121314151617181920.Button1Click(
  Sender: TObject);
var
  i:integer;
  setini:Tmeminifile;
  listip:tstringlist;
begin
try
IdNetworkCalculator1.NetworkAddress.AsString:=string(Edit1.Text);
IdNetworkCalculator1.NetworkMask.AsString:=string(Edit2.Text);
except
  begin
    showmessage('Введены не корректные данные IP -адреса или маски.');
    Exit;
  end;
end;
frmDomainInfo.ListView8.Clear;
frmDomainInfo.combobox2.clear;
frmDomainInfo.combobox1.Enabled:=false; // список компьютеров
frmDomainInfo.combobox1.Text:='';
frmDomainInfo.combobox8.Enabled:=false; /// список пользователей
frmDomainInfo.combobox8.Text:='';
for I := 0 to IdNetworkCalculator1.ListIP.Count-1 do
begin
if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),Ansiuppercase(frmDomainInfo.Caption))<>0)
and (i>=30) then break;
  frmDomainInfo.combobox2.Items.Add(IdNetworkCalculator1.ListIP[i]);
  with frmDomainInfo.listview8.Items.Add do
    begin
      caption:='';
      subitems.add(IdNetworkCalculator1.ListIP[i]);
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
      subitems.add('');
    end;
end;
   frmDomainInfo.StatusBar1.Panels[0].Text:=inttostr(frmDomainInfo.ListView8.Items.Count);
   PCCheck:=0;
   frmDomainInfo.StatusBar1.Panels[1].Text:='0';
 // сохраняем список адресов в файл
IdNetworkCalculator1.ListIP.SaveToFile(extractfilepath(application.ExeName)+'\listIP.txt');
IdNetworkCalculator1.ListIP.Clear;
/// указываем в файле настроек что последний раз открывали список адресов
if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
  begin
  SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  setini.WriteString('ConfLAN','type','IP');
  setini.WriteString('ConfLAN','IP',Edit1.Text);
  setini.WriteString('ConfLAN','subnet',Edit2.Text);
  SetInI.WriteBool('ConfLAN','LoadInfoDB',CheckBox1.Checked);
  setini.UpdateFile;
  setini.Free;
  end;
if CheckBox1.Checked then
if dataM.ConnectionDB.Connected then frmDomainInfo.readinfoforpcDB;
close;
end;


procedure TOKRightDlg1234567891011121314151617181920.CheckBox1MouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  var
 setini:Tmeminifile;
begin
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
SetInI.WriteBool('ConfLAN','LoadInfoDB',CheckBox1.Checked);
finally
setini.UpdateFile;
setini.Free;
end;
end;

procedure TOKRightDlg1234567891011121314151617181920.FormShow(Sender: TObject);
var
setini:Tmeminifile;
begin
  SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  try
  Edit1.Text:=setini.readString('ConfLAN','IP','192.168.0.0');
  Edit2.Text:=setini.readString('ConfLAN','subnet','255.255.255.0');
  if SetInI.ReadString('ConfLAN','type','FirstStart')='FirstStart' then
  CheckBox1.Checked:=false
  else CheckBox1.Checked:=SetInI.ReadBool('ConfLAN','LoadInfoDB',true);
  finally
  setini.Free;
  end;
end;

end.
