unit RemoteDesktopSettingDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Menus,
  inifiles,VCL.Dialogs, Vcl.Samples.Spin;

type
  TRemoteDesktopSetting = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    ListView1: TListView;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    Memo1: TMemo;
    CheckBox2: TCheckBox;
    Label1: TLabel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    Label2: TLabel;
    GroupBox1: TGroupBox;
    CheckBox8: TCheckBox;
    Edit1: TEdit;
    GroupBox2: TGroupBox;
    EditRDPFPS: TSpinEdit;
    PopupDeptColor: TPopupMenu;
    N161: TMenuItem;
    N241: TMenuItem;
    ComboBox1: TComboBox;
    Label3: TLabel;
    Label4: TLabel;
    procedure FormShow(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RemoteDesktopSetting: TRemoteDesktopSetting;
  TransformUser:byte; /// 1 - Новый пользователь 2 - изменить текущего пользователя 3 - удалить пользователя
  ListPassword:TstringList;
  implementation
uses umain,RemoteDesktopSettingTransformDialog,MRDUNew;
{$R *.dfm}
var
  MemFileIni:TMeminifile;
  MyListFile:TstringList;


function Crypt(varStr: WideString):WideString;
var
 k: integer;
 s: WideString;
begin
   RandSeed:=100;
   s:=varStr;
   for k:=1 to Length(s) do
    s[k]:=Chr(ord(s[k]) xor (Random(127)+1));

 Crypt:=s;
end;

function AccessRemove(s:string):byte;
var
i:byte;
const Access: Array [0..2] of string=('Полный доступ','Управление','Просмотр');
begin
for I := 0 to 2 do
  begin
    if s=Access[i] then result:=i;
  end;
end;

procedure TRemoteDesktopSetting.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TRemoteDesktopSetting.CheckBox1Click(Sender: TObject);
begin
LabeledEdit1.Enabled:=CheckBox1.Checked;
end;

procedure TRemoteDesktopSetting.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
ListPassword.Free;
end;

procedure TRemoteDesktopSetting.FormShow(Sender: TObject);
var
UnloadFile:TMemInifile;
i:integer;
const Access: Array [0..2] of string=('Полный доступ','Управление','Просмотр');
begin
caption:='Настройка прав сервера - '+MyPS;
checkbox1.Checked:=false;
LabeledEdit1.Enabled:=false;
Listview1.Clear;
ListPassword:=TStringList.Create;
UnLoadFile:=TMemIniFile.Create(ChangeFileExt( Application.Exename,'UnLF.ini'));
UnLoadFile.SetStrings(Memo1.Lines);
for I := 0 to UnLoadFile.ReadInteger('RemoteClient','Count',0) do
  begin
    with ListView1.Items.Add do
      begin
       caption:=UnLoadFile.ReadString('RemoteClient','ClientName'+inttostr(i),'Ошибка чтения данных');
       subitems.add(Access[strtoint(UnLoadFile.ReadString('RemoteClient','Access'+inttostr(i),'0'))]);
       ListPassword.Add(UnLoadFile.ReadString('RemoteClient','ClientPasswd'+inttostr(i),''));
      end;
  end;
LabeledEdit1.Text:=UnLoadFile.ReadString('Server','Port','Ошибка');
CheckBox2.Checked:=UnLoadFile.readBool('Server','View',true);
CheckBox3.Checked:=UnLoadFile.readBool('FW','uRDMServer',false);
CheckBox4.Checked:=UnLoadFile.readBool('FW','FilePrintSh',false);
CheckBox5.Checked:=UnLoadFile.readBool('FW','RDP',false);
CheckBox6.Checked:=UnLoadFile.readBool('FW','WMI',false);
CheckBox7.Checked:=UnLoadFile.readBool('FW','RPC',false);
Edit1.Text:= UnLoadFile.ReadString('Server','IP','');
UnLoadFile.Free;
end;

procedure TRemoteDesktopSetting.N1Click(Sender: TObject);  ///// НОВЫЙ ПОЛЬЗОВАТЕЛЬ
begin
TransformUser:=1;
RemoteDesktopSettionTransform.ShowModal;
RemoteDesktopSettionTransform.Caption:='Новый пользователь'
end;

procedure TRemoteDesktopSetting.N2Click(Sender: TObject);  //// изменить пользователя
begin
if ListView1.Items.Count<>0 then
  begin
   TransformUser:=2;
   RemoteDesktopSettionTransform.ShowModal;
   RemoteDesktopSettionTransform.Caption:=listview1.Selected.Caption;
  end;
end;

procedure TRemoteDesktopSetting.N3Click(Sender: TObject);  ////// удалить пользователя
begin
if ListView1.Items.Count>1 then
 begin
   if ListView1.Selected.Caption<>'' then
      ListView1.Items[ListView1.Selected.Index].Delete;
 end;
end;

procedure TRemoteDesktopSetting.OKBtnClick(Sender: TObject);
var
s:Widestring;
i:integer;
mylist2:Widestring;
VarMls:variant;
begin
try
mylist2:='';
s:='';

if RemoteDesktopSetting.Tag=1 then
 begin
 s:=('<Client>'+frmdomainInfo.labeledEdit4.Text
+'<passwd>'+frmdomainInfo.labeledEdit5.Text+'<settingFile>');
//// шифрация текста
s:=crypt(s);
 frmdomainInfo.ClientSocket1.Socket.SendText(s);
 end;

if RemoteDesktopSetting.Tag=0 then
 begin
  s:=('<Client>'+MRDForm.EditUser.Text
+'<passwd>'+MRDForm.EditPass.Text+'<settingFile>');
//// шифрация текста
s:=crypt(s);
 MRDForm.ClientSocket1.Socket.SendText(s);
 end;

 /// отправляем зашифрованное сообщение
sleep(500);
MemFileIni:=TMemIniFile.Create(ChangeFileExt( Application.Exename,'MRP.ini'));  //// создаем ini файл в памяти для последующей отправки
 for I := 0 to listview1.Items.Count-1 do
   begin
    memo1.Lines.Add(listview1.Items[i].Caption);
    memo1.Lines.Add(listview1.Items[i].SubItems[0]);
    memo1.Lines.Add(ListPassword[i]);
    MemFileIni.WriteString('RemoteClient','ClientName'+inttostr(i),listview1.Items[i].Caption);
    MemFileIni.WriteString('RemoteClient','ClientPasswd'+inttostr(i),ListPassword[i]);
    MemFileIni.WriteInteger('RemoteClient','Access'+inttostr(i),AccessRemove(Listview1.Items[i].SubItems[0]));
   end;
 MemFileIni.Writeinteger('RemoteClient','Count',listview1.Items.Count-1);
 MemFileIni.WriteString('Server','Port',LabeledEdit1.Text);
 MemFileIni.WriteBool('Server','View',CheckBox2.Checked);
 MemFileIni.WriteBool('FW','uRDMServer',CheckBox3.Checked);
 MemFileIni.WriteBool('FW','FilePrintSh',CheckBox4.Checked);
 MemFileIni.WriteBool('FW','RDP',CheckBox5.Checked);
 MemFileIni.WriteBool('FW','WMI',CheckBox6.Checked);
 MemFileIni.WriteBool('FW','RPC',CheckBox7.Checked);
 MemFileIni.WriteBool('FW','FullFW',CheckBox8.Checked);
 MemFileIni.WriteString('Server','IP',Edit1.Text);

  try
  MemFileIni.WriteString('Server','colordepth',ComboBox1.Text);
 //mrdunew.MRDuRDPViewer.RequestColorDepthChange(strtoint(ComboBox1.Text));
// mrdunew.MRDuRDPViewer.ControlInterface.RequestColorDepthChange(strtoint(ComboBox1.Text));
 except
 on E: Exception do  memo1.Lines.add('Colordepth :'+E.Message);
 end;
 try
 MemFileIni.WriteString('Server','FPS',inttostr(EditRDPFPS.Value));
 //VarMls:=variant(EditRDPFPS.Value);
  mrdunew.MRDuRDPViewer.Properties.Property_['FrameCaptureIntervalInMs']:=1000 div EditRDPFPS.Value;  //https://docs.microsoft.com/ru-ru/windows/win32/api/rdpencomapi/nf-rdpencomapi-irdpsrapisessionproperties-get_property
 //mrdunew.MRDuRDPViewer.Properties.set_Property_('FrameCaptureIntervalInMs',(1000 div EditRDPFPS.Value)); // значение передается в милисекундах
 except
 on E: Exception do  frmDomainInfo.memo1.Lines.add('FPS :'+E.Message);
 end;


  MyListFile:=TstringList.Create;
 //memo1.Lines.Add('передаем в MyListFile строки');
 MemFileIni.GetStrings(MyListFile); //// сохраняем строки в list
 UnLoadFile.SetStrings(MyListFile); //передаем настройки в ini файл для того чтобы при показе формы их считать
 for I := 0 to MyListFile.Count-1 do
  begin
    mylist2:=mylist2+(WideString(MyListFile[i]+#13#10));
  end;
 // memo1.Lines.Add(mylist2);
  //memo1.Lines.Add('шифрую файл');
  mylist2:=crypt(mylist2);
// frmdomainInfo.ClientSocket1.Socket.SendText(Mylist2);
if RemoteDesktopSetting.Tag=1 then frmdomainInfo.ClientSocket1.Socket.SendText(Mylist2);
if RemoteDesktopSetting.Tag=0 then MRDForm.ClientSocket1.Socket.SendText(Mylist2);
MemFileIni.Free;
MyListFile.Free;
RemoteDesktopSetting.Tag:=2;
 Except
   on E: Exception do
      begin
        memo1.Lines.add('Ошибка сохранения настроек на сервере:'+E.Message);
      end;
   end;

end;

end.
