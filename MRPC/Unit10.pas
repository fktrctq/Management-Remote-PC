unit Unit10;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  inifiles, Vcl.Buttons, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.Comp.ScriptCommands, FireDAC.Stan.Util,
  FireDAC.Comp.Script, FireDAC.Comp.Client, FireDAC.Phys.IBBase, Data.DB,
  FireDAC.VCLUI.Script, FireDAC.Comp.UI, FireDAC.VCLUI.Login,Themes,Registry,math,
  Vcl.Imaging.pngimage,ShellApi;

type
  TForm10 = class(TForm)
    Edit5: TEdit;
    Edit4: TEdit;
    GroupBox9: TGroupBox;
    Memo1: TMemo;
    Edit1: TEdit;
    Edit2: TEdit;
    Button3: TButton;
    Button4: TButton;
    GroupBox10: TGroupBox;
    Label2: TLabel;
    Image1: TImage;
    Label3: TLabel;
    Edit3: TEdit;
    LabeledEdit10: TLabeledEdit;
    Button7: TButton;
    LinkLabel1: TLinkLabel;
    LinkLabel2: TLinkLabel;
    LinkLabel3: TLinkLabel;
    LinkLabel4: TLinkLabel;
    procedure Button7Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure LinkLabel1Click(Sender: TObject);
    procedure LinkLabel2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure reloadLC(Sender: TObject);
    procedure LinkLabel3Click(Sender: TObject);
    procedure LinkLabel4Click(Sender: TObject);
  private
     function regeditread(patch,Section:string):string;
    function regeditwrite(patch,Section,Value:string):boolean;
  public
    { Public declarations }
  end;

var
  Form10: TForm10;
  IniSettings:TMeminiFile;
  sm: TStyleManager;
  ThrLic:TThread;
  TrLic:Tthread;
  RootPatch: HKEY;
  peremennaya:string;
  keyboardstring,activeserver:string;
  reloadL:Ttimer;
const
 stringpatch:string ='software\MRPC\MRPC';
 //HKEY_LOCAL_MACHINE  HKEY_CURRENT_USER

implementation
uses umain,unit8,unit9;
{$R *.dfm}

procedure Code(var text: WideString; password: string;  //// процедура кодирования и декодирования файла
decode: boolean);
var
i, PasswordLength: integer;
sign: shortint;
begin
PasswordLength := length(password);
if PasswordLength = 0 then Exit;
if decode then sign := -1
else sign := 1;
for i := 1 to Length(text) do
text[i] := chr(ord(text[i]) + sign *
ord(password[i mod PasswordLength + 1]));
end;

procedure Callid;
begin
unit8.CalLic;
end;


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

function TForm10.regeditwrite(patch,Section,Value:string):boolean; /// CurKey - подсекция в реестре
var      //// Запись в реестр
regFile:TregIniFile;
begin
RegFile:=TregIniFile.Create(KEY_WRITE OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.CreateKey(patch) then
RegFile.WriteString(patch,Section,value);
if Assigned(RegFile) then freeandnil(regFile);
end;



function TForm10.regeditread(patch,Section:string):string;
var      /// чтение из реестра
regFile:TregIniFile;
begin
result:='';
RegFile:=TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then
begin
result:=(regFile.ReadString(patch,Section,''));
end;
if Assigned(RegFile) then freeandnil(regFile);
end;

procedure TForm10.Button3Click(Sender: TObject);
var
s:WideString;
begin
RootPatch:=HKEY_CURRENT_USER;
s:=Edit1.Text;
code(s,'1234',true);
s:=crypt(s);
regeditwrite(stringpatch,edit2.Text,s);
end;


procedure TForm10.Button4Click(Sender: TObject);
var
s:widestring;
begin
RootPatch:=HKEY_CURRENT_USER;
s:=regeditread(stringpatch,edit2.Text);
s:=crypt(s);
code(s,'1234',false);
Memo1.Lines.add(s);
end;

procedure TForm10.Button2Click(Sender: TObject);
begin
close;
end;




procedure TForm10.Button7Click(Sender: TObject);
begin
keyboardstring:=LabeledEdit10.text;
activeserver:=Edit3.text;
peremennaya:=keyboardstring;
showmessage('Перезапустите приложение.');
RootPatch:=HKEY_CURRENT_USER;
regeditwrite(stringpatch,'isOK',keyboardstring);
regeditwrite(stringpatch,'srv',Edit3.text);
end;








procedure TForm10.reloadLC(Sender: TObject);
begin
TrLic:=unit9.Lcheck.Create(true);
TrLic.FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
TrLic.Priority:=tpLower;
TrLic.start;
if Assigned(reloadL) then
  begin
  reloadL.Interval:=10000*RandomRange(15,20);
  reloadL.Tag:=reloadL.Tag+1;
  if reloadL.tag>RandomRange(1,3) then
    begin
    reloadL.Enabled:=false;
    FreeAndNil(reloadL);
    end;
  end;
end;


procedure TForm10.FormCreate(Sender: TObject);
var
i,z:integer;
begin
try
if not Assigned(reloadL) then
begin
reloadL:=Ttimer.Create(frmDomainInfo);
reloadL.Name:='reloadLD';
z:=RandomRange(15,20);
reloadL.Interval:=10000*z;
reloadL.OnTimer:=reloadLC;
end;
RootPatch:=HKEY_CURRENT_USER;
peremennaya:=regeditread(stringpatch,'isOK');
//activeserver:=regeditread(stringpatch,'srv');
activeserver:=Edit3.text;
LabeledEdit10.Text:=peremennaya;
Callid;
 except
    memo1.Lines.Add('Ошибка')
  end;
end;



procedure TForm10.LinkLabel1Click(Sender: TObject);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar('https://skrblog.ru/'), nil, nil, SW_SHOWNORMAL);
  except
    memo1.Lines.Add('Ошибка при открытии приложения')
  end;
end;


procedure TForm10.LinkLabel2Click(Sender: TObject);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar('mailto:skrblog@mail.ru?subject=Приобретение лицензии&body=Мой ID:'+memo1.Text), nil, nil, SW_SHOWNORMAL);
  except
    memo1.Lines.Add('Ошибка при открытии приложения')
  end;
end;





procedure TForm10.LinkLabel3Click(Sender: TObject);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar('https://skrblog.ru/buy/'), nil, nil, SW_SHOWNORMAL);
  except
    memo1.Lines.Add('Ошибка при открытии приложения')
  end;
end;

procedure TForm10.LinkLabel4Click(Sender: TObject);
begin
try
shellAPI.ShellExecute(0, 'Open', PChar('https://skrblog.ru/manual/'), nil, nil, SW_SHOWNORMAL);
  except
    memo1.Lines.Add('Ошибка при открытии приложения')
  end;

end;

end.
