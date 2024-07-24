unit NewStartup;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,Dialogs,ActiveX,ComObj;

type
  TOKRightDlg2 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    ComboBox1: TComboBox;
    Label1: TLabel;
    LabeledEdit3: TLabeledEdit;
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg2: TOKRightDlg2;

implementation
uses umain;

{$R *.dfm}


function hkpath(s:string):string;
var
i:byte;
begin
i:=pos('\',s);
delete(s,1,i);
result:=s;
end;

function hk(s:string):byte;
begin
Delete(s,5,Length(s)-4);
if s='HKCR' then  result:=0;
if s='HKCU' then  result:=1;
if s='HKLM' then  result:=2;
if s='HKU\' then  result:=3;
if s='HKCC' then  result:=4;
end;

procedure TOKRightDlg2.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg2.FormShow(Sender: TObject);
begin
    try
combobox1.ItemIndex:=hk(frmDomainInfo.listview2.Selected.SubItems[2]);
LabeledEdit3.text:=hkpath(frmDomainInfo.listview2.Selected.SubItems[2]);
LabeledEdit2.text:=frmDomainInfo.listview2.Selected.SubItems[1];
LabeledEdit1.text:=frmDomainInfo.listview2.Selected.SubItems[0];
    except
     on E:Exception do
         begin
       LabeledEdit3.text:='SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
       LabeledEdit1.text:='StartNotepad';
       LabeledEdit2.text:='C:\Windows\System32\notepad.exe';
       combobox1.ItemIndex:=1;
         end;
    end;
end;

procedure TOKRightDlg2.OKBtnClick(Sender: TObject);
var
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FSWbemLocator : OLEVariant;
  FInParams       : OLEVariant;
  FOutParams      : OLEVariant;
begin
try
begin
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\default', Myuser, MyPasswd);
  FWbemObjectSet:= FWMIService.Get('StdRegProv');
  FInParams     := FWbemObjectSet.Methods_.Item('SetExpandedStringValue').InParameters.SpawnInstance_();  ////  запись значени€
    case combobox1.ItemIndex of
      0: FInParams.hDefKey:=HKEY_CLASSES_ROOT;
      1: FInParams.hDefKey:=HKEY_CURRENT_USER;
      2: FInParams.hDefKey:=HKEY_LOCAL_MACHINE;
      3: FInParams.hDefKey:=HKEY_USERS;
      4: FInParams.hDefKey:=HKEY_CURRENT_CONFIG;
    end;
   FInParams.sSubKeyName:=LabeledEdit3.text;  /// путь
   //showmessage('sValueName');
   FInParams.sValueName:=LabeledEdit1.text;     /// им€
  //showmessage('sValueName опачки');
   FInParams.sValue:=LabeledEdit2.text;         /// значение
  // showmessage('sValueName создаем');
   FOutParams    := FWMIService.ExecMethod('StdRegProv', 'SetExpandedStringValue', FInParams);   ////запрос
   // showmessage('sValueName создали');
   frmDomainInfo.memo1.Lines.Add('ƒобавление автозагрузки. '+(SysErrorMessage(FOutParams.ReturnValue)));
   OleUnInitialize;
   end;
  except
     on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('ќшибка добавлени€ автозагрузки "'+E.Message+'"');
           OleUnInitialize;
           exit;
         end;
   end;
end;

end.
