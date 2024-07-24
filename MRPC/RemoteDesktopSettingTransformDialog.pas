unit RemoteDesktopSettingTransformDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,Messages,Vcl.Dialogs;

type
  TRemoteDesktopSettionTransform = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    ComboBox1: TComboBox;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function findPrava(prava:string):byte;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RemoteDesktopSettionTransform: TRemoteDesktopSettionTransform;
  UserIndexSelected:integer;
implementation
 uses RemoteDesktopSettingDialog;
{$R *.dfm}

procedure TRemoteDesktopSettionTransform.CancelBtnClick(Sender: TObject);
begin
close;
end;

function TRemoteDesktopSettionTransform.findPrava(prava:string):byte;
var
i:byte;
begin
  for I := 0 to combobox1.Items.Count-1 do
   begin
     if prava=combobox1.Items[i] then result:=i;
   end;
end;


procedure TRemoteDesktopSettionTransform.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
TransformUser:=0; ////
end;
/// 1 - Ќовый пользователь 2 - изменить текущего пользовател€
procedure TRemoteDesktopSettionTransform.FormShow(Sender: TObject); //// загрузка формы
begin
case transformUser of
1:
  begin
  LabeledEdit1.Text:='';
  LabeledEdit2.Text:='';
  LabeledEdit3.Text:='';
  Combobox1.ItemIndex:=0;
  end;

2:
 begin
 UserIndexSelected:=RemoteDesktopSetting.ListView1.Selected.Index;
 LabeledEdit1.Text:= RemoteDesktopSetting.ListView1.Selected.Caption;
 combobox1.ItemIndex:=findprava(RemoteDesktopSetting.ListView1.Selected.SubItems[0]);
 LabeledEdit2.Text:='';
  LabeledEdit3.Text:='';
 end;


end;

end;

procedure TRemoteDesktopSettionTransform.OKBtnClick(Sender: TObject); ///// OK
begin
if LabeledEdit2.Text<>LabeledEdit3.Text then
  begin
    showmessage('ѕароли не совпадают');
    exit;
  end;

if (LabeledEdit1.Text='') or (LabeledEdit2.Text='') then
  begin
    showmessage('ѕароль/»м€ пользовател€ не могут буть пустыми');
    exit;
  end;

 begin    /// 1 - Ќовый пользователь 2 - изменить текущего пользовател€
      case transformUser of
    1:
      begin
       with RemoteDesktopSetting.ListView1.Items.Add do
       begin
       Caption:=LabeledEdit1.Text;
       subitems.add(Combobox1.Text);
       ListPassword.Add(LabeledEdit3.Text);
       end;

      end;
    2:
      begin
      RemoteDesktopSetting.ListView1.Items[UserIndexSelected].Caption:=LabeledEdit1.Text;
      RemoteDesktopSetting.ListView1.Items[UserIndexSelected].SubItems[0]:=Combobox1.Text;
      ListPassword[UserIndexSelected]:=LabeledEdit3.Text;
      end;


      end;
 end;

end;

end.
