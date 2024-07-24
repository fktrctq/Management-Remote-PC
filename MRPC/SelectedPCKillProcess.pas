unit SelectedPCKillProcess;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TOKRightDlg123456789101112131415161718192021 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    CheckBox1: TCheckBox;
    Label1: TLabel;
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg123456789101112131415161718192021: TOKRightDlg123456789101112131415161718192021;
  GroupKillProcess:TThread;
  MultiThread:boolean;
  GroupKillProcessAr:array of Tthread;

implementation
uses umain,unit3;
{$R *.dfm}

procedure TOKRightDlg123456789101112131415161718192021.CancelBtnClick(
  Sender: TObject);
begin
MultiThread:=false;
end;

procedure TOKRightDlg123456789101112131415161718192021.CheckBox1Click(
  Sender: TObject);
begin
MultiThread:=not CheckBox1.Checked;
end;

procedure TOKRightDlg123456789101112131415161718192021.FormShow(
  Sender: TObject);
begin
label1.Visible:=false;
CheckBox1.Checked:=true;
MultiThread:=not CheckBox1.Checked; //// признак остановки процесса в одном потоке
end;

procedure TOKRightDlg123456789101112131415161718192021.OKBtnClick(
  Sender: TObject);
var
i:integer;

begin
/// SelectedPCkillProc  - это список компьютеров созданный в главной форме
if checkBox1.Checked then
  begin             //// остановка процесса на списке компьютеров в одном потоке
  GroupPC:=true;
  MultiThread:=false;
  GroupselectProc:=LabeledEdit1.Text;
  GroupKillProcess:=unit3.KillProcess.Create(true);
  GroupKillProcess.FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
  GroupKillProcess.start;
  end;

  if checkBox1.Checked=false then
    begin
      GroupPC:=true;
      MultiThread:=true; //// остановки процесса на списке компов в разных потоках
      label1.Caption:='0%';
      label1.Visible:=true;
      SetLength(GroupKillProcessAr,SelectedPCkillProc.Count);
      GroupselectProc:=LabeledEdit1.Text;
      for I := 0 to SelectedPCkillProc.Count-1 do
       begin
         if SelectedPCkillProc[i]<>'' then
            begin
            NewProcMyPS:=SelectedPCkillProc[i];
            GroupKillProcessAr[i]:=unit3.KillProcess.Create(true);
            GroupKillProcessAr[i].Priority:=tpLowest;//tpLower;
            GroupKillProcessAr[i].FreeOnTerminate := true;  //// под вопросом самоуничтожение потока
            GroupKillProcessAr[i].start;
            label1.Caption:=inttostr(((100 div SelectedPCkillProc.Count-1)*i+1))+'%';
            Application.ProcessMessages;
            if MultiThread=false then break;
            sleep(500);
            end;
       end;
    SelectedPCkillProc.free; //// уничтожение списка компьютеров для завершения процесса
    end;


close;
end;

end.
