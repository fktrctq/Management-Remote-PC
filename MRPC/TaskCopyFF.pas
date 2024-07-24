unit TaskCopyFF;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TTaskCopyDelFF = class(TForm)
    EditSource: TLabeledEdit;
    EditDestination: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    CheckBox1: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    procedure selectedFolder;
    Procedure selectedFile;
  public
    { Public declarations }
  end;

var
  TaskCopyDelFF: TTaskCopyDelFF;

implementation
uses TaskEdit;
{$R *.dfm}

procedure TTaskCopyDelFF.Button2Click(Sender: TObject);
function findElement(s:string):string;
begin
  if AnsiPos(':',s)=2 then // ���� � ���� ���������� ������� ��������� �� ����������� ������������ ����� ����, ��� ����������� � ���� ����� ����� �������� $
   begin
   delete(s,2,1); // ������� ������ :
   insert('$',s,2); // ��������� �� ��� ����� $
   end;
 result:=s;
end;
begin
with TFileOpenDialog.Create(self) do
begin
 try
  Name:='DlgCopyFolder';
  Title:='����� �������� ����������';
  Options:=[fdoForceShowHidden,fdoPickFolders]; {��������}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then EditDestination.Text:=findElement(FileName)+'\';
 finally
 Destroy;
 end;
end;
end;

function StrCopySourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
result:=copy(br,1,pos(' -> ',br)-1);
end;

function StrCopyDestination(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(' -> ',br)+3);
result:=br;
end;


procedure TTaskCopyDelFF.Button3Click(Sender: TObject);
 var
 i,z:integer;
begin
if EditSource.Text='' then
begin
  ShowMessage('�� ������� ���������� ��������');
  exit;
end;
if EditDestination.Text='' then
begin
  ShowMessage('�� ������� ���������� ����������');
  exit;
end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='���������� :'+EditSource.Text+' -> '+EditDestination.Text then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:=BoolToStr(CheckBox1.Checked); // ��������� � ��������� ������� ���������� ��� ���
    SubItems.Add('���������� :'+EditSource.Text+' -> '+EditDestination.Text);    // ��������
    SubItems.Add('CopyFF');
  end;
  // ����� �������� ������� ���������� �����
 //ShowMessage(StrCopySourse('���������� :'+EditSource.Text+' -> '+EditDestination.Text));
// ShowMessage(StrCopyDestination('���������� :'+EditSource.Text+' -> '+EditDestination.Text));
end;

procedure TTaskCopyDelFF.selectedFolder;
var
i:integer;
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� �������� ��� �����������';
  Options:=[fdoForceShowHidden,fdoPickFolders]; {��������}  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     EditSource.text:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
end;

Procedure TTaskCopyDelFF.selectedFile;
var
i:integer;
begin
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='����� ����� ��� �����������';
  Options:=[fdoForceShowHidden];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     EditSource.text:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
end;

procedure TTaskCopyDelFF.Button1Click(Sender: TObject);
begin
if TaskCopyDelFF.Caption='���������� ����' then  selectedFile;
if TaskCopyDelFF.Caption='���������� �������' then  selectedFolder;

end;

end.
