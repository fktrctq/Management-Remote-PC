unit TaskSelect;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Imaging.pngimage, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ImgList,FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls, Vcl.Menus;

type
  TSelectTask = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ListView1: TListView;
    CancelB: TButton;
    OkB: TButton;
    StaticText1: TStaticText;
    Image1: TImage;
    ImageList2: TImageList;
    Button1: TButton;
    Button2: TButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CancelBClick(Sender: TObject);
    procedure OkBClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
  private
    { Private declarations }
    function readTaskTables(s:string):boolean;
    function AddOrUpdateListPcForTask(AddOrUpdate,group:boolean):boolean; // ���������� ��� ������ ������ ������ � ������
  public
    { Public declarations }
  end;

var
  SelectTask: TSelectTask;

implementation
uses MyDM,TaskEdit,umain;
{$R *.dfm}

function Log_write(fname, text:string):string;
var f:TStringList;        /// ������� ������ � ��� ����
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
exit;
end;
end;

function TSelectTask.readTaskTables(s:string):boolean;
var
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
begin
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRead.Options.AutoCommit:=false;
TransactionRead.Options.AutoStart:=false;
TransactionRead.Options.AutoStop:=false;
TransactionRead.Options.ReadOnly:=true;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=DataM.ConnectionDB;
TransactionRead.StartTransaction;
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT * FROM TABLE_TASK WHERE STATUS_TASK<>''delete''';
FDQueryRead.Open;
ListView1.Clear;
 while not FDQueryRead.Eof do
 begin
   with ListView1.Items.Add do
   begin
   if strtobool(FDQueryRead.FieldByName('START_STOP').AsString) then
   ImageIndex:=0 else ImageIndex:=1;
   caption:=FDQueryRead.FieldByName('DESCRIPTION_TASK').AsString; // �������� ������
   if (not FDQueryRead.FieldByName('WORKS_THREAD').AsBoolean)and  // ���� ����� ����������
    (FDQueryRead.FieldByName('STATUS_TASK').AsString='��������� ������') then SubItems.add('�����������') //���� ����� �������� ���������  false �� ������ ����������� ��� ���������
   else SubItems.add(FDQueryRead.FieldByName('STATUS_TASK').AsString); // ������ ������ (�����������/�����������/���������)
   SubItems.add(FDQueryRead.FieldByName('CURRENT_TASK').AsString);     // ������� �������
   SubItems.add(FDQueryRead.FieldByName('PC_RUN').AsString);      // �� ����� ���������� ������ �����������
   SubItems.add(FDQueryRead.FieldByName('COUNT_PC').AsString);    // ����� ����������� � ������
   end;
 FDQueryRead.Next;
 end;
TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
Result:=true;
except on E: Exception do
 begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
      Log_write('TASK',timetostr(now)+' ������ ������ ����� : '+e.Message);
  result:=false;
  end;
end;
end;

procedure TSelectTask.Button1Click(Sender: TObject);
begin                          // ��������
readTaskTables('');
end;



procedure TSelectTask.CancelBClick(Sender: TObject);
begin
close;
end;

procedure TSelectTask.FormShow(Sender: TObject);
begin                    //�������� ������ ��� �������� �����
readTaskTables('');
end;

function NewNameTask(s:string):string;
var
nametask:string;
SelectedPC:Tstringlist;
i:integer;
WaitName:boolean;
begin
WaitName:=false;
while not WaitName do
Begin
if not InputQuery('������� ��� ����� ������', '��� ����� ������', nametask)
        then break;
     if (nametask='') then break;
if SelectTask.ListView1.Items.Count>0 then
begin
for I := 0 to SelectTask.ListView1.Items.Count-1 do
begin
  if nametask=SelectTask.ListView1.Items[i].Caption then
  begin
    ShowMessage('������ � ����� ������ ��� ����������');
    WaitName:=false;
    nametask:='';
    Break;
  end
  else WaitName:=true;
end;
end else WaitName:=true;
End;
result:=nametask;
end;

procedure TSelectTask.N1Click(Sender: TObject);  // ����� ������
var
nametask:string;
WaitName,PCinList:boolean;
i:integer;
SelectedPC:TstringList;
begin
readTaskTables(''); // ��������� ������ �����
nametask:=NewNameTask('');
 if nametask='' then
 begin
 exit;
 end;

EditTask.ListView2.Clear; // ������� ������ ������
EditTask.ListView1.Clear; // ������� ������ �������
 try
  SelectedPC:=TstringList.Create;
  SelectedPC:=frmdomaininfo.createListpcForCheck('');
  for I := 0 to SelectedPC.count-1 do  //���������� ������ ������
  begin
    with EditTask.ListView2.Items.Add do
    begin
    ImageIndex:=0;
    Caption := inttostr(EditTask.ListView2.Items.Count);
    SubItems.Add(SelectedPC[i]);
    end;
  end;
  finally
  SelectedPC.Free;
  end;

 Edittask.Caption:=nametask;
 EditTask.TaskName.Caption:=''; // ������� label �.�. ������ �����
 EditTask.Button8.Enabled:=false; //������ ������� �� ������� , ������������ ����� ���������� ������
 Edittask.Showmodal;
end;

procedure TSelectTask.N2Click(Sender: TObject); // ������������� ������
begin
if ListView1.SelCount=1 then
begin
  EditTask.Caption:= ListView1.Selected.Caption; // ��� ������ � ���������
  EditTask.TaskName.Caption:=ListView1.Selected.Caption; // ��� ������ � label �.�. ���������� �������������� ������
  //EditTask.ReadTableSelectedTask(EditTask.TaskName.Caption); // ��������� ������ ������� � � label ������ �������� ������� �� �� ���
  EditTask.ReadResulTask(EditTask.TaskName.Caption); // ����� ������� ������ �������
  EditTask.Button8.Enabled:=true;
  if Edittask.WindowState=wsMinimized then // ���� ����� �������� �� ��������������� ��
   Edittask.WindowState:=wsNormal
   else                                        // ����� ����������
   Edittask.Show;
end;
end;

procedure TSelectTask.N3Click(Sender: TObject);  // ������ ������
begin
if ListView1.SelCount=1 then
begin
if EditTask.StatusStartStopTask(ListView1.Selected.Caption) then // ���� ������ �������� (true)
 begin
   ShowMessage('������ '+ListView1.Selected.Caption+' ��� ��������');
   exit;
 end;
if  EditTask.StatrTask(ListView1.Selected.Caption) then
 frmdomaininfo.Memo1.Lines.Add('������ ������ '+ListView1.Selected.Caption+' ������� ��������');
end;
readTaskTables(''); // �������� ������ �����
end;

procedure TSelectTask.N4Click(Sender: TObject); // ��������� ������
begin
if ListView1.SelCount=1 then
begin
if not EditTask.StatusStartStopTask(ListView1.Selected.Caption) then // ���� ������ ����������� (false)
 begin
   ShowMessage('������ ��� �����������');
   exit;
 end;
if edittask.StopTask(ListView1.Selected.Caption) then
frmDomainInfo.Memo1.Lines.Add('��������, ��������� ������ '+ListView1.Selected.Caption+' ������� ���������');
end;
readTaskTables(''); // �������� ������ �����
end;

procedure TSelectTask.N5Click(Sender: TObject); // �������� ������ �����
begin
readTaskTables('');
end;

procedure TSelectTask.OkBClick(Sender: TObject);
begin
if SelectTask.tag=1 then
AddOrUpdateListPcForTask(true,true); // �������� ������ ������ � ������
if SelectTask.tag=0 then
AddOrUpdateListPcForTask(true,false); // �������� ������ ������ ������� ������
end;

procedure TSelectTask.Button2Click(Sender: TObject);
begin
if SelectTask.tag=1 then
AddOrUpdateListPcForTask(false,true); // ��������� ������ ������ � ������
if SelectTask.tag=0 then
AddOrUpdateListPcForTask(false,false); // ��������� � ������ ������� ���������
end;


function TSelectTask.AddOrUpdateListPcForTask(AddOrUpdate,Group:boolean):boolean; // AddOrUpdate - �������� ��� ��������, Group - ��� ������ ������ ��� ��� ������ �����
var                 //���� AddOrUpdate �� ��������� ������ ������ � ������
SelectedPC:tstringlist;
i,z:integer;
PCinList:boolean;
begin
if ListView1.SelCount=1 then
begin
  if EditTask.StatusStartStopTask(ListView1.Selected.Caption) then
   begin
     showmessage('������ '+ListView1.Selected.Caption+' ��� ��������'+#10#13+' ���������� �� ��� �������� ���������.');
     exit;
   end;
  EditTask.Caption:= ListView1.Selected.Caption; // �������� ������ � ���������
  EditTask.TaskName.Caption:=ListView1.Selected.Caption; // ��� ������ � label �.�. ���������� �������������� ������
  EditTask.ReadResulTask(EditTask.TaskName.Caption); // ����� ������� ������ �������
  EditTask.Button8.Enabled:=true; //
 if AddOrUpdate then EditTask.deletelistPC(false); // ������� ���� ������ ������
  try
  if Group then // ���� ������ ��� ������ ������
  begin
    SelectedPC:=TstringList.Create;
    SelectedPC:=frmdomaininfo.createListpcForCheck('');
    if SelectedPC.Count=0 then
    begin
    if frmdomaininfo.listview8.SelCount=1 then
     SelectedPC.Add(frmdomaininfo.listview8.Selected.SubItems[0])
     else
     begin
      ShowMessage('�� ������� ������ �����������!');
      SelectedPC.Free;
      exit;
     end;
    end;
    for I := 0 to SelectedPC.count-1 do  //���������� ������ ������
    begin
    PCinList:=false;
     for z := 0 to EditTask.ListView2.Items.Count-1 do // ����� ������ ���� �� ���� ���� � ������ ��� ������
      if SelectedPC[i]=EditTask.ListView2.Items[z].SubItems[0] then // ���� ���� ��� � ������
      begin
      PCinList:=true;
      break;
      end;
      if not PCinList then // ���� ����� ��� � ������ �� ��������
      with EditTask.ListView2.Items.Add do
      begin
      ImageIndex:=0;
      Caption := inttostr(EditTask.ListView2.Items.Count);
      SubItems.Add(SelectedPC[i]);
      end;
    end;
  end;
  finally
  SelectedPC.Free;
  end;
  if not Group then  // ���� ������ ��� �������� ������ �����
  begin
   PCinList:=false;
     for z := 0 to EditTask.ListView2.Items.Count-1 do // ����� ������ ���� �� ���� ���� � ������ ��� ������
      if frmdomaininfo.combobox2.text=EditTask.ListView2.Items[z].SubItems[0] then // ���� ���� ��� � ������
      begin
      PCinList:=true;
      break;
      end;
      if not PCinList then // ���� ����� ��� � ������ �� ��������
      with EditTask.ListView2.Items.Add do
      begin
      ImageIndex:=0;
      Caption := inttostr(EditTask.ListView2.Items.Count);
      SubItems.Add(frmdomaininfo.combobox2.text);
      end;
  end;

  EditTask.Button7.Click; // ��������� ��� ���������
  SelectTask.Close; // �������� ��� ��� ����

  EditTask.ShowModal;     // ���������� ���� ������

end;

end;


end.

