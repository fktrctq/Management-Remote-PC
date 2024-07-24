unit TaskEdit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.ImgList, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Vcl.DBCtrls, Vcl.Menus,IniFiles,DateUtils,ActiveX,ComObj,CommCtrl;

type
  TEditTask = class(TForm)
    Panel1: TPanel;
    ListView1: TListView;
    Button1: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel2: TPanel;
    ListView2: TListView;
    Button2: TButton;
    Button3: TButton;
    ImageList1: TImageList;
    Button4: TButton;
    Button5: TButton;
    FDQueryProcMSI: TFDQuery;
    FDTransactionReadProcMSI: TFDTransaction;
    PopupMSI: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Button7: TButton;
    TaskName: TLabel;
    Button8: TButton;
    ImageList2: TImageList;
    Button6: TButton;
    Button9: TButton;
    PopupProc: TPopupMenu;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Button10: TButton;
    PopupPower: TPopupMenu;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    WakeOnLan1: TMenuItem;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button11: TButton;
    PopupDetalTask: TPopupMenu;
    N10: TMenuItem;
    N11: TMenuItem;
    PopupDetalTaskPC: TPopupMenu;
    N12: TMenuItem;
    N13: TMenuItem;
    Button12: TButton;
    PopupService: TPopupMenu;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    Button13: TButton;
    Button14: TButton;
    PopupActivation: TPopupMenu;
    PopupFileFolder: TPopupMenu;
    Windows1: TMenuItem;
    Office1: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    Windows71: TMenuItem;
    Win881101: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    Button15: TButton;
    Button16: TButton;
    PopupRegEdit: TPopupMenu;
    N25: TMenuItem;
    N26: TMenuItem;
    N27: TMenuItem;
    N28: TMenuItem;
    N29: TMenuItem;
    SZ1: TMenuItem;
    Button17: TButton;
    PopupRestore: TPopupMenu;
    N30: TMenuItem;
    N31: TMenuItem;
    N32: TMenuItem;
    procedure Button2Click(Sender: TObject);  /// ��������� combobox ������� �����
    procedure FormGrClose(Sender: TObject; var Action: TCloseAction);
    procedure FormGRShow(Sender: TObject);
    procedure ButOkClick(Sender: TObject);
    procedure LVPCColumnClick(Sender: TObject;Column: TListColumn);
    procedure LVPCDblClick(Sender: TObject); // ���������� ����� ������� ������
    procedure LVPCSelectItem(Sender: TObject; Item: TListItem; // ��������� shift
  Selected: Boolean);
    procedure ButCloseFormGRClick(Sender: TObject);
    procedure ListGrSelect(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ListView2ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListGrKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure ListGrKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnShow(Sender: TObject);
    function creatformNewtask(S:string):boolean; // �������� ����� ��� ��������� ��������� � ��������
    function creatformUserPass(S:string):boolean; // �������� ����� ��� ������ � ������
    function creatformService(S:string):boolean; // �������� ����� ��� �����
    procedure ButtonAddUser(Sender: TObject); // ��������� ������������ � ������
    procedure FormMsiOpen(Sender: TObject); // ������ ���������� MSI
    procedure FormProcOpen(Sender: TObject); // ������ ���������� ��������
    procedure AddTaskInListview(Sender: TObject); /// ����� �������� ��� msi �� ������
    procedure SelectComboMsiProc(Sender: TObject);
    procedure SelectComboProcMSI(Sender: TObject);
    procedure AddServiceInListview(Sender: TObject); // ��������� ������ � ������ �����
    procedure N1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject); // ��� ������ �� ������ ��������� ��� msi �������� ��� ����� �� ������ combobox
    function CreateNewTableTask(TableName:string):boolean;
    function UpdateTableTABLETASK(NameTable:string;countPC:integer):boolean;     //���������� ������� ��� ������
    procedure Button7Click(Sender: TObject);     //������� ������� �������� ����� �������+�� ����������
    procedure SaveTask(Sender: TObject);
    function AddNewItemsTable(DescriptionTable:string;countPC:integer):string;
    function AddOrUpdateListViewreadInTableTask(nameTable:string):boolean; // ������ ������ � ������� Table_task � ����� ���������� � ��� � ������ ListView
    procedure Button8Click(Sender: TObject);
    function UpdateTableTask(TableName:string):boolean; // ���������� ������� ��� ������
    function TransactionAutoCommit(auto:boolean):boolean;
    procedure ListView1Changing(Sender: TObject; Item: TListItem;
      Change: TItemChange; var AllowChange: Boolean);
    procedure ListView2Changing(Sender: TObject; Item: TListItem;
      Change: TItemChange; var AllowChange: Boolean);
    procedure Button6Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button9Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N3Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure WakeOnLan1Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    Function CreateFormDetail(s:string;z:boolean):boolean;
    procedure FormDetailPCShow(Sender: TObject); // ��������� ���������� � ���������� ������� �� ���������� ��
    procedure FormDetailTaskShow(Sender: TObject); // ��������� ���� � ���������� ������
    procedure ButDetailSaveClick(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure ListView2DblClick(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure ListView1DblClick(Sender: TObject);
    function LoadListService(NamePC:string):boolean; /// ��������� ������ �����
    procedure ButLoadService(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure Windows1Click(Sender: TObject);
    procedure Windows71Click(Sender: TObject);
    procedure Win881101Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
    function  deletelistPC(z:bool):boolean;
    procedure Button15Click(Sender: TObject);
    procedure N27Click(Sender: TObject);
    procedure N28Click(Sender: TObject);
    procedure SZ1Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure N31Click(Sender: TObject);
    procedure N32Click(Sender: TObject);
    //procedure N27Click(Sender: TObject); //������� ���� ������ ������
 private
    FormKey:Tform;
    EditsSubKeyName,EditNameKey:TlabeledEdit;
    ComboRootKey:TcomboBox;
    ButtonOk,ButtonNo:Tbutton;
    function NewsSubKeyName(CreateDel:string):boolean; //����� ��� �������� �������
    function DeleteKeyName:boolean; // ������� ����� ��� �������� ����� �������
    procedure ButtonNoClose(Sender: TObject);
    procedure FormKeyClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonOksSubKeyName(Sender: TObject); // �������� � ������ ������� ��� �������� �������
    procedure ButtonOkDelKey(Sender: TObject); // �������� ������� �������� �����
    function DateForTaskRestore(var DateRestore:string):boolean; // ���� ��� �������������� �����
    function DescriptionForRestore(var DesRestore:string):boolean; //�������� ��� ����� ����� ��������������
  public
    function StatusStartStopTask(DescriptionTask:string):boolean; // �������� �������� ��� ����������� ������
    function nametableForDescription(Descript:string):string;   // �������� ��� ������� �� ��������
    function ReadResulTask(DescriptionTable:string):bool;  // ������ ������� � ����������� ���������� ������
    function StatrTask(DescriptionTable:string):boolean;  // ������ ������
    function AddUserPass(DescriptionTable,User,Passwd:string;SavePass:boolean):boolean; // ���������� ������������ � ������
    function StopTask(DescriptionTask:string):boolean;    // ��������� ������
    function RenameTableForTask(NameTable,NewDescriptionTable:string):boolean; //������ �������� ������� ��� ������. �������� ������ � ������� TABLE_TASK
    function ThereAnyRunStopTask(statustask:boolean):integer; /// ����� ������ ������, � ������� ���� �� ���������� ��� �������������

  end;


var
 EditTask: TEditTask;


implementation
uses umain,MyDM,RunTask,TaskNewMSI,TasknewProc,TaskSelectDelMSI,TaskCopyFF,RegEditKeySave;

var
  FormGr:Tform;
  ButOk:Tbutton;
  ButLoad:Tbutton;
  ButClose:Tbutton;
  PanelB:Tpanel;
  PanelG:Tpanel;
  LVPC:Tlistview;
  ListGr:Tcombobox;
  ComboNumMisProc:Tcombobox;
  SortLV:boolean;
  sortInt:integer;
  EditUser:Tedit;
  EditPass:Tedit;
  ChekSave:TcheckBox;
{$R *.dfm}

///////////////////////////////////////////////////////////
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
//////////////////////////////////////////////////////////////////////////
function TEditTask.TransactionAutoCommit(auto:boolean):boolean;
begin
//if auto then FDTransactionReadProcMSI.Connection.StartTransaction
//else FDTransactionReadProcMSI.Commit;
//FDTransactionReadProcMSI.Options.AutoCommit:=auto;
//FDTransactionReadProcMSI.Options.AutoStart:=auto;
//FDTransactionReadProcMSI.Options.AutoStop:=auto;
end;
///////////////////////////////////////////////////////////////////////////

function CustomSortProc(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
  if ParamSort=0 then
    Result := CompareText(Item1.Caption,Item2.Caption)
  else
    if Item1.SubItems.Count>ParamSort-1 then
    begin
      if Item2.SubItems.Count>ParamSort-1 then
          case SortLV of
          true:Result := CompareText(Item1.SubItems[ParamSort-1],Item2.SubItems[ParamSort-1]);
         false:Result := CompareText(Item2.SubItems[ParamSort-1],Item1.SubItems[ParamSort-1]);
          end
      else
        Result := 1;
    end
    else
      Result:=-1;
end;

function SortThirdSubItemAsInt(Item1, Item2: TListItem; ParamSort: integer): integer; stdcall;
begin
   Result := 0;
   if strtoint( Item1.Caption ) > strtoint( Item2.Caption ) then
      Result := ParamSort
   else
   if strtoint( Item1.Caption ) < strtoint( Item2.Caption) then
      Result := -ParamSort;
end;

procedure TEditTask.FormClose(Sender: TObject; var Action: TCloseAction);
var
i:integer;
begin
PageControl1.TabIndex:=0; // ��������� �� ������ ������� ��� ���� ����� �� �������������� Listview2 ��� �������� ��������
for I := ListView2.Columns.Count-1 downto 2 do   // ������� ������� �������� ��� ����������� ���������� ������� , ��������� ������ 2 ������� � � ��� �����
 begin
  ListView2.Columns.Delete(i);
 end;

end;

procedure TEditTask.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
i:integer;
YesPC,YesTask:boolean;
begin
YesPC:=false;
YesTask:=false;
for I := 0 to ListView2.Items.Count-1 do
begin
  if ListView2.Items[i].ImageIndex=4 then
  begin
  YesPC:=true; //���� ����������� �����
  break;
  end;
end;
for I := 0 to ListView1.Items.Count-1 do
begin
  if ListView1.Items[i].ImageIndex=1 then
  begin
  Yestask:=true; //���� ����������� ������
  break;
  end;
end;
canclose:=true; // �� ������ ������, ����� ��� �������� �� �������.
if (not YesPC)and (not YesTask) then canclose:=true; // ���� ��� ����������� ����� � ������ �� ���������
if (not YesPC)and (YesTask) then // ���� ��� ������ �� ���� ����������� ������
begin
ShowMessage('������ ����������� ����, �������� ���������� � ��������� ������.');
canclose:=false;
end;
if (YesPC)and (not YesTask) then // ���� ��� ����� �� ���� ����
begin
ShowMessage('������ ������� ����, �������� ������� � ��������� ������.');
canclose:=false;
end;

if (YesPC)and (YesTask) then canclose:=true; //���� ���� ����������� ������ � ����� �� �������

end;

procedure TEditTask.FormGrClose(Sender: TObject; var Action: TCloseAction);
begin
if sender is Tform then
(sender as Tform).Release;
end;

procedure TEditTask.FormGRShow(Sender: TObject);
var
i:integer;
begin
ListGr.Clear;
if frmDomainInfo.ComboBox1.enabled then   // ���� comboBox �� ������������
begin                                         // ��������� ������ ����� � ��������� ��� ����������� ������ � AD
for I := 0 to frmDomainInfo.ComboBox1.Items.Count-1 do
 ListGr.Items.Add(frmDomainInfo.ComboBox1.Items[i]);
ListGr.Text:=frmDomainInfo.ComboBox1.Text;
ListGr.OnSelect(ListGr);
end
else  /// �����, ������������ ��� ������ IP ��� ������ ������,
begin /// ��� ��������� �� ���������������� ��� � ��������� ����� ������� ���� � ������
   for I := 0 to frmDomainInfo.ListView8.Items.Count-1 do
   begin
     with LVPC.Items.Add do
      Caption:=frmDomainInfo.ListView8.Items[i].SubItems[0];
   end;
ListGr.Text:='������ ������������ Active Directory �������� � ������������������ ������';
ListGr.Enabled:=false;
end;


end;

function statusubtask(nametable:string;CountSubTask:integer):TstringList; // �������� ������ � ��� ������� �� ������, �������� ������ ����� �������
var
i,ok,no,warning,error,AllPC:integer;
res,MainRes:string;
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
begin
try
result:=TStringList.Create;
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
//TransactionRead.Options.DisconnectAction:=xdCommit; /// ���  xdCommit
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
FDQueryRead.SQL.Text:='SELECT * FROM '+nameTable;
FDQueryRead.Open;
for i := 0 to CountSubTask-1 do
begin
  ok:=0;
  no:=0;
  warning:=0;
  error:=0;
  AllPC:=0;
  res:='';
  mainres:='';
  while not FDQueryRead.Eof do
    begin
    AllPC:=AllPC+1;
    res:=FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).AsString;
    if res='OK' then ok:=ok+1;
    if res='NO' then no:=no+1;
    if res='WARNING' then warning:=warning+1;
    if res='ERROR' then error:=error+1;
    FDQueryRead.Next;
    end;
  if (no<>0)and(ok<>0) then MainRes:='�����������';
  if no=0 then MainRes:='���������';
  if ok=0 then MainRes:='������� ����������';
  if (no+ok+warning+error=AllPC) and (ok<>0) then MainRes:='���������';

                       // ������� ��   //������� �� ���������+������   // ����� ���������� ������ � ��������������
  result.Add(MainRes+'/'+inttostr(ok)+'/'+inttostr(no+warning+error)+'/'+inttostr(error+warning));
  FDQueryRead.First; // ��������� � ������ ������ ��� ����� �� ������ �������
end;


TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
except on E: Exception do
 begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
     Log_write('TASK',timetostr(now)+' ������ ������ ����������� ��������� ������� - '+e.Message);
     frmdomaininfo.Memo1.Lines.Add('������ ������ ����������� ��������� ������� - '+e.Message);
  end;
end;

end;


function TEditTask.ReadResulTask(DescriptionTable:string):bool; /// �������� �������� ������� ��� ������ ����������� ����������
var
FDQueryRead:TFDQuery;
TransactionRead: TFDTransaction;
i,z,step:integer;
nameTable,str:string;
ResStatSubtask:tstringlist;
BEGIN
try
TransactionRead:= TFDTransaction.Create(nil);
TransactionRead.Connection:=DataM.ConnectionDB;
//TransactionRead.Options.DisconnectAction:=xdCommit; /// ���  xdCommit
TransactionRead.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionRead.Options.AutoCommit:=false;
TransactionRead.Options.AutoStart:=false;
TransactionRead.Options.AutoStop:=false;
TransactionRead.Options.ReadOnly:=true;
FDQueryRead:= TFDQuery.Create(nil);
FDQueryRead.Transaction:=TransactionRead;
FDQueryRead.Connection:=DataM.ConnectionDB;
TransactionRead.StartTransaction;
nameTable:=nametableForDescription(DescriptionTable);//��� �������� ��� ������� �� ��������
TaskName.Caption:=nameTable; /// �� ����� �������� ������� ���������� �� ���
FDQueryRead.SQL.clear;
FDQueryRead.SQL.Text:='SELECT * FROM '+nameTable;
FDQueryRead.Open;
step:=2;
ListView1.Clear;
ListView2.Clear;
step:=3;
ResStatSubtask:=TStringList.Create;
ResStatSubtask:=statusubtask(nameTable,FDQueryRead.FieldByName('TASK_QUANT').AsInteger);
 for I := 0 to FDQueryRead.FieldByName('TASK_QUANT').AsInteger-1 do   // ���� �� �������� �������
 begin
   with ListView1.Items.add do
     begin
     ImageIndex:=1;
     caption:=inttostr(FDQueryRead.FieldByName('NUMTASK'+inttostr(i)).AsInteger);
     Subitems.add(string(FDQueryRead.FieldByName('NAMETASK'+inttostr(i)).AsString));
     SubItems.Add(string(FDQueryRead.FieldByName('TYPETASK'+inttostr(i)).AsString));
     str:='';
     str:=ResStatSubtask[i];
     SubItems.Add(copy(str,1,pos('/',str)-1)); // ������ �������
     System.delete(str,1,pos('/',str));        // ������� �� OK
     SubItems.Add(copy(str,1,pos('/',str)-1)); // �� ������� ������ ��������� ������ �� �������� OK
     system.delete(str,1,pos('/',str));        // ������� �� NO
     SubItems.Add(copy(str,1,pos('/',str)-1)); // �� ������� ������ �� ��������� (������� ERROR � WARNING)
     system.delete(str,1,pos('/',str));        // ������� �� ERROR � WARNING
     SubItems.Add(str);                        // �� ������� ������ ��������� �� �������� ERROR �  WARNING
     end;
   with ListView2.Columns.Add do // ��������� ������� ��� ������ ���������� ���������� � ListWiew � �������
   begin                         // ����� ��� �������� ���� �� �� ������
     Caption:=FDQueryRead.FieldByName('NAMETASK'+inttostr(i)).AsString;
     Width:=350;
   end;
   with ListView2.Columns.Add do // ��������� ������� ��� ������ ���������� ���������� � ListWiew � �������
   begin                // ����� ��� �������� ���� �� �� ������
     Caption:='������';
     Width:=70;
   end;

 end;
 step:=4;

while not FDQueryRead.Eof do /// ��������� ����� ������ //���� �� ������� �������
begin
  with ListView2.Items.Add do
  begin
    ImageIndex:=4;
    Caption:=inttostr(ListView2.Items.Count);
    SubItems.Add(FDQueryRead.FieldByName('PC_NAME').AsString);
    for I := 0 to FDQueryRead.FieldByName('TASK_QUANT').AsInteger-1 do // ��������� ���������� ���������� �������
    begin
    if FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).Value<>null then
      begin
      if FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).AsString ='NO'then SubItems.Add('������� ����������')
      else
      SubItems.Add(FDQueryRead.FieldByName('RESULTTASK'+inttostr(i)).AsString);
      end
    else SubItems.Add('');
    if FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).Value<>null then
    SubItems.Add(FDQueryRead.FieldByName('STATUSTASK'+inttostr(i)).AsString)
    else SubItems.Add('');
    end;

  end;
  FDQueryRead.Next
end;
step:=5;


ResStatSubtask.Free;
TransactionRead.commit;
FDQueryRead.sql.clear;
FDQueryRead.close;
FDQueryRead.Free;
TransactionRead.Free;
except on E: Exception do
  begin
     if Assigned(FDQueryRead) then FDQueryRead.Free;
     if Assigned(ResStatSubtask) then ResStatSubtask.Free;

      if Assigned(TransactionRead) then
      begin
      TransactionRead.commit;
      TransactionRead.Free;
      end;
     Log_write('TASK',timetostr(now)+' ������ ������ ����������� ���������� ������ - '+e.Message);
     frmdomaininfo.Memo1.Lines.Add('������ ������ ����������� ���������� ������ - '+e.Message);
  end;
end;
END;

procedure TEditTask.OnShow(Sender: TObject);
begin
PageControl1.TabIndex:=0;
Button7.Enabled:=false;
sortInt:=1;   // ��� ���������� �������
SortLV:=true; // ��� ���������� �������
end;

procedure TEditTask.ButOkClick(Sender: TObject);
var
groupList:tstringList;
i,z:integer;
PCinList:boolean;
begin
if LVPC.Items.Count>0 then
begin
  for I := 0 to LVPC.Items.Count-1 do
  Begin
  PCinList:=false;
  if LVPC.Items[i].Checked then
  begin
   for z := 0 to ListView2.Items.Count-1 do // ����� ������ ���� �� ���� ���� � ������ ��� ������
    if LVPC.Items[i].Caption=ListView2.Items[z].SubItems[0] then // ���� ���� ��� � ������
    begin
    PCinList:=true;
    break;
    end;
    if not PCinList then
     with ListView2.Items.Add do
     begin
      ImageIndex:=0;
      Caption:=inttostr(ListView2.Items.Count);
      SubItems.add(LVPC.Items[i].Caption);
     end;
  end;

  End;
end;
end;

procedure TEditTask.ButCloseFormGRClick(Sender: TObject);
begin
FormGr.close;
end;

procedure TEditTask.LVPCColumnClick(Sender: TObject;
  Column: TListColumn);
  var
  i:integer;
begin
if LVPC.Columns[0].ImageIndex=1 then
  begin
  for I := 0 to LVPC.Items.Count-1 do
   LVPC.Items[i].Checked:=true;
   LVPC.Columns[0].ImageIndex:=2;
  exit;
  end;
if LVPC.Columns[0].ImageIndex=2 then
  begin
  for I := 0 to LVPC.Items.Count-1 do
   LVPC.Items[i].Checked:=false;
   LVPC.Columns[0].ImageIndex:=1;
  exit;
  end;
end;

procedure TEditTask.LVPCDblClick(Sender: TObject);
var
z,i:integer;
PCinList:bool;
begin
PCinList:=false;
if LVPC.SelCount=1 then
  begin
   for z := 0 to ListView2.Items.Count-1 do // ����� ������ ���� �� ���� ���� � ������ ��� ������
    if LVPC.Selected.Caption=ListView2.Items[z].SubItems[0] then // ���� ���� ��� � ������
    begin
    PCinList:=true;
    break;
    end;
    if not PCinList then
     with ListView2.Items.Add do
     begin
      ImageIndex:=0;
      Caption:=inttostr(ListView2.Items.Count);
      SubItems.add(LVPC.Selected.Caption);
     end;
  end;
end;

procedure TEditTask.LVPCSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
if LVPC.SelCount>1 then
 if(item.Selected)and(item.Checked<>true) then item.Checked:=true;
end;


procedure TEditTask.ListGrSelect(Sender: TObject); /// ����� ������ ������ �� ������
var
i:integer;
ListPC:Tstringlist;
begin
 LVPC.Clear;
 LVPC.Columns[0].ImageIndex:=1; // ���������� �������� ���� �������� ���� ������
 ListPC:=TStringList.Create;
 try
 ListPC:=frmDomainInfo.GetAllGroupPC(ListGr.Text);
 for I := 0 to ListPC.Count-1 do
   begin
     with LVPC.Items.add do
     begin
       Caption:=ListPC[i];
     end;
   end;
 finally
   ListPC.free;
 end;
end;


procedure TEditTask.ListView1Changing(Sender: TObject; Item: TListItem;
  Change: TItemChange; var AllowChange: Boolean);
  var
  i:integer;
begin
try
Button7.Enabled:=true; // �������� ������ ���������
if ListView1.Items.Count>30 then
begin
  for I := ListView1.Items.Count Downto 30 do // ������� �� �������� �� 30�� �����
   begin
     ListView1.Items[i].Selected:=true; //��������
     Button4.Click;                    // �������
   end;
frmdomaininfo.Memo1.Lines.Add('��������� ������������ ��������� ������� � ������.');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� �������, ��������� ������������ ��������� ������� � ������');
end;
end;

procedure TEditTask.ListView1DblClick(Sender: TObject);
begin
if ListView1.SelCount=1 then  //   ���� ������� ���� �������
if ListView1.ItemFocused.ImageIndex=1 then     //���� ��� ������� ���������
CreateFormDetail(ListView1.ItemFocused.SubItems[0],false); // �������� �����
end;

procedure TEditTask.ListView2Changing(Sender: TObject; Item: TListItem;
  Change: TItemChange; var AllowChange: Boolean);
begin
Button7.Enabled:=true; // �������� ������ ���������
end;

procedure TEditTask.ListView2ColumnClick(Sender: TObject; Column: TListColumn);
begin

if Column.Index=1 then
begin
  SortLV:=not SortLV;
   ListView2.CustomSort(@CustomSortProc, Column.Index);
end;

if Column.Index=0 then
begin
  sortInt:=-sortInt;
  ListView2.CustomSort(@SortThirdSubItemAsInt, sortInt);
end;


end;




procedure TEditTask.ListView2DblClick(Sender: TObject);
begin
if listview2.SelCount=0 then exit;
if listview2.Selected.ImageIndex=4 then // ���� ImageIndex=4 �� ������ ��������� � ������� � ������� ���������� �� ������ ��
CreateFormDetail(EditTask.Caption,true);
end;

procedure TEditTask.Button11Click(Sender: TObject);
begin
creatformUserPass('');
end;

procedure TEditTask.Button15Click(Sender: TObject);  /// �������
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

WT:=2;
WiteTime:=InputBox('������� �������� �� ���������� �������', '����� (���):', '1');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('�� ������� �� ������ �������� ��������, ����������� �������� �� ��������!');
  exit;
end;
if WT=0 then
begin
  ShowMessage('�� ������� �� ������ �������� ��������, ����������� �������� �� ��������!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // ������� ��������
  SubItems.Add('������� '+inttostr(WT)+' ���.');    // ��������
  SubItems.Add('TimeOut');
end;
end;



function StrDeleteSourse(s:string):string;
var
br:string;
begin
br:=s;
Delete(br, 1, pos(':',br));
 if AnsiPos(':',br)=2 then // ���� � ���� ���������� ������� ��������� �� ����������� ������������ ����� ����, ��� ����������� � ���� ����� ����� �������� $
 begin
 delete(br,2,1); // ������� ������ :
 insert('$',br,2); // ��������� �� ��� ����� $
 end;
 result:=br;
end;


procedure TEditTask.ButDetailSaveClick(Sender: TObject);
begin
if LVPC.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(LVPC,'���������� ������ �����������','');
end;

procedure TEditTask.FormDetailPCShow(Sender: TObject);
var
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
i,z:integer;
ConnectionThread: TFDConnection;
begin
try
try
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // ������������ ���� ������
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.LoginPrompt:= false;  /// ����������� ������� user password
ConnectionThRead.Connected:=true;

TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThRead;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=ConnectionThRead;
for I := 0 to EditTask.ListView2.Items.Count-1 do
begin
if (EditTask.ListView2.Items[i].Selected) and(EditTask.ListView2.Items[i].ImageIndex=4) //���� ���� ��������� �� ����� ���������� ����������� ���������� ����� �� ����
then
  begin
  TransactionWrite.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='select * from '+nametableForDescription(EditTask.Caption)+' where PC_NAME='''+EditTask.ListView2.Items[i].SubItems[0]+'''';
  FDQuery.Open;
  FDQuery.First;
  while not FDQuery.Eof do
    begin
       for z := 0 to FDQuery.FieldByName('TASK_QUANT').AsInteger-1 do
        with LVPC.Items.Add do
        begin
        Caption:=FDQuery.FieldByName('PC_NAME').AsString;
        SubItems.Add(FDQuery.FieldByName('NAMETASK'+inttostr(z)).AsString);
        SubItems.Add(FDQuery.FieldByName('RESULTTASK'+inttostr(z)).AsString);
        SubItems.Add(FDQuery.FieldByName('STATUSTASK'+inttostr(z)).AsString);
        end;
    FDQuery.Next;
    end;

  end;

end;
finally
TransactionWrite.Commit;
ConnectionThRead.Close;
FDQuery.Close;
FDQuery.Free;
TransactionWrite.Free;
ConnectionThRead.Free;
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ '+e.Message)
end;
end;

procedure TEditTask.FormDetailTaskShow(Sender: TObject);
var
TransactionWrite:TFDTransaction;
FDQuery:TFDQuery;
ConnectionThread: TFDConnection;
begin
try
try
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // ������������ ���� ������
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP ��� local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.LoginPrompt:= false;  /// ����������� ������� user password
ConnectionThRead.Connected:=true;
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=DataM.ConnectionDB;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;;
TransactionWrite.Options.ReadOnly:=true;
FDQuery:= TFDQuery.Create(nil);
FDQuery.Transaction:=TransactionWrite;
FDQuery.Connection:=DataM.ConnectionDB;

  TransactionWrite.StartTransaction;
  FDQuery.SQL.Clear;
  FDQuery.SQL.Text:='select * from '+nametableForDescription(EditTask.Caption);
  FDQuery.Open;
  FDQuery.First;
  if ListView1.SelCount=1 then
  while not FDQuery.Eof do
    begin
        with LVPC.Items.Add do
        begin
        Caption:=FDQuery.FieldByName('PC_NAME').AsString;
        SubItems.Add(FDQuery.FieldByName('NAMETASK'+inttostr(ListView1.Selected.Index)).AsString);
        SubItems.Add(FDQuery.FieldByName('RESULTTASK'+inttostr(ListView1.Selected.Index)).AsString);
        SubItems.Add(FDQuery.FieldByName('STATUSTASK'+inttostr(ListView1.Selected.Index)).AsString);
        end;
    FDQuery.Next;
    end;

finally
TransactionWrite.Commit;
ConnectionThRead.Close;
FDQuery.Close;
FDQuery.Free;
TransactionWrite.Free;
ConnectionThRead.Free;
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ '+e.Message)
end;
end;

Function TEditTask.CreateFormDetail(s:string;z:boolean):boolean;
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:='FormListGroup';
FormGr.Caption:='�������� - '+s;
FormGr.Width:=800;
FormGr.Height:=400;
if z then FormGr.OnShow:=FormDetailPCShow
else FormGr.OnShow:=FormDetailTaskShow;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


PanelG:=Tpanel.Create(FormGr);
PanelG.Parent:=FormGr;
PanelG.Name:='PanelOK';
PanelG.Height:=35;
PanelG.Align:=alTop;
PanelG.Caption:='';

ButClose:=TButton.Create(FormGr);
ButClose.Parent:=PanelG;
ButClose.Name:='SaveList';
ButClose.Caption:='��������� ������';
ButClose.Top:=5;
ButClose.Left:=5;
ButClose.Width:=100;
ButClose.OnClick:=ButDetailSaveClick;

LVPC:=Tlistview.Create(FormGr);
LVPC.Parent:=FormGr;
LVPC.Name:='LVRES';
LVPC.Align:=alClient;
LVPC.ViewStyle:=vsReport;
LVPC.ReadOnly:=true;
LVPC.RowSelect:=true;
LVPC.MultiSelect:=true;
LVPC.GridLines:=true;
lvpc.Checkboxes:=false;

  with LVPC.Columns.Add  do
  begin
    Caption:='���������';
    Width:=100;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='�������';
    Width:=300;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='���������';
    Width:=300;
  end;
  with LVPC.Columns.Add  do
  begin
    Caption:='������';
    Width:=50;
  end;
//LVPC.OnColumnClick:= LVPCColumnClick;
FormGr.Showmodal;
end;

procedure TEditTask.Button2Click(Sender: TObject);
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:='FormListGroup';
FormGr.Caption:='������ ����� � �����������';
FormGr.Width:=500;
FormGr.Height:=300;
FormGr.BorderStyle:=bsDialog;
FormGr.OnShow:=FormGrShow;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;

PanelB:=Tpanel.Create(FormGr);
PanelB.Parent:=FormGr;
PanelB.Name:='PanelBUt';
PanelB.Width:=95;
panelB.Align:=alRight;
PanelB.Caption:='';
PanelB.TabOrder:=2;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=PanelB;
ButOk.Name:='AddPC';
ButOk.Caption:='��������';
ButOk.Top:=5;
ButOk.Left:=5;
ButOk.Width:=85;
ButOk.OnClick:=ButOkClick;
ButOk.TabOrder:=0;

ButClose:=TButton.Create(FormGr);
ButClose.Parent:=PanelB;
ButClose.Name:='ClFormGr';
ButClose.Caption:='�������';
ButClose.Top:=35;
ButClose.Left:=5;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=1;

PanelG:=Tpanel.Create(FormGr);
PanelG.Parent:=FormGr;
PanelG.Name:='PanelCombo';
PanelG.Height:=30;
PanelG.Align:=alTop;
PanelG.Caption:='';
PanelG.TabOrder:=0;

ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=PanelG;
ListGr.Name:='ListGr';
ListGr.Text:='';
ListGr.Align:=alClient;
ListGr.DroppedDown:=true;
ListGr.DropDownCount:=15;
ListGr.OnSelect:=ListGrSelect;
ListGr.TabOrder:=0;


LVPC:=Tlistview.Create(FormGr);
LVPC.Parent:=FormGr;
LVPC.Name:='LVPC';
LVPC.Align:=alClient;
LVPC.ViewStyle:=vsReport;
LVPC.ReadOnly:=true;
LVPC.RowSelect:=true;
LVPC.MultiSelect:=true;
LVPC.GridLines:=true;
lvpc.Checkboxes:=true;
LVPC.SmallImages:=ImageList1;
with LVPC.Columns.Add  do
begin
  Caption:='��� ����������';
  Width:=250;
  ImageIndex:=1;
end;
LVPC.OnColumnClick:= LVPCColumnClick;
LVPC.OnDblClick:=LVPCDblClick;
LVPC.OnSelectItem:=LVPCSelectItem;
LVPC.TabOrder:=1;
FormGr.Showmodal;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TEditTask.N10Click(Sender: TObject);
begin
if listview1.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(listview1,'���������� ������ �������','');
end;

procedure TEditTask.N11Click(Sender: TObject); // �������� � ������� � ������
begin
if ListView1.SelCount=1 then
if ListView1.Selected.ImageIndex<>0 then // ���� image<>0 (�����) ������ ������ ���� ��������� � ������� � ������� ��������� � ���������� �� ������ ��
CreateFormDetail(ListView1.ItemFocused.SubItems[0],false);
end;

procedure TEditTask.N12Click(Sender: TObject);
begin
if listview2.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(listview2,'���������� ������ �����������','');
end;

procedure TEditTask.N13Click(Sender: TObject);
begin
if listview2.SelCount=0 then exit;
if listview2.Selected.ImageIndex=4 then // ���� ImageIndex=4 �� ������ ��������� � ������� � ������� ���������� �� ������ ��
CreateFormDetail(EditTask.Caption,true);
end;

procedure TEditTask.N14Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
 creatformService('��������� ������');
end;

procedure TEditTask.N15Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
creatformService('���������� ������');
end;

procedure TEditTask.N16Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
creatformService('������� ������');
end;

procedure TEditTask.N19Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
TaskCopyDelFF.Caption:='���������� �������';
TaskCopyDelFF.ShowModal;
end;

procedure TEditTask.N20Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
TaskCopyDelFF.Caption:='���������� ����';
TaskCopyDelFF.ShowModal;
end;

procedure TEditTask.N21Click(Sender: TObject);
var
patch:string;
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='������� ������� ��� ��������';
  Options:=[fdoPickFolders];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     patch:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
 if (patch='') then exit;
 if not InputQuery('������� ��� �������� ', '����:', patch) then exit;
  if (patch='') then
  begin
   showmessage('�� ������ �������');
   exit;
  end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='������� :'+patch then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='0'; //    ������� �������
    SubItems.Add('������� :'+patch);    // ����
    SubItems.Add('DelFF');
  end;

end;


procedure TEditTask.N22Click(Sender: TObject);
var
patch:string;
z,i:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
 with TFileOpenDialog.Create(self) do
 begin
 try
  Name:='DlgAddFolder';
  Title:='������� ���� ��� ��������';
  Options:=[fdoForceShowHidden];  //fdoAllowMultiSelect,fdoPickFolders,fdoForceFileSystem,fdoForceShowHidden
  if Execute then
    begin
     patch:=FileName;
    end;
 finally
 Destroy;
 end;
 end;
 if (patch='') then exit;
 if not InputQuery('���� ��� �������� ', '����:', patch) then exit;
  if (patch='') then
  begin
   showmessage('�� ������ ����');
   exit;
  end;

for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='������� :'+patch then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='1'; //    ������� ����
    SubItems.Add('������� :'+patch);    // ����
    SubItems.Add('DelFF');
  end;

end;



procedure TEditTask.N1Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
 creatformNewtask('msi'); // ������� ����� ��� ������� MSI
end;

procedure TEditTask.N30Click(Sender: TObject); /// ������� ����� ��������������
var
NameTask,DesTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
try

DesTask:='Management Remote PC';
if not DescriptionForRestore(DesTask) then  //�������� ����� ��������������
begin
 ShowMessage('�� �� ������� �������� ����� ��������������, �������� ��������...');
 exit;
end;

NameTask:='������� ����� �������������� :'+DesTask;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(1); //1- ������� ����� ��������������, 2 - ������������ �����
  SubItems.Add(NameTask); // �������� ������� + �������� ����� ��������������
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;

function TEditTask.DateForTaskRestore(var DateRestore:string):boolean; // ������� ���������� � ���������� DateRestore �������� ���� ���� ��� ���������
var
WaitName:boolean;
Rdate:TdateTime;
begin
try
DateRestore:=datetostr(now);
WaitName:=false;
while not WaitName do
Begin
if not InputQuery('������� ���� ����� ��������������', '����', DateRestore)
then
begin
DateRestore:='';
result:=false;
break;
end;
if (DateRestore='') then
  begin
  result:=false;
  break;
  end
else
 begin
  if TryStrToDate(DateRestore,Rdate) then
  begin
  result:=true;
  WaitName:=true;
  end
  else WaitName:=false;
 end;
End;
except
  begin
  ShowMessage('������ ������� ����');
  result:=false;
  end;
end;
end;

function TEditTask.DescriptionForRestore(var DesRestore:string):boolean;
var
WaitName:boolean;
begin
WaitName:=false;
while not WaitName do
BEGIN
if not InputQuery('������� �������� ����� ��������������', '��������', DesRestore)
 then
  Begin
  result:=false;
  break;
  End
 else
  Begin
  DesRestore:=trim(DesRestore);
  if (DesRestore='') then
   begin
   WaitName:=false;
   end
  else
   begin
   WaitName:=true;
   result:= true;
   end;
  End;
END;
end;

procedure TEditTask.N31Click(Sender: TObject); // ������������ �������
var
NameTask:string;
Rdate:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
try
NameTask:='������������ �������';
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if not DateForTaskRestore(Rdate) then // ���� ��� �������������� �������
begin
  ShowMessage('�� �� ������� ���� ����� ��������������, �������� ��������...');
  exit;
end;

i:=MessageDlg('��� �������������� ������� ����������� ������������ ����������.'+#10#13+'���������� ���������� ��������?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
// ����� ��������� ������ ��� ����� ����, �� ������� ��� ���� ����� ��� �������� ����� ������
// ���� ��� ����������� ����� �������������� ��������� � ��������
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(2); // 2 - ������������ ����� ������������ ����� ������, 1- ������� ����� ��������������,
  SubItems.Add(NameTask+' :'+Rdate); // �������� �������
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;


procedure TEditTask.N32Click(Sender: TObject);
var
NameTask:string;
Rdate,DesPoint:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
try
NameTask:='������������ �������';
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;


// ���� ��� ����������� ����� �������������� ��������� � ��������
if not DateForTaskRestore(Rdate) then // ���� ��� �������������� �������
begin
  ShowMessage('�� �� ������� ���� ����� ��������������, �������� ��������...');
  exit;
end;
// �������� ����� ��������������
DesPoint:='Management Remote PC';
if not DescriptionForRestore(DesPoint) then  //�������� ����� ��������������
begin
 ShowMessage('�� �� ������� �������� ����� ��������������, �������� ��������...');
 exit;
end;

i:=MessageDlg('��� �������������� ������� ����������� ������������ ����������.'+#10#13+'���������� ���������� ��������?', mtConfirmation,[mbYes,mbCancel],0);
if i=IDCancel then   exit;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(3); // 1- ������� ����� ��������������, 2 - ������������ ����� �� ����, 3- �������������� �� ���� � �������
  SubItems.Add(NameTask+' :'+Rdate+' -> '+DesPoint); // �������� �������
  SubItems.Add('RestoreWin');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;

procedure TEditTask.N3Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
creatformNewtask('proc');   // ������� ����� ��� ������� ��������
end;



procedure TEditTask.N2Click(Sender: TObject); // ��������� ����� ��� ����� ��������� MSI
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
NewMSiTask.ShowModal;
end;

procedure TEditTask.N4Click(Sender: TObject);  // ��������� ����� ��� ������ ��������
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
NewProcTask.ShowModal;
end;


procedure TEditTask.N5Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
if SelectDelMSITask.Caption<>'�������� �������� msi' then
begin
 SelectDelMSITask.Edit1.Text:='';
 SelectDelMSITask.ComboBox1.clear;
end;
SelectDelMSITask.Caption:='�������� �������� msi';
SelectDelMSITask.Edit1.TextHint:='��� ��� ����� ����� ��������� ��� ��������.';
SelectDelMSITask.ComboBox1.TextHint:='��������� ������ �������� � ���������� � ����';
SelectDelMSITask.ShowModal;
end;

procedure TEditTask.N6Click(Sender: TObject);
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

 if SelectDelMSITask.Caption<>'���������� ��������' then
begin
 SelectDelMSITask.Edit1.Text:='';
 SelectDelMSITask.ComboBox1.clear;
end;
SelectDelMSITask.Caption:='���������� ��������';
SelectDelMSITask.Edit1.TextHint:='��� �������� ��� ���������� notepad.exe.';
SelectDelMSITask.ComboBox1.TextHint:='��������� ������ ��������� � ���������� � ����';
SelectDelMSITask.ShowModal;
end;

procedure TEditTask.N7Click(Sender: TObject);  // ���������� ������
var
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='���������� ������') then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:='0'; //
  SubItems.Add('���������� ������ ������������');    // ��������
  SubItems.Add('Logout');
end;
end;

procedure TEditTask.N8Click(Sender: TObject); // ������������
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='������������') then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

WT:=4;
WiteTime:=InputBox('������� �������� ���������� ����� ������������', '����� (���):', '3');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('�� ������� �� ������ �������� ��������, ����������� �������� �� ��������!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // ������� ��������
  SubItems.Add('������������ ����������');    // ��������
  SubItems.Add('Reset');
end;
end;

procedure TEditTask.N9Click(Sender: TObject); // ���������� ������
var
i,z,WT:integer;
WiteTime:string;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]='���������� ������') then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
WT:=2;
WiteTime:=InputBox('������� �������� �� ���������� ����������', '����� (���):', '1');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('�� ������� �� ������ �������� ��������, ����������� �������� �� ��������!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // ������� �������� �� ���������� ������
  SubItems.Add('���������� ������');    // ��������
  SubItems.Add('ShDown');
end;
end;

procedure TEditTask.WakeOnLan1Click(Sender: TObject); // ��������� ����������
var
SetInI:TMeMiniFile;
BRaddress,WiteTime:string;
WT:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
BRaddress:=setini.readString('ConfLAN','broadcast','192.168.0.255');
IpBroadCast:=InputBox('������� Broadcast ����� ����� ����', 'IP-', BRaddress);
if IpBroadCast='' then
begin
  ShowMessage('�� �� ������� Broadcast ����� ����, ����������� �������� �� ��������!');
  SetInI.Free;
  exit;
end;
SetInI.WriteString('ConfLAN','broadcast',IpBroadCast);
SetInI.UpdateFile;
SetInI.Free;
WT:=4;
WiteTime:=InputBox('������� �������� ���������� ����� ���������', '����� (���):', '3');
if not TryStrToInt(WiteTime,WT) then
begin
  ShowMessage('�� ������� �� ������ �������� ��������, ����������� �������� �� ��������!');
  exit;
end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(WT); // ������� ��������
  SubItems.Add('��������� ����������. Broadcast :'+IpBroadCast);    // ��������, �� �������� � ������ ���������� broadcast �����
  SubItems.Add('WOL');
end;
end;



procedure TEditTask.Windows1Click(Sender: TObject);
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

StringKey:='';
try
   if not InputQuery('������ XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', '���� ��������:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('�� �� ����� ���� ��������� ��������');
   exit;
  end;

NameTask:='��������� Windows';
 if not InputQuery('��������� Windows', '�������� �������:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('�� �� ����� �������� �������');
   exit;
  end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask+' :'+StringKey then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(1); // ��������� windows
  SubItems.Add(NameTask+' :'+StringKey); // �������� � ������ ���������
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;

procedure TEditTask.Win881101Click(Sender: TObject);
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
StringKey:='';
try
   if not InputQuery('������ XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', '���� ��������:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('�� �� ����� ���� ��������� ��������');
   exit;
  end;

    NameTask:='��������� Office (Win 8/8.1/10...)';
 if not InputQuery('��������� Office', '�������� �������:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('�� �� ����� �������� �������');
   exit;
  end;

  for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=nametask+' :'+StringKey then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(2); // ��������� office ��� windows 8/8.1/10...
  SubItems.Add(nametask+' :'+StringKey); // �������� � ������ ���������
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;

procedure TEditTask.Windows71Click(Sender: TObject); // ��������� office ��� windows 7
var
StringKey,NameTask:string;
i,z:integer;
begin
if EditTask.ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
StringKey:='';
try
   if not InputQuery('������ XXXXX-XXXXX-XXXXX-XXXXX-XXXXX ', '���� ��������:', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('�� �� ����� ���� ��������� ��������');
   exit;
  end;

  NameTask:='��������� Office (Windows 7)';
 if not InputQuery('��������� Office', '�������� �������:', NameTask)
    then exit;
  if (NameTask='') then
  begin
   showmessage('�� �� ����� �������� �������');
   exit;
  end;

 for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=NameTask+' :'+StringKey then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;

with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=inttostr(3); // ��������� office ��� windows 7
  SubItems.Add(NameTask+' :'+StringKey); // �������� � ������ ���������
  SubItems.Add('Activate');
end;
except on E: Exception do  frmdomaininfo.Memo1.Lines.Add('������ ���������� ������� - '+e.Message);
end;
end;


procedure TEditTask.FormMsiOpen(Sender: TObject); // ��������� ������ ��������� msi ��������
begin
try
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''msi''';
FDQueryProcMSI.Open;
ListGr.Clear;
ComboNumMisProc.Clear;
 while not FDQueryProcMSI.Eof do
 begin
 ListGr.Items.Add(string(FDQueryProcMSI.FieldByName('DESCRIPTION_PROC').Value));
 ComboNumMisProc.Items.Add(string(FDQueryProcMSI.FieldByName('ID_PROC').Value));
 FDQueryProcMSI.Next;
 end;
FDQueryProcMSI.Close;
if ListGr.Items.Count>0 then
  begin
  ListGr.ItemIndex:=0;
  ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
  end;
except on E: Exception do
begin
   Log_write('TASK',timetostr(now)+' ������ ������ ������ ��������- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' - ������ ������ ������ ��������- '+e.Message);
end;
end;
end;



procedure TEditTask.FormProcOpen(Sender: TObject);  // ��������� ������� ��������� ���������
begin
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM START_PROC_MSI WHERE TYPE_PROC=''proc''';
FDQueryProcMSI.Open;
ListGr.Clear;
ComboNumMisProc.Clear;
 while not FDQueryProcMSI.Eof do
 begin
 ListGr.Items.Add(string(FDQueryProcMSI.FieldByName('DESCRIPTION_PROC').Value));
 ComboNumMisProc.Items.Add(string(FDQueryProcMSI.FieldByName('ID_PROC').Value));
 FDQueryProcMSI.Next;
 end;
FDQueryProcMSI.Close;
if ListGr.Items.Count>0 then
begin
 ListGr.ItemIndex:=0;
 ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
end;
end;

procedure TEditTask.AddTaskInListview(Sender: TObject);
var
i,z:integer;
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;

for I := 0 to ListView1.Items.Count-1 do
  begin
    if (ListView1.Items[i].SubItems[0]=ListGr.Text) and (ListView1.Items[i].Caption=ComboNumMisProc.Text) then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
with ListView1.Items.Add do
begin
  ImageIndex:=0;
  caption:=ComboNumMisProc.Text; // ����� �������� ��� msi
  if FormGr.Caption='��������� ���������' then
  begin
   SubItems.Add('��������� ��������� '+ListGr.Text);    // ��������
   SubItems.Add('msi');
  end;
  if FormGr.Caption='��������� ��������' then
  begin
  SubItems.Add('������ �������� '+ListGr.Text);    // ��������
  SubItems.Add('proc');
  end;
end;
end;

procedure TEditTask.AddServiceInListview(Sender: TObject);
var
i,z,step:integer;
begin
try
if ComboNumMisProc.Text='' then
begin
ShowMessage('������� ��� ������');
ComboNumMisProc.SetFocus;
exit;
end;

step:=0;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=FormGr.Caption+' :'+ComboNumMisProc.Text then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
 step:=1;
with EditTask.ListView1.Items.Add do
begin
step:=2;
ImageIndex:=0;
  if FormGr.Caption='��������� ������' then
  begin
   caption:=inttostr(1); // ������ ������
   SubItems.Add('��������� ������ :'+ComboNumMisProc.Text);    // ��� ������
   SubItems.Add('Service');
  end;
  if FormGr.Caption='���������� ������' then
  begin
  caption:=inttostr(2); // ��������� ������
  SubItems.Add('���������� ������ :'+ComboNumMisProc.Text);    // ��� ������
  SubItems.Add('Service');
  end;
  if FormGr.Caption='������� ������' then
  begin
  caption:=inttostr(3); // �������� ������
  SubItems.Add('������� ������ :'+ComboNumMisProc.Text);    // ��� ������
  SubItems.Add('Service');
  end;
end;
step:=3;
except on E: Exception do showmessage(E.Message);
end;
end;


procedure TEditTask.SelectComboMsiProc(Sender: TObject); // ��� ������ �� ������ ��������� ��� msi �������� ��� ����� �� ������ combobox
begin
ComboNumMisProc.ItemIndex:=ListGr.ItemIndex;
end;

procedure TEditTask.SelectComboProcMSI(Sender: TObject); // ��� ������ �� ������ ��������� ��� msi �������� ��� ����� �� ������ combobox
begin
ListGr.ItemIndex:=ComboNumMisProc.ItemIndex;
end;

procedure TEditTask.ListGrKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if (sender is Tcombobox) then
(sender as TCombobox).DroppedDown:=true;
end;

procedure TEditTask.ListGrKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key=13 then
begin
ButOk.Click;
end;
end;

function TEditTask.creatformNewtask(S:string):boolean; // ������� ����� ��� ������ ��������� msi ��� ���������
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:=s+'Form';

FormGr.Width:=350;
FormGr.Height:=100;
FormGr.BorderStyle:=bsDialog;
if s='msi' then
begin
FormGr.OnShow:=FormMsiOpen;
FormGr.Caption:='��������� ���������';
end;
if s='proc' then
begin
 FormGr.OnShow:=FormProcOpen;
 FormGr.Caption:='��������� ��������';
end;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:=s+'ComboBox';
ListGr.Text:='';
ListGr.top:=5;
ListGr.Left:=5;
ListGr.Width:=330;
ListGr.DropDownCount:=15;
Listgr.OnKeyDown:=ListGrKeyDown;
ListGr.OnSelect:= SelectComboMsiProc;
ListGr.OnKeyUp:=ListGrKeyUp;
ListGr.TabOrder:=0;

ComboNumMisProc:=TComboBox.Create(FormGr);
ComboNumMisProc.Parent:=FormGr;
ComboNumMisProc.Name:=s+'NumCombo';
ComboNumMisProc.Text:='';
ComboNumMisProc.top:=30;
ComboNumMisProc.Left:=5;
ComboNumMisProc.Width:=40;
ComboNumMisProc.OnSelect:=SelectComboProcMSI;
ComboNumMisProc.Style:=csOwnerDrawFixed;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:=s+'AddTask';
ButOk.Caption:='��������';
ButOk.Top:=40;
ButOk.Left:=160;
ButOk.Width:=85;
ButOk.OnClick:= AddTaskInListview;
ButOk.TabOrder:=1;


ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='�������';
ButClose.Top:=40;
ButClose.Left:=250;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=2;
FormGr.Showmodal;
end;

procedure TEditTask.ButtonAddUser(Sender: TObject);
begin
AddUserPass(EditTask.Caption,EditUser.Text,EditPass.text,ChekSave.Checked);
FormGr.close;
end;

function TEditTask.creatformUserPass(S:string):boolean; // ������� ����� ��� ����� ������������ � ������
begin
FormGr:=TForm.Create(EditTask);
FormGr.Name:=s+'FormUser';
FormGr.Width:=250;
FormGr.Height:=143;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:='������������ � ������';
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;


EditUser:=TEdit.Create(FormGr);
EditUser.Parent:=FormGr;
EditUser.Name:=s+'User';
EditUser.Text:='';
EditUser.top:=0;
EditUser.Left:=5;
EditUser.Width:=230;
EditUser.TextHint:='��� ������������';
EditUser.TabOrder:=0;

EditPass:=TEdit.Create(FormGr);
EditPass.Parent:=FormGr;
EditPass.Name:=s+'Passwd';
EditPass.Text:='';
EditPass.PasswordChar:=#7;
EditPass.TextHint:='������';
EditPass.top:=25;
EditPass.Left:=5;
EditPass.Width:=230;
EditPass.TabOrder:=1;

ChekSave:=TCheckBox.Create(FormGr);
ChekSave.parent:=FormGr;
ChekSave.Name:=s+'SavePass';
ChekSave.Checked:=false;
ChekSave.Caption:='��������� ����� ���������� ������';
ChekSave.Top:=52;
ChekSave.Left:=5;
ChekSave.Width:=230;
ChekSave.TabOrder:=2;

ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:=s+'AddTask';
ButOk.Caption:='��';
ButOk.Top:=75;
ButOk.Left:=50;
ButOk.Width:=85;
ButOk.OnClick:= ButtonAddUser;
ButOk.TabOrder:=2;


ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='�������';
ButClose.Top:=75;
ButClose.Left:=145;
ButClose.Width:=85;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=4;
FormGr.Showmodal;
end;

////////////////////////////////////////////////////////////////////////
procedure TEditTask.ButLoadService(Sender: TObject);
var
NamePSSearch:string;
begin
NamePSSearch:='localhost';
if not InputQuery('�������� ������ ����� ������ ��������� �����!!!', '��� ����������', NamePSSearch) then exit;
   if frmdomaininfo.ping(NamePSSearch) then
    LoadListService(NamePSSearch);
end;

function TEditTask.LoadListService(NamePC:string):boolean; /// ��������� ������ �����
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  FWbemObject     : OLEVariant;
  InstallLocation,InstallSource,Myversion,MyVendor:string;
  oEnum           : IEnumvariant;
  iValue          : LongWord;

begin
OleInitialize(nil);
ListGr.Clear;
ComboNumMisProc.clear;
 frmDomainInfo.memo1.Lines.Add('�������� ������ �����...');
 try
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text, frmDomainInfo.LabeledEdit2.Text,'','',128);
 FWbemObjectSet:= FWMIService.ExecQuery('SELECT name,caption FROM Win32_Service','WQL',wbemFlagForwardOnly);
 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
 while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// ��������
 begin
 if FWbemObject.Caption<>null then
 begin
 ListGr.Items.add(string(FWbemObject.Caption));       // �������� ������
 ComboNumMisProc.Items.add(string(FWbemObject.name));       // ��� ������
 end;
 FWbemObject:=Unassigned;
 end;
 VariantClear(FWbemObject);
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 except
 on E:Exception do
 begin
 frmDomainInfo.memo1.Lines.Add('������ �������� ����� "'+E.Message+'"');
 oEnum:=nil;
 VariantClear(FWbemObjectSet);
 VariantClear(FSWbemLocator);
 VariantClear(FWMIService);
 end;
 end;
OleUnInitialize;
ListGr.text:=('�������� ������ ����� ���������.');
ComboNumMisProc.text:=('�������� ������ �����  ���������.');
end;
//////////////////////////////////////////////////////////////////////
function TEditTask.creatformService(S:string):boolean; // ������� ����� ��� ������ �����  / �������� ��� ������� -������ �����/ ��������� ������ /�������� ������
var
step:integer;
begin
try
step:=0;
FormGr:=TForm.Create(EditTask);
FormGr.Name:='Form';
FormGr.Width:=346;
FormGr.Height:=130;
FormGr.BorderStyle:=bsDialog;
FormGr.Caption:=s;
FormGr.OnClose:=FormGrClose;
FormGr.Position:=poMainFormCenter;

step:=1;
ListGr:=TComboBox.Create(FormGr);
ListGr.Parent:=FormGr;
ListGr.Name:='CaptionService';
ListGr.Text:='';
ListGr.top:=5;
ListGr.Left:=3;
ListGr.Width:=328;
ListGr.DropDownCount:=15;
Listgr.OnKeyDown:=ListGrKeyDown;
ListGr.OnSelect:= SelectComboMsiProc;
//ListGr.OnKeyUp:=ListGrKeyUp;
ListGr.TabOrder:=0;
ListGr.TextHint:='�������� ������';
step:=2;
ComboNumMisProc:=TComboBox.Create(FormGr);
ComboNumMisProc.Parent:=FormGr;
ComboNumMisProc.Name:='NameService';
ComboNumMisProc.Text:='';
ComboNumMisProc.DropDownCount:=15;
ComboNumMisProc.OnKeyDown:=ListGrKeyDown;
ComboNumMisProc.top:=33;
ComboNumMisProc.Left:=3;
ComboNumMisProc.Width:=238;
ComboNumMisProc.OnSelect:=SelectComboProcMSI;
ComboNumMisProc.TextHint:='��� ������';
ListGr.TabOrder:=1;

step:=3;
ButLoad:=TButton.Create(FormGr);
ButLoad.Parent:=FormGr;
ButLoad.Name:='AddService';
ButLoad.Caption:='���������';
ButLoad.Hint:='��������� ������ ����� � ���������� � ����';
ButLoad.ShowHint:=true;
ButLoad.Top:=30;
ButLoad.Left:=248;
ButLoad.Width:=82;
ButLoad.OnClick:= ButLoadService;
ButLoad.TabOrder:=2;

step:=4;
ButOk:=TButton.Create(FormGr);
ButOk.Parent:=FormGr;
ButOk.Name:='AddTask';
ButOk.Caption:='��������';
ButOk.Top:=65;
ButOk.Left:=248;
ButOk.Width:=82;
ButOk.OnClick:= AddServiceInListview;
ButOk.TabOrder:=3;

step:=5;
ButClose:=TButton.Create(FormGr);
ButClose.Parent:=FormGr;
ButClose.Name:='ClFormGr';
ButClose.Caption:='�������';
ButClose.Top:=65;
ButClose.Left:=148;
ButClose.Width:=82;
ButClose.OnClick:=ButCloseFormGRClick;
ButClose.TabOrder:=4;
FormGr.Showmodal;
except on E: Exception do ShowMessage('������ �������� ������ �����');
end;
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TEditTask.N24Click(Sender: TObject);   // ������� ���������� ����� , ����� �� popupmenu
begin
Button3.Click;
end;


procedure TEditTask.N27Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
NewsSubKeyName('�������');
end;

procedure TEditTask.N28Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
NewsSubKeyName('�������');
end;

procedure TEditTask.N29Click(Sender: TObject);
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
DeleteKeyName;
end;

procedure TEditTask.ButtonNoClose(Sender: TObject);
begin
((sender as Tbutton).Parent as Tform).Close;
end;
procedure TEditTask.FormKeyClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).Release;
end;

procedure TEditTask.ButtonOksSubKeyName(Sender: TObject); // �������� ������� (������� ��� ������� ������)
var
i,z,x:integer;
des:string;
begin
if EditsSubKeyName.Text='' then begin showmessage('�� ������ ������'); exit; end;
if EditsSubKeyName.Text[Length(EditsSubKeyName.Text)]='\' then // ���� � ����� ������ ����� ������ '/' �� ������� ���
begin
 EditsSubKeyName.Text:=copy(EditsSubKeyName.Text,1,Length(EditsSubKeyName.Text)-1);
end;

if pos('�������',FormKey.Caption)<>0 then begin x:=1; Des:='�������'; end;
if pos('�������',FormKey.Caption)<>0 then begin x:=0; Des:='�������'; end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]=Des+' ������ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if length(Des+' ������ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text)>250 then
 begin
   ShowMessage('������ ������� ��������� ���������� ��������');
   exit;
 end;
with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:=inttostr(x); // 1 - ������� ������ 0 - ������� ������
    SubItems.Add(Des+' ������ :'+ComboRootKey.Text+':'+EditsSubKeyName.Text);    // ��������
    SubItems.Add('sSubKey'); //������ � �������� ������� (�������� ��� ��������)
  end;
FormKey.close;
end;

procedure TEditTask.ButtonOkDelKey(Sender: TObject); // �������� ������� �������� �����
var
i,z,x:integer;
begin
if EditsSubKeyName.Text='' then begin showmessage('�� ������ ������'); exit; end;
if EditNameKey.Text='' then begin showmessage('�� ������� ��� �����'); exit; end;
for I := 0 to EditTask.ListView1.Items.Count-1 do
  begin
    if EditTask.ListView1.Items[i].SubItems[0]='������� ���� :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text then
    begin
      z:=MessageDlg('��� ������� ��� ��������� � ������!!!'+#10#13+' �� ������� � ����� ���������???'
      , mtWarning,[mbYes,mbCancel],0);
      if z=mrcancel then exit
      else break;
    end;
  end;
if length('������� ���� :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text)>250 then
 begin
   ShowMessage('������ ������� ��������� ���������� ��������');
   exit;
 end;

with EditTask do
with ListView1.Items.Add do
  begin
    ImageIndex:=0;
    caption:='0'; // �������� �� ������ �� ���������� ��������
    SubItems.Add('������� ���� :'+ComboRootKey.Text+':'+EditsSubKeyName.Text+':'+EditNameKey.Text);    // ��������
    SubItems.Add('sNameKey'); //������ � �������� ������� (�������� ����� �������)
  end;
FormKey.close;
end;


function TEditTask.NewsSubKeyName(CreateDel:string):boolean; // ������� ��� ������� ������ �������
begin
try
FormKey:=TForm.Create(EditTask);
if CreateDel='�������' then FormKey.Caption:='������� ������ �������';
if CreateDel='�������' then FormKey.Caption:='������� ������ �������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=140;
FormKey.OnClose:=FormKeyClose;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=5;
ComboRootKey.Left:=5;
ComboRootKey.Width:=260;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������ (������ - HARDWARE\MRPC):';
EditsSubKeyName.Top:=45;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=72;
ButtonOk.Left:=110;
ButtonOk.OnClick:=ButtonOksSubKeyName;
ButtonOk.TabOrder:=1;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=72;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=2;
FormKey.ShowModal;
except on E: Exception do
begin
  Log_write('TASK',timetostr(now)+' ������ �������� ������� (��������/�������� ������� �������) - '+e.Message);
  frmdomaininfo.Memo1.Lines.Add('������ �������� ������� (��������/�������� ������� �������) - '+e.Message);
end;
end;
end;

procedure TEditTask.SZ1Click(Sender: TObject); /// ��������� ����� ����������� ������ �������
begin
if ListView1.Items.Count>=30 then
begin
 ShowMessage('���������� �������� ������ 30�� �������');
 exit;
end;
RegKeySave.Button4.Enabled:=true; // �������� ������ ���������� � ������
RegKeySave.Button1.Enabled:=false;  // ��������� ������ ������� � ������
RegKeySave.PopupEditKey.Items[1].Enabled:=false; // ��������� ������ ���� ��� ������� � ������
RegKeySave.ShowModal;  // ��������� �����
end;


function TEditTask.DeleteKeyName:boolean;
begin
try
FormKey:=TForm.Create(EditTask);
FormKey.Caption:='������� ���� �������';
FormKey.position:=poOwnerFormCenter;
FormKey.BorderStyle:=bsDialog;
FormKey.Width:=285;
FormKey.Height:=180;
FormKey.OnClose:=FormKeyClose;
/// /////////////////////
ComboRootKey:=TComboBox.Create(FormKey);
ComboRootKey.Parent:=FormKey;
ComboRootKey.Items.Add('HKEY_CLASSES_ROOT');
ComboRootKey.Items.Add('HKEY_LOCAL_MACHINE');
ComboRootKey.Items.Add('HKEY_CURRENT_USER');
ComboRootKey.Items.Add('HKEY_USERS');
ComboRootKey.Items.Add('HKEY_CURRENT_CONFIG');
ComboRootKey.ItemIndex:=1;
ComboRootKey.Style:=csOwnerDrawFixed;
ComboRootKey.Top:=5;
ComboRootKey.Left:=5;
ComboRootKey.Width:=260;
/////////////////////////
EditsSubKeyName:=TLabeledEdit.Create(FormKey);
EditsSubKeyName.Parent:=FormKey;
EditsSubKeyName.EditLabel.Caption:='������ (������ - HARDWARE\MRPC):';
EditsSubKeyName.Top:=45;
EditsSubKeyName.Left:=5;
EditsSubKeyName.Width:=260;
EditsSubKeyName.Text:='';
EditsSubKeyName.TabOrder:=0;
////////////////////////
EditNameKey:=TLabeledEdit.Create(FormKey);
EditNameKey.Parent:=FormKey;
EditNameKey.EditLabel.Caption:='��������� ��������:';
EditNameKey.Top:=83;
EditNameKey.Left:=5;
EditNameKey.Width:=260;
EditNameKey.Text:='��� ����� �������';
EditNameKey.TabOrder:=1;
////////////////////////
ButtonOk:=Tbutton.Create(FormKey);
ButtonOk.Parent:=FormKey;
ButtonOk.Caption:='��';
ButtonOk.Top:=110;
ButtonOk.Left:=110;
ButtonOk.OnClick:=ButtonOkDelKey;
ButtonOk.TabOrder:=2;
//////////////////
ButtonNo:=Tbutton.Create(FormKey);
ButtonNo.Parent:=FormKey;
ButtonNo.Caption:='������';
ButtonNo.Top:=110;
ButtonNo.Left:=190;
ButtonNo.OnClick:=ButtonNoClose;
ButtonNo.TabOrder:=3;
FormKey.ShowModal;
except on E: Exception do
begin
  Log_write('TASK',timetostr(now)+' ������ �������� ������� (��������/�������� ������� �������) - '+e.Message);
  frmdomaininfo.Memo1.Lines.Add('������ �������� ������� (��������/�������� ������� �������) - '+e.Message);
end;
end;
end;

procedure TEditTask.Button3Click(Sender: TObject); /// ������� ���������� ����������
var
i,step:integer;
NameTabl:string;
TransactionDelPC:TFDTransaction;
FDQueryDelPC:TFDQuery;
begin
try
step:=0;
if ListView2.SelCount=0 then exit;
TransactionDelPC:= TFDTransaction.Create(nil);
TransactionDelPC.Connection:=DataM.ConnectionDB;
TransactionDelPC.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionDelPC.Options.AutoCommit:=false;
TransactionDelPC.Options.AutoStart:=false;
TransactionDelPC.Options.AutoStop:=false;
FDQueryDelPC:= TFDQuery.Create(nil);
FDQueryDelPC.Transaction:=TransactionDelPC;
FDQueryDelPC.Connection:=DataM.ConnectionDB;
TransactionDelPC.StartTransaction; // �������� ����������
NameTabl:=nametableForDescription(EditTask.Caption);
for I := ListView2.Items.Count-1 downto 0 do
begin
step:=1;
  if (ListView2.Items[i].Selected) and (ListView2.Items[i].ImageIndex=4) then
  begin
  FDQueryDelPC.SQL.Clear;
  FDQueryDelPC.SQL.Text:='delete FROM '+NameTabl+' WHERE PC_NAME='''+ListView2.Items[i].SubItems[0]+'''';
  FDQueryDelPC.ExecSQL;
  end;
  step:=2;
  if (ListView2.Items[i].Selected) then ListView2.Items[i].Delete;
end;
step:=3;
TransactionDelPC.Commit;  // ��������� �� ��� ������ � ����
FDQueryDelPC.Close;
FDQueryDelPC.Free;
TransactionDelPC.Free
except on E: Exception do
begin
   if Assigned(FDQueryDelPC) then FDQueryDelPC.Free;
      if Assigned(TransactionDelPC) then
      begin
      TransactionDelPC.Rollback;
      TransactionDelPC.Free;
      end;
   Log_write('TASK',timetostr(now)+' ��� - '+inttostr(step)+' - ������ �������� ����������- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' ��� - '+inttostr(step)+' - ������ �������� ����������- '+e.Message);
end;
end;
end;


function TEditTask.nametableForDescription(Descript:string):string; /// ��� ������� �� ��������
begin
FDQueryProcMSI.SQL.clear;
FDQueryProcMSI.SQL.Text:='SELECT * FROM TABLE_TASK WHERE DESCRIPTION_TASK='''+Descript+'''';
FDQueryProcMSI.Open;
result:= FDQueryProcMSI.FieldByName('NAME_TABLE').AsString; // ��� �������� ��� ������� �� ��������
FDQueryProcMSI.SQL.clear;
end;

procedure TEditTask.N23Click(Sender: TObject); // ������� ������� �� popup menu
begin
Button4.Click;
end;

procedure TEditTask.Button4Click(Sender: TObject); /// ������� �������
var
i,z:integer;
step:integer;
nameT:string;
TransactionTaskDel:TFDTransaction;
FDQueryDelTask:TFDQuery;
begin
try
if StatusStartStopTask(EditTask.Caption) then // ���� ������ �������� (true)
 begin
   i:=MessageDlg('�� ������������� ������� ��������� � ���������� ������. ����������?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit;
 end;
step:=0;
if ListView1.SelCount=0 then exit; /// ���� �� �������� ������ �� �������
 step:=1;
nameT:=nametableForDescription(EditTask.Caption);  // ��� �������
TransactionTaskDel:= TFDTransaction.Create(nil);
TransactionTaskDel.Connection:=DataM.ConnectionDB;
TransactionTaskDel.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionTaskDel.Options.AutoCommit:=false;
TransactionTaskDel.Options.AutoStart:=false;
TransactionTaskDel.Options.AutoStop:=false;
FDQueryDelTask:= TFDQuery.Create(nil);
FDQueryDelTask.Transaction:=TransactionTaskDel;
FDQueryDelTask.Connection:=DataM.ConnectionDB;
TransactionTaskDel.StartTransaction; // �������� ����������
for I := ListView1.Items.Count-1 downto 0 do
begin
 step:=2;
  if (ListView1.Items[i].Selected) and (ListView1.Items[i].ImageIndex=1) then // ���� ������� ������� ��� ����������� � �������
  begin
   step:=3;
   FDQueryDelTask.SQL.clear;
   FDQueryDelTask.SQL.Text:='ALTER TABLE '+nameT+' DROP  TypeTask'+inttostr(i)
   +',DROP NumTask'+inttostr(i)+',DROP ResultTask'+inttostr(i)+',DROP NameTask'+inttostr(i)+',DROP StatusTask'+inttostr(i);
   FDQueryDelTask.ExecSQL;
   ListView1.Items[i].Delete;
   step:=4;
   for z := i+1 to ListView1.Items.Count do // ��������������� ������� �� ���������� �� ����������
   begin
   step:=5;
     FDQueryDelTask.SQL.clear;                            /// ������� ������� ��������������� � ����������
     FDQueryDelTask.SQL.Text:='ALTER TABLE '+nameT
     +' ALTER TypeTask'+inttostr(z)+' TO TypeTask'+inttostr(z-1)
     +',ALTER NumTask'+inttostr(z)+' TO NumTask'+inttostr(z-1)
     +',ALTER ResultTask'+inttostr(z)+' TO ResultTask'+inttostr(z-1)
     +',ALTER NameTask'+inttostr(z)+' TO NameTask'+inttostr(z-1)
     +',ALTER StatusTask'+inttostr(z)+' TO StatusTask'+inttostr(z-1);
     FDQueryDelTask.ExecSQL;
   step:=6;
   end;
  FDQueryDelTask.SQL.clear;                            /// �������� ���������� �����
  FDQueryDelTask.SQL.Text:='UPDATE '+nameT+' SET TASK_QUANT='+inttostr(ListView1.Items.Count);
  FDQueryDelTask.ExecSQL;
  end
 else
  if (ListView1.Items[i].Selected) and (ListView1.Items[i].ImageIndex=0) then  // ���� ��������� � ��� �� ��������� � ������� �� ������ ������� ������
   ListView1.Items[i].Delete;


end;
step:=7;
TransactionTaskDel.Commit;  // ��������� �� ��� ������ � ����
FDQueryDelTask.Close;
FDQueryDelTask.Free;
TransactionTaskDel.Free;
except on E: Exception do
  begin
  if Assigned(FDQueryDelTask) then FDQueryDelTask.Free;
      if Assigned(TransactionTaskDel) then
      begin
      TransactionTaskDel.Rollback;
      TransactionTaskDel.Free;
      end;
  Log_write('TASK',timetostr(now)+' ��� - '+inttostr(step)+' - ������ �������� ������� �� ������- '+e.Message);
  frmdomaininfo.Memo1.Lines.Add(' ��� - '+inttostr(step)+' - ������ �������� ������� �� ������- '+e.Message);
  end;
end;
end;


function TEditTask.RenameTableForTask(NameTable,NewDescriptionTable:string):boolean; //������ �������� ������� ��� ������. �������� ������ � ������� TABLE_TASK
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // ��������
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='UPDATE TABLE_TASK SET DESCRIPTION_TASK=:a  WHERE NAME_TABLE='''+NameTable+'''';
step:=2;
FDQueryWriteTaskTable.ParamByName('a').asstring:=NewDescriptionTable;
step:=3;
FDQueryWriteTaskTable.ExecSQL;
step:=4;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.Close;
TransactionWriteTaskTable.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryWriteTaskTable.SQL.clear;   //��������
FDQueryWriteTaskTable.Close;  /// ������� ���
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
result:=true;
Except
on E:Exception do
     begin
     if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:=false;
     Log_write('TASK',timetostr(now)+' ��� - '+inttostr(step)+' - ������ �������������� ������- '+e.Message);
     end;
end;
end;




function TEditTask.UpdateTableTABLETASK(NameTable:string;countPC:integer):boolean; // �������� ������ � ������� TABLE_TASK
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // ��������

FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='update or insert into TABLE_TASK (NAME_TABLE,COUNT_PC) VALUES (:p1,:p2)  MATCHING (NAME_TABLE)';
step:=2;
FDQueryWriteTaskTable.params.ParamByName('p1').AsString:=''+NameTable+'';
FDQueryWriteTaskTable.params.ParamByName('p2').Asinteger:=countPC; // ����� �����������
step:=3;
FDQueryWriteTaskTable.ExecSQL;
step:=4;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.Close;
result:=true;
TransactionWriteTaskTable.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryWriteTaskTable.SQL.clear;   //��������
FDQueryWriteTaskTable.Close;  /// ������� ���
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
Except
on E:Exception do
     begin
     if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:=false;
     Log_write('TASK',timetostr(now)+' ��� - '+inttostr(step)+' - ������ Update TABLE_TASK- '+e.Message);
     end;
end;
end;

function TEditTask.ThereAnyRunStopTask(statustask:boolean):integer; /// ����� ������ ������, b ������� ���� �� ���������� ��� �������������
var
FDQueryReadStatusTask:TFDQuery;
TransactionReadStatusTask: TFDTransaction;
begin
try
TransactionReadStatusTask:= TFDTransaction.Create(nil);
TransactionReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionReadStatusTask.Options.AutoCommit:=false;
TransactionReadStatusTask.Options.AutoStart:=false;
TransactionReadStatusTask.Options.AutoStop:=false;
TransactionReadStatusTask.Options.ReadOnly:=true;
FDQueryReadStatusTask:= TFDQuery.Create(nil);
FDQueryReadStatusTask.Transaction:=TransactionReadStatusTask;
FDQueryReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.StartTransaction; // ��������
FDQueryReadStatusTask.SQL.Clear;
FDQueryReadStatusTask.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE START_STOP='''+booltostr(statustask)+'''';
FDQueryReadStatusTask.open;
FDQueryReadStatusTask.Last;
result:=FDQueryReadStatusTask.RecordCount; //��������� ���������� ���������� ������� �� ��������  statustask
TransactionReadStatusTask.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryReadStatusTask.SQL.clear;   //��������
FDQueryReadStatusTask.Close;  /// ������� ���
FDQueryReadStatusTask.Free;
TransactionReadStatusTask.Free;
 except on E: Exception do
     begin
     result:=0;
     if Assigned(FDQueryReadStatusTask) then FDQueryReadStatusTask.Free;
      if Assigned(TransactionReadStatusTask) then
      begin
      TransactionReadStatusTask.Rollback;
      TransactionReadStatusTask.Free;
      end;
     Log_write('TASK',timetostr(now)+' - ������ READ StartStopTask TABLE_TASK :' +e.Message);
     end;
   end;
end;

function TEditTask.StatusStartStopTask(DescriptionTask:string):boolean; /// ����� ������ ������, ���������� ��� ���������� ����������
var                                                           //// ���� ������ ����������� �� ��������� ����������� false
FDQueryReadStatusTask:TFDQuery;
TransactionReadStatusTask: TFDTransaction;
begin
TransactionReadStatusTask:= TFDTransaction.Create(nil);
TransactionReadStatusTask.Connection:=DataM.ConnectionDB;
TransactionReadStatusTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionReadStatusTask.Options.AutoCommit:=false;
TransactionReadStatusTask.Options.AutoStart:=false;
TransactionReadStatusTask.Options.AutoStop:=false;
TransactionReadStatusTask.Options.ReadOnly:=true;
FDQueryReadStatusTask:= TFDQuery.Create(nil);
FDQueryReadStatusTask.Transaction:=TransactionReadStatusTask;
FDQueryReadStatusTask.Connection:=DataM.ConnectionDB;
try
try
TransactionReadStatusTask.StartTransaction; // ��������
FDQueryReadStatusTask.SQL.Clear;
FDQueryReadStatusTask.SQL.Text:='SELECT START_STOP FROM TABLE_TASK WHERE DESCRIPTION_TASK='''+DescriptionTask+'''';
FDQueryReadStatusTask.open;
result:=strtobool(FDQueryReadStatusTask.FieldByName('START_STOP').AsString);
//frmDomainInfo.Memo1.Lines.Add('�������� �� ������� - '+FDQueryReadStatusTask.FieldByName('START_STOP').AsString);
TransactionReadStatusTask.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryReadStatusTask.SQL.clear;   //��������
FDQueryReadStatusTask.Close;  /// ������� ���
 except on E: Exception do
     begin
      TransactionReadStatusTask.Rollback;
      FDQueryReadStatusTask.Close;
      Log_write('TASK',timetostr(now)+' - ������ READ StartStopTask TABLE_TASK :' +e.Message);
      end;
     end;
finally
FDQueryReadStatusTask.Free;
TransactionReadStatusTask.Free;
end;
end;

function TEditTask.AddNewItemsTable(DescriptionTable:string;countPC:integer):string; // ���������� �������� ������� � ������� ���� ������� � ���������� ��� ����� �������
var
step,counttable:integer;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // ��������
FDQueryWriteTaskTable.SQL.Text:='SELECT GEN_ID(GEN_FOR_TABLE_TASK,0) FROM RDB$DATABASE'; // ������� �������� ����������, ��� ���������� �� 1
FDQueryWriteTaskTable.Open;
counttable:=FDQueryWriteTaskTable.Fields[0].AsInteger+1; ///��� � ��������, ���� ������� �������� 0 �� ��� ���������� ������ ���� 1, ������������� ��������� ������� ��� ������������ ����� ������� � ID ����������
FDQueryWriteTaskTable.Close;
TransactionWriteTaskTable.Commit;
TransactionWriteTaskTable.StartTransaction;
FDQueryWriteTaskTable.SQL.Clear;
FDQueryWriteTaskTable.SQL.Text:='update or insert into TABLE_TASK (NAME_TABLE,DESCRIPTION_TASK,STATUS_TASK,START_STOP,COUNT_PC,CURRENT_TASK,PC_RUN,WORKS_THREAD,USER_NAME,PASS_USER,SAVE_PASS) VALUES (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11)  MATCHING (NAME_TABLE)';
FDQueryWriteTaskTable.params.ParamByName('p1').AsString:=''+'TASK_'+inttostr(counttable)+'';
FDQueryWriteTaskTable.params.ParamByName('p2').AsString:=''+DescriptionTable+'';
FDQueryWriteTaskTable.params.ParamByName('p3').AsString:='�����������'; // ������ ������� ������
FDQueryWriteTaskTable.params.ParamByName('p4').AsString:=booltostr(false); /// ������� ���� ��� ������ ������ ���� ����������� �����������/��� �������� ������
FDQueryWriteTaskTable.params.ParamByName('p5').Asinteger:=countPC; // ����� �����������
FDQueryWriteTaskTable.params.ParamByName('p6').AsString:='';     // ������ �������
FDQueryWriteTaskTable.params.ParamByName('p7').AsString:='';     // �� ����� ����� �����������
FDQueryWriteTaskTable.params.ParamByName('p8').AsString:=booltostr(false); // ������� ���� ��� ����� ��������� ���������� ������ ����� ������ �������� �� START_STOP
FDQueryWriteTaskTable.params.ParamByName('p9').AsString:='';   // ������������
FDQueryWriteTaskTable.params.ParamByName('p10').AsString:='';  // ������
FDQueryWriteTaskTable.params.ParamByName('p11').AsBoolean:=false; // �� ��������� ������ ����� ���������� ������
FDQueryWriteTaskTable.ExecSQL;
TransactionWriteTaskTable.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit
FDQueryWriteTaskTable.SQL.clear;   //��������
FDQueryWriteTaskTable.Close;  /// ������� ���
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;
result:='TASK_'+inttostr(counttable); // � ���������� ��� �������
Except
on E:Exception do
     begin
      if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
      if Assigned(TransactionWriteTaskTable) then
      begin
      TransactionWriteTaskTable.Rollback;
      TransactionWriteTaskTable.Free;
      end;
     result:='';
     Log_write('TASK',timetostr(now)+' ��� - '+inttostr(step)+' - ������ ������ ����� �������- '+e.Message);
     end;
end;
end;

////////////////////////////////////////////////////////////////////
function TeditTask.AddOrUpdateListViewreadInTableTask(nameTable:string):boolean;
var
i:integer;
update:boolean;
FDQueryWriteTaskTable:TFDQuery;
TransactionWriteTaskTable: TFDTransaction;
begin  // ������ ������ � ������� Table_task � ���������� ���������� � ��� � ������ ListView
try
TransactionWriteTaskTable:= TFDTransaction.Create(nil);
TransactionWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionWriteTaskTable.Options.AutoCommit:=false;
TransactionWriteTaskTable.Options.AutoStart:=false;
TransactionWriteTaskTable.Options.AutoStop:=false;
TransactionWriteTaskTable.Options.ReadOnly:=true;
FDQueryWriteTaskTable:= TFDQuery.Create(nil);
FDQueryWriteTaskTable.Transaction:=TransactionWriteTaskTable;
FDQueryWriteTaskTable.Connection:=DataM.ConnectionDB;
TransactionWriteTaskTable.StartTransaction; // ��������

update:=false;
FDQueryWriteTaskTable.SQL.clear;
FDQueryWriteTaskTable.SQL.Text:='SELECT * FROM TABLE_TASK WHERE NAME_TABLE='''+nameTable+'''';
FDQueryWriteTaskTable.Open;
for I := 0 to frmdomaininfo.TaskListView.Items.Count-1 do // ���� ���� �� ����� ������ � ������ ����� �� �������� ��� ��������
begin
 if frmdomaininfo.TaskListView.Items[i].Caption=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString then
  begin
  frmdomaininfo.TaskListView.Items[i].Caption:=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString; // �������� ������
  frmdomaininfo.TaskListView.Items[i].SubItems[0]:=FDQueryWriteTaskTable.FieldByName('STATUS_TASK').AsString; // ������ ������ (�����������/�����������/���������)
  frmdomaininfo.TaskListView.Items[i].SubItems[1]:=FDQueryWriteTaskTable.FieldByName('CURRENT_TASK').AsString;     // ������� ���������� ������
  frmdomaininfo.TaskListView.Items[i].SubItems[2]:=FDQueryWriteTaskTable.FieldByName('PC_RUN').AsString;      // �� ����� ���������� ������ �����������
  frmdomaininfo.TaskListView.Items[i].SubItems[3]:=FDQueryWriteTaskTable.FieldByName('COUNT_PC').AsString;    // ����� ����������� � ������
  update:=true;
  end;
end;
if not update then /// ���� ����� ������ ��� ������ �� ���� �������� � ������
with frmdomaininfo.TaskListView.Items.Add do
begin
caption:=FDQueryWriteTaskTable.FieldByName('DESCRIPTION_TASK').AsString; // �������� ������
SubItems.add(FDQueryWriteTaskTable.FieldByName('STATUS_TASK').AsString); // ������ ������ (�����������/�����������/���������)
SubItems.add(FDQueryWriteTaskTable.FieldByName('CURRENT_TASK').AsString);     // ������� ���������� ������
SubItems.add(FDQueryWriteTaskTable.FieldByName('PC_RUN').AsString);      // �� ����� ���������� ������ �����������
SubItems.add(FDQueryWriteTaskTable.FieldByName('COUNT_PC').AsString);    // ����� ����������� � ������
end;
TransactionWriteTaskTable.Commit;  /// ��������� ��������� , ���� ��� ������ ���� ��������� commit

FDQueryWriteTaskTable.SQL.clear;   //��������
FDQueryWriteTaskTable.Close;  /// ������� ���
FDQueryWriteTaskTable.Free;
TransactionWriteTaskTable.Free;

result:=true;
except on E: Exception do
begin
 if Assigned(FDQueryWriteTaskTable) then FDQueryWriteTaskTable.Free;
 if Assigned(TransactionWriteTaskTable) then
 begin
 TransactionWriteTaskTable.Rollback;
 TransactionWriteTaskTable.Free;
 end;
frmdomaininfo.memo1.Lines.add('������ ���������� ListView :'+E.Message);
result:=false;
end;
end;
end;
////////////////////////////////////////////////////////////////////


function TEditTask.createNewTableTask(TableName:string):boolean;     //������� ������� �������� ����� �������+�� ����������
var
i,z,step,p:integer;
ColName,ParamNum:string;
FDQueryAdd:TFDQuery;
TransactionAdd: TFDTransaction;
function incP(z:integer):integer;
begin
p:=z+1;
result:=p;
end;
begin
try
TransactionAdd:= TFDTransaction.Create(nil);
TransactionAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; //xiUnspecified
TransactionAdd.Options.AutoCommit:=false;
TransactionAdd.Options.AutoStart:=false;
TransactionAdd.Options.AutoStop:=false;
FDQueryAdd:= TFDQuery.Create(nil);
FDQueryAdd.Transaction:=TransactionAdd;
FDQueryAdd.Connection:=DataM.ConnectionDB;
TransactionAdd.StartTransaction;
step:=0;
dataM.�reateTablForTASK(TableName,'','','','',booltostr(false)); //������� ������� //������� �������
step:=1;
for I := 0 to ListView1.Items.Count-1 do /// ������� ������� ��� �����
begin
dataM.�reateTablForTASK(TableName,'TypeTask'+inttostr(i),'NumTask'+inttostr(i),'ResultTask'+inttostr(i),'NameTask'+inttostr(i),'StatusTask'+inttostr(i)) ;
end;
step:=2;
ColName:='';
ParamNum:='';
p:=3;
for i := 0 to ListView1.items.Count-1 do  // ��������� ��������� ���������� ������� ��� sql �������
Begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
End;
step:=3;
FDQueryAdd.SQL.clear;
FDQueryAdd.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME);';
 step:=4;
 FDQueryAdd.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - ����� ���������� �����
 step:=5;
p:=3;
for i := 0 to ListView1.items.Count-1 do //���� �� ������� ��������� ���� ��� ��� ��� ��� ��������� ��� ������ ������ ����������
begin
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - ��� ������ msi, proc, �������
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - ����� ����������� ������
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // ��������� ���������� ������  ResultTask
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // �������� ������� NameTask
  FDQueryAdd.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';      // ������� ��������� ������  StatusTask
end;
step:=6;
for i := 0 to ListView2.items.Count-1 do /// ���� �� ������ ��������� � ��������� sql ������ �.�. ������ ����� ���������� ������ ������ ����� � ID
begin
  FDQueryAdd.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryAdd.params.ParamByName('p3').AsString:=ListView2.items[i].subitems[0];           //PC_NAME
  step:=7;
  FDQueryAdd.ExecSQL;
end;
step:=8;
///////////////////////////////////////////////////////////////////
TransactionAdd.Commit;
FDQueryAdd.Close;
FDQueryAdd.Free;
TransactionAdd.Free;
result:=true;
Except
   on E: Exception do
      begin
      if Assigned(TransactionAdd) then
      begin
      TransactionAdd.Rollback;
      TransactionAdd.Free;
      end;
       if Assigned(FDQueryAdd) then
      begin
       FDQueryAdd.Free;
      end;
      frmdomaininfo.memo1.Lines.add('������ �������� ����� ������ (��� '+inttostr(step)+') :'+E.Message);
      result:=false;
      end;
   end;

end;

function TEditTask.UpdateTableTask(TableName:string):boolean;     //������� ���������� �������+�� ����������
var
i,z,step,p,NewZorPC:integer;
ColName,ParamNum:string;
FDQueryUpdate:TFDQuery;
TransactionUpdate: TFDTransaction;
function incP(z:integer):integer;
begin
p:=z+1;
result:=p;
end;
begin
try
step:=0;
step:=1;
for I := 0 to ListView1.Items.Count-1 do /// ������� ������� ��� �����
begin
if ListView1.Items[i].ImageIndex=0 then // ���� ����� ������ ���� ��������� �� ������� ����� �������
dataM.�reateTablForTASK(TableName,'TypeTask'+inttostr(i),'NumTask'+inttostr(i),'ResultTask'+inttostr(i),'NAMETASK'+inttostr(i),'StatusTASK'+inttostr(i)) ;
end;
step:=2;
ColName:=''; /// ����� ����� ������ � ������� �������
ParamNum:=''; // ����� ����� ������ � ����������� (:p1....:p.)
NewZorPC:=0;  // ������� ����� ������� ��� ����� ������
////////////////////////////////////////////////////////���������� ������� ������ �������� ���  (��� ��� ��� � ������) ������
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
for i := 0 to ListView1.items.Count-1 do /// ��������� � ����� ���� �� ��������� ����� ������
if ListView1.Items[i].ImageIndex=0 then NewZorPC:=NewZorPC+1;

if NewZorPC<>0 then // ���� NewZorPC ������ 0 �� ������ ���� ���������
Begin
TransactionUpdate:= TFDTransaction.Create(nil);
TransactionUpdate.Connection:=DataM.ConnectionDB;
TransactionUpdate.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;  xiUnspecified
TransactionUpdate.Options.AutoCommit:=false;
TransactionUpdate.Options.AutoStart:=false;
TransactionUpdate.Options.AutoStop:=false;
FDQueryUpdate:= TFDQuery.Create(nil);
FDQueryUpdate.Transaction:=TransactionUpdate;
FDQueryUpdate.Connection:=DataM.ConnectionDB;
p:=3;
for i := 0 to ListView1.items.Count-1 do  // ��������� ��������� ���������� ������� ��� sql �������
begin
 if ListView1.Items[i].ImageIndex=0 then // ���� ��� ����� ������
 begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
 end;
end;
step:=3;
FDQueryUpdate.SQL.clear;
FDQueryUpdate.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME)';;
 step:=4;
FDQueryUpdate.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - ����� ���������� �����
 step:=5;
p:=3;
for i := 0 to ListView1.items.Count-1 do //���� �� ������� ��������� ���� ��� ��� ��� ��� ��������� ��� ������ ������ ����������
begin
   if ListView1.Items[i].ImageIndex=0 then  // ���� ��� ����� ������
   begin
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - ��� ������ msi, proc, �������
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - ����� ����������� ������
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // ��������� ���������� ������
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // �������� �������
  FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO'; // ������� ��������� ������  //statusTask
  end;
end;
step:=6;
for i := 0 to ListView2.items.Count-1 do /// ���� �� ������ ��������� � ��������� sql ������ �.�. ������ ����� ���������� ������ ������ ����� � ID
begin
if ListView2.items[i].ImageIndex=4 then   // ��������� ��� ������ ������, ��������� �� ����� ������
  begin
  TransactionUpdate.StartTransaction;
  FDQueryUpdate.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryUpdate.params.ParamByName('p3').AsString:=''+ListView2.items[i].subitems[0]+'';           //PC_NAME
  step:=7;
  FDQueryUpdate.ExecSQL;
  TransactionUpdate.Commit;  /// ��������� ���������
  end;
end;
step:=8;

FDQueryUpdate.SQL.clear;   //��������
FDQueryUpdate.Close;  /// ������� ��� ����� ������
FDQueryUpdate.Free;
TransactionUpdate.Free;
End; /// ��������� ���������� ����� ����� ��� ������ ������
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////���������� ������� ����� �������� ��� ����� ������
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
NewZorPC:=0;
for i := 0 to ListView2.items.Count-1 do // �������� � ����� � ������� ������������ �� ����� �����
if ListView2.items[i].ImageIndex=0 then NewZorPC:=NewZorPC+1;
if NewZorPC<>0 then
Begin
TransactionUpdate:= TFDTransaction.Create(nil);
TransactionUpdate.Connection:=DataM.ConnectionDB;
TransactionUpdate.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionUpdate.Options.AutoCommit:=false;
TransactionUpdate.Options.AutoStart:=false;
TransactionUpdate.Options.AutoStop:=false;
FDQueryUpdate:= TFDQuery.Create(nil);
FDQueryUpdate.Transaction:=TransactionUpdate;
FDQueryUpdate.Connection:=DataM.ConnectionDB;
ColName:='';
ParamNum:='';
p:=3;
for i := 0 to ListView1.items.Count-1 do  // ��������� ��������� ���������� ������� ��� sql �������
Begin
   if i=ListView1.items.Count-1 then
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i);
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p));
   end
  else
   begin
   ColName:=ColName+'TypeTask'+inttostr(i)+',NumTask'+inttostr(i)+',ResultTask'+inttostr(i)+',NameTask'+inttostr(i)+',StatusTask'+inttostr(i)+',';
   ParamNum:=ParamNum+':p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',:p'+inttostr(incP(p))+',';
   end;
End;
step:=9;
FDQueryUpdate.SQL.clear;
FDQueryUpdate.SQL.Text:='update or insert into '
+TableName+' (ID_PC,TASK_QUANT,PC_NAME,'+ColName+')'    //(ID_PC,TASK_QUANT,PC_NAME)
+' VALUES(:p1,:p2,:p3,'+ParamNum+') MATCHING (PC_NAME)';;
 step:=10;
FDQueryUpdate.params.ParamByName('p2').AsInteger:=ListView1.Items.count; //TASK_QUANT - ����� ���������� �����
step:=11;
p:=3;
for i := 0 to ListView1.items.Count-1 do //���� �� ������� ��������� ���� ��� ��� ��� ��� ��������� ��� ������ ������ ����������
begin
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[1]; //      //  TypeTask - ��� ������ msi, proc, �������
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsInteger:=strtoint(ListView1.Items[i].Caption);      //NumTask - ����� ����������� ������
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO';       // ��������� ���������� ������
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:=ListView1.Items[i].SubItems[0]; // �������� �������
 FDQueryUpdate.params.ParamByName('p'+inttostr(incP(p))).AsString:='NO'; // ������� ��������� ������� StatusTask
end;
step:=12;
for i := 0 to ListView2.items.Count-1 do /// ���� �� ������ ��������� � ��������� sql ������ �.�. ������ ����� ���������� ������ ������ ����� � ID
begin
 if ListView2.items[i].ImageIndex=0 then  // ��������� ������ ��� ����� ������, ��������� ��� ������
  begin
  TransactionUpdate.StartTransaction; // ��������
  FDQueryUpdate.params.ParamByName('p1').AsInteger:=strtoint(ListView2.Items[i].Caption); //ID_PC
  FDQueryUpdate.params.ParamByName('p3').AsString:=''+ListView2.items[i].subitems[0]+'';           //PC_NAME
  step:=13;
  FDQueryUpdate.ExecSQL;
  TransactionUpdate.Commit;  /// ��������� ���������
  end;
end;
step:=14;

FDQueryUpdate.SQL.clear;   //��������
FDQueryUpdate.Close;  /// ������� ��� ����� ������
FDQueryUpdate.Free;
TransactionUpdate.Free;
End; /// ��������� ���������� ������� ��� ����� ������
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
result:=true;
Except
   on E: Exception do
      begin
      if Assigned(FDQueryUpdate) then FDQueryUpdate.Free;
      if Assigned(TransactionUpdate) then
      begin
      TransactionUpdate.Rollback;
      TransactionUpdate.Free;
      end;
      //
      frmdomaininfo.memo1.Lines.add('������ ���������� ������ (��� '+inttostr(step)+') :'+E.Message);
      result:=false;
      end;
   end;

end;


procedure TEditTask.SaveTask(Sender: TObject); /// ��������� ������
var
NameNewTask:string;
i:integer;
function imageLV (b:boolean):boolean;
var
i:integer;
begin
if b then
  begin
  for I := 0 to ListView1.Items.Count-1 do ListView1.Items[i].ImageIndex:=1;
  for I := 0 to ListView2.Items.Count-1 do ListView2.Items[i].ImageIndex:=4;
  end;
end;
begin
try
if StatusStartStopTask(EditTask.Caption) then // ���� ������ �������� (true)
 begin
    i:=MessageDlg('�� ������������� ������� ��������� � ���������� ������. ����������?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit
 end;
if (frmDomainInfo.TaskListView.Items.Count>=1) and (TaskName.Caption='') then
begin
ShowMessage('������ ����� ������ � ������������������ ������ ���������!!!');
exit;
end;
  if ListView2.Items.Count=0 then // ���� ������ ������ ���� �� �� ���������� ��������
  begin
  ShowMessage('�������� � ������ ����������');
  PageControl1.TabIndex:=1;
  Button2.SetFocus;
  exit;
  end;

  if ListView1.Items.Count=0 then // ���� ������ ������� ���� �� �� ���������� ��������
  begin
  ShowMessage('��� ���������� ������ �������� �������');
  exit;
  end;

if ListView1.Items.Count>30 then
begin
ShowMessage('��������� ���������� ����� ������� � ������, ���������� ������� '+#13#10+' �� ����� ���� ������ 30, ������� ������ �������!!!');
exit
end;


if TaskName.Caption='' then // ���� takLabel ����� ������ ��������� ����� ������
 begin
  NameNewTask:=AddNewItemsTable(EditTask.Caption,ListView2.Items.Count); // �������� ��� ������ �������, ������� ������ � ������� ��� ������ �����
  if NameNewTask='' then
  begin
   ShowMessage('�� ������� ��������� ������... ������ � �����');
   exit;
  end;

  if createNewTableTask(NameNewTask) then // ������� ������� ��� ������
  begin
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then // ��������� ������ � ������� � ������� lisView ��� �����
  imageLV(true); // ������ �������� � ������� ��� ����������� ���� ��� ��� ������� � ����� ����� � �������
  TaskName.Caption:=NameNewTask; /// ������ � ����� ��� �������, ����� ������� �� ����� ���������������
  end;
 end
else // ����� ����������� ������
  begin
  NameNewTask:=TaskName.Caption;       // ��� �������������� ��� ������� ���������� � label
  if UpdateTableTask(NameNewTask) then // ��������� ������� ��� ������
  begin
  if UpdateTableTableTask(NameNewTask,ListView2.Items.Count) then // ��������� ������ � ������� � TABLE_TASK
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then// ����� ��������� ���������� � ���������� ������� � ������� ListView ��� �����
     imageLV(true); // ������ �������� � ������� ��� ����������� ���� ��� ��� ������� � ����� ����� � �������
  end;
  end;
Button7.Enabled:=false;
Button8.Enabled:=true;

except on E: Exception do frmDomainInfo.Memo1.Lines.Add('������ ���������� ������ '+e.Message)
end;
end;

procedure TEditTask.Button7Click(Sender: TObject); /// ��������� ������
var
NameNewTask:string;
i:integer;
function imageLV (b:boolean):boolean;
var
i:integer;
begin
if b then
  begin
  for I := 0 to ListView1.Items.Count-1 do ListView1.Items[i].ImageIndex:=1;
  for I := 0 to ListView2.Items.Count-1 do ListView2.Items[i].ImageIndex:=4;
  end;
end;
begin
try
if StatusStartStopTask(EditTask.Caption) then // ���� ������ �������� (true)
 begin
    i:=MessageDlg('�� ������������� ������� ��������� � ���������� ������. ����������?', mtConfirmation,[mbYes,mbCancel],0);
   if i=IDCancel then exit
 end;
  if ListView2.Items.Count=0 then // ���� ������ ������ ���� �� �� ���������� ��������
  begin
  ShowMessage('�������� � ������ ����������');
  PageControl1.TabIndex:=1;
  Button2.SetFocus;
  exit;
  end;

  if ListView1.Items.Count=0 then // ���� ������ ������� ���� �� �� ���������� ��������
  begin
  ShowMessage('��� ���������� ������ �������� �������');
  exit;
  end;

if ListView1.Items.Count>30 then
begin
ShowMessage('��������� ���������� ����� ������� � ������, ���������� ������� '+#13#10+' �� ����� ���� ������ 30, ������� ������ �������!!!');
exit
end;


if TaskName.Caption='' then // ���� takLabel ����� ������ ��������� ����� ������
 begin
  NameNewTask:=AddNewItemsTable(EditTask.Caption,ListView2.Items.Count); // �������� ��� ������ �������, ������� ������ � ������� ��� ������ �����
  if NameNewTask='' then
  begin
   ShowMessage('�� ������� ��������� ������... ������ � �����');
   exit;
  end;

  if createNewTableTask(NameNewTask) then // ������� ������� ��� ������
  begin
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then // ��������� ������ � ������� � ������� lisView ��� �����
  imageLV(true); // ������ �������� � ������� ��� ����������� ���� ��� ��� ������� � ����� ����� � �������
  TaskName.Caption:=NameNewTask; /// ������ � ����� ��� �������, ����� ������� �� ����� ���������������
  end;
 end
else // ����� ����������� ������
  begin
  NameNewTask:=TaskName.Caption;       // ��� �������������� ��� ������� ���������� � label
  if UpdateTableTask(NameNewTask) then // ��������� ������� ��� ������
  begin
  if UpdateTableTableTask(NameNewTask,ListView2.Items.Count) then // ��������� ������ � ������� � TABLE_TASK
   if AddOrUpdateListViewreadInTableTask(NameNewTask) then// ����� ��������� ���������� � ���������� ������� � ������� ListView ��� �����
     imageLV(true); // ������ �������� � ������� ��� ����������� ���� ��� ��� ������� � ����� ����� � �������
  end;
  end;
Button7.Enabled:=false;
Button8.Enabled:=true;

except on E: Exception do frmDomainInfo.Memo1.Lines.Add('������ ���������� ������ '+e.Message)
end;
end;

function TEditTask.AddUserPass(DescriptionTable,User,Passwd:string;SavePass:boolean):boolean; // ���������� ������������ � ������ ��� ������
var
FDQueryStarTask:TFDQuery;
TransactionStarTask: TFDTransaction;
begin
try
TransactionStarTask:= TFDTransaction.Create(nil);
TransactionStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionStarTask.Options.AutoCommit:=false;
TransactionStarTask.Options.AutoStart:=false;
TransactionStarTask.Options.AutoStop:=false;
FDQueryStarTask:= TFDQuery.Create(nil);
FDQueryStarTask.Transaction:=TransactionStarTask;
FDQueryStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.StartTransaction; // ��������
NameRunTASK:=nametableForDescription(DescriptionTable); /// ��� ������� �� �������� ������ ��� �������� � �����
FDQueryStarTask.SQL.Clear;
FDQueryStarTask.SQL.Text:='update TABLE_TASK set USER_NAME=:a ,PASS_USER=:b, SAVE_PASS=:c WHERE NAME_TABLE='''+NameRunTASK+'''';
FDQueryStarTask.ParamByName('a').asstring:=user;    // ������������
FDQueryStarTask.ParamByName('b').asstring:=Passwd;  // ������
FDQueryStarTask.ParamByName('c').AsBoolean:=SavePass;// ��������� ������ ����� ������� ������ ��� ���
FDQueryStarTask.ExecSQL;
TransactionStarTask.Commit;  /// ��������� ���������
FDQueryStarTask.SQL.clear;   //��������
FDQueryStarTask.Close;  /// ������� ��� ����� ������
FDQueryStarTask.Free;
TransactionStarTask.Free;
result:=true;
except on E: Exception do
begin
if Assigned(FDQueryStarTask) then FDQueryStarTask.Free;
      if Assigned(TransactionStarTask) then
      begin
      TransactionStarTask.Rollback;
      TransactionStarTask.Free;
      end;
 frmDomainInfo.memo1.Lines.add('������ ���������� ������������ � ������ '+NameRunTASK+' : '+E.Message);
 result:=false;
end;
end;
end;

function TEditTask.StatrTask(DescriptionTable:string):boolean; // ������ ������
var
TaskThread:TThread;
FDQueryStarTask:TFDQuery;
TransactionStarTask: TFDTransaction;
begin
try
TransactionStarTask:= TFDTransaction.Create(nil);
TransactionStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot; xiUnspecified
TransactionStarTask.Options.AutoCommit:=false;
TransactionStarTask.Options.AutoStart:=false;
TransactionStarTask.Options.AutoStop:=false;
FDQueryStarTask:= TFDQuery.Create(nil);
FDQueryStarTask.Transaction:=TransactionStarTask;
FDQueryStarTask.Connection:=DataM.ConnectionDB;
TransactionStarTask.StartTransaction; // ��������
//NameRunTASK:=TaskName.Caption; // ��� ���������� ��� ��������� ������, ��� ������� ����������� �   label
NameRunTASK:=nametableForDescription(DescriptionTable); /// ��� ������� �� �������� ������ ��� �������� � �����
FDQueryStarTask.SQL.Clear;
FDQueryStarTask.SQL.Text:='update TABLE_TASK set STATUS_TASK=:a ,START_STOP=:b WHERE NAME_TABLE='''+NameRunTASK+'''';
FDQueryStarTask.ParamByName('a').asstring:='�����������';
FDQueryStarTask.ParamByName('b').asstring:=booltostr(true);  // �������� true ��� ���������� ������� ������
FDQueryStarTask.ExecSQL;
TransactionStarTask.Commit;  /// ��������� ���������
FDQueryStarTask.SQL.clear;   //��������
FDQueryStarTask.Close;  /// ������� ��� ����� ������
FDQueryStarTask.Free;
TransactionStarTask.Free;

TaskThread:=RunTask.TaskRun.Create(true);
TaskThread.FreeOnTerminate:=true;
TaskThread.Start;
result:=true;
except on E: Exception do
begin
if Assigned(FDQueryStarTask) then FDQueryStarTask.Free;
      if Assigned(TransactionStarTask) then
      begin
      TransactionStarTask.Rollback;
      TransactionStarTask.Free;
      end;
 frmDomainInfo.memo1.Lines.add('������ ������� ������ '+NameRunTASK+' : '+E.Message);
 result:=false;
end;
end;
end;

procedure TEditTask.Button8Click(Sender: TObject); /// ��������� ������
begin
if StatusStartStopTask(EditTask.Caption) then // ���� ������ �������� (true)
 begin
   showmessage('������ '+EditTask.Caption+' ��� ��������');
   exit;
 end;

 if  StatrTask(EditTask.Caption) then
 begin
 frmDomainInfo.Memo1.Lines.Add('������ ������ '+EditTask.Caption+' ������� ��������');
 EditTask.Close;
 end;
end;

function TEditTask.deletelistPC(z:bool):boolean;
var
i:integer;
begin
if z then
begin
 i:=MessageDlg('�� ������������� ������ ������� ���� ������ �����������?', mtConfirmation,[mbYes,mbCancel],0);
     if i=IDCancel then   exit;
end;
ListView2.SelectAll; //��������  ���� ������ ������
Button3.Click; //� �������
end;

procedure TEditTask.Button9Click(Sender: TObject);
begin
deletelistPC(true);
end;

function TEditTask.StopTask(DescriptionTask:string):boolean; // ��������� ������
var
StopTASK:string;
FDQueryStopTask:TFDQuery;
TransactionStopTask: TFDTransaction;
begin
try
TransactionStopTask:= TFDTransaction.Create(nil);
TransactionStopTask.Connection:=DataM.ConnectionDB;
TransactionStopTask.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
TransactionStopTask.Options.AutoCommit:=false;
TransactionStopTask.Options.AutoStart:=false;
TransactionStopTask.Options.AutoStop:=false;
FDQueryStopTask:= TFDQuery.Create(nil);
FDQueryStopTask.Transaction:=TransactionStopTask;
FDQueryStopTask.Connection:=DataM.ConnectionDB;
TransactionStopTask.StartTransaction; // ��������
StopTASK:=nametableForDescription(DescriptionTask);
FDQueryStopTask.SQL.Clear;
FDQueryStopTask.SQL.Text:='update TABLE_TASK set STATUS_TASK=:a ,START_STOP=:b WHERE NAME_TABLE='''+StopTASK+'''';
FDQueryStopTask.ParamByName('a').asstring:='��������� ������';
FDQueryStopTask.ParamByName('b').asstring:=booltostr(false); // ������������� ������ ������ � false ��� ���� ����� ���������� �
FDQueryStopTask.ExecSQL;
TransactionStopTask.Commit;  /// ��������� ���������
FDQueryStopTask.SQL.clear;   //��������
FDQueryStopTask.Close;  /// ������� ���
FDQueryStopTask.Free;
TransactionStopTask.Free;
result:=true;
except on E: Exception do
begin
 if Assigned(FDQueryStopTask) then FDQueryStopTask.Free;
      if Assigned(TransactionStopTask) then
      begin
      TransactionStopTask.Rollback;
      TransactionStopTask.Free;
      end;
 frmdomaininfo.memo1.Lines.add('������ ��������� ������ '+StopTASK+':'+E.Message);
 result:=false;
 end;
end
end;



procedure TEditTask.Button6Click(Sender: TObject); // ���������� ������
begin
if not StatusStartStopTask(EditTask.Caption) then // ���� ������ ����������� (false)
 begin
   frmDomainInfo.Memo1.Lines.Add('������ '+EditTask.Caption+' ��� �����������');
   exit;
 end;
if StopTask(EditTask.Caption) then
frmDomainInfo.Memo1.Lines.Add('��������, ��������� ������ '+EditTask.Caption+' ������� ���������');
end;

end.
