unit PcProperties;

interface

uses Winapi.Windows,System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls,
  System.Variants,Vcl.Dialogs, Vcl.ExtDlgs;

type

    Tcpu=record
    NameCPU:string;
    NumProc:string;
    FrequencyCPU:string;
    ArhCPU:string;
    end;

  TOKRightDlg12345678910111213 = class(TForm)
    ListView1: TListView;
    SpeedButton1: TSpeedButton;
    TreeView4: TTreeView;
    Panel1: TPanel;
    Panel2: TPanel;
    SpeedButton43: TSpeedButton;
    SpeedButton41: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton72: TSpeedButton;
    SpeedButton73: TSpeedButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    LVProgram: TListView;
    SpeedButton47: TSpeedButton;
    SpeedButton4: TSpeedButton;
    TabSheet3: TTabSheet;
    ListViewML: TListView;
    TabSheet4: TTabSheet;
    LVAV: TListView;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton67: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton41Click(Sender: TObject);
    procedure SpeedButton43Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton72Click(Sender: TObject);
    procedure SpeedButton73Click(Sender: TObject);
    procedure SpeedButton47Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton67Click(Sender: TObject);
  private
    function readPropertisPC:boolean;
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678910111213: TOKRightDlg12345678910111213;

implementation
uses umain,MyInventoryPC,DlgPCRename,MyDM,DlgJoinDomainOrWorkGroup,KMS,
  DlgUnjoinDomainOrWorkgroup,LoadCurencyList,MainPopup;
{$R *.dfm}





procedure TOKRightDlg12345678910111213.Button1Click(Sender: TObject);
begin
case PageControl1.ActivePageIndex of
0:
  begin // сохранить список оборудования
  frmDomainInfo.SaveTreeView(TreeView4,frmdomaininfo.ComboBox2.Text);
  end;
1:
  begin // сохранить список программ
   if LVProgram.Items.Count<>0 then
    begin
     frmDomainInfo.popupListViewSaveAs(LVProgram,'Сохранение списка программ',frmdomaininfo.ComboBox2.Text+' список программ');
    end
    else showmessage('Список пуст!');
  end;
2:
  begin
  if ListViewML.Items.Count<>0 then
    begin
     frmDomainInfo.popupListViewSaveAs(ListViewML,'Сохранение списка лицензий',frmdomaininfo.ComboBox2.Text+' список лицензий');
    end
    else showmessage('Список пуст!');
  end;
3:
  begin
  if LVAV.Items.Count<>0 then
    begin
     frmDomainInfo.popupListViewSaveAs(LVAV,'Сохранение списка антивирусного ПО',frmdomaininfo.ComboBox2.Text+' список антивирусного ПО');
    end
    else showmessage('Список пуст!');
  end;
end;

end;

procedure TOKRightDlg12345678910111213.FormShow(Sender: TObject);
begin
Caption:='Свойства '+ frmdomaininfo.Combobox2.Text;
MyUser:=frmDomainInfo.LabeledEdit1.Text;
MyPasswd:=frmDomainInfo.LabeledEdit2.Text;
Myps:= frmDomainInfo.ComboBox2.Text;
readPropertisPC; // читаем характеристики оборудования и строим дерево TreeView
frmDomainInfo.readSoftRorSelectPC(frmDomainInfo.ComboBox2.Text,LVProgram); // читаем список установленных программ
frmDomainInfo.readMicrosoftLic(frmDomainInfo.ComboBox2.Text,ListViewML); // читаем список лицензий
frmDomainInfo.readAntivirusStatus(frmDomainInfo.ComboBox2.Text,LVAV); // читаем список антивирусных продуктов
PageControl1.TabIndex:=0;
end;

procedure TOKRightDlg12345678910111213.SpeedButton1Click(Sender: TObject);
var
Mylist:TstringList;
i:integer;
SafeFileDlg:TSaveTextFileDialog;
begin
if listview1.Items.Count<>0 then
begin
SafeFileDlg:=TSaveTextFileDialog.Create(self);
SafeFileDlg.DefaultExt:='txt';
SafeFileDlg.Filter:='|*.txt';
SafeFileDlg.FilterIndex:=1;
SafeFileDlg.FileName:=frmdomaininfo.Combobox2.Text+'.txt';
  if SafeFileDlg.Execute then
    begin
    SafeFileDlg.Title:='Сохранить отчет';
    Mylist:=TstringList.Create;
    for i := 0 to listview1.Items.Count-1 do
    begin
    Mylist.add(listview1.Items[i].caption);
    end;
    Mylist.SaveToFile(SafeFileDlg.FileName);
    Mylist.Free;
    end;
FreeAndNil(SafeFileDlg);
end;

end;



procedure TOKRightDlg12345678910111213.SpeedButton2Click(Sender: TObject);
begin
if frmDomainInfo.ping(frmdomaininfo.Combobox2.Text) then frmDomainInfo.refreshinfoPC(frmdomaininfo.Combobox2.Text,frmdomaininfo.LabeledEdit1.Text,frmdomaininfo.LabeledEdit2.Text);
end;

procedure TOKRightDlg12345678910111213.SpeedButton3Click(Sender: TObject);
var
NewDescription:string;
res:string;

Begin
if not InputQuery('Описание компьютера в свойствах (Инв№)', 'Описание', NewDescription)
    then exit;
frmDomainInfo.PutInvNumberToDataBase(frmdomaininfo.Combobox2.Text,NewDescription);
frmDomainInfo.PutInvNumber(frmdomaininfo.Combobox2.Text,NewDescription);
End;

function TOKRightDlg12345678910111213.readPropertisPC:boolean;
begin
try
    DataM.FDQueryRead2.SQL.clear;      /// чтение конфигурации первого вхождения и передача в treeview
    DataM.FDQueryRead2.SQL.Text:='SELECT * FROM CONFIG_PC WHERE PC_NAME='''+
    (frmdomaininfo.Combobox2.Text)+''''+' ORDER BY DATE_INV DESC'; ////    DESC или ASC
    DataM.FDQueryRead2.Open;
    frmDomainInfo.createtreeView(DataM.FDQueryRead2,TreeView4); // Функция постороения дерева
    DataM.FDQueryRead2.SQL.clear;   //очистить
    DataM.FDQueryRead2.Close;  /// закрыть нах после чтения
 except
   on E: Exception do
   showmessage('Ошибка : '+e.Message);
   end;

end;

procedure TOKRightDlg12345678910111213.SpeedButton41Click(Sender: TObject);
begin
readPropertisPC; // читаем характеристики оборудования и строим дерево TreeView
frmDomainInfo.readSoftRorSelectPC(frmDomainInfo.ComboBox2.Text,LVProgram); // читаем список установленных программ
frmDomainInfo.readMicrosoftLic(frmDomainInfo.ComboBox2.Text,ListViewML); // читаем список лицензий
frmDomainInfo.readAntivirusStatus(frmDomainInfo.ComboBox2.Text,LVAV); // читаем список антивирусных продуктов
end;


procedure TOKRightDlg12345678910111213.SpeedButton43Click(Sender: TObject);
begin
if frmDomainInfo.ping(frmDomainInfo.Combobox2.Text) then
    begin
    OKRightDlg1234567.Caption:=frmDomainInfo.Combobox2.Text;
    OKRightDlg1234567.ShowModal;
    end;
end;


procedure TOKRightDlg12345678910111213.SpeedButton47Click(Sender: TObject);
begin
if frmDomainInfo.Ping(frmDomainInfo.ComboBox2.Text) then
frmDomainInfo.inventorySoftForSelectPC(frmDomainInfo.ComboBox2.Text,frmdomaininfo.LabeledEdit1.Text,frmdomaininfo.LabeledEdit2.Text);
end;

procedure TOKRightDlg12345678910111213.SpeedButton5Click(Sender: TObject);
begin
frmDomainInfo.InventoryMicrosoftPC(frmDomainInfo.combobox2.text,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
end;

procedure TOKRightDlg12345678910111213.SpeedButton67Click(Sender: TObject);
begin
if frmDomainInfo.ping(frmDomainInfo.Combobox2.Text) then
ReadAdnWriteAntivirusStaus(frmDomainInfo.Combobox2.Text,frmDomainInfo.labeledEdit1.Text,frmDomainInfo.labeledEdit2.Text,LVAV);
end;

procedure TOKRightDlg12345678910111213.SpeedButton6Click(Sender: TObject);
begin
FormKMS.ShowModal;
end;

procedure TOKRightDlg12345678910111213.SpeedButton72Click(Sender: TObject);
begin
if frmDomainInfo.ping(frmDomainInfo.Combobox2.Text) then
begin
OKRightDlg123456.ShowModal;
end;
end;

procedure TOKRightDlg12345678910111213.SpeedButton73Click(Sender: TObject);
begin
OKRightDlg12345678.ShowModal;
end;

end.
