unit MainPopup;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ImgList, Vcl.Menus,Clipbrd,
  Vcl.ComCtrls, Vcl.StdCtrls,ShellApi,FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client;

type
  TMainFormPopup = class(TForm)
    PCPower: TPopupMenu;
    WOL1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    ImageListMainMenu: TImageList;
    PCTask: TPopupMenu;
    PCApplication: TPopupMenu;
    PCOther: TPopupMenu;
    PCInventory: TPopupMenu;
    PCActivWinOffcie: TPopupMenu;
    Windows1: TMenuItem;
    Office1: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    SMARTHDD1: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    RDP1: TMenuItem;
    uRDM1: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    ListPCPower: TPopupMenu;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    ListPCPrinter: TPopupMenu;
    N22: TMenuItem;
    N23: TMenuItem;
    ListPCProgramMSI: TPopupMenu;
    N24: TMenuItem;
    N25: TMenuItem;
    ListPCProcess: TPopupMenu;
    N26: TMenuItem;
    N27: TMenuItem;
    N28: TMenuItem;
    ListPCInventory: TPopupMenu;
    N29: TMenuItem;
    N30: TMenuItem;
    ListPCApplication: TPopupMenu;
    RDP2: TMenuItem;
    N31: TMenuItem;
    ListPCTask: TPopupMenu;
    N32: TMenuItem;
    N33: TMenuItem;
    N34: TMenuItem;
    KMS1: TMenuItem;
    Microsoft1: TMenuItem;
    WindowsOffice1: TMenuItem;
    PopupAvProduct: TPopupMenu;
    N35: TMenuItem;
    N36: TMenuItem;
    N37: TMenuItem;
    N38: TMenuItem;
    PopupClipboard: TPopupMenu;
    N39: TMenuItem;
    ListView1: TListView;
    N40: TMenuItem;
    procedure WOL1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure Windows1Click(Sender: TObject);
    procedure Office1Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure SMARTHDD1Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure RDP1Click(Sender: TObject);
    procedure uRDM1Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
    procedure N25Click(Sender: TObject);
    procedure N26Click(Sender: TObject);
    procedure N27Click(Sender: TObject);
    procedure N28Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure RDP2Click(Sender: TObject);
    procedure N31Click(Sender: TObject);
    procedure N32Click(Sender: TObject);
    procedure N33Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N34Click(Sender: TObject);
    procedure KMS1Click(Sender: TObject);
    procedure Microsoft1Click(Sender: TObject);
    procedure WindowsOffice1Click(Sender: TObject);
    procedure N36Click(Sender: TObject);
    procedure N35Click(Sender: TObject);
    procedure N37Click(Sender: TObject);
    procedure N38Click(Sender: TObject);
    procedure N39Click(Sender: TObject);
    function DlgCloseForm(title,but1,but2,but3:string):integer;
    procedure N41Click(Sender: TObject); // диалог при закрытии программы
  private

  public
    { Public declarations }
  end;

var
  MainFormPopup: TMainFormPopup;

implementation
uses umain,My_proc,MyDM;
{$R *.dfm}





procedure TMainFormPopup.KMS1Click(Sender: TObject); // активация KMS
begin
frmDomainInfo.ActivationKMS;
end;



procedure TMainFormPopup.Microsoft1Click(Sender: TObject);  // Инвентаризация продуктов Microsoft
begin
frmDomainInfo.InventoryMicrosoft;
end;


procedure TMainFormPopup.N10Click(Sender: TObject); // добавить в домен
begin
frmDomainInfo.AddPCDomain;
end;

procedure TMainFormPopup.N11Click(Sender: TObject); // вывести комп из домена
begin
frmDomainInfo.RemovePCDomain;
end;

procedure TMainFormPopup.N12Click(Sender: TObject); // свойства ПК
begin
frmDomainInfo.PropertiesForPC;
end;

procedure TMainFormPopup.N13Click(Sender: TObject); // Присвоить инвентарный номер
begin
frmDomainInfo.RenewInventNumber;
end;

procedure TMainFormPopup.N14Click(Sender: TObject); // редактирование пользователи и группы
begin
frmDomainInfo.EditUserGroup;
end;

procedure TMainFormPopup.N15Click(Sender: TObject);  // открыть проводник
begin
frmDomainInfo.OpenFormExplorer;
end;

procedure TMainFormPopup.N16Click(Sender: TObject); // запуск задачи для текущего ПК
begin
frmDomainInfo.RunTaskForPC;
end;

procedure TMainFormPopup.N17Click(Sender: TObject);   // Создание новой задачи для текущего ПК
begin
frmDomainInfo.CreateNewTaskForPC;
end;

procedure TMainFormPopup.N18Click(Sender: TObject);  // включить группу компьютеров
begin
frmDomainInfo.PowerForListPC;
end;

procedure TMainFormPopup.N19Click(Sender: TObject);  // завершение работы на группе компьютеров
begin
frmDomainInfo.PowerOffForListPC;
end;

procedure TMainFormPopup.N1Click(Sender: TObject); // завершение сеанса на текущем компе
begin
frmDomainInfo.CloseSession;
end;

procedure TMainFormPopup.N20Click(Sender: TObject); // перезагрузка для группы компов
begin
frmDomainInfo.ResetForListPC;
end;

procedure TMainFormPopup.N21Click(Sender: TObject); // Завершение сеанса на группе компов
begin
frmDomainInfo.LogOutForListPC;
end;

procedure TMainFormPopup.N22Click(Sender: TObject);  // подключить принтер для списка компов
begin
frmDomainInfo.AddPrinterForListPC;
end;

procedure TMainFormPopup.N23Click(Sender: TObject);  // установка драйвера для списка компов
begin
frmDomainInfo.AddDriverPrintForListPC;
end;

procedure TMainFormPopup.N24Click(Sender: TObject); // Установка программы на группу компов
begin
frmDomainInfo.InstallMSIForListPC;
end;

procedure TMainFormPopup.N25Click(Sender: TObject); //Удаление программ MSI на группе компьютеров
begin
frmDomainInfo.DeleteProgramMSIForListPC;
end;

procedure TMainFormPopup.N26Click(Sender: TObject);  // запуск процесса на группе компьютеров
begin
frmDomainInfo.RunProcessForListPC;
end;

procedure TMainFormPopup.N27Click(Sender: TObject); // завершить процесс на группе компьютеров
begin
frmDomainInfo.KillProcessForListPC;
end;

procedure TMainFormPopup.N28Click(Sender: TObject); //Поиск процесса на группе компов
begin
frmDomainInfo.FindProcessForListPC;
end;

procedure TMainFormPopup.N29Click(Sender: TObject);  // инвентаризация оборудования
begin
frmDomainInfo.SpeedButton13.Click;
end;

procedure TMainFormPopup.N2Click(Sender: TObject);  // перезагрузить текущий комп
begin
frmDomainInfo.ResetCurPC;
end;

procedure TMainFormPopup.N30Click(Sender: TObject); // Инвентаризация программ
begin
frmDomainInfo.SpeedButton24.Click;
end;

procedure TMainFormPopup.N31Click(Sender: TObject); // Операции удаления и копирования каталогов и фавйлов на группу компов
begin
frmDomainInfo.CopyDelFFForListPC;
end;

procedure TMainFormPopup.N32Click(Sender: TObject);  // запуск задачи для группы компов
begin
frmDomainInfo.RunTaskForListPC;
end;

procedure TMainFormPopup.N33Click(Sender: TObject);   // создать новую задачу для группы компов
begin
frmDomainInfo.CreateNewTaskForListPC;
end;

procedure TMainFormPopup.N34Click(Sender: TObject);
begin
frmDomainInfo.OpenRestoreForm;
end;

procedure TMainFormPopup.N35Click(Sender: TObject);
begin
frmDomainInfo.SpeedButton59.Click;
end;

procedure TMainFormPopup.N36Click(Sender: TObject);
begin
frmDomainInfo.UpdateAntivirusProduct;
end;

procedure TMainFormPopup.N37Click(Sender: TObject);
begin
if frmDomainInfo.LVAntivirus.Items.Count<>0 then
     frmDomainInfo.popupListViewSaveAs(frmDomainInfo.LVAntivirus,'Сохранение списка антивирусных продуктов','Антивирусное ПО');

end;

procedure TMainFormPopup.N38Click(Sender: TObject);
begin
frmDomainInfo.LVAntivirusDblClick(frmDomainInfo.LVAntivirus);
end;



procedure TMainFormPopup.N39Click(Sender: TObject);
var
s:string;
Caller: TObject;
i,z:integer;
function CopyToClip(s:string):boolean;
begin
Clipboard.astext:=s;
end;
begin
s:='';
Caller := ((Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent;
if (Caller is TlistView) then
Begin
  if (Caller as TlistView).SelCount=0 then exit;
  for z := 0 to (Caller as TlistView).Items.Count-1 do
  if (Caller as TlistView).Items[z].Selected then
   begin
    if (Caller as TlistView).Items[z].Caption<>'' then s:=s+(Caller as TlistView).Columns[0].Caption+' - '+(Caller as TlistView).Items[z].Caption+' * ';
    for I := 0 to (Caller as TlistView).Columns.Count-2 do
    begin
      s:=s+(Caller as TlistView).Columns[i+1].Caption+' - '+(Caller as TlistView).Items[z].SubItems[i]+' * ';
    end;
    s:=s+#10;
  end;
  CopyToClip(s);    
End;

if (caller is TTreeView) then  
Begin
 for z := 1 to (caller as TTreeView).Items.Count-1 do
 if (caller as TTreeView).Items[z].Selected then 
 begin
 s:=s+(caller as TTreeView).Items[z].Text+' * ';
 End; 
 CopyToClip(s);
End;

end;

procedure TMainFormPopup.N3Click(Sender: TObject);  // Завершение работы на текущем компьютере
begin
frmDomainInfo.PowerOff;
end;

procedure TMainFormPopup.N41Click(Sender: TObject);
begin
EditMyProc.ShowModal;
end;

procedure TMainFormPopup.N4Click(Sender: TObject); // форсировать завершение сеанса
begin
frmDomainInfo.ForceCloseSession;
end;

procedure TMainFormPopup.N5Click(Sender: TObject);  // форсировать перезагрузку
begin
frmDomainInfo.rebuutSelectPC;
end;

procedure TMainFormPopup.N6Click(Sender: TObject); //форсировать завершение работы
begin
 frmDomainInfo.ForcePowerOff;
end;

procedure TMainFormPopup.N7Click(Sender: TObject); // инвентаризация оборудования
begin
frmDomainInfo.InventoryHardware;
end;

procedure TMainFormPopup.N8Click(Sender: TObject); // инвентаризация софта
begin
frmDomainInfo.InventorySoftware;
end;

procedure TMainFormPopup.N9Click(Sender: TObject); // переименовать комп
begin
frmDomainInfo.EditNamePC;
end;

procedure TMainFormPopup.Office1Click(Sender: TObject); //активация office
begin
frmDomainInfo.ActivationOffice;
end;

procedure TMainFormPopup.RDP1Click(Sender: TObject); // открыть новое RDP соединение для компа
begin
frmDomainInfo.NewRDPFormForPC;
end;

procedure TMainFormPopup.RDP2Click(Sender: TObject); // Мультиоконный RDP клиент
begin
frmDomainInfo.OpenRDPForListPC;
end;

procedure TMainFormPopup.SMARTHDD1Click(Sender: TObject);  // смарт HDD
begin
frmDomainInfo.OpenFormSMATR;
end;

procedure TMainFormPopup.uRDM1Click(Sender: TObject);  // Открыть форму для Urdm
begin
frmDomainInfo.OpenFormForuRDM;
end;

procedure TMainFormPopup.Windows1Click(Sender: TObject);  // Активация windows
begin
frmDomainInfo.ActivationWindowsPC;
end;

procedure TMainFormPopup.WindowsOffice1Click(Sender: TObject); // инвентаризация Windows и Office
begin
frmDomainInfo.InventoryMicrosoftPC(frmDomainInfo.combobox2.text,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
end;

procedure TMainFormPopup.WOL1Click(Sender: TObject);  // Включить текущий компьютер
begin
frmDomainInfo.PowerOn;
end;



function TMainFormPopup.DlgCloseForm(title,but1,but2,but3:string):integer;
var
DlgFrm:Tform;
bOk,bCancel,bNo:Tbutton;
begin
try
 DlgFrm:=Tform.Create(self);
 DlgFrm.Caption:=title;
 DlgFrm.Height:=100;
 DlgFrm.Width:=440;
 DlgFrm.BorderStyle:=bsSingle;
 DlgFrm.BorderIcons:=[biSystemMenu];
 DlgFrm.Icon:=MainFormPopup.Icon;
 DlgFrm.Position:=poMainFormCenter;

 bCancel:=TButton.Create(DlgFrm);
 bCancel.Parent:=DlgFrm;
 bCancel.Left:=10;
 bCancel.Top:=20;
 bCancel.Width:=100;
 bCancel.Caption:=but1;
 bCancel.ModalResult:=1;

 bOk:=TButton.Create(DlgFrm);
 bOk.Parent:=DlgFrm;
 bOk.Left:=130;
 bOk.Top:=20;
 bOk.Width:=160;
 bOk.Caption:=but2;
 bok.ModalResult:=3;

 bNo:=TButton.Create(DlgFrm);
 bNo.Parent:=DlgFrm;
 bNo.Left:=310;
 bNo.Top:=20;
 bNo.Width:=100;
 bNo.Caption:=but3;
 bNo.ModalResult:=2; // если форму закрыть крестиком то result=2
 result:=DlgFrm.ShowModal;
except on E: Exception do
begin
  ShowMessage('Ошибочка вышла... '+e.Message);
  result:=3; // завершить принудительной
end;
end;
 end;

end.
