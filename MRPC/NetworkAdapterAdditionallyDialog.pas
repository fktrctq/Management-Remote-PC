unit NetworkAdapterAdditionallyDialog;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.OleServer,ActiveX,ComObj,System.Variants;

type
  TOKRightDlg12345678910111213141516 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    ListView1: TListView;
    ListView2: TListView;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    ListBox1: TListBox;
    Label1: TLabel;
    Button5: TButton;
    Button6: TButton;
    ListBox2: TListBox;
    Button8: TButton;
    Button9: TButton;
    Label2: TLabel;
    Button7: TButton;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg12345678910111213141516: TOKRightDlg12345678910111213141516;
  TransformIP,TransformGateway,TransformDNSWINS:bool;
  DNSWINS:string;
implementation
uses umain,NetworkConfiguration,NetworkAdapterAddIPSubnet,
NetworkAdapterAddgateway,NetworkAdapterAddDnsWins,NetworkAdapterStaticTCPIPall;
{$R *.dfm}
 ////////////////////////////////////////////////
  function ArrayToVarArray(Arr : Array Of string):OleVariant; overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
//////////////////////////////////////////////////
function ArrayToVarArray(Arr : Array Of Word):OleVariant;overload;
var
 i : integer;
begin
    Result   :=VarArrayCreate([0, High(Arr)], varVariant);
    for i:=Low(Arr) to High(Arr) do
     Result[i]:=Arr[i];
end;
////////////////////////////////////////////////////
procedure TOKRightDlg12345678910111213141516.Button10Click(Sender: TObject);
begin   ///// изменяем шлюз
if (Listview2.Selected<>nil) then
  begin
  transformGateway:=true;
  OKRightDlg123456789101112131415161718.Edit1.Text:=Listview2.Items[Listview2.Selected.Index].Caption;
  OKRightDlg123456789101112131415161718.ShowModal;
  end;

end;

procedure TOKRightDlg12345678910111213141516.Button11Click(Sender: TObject);
begin         //// изменяем DNS

if listbox1.SelCount<>0 then
  begin
  DNSWINS:='DNS';
  TransformDNSWINS:=true;
  OKRightDlg12345678910111213141516171819.Caption:='DNS-адрес TCP/IP';
  OKRightDlg12345678910111213141516171819.Label1.Caption:='DNS-сервер';
  OKRightDlg12345678910111213141516171819.Edit1.Text:=listbox1.Items[Listbox1.ItemIndex];
  OKRightDlg12345678910111213141516171819.ShowModal;
  end;

end;

procedure TOKRightDlg12345678910111213141516.Button12Click(Sender: TObject);
begin             ////// изменяем WINS
if listbox2.SelCount<>0 then
  begin
  DNSWINS:='WINS';
  TransformDNSWINS:=true;
  OKRightDlg12345678910111213141516171819.Caption:='WINS-адрес TCP/IP';
  OKRightDlg12345678910111213141516171819.Label1.Caption:='WINS-сервер';
  OKRightDlg12345678910111213141516171819.Edit1.Text:=listbox2.Items[Listbox2.ItemIndex];
  OKRightDlg12345678910111213141516171819.ShowModal;
  end;
end;

procedure TOKRightDlg12345678910111213141516.Button1Click(Sender: TObject);
begin                      /////// добавление IP адреса
OKRightDlg1234567891011121314151617.Edit1.Text:='';
OKRightDlg1234567891011121314151617.Edit2.Text:='';
TransformIP:=false;
OKRightDlg1234567891011121314151617.ShowModal;
end;

procedure TOKRightDlg12345678910111213141516.Button2Click(Sender: TObject);
begin                    ///// удаляем строку с ip адресом
if Listview1.Selected<>nil then
Listview1.Items.Delete(listview1.Selected.Index);
end;

procedure TOKRightDlg12345678910111213141516.Button3Click(Sender: TObject);
begin
transformGateway:=false;    ////// добавляем шлюз
OKRightDlg123456789101112131415161718.Edit1.Text:='';
OKRightDlg123456789101112131415161718.ShowModal;

end;

procedure TOKRightDlg12345678910111213141516.Button4Click(Sender: TObject);
begin                    ///// удаляем строку с шлюзом
if Listview2.Selected<>nil then
Listview2.Items.Delete(listview2.Selected.Index);
end;

procedure TOKRightDlg12345678910111213141516.Button5Click(Sender: TObject);
begin         //// добавляем DNS
DNSWINS:='DNS';
TransformDNSWINS:=false;
OKRightDlg12345678910111213141516171819.Caption:='DNS-адрес TCP/IP';
OKRightDlg12345678910111213141516171819.Label1.Caption:='DNS-сервер';
OKRightDlg12345678910111213141516171819.Edit1.Text:='';
OKRightDlg12345678910111213141516171819.ShowModal;
end;

procedure TOKRightDlg12345678910111213141516.Button6Click(Sender: TObject);
begin
if listbox1.Count<>0 then listbox1.DeleteSelected;
end;

procedure TOKRightDlg12345678910111213141516.Button7Click(Sender: TObject);
begin                      ////// изменение ip адреса
if (listview1.Items.Count<>0) and (Listview1.Selected<>nil) then
begin
OKRightDlg1234567891011121314151617.Edit1.Text:=Listview1.Items[Listview1.Selected.Index].Caption;
OKRightDlg1234567891011121314151617.Edit2.Text:=Listview1.Items[Listview1.Selected.Index].SubItems[0];
TransformIP:=true;
OKRightDlg1234567891011121314151617.ShowModal;

end;

end;

procedure TOKRightDlg12345678910111213141516.Button8Click(Sender: TObject);
begin             ////// добавляем WINS
DNSWINS:='WINS';
TransformDNSWINS:=false;
OKRightDlg12345678910111213141516171819.Caption:='WINS-адрес TCP/IP';
OKRightDlg12345678910111213141516171819.Label1.Caption:='WINS-сервер';
OKRightDlg12345678910111213141516171819.Edit1.Text:='';
OKRightDlg12345678910111213141516171819.ShowModal;
end;

procedure TOKRightDlg12345678910111213141516.Button9Click(Sender: TObject);
begin
if listbox2.Count<>0 then listbox2.DeleteSelected;
end;

procedure TOKRightDlg12345678910111213141516.FormShow(Sender: TObject);
var
//  FSWbemLocator   : OLEVariant;
//  FWMIService     : OLEVariant;
//  FWbemObjectSet  : OLEVariant;
//  MyError       : integer;
  i             : integer;
begin
 try
begin
  listView1.Clear;
  ListView2.Clear;
  ListBox1.Clear;
  ListBox2.Clear;
  button1.Enabled:=true;
  button2.Enabled:=true;
  button3.Enabled:=true;
  button4.Enabled:=true;
  button5.Enabled:=true;
  button6.Enabled:=true;
  button7.Enabled:=true;
  button8.Enabled:=true;
  button9.Enabled:=true;
  button10.Enabled:=true;
  button11.Enabled:=true;
  button12.Enabled:=true;

           if DHCPBool<>'True' then
             begin
               if VarType(masIP) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(masIP,1) to VarArrayHighBound(masIP,1)-1 do
                          begin //// добавляем список IP адресов
                          ListView1.Items.Add;
                          ListView1.Items[i].Caption:=(string(masIP[i]));
                          end;
                      end;
                if VarType(masSubnet) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(masSubnet,1) to VarArrayHighBound(masSubnet,1)-1 do
                        begin  //// добавляем список масок
                        ListView1.Items[i].SubItems.add(string(masSubnet[i]));
                        end;
                      end;
                if VarType(masGateway) and VarTypeMask=varVariant then
                       begin
                        for i :=VarArrayLowBound(masGateway,1) to VarArrayHighBound(masGateway,1) do
                        begin //// добавляем список шлюзов
                        ListView2.Items.Add.Caption:=(string(masGateway[i]));
                        end;
                      end;
                 if VarType(masDNS) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(masDNS,1) to VarArrayHighBound(masDNS,1) do
                        begin  //// добавляем списк DNS серверов
                        listBox1.Items.Add(string(masDNS[i]));
                        end;
                      end;
                  if YesWins<>'' then /// wins сервер
                  listBox2.Items.Add(YesWins);
                  if YesWins1<>'' then /// wins сервер
                  listBox2.Items.Add(YesWins1);
             end
             else
             begin
               ListView1.Items.Add;
               ListView1.Items[i].Caption:=('DHCP включен');
               button1.Enabled:=false;
               button2.Enabled:=false;
               button3.Enabled:=false;
               button4.Enabled:=false;
               button5.Enabled:=false;
               button6.Enabled:=false;
               button7.Enabled:=false;
               button8.Enabled:=false;
               button9.Enabled:=false;
               button10.Enabled:=false;
               button11.Enabled:=false;
               button12.Enabled:=false;
             end;

end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(' - Ошибка получения данных "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');

           exit;
         end;

end;
end;

procedure TOKRightDlg12345678910111213141516.OKBtnClick(Sender: TObject);
var
i:integer;
NEWIPSettings:TThread;
begin
try
YesWins:='';
YesWins1:='';
if listview1.Items.Count<>0 then
    begin  //// добавление ip b mask в массивы
      MasIP:=VarArrayCreate([0,listview1.Items.Count-1],VarVariant);
      MasSubnet:=VarArrayCreate([0,listview1.Items.Count-1],VarVariant);
      for I := 0 to listview1.Items.Count-1 do
        begin
          MasIP[i]:=listview1.Items[i].Caption;
          MasSubnet[i]:=listview1.Items[i].SubItems[0];
         // frmDomainInfo.memo1.Lines.Add('IP+ mask -'+inttostr(i));
        end;
    end;
if listview2.Items.Count<>0 then
    begin  ///// добавление шлюзов в массив
      MasGateway:=VarArrayCreate([0,listview2.Items.Count-1],VarVariant);
      MasGatewayMetric:=VarArrayCreate([0,listview2.Items.Count-1],VarVariant);
      for I := 0 to listview2.Items.Count-1 do
        begin
         masGateway[i]:=(listview2.Items[i].Caption);
         MasGatewayMetric[i]:= inttostr(i+1);
       //  frmDomainInfo.memo1.Lines.Add('Gateway -'+inttostr(i));
        end;
    end;
if listbox1.Items.Count<>0 then
    begin   ///// добавление DNS
     MasDNS:=VarArrayCreate([0,listbox1.Items.Count-1],VarVariant);
      for I := 0 to listbox1.Items.Count-1 do
        begin
          MasDNS[i]:=listbox1.Items[i];
         // frmDomainInfo.memo1.Lines.Add('DNS -'+inttostr(i));
        end;
    end;
if listbox2.Items.Count<>0 then
    begin       //// добавление WINS
      for I := 0 to listbox2.Items.Count-1 do
        begin
          if i=0 then YesWins:=listbox2.Items[0];
          if i=1 then YesWins1:=listbox2.Items[1];
        end;
    end;
NEWIPSettings:=NetworkAdapterStaticTCPIPAll.NetworkAdapterStaticIPALL.Create(false);
OKRightDlg123456789101112131415.SpeedButton1.Enabled:=False;
OKRightDlg123456789101112131415.OKBtn.Enabled:=false;
close;

except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка назначения статических параметров адаптера "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;

end.
