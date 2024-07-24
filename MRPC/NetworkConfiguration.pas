unit NetworkConfiguration;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Mask,ComObj,System.Variants,Vcl.OleServer,ActiveX;

type
  TOKRightDlg123456789101112131415 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    MaskEdit1: TMaskEdit;
    MaskEdit2: TMaskEdit;
    MaskEdit3: TMaskEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Panel1: TPanel;
    CheckBox3: TCheckBox;
    Panel2: TPanel;
    CheckBox4: TCheckBox;
    MaskEdit4: TMaskEdit;
    MaskEdit5: TMaskEdit;
    Label4: TLabel;
    Label5: TLabel;
    SpeedButton1: TSpeedButton;
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure MaskEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure MaskEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure MaskEdit3KeyPress(Sender: TObject; var Key: Char);
    procedure MaskEdit4KeyPress(Sender: TObject; var Key: Char);
    procedure MaskEdit5KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  OKRightDlg123456789101112131415: TOKRightDlg123456789101112131415;
  NewNetworkDHCPEnable,NewNetworkAdapterStaticIPALL,
  NewPropertiesNetworkAdapter,NewNetworkAdapterDNS:Tthread;
  IPAddress,IPSubnetMask,IPGateway,IPDNS,IPDNS2:string;
  DHCPBool,YesIPAddress,YesMask,
  YesGateway,YesDNS1,YesDNS2,YesWINS,YesWins1:string;
  masIP,MasSubnet,MasGateway,masDNS,MasGatewayMetric: variant;


implementation
uses umain,NetworkAdapterDHCP,NetworkAdapterStaticTCPIPAll,
PropertiesNetworkThread,NetworkAdapterDNSThread,NetworkAdapterAdditionallyDialog;
{$R *.dfm}
/////////////////////////////////////////////////////////////////////////
function ArrayToVarArray(Arr : Array Of string):OleVariant;overload;
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
/////////////////////////////////////////////////////////////////////////
procedure TOKRightDlg123456789101112131415.CheckBox1Click(Sender: TObject);
begin
if checkBox1.Checked=true then
begin
  checkBox2.Checked:=false;
  checkBox3.Enabled:=true;
  panel1.Enabled:=false;
  MaskEdit1.color:=clSilver;
  MaskEdit2.color:=clSilver;
  MaskEdit3.color:=clSilver;
  MaskEdit1.Text:='';
  MaskEdit2.Text:='';
  MaskEdit3.Text:='';
  OkBtn.Enabled:=true;

end;
end;

procedure TOKRightDlg123456789101112131415.CheckBox2Click(Sender: TObject);
begin
if checkBox2.Checked=true then
begin
  checkBox1.Checked:=false;
  checkBox3.Checked:=false;
  checkBox3.Enabled:=false;
  checkBox4.Checked:=true;
  panel2.Enabled:=true;
  MaskEdit1.color:=clWindow;
  MaskEdit2.color:=clWindow;
  MaskEdit3.color:=clWindow;
  MaskEdit4.color:=clWindow;
  MaskEdit5.color:=clWindow;
  panel1.Enabled:=true;
  OkBtn.Enabled:=true;
end;
end;

procedure TOKRightDlg123456789101112131415.CheckBox3Click(Sender: TObject);
begin
if checkBox3.Checked=true then
begin
  checkBox4.Checked:=false;
  panel2.Enabled:=false;
    MaskEdit4.color:=clSilver;
    MaskEdit5.color:=clSilver;
    MaskEdit4.Text:='';
    MaskEdit5.Text:='';
    IPDNS:='';
    IPDNS2:='';
    OkBtn.Enabled:=true;
end;
end;

procedure TOKRightDlg123456789101112131415.CheckBox4Click(Sender: TObject);
begin
if checkBox4.Checked=true then
begin
  checkBox3.Checked:=false;
  panel2.Enabled:=true;
  MaskEdit4.color:=clWindow;
  MaskEdit5.color:=clWindow;
  OkBtn.Enabled:=true;
end;
end;

procedure TOKRightDlg123456789101112131415.FormShow(Sender: TObject);
var
  FSWbemLocator   : OLEVariant;
  FWMIService     : OLEVariant;
  FWbemObjectSet  : OLEVariant;
  MyError       : integer;
   FWbemObject:    OLEVariant;
  oEnum : IEnumvariant;
  i             : integer;
  iValue        : LongWord;

begin
 try
begin
  caption:=frmDomainInfo.listview7.Selected.SubItems[6];
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Загрузка данный сетевого интерфейса');
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
                    DHCPBool:=(vartostr(FWbemObject.DHCPEnabled));
                    masIP:=((FWbemObject.IPAddress));
                    if VarType(masIP) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(masIP,1) to VarArrayHighBound(masIP,1) do
                        if i=0 then YesIPAddress:=(string(masIP[i]));
                      end;
                    MasSubnet:=null;
                    MasSubnet:=(FWbemObject.IPSubnet);
                    if VarType(MasSubnet) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(MasSubnet,1) to VarArrayHighBound(MasSubnet,1) do
                        if i=0 then YesMask:=(string(MasSubnet[i]));
                      end;
                    MasGateway:=null;
                    MasGateway:=(FWbemObject.DefaultIPGateway);
                    if VarType(MasGateway) and VarTypeMask=varVariant then
                       begin
                        for i :=VarArrayLowBound(MasGateway,1) to VarArrayHighBound(MasGateway,1) do
                        if i=0 then YesGateway:=(string(MasGateway[i]));
                      end;
                    masDNS:=null;
                    masDNS:=(FWbemObject.DNSServerSearchOrder);
                    if VarType(masDNS) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(masDNS,1) to VarArrayHighBound(masDNS,1) do
                        begin
                        if i=0 then YesDNS1:=(string(masDNS[i]));
                        if i=1 then YesDNS2:=(string(masDNS[i]));
                        end;
                      end;
                     if FWbemObject.WINSPrimaryServer<>null then /// wins сервер
                          YesWins:=(string(FWbemObject.WINSPrimaryServer));
                     if FWbemObject.WINSSecondaryServer<>null then /// wins сервер
                          YesWins1:=(string(FWbemObject.WINSSecondaryServer));

              FSWbemLocator:= Unassigned;
              FWbemObject:=Unassigned;
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

if DHCPBool='True' then
  begin
  checkBox1.Enabled:=true;
  checkBox1.Checked:=true;
  checkBox2.Checked:=false;
  checkBox3.Enabled:=true;
  checkBox3.Checked:=true;
  checkBox4.Checked:=false;
  panel1.Enabled:=false;
  MaskEdit1.color:=clSilver;
  MaskEdit2.color:=clSilver;
  MaskEdit3.color:=clSilver;
  MaskEdit1.Text:='';
  MaskEdit2.Text:='';
  MaskEdit3.Text:='';
  panel2.Enabled:=false;
  MaskEdit4.color:=clSilver;
  MaskEdit5.color:=clSilver;
  MaskEdit4.Text:='';
  MaskEdit5.Text:='';
  OkBtn.Enabled:=false;
  SpeedButton1.Enabled:=false;
  end;
if DHCPBool='False' then
  begin
  checkBox1.Enabled:=true;
  checkBox1.Checked:=False;
  checkBox2.Checked:=true;
  checkBox3.Enabled:=true;
  checkBox3.Checked:=False;
  checkBox4.Checked:=true;
  checkBox4.Enabled:=true;
  panel1.Enabled:=True;
  MaskEdit1.color:=clWindow;
  MaskEdit2.color:=clWindow;
  MaskEdit3.color:=clWindow;
  MaskEdit1.Text:=YesIPAddress;
  MaskEdit2.Text:=YesMask;
  MaskEdit3.Text:=YesGateway;
  panel2.Enabled:=True;
  MaskEdit4.color:=clWindow;
  MaskEdit5.color:=clWindow;
  MaskEdit4.Text:=YesDNS1;
  MaskEdit5.Text:=YesDNS2;
  SpeedButton1.Enabled:=true;
  OkBtn.Enabled:=false;
  end;


end;

procedure TOKRightDlg123456789101112131415.MaskEdit1KeyPress(Sender: TObject;
  var Key: Char);
begin
OkBtn.Enabled:=true;
end;

procedure TOKRightDlg123456789101112131415.MaskEdit2KeyPress(Sender: TObject;
  var Key: Char);
begin
OkBtn.Enabled:=true;
end;

procedure TOKRightDlg123456789101112131415.MaskEdit3KeyPress(Sender: TObject;
  var Key: Char);
begin
OkBtn.Enabled:=true;
end;

procedure TOKRightDlg123456789101112131415.MaskEdit4KeyPress(Sender: TObject;
  var Key: Char);
begin
OkBtn.Enabled:=true;
end;

procedure TOKRightDlg123456789101112131415.MaskEdit5KeyPress(Sender: TObject;
  var Key: Char);
begin
OkBtn.Enabled:=true;
end;

procedure TOKRightDlg123456789101112131415.OKBtnClick(Sender: TObject);
begin
if (checkBox1.Checked=true) and (checkBox3.Checked=true) then
begin          ////////////// включаем DHCP
   NewNetworkDHCPEnable:=NetworkAdapterDHCP.NetworkAdapterDHCPEnable.Create(false);
   exit;
end;

if (checkBox2.Checked=true) and (checkBox4.Checked=true) then
begin
 if (MaskEdit1.text<>'') and(MaskEdit2.text<>'') then
   begin
   MasIP:=VarArrayCreate([0,0],VarVariant);
   MasSubnet:=VarArrayCreate([0,0],VarVariant);
   MasIP[0]:=MaskEdit1.text;
   MasSubnet[0]:=MaskEdit2.text;
   end;
 if MaskEdit3.text<>'' then
   begin
   MasGateway:=VarArrayCreate([0,0],VarVariant);
   MasGatewayMetric:=VarArrayCreate([0,0],VarVariant);
   MasGateway[0]:=MaskEdit3.text;
   MasGatewayMetric[0]:='1';
   end;
 if MaskEdit4.text<>'' then
   begin
    MasDNS:=VarArrayCreate([0,0],VarVariant);
    MasDNS[0]:=MaskEdit4.text;
   end;
 if (MaskEdit4.text<>'')and(MaskEdit5.text<>'')  then
   begin
    MasDNS:=VarArrayCreate([0,1],VarVariant);
    MasDNS[0]:=MaskEdit4.text;
    MasDNS[1]:=MaskEdit5.text;
   end;
 NewNetworkAdapterStaticIPALL:=NetworkAdapterStaticTCPIPALL.NetworkAdapterStaticIPALL.Create(false);
 exit;
 end;

 if (checkBox1.Checked=true) and (checkBox4.Checked=true) then
begin   /////включаем DHCP и ручками назначаем DNS
   if MaskEdit4.text<>'' then MasDNS:=ArrayToVarArray([MaskEdit4.text]);
   if (MaskEdit4.text<>'')and(MaskEdit5.text<>'')  then MasDNS:=ArrayToVarArray([MaskEdit4.text,MaskEdit5.text]);
    NewNetworkDHCPEnable:=NetworkAdapterDHCP.NetworkAdapterDHCPEnable.Create(false);
   //NewNetworkAdapterDNS:=NetworkAdapterDNSThread.NetworkAdapterDNS.Create(false);
   exit;
end;

end;

procedure TOKRightDlg123456789101112131415.SpeedButton1Click(Sender: TObject);
begin
OKRightDlg12345678910111213141516.ShowModal;
end;

end.
