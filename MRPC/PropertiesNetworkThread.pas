unit PropertiesNetworkThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  PropertiesNetworkInterface = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain,propertiesNetworkDialog;




{ PropertiesNetworkInterface }

procedure PropertiesNetworkInterface.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  MyError       : integer;
  mas           : variant;
  res           : string;
  i             : integer;
  iValue        : LongWord;
begin
try
begin
         //Mas:= VarArrayCreate([0, 2], varVariant);
       //Mas1:= VarArrayCreate([0, 10], varVariant);
   res:='';
  OKRightDlg1234567891011121314.listview1.Clear;
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  frmDomainInfo.memo1.Lines.Add('Загрузка данный сетевого интерфейса');
  OleInitialize(nil); ///// нахуй незнаю зачем
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = '+'"'+NetworkInterfaceID+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
       begin
                  with OKRightDlg1234567891011121314.listview1 do
                    begin
                    Items.Add.Caption:='Описание - '+(vartostr(FWbemObject.Description));
                    Items.Add.Caption:='Физический адрес - '+ (vartostr(FWbemObject.MACAddress)); // ok
                    Items.Add.Caption:='DHCP включен - '+(vartostr(FWbemObject.DHCPEnabled)); ///ok
                    mas:=((FWbemObject.IPAddress));
                    if VarType(mas) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                        res:=res+(string(mas[i]))+'/';
                      end;
                        Items.Add.Caption:='Адрес IPv4/IPv6 - '+res;
                    res:='';
                    Varclear(mas);
                    //mas:=null;
                    mas:=(FWbemObject.IPSubnet);
                    if VarType(mas) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                        res:=res+(string(mas[i]))+'/';
                      end;
                        Items.Add.Caption:='Маска подсети IPv4 - '+res;
                    res:='';
                    Varclear(mas);
                    //mas:=null;
                    mas:=(FWbemObject.DefaultIPGateway);
                    if VarType(mas) and VarTypeMask=varVariant then
                       begin
                        for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                        res:=res+(string(mas[i]))+'/';
                       end;
                       Items.Add.Caption:='Шлюз по умолчанию - '+res;
                    res:='';
                    Items.Add.Caption:='DHCP сервер IPv4 - '+(vartostr(FWbemObject.DHCPServer));
                    Varclear(mas);
                   // mas:=null;
                    mas:=(FWbemObject.DNSServerSearchOrder);
                    if VarType(mas) and VarTypeMask=varVariant then
                      begin
                        for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                        res:=res+(string(mas[i]))+'/';
                      end;
                    Items.Add.Caption:='DNS сервер IPv4 - '+res;
                    res:='';
                    Items.Add.Caption:='WINS сервер - '+(vartostr(FWbemObject.WINSPrimaryServer));
                  end;

end;
  Varclear(mas);
  oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add(' - Ошибка получения данных "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
end;


end.
