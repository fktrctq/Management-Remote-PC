unit SelectedPCInstallDriverPrint;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  SelectedPCInstallDriver = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
 var
 MyIdIcmpClient: TIdIcmpClient;

implementation
uses umain,AddPrintDriverDialog;

{ SelectedPCInstallDriver }
function finditem(s,name:string;image:integer):boolean;
var
z:integer;
begin
 for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
    if frmDomainInfo.ListView8.Items[z].SubItems[0]=name then
    begin
    frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=image;
    frmDomainInfo.ListView8.Items[z].SubItems[1]:=s;
    break;
    end;
end;

 function ping(s:string):boolean;
var
z:integer;
begin
try
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  finditem('�������� �������� �������� �������',s,2);
  end
else
  begin
  result:=true; ///��������
  frmDomaininfo.Memo1.Lines.Add('IP ����� - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  �����='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'��'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;

   except
    begin
    result:=false;
    finditem('���� �� ��������',s,2);
    end;
   end;
end;

function InstDrivForPrint(ListPCForinstal:TstringList):boolean;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  objDriver    : OLEVariant;
  ParObjDriver : OLEVariant;
  MyError,i,z       : integer;

  function GetObject(const objectName: String): IDispatch;
  var
    chEaten: Integer;
    BindCtx: IBindCtx;//for access to a bind context
    Moniker: IMoniker;//Enables you to use a moniker object
    begin
    OleCheck(CreateBindCtx(0, bindCtx));
    OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));//Converts a string into a moniker that identifies the object named by the string
    OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));//Binds to the specified object
    end;
  begin
////////////////////////////////////
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
/////////////////////////////////////
GroupPC:=false;
for I := 0 to ListPCForinstal.Count-1 do
if (ListPCForinstal[i]<>'') and (ping(ListPCForinstal[i])) then
Begin
    try
    begin
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      frmDomainInfo.memo1.Lines.Add('��������� ��������  '+DriverName+' �� '+ListPCForinstal[i]);
      finditem('��������� ��������  '+DriverName,ListPCForinstal[i],16) ;
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer(ListPCForinstal[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
      FWMIService.Security_.impersonationlevel:=3;
      FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
      objDriver:=FWMIService.get('Win32_PrinterDriver');
      ParObjDriver:=FWMIService.get('Win32_PrinterDriver');
      if FilePath<>'' then ParObjDriver.Infname:=FilePath;
      ParObjDriver.Name:=DriverName;
      if DriverPath<>'' then ParObjDriver.DriverPath:=DriverPath;

      frmDomainInfo.memo1.Lines.Add('���� - '+FilePath);
      frmDomainInfo.memo1.Lines.Add('������� - '+DriverName);
      frmDomainInfo.memo1.Lines.Add('DLL - '+DriverPath);

      MyError:=objDriver.AddPrinterDriver(ParObjDriver);
      if MyError=0 then
          begin
          finditem('��������� ��������  '+DriverName+' : '+SysErrorMessage(MyError),ListPCForinstal[i],1);
          end
        else
           begin
           finditem('��� ��������� �������� '+DriverName+' �������� ��������. ������ : '+SysErrorMessage(MyError),ListPCForinstal[i],2);
           end;
      frmDomainInfo.memo1.Lines.Add('�������� ��������� �������� '+DriverName+'  �� '+ListPCForinstal[i]+' : '+SysErrorMessage(MyError));
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      VariantClear(objDriver);
      VariantClear(ParObjDriver);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
    end;

    except
    on E:Exception do
    begin
    finditem('��� ��������� �������� '+DriverName+'  �������� ��������. ������ : '+E.Message,ListPCForinstal[i],2);
    frmDomainInfo.memo1.Lines.Add('������ ��������� �������� �� '+ListPCForinstal[i]+' : "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    OleUnInitialize;
    end;
    end;

End;
if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
end;



procedure SelectedPCInstallDriver.Execute;
begin
InstDrivForPrint(SelectedPCInstalDrv);
freeandnil(SelectedPCInstalDrv);
end;

end.
