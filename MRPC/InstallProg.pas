unit InstallProg;

interface

uses
   System.Classes,System.Variants,ActiveX,
   ComObj,CommCtrl,Dialogs,SysUtils,Vcl.Controls,IdIcmpClient,ShlObj;

type
  InstallProgram = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
var
Myidicmpclient:TIdIcmpClient;

ThreadVar
ThreadPC      :string;


implementation
uses umain;

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



{ InstallProgram }
procedure StartInstall;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
   Resinstall,b,z :integer;
BEGIN
      Myidicmpclient:=TIdIcmpClient.Create;
      Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
      Myidicmpclient.PacketSize:=32;
      Myidicmpclient.Port:=0;
      Myidicmpclient.Protocol:=1;
      Myidicmpclient.ReceiveTimeout:=pingtimeout;
      ///////////////////////////////
if ping(ThreadPC) then
  begin
      try
          frmDomainInfo.memo1.Lines.Add('������� ������� ��������� ��������� '+ExtractFileName(NewinstallProgramm)+' �� '+ThreadPC+'.');
          finditem('������� ������� ��������� ��������� '+ExtractFileName(NewinstallProgramm),ThreadPC,14);
          OleInitialize(nil);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService  := FSWbemLocator.ConnectServer(ThreadPC, 'root\CIMV2', MyUser, MyPasswd,'','',128);
          FWbemObject   := FWMIService.Get('Win32_Product');
          Resinstall:=FWbemObject.install(NewinstallProgramm,KeyNewInstallProgram,InstallBool);
          if Resinstall=0 then
            Begin
            finditem('��������� ��������� ' +ExtractFileName(NewinstallProgramm)
            +' : '+SysErrorMessage(Resinstall),ThreadPC,1);
            End
          else
             Begin
             finditem('��� ��������� ��������� '
                  +ExtractFileName(NewinstallProgramm) +' �������� �������� : '
                  +SysErrorMessage(Resinstall),ThreadPC,2)
             End;
          frmDomainInfo.memo1.Lines.Add('��������� ��������� '
          +ExtractFileName(NewinstallProgramm)+' �� '
          +ThreadPC+' : '+SysErrorMessage(Resinstall));
          FWbemObject:=Unassigned;
          VariantClear(FWbemObject);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
          OleUnInitialize;
        except
          on E:Exception do
           Begin
           finditem('������ ��������� ��������� '
           +ExtractFileName(NewinstallProgramm) +' : '+E.Message,ThreadPC,2);
             frmDomainInfo.memo1.Lines.Add('������ ��������� ��������� '+ExtractFileName(NewinstallProgramm)+' �� '+ThreadPC+' : "'+E.Message+'"');
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             VariantClear(FWbemObject);
             VariantClear(FWMIService);
             VariantClear(FSWbemLocator);
             OleUnInitialize;
             exit;
           End;
       end;
  end;
END;


procedure InstallProgram.Execute;
begin
ThreadPC:=NewProgramMyPS;
StartInstall;
end;

end.
