unit Unit5;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  NewProcess = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;





implementation
uses umain,FormNewProcess;
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
MyIdIcmpClient: TIdIcmpClient;
begin
try
result:=false;
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
Myidicmpclient.host:=s;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
  begin
  result:=false;
  finditem('Превышен интервал ожидания запроса',s,2);
  end
else
  begin
  result:=true; ///доступен
  frmDomaininfo.Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
  end;

   except
    begin
    result:=false;
    finditem('Узел не доступен',s,2);
    end;
   end;
end;



{ NewProcess }
function MyProcess(ThreadPC,user,pass:string):boolean;
var
MyError:integer;
 FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  objProcess    : OLEVariant;
  objConfig     : OLEVariant;
  ProcessID,z   : Integer;

  const
  wbemFlagForwardOnly = $00000020;
  HIDDEN_WINDOW       = 1;
begin
//if GroupPC=true then
//if ping(ThreadPC) then    ///// если компьютер доступен то запусаем процесс
try
        begin
          OleInitialize(nil);
          frmDomainInfo.memo1.Lines.Add('Запуск процесса на - '+ThreadPC);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService   := FSWbemLocator.ConnectServer(ThreadPC, 'root\CIMV2', User, Pass,'','',128);
          FWMIService.security_.AuthenticationLevel:=6;
          FWMIService.security_.ImpersonationLevel:=3;
          //FWMIService.security_.Privileges.AddAsString('SeEnableDelegationPrivilege');
          FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
          objConfig     := FWbemObject.SpawnInstance_;
          objConfig.ShowWindow := HIDDEN_WINDOW;
          objProcess    := FWMIService.Get('Win32_Process');
          MyError:=objProcess.Create(MyCommandLine, null, objConfig, (ProcessID));
           if MyError=0 then
          begin
          finditem('Запуск процесса '+MyCommandLine+' : '+': '+SysErrorMessage(MyError),ThreadPC,1);
          end
        else
           begin
           finditem('При запуске процесса '+MyCommandLine+' возникли проблемы : '+SysErrorMessage(MyError),ThreadPC,2);
           end;

          frmDomainInfo.memo1.Lines.Add('Запуск процесса '+MyCommandLine+' на '+ThreadPC+' : '+SysErrorMessage(MyError));
          frmDomainInfo.memo1.Lines.Add('---------------------------');
          FWbemObject:=Unassigned;
          objProcess:=Unassigned;
          VariantClear(FWbemObject);
          VariantClear(objConfig);
          VariantClear(objProcess);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
          OleUnInitialize;
        end;
     except
on E:Exception do
         begin
         finditem('При запуске процесса '+MyCommandLine+' возникли проблемы : "'+E.Message+'"',ThreadPC,2);
         frmDomainInfo.memo1.Lines.Add('При запуске процесса '+MyCommandLine+' на '+ThreadPC+' возникли проблемы. - "'+E.Message+'"');
         frmDomainInfo.memo1.Lines.Add('---------------------------');
         OleUnInitialize;
         end;

      end;
end;

procedure NewProcess.Execute;
begin
MyProcess(NewProcMyPS,frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);

end;

end.
