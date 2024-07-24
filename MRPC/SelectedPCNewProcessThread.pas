unit SelectedPCNewProcessThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  SelectedPCNewProcess = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
var
MyIdIcmpClient: TIdIcmpClient;


implementation
uses umain;

{ SelectedPCNewProcess }
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
        for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
          if frmDomainInfo.ListView8.Items[z].SubItems[0]=s then
          begin
          frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=2;
          frmDomainInfo.ListView8.Items[z].SubItems[1]:='Превышен интервал ожидания запроса';
          break;
          end;
        frmDomainInfo.memo1.Lines.Add('---------------------------');
        frmDomaininfo.memo1.Lines.Add(s+' - Превышен интервал ожидания запроса');
        result:=false;
      end
    else
      begin
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      frmDomaininfo.Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
      result:=true;
      end;
  except
  on E: Exception do
    begin
      for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
        if frmDomainInfo.ListView8.Items[z].SubItems[0]=s then
        begin
        frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=2;
        frmDomainInfo.ListView8.Items[z].SubItems[1]:='Узел не доступен';
        break;
        end;
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      frmDomaininfo.memo1.Lines.add(s+' - Узел не доступен.');
      result:=false;
    end;
  end;

end;

 procedure SelectedPCNewProcess.Execute;
 begin

 end;


{procedure SelectedPCNewProcess.Execute;
var
MyError:integer;
 FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  objProcess    : OLEVariant;
  objConfig     : OLEVariant;
  ProcessID,i,z   : Integer;
  const
  wbemFlagForwardOnly = $00000020;
  HIDDEN_WINDOW       = 1;
begin
///////////////////////////
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
///////////////////////////////
for I := 0 to SelectedPCNewProc.Count-1 do
if (SelectedPCNewProc[i]<>'') and (ping(SelectedPCNewProc[i])) then
  begin
    try
        begin
          OleInitialize(nil);
          frmDomainInfo.memo1.Lines.Add('Запуск процесса на - '+SelectedPCNewProc[i]);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService   := FSWbemLocator.ConnectServer(SelectedPCNewProc[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
          FWbemObject   := FWMIService.Get('Win32_ProcessStartup');
          objConfig     := FWbemObject.SpawnInstance_;
          objConfig.ShowWindow := HIDDEN_WINDOW;
          objProcess    := FWMIService.Get('Win32_Process');
          MyError:=objProcess.Create(MyCommandLine, null, objConfig, (ProcessID));
          FWbemObject:=Unassigned;
          objProcess:=Unassigned;
          if MyError=0 then
          begin
            for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
              if frmDomainInfo.ListView8.Items[z].SubItems[0]=SelectedPCNewProc[i] then
               begin
                frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=1;
                frmDomainInfo.ListView8.Items[z].SubItems[1]:='Запуск процесса '+MyCommandLine+' прошел успешно, ID - '+inttostr(ProcessID)
                +': '+SysErrorMessage(MyError);
                break;
                end;
          end
        else
           begin
           for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
              if frmDomainInfo.ListView8.Items[z].SubItems[0]=SelectedPCNewProc[i] then
                begin
                frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=2;
                frmDomainInfo.ListView8.Items[z].SubItems[1]:='При запуске процесса '+MyCommandLine+' возникли проблемы : '+SysErrorMessage(MyError);
                break;
                end;
           end;
          frmDomainInfo.memo1.Lines.Add('Запуск процесса '+MyCommandLine+' на '+SelectedPCNewProc[i]+' : '+SysErrorMessage(MyError));
          frmDomainInfo.memo1.Lines.Add('---------------------------');
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
               for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
                if frmDomainInfo.ListView8.Items[z].SubItems[0]=SelectedPCNewProc[i] then
                begin
                frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=2;
                frmDomainInfo.ListView8.Items[z].SubItems[1]:='При запуске процесса '+MyCommandLine+' на '+SelectedPCNewProc[i]+' возникли проблемы : "'+E.Message+'"';
                break;
                end;
               frmDomainInfo.memo1.Lines.Add('При запуске процесса '+MyCommandLine+' на '+SelectedPCNewProc[i]+' возникли проблемы. - "'+E.Message+'"');
               frmDomainInfo.memo1.Lines.Add('---------------------------');
               OleUnInitialize;
             end;
    end;
 end;
FreeAndNil(Myidicmpclient);
SelectedPCNewProc.Clear;
GroupPC:=false;
end;}


end.
