unit SelectedPCShotDownThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  SelectedPCShotDown = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

var
MyIdIcmpClient: TIdIcmpClient;

implementation
uses umain;


{ SelectedPCShotDown }

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

function ShotDownResetClosesession(listpc:tstringlist):boolean;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResShutdown,i,z:integer;
  Operation     :string;
begin

Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;

    for I := 0 to listpc.Count-1 do
    if (listpc[i]<>'') and (ping(listpc[i])) then
      begin
          try
              frmDomainInfo.memo1.Lines.Add('���������� ������, ������, ������������ - '+listpc[i]);
              OleInitialize(nil);
              FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
              FWMIService   := FSWbemLocator.ConnectServer(listpc[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
              //FWMIService    := GetObject('winmgmts:{impersonationLevel=impersonate}!\\'+MyPS+'\root\CIMV2');
              FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
              oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
              if oEnum.Next(1, FWbemObject, iValue) = 0 then
                case myShutdown of
                 0:
                  begin
                    ResShutdown:=FWbemObject.Win32Shutdown(0);
                    frmDomainInfo.memo1.Lines.Add('���������� ������. '+SysErrorMessage(ResShutdown));
                    Operation:='"���������� ������."';
                    if ResShutdown<>0 then
                      begin
                      ResShutdown:=FWbemObject.Win32Shutdown(4);
                      frmDomainInfo.memo1.Lines.Add('�������������� ���������� ������. '+SysErrorMessage(ResShutdown));
                      Operation:='"�������������� ���������� ������."';
                      end;
                    frmDomainInfo.memo1.Lines.Add('-----------------------------');
                  end;
                  1:
                  begin
                   ResShutdown:=FWbemObject.Win32Shutdown(4);
                   frmDomainInfo.memo1.Lines.Add('�������������� ���������� ������. '+SysErrorMessage(ResShutdown));
                   frmDomainInfo.memo1.Lines.Add('-----------------------------');
                   Operation:='"�������������� ���������� ������."';
                  end;
                   2:
                   begin
                      ResShutdown:=FWbemObject.Win32Shutdown(2);
                     frmDomainInfo.memo1.Lines.Add('������������. '+SysErrorMessage(ResShutdown));
                     Operation:='"������������."';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(6);
                        frmDomainInfo.memo1.Lines.Add('�������������� ������������. '+SysErrorMessage(ResShutdown));
                        Operation:='"�������������� ������������."';
                        end;
                      frmDomainInfo.memo1.Lines.Add('-----------------------------');
                   end;
                   3:
                   begin
                   ResShutdown:=FWbemObject.Win32Shutdown(6);
                   frmDomainInfo.memo1.Lines.Add('�������������� ������������. '+SysErrorMessage(ResShutdown));
                   frmDomainInfo.memo1.Lines.Add('-----------------------------');
                   Operation:='"�������������� ������������."';
                   end;
                   4:
                   begin
                     ResShutdown:=FWbemObject.Win32Shutdown(1);
                     frmDomainInfo.memo1.Lines.Add('���������� ������. '+SysErrorMessage(ResShutdown));
                     Operation:='"���������� ������."';
                     if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(5);
                        frmDomainInfo.memo1.Lines.Add('�������������� ���������� ������. '+SysErrorMessage(ResShutdown));
                        Operation:='"�������������� ���������� ������."';
                        end;
                     frmDomainInfo.memo1.Lines.Add('-----------------------------');
                    end;
                    5:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(5);
                      frmDomainInfo.memo1.Lines.Add('�������������� ���������� ������. '+SysErrorMessage(ResShutdown));
                      Operation:='"�������������� ���������� ������."';
                      if ResShutdown<>0 then
                        begin
                        ResShutdown:=FWbemObject.Win32Shutdown(12);
                        frmDomainInfo.memo1.Lines.Add('��������� ���������� ������. '+SysErrorMessage(ResShutdown));
                        Operation:='"��������� ���������� ������."';
                        end;
                      frmDomainInfo.memo1.Lines.Add('-----------------------------');
                    end;
                    6:
                    begin
                      ResShutdown:=FWbemObject.Win32Shutdown(12);
                      frmDomainInfo.memo1.Lines.Add('��������� ���������� ������. '+SysErrorMessage(ResShutdown));
                      frmDomainInfo.memo1.Lines.Add('-----------------------------');
                       Operation:='"��������� ���������� ������"';
                    end;

                  end;
////////////////////////////////////////////////////////////////////////////////////////////////
          if ResShutdown=0 then
          begin
          finditem('�������� '+Operation+' : '+SysErrorMessage(ResShutdown),listpc[i],1 );
          end
        else
           begin
           finditem('��� ���������� �������� '+Operation+' �������� �������� :'+SysErrorMessage(ResShutdown),listpc[i],2);
           end;
///////////////////////////////////////////////////////////////////////////////
              FWbemObject:=Unassigned;
              VariantClear(FWbemObject);
              oEnum:=nil;
              VariantClear(FWbemObjectSet);
              VariantClear(FWMIService);
              VariantClear(FSWbemLocator);
              OleUnInitialize;
              except
            on E:Exception do
             begin
             finditem('��� ���������� �������� '+Operation+' �������� ��������. ������ "'+E.Message+'"',listpc[i],2);
             frmDomainInfo.memo1.Lines.Add(listpc[i]+'  ������ - "'+E.Message+'"');
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             OleUnInitialize;
             end;
            end;
       end; ///���������� �����
FreeAndNil(Myidicmpclient);
/// ������� ������ �����������
end;

procedure SelectedPCShotDown.Execute;
begin
ShotDownResetClosesession(SelectedPCShutDown);
SelectedPCShutDown.Free;
end;

end.
