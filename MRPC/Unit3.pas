unit Unit3;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,IdIcmpClient;

type
  KillProcess = class(TThread)



  private


  protected
    procedure Execute; override;
  end;

 ThreadVar
 ListPCKill:tstringList;
 ThreadPCKillProc:string;

implementation
uses umain,SelectedPCKillProcess;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  Myidicmpclient:TIdIcmpClient;
  iValue        : LongWord;

{ KillProcess }

{ping}

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



Function killSelectedProcessForGroupPC(NewPCinGroup:string):boolean;
var ///////////////////////////////////// ���������� �������� �� ������ ����� � ������ �������
YesProc:boolean;
i,z,ErrorProcKill:integer;
BEGIN
      Myidicmpclient:=TIdIcmpClient.Create;
      Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
      Myidicmpclient.PacketSize:=32;
      Myidicmpclient.Port:=0;
      Myidicmpclient.Protocol:=1;
      Myidicmpclient.ReceiveTimeout:=pingtimeout;
      ///////////////////////////////
if ping(NewPCinGroup) then
Begin  ////////////////////////////////////////////// ���������� �������� �� ������ �����
      try
       Begin
        YesProc:=false;
        OleInitialize(nil); ///// ����� ������ �����
        FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
        FWMIService  := FSWbemLocator.ConnectServer(NewPCinGroup, 'root\CIMV2', MyUser, MyPasswd,'','',128);
        FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"%'+GroupselectProc+'%"','WQL',wbemFlagForwardOnly);
        oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
        while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
         begin
          ErrorProcKill:=FWbemObject.terminate();
          frmDomainInfo.memo1.Lines.Add('�������� ���������� �������� '+GroupselectProc+' �� '+NewPCinGroup+' : '+SysErrorMessage(ErrorProcKill));
          frmDomainInfo.memo1.Lines.Add('---------------------------');
          FWbemObject:=Unassigned;

          finditem('�������� ���������� �������� '+
          GroupselectProc+' : '+SysErrorMessage(ErrorProcKill),NewPCinGroup,1);

          YesProc:=true;
         end;

        if YesProc=false then
        begin
        finditem('������� '+GroupselectProc+' �� ������',NewPCinGroup,6);
        frmDomainInfo.memo1.Lines.Add('������� '+GroupselectProc+' �� '+NewPCinGroup+' �� ������.');
        end;

        End;

  except
  on E:Exception do
  if YesProc=false then
  begin
  finditem('��� ���������� �������� '+GroupselectProc+' ��������� ������. - "'+E.Message+'"',NewPCinGroup,2);
  frmDomainInfo.memo1.Lines.Add('��� ���������� �������� '+GroupselectProc+' �� '+NewPCinGroup+' ��������� ������. - "'+E.Message+'"');
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  end;
  end;

End; /////////////////////////////////////////////// ���������� �������� �� �� ������ �����  ������ �������
VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
Myidicmpclient.Free;
OleUnInitialize;
END;



Function KillSelectedProcessNoGroup(namePC:string):boolean;
var    ///////////////////////// ���������� ��������� �� ����� ������
YesProc:boolean;
i,z,y,b:integer;
ListProcess:TstringList;
ErrorProcKill:integer;
BEGIN
try
           OleInitialize(nil);
           ListProcess:=TStringList.Create;
           for b := 0 to SelectProc.Count-1 do
           ListProcess.Add(SelectProc[b]);
           SelectProc.Free;
           FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
           FWMIService  := FSWbemLocator.ConnectServer(namePC, 'root\CIMV2', MyUser, MyPasswd,'','',128);
           //
           for y := 0 to ListProcess.Count-1 do
           Begin
           try
             YesProc:=false;
             if Multithread=false then FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE ProcessId='+ListProcess.ValueFromIndex[y],'WQL',wbemFlagForwardOnly);
             //if Multithread=true then FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"%'+SelectProc+'%"','WQL',wbemFlagForwardOnly);
             oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
             while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
               begin
               ErrorProcKill:=FWbemObject.terminate();
               frmDomainInfo.memo1.Lines.Add('���������� �������� '+ListProcess.names[y]+' - ID:'+ListProcess.ValueFromIndex[y]+' :'+SysErrorMessage(ErrorProcKill));
               frmDomainInfo.memo1.Lines.Add('---------------------------');
               FWbemObject:=Unassigned;
               YesProc:=true;

               finditem('���������� �������� '+ListProcess.names[y]+' - ID:'
               +ListProcess.ValueFromIndex[y]+' :'+SysErrorMessage(ErrorProcKill),namePC,1)
               end;

             if YesProc=false then
              begin
               frmDomainInfo.memo1.Lines.Add('������� '+ListProcess.names[y]+' c ID: '+ListProcess.ValueFromIndex[y]+' �� ������.');
               finditem('������� '+ListProcess.names[y]+' c ID: '
               +ListProcess.ValueFromIndex[y]+' �� �������',namePC,6)
              end;

           except
           on E:Exception do
            begin
            frmDomainInfo.memo1.Lines.Add('������ ���������� �������� '+ListProcess.names[y]+' c ID: '+ListProcess.ValueFromIndex[y]+' : "'+E.Message+'"');
            finditem('������ ���������� �������� '+ListProcess.names[y]
            +' c ID: '+ListProcess.ValueFromIndex[y]+' : "'+E.Message+'"',namePC,2)
            end;
           end;

        End;
except
  on E:Exception do
    begin
    finditem('��� ���������� �������� ��������� ������. - "'+E.Message+'"',namePC,2);
    frmDomainInfo.memo1.Lines.Add('����� ������ ���������� ���������  "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    end;
  end;

VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
ListProcess.Free; /// ������� ������ ��������� ��� ����������
END;



function killSelectedProcess(ListPCKill:TstringList):boolean;
var ///////////////////////////////////// ���������� �������� �� ��������� ������� � ����� ������
YesProc:boolean;
i,z,ErrorProcKill:integer;
BEGIN
      Myidicmpclient:=TIdIcmpClient.Create;
      Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
      Myidicmpclient.PacketSize:=32;
      Myidicmpclient.Port:=0;
      Myidicmpclient.Protocol:=1;
      Myidicmpclient.ReceiveTimeout:=pingtimeout;

      for I := 0 to ListPCKill.Count-1 do
      if (ListPCKill[i]<>'') and (ping(ListPCKill[i])) then
          BEGIN  //////////////////////////////////////////////
               try
                Begin
                 YesProc:=false;
                 OleInitialize(nil); ///// ����� ������ �����
                 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                 FWMIService  := FSWbemLocator.ConnectServer(ListPCKill[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
                 FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"%'+GroupselectProc+'%"','WQL',wbemFlagForwardOnly);
                 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;

                   while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                   begin
                   ErrorProcKill:=FWbemObject.terminate();
                   frmDomainInfo.memo1.Lines.Add('���������� �������� �� '+ListPCKill[i]+' :'+SysErrorMessage(ErrorProcKill));
                   frmDomainInfo.memo1.Lines.Add('---------------------------');
                   FWbemObject:=Unassigned;
                   YesProc:=true;
                   end;

                   if (YesProc=true) and (ErrorProcKill=0) then
                   begin
                   finditem('���������� �������� '+GroupselectProc+' ������ �������',ListPCKill[i],1);
                   end;

                   if (YesProc=true)and (ErrorProcKill<>0) then
                   begin
                   finditem('��� ���������� �������� '+GroupselectProc+' �������� ��������.- '+SysErrorMessage(ErrorProcKill),ListPCKill[i],2);
                   end;

                   if (YesProc=false)then
                   finditem('������� '+GroupselectProc+' �� ������',ListPCKill[i],6);

                End;

          except
            on E:Exception do
              begin
              finditem('��� ���������� �������� ��������� ������. - "'+E.Message+'"',ListPCKill[i],2);
              frmDomainInfo.memo1.Lines.Add('������  "'+E.Message+'"');
              frmDomainInfo.memo1.Lines.Add('---------------------------');
              VariantClear(FWbemObject);
              oEnum:=nil;
              VariantClear(FWbemObjectSet);
              VariantClear(FWMIService);
              VariantClear(FSWbemLocator);
              end;
            end;
      END;  /////// ��������� �����

           VariantClear(FWbemObject);
           oEnum:=nil;
           VariantClear(FWbemObjectSet);
           VariantClear(FWMIService);
           VariantClear(FSWbemLocator);
           OleUnInitialize;
           ListPCKill.Free;
           Myidicmpclient.Free;
           OleUnInitialize;
END;

procedure KillProcess.Execute;
var
i:integer;
begin

if GroupPC=false then
  begin           /// ���������� ������ ��������� �� ����� ������
  KillSelectedProcessNoGroup(NewProcMyPS);
  end;

if GroupPC then  /// ���������� ������ �������� �� ������ �����
  begin
  if not MultiThread then /// ������� ���������� �������� � ����� ������
  begin
   killSelectedProcess(SelectedPCkillProc);   ///// �� ������ ������ � ����� ������
  end;
  if  MultiThread then   /// ������� ���������� �������� � ������ �������
    killSelectedProcessForGroupPC(NewProcMyPS);  //// �� ������ ������ � ������ ������� - ���������� �� ����� SelectedPCKillProcess
  end;
end;

end.
