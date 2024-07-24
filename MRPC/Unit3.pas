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



Function killSelectedProcessForGroupPC(NewPCinGroup:string):boolean;
var ///////////////////////////////////// завершение процесса на группе машин в разных потоках
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
Begin  ////////////////////////////////////////////// завершение процесса на группе машин
      try
       Begin
        YesProc:=false;
        OleInitialize(nil); ///// нахуй незнаю зачем
        FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
        FWMIService  := FSWbemLocator.ConnectServer(NewPCinGroup, 'root\CIMV2', MyUser, MyPasswd,'','',128);
        FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"%'+GroupselectProc+'%"','WQL',wbemFlagForwardOnly);
        oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
        while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
         begin
          ErrorProcKill:=FWbemObject.terminate();
          frmDomainInfo.memo1.Lines.Add('Операция Завершение процесса '+GroupselectProc+' на '+NewPCinGroup+' : '+SysErrorMessage(ErrorProcKill));
          frmDomainInfo.memo1.Lines.Add('---------------------------');
          FWbemObject:=Unassigned;

          finditem('Операция Завершение процесса '+
          GroupselectProc+' : '+SysErrorMessage(ErrorProcKill),NewPCinGroup,1);

          YesProc:=true;
         end;

        if YesProc=false then
        begin
        finditem('Процесс '+GroupselectProc+' не найден',NewPCinGroup,6);
        frmDomainInfo.memo1.Lines.Add('Процесс '+GroupselectProc+' на '+NewPCinGroup+' не найден.');
        end;

        End;

  except
  on E:Exception do
  if YesProc=false then
  begin
  finditem('При завершении процесса '+GroupselectProc+' произошла ошибка. - "'+E.Message+'"',NewPCinGroup,2);
  frmDomainInfo.memo1.Lines.Add('При завершении процесса '+GroupselectProc+' на '+NewPCinGroup+' произошла ошибка. - "'+E.Message+'"');
  frmDomainInfo.memo1.Lines.Add('---------------------------');
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  end;
  end;

End; /////////////////////////////////////////////// завершение процесса на на группе машин  разных потоках
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
var    ///////////////////////// завершение процессов на одной машине
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
               frmDomainInfo.memo1.Lines.Add('завершение процесса '+ListProcess.names[y]+' - ID:'+ListProcess.ValueFromIndex[y]+' :'+SysErrorMessage(ErrorProcKill));
               frmDomainInfo.memo1.Lines.Add('---------------------------');
               FWbemObject:=Unassigned;
               YesProc:=true;

               finditem('завершение процесса '+ListProcess.names[y]+' - ID:'
               +ListProcess.ValueFromIndex[y]+' :'+SysErrorMessage(ErrorProcKill),namePC,1)
               end;

             if YesProc=false then
              begin
               frmDomainInfo.memo1.Lines.Add('Процесс '+ListProcess.names[y]+' c ID: '+ListProcess.ValueFromIndex[y]+' не найден.');
               finditem('Процесс '+ListProcess.names[y]+' c ID: '
               +ListProcess.ValueFromIndex[y]+' не неайден',namePC,6)
              end;

           except
           on E:Exception do
            begin
            frmDomainInfo.memo1.Lines.Add('Ошибка завершения процесса '+ListProcess.names[y]+' c ID: '+ListProcess.ValueFromIndex[y]+' : "'+E.Message+'"');
            finditem('Ошибка завершения процесса '+ListProcess.names[y]
            +' c ID: '+ListProcess.ValueFromIndex[y]+' : "'+E.Message+'"',namePC,2)
            end;
           end;

        End;
except
  on E:Exception do
    begin
    finditem('При завершении процесса произошла ошибка. - "'+E.Message+'"',namePC,2);
    frmDomainInfo.memo1.Lines.Add('Общая ошибка завершения процессов  "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    end;
  end;

VariantClear(FWbemObject);
oEnum:=nil;
VariantClear(FWbemObjectSet);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
ListProcess.Free; /// очищаем список процессов для завершения
END;



function killSelectedProcess(ListPCKill:TstringList):boolean;
var ///////////////////////////////////// завершение процесса на выбранных машинах в одном потоке
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
                 OleInitialize(nil); ///// нахуй незнаю зачем
                 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                 FWMIService  := FSWbemLocator.ConnectServer(ListPCKill[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
                 FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Caption LIKE'+'"%'+GroupselectProc+'%"','WQL',wbemFlagForwardOnly);
                 oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;

                   while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                   begin
                   ErrorProcKill:=FWbemObject.terminate();
                   frmDomainInfo.memo1.Lines.Add('завершение процесса на '+ListPCKill[i]+' :'+SysErrorMessage(ErrorProcKill));
                   frmDomainInfo.memo1.Lines.Add('---------------------------');
                   FWbemObject:=Unassigned;
                   YesProc:=true;
                   end;

                   if (YesProc=true) and (ErrorProcKill=0) then
                   begin
                   finditem('Завершение процесса '+GroupselectProc+' прошло успешно',ListPCKill[i],1);
                   end;

                   if (YesProc=true)and (ErrorProcKill<>0) then
                   begin
                   finditem('При завершении процесса '+GroupselectProc+' возникли проблемы.- '+SysErrorMessage(ErrorProcKill),ListPCKill[i],2);
                   end;

                   if (YesProc=false)then
                   finditem('Процесс '+GroupselectProc+' не найден',ListPCKill[i],6);

                End;

          except
            on E:Exception do
              begin
              finditem('При завершении процесса произошла ошибка. - "'+E.Message+'"',ListPCKill[i],2);
              frmDomainInfo.memo1.Lines.Add('Ошибка  "'+E.Message+'"');
              frmDomainInfo.memo1.Lines.Add('---------------------------');
              VariantClear(FWbemObject);
              oEnum:=nil;
              VariantClear(FWbemObjectSet);
              VariantClear(FWMIService);
              VariantClear(FSWbemLocator);
              end;
            end;
      END;  /////// окончание цикла

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
  begin           /// завершение списка процессов на одной машине
  KillSelectedProcessNoGroup(NewProcMyPS);
  end;

if GroupPC then  /// завершение одного процесса на группе машин
  begin
  if not MultiThread then /// признак завершения процесса в одном потоке
  begin
   killSelectedProcess(SelectedPCkillProc);   ///// на группу компов в одном потоке
  end;
  if  MultiThread then   /// признак завершения процесса в разных потоках
    killSelectedProcessForGroupPC(NewProcMyPS);  //// на группе компов в разных потоках - вызывается из формы SelectedPCKillProcess
  end;
end;

end.
