unit SelectedPCDeleteProgramThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,
  IdIcmpClient;

type
  SelectedPCDeleteProgram = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
 var
  MyIdIcmpClient: TIdIcmpClient;
  CurrentPCDiffThread:string; /// для имени компа при запуске удаления в разных потоках
implementation
uses umain,SelectedPCDeleteProgramDialog;


{ SelectedPCDeleteProgram }
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

 function DeleteProgramInDifferentTread(NamePC:string):boolean;    ///// удаление программы в разных потоках
var
  FSWbemLocator   : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResDelProgram,i,z: integer;
  YesProg:boolean;


begin
     ////////////////////////////////////////////
      Myidicmpclient:=TIdIcmpClient.Create;
      Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
      Myidicmpclient.PacketSize:=32;
      Myidicmpclient.Port:=0;
      Myidicmpclient.Protocol:=1;
      Myidicmpclient.ReceiveTimeout:=pingtimeout;
      ///////////////////////////////

    if ping(NamePC) then
      Begin
        try
            begin
                OleInitialize(nil);
                YesProg:=false;
                frmDomainInfo.memo1.Lines.Add('---------------------------');
                frmDomainInfo.memo1.Lines.Add(NamePC+' - Запущен процесс удаления программы :'+GroupPCSelectProg);
                finditem('Запущен процесс удаления программы :'+GroupPCSelectProg,NamePC,13) ;
                FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd,'','',128);
                if equallyName=false then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+GroupPCSelectProg+'%"','WQL',wbemFlagForwardOnly);
                if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+GroupPCSelectProg+'"','WQL',wbemFlagForwardOnly);
                oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
                 while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                  begin
                  ResDelProgram:=FWbemObject.Uninstall(); //// Удалениме программы     '+'"'+GroupPCSelectProg+'"'
                  FWbemObject:=Unassigned;
                  YesProg:=true;
                  end;
                  if YesProg=true then
                      begin
                        if ResDelProgram=0 then
                          begin
                          finditem('Удаление программы '+GroupPCSelectProg+' завершилось : '+SysErrorMessage(ResDelProgram),NamePC,1);
                          frmDomainInfo.memo1.Lines.Add('Удаление программы '+GroupPCSelectProg+' на '+NamePC+' завершилось : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end
                        else
                          begin
                          finditem('При удалении программы '+GroupPCSelectProg+' возникли проблемы : '+SysErrorMessage(ResDelProgram),NamePC,2);
                          frmDomainInfo.memo1.Lines.Add('При удалении программы '+GroupPCSelectProg+' на '+NamePC+' возникли проблемы : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end;
                       end;
                    if YesProg=false then
                      begin
                      finditem('Программа '+GroupPCSelectProg+' не найдена',NamePC,6);
                      frmDomainInfo.memo1.Lines.Add('Программа '+GroupPCSelectProg+' на '+NamePC+' не найдена.');
                      frmDomainInfo.memo1.Lines.Add('---------------------------');
                      end;
             VariantClear(FWbemObject);
            oEnum:=nil;
            VariantClear(FWbemObjectSet);
            VariantClear(FWMIService);
            VariantClear(FSWbemLocator);
            OleUnInitialize;
            end;
            except
              on E:Exception do
               begin
               finditem('При удалении программы '+GroupPCSelectProg+' произошла ошибка. - "'+E.Message+'"',NamePC,2);
               frmDomainInfo.memo1.Lines.Add('Ошибка удаления программы '+GroupPCSelectProg+' на '+NamePC+': "'+E.Message+'"');
               frmDomainInfo.memo1.Lines.Add('---------------------------');
               VariantClear(FWbemObject);
               oEnum:=nil;
               VariantClear(FWbemObjectSet);
               VariantClear(FWMIService);
               VariantClear(FSWbemLocator);
               OleUnInitialize;
               end;
            end;
      End;

if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
end;



function SelectedPCDeleteProgramOneThread(MyPCstringList:TstringList):boolean;    ///// удаление программы в одном потоке
var
  FSWbemLocator   : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResDelProgram,i,z: integer;
  YesProg:boolean;


begin
     ////////////////////////////////////////////
      Myidicmpclient:=TIdIcmpClient.Create;
      Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
      Myidicmpclient.PacketSize:=32;
      Myidicmpclient.Port:=0;
      Myidicmpclient.Protocol:=1;
      Myidicmpclient.ReceiveTimeout:=pingtimeout;
      ///////////////////////////////
 for I := 0 to myPCStringList.Count-1 do
    if ping(myPCStringList[i]) then
      Begin
        try
            begin
                OleInitialize(nil);
                YesProg:=false;
                frmDomainInfo.memo1.Lines.Add('---------------------------');
                frmDomainInfo.memo1.Lines.Add(myPCStringList[i]+' - Запущен процесс удаления программы :'+GroupPCSelectProg);
                finditem('Запущен процесс удаления программы :'+GroupPCSelectProg,myPCStringList[i],13) ;
                FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                FWMIService  := FSWbemLocator.ConnectServer(myPCStringList[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
                if equallyName=false then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+GroupPCSelectProg+'%"','WQL',wbemFlagForwardOnly);
                if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+GroupPCSelectProg+'"','WQL',wbemFlagForwardOnly);
                oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
                 while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                  begin
                  ResDelProgram:=FWbemObject.Uninstall(); //// Удалениме программы     '+'"'+GroupPCSelectProg+'"'
                  FWbemObject:=Unassigned;
                  YesProg:=true;
                  end;
                  if YesProg=true then
                      begin
                        if ResDelProgram=0 then
                          begin
                          finditem('Удаление программы '+GroupPCSelectProg+' завершилось : '+SysErrorMessage(ResDelProgram),myPCStringList[i],1);
                          frmDomainInfo.memo1.Lines.Add('Удаление программы '+GroupPCSelectProg+' на '+myPCStringList[i]+' завершилось : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end
                        else
                          begin
                          finditem('При удалении программы '+GroupPCSelectProg+' возникли проблемы : '+SysErrorMessage(ResDelProgram),myPCStringList[i],2);
                          frmDomainInfo.memo1.Lines.Add('При удалении программы '+GroupPCSelectProg+' на '+myPCStringList[i]+' возникли проблемы : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end;
                       end;
                    if YesProg=false then
                      begin
                      finditem('Программа '+GroupPCSelectProg+' не найдена',myPCStringList[i],6);
                      frmDomainInfo.memo1.Lines.Add('Программа '+GroupPCSelectProg+' на '+myPCStringList[i]+' не найдена.');
                      frmDomainInfo.memo1.Lines.Add('---------------------------');
                      end;
             VariantClear(FWbemObject);
            oEnum:=nil;
            VariantClear(FWbemObjectSet);
            VariantClear(FWMIService);
            VariantClear(FSWbemLocator);
            OleUnInitialize;
            end;
            except
              on E:Exception do
               begin
               finditem('При удалении программы '+GroupPCSelectProg+' произошла ошибка. - "'+E.Message+'"',myPCStringList[i],2);
               frmDomainInfo.memo1.Lines.Add('Ошибка удаления программы '+GroupPCSelectProg+' на '+myPCStringList[i]+': "'+E.Message+'"');
               frmDomainInfo.memo1.Lines.Add('---------------------------');
               VariantClear(FWbemObject);
               oEnum:=nil;
               VariantClear(FWbemObjectSet);
               VariantClear(FWMIService);
               VariantClear(FSWbemLocator);
               OleUnInitialize;
               end;
            end;
      End;

if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
if Assigned(myPCStringList) then freeandnil(myPCStringList);
end;


procedure SelectedPCDeleteProgram.Execute;
begin
if GroupPCDeleteProg then SelectedPCDeleteProgramOneThread(SelectedPCDeleteProg)
else DeleteProgramInDifferentTread(CurrentPCDiffThread);
end;


end.
