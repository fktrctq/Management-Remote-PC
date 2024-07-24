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
  CurrentPCDiffThread:string; /// ��� ����� ����� ��� ������� �������� � ������ �������
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

 function DeleteProgramInDifferentTread(NamePC:string):boolean;    ///// �������� ��������� � ������ �������
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
                frmDomainInfo.memo1.Lines.Add(NamePC+' - ������� ������� �������� ��������� :'+GroupPCSelectProg);
                finditem('������� ������� �������� ��������� :'+GroupPCSelectProg,NamePC,13) ;
                FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                FWMIService  := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', MyUser, MyPasswd,'','',128);
                if equallyName=false then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+GroupPCSelectProg+'%"','WQL',wbemFlagForwardOnly);
                if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+GroupPCSelectProg+'"','WQL',wbemFlagForwardOnly);
                oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
                 while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                  begin
                  ResDelProgram:=FWbemObject.Uninstall(); //// ��������� ���������     '+'"'+GroupPCSelectProg+'"'
                  FWbemObject:=Unassigned;
                  YesProg:=true;
                  end;
                  if YesProg=true then
                      begin
                        if ResDelProgram=0 then
                          begin
                          finditem('�������� ��������� '+GroupPCSelectProg+' ����������� : '+SysErrorMessage(ResDelProgram),NamePC,1);
                          frmDomainInfo.memo1.Lines.Add('�������� ��������� '+GroupPCSelectProg+' �� '+NamePC+' ����������� : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end
                        else
                          begin
                          finditem('��� �������� ��������� '+GroupPCSelectProg+' �������� �������� : '+SysErrorMessage(ResDelProgram),NamePC,2);
                          frmDomainInfo.memo1.Lines.Add('��� �������� ��������� '+GroupPCSelectProg+' �� '+NamePC+' �������� �������� : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end;
                       end;
                    if YesProg=false then
                      begin
                      finditem('��������� '+GroupPCSelectProg+' �� �������',NamePC,6);
                      frmDomainInfo.memo1.Lines.Add('��������� '+GroupPCSelectProg+' �� '+NamePC+' �� �������.');
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
               finditem('��� �������� ��������� '+GroupPCSelectProg+' ��������� ������. - "'+E.Message+'"',NamePC,2);
               frmDomainInfo.memo1.Lines.Add('������ �������� ��������� '+GroupPCSelectProg+' �� '+NamePC+': "'+E.Message+'"');
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



function SelectedPCDeleteProgramOneThread(MyPCstringList:TstringList):boolean;    ///// �������� ��������� � ����� ������
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
                frmDomainInfo.memo1.Lines.Add(myPCStringList[i]+' - ������� ������� �������� ��������� :'+GroupPCSelectProg);
                finditem('������� ������� �������� ��������� :'+GroupPCSelectProg,myPCStringList[i],13) ;
                FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
                FWMIService  := FSWbemLocator.ConnectServer(myPCStringList[i], 'root\CIMV2', MyUser, MyPasswd,'','',128);
                if equallyName=false then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption LIKE '+'"%'+GroupPCSelectProg+'%"','WQL',wbemFlagForwardOnly);
                if equallyName then FWbemObjectSet:= FWMIService.ExecQuery('SELECT Caption FROM Win32_Product WHERE Caption ='+'"'+GroupPCSelectProg+'"','WQL',wbemFlagForwardOnly);
                oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
                 while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
                  begin
                  ResDelProgram:=FWbemObject.Uninstall(); //// ��������� ���������     '+'"'+GroupPCSelectProg+'"'
                  FWbemObject:=Unassigned;
                  YesProg:=true;
                  end;
                  if YesProg=true then
                      begin
                        if ResDelProgram=0 then
                          begin
                          finditem('�������� ��������� '+GroupPCSelectProg+' ����������� : '+SysErrorMessage(ResDelProgram),myPCStringList[i],1);
                          frmDomainInfo.memo1.Lines.Add('�������� ��������� '+GroupPCSelectProg+' �� '+myPCStringList[i]+' ����������� : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end
                        else
                          begin
                          finditem('��� �������� ��������� '+GroupPCSelectProg+' �������� �������� : '+SysErrorMessage(ResDelProgram),myPCStringList[i],2);
                          frmDomainInfo.memo1.Lines.Add('��� �������� ��������� '+GroupPCSelectProg+' �� '+myPCStringList[i]+' �������� �������� : '+SysErrorMessage(ResDelProgram));
                          frmDomainInfo.memo1.Lines.Add('---------------------------');
                          end;
                       end;
                    if YesProg=false then
                      begin
                      finditem('��������� '+GroupPCSelectProg+' �� �������',myPCStringList[i],6);
                      frmDomainInfo.memo1.Lines.Add('��������� '+GroupPCSelectProg+' �� '+myPCStringList[i]+' �� �������.');
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
               finditem('��� �������� ��������� '+GroupPCSelectProg+' ��������� ������. - "'+E.Message+'"',myPCStringList[i],2);
               frmDomainInfo.memo1.Lines.Add('������ �������� ��������� '+GroupPCSelectProg+' �� '+myPCStringList[i]+': "'+E.Message+'"');
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
