unit MyInventoryPC;

interface

uses
  System.Classes,ActiveX,ComObj,Variants,FireDAC.Stan.Intf, FireDAC.Stan.Option
  , FireDAC.Stan.Param,System.SysUtils,Vcl.Forms,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.Client, Data.DB,FireDAC.UI.Intf
  ,FireDAC.Comp.UI,FireDAC.Phys.FB,inifiles,SqlTimSt,IdIcmpClient,System.StrUtils,
  WinSock,Windows,System.TimeSpan,IdTCPClient;

type
/////////////////////////////////////////////
  TSMARTitem=record
    ImageIndex:integer; /// image
    IDname:integer;  //item.SubItems[0]
    Val:integer; // item.SubItems[1]
    CurWorst:integer; // item.SubItems[2]
    CurTresh:integer;  //  item.SubItems[3]
    CurValue:string[30]; //item.SubItems[4]
    PersentResult : integer; /// item.SubItems[5]

  end;
/////////////////////////////////////////////
  InventoryConfig = class(TThread)

  private
  function ConnectWMI(NamePC,User,Pass:string): Boolean;
  function DisconnectDB:boolean;
  function connectDB:boolean;
  procedure ScanPrinter(var printerName,countprinter :string);
  procedure NetworkInterface(var NIName,NIMAC,NISpeed,NIIP,NIMask,
NIGateway,NIDHCP,NIDNS,NIWINS,NIcount,NIcountIP,NIcountGW,NIcountDNS,
NIcountDHCP,NIcountWINS :string);
  procedure osconfig(var nameOS,VerOS,OSKey,invNumber,OSArchit,OSSP,InstalDate:string);
  procedure ScanHDD(var hddType,hddName,hddSize,hddsn,counthdd,InterfaceType,HDDfirmware,HDDMYsmart,HDDsmartOS,HDDtemp :string);
  procedure DesktopMon(var screenHW,NameMon,NamufactureMon,countMonitor,DPI :string);
  procedure videocard(var card,cardmem,countcard :string);
  procedure memorySpeed (var memSpeed:string );
  Procedure memory(var CountMem,SummSizeMem,SizeMem:string );
  procedure procConf(var ProcName,ProcCore,ProcSpeed,ProcArh,ProcSoc,ProcLogProc,CountProc:string);
  procedure monterboard(var montrName,montrSN,montrManuf:String); /// получение данных о материнской плате
  procedure Nameuser (var UserNamePC:String);
  procedure CDDVDRom (var CDRomName,CDRomCount,CDRomType:String);
  procedure SoundDevice (var SoundName,SoundCount:String);
  procedure BIOSMB(var BIOSDescription,BIOSCount,BIOSSN,SMBIOSBIOSVersionMM,BIOSStatus,BiosPrimary,BIOSManufac,Characteristics:string);
  function DeleteItemVarArray(a:variant{p,n:integer}):variant; ///удаление элемента массива
  function SmartSelect(selectHDD:string):string;
  function imageitems(Item: TSMARTitem):boolean; /////// оценка состояния атрибута SMART
  function Log_write(level:integer;fname, text:string):string;
  function Obrez(s:string):string;
  function Dectobin(id, col: Integer): string;
  function BIN2DEC(BIN: string): LONGINT;
  function MySendARP(const MyIPAddress: String): String;
  function WriteHardware(s,z:string):bool;
  function get135Port(s:string):bool;
  function SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
  protected
    procedure Execute; override;
  end;
function SendARP(DestIp: DWORD; srcIP: DWORD; pMacAddr: pointer; PhyAddrLen: Pointer): DWORD;stdcall; external 'iphlpapi.dll';


implementation
uses uMain,PingForInventoryPC;
ThreadVar
  FSWbemLocator       : OLEVariant;
  FWMIService         : OLEVariant;
  TransactionWrite    : TFDTransaction;
  FDQueryWrite        : TFDQuery;
  ConnectionThread: TFDConnection;
  CurrentPC,VerWin,
  InvUser,InvPass  : string;
  AnswerMAC        : string;
  AlllifeHDD,PersentageAtribute       : integer;
  tempMin,tempMax,tempCurr:integer;
  minLifeCurrentHDDforPC:integer;
  minLifeCurrentHDDforPCOS:string;
  readSmart:bool;


function InventoryConfig.Log_write(level:integer;fname, text:string):string;
var f:TStringList;        /// функция записи в лог файл
begin
try
  if not DirectoryExists('log') then CreateDir('log');
      f:=TStringList.Create;
      try
        if FileExists(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log') then
          f.LoadFromFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
        case level of
        0: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Info'+StringOfChar(' ',19)+text);
        1: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Warning'+StringOfChar(' ',13)+text);
        2: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Error'+StringOfChar(' ',17)+text);
        3: f.Insert(0,DateTimeToStr(Now)+chr(9)+'Critical Error'+StringOfChar(' ',5)+text);
        end;
        while f.Count>1000 do f.Delete(1000);
        f.SaveToFile(ExtractFilePath(Application.ExeName)+'log\'+fname+'.log');
      finally
        f.Destroy;
      end;
except
  exit;
end;
end;


function InventoryConfig.Obrez(s:string):string;
  begin
  result:='';
  delete(s,1,pos('\',s));
  delete(s,1,pos('\',s));
  result:=s;
  //result:=copy(s,1,pos('\',s)-1);
  //Memo2.Lines.Add(result);
  end;

function InventoryConfig.Dectobin(id, col: Integer): string;
var
  delimoe: Integer;
  StrBin: String;
begin
  delimoe := id;
  StrBin := '';
  repeat
    StrBin := IntToStr(delimoe mod 2) + StrBin;
    delimoe := delimoe div 2;
  until (delimoe = 0);
  while Length(StrBin) <> col do
    StrBin := '0' + StrBin;

  Result := StrBin;
end;

function InventoryConfig.BIN2DEC(BIN: string): LONGINT;
var
  J: LONGINT;
  Error: Boolean;
  DEC: LONGINT;
begin
  DEC := 0;
  Error := False;
  for J := 1 to Length(BIN) do
  begin
    if (BIN[J] <> '0') and (BIN[J] <> '1') then
      Error := True;
    if BIN[J] = '1' then
      DEC := DEC + (1 shl (Length(BIN) - J));
    { (1 SHL (Length(BIN) - J)) = 2^(Length(BIN)- J) }
  end;
  if Error then
    BIN2DEC := 0
  else
    BIN2DEC := DEC;
end;

/////////////////////////////////////// ARP
function InventoryConfig.MySendARP(const MyIPAddress: String): String;
var
  DestIP: ULONG;
  MacAddr: Array [0..5] of Byte;
  MacAddrLen: ULONG;
  SendArpResult: Cardinal;
begin
try
if (MyIPAddress ='') or (MyIPAddress ='0.0.0.0') then
 begin
   result:='';
   exit;
 end;
  DestIP := inet_addr(PAnsiChar(AnsiString(MyIPAddress)));
  MacAddrLen := Length(MacAddr);
  SendArpResult := SendARP(DestIP, 0, @MacAddr, @MacAddrLen);
  if SendArpResult = NO_ERROR then
    Result := Format('%2.2X-%2.2X-%2.2X-%2.2X-%2.2X-%2.2X',
                     [MacAddr[0], MacAddr[1], MacAddr[2],
                      MacAddr[3], MacAddr[4], MacAddr[5]])
  else
    Result := '';
  NetApiBufferFree(@MacAddr);
  NetApiBufferFree(@MacAddrLen)
Except
   on E: Exception do
      begin
      Result := '';
      Log_write(2,'Hardware',MyIPAddress+': Ошибка ARP  - '+E.Message);
      end;
  end;
end;
/////////////////////////////////////////////////////////////
function InventoryConfig.WriteHardware(s,z:string):bool;
begin
try
TransactionWrite.StartTransaction;
FDQueryWrite.SQL.Clear;
FDQueryWrite.SQL.Text:='update or insert into ALL_HARDWARE '
+'(NAME_HARDWARE,TYPE_HARDWARE) VALUES ('''+trim(s)+''','''+z+''') MATCHING (NAME_HARDWARE,TYPE_HARDWARE)';
FDQueryWrite.ExecSQL;
TransactionWrite.Commit;
FDQueryWrite.close;
result:=true;
except
result:=false;
end;
end;

//////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////
function InventoryConfig.get135Port(s:string):bool;
var
ScanTCPPort         : TIdTCPClient;
begin
try
if not RPCport then  /// если не проверям 135 порт
begin
result:=true;
exit;
end;
ScanTCPPort:=TIdTCPClient.Create;
ScanTCPPort.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
ScanTCPPort.Port:=135;
ScanTCPPort.ConnectTimeout:=pingtimeout;
ScanTCPPort.Host:=s;
try
ScanTCPPort.Connect;
if ScanTCPPort.Connected then result:=true
else result:=false;
finally
ScanTCPPort.Disconnect;
ScanTCPPort.Free;
end;
 Except
      begin
      result:=false;
      Log_write(0,'Hardware',s+': scan 135 potr Connect timeout  ');
      end;
  end;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
function InventoryConfig.SelectedPing(Var resIP:string; HostName:string;Timeout:integer):boolean;
var
avalible:boolean;
begin
resIP:='';
AnswerMAC:='';
  case PingType of
    1:
    begin
    avalible:=PingIdIcmp(resIP,HostName,pingtimeout);
    end;
    2:
    begin
    avalible:=PingGetaddrinfo(resIP,HostName,pingtimeout);
    end;
    3:
    begin
    avalible:=PingGetHostByName(resIP,HostName,pingtimeout);
    end;
  END;
  if not avalible then
  begin
  try
    TransactionWrite.StartTransaction;
     FDQueryWrite.SQL.Clear;
     FDQueryWrite.SQL.Text:='update or insert into MAIN_PC (PC_NAME,DATE_INV,ERROR_INV) VALUES ('''+HostName+''','''+datetostr(date)+''',''Превышен интервал ожидания запроса'') MATCHING (PC_NAME)';
     FDQueryWrite.ExecSQL;
     TransactionWrite.Commit;
     FDQueryWrite.close;
     frmDomainInfo.statusbarInv.Panels[3].Text:=HostName+' - не доступен';
   except on E: Exception do
   Log_write(1,'Software',HostName+' Ошибка SelectedPing '+e.Message) ;
   end;
  end
  else
  begin
  AnswerMAC:=MySendARP(resIP);
  end;
  Result:=avalible;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////
function InventoryConfig.imageitems(Item: TSMARTitem):boolean; /////// оценка состояния атрибута SMART
var   //subitemSMART     : TSMARTitem;
  i,LifeHDD: Integer;
  s:string;

 function exceptAtribute(z:integer):bool;
 var
 i:integer;
 const
  mass: array [0 .. 2] of integer = (190,194,199);
 begin
 result:=false;
 for I := 0 to Length(mass) do
   begin
     if mass[i]=z then
     begin
       result:=true;
       break;
     end;
   end;
 end;

  function ocenka(th:integer):integer;
  begin
  result:=0;
    case th of
    0:  result:=5;
    1:  result:=5;
    2:  result:=6;
    3:  result:=7;
    4:  result:=8;
    5:  result:=9;
    6:  result:=10;
    7:  result:=11;
    8:  result:=13;
    9,10: result:=round(th+((th/100)*60));
    11..20:result:=round(th+((th/100)*50));
    21..30:result:=round(th+((th/100)*45));
    31..40:result:=round(th+((th/100)*40));
    41..50: result:=round(th+((th/100)*35));
    51..70:result:=round(th+((th/100)*30));
    71..90:result:=round(th+((th/100)*25));
    91..110:result:=round(th+((th/100)*20));
    111..140:result:=round(th+((th/100)*15));
    else result:=round(th+((th/100)*10));
    end;
  end;
begin
 try
  lifehdd:=0;
  item.PersentResult:=0;
  case item.IDname of
/////////////////////////////////////////////////////////////////////////////////////////////////////
  1:
    begin
    if ((item.CurWorst)<=ocenka(item.CurTresh)) or
     ((item.Val)<=ocenka(item.CurTresh)) then Item.ImageIndex:=6;
    if ((item.CurWorst)<=(item.CurTresh)) or
     ((item.Val)<=(item.CurTresh)) then Item.ImageIndex:=3
     else Item.ImageIndex:=2;
    if (Item.CurValue='0') then begin item.ImageIndex:=2; end; /// если текущее значение равно 0 то все ок
    if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
    if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
    end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
    2: begin
       if((item.CurWorst)<=(item.CurTresh)) or
        ((item.Val)<=(item.CurTresh))
        then item.ImageIndex:=3 else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
     end;
   4:
     begin
       if(item.CurWorst<=(item.CurTresh)) or
        (item.Val<=(item.CurTresh))
        then item.ImageIndex:=3 else item.ImageIndex:=2;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
     end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
  3:
  begin
   //if(item.CurWorst>=item.CurRAW) or (item.Val>=item.CurRAW)
   //     then item.ImageIndex:=3 else item.ImageIndex:=2;
   item.ImageIndex:=2;
  end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
   5:
     begin
     if (item.CurValue='0') then begin item.ImageIndex:=2; end;
     if item.CurTresh<>0 then
        begin
        if (item.Val<=(item.CurTresh))or
        (item.CurWorst<=(item.CurTresh)) then item.ImageIndex:=3
             else
             if (item.Val<=ocenka(item.CurTresh)+15) or
                 (item.CurWorst<=ocenka(item.CurTresh)+15) then Item.ImageIndex:=6
              else
             if (item.Val<=ocenka((item.CurTresh)+10)) or
                 (item.CurWorst<=ocenka((item.CurTresh)+10)) then Item.ImageIndex:=7
              else Item.ImageIndex:=2;
        end
          else Item.ImageIndex:=4; /// если значение  item.CurRAW =  пусто то надо бы обновит
      if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+20;
      end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
   7:
      begin
      if (item.CurValue='0') then begin item.ImageIndex:=2; end;
      if (item.Val<=(item.CurTresh))or
          (item.CurWorst<=(item.CurTresh)) then item.ImageIndex:=3
          else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   8:      /// содержит среднюю производительность операций позиционирования головок
   begin
      if (item.CurValue='0') then begin item.ImageIndex:=2; end;
      if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=3
          else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
   9:
      begin ///// количество отработанных часов
        Item.ImageIndex:=2;
        if item.Val<=35 then item.ImageIndex:=7;
        if item.Val<=25 then  item.ImageIndex:=6;
        if (item.Val<=item.CurTresh+10)or
          (item.CurWorst<=item.CurTresh+10) then
         item.ImageIndex:=3;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+8;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+12;
     end;
///////////////////////////////////////////////////////////////////////////////////////
    10:  ///Количество повторных попыток запуска вращения.
        begin
        if (item.Val<=item.CurTresh)or
        (item.CurWorst<=item.CurTresh) then item.ImageIndex:=3
             else
             if (item.Val<=ocenka(item.CurTresh)) or
                 (item.CurWorst<=ocenka(item.CurTresh)) then Item.ImageIndex:=6
              else
             if (item.Val<=ocenka(item.CurTresh+10)) or
                 (item.CurWorst<=ocenka(item.CurTresh+10)) then Item.ImageIndex:=7
              else Item.ImageIndex:=2;
       if (item.CurValue='0') then item.ImageIndex:=2;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+20;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////
    11,12:  //11-Повторная калибровка или калибровка// 12-  включения / выключения жесткого диска
      begin
       if (item.CurValue='0') then begin item.ImageIndex:=2; end;
       if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=3
          else item.ImageIndex:=2;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////
    13,100: if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=7
          else item.ImageIndex:=2;
///////////////////////////////////////////////////////////////////////////////////////////////////////
    148..151,160,161,163..168:
      begin
      if (item.Val<=ocenka(item.CurTresh)) or
      (item.CurWorst<=ocenka(item.CurTresh)) then Item.ImageIndex:=6
      else Item.ImageIndex:=2;
      end;
/////////////////////////////////////////////////////////////////////////////////////////////////
    169:   ///Remain Life Percentage SSD
      begin
        try
        if (item.Val<=ocenka(item.CurTresh)) or
      (item.CurWorst<=ocenka(item.CurTresh)) then Item.ImageIndex:=6
      else Item.ImageIndex:=2;

         except
         Item.ImageIndex:=4;
        end;
      if (Item.ImageIndex=6)or(Item.ImageIndex=4) then lifehdd:=lifehdd+10;
      end;
///////////////////////////////////////////////////////////////////////////////////////////////////
    170:      ////параметр аналогичный ID 8
         begin
         if item.CurTresh<>0 then
            begin
            if (item.Val<=item.CurTresh)or
            (item.CurWorst<=item.CurTresh) then item.ImageIndex:=3
                 else
                 if (item.Val<=ocenka(item.CurTresh)) or
                     (item.CurWorst<=ocenka(item.CurTresh)) then Item.ImageIndex:=6
                     else Item.ImageIndex:=2;
            end
              else
              Item.ImageIndex:=2; /// если значение  item.CurRAW =  пусто то надо бы обновит
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          if (Item.ImageIndex=6)or(Item.ImageIndex=4) then lifehdd:=lifehdd+5;
          end;
 //////////////////////////////////////////////////////////////////////////////////////////////////
      171,172,173,174:
           begin
       if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=7
          else item.ImageIndex:=2;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
          end;
/////////////////////////////////////////////////////////////////////////////////////////////////
      175: item.ImageIndex:=4;
/////////////////////////////////////////////////////////////////////////////////////////////////
      176: if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=7
          else item.ImageIndex:=2;
////////////////////////////////////////////////////////////////////////////////////////////////
      177,179,180:
      begin
       if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=7
          else item.ImageIndex:=2;
         if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
       end;
///////////////////////////////////////////////////////////////////////////////////////////////
      181,182,183:
      begin
       if (item.Val<=item.CurTresh)or
          (item.CurWorst<=item.CurTresh) then item.ImageIndex:=7
          else item.ImageIndex:=2;
         if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
      end;
//////////////////////////////////////////////////////////////////////////////////////////
      184:
        begin   //// критический параметр   End-to-End error / IOEDC
        if item.CurValue='0' then item.ImageIndex:=2  else
          begin
           if (item.Val<=ocenka(item.CurTresh))or
            (item.CurWorst<=ocenka(item.CurTresh)) then
            item.ImageIndex:=3
            else
             if  item.CurValue>='1' then  item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
         if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
         if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
/////////////////////////////////////////////////////////////////////////////////////////
       185:
          begin
           if (item.Val<=ocenka(item.CurTresh))or
          (item.CurWorst<=ocenka(item.CurTresh)) then
          item.ImageIndex:=6
          else item.ImageIndex:=2;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          end;
//////////////////////////////////////////////////////////////////////////////////////////
       187:
          begin
          if (item.CurValue='0') then begin item.ImageIndex:=2; end;
          if (item.Val<=ocenka(item.CurTresh))or
          (item.CurWorst<=ocenka(item.CurTresh)) then
          item.ImageIndex:=3
          else item.ImageIndex:=2;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=4 then lifehdd:=lifehdd+10;
          end;
/////////////////////////////////////////////////////////////////////////////////////////
        188:
           begin
            if (item.CurValue='0') then begin item.ImageIndex:=2; end;
            if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then
                item.ImageIndex:=6
            else item.ImageIndex:=2;
             if (item.Val<=(item.CurTresh))or
               (item.CurWorst<=(item.CurTresh)) then
                item.ImageIndex:=3;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+20;
           end;
////////////////////////////////////////////////////////////////////////////////////////////////
        189:
          begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
        190:
          begin
          item.ImageIndex:=2;
          if (strtoint(item.CurValue)>=45) then item.ImageIndex:=7
          else if (strtoint(item.CurValue)>=50) then item.ImageIndex:=6
          else if (strtoint(item.CurValue)>=60) then item.ImageIndex:=3
          else item.ImageIndex:=2;

          if Item.ImageIndex=7 then Item.PersentResult:=75;
          if Item.ImageIndex=6 then Item.PersentResult:=50;
          if Item.ImageIndex=3 then Item.PersentResult:=0;
          if Item.ImageIndex=2 then Item.PersentResult:=100;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
        191:
          begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then
               item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
        192:
          begin
           if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then
                item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
        193:
          begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then
               item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        194:
         begin
          item.ImageIndex:=2;
          if (strtoint(item.CurValue)>=45) then item.ImageIndex:=7
          else if (strtoint(item.CurValue)>=50) then item.ImageIndex:=6
          else if (strtoint(item.CurValue)>=60) then item.ImageIndex:=3
          else item.ImageIndex:=2;

          if Item.ImageIndex=7 then Item.PersentResult:=75;
          if Item.ImageIndex=6 then Item.PersentResult:=50;
          if Item.ImageIndex=3 then Item.PersentResult:=0;
          if Item.ImageIndex=2 then Item.PersentResult:=100;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        195:
          begin
            if (item.Val<=ocenka(item.CurTresh)) then
               item.ImageIndex:=6;
             if (item.Val<=(item.CurTresh)) then
               item.ImageIndex:=3
            else item.ImageIndex:=2;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        196:
          begin
          if (item.CurValue='0') then begin item.ImageIndex:=2; end;
          if (item.Val<=ocenka(item.CurTresh)) then item.ImageIndex:=6;
          if (item.Val<=item.CurTresh) then  item.ImageIndex:=3
             else item.ImageIndex:=2;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
         197:
            begin
            item.ImageIndex:=2;
            if item.Val<=ocenka(item.CurTresh) then item.ImageIndex:=3;
              if (item.CurValue<>'0') then
               begin
               if strtoint(item.CurValue)<=9 then  item.ImageIndex:=2;
               if strtoint(item.CurValue)>=10 then  item.ImageIndex:=7;
               if strtoint(item.CurValue)>=50 then  item.ImageIndex:=6;
               if strtoint(item.CurValue)>=100 then  item.ImageIndex:=3;
               end
              else  item.ImageIndex:=4;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
          198:
             begin
             item.ImageIndex:=2;
            if (item.Val<=ocenka(item.CurTresh)) then
                 item.ImageIndex:=3
                else
              if (strtoint(item.CurValue)<>0) then
               begin
               if strtoint(item.CurValue)<=9 then  item.ImageIndex:=2;
               if strtoint(item.CurValue)>=10 then  item.ImageIndex:=7;
               if strtoint(item.CurValue)>=50 then  item.ImageIndex:=6;
               if strtoint(item.CurValue)>=100 then  item.ImageIndex:=3;
               end
              else  item.ImageIndex:=4;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         199:
         begin
               if strtoint(item.CurValue)=0 then item.ImageIndex:=2;
               if strtoint(item.CurValue)<=19 then  item.ImageIndex:=2;
               if strtoint(item.CurValue)>=20 then  item.ImageIndex:=7;
               if strtoint(item.CurValue)>=50 then  item.ImageIndex:=6;
               if (item.Val<=item.CurTresh)or
                 (item.CurWorst<=item.CurTresh) then
                item.ImageIndex:=3;
            if Item.ImageIndex=7 then Item.PersentResult:=75;
            if Item.ImageIndex=6 then Item.PersentResult:=50;// заменить кабель
            if Item.ImageIndex=3 then Item.PersentResult:=0; /// заменить кабель
            if Item.ImageIndex=2 then Item.PersentResult:=100;
         end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
         200: /// содержит частоту возникновения ошибок при записи
            begin
               if (strtoint(item.CurValue)=0) then item.ImageIndex:=2
                else
               begin
               if strtoint(item.CurValue)<=5 then  item.ImageIndex:=2;
               if strtoint(item.CurValue)>=6 then  item.ImageIndex:=7;
               if strtoint(item.CurValue)>=30 then  item.ImageIndex:=6
               else  item.ImageIndex:=2;
               end;
              if (item.Val<=item.CurTresh)or
                 (item.CurWorst<=item.CurTresh) then
                item.ImageIndex:=3;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+10; //
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15; ///
         end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
         201:  //содержит частоту возникновения ошибок чтения, произошедших по вине программного обеспечения
            begin
            begin
             item.ImageIndex:=2;
            if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then
                 item.ImageIndex:=3
                else
              if (strtoint(item.CurValue)<>0) then
               begin
               if strtoint(item.CurValue)<=9 then  item.ImageIndex:=2;
               if strtoint(item.CurValue)>=10 then  item.ImageIndex:=7;
               if strtoint(item.CurValue)>=50 then  item.ImageIndex:=6;
               if strtoint(item.CurValue)>=100 then  item.ImageIndex:=3;
               end;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////
         202: //  содержание атрибута — загадка, но проанализировав различные диски, могу констатировать, что ненулевое значение — это плохо
            begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
             if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
         203:   //содержит количество ошибок ECC
            begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////
         204:  //Количество ошибок, исправленных с помощью программного обеспечения для внутреннего исправления ошибок.
            begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
         205: /// Подсчет ошибок из-за высокой температуры.
            begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
          206: //Высота головок над поверхностью диска. Если оно слишком низкое, вероятность падения головы выше; если слишком высокий, ошибки чтения / записи более вероятны.
            begin
             if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
          207:  // Величина импульсного тока, используемого для вращения привода
            begin
            if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
          208:  // Количество циклов, необходимых для раскрутки
                //Количество повторных попыток во время ускорения из-за низкого доступного тока.
            begin
             if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10; /// проверить блок питания
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
          209:  /// Производительность диска во время автономных операций
                  //Производительность поиска накопителя во время внутренних самопроверок.
            begin
             if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////
          210:  /// вибрации во время записи
            begin
              if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////
       211: //Вибрация, возникающая во время операций записи.
        begin
          if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
       212: //Шок, возникший во время операций записи.
        begin
         if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        218:  // CRC Error Count (kingston SSD)
          begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
       220:  //Расстояние диска сместилось относительно шпинделя (обычно из-за удара или температуры).
        begin
         if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       221: //Параметр SMART Disk Shift указывает, что расстояние диска смещено относительно шпинделя,
            // что может быть вызвано механическим ударом или высокой температурой.
        begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
      222: ///  Количество отработанных часов .   Жесткие диски, поддерживающие этот атрибут
            //Samsung, Seagate, IBM (Hitachi), Fujitsu (не все модели), Maxtor, Western Digital (не все модели)
         begin
         Item.ImageIndex:=2;
          if item.Val<=35 then item.ImageIndex:=7;
          if item.Val<=25 then  item.ImageIndex:=6;
          if (item.Val<=item.CurTresh+10)or
            (item.CurWorst<=item.CurTresh+10) then
           item.ImageIndex:=3;
           if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+8;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+12;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
      223:  //Параметр SMART Load / Unload Retry Count указывает количество головок дисков, которые входят / выходят из зоны
        begin
        if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
       224: //Параметр SMART Load Friction указывает скорость трения между механическими частями жесткого диска.
            // Увеличение значения означает, что существует проблема с механической подсистемой привода.
        begin
         if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+15;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       225: ///Load / Unload Cycle Count SMART параметр указывает общее количество циклов загрузки.
        begin
         if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////
       226:  ///Общее время нагрузки на привод магнитных головок (время, не проведенное в зоне парковки)
        begin
          if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       227: // указывает на количество попыток скомпенсировать изменения скорости вращения диска.
        begin
        if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
       228://Число циклов выключения, которые подсчитываются, когда происходит «событие втягивания» и головки загружаются с носителя, например, когда машина выключена, переведена в спящий режим или находится в режиме ожидания.
        begin
         if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
       230:   // Head Amplitude GMR указывает расстояние перемещения головки между операциями   samsung, Seagate, IBM (Hitachi), Fujitsu (не все модели), Maxtor, Western Digital (не все модели)
        begin  /// для ssd критический параметр с worst=90  наилучший =100
        if (item.Val<=item.CurTresh)or
               (item.CurWorst<=item.CurTresh) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
         if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        231:  //Указывает приблизительный оставшийся срок службы SSD
          Item.ImageIndex:=4;
          {begin
          if strtoint(item.CurValue)<=30 then item.ImageIndex:=7;
          if strtoint(item.CurValue)<=20 then item.ImageIndex:=6;
          if strtoint(item.CurValue)<=10 then item.ImageIndex:=3
           else  item.ImageIndex:=2;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+10;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
          end; }
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        232:  //Число циклов физического стирания, выполненных на твердотельном накопителе, в процентах от максимального количества циклов физического стирания, которое рассчитан накопитель.
          begin
          if (item.Val<=item.CurTresh+5)or
               (item.CurWorst<=item.CurTresh+5) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+30;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
         233: //Индикатор износа носителя (SSD) или часы включения    Lifetime Nand Writes
           begin
             if item.CurTresh=0 then
                begin
                if (item.Val<=35) or (item.CurWorst<=35) then
                 item.ImageIndex:=6;
                if (item.Val<=10) or (item.CurWorst<=10) then
                 item.ImageIndex:=7
                   else item.ImageIndex:=2;
                end
             else
                begin
                   if (item.Val<=item.CurTresh+5)or
                   (item.CurWorst<=item.CurTresh+5) then  item.ImageIndex:=3
                    else item.ImageIndex:=2;
                end;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+10;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
           end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
         234:
          begin  //Среднее количество стираемых И максимальное количество стираемых
           if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////
         235:// Good Block Count AND System(Free) Block Count
          begin
            if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////
        240:  //Время, затрачиваемое на позиционирование приводных головок
          begin
            if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////
         241: //Всего написано LBA
           begin
           if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
           end;
/////////////////////////////////////////////////////////////////////////////////////////////////
         242: // Всего прочитанных LBA
          begin
           if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
         243:  //Всего написано LBA Расширено
          begin
           if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////
         244: // Всего прочитанных LBA Расширено  or Average Erase Count (Kingston SSD)
          begin
          if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
/////////////////////////////////////////////////////////////////////////////////////////////
         245: ///Max Erase Count (kingston)
          begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////
          246: ///Total Erase Count (kingston)
          begin
          if (item.Val<=ocenka(item.CurTresh))or
               (item.CurWorst<=ocenka(item.CurTresh)) then item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
///////////////////////////////////////////////////////////////////////////////////////
          249: //Необработанное значение сообщает о количестве записей в NAND с шагом 1 ГБ
           begin
           if (item.Val<=10) or (item.CurWorst<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+10;

           end;
/////////////////////////////////////////////////////////////////////////////////////////////////
          250:  ///Количество ошибок при чтении с диска.
            begin
            if item.CurTresh<>0then
               begin
              if (item.Val<=ocenka(item.CurTresh))or
                 (item.CurWorst<=ocenka(item.CurTresh)) then  item.ImageIndex:=3
                 else item.ImageIndex:=2;
               end
             else
              begin
                if (item.Val<=10) or (item.CurWorst<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
              end;
            if Item.ImageIndex=3 then  lifehdd:=lifehdd+20;
            if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
          251:  //Атрибут Minimum Spares Remaining указывает количество оставшихся запасных блоков в процентах от общего количества доступных запасных блоков.
            begin
              if (item.Val<=10) or (item.CurWorst<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
           252: //Недавно добавленная блокировка сбойной флэш-памяти указывает общее количество сбойных флэш-блоков, обнаруженных накопителем с момента его первоначальной инициализации в процессе производства.
            begin
               if (item.Val<=10) or (item.CurWorst<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
               if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
          254: //Количество обнаруженных событий «Свободное падение»
            begin
            if (item.Val<=10) or (item.CurWorst<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////
          else item.ImageIndex:=5;
  end;
   except
  item.ImageIndex:=4;
  end;

//////////////////// столбец ресурс
if not exceptAtribute(item.IDname) then
 begin
  case item.ImageIndex of
   7: item.PersentResult:=75;
   6: item.PersentResult:=50;
   3: item.PersentResult:=25;
   2: item.PersentResult:=100;
  end;

 end;
/////////////////////////////////
  AllLifehdd:=AllLifehdd+lifehdd;
  PersentageAtribute:=Item.PersentResult;
/////////////////////////////
Item.ImageIndex:=0;
Item.IDname:=0;
Item.Val:=0;
Item.CurWorst:=0;
Item.CurTresh:=0;
Item.CurValue:='';
Item.PersentResult:=0;
end;





function InventoryConfig.SmartSelect(selectHDD:string):string;
var // SMART Выбранного диска
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObjectVendorTresh: OLEVariant;
  oEnum,oEnumSM,oEnumVT: IEnumvariant;
  FWbemObject,FWbemObjectVT: OLEVariant;
  iValue: LongWord;
  SmartMAS: array of byte;
  ThresholdsArr: array of byte;
  Vendorspec4: array of byte;
  i,z,atribute190,atribute194,Atribute9,Atribute241,Atribute242,Atribute243: Integer;
  predID, SelfStat, selfStatProcent: Integer; //  Atribute241 - написано LBA/ Atribute242- прочитано LBA/Atribute243- дополнительно к написано LBA
  testImpl, SrtCapab,vendorspecific4,SpinUpTime,Atribute234,Atribut235,
  Atribute240,Atribute169,Atribute175: string;
  //AllTime:TTimeSpan;
  VendorSMART:TSMARTitem;
  YESsmart:bool;
//////////////////////////////////////////////////////////////////////
  function bubblesort(s: string):bool; /// функция сортировки и расчета миним и максим температуры
  var
  A: array[0..3] of integer;
  //B: array[0..3] of integer;
  i,j,p,z : integer;
  begin
  for z := 0 to 3 do
    begin
      a[z]:=strtoint(copy(s,1,pos(',',s)-1));
      delete(s,1,pos(',',s));
    end;
  for I :=1 to 3 do
  begin
    for j := 1 to 3-i do
    begin
     if A[j]>A[j+1] then
       begin
       p:=A[j];
       A[j]:=A[j+1];
       A[j+1]:=P;
       end;
    end;
  end;
  s:='';
  p:=0;
   for z := 0 to 3 do
   begin
   if (a[z]>0)and(a[z]<80) then
     begin
     s:=s+inttostr(a[z])+',';
     inc(p);
     end;
   end;
   if p>0 then
     begin
       case p of
       1:
       begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          tempMin:=strtoint(copy(s,1,pos(',',s)-1))
           else tempMax:=strtoint(copy(s,1,pos(',',s)-1));
       end;
       2:
       begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          tempMin:=strtoint(copy(s,1,pos(',',s)-1));
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          tempMax:=strtoint(copy(s,1,pos(',',s)-1));
       end;
        3:
         begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          tempmin:=strtoint(copy(s,1,pos(',',s)-1));
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
          delete(s,1,pos(',',s));//// удаляем вторую цифру и знак запятой т.к. нам надо MAX и MIN
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          tempMax:=strtoint(copy(s,1,pos(',',s)-1));
         end;
        4:
         begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          tempMin:=strtoint(copy(s,1,pos(',',s)-1));
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
          delete(s,1,pos(',',s));//// удаляем вторую цифру и знак запятой т.к. нам надо MAX и MIN
          delete(s,1,pos(',',s));//// удаляем третью цифру и знак запятой т.к. нам надо MAX и MIN
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          tempMax:=strtoint(copy(s,1,pos(',',s)-1));
         end;
       end;
     end;
  end;
///////////////////////////////////////////////////////////////////////////////////
  function delflags(s: string): string;
  begin
    if s <> '' then
    begin
      delete(s, Length(s) - 2, Length(s));
      Result := s;
    end
    else
      Result := '';
  end;
/////////////////////////////////////////////////////////////////////////////////////////////
begin
VendorSMART.ImageIndex:=0;
VendorSMART.IDname:=0;
VendorSMART.Val:=0;
VendorSMART.CurWorst:=0;
VendorSMART.CurTresh:=0;
VendorSMART.CurValue:='';
VendorSMART.PersentResult:=0;
Alllifehdd:=0;
YESsmart:=false;
tempCurr:=0;
 try
    fillchar(SmartMAS,sizeof(SmartMAS),0);
    fillchar(ThresholdsArr,sizeof(ThresholdsArr),0);
    fillchar(Vendorspec4,sizeof(Vendorspec4),0);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(CurrentPC,'root\WMI', InvUser, InvPass,'','',128);


    /// /////////////////////////////////////////////////////////////////////////////
    FWbemObjectVendorTresh := FWMIService.ExecQuery
    ///// заполняем массив с пороговыми значениями
      ('SELECT * FROM MSStorageDriver_FailurePredictThresholds WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',wbemFlagForwardOnly);
    oEnumVT := IUnknown(FWbemObjectVendorTresh._NewEnum) as IEnumvariant;
    while oEnumVT.Next(1, FWbemObjectVT, iValue) = 0 do
    begin
        SetLength(ThresholdsArr,VarArrayHighBound(FWbemObjectVT.properties_.Item('VendorSpecific',0).value, 1)); /// задаем длинну массива
        for i := 0 to VarArrayHighBound(FWbemObjectVT.properties_.Item('VendorSpecific', 0).value, 1) do
        begin
          ThresholdsArr[i] := FWbemObjectVT.properties_.Item('VendorSpecific',0).value[i];
        end;
      FWbemObjectVT := Unassigned;
    end;
    oEnumVT := nil;
    FWbemObjectVendorTresh := Unassigned;
    FWbemObjectVT:= Unassigned;
    VariantClear(FWbemObjectVT);
    VariantClear(FWbemObjectVendorTresh);
    /// /// непосредственно данные SMART
    FWbemObjectSet := Unassigned;
    VariantClear(FWbemObjectSet);
    oEnum:=nil;
    FWbemObjectSet := FWMIService.ExecQuery ('SELECT * FROM MSStorageDriver_ATAPISmartData WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL', wbemFlagForwardOnly);   // MSStorageDriver_ATAPISmartData или MSStorageDriver_FailurePredictData
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
        if copy(Dectobin(FWbemObject.ErrorLogCapability, 8), 8, 1) <> '1' then
          begin
          oEnum := nil;
          VariantClear(FWbemObject);
          VariantClear(FWbemObjectSet);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
          result:= 'Not Supported';
          exit;
          end;
          YESsmart:=true;
        /////////////////////////////////////////////////////////////////////////////////////////////

        SetLength(SmartMAS, VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific',0).value, 1));
        for i := 0 to VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific', 0).value, 1) do
        begin
          SmartMAS[i] := FWbemObject.properties_.Item('VendorSpecific',0).value[i]
        end;
          predID := 0;
          /// / предыдущий ID необходим для того чтобы не выводились непонятные ID после последнего максимального
          for i := 0 to 29 do  // почему 29 , данные располагаются с 0 го по 361 байт
          begin
            if (SmartMAS[i * 12 + 2] <> 0) and (SmartMAS[i * 12 + 2] > predID) then /// если ID больше 0 и текущий ID больше предыдущего
            begin
              with VendorSMART do
              begin
                predID := SmartMAS[i * 12 + 2];
                IDname := (SmartMAS[i * 12 + 2]); // ID
                Val:=(SmartMAS[i * 12 + 5]);   // 4 текущее value get 6th column where actual normalized data is stored
                CurWorst:=(SmartMAS[i * 12 + 6]);   // 5 наихудшее значение worst get 7th column where worst normalized data is stored

//////////////////////////////////////////////////////////////////////////////////// Threshold Пороговое занчение
                //if ThresholdsArr[i * 12 + 2] = SmartMAS[i * 12 + 2] then memo2.Lines.add(IntToStr(ThresholdsArr[i * 12 + 3]));
                if ThresholdsArr[i * 12 + 2] <> SmartMAS[i * 12 + 2] then
                  begin
                    for z := 0 to 29 do
                    begin
                     if ThresholdsArr[i * 12 + 2] = SmartMAS[z * 12 + 2] then
                      begin
                      CurTresh:=(ThresholdsArr[z * 12 + 3]);
                      break;
                      end
                      else if z=29 then CurTresh:=0;;
                      end;
                  end
                else  CurTresh:=(ThresholdsArr[i * 12 + 3]);
//////////////////////////////////////////////////////////////////////////////////////////////////////////
                case SmartMAS[i * 12 + 2] of
                // 16 в 8й степени          6й степени                 4й степени            2й степени
                  1,2,4 .. 8: CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  3:  begin
                      SpinUpTime:='';
                      if SmartMAS[i * 12 + 11]<>0 then SpinUpTime:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 7]);
                      CurValue:=(SpinUpTime);
                      end;
                  9:begin //// отработаных часов
                    atribute9:=0;
                    atribute9:=(((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                    CurValue:=(inttostr(Atribute9));
                    end;
                  10 .. 13: CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  22: CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                 // 100: SubItems.Add(inttostr(BIN2DEC(copy((Dectobin(SmartMAS[i * 12 + 7],6)),4,3))));    ????
                  100:CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  169:
                     begin     //// оставшийся процент жизни SSD
                     Atribute169:='';
                     if SmartMAS[i * 12 + 11]<>0 then Atribute169:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 7]);
                      CurValue:=(Atribute169);
                     end;
                  170 .. 174:CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  175 :///// 175 расчитывается по байтам 0-1 последний результат теста в микросекундах, 2-3: минуты с момента последнего теста, 4-5: количество тестов в течение жизни
                  begin
                   Atribute175:='';
                   if SmartMAS[i * 12 + 7]<>0 then  Atribute175:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                   if SmartMAS[i * 12 + 8]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                   if SmartMAS[i * 12 + 9]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                   if SmartMAS[i * 12 + 10]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                   if SmartMAS[i * 12 + 11]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 7]);
                   CurValue:=( Atribute175);
                  end;

                  176 .. 189:CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  191 .. 193:CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  195 ..233,236..239,244..254:CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  190:begin
                      atribute190:=0;
                      CurValue:=(IntToStr(SmartMAS[i * 12 + 7])); // температура
                      atribute190:=(SmartMAS[i * 12 + 7]);
                      tempCurr:=(atribute190);
                      end;
                  194: begin
                        atribute194:=0;
                        atribute194:=(SmartMAS[i * 12 + 7]);
                        CurValue:=(IntToStr(SmartMAS[i * 12 + 7])); // температура
                        bubblesort(inttostr(SmartMAS[i * 12 + 8])+','+(IntToStr(SmartMAS[i * 12 + 9]))+','+(IntToStr(SmartMAS[i * 12 + 10]))+','+(IntToStr(SmartMAS[i * 12 + 11]))+',');
                        tempCurr:=(atribute194);
                       end;
                  234: begin //Decoded as: byte 0-1-2 = average erase count (big endian) and byte 3-4-5 = max erase count (big endian)
                       Atribute234:='';
                       Atribute234:=inttostr((SmartMAS[i * 12 + 7])+(SmartMAS[i * 12 + 8]*256)+(SmartMAS[i * 12 + 9]+65536))+'/';
                       Atribute234:=Atribute234+inttostr((SmartMAS[i * 12 + 10]+16777216)+(SmartMAS[i * 12 + 11]*4294967296)+(SmartMAS[i * 12 + 12]+1099511627776));
                       CurValue:=(Atribute234);
                       end;
                  235: begin
                       Atribut235:='';
                       Atribut235:=inttostr((SmartMAS[i * 12 + 7])+(SmartMAS[i * 12 + 8]*256)+(SmartMAS[i * 12 + 9]+65536))+'/';
                       Atribut235:=Atribut235+inttostr((SmartMAS[i * 12 + 10]+16777216)+(SmartMAS[i * 12 + 11]*4294967296));
                       CurValue:=(Atribut235);
                       end;
                  240:  begin
                      Atribute240:='';
                      if SmartMAS[i * 12 + 11]<>0 then Atribute240:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 7]);
                       CurValue:=(Atribute240);
                      end;
                   241:  /// написано LBA
                       begin
                       Atribute241:=0;
                       Atribute241:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       CurValue:=(inttostr(Atribute241));
                       //if sizeLBA<>0 then
                      //  begin
                       //  Atribute241:=(((Atribute241*SizeLBA) div 1024) div 1024)div 1024;
                       // end;
                      // CurValue:=inttostr(Atribute241) +' Гб';
                       end;
                   242:  // Прочитано LBA
                       begin
                       Atribute242:=0;
                       Atribute242:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       CurValue:=(inttostr(Atribute242));
                       //if sizeLBA<>0 then
                       // begin
                       //  Atribute242:=(((Atribute242*SizeLBA)div 1024)div 1024)div 1024;
                       //  CurValue:=inttostr(Atribute242)+' Гб';
                       // end;

                       end;
                   243:  // Дополнительно написано LBA
                       begin
                       Atribute243:=0;
                       Atribute243:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       CurValue:=(inttostr(Atribute243));
                      // if sizeLBA<>0 then
                      //  begin
                       //  Atribute243:=(((Atribute243*SizeLBA) div 1024)div 1024)div 1024;
                       //  Atribute243:=Atribute243+Atribute241;
                       //  StaticText20.Caption:='Всего записано - '+inttostr(Atribute243)+' Гб';
                      //  end;
                       end;

                else
                  CurValue:=(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                end;
                imageitems(VendorSMART);
 ////////////////////////////////////////////////////запись в лог файл данных SMART
 { Log_write('SMARTRes',datetostr(date)+'/'+timetostr(time)
  +'/'+CurrentPC+' - ImageIndex -'+inttostr(VendorSMART.ImageIndex)+
                 ' - IDname - '+inttostr(VendorSMART.IDname)+
                  ' - Val - '+inttostr(VendorSMART.Val)+
                  ' - CurWorst - '+inttostr(VendorSMART.CurWorst)+
                  ' - CurTresh - '+inttostr(VendorSMART.CurTresh)+
                  ' - CurValue - '+(VendorSMART.CurValue)+
                   ' - PersentageAtribute - '+inttostr(PersentageAtribute)
                 ); }
 ///////////////////////////////////////////////////////////////////////////////
              end;               /// with VendorSMART do
            end;/// if (smartMas[i*12+2]<>0) and (smartmas[i*12+2]>predID) then

          end;  /// for I := 0 to 29 do
      FWbemObject := Unassigned;
      SmartMAS := nil;
      ThresholdsArr := nil;  ///???
      Vendorspec4:=nil;     /// ???
      break;                /// ???
  end;
 ////////////////////////////////////////////////////////
 if YESsmart then
 begin
  if Alllifehdd>=100 then   /////// вывод жизни диска
    begin
     Alllifehdd:=5;
     end
  else
    begin
    Alllifehdd:=100-Alllifehdd;
    end;
//////////////////////////////////////////////////////////
   if minLifeCurrentHDDforPC>Alllifehdd then
   minLifeCurrentHDDforPC:= Alllifehdd;
////////////////////////////////////////////////////////
   result:=inttostr(Alllifehdd);
   //// запись в лог файл оценки состояния HDD
  // Log_write('SMARTRes',datetostr(date)+'/'+timetostr(time)+' - Alllifehdd -  '+inttostr(Alllifehdd));
  end
else result:='Not supported';

 ///////////////////////////////////////////////////////////////
  FWbemObject := Unassigned;
  VariantClear(FWbemObject);
  oEnum := nil;
  except
    on E: Exception do
    begin
      Log_write(1,'Hardware',CurrentPC+' - ошибка загрузки данных SMART-'+e.Message);
      oEnum := nil;
      VariantClear(FWbemObject);
      VariantClear(FWbemObjectSet);;
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      result:=e.Message;
      exit;
    end;
  end;
  oEnum := nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);;
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  SmartMAS := nil;
  ThresholdsArr := nil;
  Vendorspec4:=nil;
  selectHDD:='';
end;
///////////////////////////////////////////////////////////////////////////////////////////
function InventoryConfig.DeleteItemVarArray(a:variant{p,n:integer}):variant; ///удаление элемента массива
var i,z,x:integer;
    v:array of variant;
begin
  x:=0;
 for i:=0 to VarArrayHighBound(a,1) do
 begin
 if not VarIsEmpty(a[i]) then inc(z);
 end;
 SetLength(v,z);
 for i:=0 to VarArrayHighBound(a,1) do
   begin
     if not VarIsEmpty(a[i]) then
       begin
        v[x]:=a[i];
        inc(x);
       end;
   end;
z:=0;
result:=v;
VarClear(a);
v:=Unassigned;
end;
///////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.BIOSMB(var BIOSDescription,BIOSCount,BIOSSN,SMBIOSBIOSVersionMM,BIOSStatus,BiosPrimary,BIOSManufac,Characteristics:string);
var
i,z:integer;
mj,mn,s:string;
BIOSChr: variant;
iValue: LongWord;
oEnum : IEnumvariant;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
function bioscharact(i:integer):string;
begin
  case i of
  4:result:= 'ISA'; 5:result:= 'MCA'; 6:result:= 'EISA'; 7:result:='PCI';8:result:= 'PC Card (PCMCIA)';
  9:result:= 'Plug and Play'; 10:result:='APM'; 11:result:='Обновление BIOS (Flash)'; 12:result:= 'Дублирование BIOS';
  13:result:='VL-VESA';14:result:= 'ESCD';15:result:='Загрузка с CD';16:result:='Выбор загрузки';17:result:='ПЗУ BIOS установлен в разъем';
  18:result:='Загрузка с PC Card (PCMCIA)';19:result:= 'Enhanced Disk Drive'; 26:result:='Print Screen';
  31:result:= 'NEC PC-98'; 32:result:='ACPI';33:result:='USB Legacy';34:result:='AGP'; 35:result:='Загрузка c I2O';
  36:result:='Загрузка c LS-120'; 37:result:='Загрузка с ATAPI ZIP Drive';38:result:='1394 Boot';
  39:result:= 'Smart Battery'
  else result:='';
  end;
end;
begin
try
    i:=0;
       FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_BIOS','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
            begin
              try
              //////////////////////////////////////////////////////////// поддерживаемые функции
              if VarIsArray(FWbemObject.BiosCharacteristics) then
               begin
                 BIOSChr:=FWbemObject.BiosCharacteristics;
                 for z := 0 to VarArrayHighBound(BIOSChr,1) do
                 begin
                  s:=bioscharact(integer(BIOSChr[z]));
                  if s<>'' then Characteristics:=Characteristics+s+'/**/';
                  end;
               end;
               ////////////////////////////////////////////////////////////
                if FWbemObject.Description<>null then
                begin
                BIOSDescription:=BIOSDescription+trim(vartostr(FWbemObject.Description))+'/**/';
                WriteHardware(string(FWbemObject.Description),'BIOS'); /// запись в таблицу ALL_HARDWARE
                end;
              if FWbemObject.SerialNumber<>null then BIOSSN:=trim(vartostr(FWbemObject.SerialNumber)+'/**/');
              if FWbemObject.SMBIOSBIOSVersion<>null then  SMBIOSBIOSVersionMM:=trim(vartostr(FWbemObject.SMBIOSBIOSVersion));
              if FWbemObject.SMBIOSMajorVersion<>null then mj:=' ver '+vartostr(FWbemObject.SMBIOSMajorVersion);
              if FWbemObject.SMBIOSMinorVersion<>null then mn:='.'+vartostr(FWbemObject.SMBIOSMinorVersion);
              SMBIOSBIOSVersionMM:=SMBIOSBIOSVersionMM+mj+mn+'/**/';
              if FWbemObject.status<>null then BIOSStatus:=trim(vartostr(FWbemObject.status)+'/**/');
              if FWbemObject.PrimaryBIOS<>null then BiosPrimary:=trim(VarToStr(FWbemObject.PrimaryBIOS)+'/**/');
              if FWbemObject.Manufacturer<>null then BIOSManufac:=trim(vartostr(FWbemObject.Manufacturer)+'/**/');
              FWbemObject:=Unassigned;
              inc(i);
              BIOSCount:=inttostr(i);
              except
               BIOSDescription:='unknown/**/';
               BIOSCount:='0';
               FWbemObject:=Unassigned;
              end;
            end;
         if BIOSDescription='' then BIOSDescription:='unknown/**/';
         if BIOSCount='' then BIOSCount:='0';
         if Characteristics='' then Characteristics:='unknown/**/' ;

         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         VarClear(BIOSChr);/// очистка варианта
        //BIOSChr:=DeleteItemVarArray(BIOSChr); /// очистка массива  , упс ошибка
     except
      on E:Exception do
        begin
         Log_write(1,'Hardware',CurrentPC+' - BIOS -'+e.Message);
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         BIOSDescription:='unknown/**/';
         BIOSCount:='0';
        end;
      end;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.SoundDevice (var SoundName,SoundCount:String);
var
i:integer;
iValue              : LongWord;
oEnum               : IEnumvariant;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
    i:=0;
       FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_SoundDevice','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
            begin
              try
              if FWbemObject.Caption<>null then
              SoundName:=SoundName+trim(string(FWbemObject.Caption))+'/**/';
              WriteHardware(string(FWbemObject.Caption),'Sound device'); /// запись в таблицу ALL_HARDWARE
              FWbemObject:=Unassigned;
              inc(i);
              SoundCount:=inttostr(i);
              except
               soundName:='unknown/**/';
               SoundCount:='1';
               FWbemObject:=Unassigned;
              end;
            FWbemObject:=Unassigned;
            end;
         if SoundName='' then SoundName:='unknown/**/';
         if SoundCount='' then SoundCount:='0';

         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
     except
      on E:Exception do
        begin
         Log_write(1,'Hardware',CurrentPC+' - DeviceSound-'+e.Message);
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         SoundName:='unknown/**/';
         SoundCount:='0';
        end;
      end;
end;


procedure InventoryConfig.CDDVDRom (var CDRomName,CDRomCount,CDRomType:String);
var
i:integer;
iValue         : LongWord;
oEnum          : IEnumvariant;
FWbemObjectSet : OLEVariant;
FWbemObject    : OLEVariant;
begin
try
     i:=0;
       FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM  Win32_CDROMDrive','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
            begin
              try
                if FWbemObject.Name<>null then
                  begin
                     CDRomName:=CDRomName+trim(string(FWbemObject.Name))+'/**/';
                     WriteHardware(string(FWbemObject.Name),'CD/DVD-ROM'); /// запись в таблицу ALL_HARDWARE
                  if FWbemObject.MediaType<>null then
                     CDRomType:=CDRomType+trim(string(FWbemObject.MediaType))+'/**/'
                  else
                     CDRomType:=CDRomType+'unknown/**/';
                  inc(i);
                  CDRomCount:=inttostr(i);
                  end;
                FWbemObject:=Unassigned;
                except
                 CDRomName:='unknown/**/';
                 CDRomType:='unknown/**/';
                 CDRomCount:='1';
                end;
            end;
         if CDRomCount='' then CDRomCount:='0';
         if CDRomName='' then  CDRomName:='unknown/**/';
         if CDRomType='' then  CDRomType:='unknown/**/';         
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
     except
      on E:Exception do
        begin
         Log_write(1,'Hardware',CurrentPC+' - CD/DVD-'+e.Message);
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         CDRomName:='unknown/**/';
         CDRomType:='unknown/**/';
         CDRomCount:='0';
        end;
      end;
end;


procedure InventoryConfig.Nameuser (var UserNamePC:String);
var
iValue              : LongWord;
oEnum               : IEnumvariant;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
       FWbemObjectSet:= FWMIService.ExecQuery('SELECT UserName FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Пользователь
            begin
            try
            if FWbemObject.UserName<>null then
            UserNamePC:=trim(string(FWbemObject.UserName));
            FWbemObject:=Unassigned;
            except
             UserNamePC:='unknown';
            end;
            end;
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
     except
      on E:Exception do
        begin
         Log_write(1,'Hardware',CurrentPC+' - UserName-'+e.Message);
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         UserNamePC:='unknown';
        end;
      end;
end;


procedure InventoryConfig.monterboard(var montrName,montrSN,montrManuf:String); /// получение данных о материнской плате
var
iValue              : LongWord;
oEnum               : IEnumvariant;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
       FWbemObjectSet:= FWMIService.ExecQuery('SELECT SerialNumber,Product,Manufacturer FROM Win32_BaseBoard','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// мАТЕРИНСКАЯ ПЛАТА
            begin
       ///////////////////
              try
              if FWbemObject.SerialNumber<>null then
               montrSN:=trim(string(FWbemObject.SerialNumber));
              except
               montrSN:='unknown';
              end;
        /////////////////////
              try
              if FWbemObject.Product<>null then
                 montrName:=trim(string(FWbemObject.Product));
                 WriteHardware(montrName,'Motherboard'); /// запись в таблицу ALL_HARDWARE
              except
              montrName:='unknown';
              end;
       //////////////////////
              try
              if FWbemObject.Manufacturer<>null then
                 montrManuf:=trim(string(FWbemObject.Manufacturer));
              except
              montrManuf:='unknown';
              end;
         ////////////////////
              FWbemObject:=Unassigned;
            end;
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
     except
      on E:Exception do
        begin
         Log_write(1,'Hardware',CurrentPC+' - Motherboard-'+e.Message);
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet)
        end;
      end;
end;
//////////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.procConf(var ProcName,ProcCore,ProcSpeed,ProcArh,ProcSoc,ProcLogProc,CountProc:string);
var
i:byte;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
       try
       begin
        i:=0;
        FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Processor','WQL',wbemFlagForwardOnly);
        oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
        while oEnum.Next(1, FWbemObject, iValue) = 0 do  /// процессор   NumberCore,Arhitecture
          begin
              if string(FWbemObject.DeviceID)<>null then
                begin
            ////////////////////
                try
                if (FWbemObject.DataWidth)<>null then
                  ProcArh:=trim(ProcArh+inttostr(FWbemObject.DataWidth))+'/**/';
                except
                 ProcArh:=trim(ProcArh+'unknown')+'/**/';
                end;
           ///////////////////////
                try
                begin
                ProcName:=trim(ProcName+string(FWbemObject.Name))+'/**/';
                WriteHardware(FWbemObject.Name,'Processor'); /// запись в таблицу ALL_HARDWARE
                end;
                except
                ProcName:=trim(ProcName+'unknown')+'/**/';
                end;
           ////////////////////////
                try
                if FWbemObject.MaxClockSpeed<>null then
                ProcSpeed:=trim(ProcSpeed+inttostr(FWbemObject.MaxClockSpeed))+'/**/';
                except
                ProcSpeed:=trim(ProcSpeed+'unknown')+'/**/';
                end;
          /////////////////////////
                try
                if FWbemObject.SocketDesignation<>null then
                ProcSoc:=trim(ProcSoc+string(FWbemObject.SocketDesignation))+'/**/';
                except
                ProcSoc:=trim(ProcSoc+'unknown')+'/**/';
                end;
        //////////////////////////
                try
                if FWbemObject.NumberOfLogicalProcessors<>null then
                ProcLogProc:=trim(ProcLogProc+inttostr(FWbemObject.NumberOfLogicalProcessors))+'/**/';
                except
                ProcLogProc:=trim(ProcLogProc+inttostr(i))+'/**/';
                end;
                inc(i);
                CountProc:=inttostr(i);
       /////////////////////////
                try
                if (FWbemObject.NumberOfCores)<>null then
                ProcCore:=trim(ProcCore+inttostr(FWbemObject.NumberOfCores))+'/**/';
                Except
                ProcCore:=CountProc+'/**/';
                end;
        ///////////////////////////
                 end;
           FWbemObject:=Unassigned; /// обязательно для освобождения памяти на слежующую итерацию  while oEnum.Next(1, FWbemObject, iValue) = 0 do
            end;
        end;
         Except
            on E:Exception do
              begin
              Log_write(1,'Hardware',CurrentPC+' - Processor -'+e.Message);
              VariantClear(FWbemObject);
              oEnum:=nil;
              VariantClear(FWbemObjectSet)
              end;
          end;
        VariantClear(FWbemObject);
        oEnum:=nil;
        VariantClear(FWbemObjectSet);
end;
//////////////////////////////////////////////////////////////////////////////////////
Procedure InventoryConfig.memory(var CountMem,SummSizeMem,SizeMem:string );
var
i,Bank,Summ:integer;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
  Bank:=0;
  summ:=0;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT tag,Speed,Capacity FROM Win32_PhysicalMemory','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////Объем Оперативная память
            begin
            for I := 0 to 32 do
             begin
              if string(FWbemObject.tag)='Physical Memory'+' '+inttostr(i) then
                begin
                   if (FWbemObject.Capacity)<>null then
                     begin
                      inc(Bank);
                      summ:=summ+ (FWbemObject.Capacity/1024/1024);
                      Sizemem:=trim(sizemem+(inttostr(FWbemObject.Capacity/1024/1024)))+'/**/';
                     end;
                end;
              end;
            FWbemObject:=Unassigned;
            CountMem:=inttostr(bank); //// количество занятых слотов
            SummSizeMem:=inttostr(summ);   //// общая память         
         end;
         //// если не попали в цикл то проверяем переменные на пустые значения
        if CountMem='' then CountMem:='0';
        if SummSizeMem='' then SummSizeMem:='0';
        if Sizemem='' then  Sizemem:='0';        
Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - MemorySize -'+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    CountMem:='0';
    SummSizeMem:='0';
    Sizemem:='0/**/';
    end;
  end;
  oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
end;
//////////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.memorySpeed (var memSpeed:string );
var
i:integer;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT tag,Speed FROM Win32_PhysicalMemory','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////Объем Оперативная память
            begin
            for I := 0 to 32 do
              begin
              if string(FWbemObject.tag)='Physical Memory'+' '+inttostr(i) then
                begin
                   if (FWbemObject.Speed)<>null then
                     begin
                     memSpeed:=trim(memSpeed+(string(FWbemObject.speed)))+'/**/';
                     end;
                end;
              end;
            FWbemObject:=Unassigned;
            end;
   memSpeed:=StringReplace(memSpeed,' ','',[rfReplaceAll]);
  if memSpeed='' then memSpeed:= '0';

  Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - Memory Speed - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    memSpeed:='0/**/';
    end;
   end;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
end;
//////////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.videocard(var card,cardmem,countcard :string);
var i:byte;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
i:=0;
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_VideoController','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Видеокарта
              begin
              inc(i);
           ////////////////////////////////////////
              try
                if FWbemObject.AdapterRAM<>null then
                    begin
                     cardmem:=cardmem+trim(inttostr(FWbemObject.AdapterRAM/1024/1024))+'/**/';
                     WriteHardware(string(FWbemObject.Description),'Video Card'); /// запись в таблицу ALL_HARDWARE
                     card:=card+trim(string(FWbemObject.Description))+'/**/';
                    end
                else
                begin
                cardmem:=cardmem+'0/**/';
                card:=card+trim(string(FWbemObject.Description))+'/**/';
                end;
              FWbemObject:=Unassigned;
              except
               cardmem:=cardmem+'0/**/';
               card:=card+'unknown/**/';
               FWbemObject:=Unassigned;
              end;
           ////////////////////////////////////
            end;
           cardmem:=StringReplace(cardmem,'-','',[rfReplaceAll]);/// убираем лишние дефисы
           countcard:=inttostr(i);
           if cardmem='' then cardmem:='0/**/';
           if card='' then card:='unknown/**/';


  Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - VideoController - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet)
    end;
   end;
   FWbemObject:=Unassigned;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
end;
///////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.DesktopMon(var screenHW,NameMon,NamufactureMon,countMonitor,DPI :string);
var i:byte;
oEnum               : IEnumvariant;
iValue              : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
i:=0;
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT PixelsPerXLogicalInch,PixelsPerYLogicalInch,ScreenHeight,ScreenWidth,Caption,MonitorManufacturer FROM Win32_DesktopMonitor','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// монитор
             if FWbemObject.ScreenWidth<>null then
              begin
             ///////////////////
                 try
                if (FWbemObject.PixelsPerXLogicalInch)<>null then
                 DPI:=DPI+string(FWbemObject.PixelsPerXLogicalInch)+'x'+string(FWbemObject.PixelsPerYLogicalInch)+'/**/'
                 else DPI:=DPI+'unknown/**/';
                 except
                  DPI:=DPI+'unknown/**/';
                 end;
            //////////////////////
                 try
                if (FWbemObject.ScreenHeight)<>null then
                 screenHW:=screenHW+string(FWbemObject.ScreenWidth)+'x'+string(FWbemObject.ScreenHeight)+'/**/'
                 else screenHW:=screenHW+'unknown/**/';
                  except
                  screenHW:=screenHW+'unknown/**/';
                 end;
          ///////////////////////
                 try
                if (FWbemObject.Caption)<>null then
                begin
                 NameMon:=NameMon+string(FWbemObject.Caption)+'/**/';
                 WriteHardware(string(FWbemObject.Caption),'Monitor'); /// запись в таблицу ALL_HARDWARE
                 inc(i);
                end
                 else NameMon:=NameMon+'unknown/**/';
                 except
                  NameMon:=NameMon+'unknown/**/';
                 end;
          ////////////////////////
                try
                if (FWbemObject.MonitorManufacturer)<>null then
                begin
                 NamufactureMon:=NamufactureMon+string(FWbemObject.MonitorManufacturer)+'/**/'
                end
                 else NamufactureMon:=NamufactureMon+'unknown/**/';
                 except
                 NamufactureMon:=NamufactureMon+'unknown/**/';
                end;
       //////////////////////////
        FWbemObject:=Unassigned;
              end;
        FWbemObject:=Unassigned;
        countMonitor:=inttostr(i);
          
  Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - DesktopMonitor - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet)
    end;
   end;
   FWbemObject:=Unassigned;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
end;
/////////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.ScanHDD(var hddType,hddName,hddSize,hddsn,counthdd,InterfaceType,HDDfirmware,HDDMYsmart,HDDsmartOS,HDDtemp :string);
var
i,indexHDD:integer;
oEnum             : IEnumvariant;
iValue            : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
function snhddPhisical(idHdd:string):string; ///// функция извлечения serial number
var
FWbemObjectSetSN  : OLEVariant;
oEnumSN           : IEnumvariant;
FWbemObjectSN     : OLEVariant;
iValueSN          : LongWord;
oEnum             : IEnumvariant;
iValue            : LongWord;
begin
try
idHdd:=copy(idHdd,5,length(idHdd));
Result := '';
FWbemObjectSetSN:= FWMIService.ExecQuery
('SELECT * FROM Win32_PhysicalMedia WHERE Tag LIKE "%'+idHdd+'%"', 'WQL',wbemFlagForwardOnly);
oEnumsn:= IUnknown(FWbemObjectSetSN._NewEnum) as IEnumVariant;
while oEnumSN.Next(1, FWbemObjectSN, iValueSN) = 0 do
  begin
   result:=trim(vartostr(FWbemObjectSN.SerialNumber))
  end;
oEnumSN:=nil;
VariantClear(FWbemObjectSN);
VariantClear(FWbemObjectSetSN);
except
result:='unknown';
end;
end;

begin
try
      i:=0;
      indexHDD:=0;
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * From Win32_DiskDrive WHERE MediaType=''Fixed hard disk media''','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// HDD
            begin
           /////////////
              try
              if FWbemObject.Caption<>null then
              begin
                hddName:=trim(hddName+string(FWbemObject.Caption))+'/**/';
                WriteHardware(string(FWbemObject.Caption),'HDD'); /// запись в таблицу ALL_HARDWARE
                inc(i);
              end;
              except
               hddName:=trim(hddName+'unknown')+'/**/';
                inc(i);
              end;
            ///////////// type
              try
              if FWbemObject.MediaType<>null then  hddType:=trim(hddType+string(FWbemObject.MediaType))+'/**/';
              except
              hddType:=trim(hddType+'unknown')+'/**/';
              end;
           /////////// serial number
             hddsn:=hddsn+snhddPhisical(vartostr(FWbemObject.DeviceID))+'/**/';
           /////////////// size
              try
              if FWbemObject.size<>null then
               hddSize:=hddSize+inttostr(((FWbemObject.size)/1024/1024/1024))+'/**/'
               else hddSize:=hddSize+'unknown/**/';
              except
               hddSize:=hddSize+'unknown/**/';
              end;
        ///////////////   type interface
             try
              if FWbemobject.InterfaceType<>null then
                InterfaceType:=trim(InterfaceType+vartostr(FWbemobject.InterfaceType))+'/**/'
              else InterfaceType:=InterfaceType+'unknown/**/';
             except
              InterfaceType:=InterfaceType+'unknown/**/';
             end;
         ///////////////  FW.ver
              try
              if FWbemobject.FirmwareRevision<>null then
               HDDfirmware:=trim(HDDfirmware+vartostr(FWbemobject.FirmwareRevision))+'/**/'
              else HDDfirmware:=HDDfirmware+'unknown/**/';
              except
               HDDfirmware:=HDDfirmware+'unknown/**/';
              end;
          ////////////////
            if readSmart then
            begin
             try  ///MY SMART
              HDDMYsmart:=HDDMYsmart+(SmartSelect(vartostr(FWbemobject.PNPDeviceID)))+'/**/';
             except
               HDDMYsmart:=HDDMYsmart+'Error function/**/'
             end;
           end ;
          // else HDDMYsmart:='unknown';
          ///////////////////////////////////////////
          try  //// SMART OS
             if vartostr(FWbemObject.status)<>'' then
              begin   //minLifeCurrentHDDforPCOS
              if vartostr(FWbemObject.status)<>'OK' then minLifeCurrentHDDforPCOS:='WARNING';
              HDDsmartOS:=HDDsmartOS+vartostr(FWbemObject.status)+'/**/';
              end
             else  HDDsmartOS:=HDDsmartOS+'unknown/**/';
          except
           HDDsmartOS:=HDDsmartOS+'unknown/**/';
          end;
          ////////////////////////////////////////////////
          try
          HDDtemp:=HDDtemp+inttostr(tempCurr)+'/**/'
          except
          HDDtemp:=HDDtemp+inttostr(tempCurr)+'unknown/**/';
          end;
          ///////////////////////////////////////////////
          FWbemObject:=Unassigned;
          inc(indexHDD);
          end;
           counthdd:=inttostr(i);
  Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - DiskDrive - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet)
    end;
   end;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
end;



////////////////////////////////////////////////////////////////////////////////////


procedure InventoryConfig.osconfig(var nameOS,VerOS,OSKey,invNumber,OSArchit,OSSP,InstalDate:string);
var
oEnum             : IEnumvariant;
iValue            : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT *'
+' FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Операционная система
    begin
   ////////// /////
    try
    if FWbemObject.Caption<>null then
    begin
     nameOS:=trim(string(FWbemObject.Caption));
     WriteHardware(string(FWbemObject.Caption),'OS'); /// запись в таблицу ALL_HARDWARE
    end;
    except
    nameOS:='unknown';
    end;
   ///////////////
    try
    if FWbemObject.version<>null then   VerOS:=trim(string(FWbemObject.version));
    except
    VerOS:= 'unknown';
    end;
    ///////////
    try
    if FWbemObject.SerialNumber<>null then OSKey:=trim(string(FWbemObject.SerialNumber));
    except
    OSKey:='unknown';
    end;
    //////////
    try
    if FWbemObject.Description<>null then invNumber:=trim(string(FWbemObject.Description));
    except
    invNumber:='unknown';
    end;
    //////////
    try
    if FWbemObject.OSArchitecture<>null then OSArchit:=trim(string(FWbemObject.OSArchitecture));
    except
    OSArchit:='unknown';
    end;
    try
    if FWbemObject.CSDVersion<>null then OSSP:=trim(string(FWbemObject.CSDVersion)); /// Service Pack
    except
     OSSP:= 'unknown';
    end;
    try
    if FWbemObject.InstallDate<>null then InstalDate:=trim((FWbemObject.InstallDate));
    except
    InstalDate:='01.01.1999';
    end;
    FWbemObject:=Unassigned;
    end;
    FWbemObject:=Unassigned;
 Except
  on E:Exception do
     begin
     Log_write(1,'Hardware',CurrentPC+' - OS - '+e.Message);
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet)
     end;
   end;
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
end;

///////////////////////////////////////////////////////////////////////////////
procedure InventoryConfig.NetworkInterface(var NIName,NIMAC,NISpeed,NIIP,NIMask,
NIGateway,NIDHCP,NIDNS,NIWINS,NIcount,NIcountIP,NIcountGW,NIcountDNS,
NIcountDHCP,NIcountWINS :string);
var
FWbemObjectNI   : OLEVariant;
FWbemObjectSetNI: OLEVariant;
oEnumNI : IEnumvariant;
oEnum             : IEnumvariant;
iValue            : LongWord;
mas           : variant;
res           : string;
i,z,countIP,countGW,countDNS,countDHCP,countWINS: integer;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
try
countIP:=0; //// обнуляем количество IP адресов
countGW:=0; /// обнуляем шлюзы
CountDNS:=0; /// обнуляем dns
CountDHCP:=0; /// обнуляем DHCP
CountWINS:=0; /// обнуляем WINS
z:=0;
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapter where AdapterTypeId=0','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Сетевой интерфейс
          begin
          if (FWbemObject.Description<>null) and (FWbemObject.MACAddress<>null) then
          begin
               try
               if FWbemObject.Description<>null then
                  begin
                  NIName:=trim(NIName+string(FWbemObject.Description))+'/**/';
                  WriteHardware(string(FWbemObject.Description),'Network Interfaces'); /// запись в таблицу ALL_HARDWARE
                  end
                  else  NIName:=trim(NIName+'unknown')+'/**/';
               except
                 NIName:=trim(NIName+'unknown')+'/**/';
               end;
              ///////////
               try
               if FWbemObject.MACAddress<>null then
                  NIMAC:=trim(NIMAC+string(FWbemObject.MACAddress))+'/**/'
                  else  NIMAC:=trim(NIMAC+'unknown')+'/**/';
               except
                NIMAC:=trim(NIMAC+'unknown')+'/**/';
               end;
              ///////////
              try
              if FWbemObject.speed<>null then
                  begin
                  if (FWbemObject.speed/1000/1000)<10000 then
                    NISpeed:=trim(NISpeed+inttostr(FWbemObject.speed/1000/1000))+'/**/'
                   else NISpeed:=NISpeed+'unknown/**/'
                  end
              else NISpeed:=NISpeed+'unknown/**/';
               except
                NISpeed:=NISpeed+'unknown/**/';
               end;
               inc(z); /// количество интерфейсов

                  FWbemObjectSetNI:= FWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration','WQL',wbemFlagForwardOnly);
                  oEnumNI:= IUnknown(FWbemObjectSetNI._NewEnum) as IEnumVariant;
                  countIP:=0; //// обнуляем количество IP адресов
                  countGW:=0; /// обнуляем шлюзы
                  CountDNS:=0; /// обнуляем dns
                  CountDHCP:=0; /// обнуляем DHCP
                  CountWINS:=0; /// обнуляем WINS
               try
               while oEnumNI.Next(1, FWbemObjectNI, iValue) = 0 do
                begin
                  if FWbemObjectNI.InterfaceIndex=FWbemObject.InterfaceIndex then
                     begin
                      //////////////////////////////////////// ip
                        mas:=(FWbemObjectNI.IPAddress);
                          if VarType(mas) and VarTypeMask=varVariant then
                            begin
                            for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                              begin
                              NIIP:=NIIP+(string(mas[i]))+'/**/';
                              inc(countIP);
                              end;
                            end;
                          Varclear(mas);
                       ////////////////////////////////////mask
                         mas:=(FWbemObjectNI.IPSubnet);
                         if VarType(mas) and VarTypeMask=varVariant then
                                begin
                                  for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                                  NIMask:=NIMask+(string(mas[i]))+'/**/';
                                end;
                                Varclear(mas);
                        ////////////////////////////////////Gateway
                       mas:=(FWbemObjectNI.DefaultIPGateway);
                              if VarType(mas) and VarTypeMask=varVariant then
                                 begin
                                  for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                                  begin
                                  NIGateway:=NIGateway+(string(mas[i]))+'/**/';
                                  inc(countGW);
                                  end;
                                 end;
                              Varclear(mas);
                       ///////////////////////////////////////////DHCP
                       if FWbemObjectNI.DHCPServer<>null then
                         begin
                         NIDHCP:=NIDHCP+(vartostr(FWbemObjectNI.DHCPServer))+'/**/';
                         inc(countDHCP);
                         end;

                       ///////////////////////////////////////////WINS
                       if FWbemObjectNI.WINSPrimaryServer<>null then
                         begin
                         NIWINS:=NIWINS+(vartostr(FWbemObjectNI.WINSPrimaryServer))+'/**/';
                         inc(countWINS);
                         end;

                       ///////////////////////////////////////////DNS
                           mas:=(FWbemObjectNI.DNSServerSearchOrder);
                               if VarType(mas) and VarTypeMask=varVariant then
                                 begin
                                 for i :=VarArrayLowBound(mas,1) to VarArrayHighBound(mas,1) do
                                   begin
                                   NIDNS:=NIDNS+(string(mas[i]))+'/**/';
                                   inc(countDNS);
                                   end;
                                 end;
                                 Varclear(mas);
                       ///////////////////////////////////////////////////////
                      end; ///if NIID=NIID2 then
                     FWbemObjectNI:=Unassigned;
                end;
              except
              countIP:=0; //// обнуляем количество IP адресов
              countGW:=0; /// обнуляем шлюзы
              CountDNS:=0; /// обнуляем dns
              CountDHCP:=0; /// обнуляем DHCP
              CountWINS:=0; /// обнуляем WINS
              end;

          end; /// while oEnumNI.Next

     NICountIP:=NICountIP+inttostr(countip)+'/**/'; /// количество IP адресов на интерфейсеJ
     NIcountGW:=NIcountGW+inttostr(countGW)+'/**/'; /// на каком интерфейсе шлюз
     NIcountDNS:=NIcountDNS+inttostr(countDNS)+'/**/'; ///на каких интерфейсах DNS
     NIcountDHCP:=NIcountDHCP+inttostr(countDHCP)+'/**/';/// на каких интерфейсах включен DHCP
     NIcountWINS:=NIcountWINS+inttostr(countWINS)+'/**/';/// на каких интерфейсах включен WINS
     FWbemObject:=Unassigned;
     VariantClear(FWbemObjectNI);
     oEnumNI:=nil;
     VariantClear(FWbemObjectSetNI);
     end;///  while oEnum.Next
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  NIcount:=inttostr(z); /// количество интерфейсав


Except
  on E:Exception do
    begin
    Log_write(1,'Hardware',CurrentPC+' - NetworkInterface - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWbemObjectNI);
    oEnumNI:=nil;
    VariantClear(FWbemObjectSetNI);
    end;
   end;


end;
/////////////////////////////////////////////////////////////////////
procedure InventoryConfig.ScanPrinter(var printerName,countprinter :string);
var i:byte;
oEnum             : IEnumvariant;
iValue            : LongWord;
FWbemObjectSet      : OLEVariant;
FWbemObject         : OLEVariant;
begin
i:=0;
try
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// принтер
              begin
           ////////////////////////////////////////
              try
                if FWbemObject.Caption<>Null then
                    begin
                     WriteHardware(string(FWbemObject.Caption),'Printer'); /// запись в таблицу ALL_HARDWARE
                     printerName:=printerName+trim(string(FWbemObject.Caption))+'/**/';
                     inc(i);
                    end;
              except
              printerName:=printerName+'unknown/**/';
              end;
             FWbemObject:=Unassigned;
             end;
FWbemObject:=Unassigned;
countprinter:=inttostr(i);
  Except
  on E:Exception do
    begin
   Log_write(1,'Hardware',CurrentPC+' - Printer - '+e.Message);
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    countprinter:=inttostr(i);
    end;
   end;
   oEnum:=nil;
   VariantClear(FWbemObject);
   VariantClear(FWbemObjectSet);
   if printerName='' then  printerName:= 'unknown/**/';

end;
//////////////////////////////////////////////////////////////////////////


function convertdate(s:string;n:string):string;
var
d:tdatetime;
begin
  try
 if length(s)=25 then
  if n='date' then
  begin   /////  день           месяц         год
    result:=copy(s,7,2)+'.'+copy(s,5,2)+'.'+copy(s,1,4);
    d:=strtodate(result);
  end;
  Except
  on E:Exception do
    begin
    result:='01/01/1999';
    end;
   end;

try
  if n='time' then
  begin
  result:=copy(s,9,2)+':'+copy(s,11,2)+':'+copy(s,11,2);
  d:=strtotime(result);
  end;
Except
  on E:Exception do
    begin
    result:='00:00:00';
    end;
   end;
  end;

function InventoryConfig.connectDB:boolean;
begin
try
ConnectionThRead:=TFDConnection.Create(nil);
ConnectionThRead.DriverName:='FB';
ConnectionThRead.Params.database:=databaseName; // расположение базы данных
ConnectionThRead.Params.Add('server='+databaseServer);
ConnectionThRead.Params.Add('port='+databasePort);
ConnectionThRead.Params.Add('protocol='+databaseProtocol);  //TCPIP или local
ConnectionThRead.Params.Add('CharacterSet=UTF8');
ConnectionThRead.Params.add('sqlDialect=3');
ConnectionThRead.Params.DriverID:=databaseDriverID;
ConnectionThRead.Params.UserName:=databaseUserName;
ConnectionThRead.Params.Password:=databasePassword;
ConnectionThRead.Connected:=true;
ConnectionThRead.LoginPrompt:= false;  /// отображение диалога user password
TransactionWrite:= TFDTransaction.Create(nil);
TransactionWrite.Connection:=ConnectionThRead;
TransactionWrite.Options.Isolation:=xiReadCommitted; ///xiReadCommitted; //xiSnapshot;
FDQueryWrite:=TFDQuery.Create(nil);
FDQueryWrite.Transaction:=TransactionWrite;
FDQueryWrite.Connection:=ConnectionThRead;
OleInitialize(nil);
result:=true;
except on E: Exception do
begin
result:=false;
Log_write(2,'Hardware',' Ошибка подключения к базе данных '+e.Message);
end;
end;
end;

function InventoryConfig.DisconnectDB:boolean;
begin
try
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize();
FDQueryWrite.Close;
if Assigned(FDQueryWrite) then FDQueryWrite.Free;
if TransactionWrite.Active then TransactionWrite.Commit;
if Assigned(TransactionWrite) then TransactionWrite.Free;
frmDomainInfo.statusbarInv.Panels[3].Text:='Инвентаризация завершена';
if Assigned(ConnectionThRead) then //закрываем соединение
  begin
  ConnectionThRead.Connected:=false;
  ConnectionThRead.Close;
  ConnectionThRead.Free;
  end;
result:=true;
except on E: Exception do
begin
result:=false;
Log_write(2,'Hardware',' Ошибка отключения от БД '+e.Message);
end;
end;
end;

function InventoryConfig.ConnectWMI(NamePC,User,Pass:string): Boolean;
begin
 try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService   := FSWbemLocator.ConnectServer(NamePC, 'root\CIMV2', User, Pass,'','',128);
    FWMIService.Security_.impersonationlevel:=3;
    FWMIService.Security_.authenticationLevel := 6;
    Result:=True;
 except
    begin
    result:=false;
    Log_write(1,'Hardware',NamePC+' - RPC сервервер не доступен');
    end;
 end;
end;


procedure InventoryConfig.Execute;
var
card,memcard,countcard,resIP:string;
hddType,hddName,hddSize,hddsn,counthdd,InterfaceType,HDDfirmware,HDDMySMART,HddSMARTOs,HDDTemp :string;
ProcName,ProcCore,ProcSpeed,ProcArh,ProcSoc,ProcLogProc,CountProc:string;
montrName,montrSN,montrManuf:String;
nameOS,VerOS,OSKey,invNumber,OSArchit,OSSP,InstalDate,UserNamePC:string;
NIName,NIMAC,NISpeed,NIIP,NIMask,NIGateway,NIDHCP,NIDNS,NIWINS,NIcount,NICountIP,
NIcountGW,NIcountDNS,NIcountDHCP,NIcountWINS :string;
CountMem,SummSizeMem,SizeMem,memSpeed:string;
screenHW,NameMon,NamufactureMon,countMonitor,DPI :string;
CDRomName,CDRomCount,CDRomType:String;
SoundName,SoundCount:String;
printerName,countprinter :string;
BIOSCount, BIOSDescription,BIOSSN,SMBIOSBIOSVersionMM,BIOSStatus,BiosPrimary,BIOSManufac,Characteristics:string;
resinv:string;
i:integer;
CurrentListPC:TstringList;
CountPCOK:integer;
invOK:string;
SetInI:Tmeminifile;
function nullstring(s:string):string;
begin
  if s='' then result:=' '  /// if s='' then result:='unknown/**/'
  else result:=s;
  s:='';
end;

begin
try
if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini')  then
begin
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
readSmart:=SetInI.ReadBool('SMART','read',true);
SetInI.Free;
end;
////////////////////////////////////////////////
if not connectDB then
begin
Log_write(2,'Hardware',' Не могу подключится к БД, продолжение инвентаризации не возможно');
frmDomainInfo.statusbarInv.Panels[3].Text:='Отсутствует подключение к БД';
InventConf:=false;
SolveExitInvConf:=true; /// признак завершения потока
exit;
end;
Log_write(0,'Hardware','Запуск инвентаризации');
CurrentListPC:=TstringList.Create;
CurrentListPC.Text:=ListPCConf.Text;
ListPCConf.Free;
InvUser:=MyUser;
InvPass:=MyPasswd;
CountPCOK:=0;
resinv:='';
SolveExitInvConf:=false;
for I := 0 to CurrentListPC.Count-1 do
begin
        if not inventConf then  break;  /// если остановили инвентаризацию, выход из цикла
         if SelectedPing(resIP,CurrentListPC[i],pingtimeout) and(get135Port(CurrentListPC[i])) then
         begin
          if ConnectWMI(CurrentListPC[i],InvUser, InvPass) then
            begin
            try
            CurrentPC:=CurrentListPC[i];
//////////////////////////////////////////////////////
            card:='';
            memcard:='';
            countcard:='';
            videocard(card,memcard,countcard);
            card:=nullstring(card);
            memcard:=nullstring(memcard);
            //countcard:=nullstring(countcard);
/////////////////////////////////////////////////
            hddName:='';
            hddType:='';
            hddSize:='';
            hddsn:='';
            counthdd:='';
            InterfaceType:='';
            HDDfirmware:='';
            HDDMySMART:='';
            HddSMARTOs:='';
            HDDTemp:='';
            minLifeCurrentHDDforPC:=100; // для находжения наименьшего занчения жизни диска на текущем компьютере (HDDMySMART)
            minLifeCurrentHDDforPCOS:='OK';  /// наихудший статус диска Оценка опер.сист
            scanhdd(hddType,hddName,hddSize,hddsn,counthdd,InterfaceType,HDDfirmware,HDDMySMART,HddSMARTOs,HDDTemp);
            hddName:=nullstring(hddName);
            hddType:=nullstring(hddType);
            hddSize:=nullstring(hddSize);
            hddsn:=nullstring(hddsn);
           //counthdd:=nullstring(counthdd);
            InterfaceType:=nullstring(InterfaceType);
            HDDfirmware:=nullstring(HDDfirmware);

/////////////////////////////////////////////////////
            ProcName:='';
            ProcCore:='';
            ProcSpeed:='';
            ProcArh:='';
            ProcSoc:='';
            ProcLogProc:='';
            procConf(ProcName,ProcCore,ProcSpeed,ProcArh,ProcSoc,ProcLogProc,CountProc);
            ProcName:=nullstring(ProcName);
            //ProcCore:=nullstring(ProcCore);
            ProcSpeed:=nullstring(ProcSpeed);
            ProcArh:=nullstring(ProcArh);
            ProcSoc:=nullstring(ProcSoc);
            //ProcLogProc:=nullstring(ProcLogProc);
/////////////////////////////////////////////////////////
            montrName:='';
            montrSN:='';
            montrManuf:='';
            monterboard(montrName,montrSN,montrManuf);
            montrName:=nullstring(montrName);
            montrSN:=nullstring(montrSN);
            montrManuf:=nullstring(montrManuf);
//////////////////////////////////////////////////////////
            nameOS:='';
            VerOS:='';
            OSKey:='';
            invNumber:='';
            OSArchit:='';
            OSSP:='';
            InstalDate:='';
            osconfig(nameOS,VerOS,OSKey,invNumber,OSArchit,OSSP,InstalDate);
            nameOS:=nullstring(nameOS);
            VerOS:=nullstring(VerOS);
            VerWin:=VerOS;
            OSKey:=nullstring(OSKey);
            //invNumber:=nullstring(invNumber);
            OSArchit:=nullstring(OSArchit);
            OSSP:=nullstring(OSSP);
            //InstalDate:=nullstring(InstalDate);
////////////////////////////////////////////////////////////////////////
            UserNamePC:='';
            Nameuser (UserNamePC);
            UserNamePC:=nullstring(UserNamePC);
////////////////////////////////////////////////////////////////////////////
            NIName:='';
            NIMAC:='';
            NISpeed:='';
            NIIP:='';
            NIMask:='';
            NIGateway:='';
            NIDHCP:='';
            NIDNS:='';
            NIWINS:='';
            NIcount:='';
            NICountIP:='';
            NIcountGW:='';
            NIcountDNS:='';
            NIcountDHCP:='';
            NIcountWINS:='';
            NetworkInterface(NIName,NIMAC,NISpeed,NIIP,NIMask,NIGateway,NIDHCP,
            NIDNS,NIWINS,NIcount,NICountIP,NIcountGW,NIcountDNS,NIcountDHCP,NIcountWINS);
            NIName:=nullstring(NIName);
            NIMAC:=nullstring(NIMAC);
            NISpeed:=nullstring(NISpeed);
            NIIP:=nullstring(NIIP);
            NIMask:=nullstring(NIMask);
            NIGateway:=nullstring(NIGateway);
            NIDHCP:=nullstring(NIDHCP);
            NIDNS:=nullstring(NIDNS);
            NIWINS:=nullstring(NIWINS);
           // NIcount:=nullstring(NIcount);
           // NICountIP:=nullstring(NICountIP);
           // NIcountGW:=nullstring(NIcountGW);
           // NIcountDNS:=nullstring(NIcountDNS);
           // NIcountDHCP:=nullstring(NIcountDHCP);
           // NIcountWINS:=nullstring(NIcountWINS);
////////////////////////////////////////////////////////////////////////
            CountMem:='';
            SummSizeMem:='';
            SizeMem:='';
            memSpeed:='';
            memorySpeed(memSpeed);
            memory(CountMem,SummSizeMem,SizeMem);
            // CountMem:=nullstring(CountMem);
            SummSizeMem:=nullstring(SummSizeMem);
            SizeMem:=nullstring(SizeMem);
            memSpeed:=nullstring(memSpeed);
/////////////////////////////////////////////////////////////////////////////
            screenHW:='';
            NameMon:='';
            NamufactureMon:='';
            countMonitor:='';
            DPI:='';
            DesktopMon(screenHW,NameMon,NamufactureMon,countMonitor,DPI);
            screenHW:=nullstring(screenHW);
            NameMon:=nullstring(NameMon);
            NamufactureMon:=nullstring(NamufactureMon);
            //countMonitor:=nullstring(countMonitor);
            DPI:=nullstring(DPI);
////////////////////////////////////////////////////////////////////////////
            CDRomName:='';
            CDRomCount:='';
            CDRomType:='';
            CDDVDRom (CDRomName,CDRomCount,CDRomType);
            CDRomName:=nullstring(CDRomName);
            //CDRomCount:=nullstring(CDRomCount);
            CDRomType:=nullstring(CDRomType);
//////////////////////////////////////////////////////////////////////////
            SoundName:='';
            SoundCount:='';
            SoundDevice (SoundName,SoundCount);
            SoundName:=nullstring(SoundName);
            //SoundCount:=nullstring(SoundCount);
////////////////////////////////////////////////////////////////////////////
            printerName:='';
            countprinter:='';
            ScanPrinter(printerName,countprinter);
            printerName:=nullstring(printerName);
            ///countprinter:=nullstring(countprinter);
///////////////////////////////////////////////////////////////////////////// BIOSCHARACTERISTICS
            BIOSDescription:='';
            BIOSCount:='';
            BIOSSN:='';
            SMBIOSBIOSVersionMM:='';
            BIOSStatus:='';
            BiosPrimary:='';
            BIOSManufac:='';
            Characteristics:='';
            BIOSMB(BIOSDescription, BIOSCount,BIOSSN,SMBIOSBIOSVersionMM,BIOSStatus
            ,BiosPrimary,BIOSManufac,Characteristics);
/// /////////////////////////////////////////////////////////////////////////
            TransactionWrite.StartTransaction;
            FDQueryWrite.SQL.Clear;
            FDQueryWrite.SQL.Text:=
            'update or insert into CONFIG_PC (PC_NAME,MONTERBOARD,MONTERBOARD_SN,'
            +'MONTERBOARD_MANUFACTURE,PROCESSOR,PROCESSOR_CORE,PROCESSOR_LOGPROC,'
            +'PROCESSOR_SPEED,PROCESSOR_ARCH,PROCESSOR_SOKET,COUNT_PROC,MEMORY_TYPE,'
            +'MEMORY_SIZE,SUMM_MEM_SIZE,COUNT_MEM,VIDEOCARD,VIDEOCARD_MEM,'
            +'COUNT_VIDEOCARD,HDD_NAME,HDD_TYPE,HDD_SIZE,HDD_SN,COUNT_HDD,OS_NAME,OS_VER,'
            +'OS_KEY,OS_TYPE,INV_NUMBER,OS_SP,OS_DATEINSTALL,NETWORKINTERFACE,NETWORK_MAC,'
            +'NETWORK_SPEED,NETWORK_IP,NETWORK_MASK,NETWORK_GATEWAY,NETWORK_DHCP,'
            +'NETWORK_DNS,NETWORK_WINS,COUNT_NETWORK,DATE_INV,RESULT_INVENT,USER_NAME,'
            +'NETWORK_COUNT_IP,NETWORK_COUNT_GATEWAY,NETWORK_COUNT_DNS,NETWORK_COUNT_DHCP'
            +',NETWORK_COUNT_WINS,OS_TIMEINSTALL,MONITOR_HW,MONITOR_NAME,MONITOR_MANUF'
            +',MONITOR_DPI,MONITOR_COUNT,TIME_INV,HDD_INTERFACETYPE,HDD_FIRMWARE'
            +',DVDROM_NAME,DVDROM_TYPE,DVDROM_COUNT,SOUND_NAME,SOUND_COUNT,ANSWER_MAC'
            +',PRINTER_NAME,PRINTER_COUNT,HDD_MY_SMART,HDD_SMART_OS,HDD_CUR_TEMP,BIOS_DESCRIPTION'
            +',BIOS_COUNT,BIOSSN,SMBIOSBIOSVERSION,BIOSSTATUS,BIOSPRIMARY,BIOSMANUFAC,BIOSCHARACTERISTICS)'
            +'VALUES(:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8,:p9,:p10,:p11,:p12,:p13,:p14'
            +',:p15,:p16,:p17,:p18,:p19,:p20,:p21,:p22,:p23,:p24,:p25,:p26,:p27'
            +',:p28,:p29,:p30,:p31,:p32,:p33,:p34,:p35,:p36,:p37,:p38,:p39,:p40'
            +',:p41,:p42,:p43,:p44,:p45,:p46,:p47,:p48,:p49,:p50,:p51,:p52,:p53'
            +',:p54,:p55,:p56,:p57,:p58,:p59,:p60,:p61,:p62,:p63,:p64,:p65,:p66,:p67,:p68'
            +',:p69,:p70,:p71,:p72,:p73,:p74,:p75,:p76)'
            +' MATCHING ('+InvWarning+')'; // в переманной строка с полями таблицы для определения нарушения конфигурации
            FDQueryWrite.params.ParamByName('p1').AsString:=''+CurrentListPC[i]+'';
            FDQueryWrite.params.ParamByName('p2').AsString:=''+leftstr(montrName,250)+'';
            FDQueryWrite.params.ParamByName('p3').AsString:=''+leftstr(montrSN,250)+'';
            FDQueryWrite.params.ParamByName('p4').AsString:=''+leftstr(montrManuf,250)+'';
            FDQueryWrite.params.ParamByName('p5').AsString:=''+leftstr(ProcName,250)+'';
            FDQueryWrite.params.ParamByName('p6').AsString:=''+leftstr(ProcCore,100)+'';
            FDQueryWrite.params.ParamByName('p7').AsString:=''+leftstr(ProcLogProc,250)+'';
            FDQueryWrite.params.ParamByName('p8').AsString:=''+leftstr(ProcSpeed,100)+'';
            FDQueryWrite.params.ParamByName('p9').AsString:=''+leftstr(ProcArh,100)+'';
            FDQueryWrite.params.ParamByName('p10').AsString:=''+leftstr(ProcSoc,250)+'';
            FDQueryWrite.params.ParamByName('p11').AsString:=''+CountProc+'';
            FDQueryWrite.params.ParamByName('p12').AsString:=''+leftstr(memSpeed,300)+'';
            FDQueryWrite.params.ParamByName('p13').AsString:=''+leftstr(SizeMem,300)+'';
            FDQueryWrite.params.ParamByName('p14').AsString:=''+LEFTSTR(SummSizeMem,100)+'';
            FDQueryWrite.params.ParamByName('p15').AsString:=''+CountMem+'';
            FDQueryWrite.params.ParamByName('p16').AsString:=''+leftstr(card,250)+'';
            FDQueryWrite.params.ParamByName('p17').AsString:=''+leftstr(memcard,250)+'';
            FDQueryWrite.params.ParamByName('p18').AsString:=''+(countcard)+'';
            FDQueryWrite.params.ParamByName('p19').AsString:=''+leftstr(hddName,1000)+'';
            FDQueryWrite.params.ParamByName('p20').AsString:=''+leftstr(hddType,1000)+'';
            FDQueryWrite.params.ParamByName('p21').AsString:=''+leftstr(hddSize,1000)+'';
            FDQueryWrite.params.ParamByName('p22').AsString:=''+leftstr(hddsn,1000)+'';
            FDQueryWrite.params.ParamByName('p23').AsString:=''+(counthdd)+'';
            FDQueryWrite.params.ParamByName('p24').AsString:=''+leftstr(nameOS,200)+'';
            FDQueryWrite.params.ParamByName('p25').AsString:=''+leftstr(VerOS,200)+'';
            FDQueryWrite.params.ParamByName('p26').AsString:=''+leftstr(OSKey,200)+'';
            FDQueryWrite.params.ParamByName('p27').AsString:=''+leftstr(OSArchit,200)+'';
            FDQueryWrite.params.ParamByName('p28').AsString:=''+leftstr(invNumber,200)+'';
            FDQueryWrite.params.ParamByName('p29').AsString:=''+leftstr(OSSP,200)+'';
            FDQueryWrite.params.ParamByName('p30').AsString:=''+convertdate(InstalDate,'date')+'';
            FDQueryWrite.params.ParamByName('p31').AsString:=''+leftstr(NIName,300)+'';
            FDQueryWrite.params.ParamByName('p32').AsString:=''+leftstr(NIMAC,300)+'';
            FDQueryWrite.params.ParamByName('p33').AsString:=''+leftstr(NISpeed,300)+'';
            FDQueryWrite.params.ParamByName('p34').AsString:=''+leftstr(NIIP,300)+'';
            FDQueryWrite.params.ParamByName('p35').AsString:=''+leftstr(NIMask,300)+'';
            FDQueryWrite.params.ParamByName('p36').AsString:=''+leftstr(NIGateway,300)+'';
            FDQueryWrite.params.ParamByName('p37').AsString:=''+leftstr(NIDHCP,300)+'';
            FDQueryWrite.params.ParamByName('p38').AsString:=''+leftstr(NIDNS,300)+'';
            FDQueryWrite.params.ParamByName('p39').AsString:=''+leftstr(NIWINS,300)+'';
            FDQueryWrite.params.ParamByName('p40').AsString:=''+(NIcount)+'';
            FDQueryWrite.params.ParamByName('p41').AsString:=''+datetostr(date)+'';
            FDQueryWrite.params.ParamByName('p42').AsString:='OK';
            FDQueryWrite.params.ParamByName('p43').AsString:=''+leftstr(UserNamePC,100)+'';
            FDQueryWrite.params.ParamByName('p44').AsString:=''+leftstr(NiCountIP,100)+'';
            FDQueryWrite.params.ParamByName('p45').AsString:=''+leftstr(NiCountGW,100)+'';
            FDQueryWrite.params.ParamByName('p46').AsString:=''+leftstr(NiCountDNS,100)+'';
            FDQueryWrite.params.ParamByName('p47').AsString:=''+leftstr(NiCountDHCP,100)+'';
            FDQueryWrite.params.ParamByName('p48').AsString:=''+leftstr(NiCountWINS,100)+'';
            FDQueryWrite.params.ParamByName('p49').AsString:=''+convertdate(InstalDate,'time')+'';
            FDQueryWrite.params.ParamByName('p50').AsString:=''+leftstr(screenHW,200)+'';
            FDQueryWrite.params.ParamByName('p51').AsString:=''+leftstr(NameMon,500)+'';
            FDQueryWrite.params.ParamByName('p52').AsString:=''+leftstr(NamufactureMon,500)+'';
            FDQueryWrite.params.ParamByName('p53').AsString:=''+leftstr(DPI,200)+'';
            FDQueryWrite.params.ParamByName('p54').AsString:=''+leftstr(countMonitor,100)+'';
            FDQueryWrite.params.ParamByName('p55').AsString:=''+timetostr(time)+'';
            FDQueryWrite.params.ParamByName('p56').AsString:=''+leftstr(InterfaceType,1000)+'';
            FDQueryWrite.params.ParamByName('p57').AsString:=''+leftstr(HDDfirmware,1000)+'';
            FDQueryWrite.params.ParamByName('p58').AsString:=''+leftstr(CDRomName,1000)+'';
            FDQueryWrite.params.ParamByName('p59').AsString:=''+leftstr(CDRomType,1000)+'';
            FDQueryWrite.params.ParamByName('p60').AsString:=''+leftstr(CDRomCount,5)+'';
            FDQueryWrite.params.ParamByName('p61').AsString:=''+leftstr(SoundName,1500)+'';
            FDQueryWrite.params.ParamByName('p62').AsString:=''+leftstr(SoundCount,5)+'';
            FDQueryWrite.params.ParamByName('p63').AsString:=''+leftstr(AnswerMAC,30)+'';
            FDQueryWrite.params.ParamByName('p64').AsString:=''+leftstr(Printername,500)+'';
            FDQueryWrite.params.ParamByName('p65').AsString:=''+(countprinter)+'';
            FDQueryWrite.params.ParamByName('p66').AsString:=''+leftstr(HDDMySMART,1000)+'';
            FDQueryWrite.params.ParamByName('p67').AsString:=''+leftstr(HddSMARTOs,1000)+'';
            FDQueryWrite.params.ParamByName('p68').AsString:=''+leftstr(HDDTemp,1000)+'';
            ///////////////////////////////////////////////////////////////////////
            ///  BIOSDescription,BIOSCount,BIOSSN,SMBIOSBIOSVersionMM,BIOSVersion,BiosPrimary,BIOSManufac
            //// ver 3.0
            FDQueryWrite.params.ParamByName('p69').AsString:=''+leftstr(BIOSDescription,500)+'';
            FDQueryWrite.params.ParamByName('p70').AsString:=''+leftstr(BIOSCount,3)+'';
            FDQueryWrite.params.ParamByName('p71').AsString:=''+leftstr(BIOSSN,200)+'';
            FDQueryWrite.params.ParamByName('p72').AsString:=''+leftstr(SMBIOSBIOSVersionMM,200)+'';
            FDQueryWrite.params.ParamByName('p73').AsString:=''+leftstr(BIOSStatus,50)+'';
            FDQueryWrite.params.ParamByName('p74').AsString:=''+leftstr(BiosPrimary,10)+'';
            FDQueryWrite.params.ParamByName('p75').AsString:=''+leftstr(BIOSManufac,300)+'';
            FDQueryWrite.params.ParamByName('p76').AsString:=''+leftstr(Characteristics,500)+'';
            ///////////////////////////////////////////////////////////////////////////////////
            invOK:='OK';
            FDQueryWrite.ExecSQL;
            TransactionWrite.commit;
            FDQueryWrite.close;
            try
            TransactionWrite.StartTransaction;
            FDQueryWrite.SQL.Clear;                                                              /// ERROR_INV Зачищаем ошибки прошлой инвентаризации
            FDQueryWrite.SQL.Text:='update or insert into '                                      ///minLifeCurrentHDDforPCOS
            +' MAIN_PC (PC_NAME,DATE_INV,INV_NUMBER,PC_OS,RESULT_INV,ERROR_INV,ANSWER_MAC,HDD_MY_SMART,HDD_SMART_OS,CUR_USER_NAME) VALUES'
            +' ('''+CurrentListPC[i]+''','''+datetostr(date)+''','''+leftstr(invNumber,200)+''','''+leftstr(nameOS,200)+''','''+invOK+''','''+resinv+''','''+leftstr(AnswerMAC,30)+''','''+inttostr(minLifeCurrentHDDforPC)+''','''+leftstr(minLifeCurrentHDDforPCOS,20)+''','''+leftstr(UserNamePC,100)+''')'
            +' MATCHING (PC_NAME)';
            FDQueryWrite.ExecSQL;
            TransactionWrite.commit;
            FDQueryWrite.Close;
            except on E: Exception do
            begin
            Log_write(2,'Hardware',CurrentPC+' - Ошибка записи в базу данных: '+e.Message);
            TransactionWrite.commit;
            end;
            end;

            inc(CountPCOK);
            frmDomainInfo.statusbarInv.Panels[3].Text:=CurrentPC+' - OK';
             Except
              on E:Exception do
                begin
                if TransactionWrite.Active then TransactionWrite.Rollback;
                Log_write(2,'Hardware',CurrentPC+' - Ошибка инвентаризации: '+e.Message);
                TransactionWrite.StartTransaction;
                FDQueryWrite.SQL.Clear;
                FDQueryWrite.SQL.Text:='update or insert into MAIN_PC (PC_NAME,DATE_INV,ERROR_INV) VALUES'
                +' ('''+CurrentListPC[i]+''','''+datetostr(date)+''','''+e.Message+''') MATCHING (PC_NAME)';
                FDQueryWrite.ExecSQL;
                TransactionWrite.Commit;
                FDQueryWrite.Close;
                VariantClear(FWMIService);
                VariantClear(FSWbemLocator);
                end
              end;
            //////////////////////////////////////////////
          end; // если доступен WMI
         end; // если комп в сети
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
frmDomainInfo.statusbarInv.Panels[4].Text:=inttostr(CurrentListPC.Count)+'/'+inttostr(i+1)+'/'+inttostr(CountPCOK)+' - OK';
if not inventConf then  break;  /// если остановили инвентаризацию, выход из цикла
end; // цикл по компьютерам

 Except
  on E:Exception do
    begin
    Log_write(3,'Hardware',CurrentPC+'Общая ошибка инвентаризации: '+e.Message);
    frmDomainInfo.statusbarInv.Panels[3].Text:=CurrentPC+' - Общая ошибка';
    end;
   end;
if Assigned(CurrentListPC) then CurrentListPC.Free;
DisconnectDB; // закрываем все нах соединения
Log_write(0,'Hardware',' Инвентаризация оборудования завершена');
InventConf:=false; /// инвентаризация закончена
SolveExitInvConf:=true; /// разрешить закрыть программу или запустить инвентаризацию повторно
end;

end.
