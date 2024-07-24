unit WOLThread;

interface

uses
WinTypes, Messages, SysUtils, Classes, IdBaseComponent,
  IdComponent, IdUDPBase, IdUDPClient,idGlobal;



type
  WOL = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;
var
MACList:TstringList;
implementation
uses umain;

procedure MyWOL(aMacAddress: string);
var
  i, j: Byte;
  lBuffer: tbytes;
  lUDPClient: TIDUDPClient;
begin
frmDomainInfo.Memo1.Lines.Add('Отправляю WOL на адрес '+aMacAddress);
    aMacAddress := StringReplace(uppercase(aMacAddress), '-', '', [rfReplaceAll]);
    aMacAddress := StringReplace(aMacAddress, ':', '', [rfReplaceAll]);
   // frmDomainInfo.Memo1.Lines.Add('2');
  try
    SetLength(lbuffer,117);
    for i := 1 to 6 do
    begin
      lBuffer[i] :=StrToIntDef('$' + aMacAddress[(i * 2) - 1] + aMacAddress[i * 2],0);
    end;
    //frmDomainInfo.Memo1.Lines.Add('3');
    lBuffer[7] := $00;
    lBuffer[8] := $74;
    lBuffer[9] := $FF;
    lBuffer[10] := $FF;
    lBuffer[11] := $FF;
    lBuffer[12] := $FF;
    lBuffer[13] := $FF;
    lBuffer[14] := $FF;
    for j := 1 to 16 do
    begin
      for i := 1 to 6 do
      begin
        lBuffer[15 + (j - 1) * 6 + (i - 1)] := lBuffer[i];
      end;
    end;
    lBuffer[116] := $00;
    lBuffer[115] := $40;
    lBuffer[114] := $90;
    lBuffer[113] := $90;
    lBuffer[112] := $00;
    lBuffer[111] := $40;
    try
      lUDPClient := TIdUDPClient.Create(nil);
      lUDPClient.BroadcastEnabled := true;
      lUDPClient.Host := IpBroadCast;   ///// траблы с указанием IP адреса (Необходимо определять подсеть в какую слать пакет, УТОЧНИТЬ)
      lUDPClient.Port := 9;
      lUDPClient.SendBuffer(lUDPClient.Host,lUDPClient.Port ,tidbytes(lBuffer));
      frmDomainInfo.Memo1.Lines.Add('Магический пакет отправлен');
    finally
      lUDPClient.Free;
    end;
  except
   on E: Exception do frmDomainInfo.Memo1.Lines.Add('Error: '+E.Message+' WakeUp: '+aMacAddress);
  end;
end;



procedure WOL.Execute;
var
i,z:integer;
begin
MACList:=TStringList.Create;
for I := 0 to ListMACAddress.Count-1 do
begin
MACList.add(ListMACAddress[i]);
end;
ListMACAddress.Free;

for I := 0 to MACList.Count-1 do
begin
MyWol(MACList[i]);
  for z := 0 to frmDomaininfo.listview8.Items.Count-1 do
  begin
  if frmDomaininfo.listview8.Items[z].SubItems[2]=MACList[i] then
    begin
    frmDomaininfo.ListView8.Items[z].SubItemImages[1]:=1;
    frmDomainInfo.ListView8.Items[z].SubItems[1]:='WOL пакет отправлен';
    break;
    end;
  end;
end;
MACList.Free;
end;


 end.
