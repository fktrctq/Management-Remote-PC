unit FormCopyFF;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,ShellAPI,ClipBrd,ShlObj,IdIcmpClient,
  Vcl.ExtCtrls;

const
  netapi32lib = 'netapi32.dll';
  NERR_Success = NO_ERROR;
  wbemFlagForwardOnly = $00000020;
  wbemFlagReturnImmediately=16;

type
  TForm11 = class(TForm)
    TaskDialog1: TTaskDialog;
    procedure CancelCopy(Sender: TObject);
    procedure TaskDialog1VerificationClicked(Sender: TObject);
    procedure FormFFClose(Sender: TObject; var Action: TCloseAction);

  private
  //  procedure SetPosition(APos: Int64);
  public
  function CopyFileFolder(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean; /// ��� ������� ����������� � ������
  function FunctionCopyFF(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):integer; // ����������� �� � ������
  procedure Clipboard_DataSend(const DataPaths: TStrings;const MoveType: Integer);
  function FunctionStartThreadCopyFileEx(FSource,FDest,username,passwd:string;
CancelCopyFF:boolean;CountNewForm:integer;ListPC:TstringList):boolean; /// ������� ��� ������� CopyFileEx ����������� � ������
  function CopyFileFolderListPC(FSource,FDest,userPC,Paswd,StrListPC:string; // ������� ������� ����������� � ������ ��� ������ �����������
  FindCreateFolder,CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
  end;
type
 TstrForCopy = record
    FSource :String;   /// �������� ����
    FDest:string;      //���� ����������
    ListPC:string;
    UserName:string;
    PassWd:string;
    PathCreate:bool;  // �������� � �������� ����������
    CancelCopyFF:boolean;   // �������� ��� ��� �������� �����������
    OwnerForm:Tform;        // ������������ ����� �������� ����, ����� �� ���������
    ProgBar:TprogressBar;      // �������� ���, ����� �������� ������ �����������
    NumCount:integer;
    TypeOperation:integer;     // ��� �������, ����������, �������� � �.�. ������ ��� ��������� TSHFileOpStruct
  end;
type
SHFILEOPSTRUCTA = packed record
    Wnd: HWND;
    wFunc: UINT;
    pFrom: PAnsiChar;
    pTo: PAnsiChar;
    fFlags: FILEOP_FLAGS;
    fAnyOperationsAborted: BOOL;
    hNameMappings: Pointer;
    lpszProgressTitle: PAnsiChar; { ������������ ������ ��� ������������� ����� FOF_SIMPLEPROGRESS }
  end;
  function ThreadCopyFileEx(paramCopyFF:pointer):boolean; // ����������� � ������ c ������� �������� ����
  function ThreadCopyFF(paramCopyFF:pointer):boolean;
  //function NetApiBufferFree(Buffer: Pointer): DWORD; stdcall;
  //  external netapi32lib;
var
  Form11: TForm11;
  RunCopyFF         :array [0..2000] of TstrForCopy;///��������� ��� ����������� � ������
  treadID: array [0..2000] of longword;
  res : array [0..2000] of integer;
  CopyCancelFF: array [0..2000] of boolean;

ThreadVar
PointForCopyFF: ^TstrForCopy;

implementation
uses umain;

{$R *.dfm}
procedure TForm11.FormFFClose(Sender: TObject; var Action: TCloseAction); // ����������� ����� ��� ��������
begin
(sender as Tform).Release;
end;



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
Myidicmpclient:TIdIcmpClient;
begin
try
result:=false;
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=pingtimeout;
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
if Assigned(MyIdIcmpClient) then freeandnil(MyIdIcmpClient);
end;

function Shell_Str(Strs: TStrings): string;
{
  ������� ����������� TStirngs � ����-������ ��� Shell
  ��� ����-������ ��� ������ ������, ��� ������ ��������� ������ #0
  � ��� ����-������ ������ ������������� #0#0
}
var
  i: Integer;
begin
  for i := 0 to Strs.Count - 1 do
    Result := Result + Strs.Strings[i] + #0;
  Result := Result + #0;
end;

procedure TForm11.CancelCopy(Sender: TObject);
var
NumThread:integer;
begin
if (sender as Tbutton).owner is Tform then
begin
  NumThread:=((sender as Tbutton).owner as Tform).tag;
 // RunCopyFF[NumThread].CancelCopyFF:=true;
  //CopyCancelFF[NumThread]:=true;
  //CloseHandle(res[NumThread]);
 //((sender as Tbutton).owner as Tform).Release;
end;
end;

function FindAdncreateDir(path,NamePC:string):bool;// �������� � �������� ����������
begin
   try
    if not DirectoryExists(ExtractFileDir('\\'+NamePC+'\'+path)) then //���� ��� ��������
    begin
     if not ForceDirectories('\\'+NamePC+'\'+ExtractFileDir(path)) then/// ������� ���� ���� �� �����
     begin
     finditem('�������� ���������� ' +ExtractFileDir(path)+' : ��������� ������ ',NamePC,2);
     result:=false;
     end
     else result:=true; // ���������� �������
    end
    else result:=true; // ���������� ����
  except on E: Exception do
     begin
     frmdomaininfo.Memo1.Lines.Add(NamePC+' : ������ �������� ���������� - '+e.Message);
     finditem('�������� ���������� ' +ExtractFileDir(path)+' :' +e.Message,NamePC,2);
     result:=false;
     end;
   end;
end;

///////////////////////////////////////////////////////////
function ThreadCopyFileEx(paramCopyFF:pointer):boolean; // ����������� � ������ c ������� �������� ����
 var
 i:integer;
 htoken:THandle;
 step:string;
 listpc:tstringlist;
 function CopyCallBack (  /// ������� ��������� ��� CopyFileEx
 TotalFileSize,
 TotalBytesTransferred,
 StreamSize,
 StreamBytesTransferred: Int64;
 StreamNumber,
 CallBackReasom: DWORD;
 SrcFile,
 DestFile: THandle): DWORD; stdcall;
begin
   PointForCopyFF.ProgBar.Position:=TotalBytesTransferred;
end;
begin
try
listpc:=tstringlist.Create;
PointForCopyFF:=paramCopyFF;
listpc.CommaText:=PointForCopyFF.ListPC;
for I := 0 to ListPC.Count-1 do
if ping(ListPC[i]) then
BEGIN
 try
 ////////////////////////////////////// ����������� ������������ �� ��������� �����
   try
   if not (LogonUserA (PAnsiChar(PointForCopyFF.UserName), PAnsiChar (ListPC[i]),  // ������� ������� �� ���� � ����
   PAnsiChar (PointForCopyFF.PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
   LOGON32_PROVIDER_WINNT50, hToken))
   then GetLastError();
   except on E: Exception do frmdomaininfo.Memo1.Lines.Add(ListPC[i]+' : ������ LogonUser - '+e.Message)
   end;
   //////////////////////
  // if FindAdncreateDir(PointForCopyFF.FDest,ListPC[i]) then // ���� ������� ���� ��� ������. �� ��������
  //����� ������������ ������� ���������� ���� ����������
   begin
      frmdomaininfo.Memo1.Lines.Add(listpc[i]+'  �������- ''\\'+ListPC[i]+'\'+PointForCopyFF.FDest);
      if  CopyFileEx(PChar(PointForCopyFF.FSource), //lpExistingFileName - ��������
                    PChar('\\'+ListPC[i]+'\'+PointForCopyFF.FDest),   //pNewFileName - ����������  // ���� ��������  \\���������\c$\temp\
                    nil,//@CopyCallBack,                 //@CopyCallBack - ������� ���� ���������� �������� ��������� �����������
                    nil,                           //pData - ������ ��� ���������� �������
                    @CopyCancelFF[PointForCopyFF.NumCount],   // pbCancel - ���� ��� ��������� �����������
                    0)     // dwCopyFlags -  COPY_FILE_FAIL_IF_EXISTS - ���� ���� ���������� ����������, �� ����������� �� ����������. COPY_FILE_RESTARTABLE - ���������, ��� ���� � �������� ����������� ��������� ������, �� ����������� ����� ����� ���������� �������, ������ ��� lpExistingFileName, lpNewFileName �� �� �������� ��� ���� � ���������� ������ �������.
      then finditem('����������� ����� ' +ExtractFileName(PointForCopyFF.FSource)+' : �������� ��������� �������',ListPC[i],1)
      else
      begin
      GetLastError(); // ����� ������ ������ �� ���������� �����������
      finditem('����������� ����� ' +ExtractFileName(PointForCopyFF.FSource)+' : ��������� ������ ',ListPC[i],2)
      end;
    end; // ����� ����������� �����
    except
      on E: Exception do
      begin
       finditem('����������� ����� ' +ExtractFileName(PointForCopyFF.FSource)+': ������ ThreadCopyFileEx  : '+E.Message,ListPC[i],2);
       frmdomaininfo.Memo1.Lines.Add(Format(step+' : '+ListPC[i]+': ������ ThreadCopyFileEx  :  %s',[E.Message]));
       end;
    end;
CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
END;

ListPC.Free;
except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format(step+' : ������ ThreadCopyFileEx  :  %s',[E.Message]));
  end;
EndThread(0);    // ������� �����
end;

function TForm11.FunctionStartThreadCopyFileEx(FSource,FDest,username,passwd:string;
CancelCopyFF:boolean;CountNewForm:integer;ListPC:TstringList):boolean;
///////////////////////////////                             /// ������� ������� ������ ����������� ����� CopeFileEx
 function FileSize (const AFileName: string): Int64;
 var
   FS: TFileStream;
 begin
   FS := TFileStream.Create(AFileName,fmOpenRead);
   try
     Result := FS.Size
   finally
     FS.Free;
   end;
 end;
begin
try
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // ����������� ������
  RunCopyFF[CountNewForm].ListPC:=ListPC.CommaText;
  RunCopyFF[CountNewForm].UserName:=username;
  RunCopyFF[CountNewForm].PassWd:=passwd;
  //RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// ����� ����� ����������� ��� �������� ����, ����� ������ ��� ����������� ����� ����� �����������
 // RunCopyFF[CountNewForm].ProgBar:=FProgressBar;  // � ����� ������� ���� ���������� ������� ����������, ��������� ��� CopyFileEx
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFileEx),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format(' - ������ FunctionStartThreadCopyFileEx  :  %s',[E.Message]));
  end;
 end;

procedure TForm11.TaskDialog1VerificationClicked(Sender: TObject);
begin

end;

//////////////////////////////////////////////////////////////

function ThreadCopyFF(paramCopyFF:pointer):boolean; // ����������� � ������ ��� ������ �����
var
SHFileOpStruct : TSHFileOpStruct;
begin
PointForCopyFF:=paramCopyFF;
///////////////////////////// ����� ������� ������� � vista , �������� ����� � ��������
  with SHFileOpStruct do
    begin
      Wnd := PointForCopyFF.OwnerForm.Handle;
      case PointForCopyFF.TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      pFrom := pchar(PointForCopyFF.FSource);
      pTo := pchar(PointForCopyFF.FDest);
      fFlags := 0;//FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR or FOF_RENAMEONCOLLISION or FOF_SIMPLEPROGRESS;
      fAnyOperationsAborted := PointForCopyFF.CancelCopyFF;
      hNameMappings := nil;
      lpszProgressTitle :=nil;
    end;
  SHFileOperation( SHFileOpStruct);//��������������� ������� �����������/�����������
EndThread(0);    // ������� �����
end;

///////////////////////////////////////////////////////////////////////
function ThreadCopyFFSelectPC(paramCopyFF:pointer):boolean; // ����������� � ������ ��� �����  �����������
var
SHFileOpStruct : TSHFileOpStruct;
ListPC,ListDel:Tstringlist;
rescopy,i,z:integer;
htoken:THandle;
begin
try
PointForCopyFF:=paramCopyFF;
try
ListPC:=tstringlist.Create;
ListPC.CommaText:=PointForCopyFF.ListPC; // ������ �����������
///////////////////////////// ����� ������� ������� � vista , �������� ����� � �������� ��� ������ �����������
  with SHFileOpStruct do
    begin
      Wnd := PointForCopyFF.OwnerForm.Handle;
      case PointForCopyFF.TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4: wFunc :=  FO_RENAME;
      end;
      fFlags :=FOF_NO_UI or FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR;
      fAnyOperationsAborted := PointForCopyFF.CancelCopyFF;
      hNameMappings := nil;                //����� ������� ������������, ���� ���� �������� ���������� ������� ������������� ����, ������� �������� ������ � ����� ����� ��������������� ������
      lpszProgressTitle :=nil;            // ��������� �� ��������� ����������� ���� ���������
    end;
    for I := 0 to ListPC.Count-1 do
    if ping(ListPC[i]) then
    Begin

     try ////////////////////////////////////// ����������� ������������ �� ��������� �����
     if not (LogonUserA (PAnsiChar(PointForCopyFF.UserName), PAnsiChar (ListPC[i]),  // ������� ������� �� ���� � ����
     PAnsiChar (PointForCopyFF.PassWd), 9,          //9-LOGON32_LOGON_NEW_CREDENTIALS
     LOGON32_PROVIDER_WINNT50, hToken))
     then GetLastError();
     except on E: Exception do frmdomaininfo.Memo1.Lines.Add(ListPC[i]+' : ������ LogonUser - '+e.Message)
     end;
     ////////////////////////////////////////////
     try
     if PointForCopyFF.PathCreate then FindAdncreateDir(PointForCopyFF.FDest,ListPC[i]); //��������� � ������� ������� ���� ��� ���.

     if PointForCopyFF.TypeOperation=2 then //���� �������� ����������� �� ���������
      begin
      SHFileOpStruct.pTo := pchar('\\'+ListPC[i]+'\'+PointForCopyFF.FDest);   // ���� ��������
      SHFileOpStruct.pFrom := pchar(PointForCopyFF.FSource); //���� �������� �����������, �� ���� ��������
      rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/�����������/��������
      if rescopy=0 then finditem('�������� ����������� ������� ���������',ListPC[i],1)
      else finditem('������� ����������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',ListPC[i],2);
      end;

      if PointForCopyFF.TypeOperation=3 then //���� �������� ��������
       begin
       ListDel:=TStringList.Create;
       try
       ListDel.CommaText:=PointForCopyFF.FDest;
       for z := 0 to ListDel.Count-1 do
        begin
        SHFileOpStruct.pTo := pchar('');  // �� ��� �����
        SHFileOpStruct.pFrom:= pchar('\\'+ListPC[i]+'\'+ListDel[z]+#0+#0); // ��� ��� ������� , ��������� ������� 0 �.�. �������� ������ � ������ �� ����������
        rescopy:=SHFileOperation(SHFileOpStruct); //��������������� ������� �����������/�����������/��������
        if rescopy=0 then finditem('�������� �������� ������� ���������',ListPC[i],1)
        else finditem('������� �������� ����������� �������: '+SysErrorMessage(rescopy)+' ('+inttostr(rescopy)+')',ListPC[i],2);
        end;
       finally
       ListDel.Free;
       end;
       end;

      CloseHandle(htoken);  //��������� ����� ��������� ��� LogonUser
      except
      on E: Exception do
      begin
      case PointForCopyFF.TypeOperation of
       2:begin
       frmdomaininfo.Memo1.Lines.Add(Format('������ ������� ����������� �� ���������  %s :  %s',[listpc[i],E.Message]));
       finditem('������ ������� �����������: '+E.Message,ListPC[i],2);
       end;
       3:begin
        frmdomaininfo.Memo1.Lines.Add(Format('������ ������� �������� �� ����������  %s :  %s',[listpc[i],E.Message]));
       finditem('������ ������� ��������: '+E.Message,ListPC[i],2);
       end;
      end;
      end;
      end;
    End; //����
finally
  ListPC.Free;
end;
except
    on E: Exception do
     frmdomaininfo.Memo1.Lines.Add(Format('������ ������� �����������/�������� �� ������ �����������  :  %s',[E.Message]));
  end;
EndThread(0);    // ������� �����
end;
////////////////////////////////////////////////////////////////////////



function TForm11.FunctionCopyFF(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):integer;
var       //������� �����������/����������� �� � ������
SHFileOpStruct : TSHFileOpStruct;
FStyleY:bool;
begin
  with SHFileOpStruct do
    begin
      FStyleY:=false;
      if OwnerForm.FormStyle=fsStayOnTop then
      begin
      OwnerForm.FormStyle:=fsNormal; // ��� ���� ����� ����� �� ����������
      FStyleY:=true;
      end;
      Wnd :=OwnerForm.Handle;
      case TypeOperation of   //   FO_DELETE - �������, FO_MOVE - ����������� ,FO_RENAME - �������������
      1: wFunc := FO_MOVE;
      2: wFunc := FO_COPY;
      3: wFunc := FO_DELETE;
      4:wFunc :=  FO_RENAME;
      end;
      pFrom := pchar(FSource);
      pTo := pchar(FDest);
      fFlags := 0;//FOF_NOCONFIRMATION or FOF_NOCONFIRMMKDIR or FOF_RENAMEONCOLLISION or FOF_SIMPLEPROGRESS;
      fAnyOperationsAborted := CancelCopyFF;
      hNameMappings := nil;
      lpszProgressTitle :=nil;
    end;
    result:=SHFileOperation( SHFileOpStruct ); //��������������� ������� �����������/�����������
  if FStyleY then OwnerForm.FormStyle:=fsStayOnTop;
end;


function TForm11.CopyFileFolder(FSource,FDest:string;CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
///////////////////////////////                             /// ������� ������� ������� � ������  SHFileOperation
begin
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // ����������� ������
  RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// ����� ����� ������ ����������� ��� �������� ���� �����
  RunCopyFF[CountNewForm].NumCount:=CountNewForm;
  RunCopyFF[CountNewForm].TypeOperation:=TypeOperation;
  //CopyCancelFF[CountNewForm]:=false;
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFF),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 end;

 function TForm11.CopyFileFolderListPC(FSource,FDest,userPC,Paswd,StrListPC:
 string;FindCreateFolder,CancelCopyFF:boolean;OwnerForm:Tform;CountNewForm,TypeOperation:integer):boolean;
///////////////////////////////                             /// ������� ������� �������/����������/��������/�������������� � ������ ��� ������ ����������� SHFileOperation
begin
  RunCopyFF[CountNewForm].UserName:=userPC;
  RunCopyFF[CountNewForm].PassWd:=Paswd;
  RunCopyFF[CountNewForm].ListPC:=StrListPC;
  RunCopyFF[CountNewForm].FSource:=FSource;
  RunCopyFF[CountNewForm].FDest:=FDest;
  RunCopyFF[CountNewForm].CancelCopyFF:=false;    // ����������� ������
  RunCopyFF[CountNewForm].OwnerForm:=OwnerForm; /// ����� ����� ������ ����������� ��� �������� ���� �����
  RunCopyFF[CountNewForm].NumCount:=CountNewForm;
  RunCopyFF[CountNewForm].TypeOperation:=TypeOperation;
  RunCopyFF[CountNewForm].PathCreate:=FindCreateFolder;
  //CopyCancelFF[CountNewForm]:=false;
  res[CountNewForm]:=BeginThread(nil,0,addr(ThreadCopyFFSelectPC),Addr(RunCopyFF[CountNewForm]),0,treadID[CountNewForm]); ///
  CloseHandle(res[CountNewForm]);
 end;


procedure TForm11.Clipboard_DataSend(const DataPaths: TStrings;const MoveType: Integer);
{  ���������� �����/����� � ����� ������ ��
  5 = �����������(����� �� ������ Ctrl+C)
  2 = �������(����� �� ������ Ctrl+X)  DROPEFFECT_MOVE
  ���� ����� ����� ���� ��������(Ctrl+V) ��� ������ � ����� �������� ���������.
  uses ShlObj, ClipBrd, Windows; }
var
  DropFiles: PDropFiles;
  hGlobal: THandle;
  iLen: Integer;
  f: Cardinal;
  d: PCardinal;
  DataSpecialList: string;
  // ������ �������(FullPaths) ������/����� ������� ���� ����������

begin
  Clipboard.Open;
  Clipboard.Clear;
  // ��������������� ������ � ����-������
  DataSpecialList := Shell_Str(DataPaths);

  iLen := Length(DataSpecialList) * SizeOf(Char);
  hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT,
    SizeOf(TDropFiles) + iLen);
  Win32Check(hGlobal <> 0);
  DropFiles := GlobalLock(hGlobal);
  DropFiles^.pFiles := SizeOf(TDropFiles);

{$IFDEF UNICODE}
  DropFiles^.fWide := true;
{$ENDIF}
  Move(DataSpecialList[1], (PansiChar(DropFiles) + SizeOf(TDropFiles))^, iLen);
  SetClipboardData(CF_HDROP, hGlobal);
  GlobalUnlock(hGlobal);
  // FOR COPY
  begin
    f := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
    hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT,
      SizeOf(dword));
    d := PCardinal(GlobalLock(hGlobal));
    d^ := MoveType; // 2-Cut, 5-Copy
    SetClipboardData(f, hGlobal);
    GlobalUnlock(hGlobal);
  end;
  Clipboard.Close;
end;


end.
