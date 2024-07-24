unit Unit1;

interface

uses
  Winapi.Windows,Winapi.Messages, System.SysUtils, System.Classes,
   Vcl.SvcMgr, Vcl.Dialogs, Vcl.ExtCtrls,VCL.Forms,TlHelp32;

type
  TServiceRDM = class(TService)
    Timer1: TTimer;
    procedure Timer1Timer(Sender: TObject);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    function processExists(exeFileName: string): Boolean;
    function Killprocess(exeFileName: string): Boolean;
    function CreateNewProcess(nameProc:string):bool;
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure ServiceShutdown(Sender: TService);
    procedure ServicePause(Sender: TService; var Paused: Boolean);
    procedure WMQueryEndSession(var Message : TMessage); message WM_QUERYENDSESSION;
    procedure timerFW;
    procedure TimeFWstart(Sender: TObject);
     //// ���������� ��������� ��� ������������ ��� ����������� Windows
     private
    { Private declarations }
  public
    function GetServiceController: TServiceController; override;

    { Public declarations }
  end;

var
  ServiceRDM: TServiceRDM;
  CurrentSessionID:cardinal;
  TimeFW:Ttimer;
implementation
 uses FWW;
{$R *.dfm}
function WTSQueryUserToken(SessionId: ULONG; var phToken: THandle): BOOL; stdcall;
external 'Wtsapi32.dll';
function WTSGetActiveConsoleSessionId: DWORD; stdcall;
external 'Kernel32.dll';
//function CreateEnvironmentBlock(var lpEnvironment: Pointer; hToken: THandle; bInherit: BOOL): BOOL; stdcall;
// external 'Userenv.dll';
//function DestroyEnvironmentBlock(lpEnvironment: Pointer): BOOL; stdcall;
// external 'Userenv.dll';

//////////////////////////////////////////////////////
function Log_write(fname, text:string):string;
var f:TStringList;
begin
  //if not DirectoryExists(ExtractFilePath(Application.ExeName)+'logs') then
  // CreateDir(ExtractFilePath(Application.ExeName)+'logs');
  f:=TStringList.Create;
  try
    if FileExists(ExtractFilePath(Application.ExeName)+fname+'.log') then
      f.LoadFromFile(ExtractFilePath(Application.ExeName)+fname+'.log');
    f.Insert(0,DateTimeToStr(Now)+chr(9)+text);
    while f.Count>1000 do f.Delete(1000);
    f.SaveToFile(ExtractFilePath(Application.ExeName)+fname+'.log');
  finally
    f.Destroy;
  end;
end;
//////////////////////////////////////////////////////// ����� ��������
function TServiceRDM.processExists(exeFileName: string): Boolean;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
try
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  Result := False;
  while Integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
      UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) =
      UpperCase(ExeFileName))) then
    begin
      Result := True;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);

    except on E: Exception do
    Log_Write('service','������ ������ ��������: '+e.ClassName +': '+ e.Message);
    end
end;
///////////////////////////////////////////////
function TServiceRDM.Killprocess(exeFileName: string): Boolean; ////����� � ��������� �������
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
  ErrorCode: Cardinal;
begin
 try
      FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
      FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
      ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
      Result := False;
      while Integer(ContinueLoop) <> 0 do
      begin
        if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
          UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) =
          UpperCase(ExeFileName))) then
            begin
              TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID),0);
               Result := True;
            end;
      ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
      end;
      CloseHandle(FSnapshotHandle);
     //Log_Write('service','���������� �������� ������ �������');
    //RaiseLastOSError(ErrorCode);
    except on E: Exception do
    Log_Write('service','������ ���������� �������� '+exeFileName+': '+e.ClassName +': '+ e.Message);
    end;
end;


function TServiceRDM.CreateNewProcess(nameProc:string):bool;
var
  CmdLine: String;
  ErrorCode: Cardinal;
 // ConnSessID: Cardinal;
  Token,NewToken: THandle;
  App: AnsiString;
  FProcessInfo: PROCESS_INFORMATION;
  FStartupInfo: STARTUPINFOW;
  hLib: THandle;
  WTSError:integer;
  env: Pointer;
  dwCreationFlags:DWORD;
begin
result:=false;
 if WTSGetActiveConsoleSessionId<>0  then
       begin
       try
       CurrentSessionID:=WTSGetActiveConsoleSessionId; //// ���������� ������������� ������
        FStartupInfo.cb:=sizeOf(STARTUPINFO);
        FStartupInfo.dwFlags:=STARTF_FORCEOFFFEEDBACK;
        FStartupInfo.lpDesktop := pchar('winsta0\default');
        CmdLine:=ExtractFilePath(Application.ExeName)+nameProc;
        WTSQueryUserToken(CurrentSessionID, Token);
       if not DuplicateTokenEx(Token,MAXIMUM_ALLOWED,nil,SecurityImpersonation, TokenImpersonation, NewToken) then
         begin
          //Log_Write('service','DuplicateTokenEx: ' + SysErrorMessage(GetLastError));
          Result := False;
          CloseHandle(NewToken);
          CloseHandle(Token);
          exit;
         end;
        //DuplicateTokenEx(Token,MAXIMUM_ALLOWED,nil,SecurityDelegation,TokenPrimary, NewToken);
        //env:=nil;
       // CreateEnvironmentBlock(env, NewToken, False);    ///��������� ���������� ����� ��� ���������� ������������
        {env := nil;
         if (CreateEnvironmentBlock(env, Token, False)) then
            begin
              dwCreationFlags := dwCreationFlags or CREATE_UNICODE_ENVIRONMENT or PROFILE_USER;
              Log_Write('service','dwCreationFlags = CREATE_UNICODE_ENVIRONMENT');
            end
            else
            begin
              env := nil;
              Log_Write('service','dwCreationFlags =nil');
            end; }
        //SetEnvironmentVariable('Param', 'Value');
           result:=CreateProcessAsUser(NewToken, ///hToken ����� ��������� ������������
            PWideChar(CmdLine),                        ///lpApplicationName ��� ������, ������� ������ ���� ��������
            nil,           ////lpCommandLine ����������� ����
            nil,                      ///lpProcessAttributes ��������� �� ��������� SECURITY_ATTRIBUTES, ������� ���������� ���������� ������������
            nil,                     ///lpThreadAttributes ��������� �� ��������� SECURITY_ATTRIBUTES, ������� ���������� ���������� ������������
            false,                 ///???bInheritHandles ���� ���� �������� ����� �������� ������ , ������ ����������� ���������� ����������� �������� ����������� ����� ���������
            0,///              dwCreationFlags �����, ������� ��������� ������� ���������� � ��������� ��������  0 ��� CREATE_UNICODE_ENVIRONMENT
            nil,                 ///env -  lpEnvironment  ��������� �� ���� ����� ��� ������ ��������. ���� ���� �������� ����� NULL , ����� ������� ���������� ����� ����������� ��������.
            nil,                /// lpCurrentDirectory  ������ ���� � �������� �������� ��� ��������. ������ ����� ����� ��������� ���� UNC.
            FStartupInfo,      /// lpStartupInfo
            FProcessInfo
            );

        CloseHandle(NewToken);
        CloseHandle(Token);
        CloseHandle(FProcessInfo.hProcess);
        CloseHandle(FProcessInfo.hThread);
        //DestroyEnvironmentBlock(env);
          Log_Write('service','������ �������� ������ �������');
          except on E: Exception do
          Log_Write('service','������ ������� �������� '+CmdLine+' :'+e.ClassName +': '+ e.Message);
          end;
       end;

end;
//////////////////////////////////////////////////////////////////////////////

procedure TServiceRDM.WMQueryEndSession(var Message : TMessage);
 begin   //// ���������� ��������� ��� ������������ ��� ����������� Windows
  try
if (ProcessExists('uRDMServer.exe'))=true then
   begin
   timer1.Enabled:=false;
   Killprocess('uRDMServer.exe');
   end;
Log_Write('service','���������� ������ "Shutdown"');
     except on E: Exception do
       Log_Write('service','������ ���������� ����������: '+e.ClassName +': '+ e.Message);
     end;
end;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  ServiceRDM.Controller(CtrlCode);
end;



function TServiceRDM.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TServiceRDM.ServicePause(Sender: TService; var Paused: Boolean);
begin
//if ProcessExists('uRDMServer.exe')=true then
//   begin
//   timer1.Enabled:=false;
//   Killprocess('uRDMServer.exe');
//   end;
//Log_Write('service','������ ��������������');
end;

procedure TServiceRDM.ServiceShutdown(Sender: TService);
begin
try
if (ProcessExists('uRDMServer.exe'))=true then
   begin
   timer1.Enabled:=false;
   Killprocess('uRDMServer.exe');
   end;
Log_Write('service','���������� ������ "Shutdown"');
     except on E: Exception do
       Log_Write('service','������ ���������� ������ ��� shutdown: '+e.ClassName +': '+ e.Message);
     end;
end;

procedure TServiceRDM.TimeFWstart(Sender: TObject);
begin
FWW.startFW;
end;

procedure TServiceRDM.timerFW;
begin
timeFW:=Ttimer.Create(self);
timeFW.Name:='TFW';
timeFW.interval:=12000;
timeFW.OnTimer:=TimeFWstart;
end;

procedure TServiceRDM.ServiceStart(Sender: TService; var Started: Boolean);
begin
CurrentSessionID:=0;
timer1.Enabled:=true;
if not ProcessExists('uRDMServer.exe') then    //// ������� �� ������� ���� ��� �� ��������� �����
begin
CreateNewProcess('uRDMServer.exe');
end;
Log_Write('service','������ �����������');
FWW.startFW;// �������� FireWall
//timerFW; //������ ������� ��� �������� Firewall c ���������� ������ ��������� ��� �� ������
end;

procedure TServiceRDM.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
if ProcessExists('uRDMServer.exe') then
   begin
   timer1.Enabled:=false;
   Killprocess('uRDMServer.exe');
   end;
Log_Write('service','������ �����������');
end;

procedure TServiceRDM.Timer1Timer(Sender: TObject);
begin
if CurrentSessionID<> WTSGetActiveConsoleSessionId then ///// ���� ��������� ������ ������������, ��������� �������
begin
 if ProcessExists('uRDMServer.exe')=true then
   begin
   CurrentSessionID:=WTSGetActiveConsoleSessionId;
   Killprocess('uRDMServer.exe');
   end;
end;
if not ProcessExists('uRDMServer.exe') then    //// ������� �� ������� ���� ��� �� ��������� �����
CreateNewProcess('uRDMServer.exe');

end;

end.

