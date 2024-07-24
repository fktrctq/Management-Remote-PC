unit MessageSystem;

interface


uses
  Windows;

const
  // ��������� WM_WTSSESSION_CHANGE ���������� ���������� �� ���������� � ��������� ������..
  WM_WTSSESSION_CHANGE = $2B1;

  // wParam values:
  WTS_CONSOLE_CONNECT = 1;
  WTS_CONSOLE_DISCONNECT = 2;
  WTS_REMOTE_CONNECT = 3;
  WTS_REMOTE_DISCONNECT = 4;
  WTS_SESSION_LOGON = 5;
  WTS_SESSION_LOGOFF = 6;
  WTS_SESSION_LOCK = 7;
  WTS_SESSION_UNLOCK = 8;
  WTS_SESSION_REMOTE_CONTROL = 9;

  // ������ ����������� � ������, ��������� � �������, ������������� � ����
  // ���������������� ��������� ��������� hWnd, ������ ���� ��������.
  NOTIFY_FOR_THIS_SESSION = 0;
  // ��� ����������� � ������ ������ ���� ��������.
  NOTIFY_FOR_ALL_SESSIONS = 1;


function RegisterSessionNotification(Wnd: HWND; dwFlags: DWORD): Boolean;
function UnRegisterSessionNotification(Wnd: HWND): Boolean;
function GetCurrentSessionID: Integer;

implementation

function RegisterSessionNotification(Wnd: HWND; dwFlags: DWORD): Boolean;
 // ������� RegisterSessionNotification ������������ ��������� ����
  // ��� ��������� ����������� �� ��������� ������.
  // ���������:
  // hWnd: ���������� ���� ��� ��������� ����������� �� ��������� ������.
  // dwFlags: ����������, ����� ����������� ������ ������ ���� ��������:d:
  // (NOTIFY_FOR_THIS_SESSION, NOTIFY_FOR_ALL_SESSIONS)
type
  TWTSRegisterSessionNotification = function(Wnd: HWND; dwFlags: DWORD): BOOL; stdcall;
var
  hWTSapi32dll: THandle;
  WTSRegisterSessionNotification: TWTSRegisterSessionNotification;
begin
  Result := False;
  hWTSAPI32DLL := LoadLibrary('Wtsapi32.dll');
  if (hWTSAPI32DLL > 0) then
  begin
    try @WTSRegisterSessionNotification :=
        GetProcAddress(hWTSAPI32DLL, 'WTSRegisterSessionNotification');
      if Assigned(WTSRegisterSessionNotification) then
      begin
        Result:= WTSRegisterSessionNotification(Wnd, dwFlags);
      end;
    finally
      if hWTSAPI32DLL > 0 then
        FreeLibrary(hWTSAPI32DLL);
    end;
  end;
end;

function UnRegisterSessionNotification(Wnd: HWND): Boolean;
// ������� RegisterSessionNotification �������� ����������� ���������� ����
  // ���������:
  // hWnd: ���������� ����
type
  TWTSUnRegisterSessionNotification = function(Wnd: HWND): BOOL; stdcall;
var
  hWTSapi32dll: THandle;
  WTSUnRegisterSessionNotification: TWTSUnRegisterSessionNotification;
begin
  Result := False;
  hWTSAPI32DLL := LoadLibrary('Wtsapi32.dll');
  if (hWTSAPI32DLL > 0) then
  begin
    try @WTSUnRegisterSessionNotification :=
        GetProcAddress(hWTSAPI32DLL, 'WTSUnRegisterSessionNotification');
      if Assigned(WTSUnRegisterSessionNotification) then
      begin
        Result:= WTSUnRegisterSessionNotification(Wnd);
      end;
    finally
      if hWTSAPI32DLL > 0 then
        FreeLibrary(hWTSAPI32DLL);
    end;
  end;
end;

function GetCurrentSessionID: Integer;
 // ��������� �������������� ������ �� �������� ��������
type
  TProcessIdToSessionId = function(dwProcessId: DWORD; pSessionId: DWORD): BOOL; stdcall;
var
  ProcessIdToSessionId: TProcessIdToSessionId;
  hWTSapi32dll: THandle;
  Lib : THandle;
  pSessionId : DWord;
begin
  Result := 0;
  Lib := GetModuleHandle('kernel32');
  if Lib <> 0 then
  begin
    ProcessIdToSessionId := GetProcAddress(Lib, '1ProcessIdToSessionId');
    if Assigned(ProcessIdToSessionId) then
    begin
      ProcessIdToSessionId(GetCurrentProcessId(), DWORD(@pSessionId));
      Result:= pSessionId;
    end;
  end;
end;

end.
