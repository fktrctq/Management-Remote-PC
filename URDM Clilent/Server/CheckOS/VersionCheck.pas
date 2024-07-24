UNIT VersionCheck;

INTERFACE
TYPE
  TWinVersions    = (vWinUnknown,
                     vWin32S,
                     vWin95,   vWin95OSR2,
                     vWin98,   vWin98SE,
                     vWinME,
                     vWinNT3,  vWinNT4,
                     vWin2000, vWinXP, vWin2003Server,
                     vWinVista, vWin7, vWin8, vWin81,
                     vWin10);
  TOSversions     = (o32bit, o64bit);

CONST
  TWinVersionsStr : Array[TWinVersions] of String[25] =
                    ('Unknown Windows Version',
                     'Windows 32S',
                     'Windows 95',        'Windows 95 OSR2',
                     'Windows 98',        'Windows 98 SE',
                     'Windows Millenium',
                     'Windows NT 3.x',    'Windows NT 4.x',
                     'Windows 2000',      'Windows XP',      'Windows Server 2003',
                     'Windows Vista',     'Windows 7',
                     'Windows 8', 'Windows 8.1',
                     'Windows 10');
  TOSversionsStr  : Array[TOSversions] of String[6] =
                    ('32 Bit', '64 bit');


Function WinVersionMajor : integer;
Function WinVersionMinor : integer;
Function OSversion : TOSversions;
function WMIVer:string;
Function SPinfo : String;


IMPLEMENTATION
USES
  Windows,
  CommCtrl,ActiveX,ComObj;

{$R VersionCheck.res}
// The resource file attaches the Version manifest to this unit
// and thus to the program.
// A manifest is required since Windows 8.1 to detect the actual
// version number of the OS .......
// The manifest is saved in the file: VersionCheck.manifest
// To create the resource file issue the command:
//        "C:\.......\BRCC32.exe" VersionCheck.rc
// The content of this RC file is just 1 line:
//        1 24 VersionCheck.manifest
//
// The manifest and procedure to add the manifest to a Delphi
// program is described by:
//        Glenn9999, to be found on the Tek-Tip Forums:
//        http://www.tek-tips.com/faqs.cfm?fid=7848

function WMIVer:string;
var
FWbemObjectSet: OLEVariant;
FWMIService   : OLEVariant;
oEnum         : IEnumvariant;
FWbemObject   : OLEVariant;
FSWbemLocator : OLEVariant;
iValue        : LongWord;
res:string;
const
wbemFlagForwardOnly = $00000020;
begin
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', ''); ///WbemUser, WbemPassword
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_OperatingSystem','WQL',wbemFlagForwardOnly);
       oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Операционная система
            begin
            result:= string(FWbemObject.version);
            end;
         VariantClear(FWbemObject);
         oEnum:=nil;
         VariantClear(FWbemObjectSet);
         VariantClear(FWMIService);
         VariantClear(FSWbemLocator);

end;

Function SPinfo : String;
Var
  OS : TOSVERSIONINFO;
Begin
  Result                 := '';
  OS.dwPlatformId        := 0;
  OS.dwOSVersionInfoSize := SizeOf(OSVERSIONINFO);
  GetVersionEx(OS);
  Result := OS.szCSDVersion;
End;

Function OSversion : TOSversions;
Type
  TIsWow64Process = function( // Type of IsWow64Process API fn
           Handle  : THandle;
       Var Res     : Boolean) : Boolean; stdcall;
Var       
  IsWow64Process    : TIsWow64Process;  // IsWow64Process fn reference
  IsWow64           : Boolean;
Begin
  IsWow64Process := GetProcAddress(GetModuleHandle('kernel32'), 'IsWow64Process' );
  If Assigned(IsWow64Process) then
    begin
      IsWow64Process(GetCurrentProcess, IsWow64);
      If IsWow64 then Result := o64bit
      else Result := o32bit;
    end
  else Result := o32bit;
end;

Function WinVersionMajor : integer;
Var
  OS : TOSVERSIONINFOEX;
Begin
  OS.dwPlatformId        := 0;
  OS.dwOSVersionInfoSize := SizeOf(OSVERSIONINFOEX);
  GetVersionEx(OS);
  Case OS.dwPlatformId of
     VER_PLATFORM_WIN32_NT:
       begin
         If (OS.dwMajorVersion = 10)  then Result := 10
         else
         If (OS.dwMajorVersion = 6)  then Result := 6
         else
         If (OS.dwMajorVersion = 6) then Result := 6
         else
         If (OS.dwMajorVersion = 6)  then Result := 6
         else
         If (OS.dwMajorVersion = 6) then Result := 6
         else
         If (OS.dwMajorVersion = 5) then Result := 5
         else
         If (OS.dwMajorVersion = 5)  then Result := 5
         else
         If (OS.dwMajorVersion = 5)  then Result := 5
         else
         If (OS.dwMajorVersion = 4) then Result := 4
         else
         If (OS.dwMajorVersion < 4) then Result := 3
         else;
       end;
  end;
End;


Function WinVersionMinor : integer;
Var
  OS : TOSVERSIONINFOEX;
Begin
  OS.dwPlatformId        := 0;
  OS.dwOSVersionInfoSize := SizeOf(OSVERSIONINFOEX);
  GetVersionEx(OS);
  Case OS.dwPlatformId of
     VER_PLATFORM_WIN32_NT:
       begin
         If (OS.dwMinorVersion = 0) then Result := 0
         else
         If (OS.dwMinorVersion = 3) then Result :=3 // vWin81
         else
         If (OS.dwMinorVersion = 2) then Result := 2 //vWin8
         else
         If  (OS.dwMinorVersion = 1) then Result :=1 // vWin7
         else
         If (OS.dwMinorVersion = 0) then Result := 0; //vWinVista
       end;
  end;
End;
INITIALIZATION
  // this call is necessary for the newer "theme" controls.  Note YMMV on Common Controls 6.0
  // ("theme controls") and you may require other corrective changes in your programs for
  // those to work (Glenn9999)
  InitCommonControls;
END.
