program MRRDP;

uses
  Vcl.Forms,
  Vcl.Dialogs,
  Windows,
  Registry,
  System.SysUtils,
  URDP in 'URDP.pas' {FormRDP},
  Vcl.Themes,
  Vcl.Styles,
  MyDMRDP in 'MyDMRDP.pas' {DataMod: TDataModule},
  LoadListPC in 'LoadListPC.pas' {FormLoadList},
  Unit1 in 'Unit1.pas',
  ADWork in 'ADWork.pas';

{$R *.res}
 function regeditread(patch,Section:string):string;
var      /// чтение из реестра
regFile:TregIniFile;
begin
result:='';
RegFile:=TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
RegFile.RootKey:=RootPatch;
if regFile.KeyExists(patch) then
begin
result:=(regFile.ReadString(patch,Section,''));
end;
if Assigned(RegFile) then freeandnil(regFile);
end;


function CheckCopyRUN(s:string): boolean;
begin
  HM:= OpenMutex(MUTEX_ALL_ACCESS, false, pchar(s));
  Result:= (HM <> 0);
  if HM = 0 then HM:= CreateMutex(nil, false, pchar(s));
end;

function CheckOwnerRUN(s:string): boolean;
begin
  HML:= OpenMutex(MUTEX_ALL_ACCESS, false, pchar(s));
  Result:= (HML <> 0);
  if HML = 0 then HML:= CreateMutex(nil, false, pchar(s));
end;

function FindProc(TipForm,NameForm:string):bool;
//var
//hwn,OwnerHwn: THandle;
begin           //'TForm1', 'Form1'
  try
 // hwn := FindWindow(pchar(TipForm),nil); //pchar(TipForm)
  if FindWindow(pchar(TipForm),nil)<>0 then result:=true
  else result:=false;
  except
   result:=false;
   end;
end;

function ActProc(TipForm,NameForm:string):bool;

begin           //'TForm1', 'Form1'
  try
  hwn := FindWindow(pchar(TipForm),nil); //pchar(TipForm)
  if hwn<>0 then
  begin
  //OwnerHwn:=GetWindow(hwn,GW_OWNER); /// получаем родителя
  ShowWindow(hwn, SW_SHOWMAXIMIZED); //SW_RESTORE   // SW_SHOWMAXIMIZED  //SW_SHOWNORMAL
  SetForegroundWindow(hwn); //GetWindow(hwn,GW_OWNER)
  result:=true;
  end
  else result:=false;
  except
   result:=false;
   end;
end;


begin
 if CheckCopyRUN('RDPForMRPC') then /// если мьютекс есть значит уже запушено
 begin
  if  ActProc('TFormRDP','') then /// ищем приложение и переводим на передний план
  halt; // если нашли приложение то текущее убиваем
 end;


  {if not FindProc('TfrmDomainInfo','') then // если не запущен Management Remote PC
    begin
    showmessage('Запуск только из основного приложения "Management Remote PC".');
    halt;
    end;}



  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Turquoise Gray');
  Application.Title := 'RDP клиент (skrblog.ru)';
  Application.CreateForm(TFormRDP, FormRDP);
  Application.CreateForm(TDataMod, DataMod);
  Application.CreateForm(TFormLoadList, FormLoadList);
  Application.Run;
end.
