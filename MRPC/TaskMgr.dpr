program TaskMgr;

uses
  Forms,
  Windows,
  Registry,
  System.SysUtils,
  uMain in 'uMain.pas' {frmDomainInfo},
  Unit2 in 'Unit2.pas',
  Unit3 in 'Unit3.pas',
  Unit4 in 'Unit4.pas',
  Unit5 in 'Unit5.pas',
  FormNewProcess in 'FormNewProcess.pas' {NewProcForm},
  Unit6 in 'Unit6.pas',
  Unit7 in 'Unit7.pas',
  NewServic in 'NewServic.pas' {OKRightDlg},
  PCShutDown in 'PCShutDown.pas',
  NewStartup in 'NewStartup.pas' {OKRightDlg2},
  DlgNewInstall in 'DlgNewInstall.pas' {OKRightDlg1},
  DeleteProg in 'DeleteProg.pas',
  UninstallOtherProg in 'UninstallOtherProg.pas' {OKRightDlg12},
  HDDChkdsk in 'HDDChkdsk.pas',
  DlgChkdsk in 'DlgChkdsk.pas' {OKRightDlg123},
  DefragHDD in 'DefragHDD.pas' {OKRightDlg1234},
  DefragForce in 'DefragForce.pas',
  Dismount in 'Dismount.pas',
  HDDMount in 'HDDMount.pas',
  HDDFormat in 'HDDFormat.pas',
  DlgFormatHDD in 'DlgFormatHDD.pas' {OKRightDlg12345},
  DefragAnalysis in 'DefragAnalysis.pas',
  DlgJoinDomainOrWorkGroup in 'DlgJoinDomainOrWorkGroup.pas' {OKRightDlg123456},
  DlgPCRename in 'DlgPCRename.pas' {OKRightDlg1234567},
  DlgUnjoinDomainOrWorkgroup in 'DlgUnjoinDomainOrWorkgroup.pas' {OKRightDlg12345678},
  AddNetworkPrint in 'AddNetworkPrint.pas',
  NetworkAddPrint in 'NetworkAddPrint.pas' {OKRightDlg123456789},
  AddPrintDriverThread in 'AddPrintDriverThread.pas',
  AddPrintDriverDialog in 'AddPrintDriverDialog.pas' {OKRightDlg12345678910},
  PrintCancelAllJob in 'PrintCancelAllJob.pas',
  PrintTestPage in 'PrintTestPage.pas',
  PrintRenamePrint in 'PrintRenamePrint.pas',
  PrintRenamePrintDialog in 'PrintRenamePrintDialog.pas' {OKRightDlg1234567891011},
  PrintResumePrint in 'PrintResumePrint.pas',
  PrintPausePrint in 'PrintPausePrint.pas',
  PrinterSetDefaultPrinter in 'PrinterSetDefaultPrinter.pas',
  ChangeOwnerThread in 'ChangeOwnerThread.pas',
  NewUserProfileDialog in 'NewUserProfileDialog.pas' {OKRightDlg123456789101112},
  PrintDeleteThread in 'PrintDeleteThread.pas',
  PcProperties in 'PcProperties.pas' {OKRightDlg12345678910111213},
  UserProfileDelete in 'UserProfileDelete.pas',
  PropertiesNetworkDialog in 'PropertiesNetworkDialog.pas' {OKRightDlg1234567891011121314},
  PropertiesNetworkThread in 'PropertiesNetworkThread.pas',
  NetworkAdapterDisable in 'NetworkAdapterDisable.pas',
  NetworkAdapterEnable in 'NetworkAdapterEnable.pas',
  NetworkConfiguration in 'NetworkConfiguration.pas' {OKRightDlg123456789101112131415},
  NetworkAdapterDHCP in 'NetworkAdapterDHCP.pas',
  NetworkAdapterStaticTCPIPAll in 'NetworkAdapterStaticTCPIPAll.pas',
  NetworkAdapterDNSThread in 'NetworkAdapterDNSThread.pas',
  NetworkAdapterAdditionallyDialog in 'NetworkAdapterAdditionallyDialog.pas' {OKRightDlg12345678910111213141516},
  NetworkAdapterAddIPSubnet in 'NetworkAdapterAddIPSubnet.pas' {OKRightDlg1234567891011121314151617},
  NetworkAdapterAddGateway in 'NetworkAdapterAddGateway.pas' {OKRightDlg123456789101112131415161718},
  NetworkAdapterAddDnsWins in 'NetworkAdapterAddDnsWins.pas' {OKRightDlg12345678910111213141516171819},
  SelectedPCNewProcessThread in 'SelectedPCNewProcessThread.pas',
  SelectedPCShotDownThread in 'SelectedPCShotDownThread.pas',
  GeneralPingPC in 'GeneralPingPC.pas',
  IPaddress in 'IPaddress.pas' {OKRightDlg1234567891011121314151617181920},
  SelectedPCInstallDriverPrint in 'SelectedPCInstallDriverPrint.pas',
  Vcl.Themes,
  Vcl.Styles,
  SelectedPCAddPrinterThread in 'SelectedPCAddPrinterThread.pas',
  SelectedPCKillProcess in 'SelectedPCKillProcess.pas' {OKRightDlg123456789101112131415161718192021},
  SelectedPCDeleteProgramThread in 'SelectedPCDeleteProgramThread.pas',
  SelectedPCDeleteProgramDialog in 'SelectedPCDeleteProgramDialog.pas' {OKRightDlg12345678910111213141516171819202122},
  WindowsActivation in 'WindowsActivation.pas' {ActivationWindows},
  OfficeActivation in 'OfficeActivation.pas' {ActivateOffice},
  RemoteDesktopSettingDialog in 'RemoteDesktopSettingDialog.pas' {RemoteDesktopSetting},
  RemoteDesktopSettingTransformDialog in 'RemoteDesktopSettingTransformDialog.pas' {RemoteDesktopSettionTransform},
  MRDUNew in 'MRDUNew.pas' {MRDForm},
  SettingsProgramForm in 'SettingsProgramForm.pas' {SettingsProgram},
  WOLThread in 'WOLThread.pas',
  MyDM in 'MyDM.pas' {DataM: TDataModule},
  MyInventoryPC in 'MyInventoryPC.pas',
  ExportExcel in 'ExportExcel.pas',
  MyInventorySoft in 'MyInventorySoft.pas',
  SMARTStat in 'SMARTStat.pas' {Form8},
  PrintProperty in 'PrintProperty.pas' {PrinterProperty},
  PrintSecuretyEdit in 'PrintSecuretyEdit.pas' {PrintSecuretyEditor},
  unit8 in 'unit8.pas',
  Unit9 in 'Unit9.pas',
  Unit10 in 'Unit10.pas' {Form10},
  OneStartForm in 'OneStartForm.pas' {FormOneStart},
  AddPrintLAN in 'AddPrintLAN.pas',
  ShareFS in 'ShareFS.pas',

  FormFSShare in 'FormFSShare.pas' {FormShareFS},
  FolderSecuretyEdit in 'FolderSecuretyEdit.pas' {FrmSecFolder},
  FindProcessGroup in 'FindProcessGroup.pas',
  FormFindProcess in 'FormFindProcess.pas' {Form12},
  EventWindows in 'EventWindows.pas',
  LogFileEventProperties in 'LogFileEventProperties.pas' {PropertiesFileEvent},
  LoadCurencyList in 'LoadCurencyList.pas',
  PropertyAccount in 'PropertyAccount.pas' {FormUserAccount},
  AccountLocal in 'AccountLocal.pas' {FormLocalAccount},
  PropertyGroup in 'PropertyGroup.pas' {FormPropertyGroup},
  NewRDPWin in 'NewRDPWin.pas' {RDPWin},
  RemoteExplorer in 'RemoteExplorer.pas' {MRPCExplorer},
  uStackTrace in 'uStackTrace.pas',
  CopyPasteFF in 'CopyPasteFF.pas',
  FormCopyFF in 'FormCopyFF.pas' {Form11},
  RenamePC in 'RenamePC.pas',
  PCJoinDomainOrWorkGroup in 'PCJoinDomainOrWorkGroup.pas',
  PCUnJoinDomainOrWorkgroup in 'PCUnJoinDomainOrWorkgroup.pas',
  FormForCopyPCList in 'FormForCopyPCList.pas' {FormCopyFFFGropPC},
  RunTASK in 'RunTASK.pas',
  TaskEdit in 'TaskEdit.pas' {EditTask},
  TaskSelect in 'TaskSelect.pas' {SelectTask},
  TaskNewMSI in 'TaskNewMSI.pas' {NewMSITask},
  TaskNewProc in 'TaskNewProc.pas' {NewProcTask},
  TaskSelectDelMSI in 'TaskSelectDelMSI.pas' {SelectDelMSITask},
  TaskCopyFF in 'TaskCopyFF.pas' {TaskCopyDelFF},
  MainPopup in 'MainPopup.pas' {MainFormPopup},
  RegEditForm in 'RegEditForm.pas' {RegEdit},
  RegEditKeySave in 'RegEditKeySave.pas' {RegKeySave},
  RunTaskRegEdit in 'RunTaskRegEdit.pas',
  RestoreWindows in 'RestoreWindows.pas' {FormRestoreWin},
  KMS in 'KMS.pas' {FormKMS},
  InventoryWindowsOffice in 'InventoryWindowsOffice.pas',
  ErrorLic in 'ErrorLic.pas' {CodeErrorLic},
  RunTaskOsher in 'RunTaskOsher.pas',
  MessageSystem in 'MessageSystem.pas',
  PingForScan in 'PingForScan.pas',
  PingForMain in 'PingForMain.pas',
  PingForTask in 'PingForTask.pas',
  EditProcMsi in 'EditProcMsi.pas' {FormEditMsiProc},
  My_Proc in 'My_Proc.pas' {EditMyProc};

{$R *.res}

/// stringpatch:string ='software\MRPC\MRPC'
// peremennaya:=regeditread(stringpatch,'isOK');
function regeditread(patch, Section: string): string;
var
/// чтение из реестра
  regFile: TregIniFile;
begin
  result := '';
  regFile := TregIniFile.Create(KEY_READ OR KEY_WOW64_64KEY);
  regFile.RootKey := RootPatch;
  if regFile.KeyExists(patch) then
    begin
      result := (regFile.ReadString(patch, Section, ''));
    end;
  if Assigned(regFile) then
    freeandnil(regFile);
end;

function CheckMut: boolean;
begin
  HM := OpenMutex(MUTEX_ALL_ACCESS, false, 'MForMRPC');
  /// есть ли мьютекс
  result := (HM <> 0);
  if HM = 0 then
    HM := CreateMutex(nil, false, 'MForMRPC');
  /// если нет тоосоздаем
end;

function ActProc(TipForm, NameForm: string): bool;
var
  hwn: THandle;
begin // 'TForm1', 'Form1'
  try
    hwn := FindWindow(pchar(TipForm), nil);
    if hwn <> 0 then
      begin
        // OwnerHwn:=GetWindow(hwn,GW_OWNER); /// получаем родителя
        ShowWindow(GetWindow(hwn, GW_OWNER), SW_SHOWNORMAL);
        // SW_RESTORE   // SW_SHOWMAXIMIZED  //SW_SHOWNORMAL
        SetForegroundWindow(GetWindow(hwn, GW_OWNER));
        result := true;
      end
    else
      result := false;
  except
    result := false;
  end;
end;

begin
  if CheckMut then
  /// если уже запущенна то ищем приложение
    begin
      if ActProc('TfrmDomainInfo', '') then
        halt;
    end;

  Application.Initialize;
  TStyleManager.TrySetStyle('Light');
  Application.CreateForm(TfrmDomainInfo, frmDomainInfo);
  Application.CreateForm(TNewProcForm, NewProcForm);
  Application.CreateForm(TOKRightDlg, OKRightDlg);
  Application.CreateForm(TOKRightDlg2, OKRightDlg2);
  Application.CreateForm(TOKRightDlg1, OKRightDlg1);
  Application.CreateForm(TOKRightDlg12, OKRightDlg12);
  Application.CreateForm(TOKRightDlg123, OKRightDlg123);
  Application.CreateForm(TOKRightDlg1234, OKRightDlg1234);
  Application.CreateForm(TOKRightDlg12345, OKRightDlg12345);
  Application.CreateForm(TOKRightDlg123456, OKRightDlg123456);
  Application.CreateForm(TOKRightDlg1234567, OKRightDlg1234567);
  Application.CreateForm(TOKRightDlg12345678, OKRightDlg12345678);
  Application.CreateForm(TOKRightDlg123456789, OKRightDlg123456789);
  Application.CreateForm(TOKRightDlg12345678910, OKRightDlg12345678910);
  Application.CreateForm(TOKRightDlg1234567891011, OKRightDlg1234567891011);
  Application.CreateForm(TOKRightDlg123456789101112, OKRightDlg123456789101112);
  Application.CreateForm(TOKRightDlg12345678910111213, OKRightDlg12345678910111213);
  Application.CreateForm(TOKRightDlg1234567891011121314, OKRightDlg1234567891011121314);
  Application.CreateForm(TOKRightDlg123456789101112131415, OKRightDlg123456789101112131415);
  Application.CreateForm(TOKRightDlg12345678910111213141516, OKRightDlg12345678910111213141516);
  Application.CreateForm(TOKRightDlg1234567891011121314151617, OKRightDlg1234567891011121314151617);
  Application.CreateForm(TOKRightDlg123456789101112131415161718, OKRightDlg123456789101112131415161718);
  Application.CreateForm(TOKRightDlg12345678910111213141516171819, OKRightDlg12345678910111213141516171819);
  Application.CreateForm(TOKRightDlg1234567891011121314151617181920, OKRightDlg1234567891011121314151617181920);
  Application.CreateForm(TOKRightDlg123456789101112131415161718192021, OKRightDlg123456789101112131415161718192021);
  Application.CreateForm(TOKRightDlg12345678910111213141516171819202122, OKRightDlg12345678910111213141516171819202122);
  Application.CreateForm(TActivationWindows, ActivationWindows);
  Application.CreateForm(TActivateOffice, ActivateOffice);
  Application.CreateForm(TRemoteDesktopSetting, RemoteDesktopSetting);
  Application.CreateForm(TRemoteDesktopSettionTransform, RemoteDesktopSettionTransform);
  Application.CreateForm(TMRDForm, MRDForm);
  Application.CreateForm(TSettingsProgram, SettingsProgram);
  Application.CreateForm(TDataM, DataM);
  Application.CreateForm(TForm8, Form8);
  Application.CreateForm(TPrinterProperty, PrinterProperty);
  Application.CreateForm(TPrintSecuretyEditor, PrintSecuretyEditor);
  Application.CreateForm(TForm10, Form10);
  Application.CreateForm(TFormOneStart, FormOneStart);
  Application.CreateForm(TFormShareFS, FormShareFS);
  Application.CreateForm(TFrmSecFolder, FrmSecFolder);
  Application.CreateForm(TForm12, Form12);
  Application.CreateForm(TPropertiesFileEvent, PropertiesFileEvent);
  Application.CreateForm(TFormUserAccount, FormUserAccount);
  Application.CreateForm(TFormLocalAccount, FormLocalAccount);
  Application.CreateForm(TFormPropertyGroup, FormPropertyGroup);
  Application.CreateForm(TRDPWin, RDPWin);
  Application.CreateForm(TMRPCExplorer, MRPCExplorer);
  Application.CreateForm(TForm11, Form11);
  Application.CreateForm(TFormCopyFFFGropPC, FormCopyFFFGropPC);
  Application.CreateForm(TEditTask, EditTask);
  Application.CreateForm(TSelectTask, SelectTask);
  Application.CreateForm(TNewMSITask, NewMSITask);
  Application.CreateForm(TNewProcTask, NewProcTask);
  Application.CreateForm(TSelectDelMSITask, SelectDelMSITask);
  Application.CreateForm(TTaskCopyDelFF, TaskCopyDelFF);
  Application.CreateForm(TMainFormPopup, MainFormPopup);
  Application.CreateForm(TRegEdit, RegEdit);
  Application.CreateForm(TRegKeySave, RegKeySave);
  Application.CreateForm(TFormRestoreWin, FormRestoreWin);
  Application.CreateForm(TFormKMS, FormKMS);
  Application.CreateForm(TCodeErrorLic, CodeErrorLic);
  Application.CreateForm(TFormEditMsiProc, FormEditMsiProc);
  Application.CreateForm(TEditMyProc, EditMyProc);
  Application.Run;

end.
