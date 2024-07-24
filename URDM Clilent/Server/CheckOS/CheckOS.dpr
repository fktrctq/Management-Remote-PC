program CheckOS;

uses
  Forms,
  CheckOS_Main in 'CheckOS_Main.pas' {MainForm},
  VersionCheck in 'VersionCheck.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
