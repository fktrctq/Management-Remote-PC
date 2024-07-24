program uRDMServer;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Form1},
  VersionCheck in 'CheckOS\VersionCheck.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar :=false;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

