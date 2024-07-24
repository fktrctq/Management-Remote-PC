UNIT CheckOS_Main;

INTERFACE

USES
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

TYPE
  TMainForm = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    L3: TLabel;
    L2: TLabel;
    L1: TLabel;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

VAR
  MainForm: TMainForm;

IMPLEMENTATION
USES
  VersionCheck;

{$R *.dfm}

procedure TMainForm.FormActivate(Sender: TObject);
begin
  L1.Caption := TWinVersionsStr[WinVersion];
  L2.Caption := TOSversionsStr[OSversion];
  L3.Caption := SPinfo;
end;

END.



