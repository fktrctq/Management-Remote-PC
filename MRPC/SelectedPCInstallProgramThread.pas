unit SelectedPCInstallProgramThread;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils,
  IdIcmpClient,Vcl.Controls;

type
  SelectedPCInstallProgram = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

var
MyIdIcmpClient: TIdIcmpClient;

implementation
uses umain;
{ SelectedPCInstallProgram }



procedure SelectedPCInstallProgram.Execute;
begin
END;


end.
