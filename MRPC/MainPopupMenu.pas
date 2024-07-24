unit MainPopupMenu;

interface

uses
  System.SysUtils, System.Classes, Vcl.Menus, Vcl.ImgList, Vcl.Controls;

type
  TPMData = class(TDataModule)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PMData: TPMData;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

end.
