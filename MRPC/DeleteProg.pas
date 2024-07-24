unit DeleteProg;

interface

uses
   System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  DeleteProgram = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;


{ DeleteProgram }

procedure DeleteProgram.Execute;
var
  FSWbemLocator   : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResDelProgram: integer;
  i:integer;
  CurrentListProgram:TstringList;
  CurrentPC,CurrentUser,CurrentPass:string;
begin
  try
  CurrentListProgram:=Tstringlist.Create;
  CurrentListProgram.Text:=listDelProg.Text;
  listDelProg.Free;
  CurrentPC:=MyPS;
  CurrentUser:=MyUser;
  CurrentPass:=MyPasswd;
  for I := 0 to CurrentListProgram.Count-1 do
    BEGIN
    try
      Begin
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      frmDomainInfo.memo1.Lines.Add('Запущен процесс удаления программы '+timetostr(time)+':'+CurrentListProgram[i]);
      OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(CurrentPC, 'root\CIMV2', CurrentUser, CurrentPass);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Product WHERE Caption = '+'"'+CurrentListProgram[i]+'"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
        begin
        ResDelProgram:=FWbemObject.Uninstall(); //// Удалениме программы
        FWbemObject:=Unassigned;
        end;
      frmDomainInfo.memo1.Lines.Add(CurrentPC+' Удаление программы '+CurrentListProgram[i] +' завершено. '+timetostr(time)+' : '+SysErrorMessage(ResDelProgram));
      frmDomainInfo.memo1.Lines.Add('---------------------------');
      VariantClear(FWbemObject);
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
      End;
      except
      on E:Exception do
      frmDomainInfo.memo1.Lines.Add(CurrentPC+' Ошибка удаления программы '+CurrentListProgram[i] +' "'+E.Message+'"');
      end;
    END;

  except
  on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add(CurrentPC+' Ошибка удаления программ - "'+E.Message+'"');
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    OleUnInitialize;
    CurrentListProgram.Free;
    end;
  end;

CurrentListProgram.Free;
end;

end.
