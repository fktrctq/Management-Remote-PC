unit Unit7;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Newservice = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses uMain,NewServic;






{ Newservice }   /// �������� ����� ������

procedure Newservice.Execute;
var
ResServic:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObject   : OLEVariant;
begin
     try
      begin
      OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIService  := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text, 'root\CIMV2', frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text,'','',128);
     FWbemObject   := FWMIService.Get('Win32_Service');
     ResServic:=FWbemObject.create(NewServName, NewServDispName, NewServPathName,NewServType ,1 ,NewServStartMode ,NewServDesktop,NewServStartName,NewServStartPassword);  ////// ������� ������
     FWbemObject:=Unassigned;
     frmDomainInfo.memo1.Lines.Add('�������� ������. '+SysErrorMessage(ResServic));
     VariantClear(FWbemObject);
     VariantClear(FWMIService);
     VariantClear(FSWbemLocator);
     OleUnInitialize;
      end;
    except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ "'+E.Message+'"');
           OleUnInitialize;
          exit;
         end;
      end;


end;

end.
