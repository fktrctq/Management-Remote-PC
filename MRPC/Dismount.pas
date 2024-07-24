unit Dismount;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  DismountVolume = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
const
ResultHDD:array [0..4] of string=('������� Dismount ������� ���������', '��� �������','Volume Has Mount Points','��� �� ������������ ��������� " No-Autoremount State"','Force Option Required');


{ DismountVolume }

procedure DismountVolume.Execute;
var
resDismount:integer;
begin
try
  begin
     OleInitialize(nil); ///// ����� ������ �����
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+SelectedHDD+'%"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            frmDomainInfo.memo1.Lines.Add('Dismount -  '+(string(FWbemObject.caption)));
            resDismount:=FWbemObject.Dismount(true,true); //// ���������� �����
            FWbemObject:=Unassigned;
           end;
      OleUnInitialize;  /// ���� ����� ���������� �� ����
      frmDomainInfo.memo1.Lines.Add('Dismount - '+ResultHDD[resDismount]);
  end;
 except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ Dismount "'+E.Message+'"');
          exit;
         end;
      end;
end;

end.
