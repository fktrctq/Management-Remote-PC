unit HDDMount;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  MountHdd = class(TThread)
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
ResultHDD:array [0..2] of string=('������� Mount ������� ���������', '��� �������','����������� ������');




{ MountHdd }

procedure MountHdd.Execute;
var
resMount:integer;
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
            frmDomainInfo.memo1.Lines.Add('Mount -  '+(string(FWbemObject.caption)));
            resMount:=FWbemObject.Mount(); //// ���������� �����
            FWbemObject:=Unassigned;
           end;
      OleUnInitialize;  /// ���� ����� ���������� �� ����
      frmDomainInfo.memo1.Lines.Add('Mount - '+ResultHDD[resMount]);
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
