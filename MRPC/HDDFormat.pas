unit HDDFormat;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  FormatHDD = class(TThread)
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
ResultHDD:array [0..18] of string=('������� Format ������� ���������',
'�������� ������� �� ��������������','������������� �������� ��������','������ ��������',
'����� �������','������ �� ������ ������ ����������','Volume write protected','���� ���������� ����','���������� ������� ��������������',
'������ �����/������ (I/O)','������������ ����� ����','��� �������� � �����',
'��� ������� ���','������� ������� ����� ����','��� �� ���������� (Not Mounted)'
,'��������� ������ ��������','������� ������ ��������','������ �������� ��������� 32 ����','����������� ������');

{ FormatHDD }
    ///FormatClusterSize,,FormatFileSystem,FormatLabel,FormatQuickFormat,FormatEnableCompression
procedure FormatHDD.Execute;
var
resFormat:integer;
begin
try
  begin
     OleInitialize(nil);
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Volume WHERE DeviceID LIKE '+'"%'+SelectedHDD+'%"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            frmDomainInfo.memo1.Lines.Add('�������������� ����� -  '+(string(FWbemObject.caption)));
            resFormat:=FWbemObject.Format(FormatFileSystem,FormatQuickFormat,FormatClusterSize,FormatLabel,FormatEnableCompression); //// ���������� �����
            frmDomainInfo.memo1.Lines.Add('�������������� ����� - '+(string(FWbemObject.caption))+'. '+ResultHDD[resFormat]);
            FWbemObject:=Unassigned;
           end;
      VariantClear(FWbemObject);
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;

  end;
 except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ �������������� "'+E.Message+'"');
          exit;
         end;
      end;
end;

end.
