unit Unit6;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Service = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses uMain;

{ Service }

procedure Service.Execute;
var
ResServic:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
FWbemObjectSet: OLEVariant;
FWbemObject   : OLEVariant;
oEnum         : IEnumvariant;
iValue        : LongWord;
Const OWN_PROCESS = 16;
 NOT_INTERACTIVE = True;
 ErrorCodeService: array [0..24] of string=('�������� ������� ���������',
 '�������� �� ��������������','� ������������ ��� ������������ �������',
 '������ ���������� ����������, ��������� �� ��� ������� ������ ���������� ������',
 '����������� ����������� ��� �������������� ��� ���������� ��� ������.'
 ,'����������� ��� �������� ���������� �� ����� ���� ��������� ������ ��-�� ��������� ������ '
 ,'������ �� ��������','������ �� �������� �� ������ ������������',
 '����������� ���� ��� ������� ������','���� � ������������ ����� ������ �� ������'
 ,'������ ��� ��������.','���� ������ ��� ���������� ����� ������ �������������',
 '�����������, �� ������� ��������� ��� ������, ���� ������� �� �������.',
 '������ �� ������� ����� ������, ����������� ��� ��������� ������.',
 ' ������ ���� ��������� �� �������.','������ �� ����� ���������� �������� ����������� ��� ������� � �������.'
 ,'��� ������ ��������� �� �������','������ �� ����� ������ ����������',' ��� ������� ������ ����� ����������� �����������'
 ,'������ ����������� ��� ��� �� ������','��� ������ �������� ������������ �������'
 ,'������ �������� ������������ ���������.','������� ������, ��� ������� ����������� ��� ������, �������� ������������ ��� �� ����� ���������� �� ������ ������.'
 ,'������ ���������� � ���� ������ ��������, ��������� �� �������.',
 '� ��������� ����� ������ �������������� � �������');
begin
      try
      begin
        OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
      begin
      if ActionServic=3 then FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Service','WQL',wbemFlagForwardOnly)
      else FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Service WHERE Name = '+'"'+SelectServic+'"','WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            frmDomainInfo.memo1.Lines.Add('---------------------------');
            case ActionServic of
            ///////////////////
            1:
            begin
            frmDomainInfo.memo1.Lines.Add('������ ������...');
            ResServic:=FWbemObject.StartService();  ////// ��������� ������
            FWbemObject:=Unassigned;
            end;
            /////////////////
            2:
            begin
            frmDomainInfo.memo1.Lines.Add('��������� ������...');
            ResServic:=FWbemObject.StopService();  ////// ���������� ������
            FWbemObject:=Unassigned;
            end;
            3:
            begin
            //ResServic:=FWbemObject3.create;//('"DbService"', '"Personnel Database"', '"C:\Program Files (x86)\2gis\3.0\2GISUpdateService.exe"',OWN_PROCESS ,1 ,'"Automatic"' ,NOT_INTERACTIVE,'".\LocalSystem"','');  ////// ������� ������
            //FWbemObject3:=Unassigned;
            end;
            //////////////////
            4:
            begin
            frmDomainInfo.memo1.Lines.Add('����� ���� ������� ������...');
            ResServic:=FWbemObject.ChangeStartMode(TypeRunService); //// ����� ���� �������
            FWbemObject:=Unassigned;
            end;
            5:
            begin
            frmDomainInfo.memo1.Lines.Add('�������� ������...');
            ResServic:=FWbemObject.Delete(); //// ��������� ������
            FWbemObject:=Unassigned;
            end;
            end;

            end;
         if ResServic>24 then frmDomainInfo.memo1.Lines.Add('C����� - '+SysErrorMessage(ResServic))
          else frmDomainInfo.memo1.Lines.Add('C����� - '+ErrorCodeService[ResServic]);


         frmDomainInfo.memo1.Lines.Add('---------------------------');
          VariantClear(FWbemObject);
          oEnum:=nil;
          VariantClear(FWbemObjectSet);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
         OleUnInitialize;  /// ���� ����� ���������� �� ����
      end;


      end;
      except
        on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ "'+E.Message+'"');
          exit;
         end;
      end;

end;





end.
