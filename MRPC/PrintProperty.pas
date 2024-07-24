unit PrintProperty;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.ExtCtrls,ActiveX,ComObj, Vcl.Buttons;


type
  TPrinterProperty = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    CheckBox1: TCheckBox;
    LabeledEdit1: TLabeledEdit;
    CheckBox2: TCheckBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    LabeledEdit2: TLabeledEdit;
    LinkLabel1: TLinkLabel;
    Button4: TButton;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure Button3Click(Sender: TObject);
    procedure propertyselectprint;
    procedure CheckBox1Click(Sender: TObject);
    procedure PrintConfig;
    procedure FormShow(Sender: TObject);
    procedure ClearObject;
    function PrinterNetworkShared(s:boolean):string;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure LinkLabel1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    function Win32ACEDecompose(VarArray:variant):bool;
    function Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PrinterProperty: TPrinterProperty;
  DriverForSelectprinter:string;
  FormDriver:Tform;
  MemoDriver:Tmemo;
const
PrintCapabilities :array [0..21] of string =('����������','������','������� ������'
,'������������� ������','�����������','�������������','��������','���������� ������'
,'����������','������ �� ��������','�����������','�����-����� ������','�������������'
,'�������������, ������ �� �������� ����','�������������, ������ �� ��������� ����'
,'�������','������','�������� �������','�������� ������','������� ��������',
'���������� ��������','������ ��������');
PrintAvailability :array [1..21] of string =('������','����������','������ ��������'
,'��������������','� �����','�� ���������','����������','������� �����','��������'
,'����������','�� ������������','������ ���������','����������������-����������'
,'����������������-����� ������� �����������������','���������������� - ����� ��������'
,'���� �������','���������������� - ��������������','��������������','�� �����'
,'�� ��������','Quiesced (����� ����������)');
PrintConfigManagerErrorCode :array [0..31] of string =('���������� �������� ���������'
,'���������� ��������� �����������','��� ��������� ��� ����� ����������'
,'��������, ������� ��� ����� ���������� ���������, ��� � ����� ������� ������������ ������ ��� ������ ��������'
,'��� ���������� �� �������� ������� �������. ���� �� ��������� ��� ��� ������ ����� ���� ����������'
,'������� ��� ����� ���������� ��������� � �������, ������� Windows �� ����� ���������'
,'������������ �������� ��� ���������� ����������� � ������� ������������'
,'�� ������� �������������',
'��������� ��������� ��� ���������� �����������'
,'���������� �� �������� ������� �������. ����������� �������� ����������� �������� � �������� ��� ����������'
,'���������� �� ����� �����������','���������� �� �������','���������� �� ����� ����� ���������� ��������� �������� ��� �������������'
,'Windows �� ����� ��������� ������� ����������','���������� �� ����� �������� ������� �������, ���� ��������� �� ����� ������������'
,'���������� �� �������� ������� ������� ��-�� ��������� �������� � ��������� �������������'
,'Windows �� ����� ���������� ��� �������, ������� ���������� ����������.'
,'���������� ����������� ����������� ��� �������'
,'�������������� �������� ��� ����� ����������','���� ��� ������������� ���������� VxD'
,'������ ����� ���� ���������','���� �������: ���������� �������� ������� ��� ����� ����������.'
,'���������� ���������','���� �������: ���������� �������� ������� ��� ����� ����������'
,'���������� �����������, �������� ����������� ��� �� ����������� ��� ��� ��������'
,'Windows ��� ��� ����������� ����������','Windows ��� ��� ����������� ����������.'
,'���������� �� ����� ���������� ������������ �������','�������� ��������� �� �����������'
,'���������� ���������. �������� ���������� �� ���������� ��������� �������'
,'���������� ���������� ������ IRQ, ������� ���������� ������ ����������'
,'���������� �� �������� ������� �������. Windows �� ����� ��������� ����������� �������� ���������.'
);
PrintDetectedErrorState :array [0..11] of string =('����������','������','��'
,'���� ������',
'��� ������','���� ������','��� ������','������� ������','�������� ������'
,'�������  �����','���������� ������������','�������� ����� ��������');
printExtendedDetectedErrorState :array [0..15] of string =('����������� ������'
,'������','������ ���','���� ������','��� ������','���� ������','��� ������'
,'������� ������ ��������','�������','Service Requested','�������� ����� ��������'
,'�������� � �������','�� ������� ����������� ��������','��������� ������������� ������������'
,'������������ ������','������ ����������');
PrintExtendedPrinterStatus :array [1..18] of string = ('����������','����������'
,'�������� (������ Idle)','�����','��������','Stopped Printing','Offline','������ ��������������'
,'������','�����','�� ���������','����� ��������','���������','�������������'
,'����������������','� �������� ��������','I/O Active','������ ������');
PrintLanguagesSupported :array [1..50] of string =('Other','Unknown','PCL ','HPGL'
,'PJL','PS','PSPrinter','IPDS','PPDS','EscapeP','Epson','DDIF','Interpress'
,'ISO6429','Line Data','MODCA','REGIS','SCS','SPDL','TEK4014','PDS','IGP','CodeV'
,'DSCDSE','WPS','LN03','CCITT','QUIC','CPAP','DecPPL','Simple Text','NPAP','DOC','imPress'
,'Pinwriter','NPDL','NEC201PL','Automatic','Pages','LIPS','TIFF','Diagnostic'
,'CaPSL','EXCL','LCDS','XES','MIME','XPS','HPGL2','PCLXL');
PrintMarkingTechnology :array [1..27] of string=('Other','�����������','���������������������� ���������','����� ����������������������'
,'����������������������','��������� �� 9pin','��������� �� 24pin','������� �������� ���������� ������� � ������� ��������'
,'������� ���������� ������� ','Impact Band','Impact Other','�������� ������','�������� �������','�������� ������'
,'Pen','����������� ��������','�������������������','�������������','�������� ������','�������������','������������������'
,'Photographic Microfiche','Photographic Imagesetter','Photographic Other','������ ���������','eBeam','Typesetter');
PrintPrinterStatus  : array [1..7] of string =('Other','����������� �����','����� Idle','����� ������','��������','��������� ������','�� � ����');


implementation
uses umain,PrintSecuretyEdit,AddPrintLan;



{$R *.dfm}
function VarToInt(AVariant: variant): integer;
begin
  //*** ���� NULL ��� �� ��������, �� ������ �������� �� ���������
  Result := 0;
  if VarIsNull(AVariant) then
    Result := 0
  else
    {//*** ���� ��������, �� ������ ��������}
     if VarIsOrdinal(AVariant) then
      Result := StrToInt(VarToStr(AVariant));
end;

function TPrinterProperty.PrinterNetworkShared(s:boolean):string;
var
FSWbemLocator    : OLEVariant;
  FWMIService    : OLEVariant;
  FWbemObjectSet : OLEVariant;
  FWbemObject    : OLEVariant;
  oEnum          : IEnumvariant;
  MyError,i    : integer;
begin
try
begin
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
     while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
        FWbemObject.Shared:=CheckBox1.Checked;
        FWbemObject.Published:=CheckBox2.Checked;
        FWbemObject.ShareName:=LabeledEdit1.Text;
        FWbemObject.Location:=LabeledEdit2.Text;
        result:=vartostr(FWbemObject.put_());
      end;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ �������� SharedNetwork "'+E.Message);
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;
end;
    If oEnum<>nil then oEnum:=nil;
    VariantClear(FWbemObject);
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
    OleUnInitialize;
end;




procedure TPrinterProperty.ClearObject;
begin
memo1.Clear;
LabeledEdit1.Text:='';
Label1.Caption:='����:';
Label2.Caption:='�������:';
Label3.Caption:='��������� ������� ������:';
Label4.Caption:='������� ������:';
LinkLabel1.Caption:='';
CheckBox1.Checked:=false;
CheckBox2.Checked:=false;
DriverForSelectprinter:='';
end;


procedure TPrinterProperty.CheckBox1Click(Sender: TObject);
begin
LabeledEdit1.Enabled:=CheckBox1.Checked;
CheckBox2.enabled:=CheckBox1.Checked;
if CheckBox1.Checked then LabeledEdit1.Text:=SelectedPrint
else
begin
LabeledEdit1.Text:='';
CheckBox2.Checked:=false;
end;

end;




procedure TPrinterProperty.FormClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as Tform).FreeOnRelease;
end;

procedure TPrinterProperty.FormShow(Sender: TObject);
begin
MyPS:=frmDomainInfo.ComboBox2.text;
MyUser:=frmDomainInfo.LabeledEdit1.text;
MyPasswd:=frmDomainInfo.LabeledEdit2.text;
ClearObject;
propertyselectprint;
//PrintConfig;
end;

procedure TPrinterProperty.LinkLabel1Click(Sender: TObject);
var
 FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  massArray: variant;
  i:integer;
begin
try
begin
FormDriver:=TForm.Create(PrinterProperty);
FormDriver.Caption:='��������: '+ DriverForSelectprinter;
FormDriver.Height:=400;
FormDriver.Width:=400;
FormDriver.Position:=poMainFormCenter;
FormDriver.BorderStyle:=bsDialog;
FormDriver.OnClose:=FormClose;
MemoDriver:=Tmemo.Create(FormDriver);
MemoDriver.parent:=FormDriver;
MemoDriver.Align:=alClient;
MemoDriver.ScrollBars:=ssVertical;
MemoDriver.Clear;
FormDriver.Show;
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_PrinterDriver WHERE Name LIKE '+'"%'+DriverForSelectprinter+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
     while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
      MemoDriver.Lines.Add('_________________________________________________');
      if FWbemObject.Name<>null then
      MemoDriver.Lines.Add('��� ��������: '+string(FWbemObject.Name));
      if FWbemObject.Version<>null then
     MemoDriver.Lines.Add('������ ��������: '+string(FWbemObject.Version));
      if FWbemObject.SupportedPlatform<>null then
      MemoDriver.Lines.Add('���������: '+string(FWbemObject.SupportedPlatform));
      if FWbemObject.OEMUrl<>null then
      MemoDriver.Lines.Add('OEMUrl: '+string(FWbemObject.OEMUrl));
      if FWbemObject.ConfigFile<>null then
      MemoDriver.Lines.Add('���� ������������: '+string(FWbemObject.ConfigFile));
      if FWbemObject.DataFile<>null then
      MemoDriver.Lines.Add('���� ������: '+string(FWbemObject.DataFile));
      if FWbemObject.DriverPath<>null then
     MemoDriver.Lines.Add('DriverPath: '+string(FWbemObject.DriverPath));
      if FWbemObject.FilePath<>null then
      MemoDriver.Lines.Add('FilePath: '+string(FWbemObject.FilePath));
      if FWbemObject.HelpFile<>null then
      MemoDriver.Lines.Add('HelpFile: '+string(FWbemObject.HelpFile));
      if VarIsArray(FWbemObject.DependentFiles) then
            begin
            massArray:=FWbemObject.DependentFiles;
            MemoDriver.Lines.Add('��������� �����__________________________________');
            for i := 0 to VarArrayHighBound(massArray, 1) do
              begin
              MemoDriver.Lines.add(vartostr(massArray[i]));
              end;
            MemoDriver.Lines.Add('_________________________________________________');
            VarClear(massArray);
            end;
    FWbemObject:=Unassigned;
    end;
end;
except
  on E:Exception do
           begin
             frmDomainInfo.memo1.Lines.Add('������ �������� '+E.Message);
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             OleUnInitialize;
           end;
  end;

 If oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TprinterProperty.PrintConfig;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
begin
try
begin;
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_PrinterConfiguration WHERE Description LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
     while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
      //Application.ProcessMessages;
      //////////////////////////////////////////////////////
      if FWbemObject.Color<>null then
        begin
        if vartoint(FWbemObject.Color)=1 then  label4.Caption:='������� ������: ���'
        else label4.Caption:='������� ������: ��';
        end;
     { if FWbemObject.DriverVersion<>null then
        begin
        memo1.Lines.add('������ �������� FWbemObject.DriverVersion: '+string(FWbemObject.DriverVersion));
        memo1.Lines.add('------------------------------------------------------');
        end;}
      if FWbemObject.duplex<>null then
        begin
        if boolean(FWbemObject.duplex) then memo1.Lines.add('������������� ������: ��')
        else memo1.Lines.add('������������� ������: ���');
        end;
       if (FWbemObject.HorizontalResolution<>null)and (FWbemObject.VerticalResolution<>null) then
        begin
        memo1.Lines.add('���������� ������ (�����/����): '+string(FWbemObject.HorizontalResolution)+'x'+string(FWbemObject.VerticalResolution));
        memo1.Lines.add('------------------------------------------------------');
        end;
        //////////////////////////////////////////////////////////
        if VarIsNumeric(FWbemObject.ICMMethod) then
        begin
          case vartoint(FWbemObject.ICMMethod) of
          1:memo1.Lines.add('����� ���������� ������: �������� ');
          2:memo1.Lines.add('����� ���������� ������: Windows ');
          3:memo1.Lines.add('����� ���������� ������: �� ������ �������� ���������� ');
          4:memo1.Lines.add('����� ���������� ������: ���������� ');
          else memo1.Lines.add('����� ���������� ������: ���������� ');
          end;

        end;
       if VarIsNumeric(FWbemObject.ICMIntent) then
          begin
          case vartoint(FWbemObject.ICMIntent) of
          1:memo1.Lines.add('����� ������������� ������: ��������� ');
          2:memo1.Lines.add('����� ������������� ������: ��������');
          3:memo1.Lines.add('����� ������������� ������: ������ ����');
          else memo1.Lines.add('����� ���������� ������: ���������� ');
          end;
        end;
       ///////////////////////////////////////////////////////////
         if VarIsNumeric(FWbemObject.MediaType) then
        begin
          case vartoint(FWbemObject.MediaType) of
          1:memo1.Lines.add('��� ��������, �� ������� �������� �������: �������� ');
          2:memo1.Lines.add('��� ��������, �� ������� �������� �������: Transparency ');
          3:memo1.Lines.add('��� ��������, �� ������� �������� �������: ���������');
          else memo1.Lines.add('����� ���������� ������: ���������� ');
          end;
        end;
       //////////////////////////////////////////////////////////////////////
       if VarIsNumeric(FWbemObject.Orientation) then
        begin
          case vartoint(FWbemObject.Orientation) of
          1:memo1.Lines.add('���������� ������ ��� ������: ������� ');
          2:memo1.Lines.add('���������� ������ ��� ������: ������');
          else memo1.Lines.add('���������� ������ ��� ������: ��, ����� ��� ��! ');
          end;
        end;
       ///////////////////////////////////////////////////////////////////////////////
     FWbemObject:=Unassigned;
      end;
end;
  except
  on E:Exception do
           begin
             frmDomainInfo.memo1.Lines.Add('������ �������� "'+E.Message);
             frmDomainInfo.memo1.Lines.Add('---------------------------');
             OleUnInitialize;
           end;
  end;

If oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;





procedure TPrinterProperty.propertyselectprint;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  masvar: variant;
  i:integer;
  PrintPropertyString:string;
begin
   try
begin
  memo1.clear;
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
     while oEnum.Next(1, FWbemObject, iValue) = 0 do
      begin
      //////////////////////////////////////////////////////
      if FWbemObject.DriverName<>null then
        begin
        LinkLabel1.Caption:='<a>'+string(FWbemObject.DriverName)+'</a>';
        DriverForSelectprinter:=string(FWbemObject.DriverName);
        end
        else LinkLabel1.Caption:='����������';
        ////////////////////////////////////////////
         if FWbemObject.PortName<>null then
        begin
         Label1.Caption:='����: '+string(FWbemObject.PortName);
        end
        else Label1.Caption:='����: ����������';
              //////////////////////////////////////////////////////
        if FWbemObject.PrintProcessor<>null then
        begin
         Label3.Caption:='��������� ������� ������: '+string(FWbemObject.PrintProcessor);
        end else Label3.Caption:='��������� ������� ������: ����������';
      ////////////////////////////////////////////////////////
     if (FWbemObject.ServerName<>null)  then
      memo1.Lines.add('��� �������, ������� ��������� ��������� (FWbemObject.ServerName): '+string(FWbemObject.ServerName));
      /////////////////////////////////////////////////////////////////////////////////
     {    if FWbemObject.Hidden<>null then //���� TRUE , ������� ����� �� ������������� ����.
     if boolean(FWbemObject.Hidden)=false then
     begin
      memo1.Lines.add('������� �������� ��� ����������� ������ (FWbemObject.Hidden):'+vartostr(FWbemObject.Hidden));
      memo1.Lines.add('------------------------------------------------------');
     end; }
     ////////////////////////////////////////////////////////////////////////
      //Shared ������� �������� ��� ������� ������
      CheckBox1.Checked:=boolean(FWbemObject.Shared);
      LabeledEdit1.Enabled:=CheckBox1.Checked;
      CheckBox2.enabled:=CheckBox1.Checked;
      ////////////////////////////////////////////////////////////////////////////
      if (FWbemObject.ShareName<>null)  then
     begin
     // memo1.Lines.add('��� ���������� ��� ������������ ������� (FWbemObject.ShareName): '+string(FWbemObject.ShareName));
     // memo1.Lines.add('------------------------------------------------------');
      LabeledEdit1.Text:=string(FWbemObject.ShareName);
     end;
      ///////////////////////////////////////////////////////////////
     { if (FWbemObject.status<>null)  then
      memo1.Lines.add('������� ������ �������� (FWbemObject.status): '+string(FWbemObject.status));
       /// ///////////////////////////////////////////////////
     if (FWbemObject.network<>null)  then
      memo1.Lines.add('������� ������� (FWbemObject.network): '+booltostr(FWbemObject.network));
     if ((FWbemObject.local)<>null) then
      memo1.Lines.add('FWbemObject.local: '+booltostr(FWbemObject.local));
      memo1.Lines.add('------------------------------------------------------');}
     /////////////////////////////////////////////////////////////////////
     if FWbemObject.Published<>null then
     begin
      CheckBox2.Checked:= boolean(FWbemObject.Published);
     end;
////////////////////////////////////////////////////////////// ������������
     if FWbemObject.Location<>null then LabeledEdit2.Text:=string(FWbemObject.Location)
     else LabeledEdit2.Text:='';
       //////////////////////////////////////////////////////////
      { masvar:=FWbemObject.Capabilities;
       PrintPropertyString:='';
       for i := 0 to VarArrayHighBound(masvar, 1) do
        begin
        PrintPropertyString := PrintPropertyString + PrintCapabilities[vartoint(masvar[i])] + ', ';
        end;
       memo1.lines.add('�������������� �������: '+PrintPropertyString);
       VarClear(masvar);
       memo1.Lines.add('------------------------------------------------------'); }
      //////////////////////////////////////////////
      {if FWbemObject.Availability<>null then
         memo1.lines.add('�����������: '+(PrintAvailability[integer(FWbemObject.Availability)]));
         memo1.Lines.add('------------------------------------------------------'); }
      ////////////////////////////////////////////////
      if FWbemObject.ConfigManagerErrorCode<>null then
         memo1.lines.add('������: '+PrintConfigManagerErrorCode[integer(FWbemObject.ConfigManagerErrorCode)]);
         /////////////////////////////////////////////
       if (FWbemObject.DetectedErrorState<>null)  then
         begin
          if integer(FWbemObject.DetectedErrorState)<>0 then
          memo1.lines.add('��������� ��������: '+PrintDetectedErrorState[integer(FWbemObject.DetectedErrorState)]);
         end;
        //////////////////////////////////////////////////
        if (FWbemObject.ExtendedDetectedErrorState<>null)  then
          begin
           if (integer(FWbemObject.ExtendedDetectedErrorState)<>0) then
           memo1.lines.add('���������� �� �������: '+PrintExtendedDetectedErrorState[integer(FWbemObject.ExtendedDetectedErrorState)]);
          end;
          /////////////////////////////////////////////////
          if (FWbemObject.ExtendedPrinterStatus<>null)  then
          begin
          if (integer(FWbemObject.ExtendedPrinterStatus)<>0) then
           memo1.lines.add('���������� � ������� ���������: '+PrintExtendedPrinterStatus[integer(FWbemObject.ExtendedPrinterStatus)]);
          end;
         ///////////////////////////////////////////
         if VarIsArray(FWbemObject.LanguagesSupported) then
           begin
           masvar:=FWbemObject.LanguagesSupported;
           PrintPropertyString:='';
           for i := 0 to VarArrayHighBound(masvar, 1) do
            begin
            PrintPropertyString := PrintPropertyString + PrintLanguagesSupported[integer(masvar[i])] + ', ';
            end;
            memo1.lines.add('�������������� ����� ������: '+PrintPropertyString);
            VarClear(masvar);
           end;
          ///////////////////////////////////////
          if (FWbemObject.MarkingTechnology<>null)  then
          begin
           if (integer(FWbemObject.MarkingTechnology)<>0) then
           memo1.lines.add('���������� ������: '+PrintMarkingTechnology[integer(FWbemObject.MarkingTechnology)]);
          end;
          //////////////////////////////////////////////////
          if VarIsArray(FWbemObject.PrinterPaperNames) then
            begin
            masvar:=FWbemObject.PrinterPaperNames;
            PrintPropertyString:='';
            for i := 0 to VarArrayHighBound(masvar, 1) do
              begin
              PrintPropertyString := PrintPropertyString + vartostr(masvar[i]) + ', ';
              end;
            memo1.Lines.Add('_________________________________________________');
            memo1.lines.add('�������������� ������� ������: '+PrintPropertyString);
            memo1.Lines.Add('_________________________________________________');
            VarClear(masvar);
            end;
          ////////////////////////////////////////////////////////////
         if (FWbemObject.HorizontalResolution<>null) and (FWbemObject.VerticalResolution<>null) then
          begin
          memo1.lines.add('���������� ��������: '+string(FWbemObject.HorizontalResolution)+'x'+string(FWbemObject.VerticalResolution));
          end;
          /// /////////////////////////////////////////////////////////
          if FWbemObject.MaxCopies<>null then
          memo1.lines.add('������������ ���������� ����� (FWbemObject.MaxCopies): '+(vartostr(FWbemObject.MaxCopies)));
         ////////////////////////////////////////////////////////////////
          if FWbemObject.MaxNumberUp<>null then
          memo1.lines.add('������������ ���������� ������� �� ����� (FWbemObject.MaxNumberUp): '+(vartostr(FWbemObject.MaxNumberUp)));
        end;
    FWbemObject:=Unassigned;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������ ��������:'+E.Message);
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
         end;

end;
  If oEnum<>nil then oEnum:=nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TPrinterProperty.SpeedButton1Click(Sender: TObject);
begin
PrintConfig;
end;

procedure TPrinterProperty.SpeedButton2Click(Sender: TObject);
begin
propertyselectprint;
end;


procedure TPrinterProperty.Button1Click(Sender: TObject);
begin
close;
end;

procedure TPrinterProperty.Button2Click(Sender: TObject);
var
i:string;
begin
i:=PrinterNetworkShared(CheckBox1.Checked);
frmDomainInfo.memo1.Lines.Add('��������� �������� PrinterNetworkShared: '+i);
end;

procedure TPrinterProperty.Button3Click(Sender: TObject);
begin

//PrintConfig;
end;

function TPrinterProperty.Win32TrusteeDecompose(VarTrustee:Olevariant):bool;
var
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
ACETrustee    : Olevariant;
i:integer;
SIDforStr     : string;
begin
SIDforStr:='';
result:=false;
if not(varisnull(VarTrustee)) then
  begin
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  ACETrustee:=FWMIService.get('WIN32_Trustee');
  ACETrustee:=VarTrustee;
  if not(varisnull(ACETrustee.TIME_CREATED)) then memo1.Lines.Add(string(ACETrustee.TIME_CREATED));
  if not(varisnull(ACETrustee.Domain)) then memo1.Lines.Add(string(ACETrustee.Domain));
  if not(varisnull(ACETrustee.name)) then memo1.Lines.Add(string(ACETrustee.name));
  if not(varisnull(ACETrustee.SidLength)) then memo1.Lines.Add(string(ACETrustee.SidLength));
  if not(varisnull(ACETrustee.SIDString)) then memo1.Lines.Add(string(ACETrustee.SIDString));
  if varisarray(ACETrustee.sid) then
  begin
    for I := 0 to VarArrayHighBound(ACETrustee.sid, 1) do
    begin
     SIDforStr:=SIDforStr+string(ACETrustee.sid[i])+'-';
    end;
  memo1.Lines.Add(SIDforStr);
  end;

  OleUnInitialize;
  result:=true;
  end;
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
VariantClear(ACETrustee);
VariantClear(VarTrustee);
end;

function TPrinterProperty.Win32ACEDecompose(VarArray:variant):bool;
var
i:integer;
FSWbemLocator : OLEVariant;
FWMIService   : OLEVariant;
ACETrustee    : OLEVAriant;
Win32ACE      : OleVariant;
begin
result:=false;
OleInitialize(nil);
FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
FWMIService.Security_.impersonationlevel:=3;
FWMIService.Security_.authenticationLevel := 6;
ACETrustee:=FWMIService.get('WIN32_Trustee');
Win32ACE:= FWMIService.get('WIN32_ACE');
for i := 0 to VarArrayHighBound(VarArray, 1) do
begin
  memo1.Lines.Add('���� VarIsArray(VarArray['+inttostr(i)+'])');
  Win32ACE:=varArray[i];

  ACETrustee:=Win32ACE.Trustee;
  Win32TrusteeDecompose(ACETrustee);
  if not(VarIsNull(Win32ACE.TIME_CREATED)) then
  Memo1.Lines.Add('����� ��������:'+string(Win32ACE.TIME_CREATED));
  if not(VarIsNull(Win32ACE.AccessMask)) then
  Memo1.Lines.Add('����� ������� (����� ����):'+string(Win32ACE.AccessMask));
  if not(VarIsNull(Win32ACE.AceFlags)) then
  Memo1.Lines.Add('������������ ����:'+string(Win32ACE.AceFlags));
  if not(VarIsNull(Win32ACE.AceType)) then
  Memo1.Lines.Add('������ 0/1/2:'+string(Win32ACE.AceType));
  if not(VarIsNull(Win32ACE.GuidInheritedObjectType)) then
  Memo1.Lines.Add('Guid ��������:'+string(Win32ACE.GuidInheritedObjectType));
  if not(VarIsNull(Win32ACE.GuidObjectType)) then
  Memo1.Lines.Add('Guid �������:'+string(Win32ACE.GuidObjectType));



  VariantClear(Win32ACE);
  Result:=true;
end;
VariantClear(Win32ACE);
VariantClear(ACETrustee);
VariantClear(FWMIService);
VariantClear(FSWbemLocator);
OleUnInitialize;
end;

procedure TPrinterProperty.Button4Click(Sender: TObject);
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  PrintSecurityDescriptor:OLEVariant;
  SDACEDACLArray       :variant;
  SDACESACLArray       :Variant;
  SDTrusteeGroup  :Olevariant;
  SDTrusteeOwner  :Olevariant;
  NewSecurityDescriptor: pSecurityDescriptor;
  oEnum : IEnumvariant;
  masvar: variant;
  SDescriptorError:integer;
  PrintPropertyString:string;



begin
PrintSecuretyEditor.ShowModal;
{try
begin
  memo1.clear;
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6;
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Printer WHERE DeviceID LIKE '+'"%'+SelectedPrint+'%"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;

  PrintSecurityDescriptor:=FWMIService.get('Win32_SecurityDescriptor');
  //SDACEDACLArray:= FWMIService.get('Win32_ACE');
  //SDACESACLArray:= FWMIService.get('Win32_ACE');
  SDTrusteeGroup:= FWMIService.get('Win32_Trustee');
  SDTrusteeOwner:= FWMIService.get('Win32_Trustee');

     if oEnum.Next(1, FWbemObject, iValue) = 0 then
      begin
      SDescriptorError:=FWbemObject.GetSecurityDescriptor(PrintSecurityDescriptor);

      if VarIsArray(PrintSecurityDescriptor.dacl) then
      begin
      SDACEDACLArray:=PrintSecurityDescriptor.dacl;
      if Win32ACEDecompose(SDACEDACLArray)then memo1.Lines.add('Win32ACEDecompose(SDACEDACLArray)');
      end
        else Memo1.Lines.Add('PrintSecurityDescriptor.dacl �� ������');

      if VarIsArray(PrintSecurityDescriptor.sacl) then
      begin
      SDACESACLArray:=PrintSecurityDescriptor.sacl;
      if Win32ACEDecompose(SDACESACLArray)then memo1.Lines.add('Win32ACEDecompose(SDACESACLArray)');
      end
        else Memo1.Lines.Add('PrintSecurityDescriptor.sacl �� ������');

      SDTrusteeGroup:=PrintSecurityDescriptor.group;
      if Win32TrusteeDecompose(SDTrusteeGroup) then Memo1.Lines.Add('��������� ������ SDTrusteeGroup');

      SDTrusteeOwner:=PrintSecurityDescriptor.owner;
      if Win32TrusteeDecompose(SDTrusteeOwner) then Memo1.Lines.Add('��������� ������ SDTrusteeOwner');


      ///BinarySDToSecurityDescriptor(NewSecurityDescriptor,PrintSecurityDescriptor,
     /// pchar(MyPS),pchar(MyUser),pchar(MyPasswd),0);
     // GetSecurityDescriptorDacl(NewSecurityDescriptor,true,null,true);
      end;
   memo1.Lines.Add('��������� �������� SysErrorMessage: '+  SysErrorMessage(SDescriptorError));

end;
except
on E:Exception do
         begin

           memo1.Lines.Add('������ �������� SysErrorMessage: '+  SysErrorMessage(SDescriptorError));
           memo1.Lines.Add('������ �������� PrintSecurityDescriptor: '+E.Message);
           memo1.Lines.Add('---------------------------');
           OleUnInitialize;
         end;

end;
 If oEnum<>nil then oEnum:=nil;
 VarClear(SDACEDACLArray);
 VarClear(SDACESACLArray);
 VariantClear(SDTrusteeGroup);
 VariantClear(SDTrusteeOwner);
 VariantClear(PrintSecurityDescriptor);
 VariantClear(FWbemObject);
 VariantClear(FWbemObjectSet);
 VariantClear(FWMIService);
 VariantClear(FSWbemLocator);
 OleUnInitialize; }
end;

end.
