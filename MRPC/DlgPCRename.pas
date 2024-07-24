unit DlgPCRename;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,System.Variants,ActiveX,ComObj;
type
 TPropForRenamePC = record
    NamePC :String;
    User:string;
    Passwd:string;
    NewNamePC:string;
    PasswordDomain:string;
    UserDomain:string;
    RebootAfterUnjoinDomain:bool;
  end;
type
  TOKRightDlg1234567 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    NamePCEdit: TLabeledEdit;
    UserAdmin: TLabeledEdit;
    PassAdmin: TLabeledEdit;
    RebootAfter: TCheckBox;
    procedure OKBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
   // function RenamePCFunction(NamePC,User,Passwd,NewNamePC,PasswordDomain,UserDomain:string;
//RebootAfterUnjoinDomain:bool):bool;
  public
    { Public declarations }
  end;
  function RenamePCFunction(ParamRenamePC:pointer):integer;
var
  OKRightDlg1234567: TOKRightDlg1234567;


implementation
uses umain;

ThreadVar
PointForFolder: ^TPropForRenamePC;
var
RunRenamePC:TPropForRenamePC;
{$R *.dfm}

function RenamePCFunction(ParamRenamePC:pointer):integer;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  ResRename:integer;
  FWbemObjectSet: OLEVariant;
  oEnum : IEnumvariant;
  iValue        : LongWord;
begin
try
  OleInitialize(nil);
  PointForFolder:=ParamRenamePC;
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer( PointForFolder.NamePC, 'root\CIMV2', PointForFolder.User, PointForFolder.Passwd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
    frmDomainInfo.memo1.Lines.Add('�������������� ����������....');
    if (PointForFolder.PasswordDomain='')and (PointForFolder.UserDomain<>'') then ResRename:=FWbemObject.Rename(PointForFolder.NewNamePC,null,PointForFolder.UserDomain);
    if (PointForFolder.UserDomain='')and (PointForFolder.PasswordDomain<>'') then ResRename:=FWbemObject.Rename(PointForFolder.NewNamePC,PointForFolder.PasswordDomain,null);
    if (PointForFolder.PasswordDomain='')and(PointForFolder.UserDomain='') then ResRename:=FWbemObject.Rename(PointForFolder.NewNamePC,null,null);
    if (PointForFolder.PasswordDomain<>'')and(PointForFolder.UserDomain<>'') then ResRename:=FWbemObject.Rename(PointForFolder.NewNamePC,PointForFolder.PasswordDomain,PointForFolder.UserDomain);

    if (PointForFolder.RebootAfterUnjoinDomain) and (ResRename<>5) then frmDomainInfo.rebuutSelectPC  //5 - �������� � �������
      else frmDomainInfo.memo1.Lines.Add('���������� ������������� ���������');
    frmDomainInfo.memo1.Lines.Add(PointForFolder.NamePC+'---�������������� ����������. ('+inttostr(ResRename)+') '+SysErrorMessage(ResRename));
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    FWbemObject:=Unassigned;
  end;
  VariantClear(FWbemObject);
  if oEnum<>nil then oEnum:=nil;
  FWbemObjectSet:=Unassigned;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
  except
  on E:Exception do
    begin
    frmDomainInfo.memo1.Lines.Add(PointForFolder.NamePC+'---������ �������������� ���������� "'+E.Message+'"');
    frmDomainInfo.memo1.Lines.Add('---------------------------');
    OleUnInitialize;
    exit;
    end;
  end;
NetApiBufferFree(ParamRenamePC); /// ������� ������
EndThread(0);    // ������� �����
end;

 //NewNamePC,UserDomain,PasswordDomain
procedure TOKRightDlg1234567.FormShow(Sender: TObject);
begin
UserAdmin.Text:=frmDomainInfo.LabeledEdit1.Text;
PassAdmin.Text:=frmDomainInfo.LabeledEdit2.Text;
end;

procedure TOKRightDlg1234567.OKBtnClick(Sender: TObject);
var
res:integer;
TreadIDR:LongWord;
begin
/// ������� ��������������
RunRenamePC.NamePC:=OKRightDlg1234567.Caption; // ����� ���� ���������������
RunRenamePC.User:= frmDomainInfo.LabeledEdit1.Text;       //  �����
RunRenamePC.Passwd:=frmDomainInfo.LabeledEdit2.Text;       // ������
RunRenamePC.NewNamePC:=NamePCEdit.Text;                        // ����� ���
RunRenamePC.UserDomain:= UserAdmin.Text;                        //�����
RunRenamePC.PasswordDomain:=PassAdmin.Text;                        // ������
RunRenamePC.RebootAfterUnjoinDomain:=RebootAfter.Checked;                  // ������������ ���� ����������
res:=BeginThread(nil,0,addr(RenamePCFunction),Addr(RunRenamePC),0,treadIDR); ///
 CloseHandle(res);
// � ������. unit - RenamePC
{NewNamePC:=NamePCEdit.Text;
UserDomain:=UserAdmin.Text;
PasswordDomain:=PassAdmin.Text;
NewRenamePC:=RenamePC.PCRename.Create(true);
NewRenamePC.FreeOnTerminate := true;  //// ��� �������� ��������������� ������
NewRenamePC.start;}

end;

end.
