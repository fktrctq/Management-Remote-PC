unit DlgJoinDomainOrWorkGroup;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls,ActiveX,ComObj,CommCtrl,System.Variants;
 type
 TPropJoinDomWG = record
    MyPS :String;
    MyUser:string;
    MyPasswd:string;
    NameDomain:string;
    PasswordDomain:string;
    UserDomain:string;
    Oper:integer;
    RebootAfterUnjoinDomain:bool;
  end;
type
  TOKRightDlg123456 = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    ComboBox1: TComboBox;
    RebootAfter: TCheckBox;
    NumOper: TEdit;
    procedure OKBtnClick(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
  private
    function summOperation(const z:integer):integer;
    {function JoinDomainOrWorkGroup(MyPS,MyUser,MyPasswd,NameDomain,
    PasswordDomain,UserDomain:string;Oper:integer;RebootAfterUnjoinDomain:bool):bool;}

  public
    { Public declarations }
  end;
function JoinDomainOrWorkGroup(paramUnJwG:pointer):integer;

var
  OKRightDlg123456: TOKRightDlg123456;
Const JOIN_DOMAIN = 1 ; // ���������� � ������
Const ACCT_CREATE = 2 ; // ������� ������
Const ACCT_DELETE = 4 ; // ������ �� ������
Const WIN9X_UPGRADE = 16 ; // ��������
Const DOMAIN_JOIN_IF_JOINED = 32 ;  //��������� �������������� � ������ ������, ���� ���� ��������� ��� ����������� � ������.
Const JOIN_UNSECURE = 64  ;        //��������� ������������ ����������.
Const MACHINE_PASSWORD_PASSED = 128 ;
Const DEFERRED_SPN_SET = 256;
Const INSTALL_INVOCATION = 262144;

implementation
uses umain,PCJoinDomainOrWorkGroup;
ThreadVar
PointForFolder: ^TPropJoinDomWG;
var
RunUnjPC:TPropJoinDomWG;
{$R *.dfm}

function TOKRightDlg123456.summOperation(const z:integer):integer;
begin
case ComboBox1.ItemIndex of
0:result:=JOIN_DOMAIN+ACCT_CREATE;
1:result:=JOIN_DOMAIN;
2:result:= ACCT_CREATE;
3:result:=WIN9X_UPGRADE;
4:result:=DOMAIN_JOIN_IF_JOINED;
5:result:=JOIN_UNSECURE;
6:result:=0;
else result:=0;
end;
end;

function JoinDomainOrWorkGroup(paramUnJwG:pointer):integer;
var
 MyError:integer;
 FWbemObjectSet: OLEVariant;
 oEnum: IEnumvariant;
 FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObject   : OLEVariant;
  iValue        : LongWord;
Const JOIN_DOMAIN = 1 ; // ���������� � ������
Const ACCT_CREATE = 2 ; // ������� ������
Const ACCT_DELETE = 4 ; // ������ �� ������
Const WIN9X_UPGRADE = 16 ; // ��������
Const DOMAIN_JOIN_IF_JOINED = 32 ;  //��������� �������������� � ������ ������, ���� ���� ��������� ��� ����������� � ������.
Const JOIN_UNSECURE = 64  ;        //��������� ������������ ����������.
Const MACHINE_PASSWORD_PASSED = 128 ;
Const DEFERRED_SPN_SET = 256;
Const INSTALL_INVOCATION = 262144;
begin
try
  PointForFolder:=paramUnJwG;
  OleInitialize(nil);
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(PointForFolder.MyPS, 'root\CIMV2', PointForFolder.MyUser, PointForFolder.MyPasswd,'','',128);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 ��� WbemAuthenticationLevelPktPrivacy
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT name FROM Win32_ComputerSystem','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  if oEnum.Next(1, FWbemObject, iValue) = 0 then
  begin
     MyError:=FWbemObject.JoinDomainOrWorkgroup(PointForFolder.NameDomain,PointForFolder.PasswordDomain,PointForFolder.UserDomain,null,
     JOIN_DOMAIN+ACCT_CREATE);
     if PointForFolder.RebootAfterUnjoinDomain then frmDomainInfo.rebuutSelectPC
      else frmDomainInfo.memo1.Lines.Add('���������� ������������� ���������');
     frmDomainInfo.memo1.Lines.Add('�������� - '+inttostr(PointForFolder.Oper)+' - �������� � �����. ('+inttostr(MyError)+') '+SysErrorMessage(MyError));
     frmDomainInfo.memo1.Lines.Add('---------------------------');
  end;
  VariantClear(FWbemObject);
  oEnum:=nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('������  "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;
//NetApiBufferFree(paramUnJwG); /// ������� ������
EndThread(0);    // ������� �����
end;

procedure TOKRightDlg123456.CancelBtnClick(Sender: TObject);
begin
close;
end;

procedure TOKRightDlg123456.ComboBox1Select(Sender: TObject);
begin
NumOper.Text:=inttostr(summOperation(ComboBox1.ItemIndex));
end;

procedure TOKRightDlg123456.OKBtnClick(Sender: TObject);
var
oper:integer;
Res:integer;
TreadIDU:Longword;
NewJoinDomainOrWorkGroup:TThread;
begin
Oper:= summOperation(ComboBox1.ItemIndex);
if NumOper.Text<>'' then Oper:=strtoint(NumOper.Text);
{NumOper.Text:=inttostr(Oper);
RunUnjPC.MyPS:=frmdomaininfo.combobox2.text;  // bvz ����������
RunUnjPC.MyUser:=frmdomaininfo.labelEdEdit1.text; //������������
RunUnjPC.MyPasswd:=frmdomaininfo.labelEdEdit2.text; // ������
RunUnjPC.NameDomain:=LabeledEdit2.Text;              // ��� ������
RunUnjPC.PasswordDomain:= LabeledEdit3.Text;            // ������ ������
RunUnjPC.UserDomain:=LabeledEdit2.Text;           //������������ ������
RunUnjPC.Oper:=Oper;                        // ��������
RunUnjPC.RebootAfterUnjoinDomain:=RebootAfter.Checked;         // ������������ ���� �����
res:=BeginThread(nil,0,addr(JoinDomainOrWorkGroup),Addr(RunUnjPC),0,treadIDU); ///
CloseHandle(res); }
NameDomain:=LabeledEdit1.Text;  // ����� �����
UserDomain:=LabeledEdit2.Text;  // ������������ ������
PasswordDomain:=LabeledEdit3.Text;  // ������ ������������ ������
RebootAfterUnjoinDomain:=RebootAfter.Checked;  // ������������� ��� ��� ���� ����� ��������
JDOWPC:= frmdomaininfo.combobox2.text;  // ����������
JDOWUSER:=frmdomaininfo.labelEdEdit1.text; // ������������
JDOWPASS:=frmdomaininfo.labelEdEdit2.text; // ������
JDOWOper:= Oper;
NewJoinDomainOrWorkGroup:=PCJoinDomainOrWorkGroup.JoinDomainOrWorkGroup.Create(true);
NewJoinDomainOrWorkGroup.FreeOnTerminate:=true;
NewJoinDomainOrWorkGroup.Start;
end;

end.
