unit URDP;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.OleCtrls, MSTSCLib_TLB, Vcl.ComCtrls;

type
  TFormRDP = class(TForm)
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    Label1: TLabel;
    Memo1: TMemo;
    Splitter1: TSplitter;
    ComboKomp: TComboBoxEx;
    PageControl1: TPageControl;
    SpeedButton4: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure NewRdpAuthenticationWarningDismissed(Sender: TObject);
    procedure NewRdpAuthenticationWarningDisplayed(Sender: TObject);
    procedure NewRdpAutoReconnected(Sender: TObject);
    procedure NewRdpAutoReconnecting(ASender: TObject; disconnectReason,
      attemptCount: Integer);
    procedure NewRdpAutoReconnecting2(ASender: TObject;
      disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
      maxAttemptCount: Integer);
    procedure NewRdpChannelReceivedData(ASender: TObject; const chanName,
      data: WideString);
    procedure NewRdpConfirmClose(Sender: TObject);
    procedure NewRdpConnected(Sender: TObject);
    procedure NewRdpConnecting(Sender: TObject);
    procedure NewRdpDisconnected(ASender: TObject; discReason: Integer);
    procedure NewRdpExit(Sender: TObject);
    procedure NewRdpFatalError(ASender: TObject; errorCode: Integer);
    procedure NewRdpLoginComplete(Sender: TObject);
    procedure NewRdpLogonError(ASender: TObject; lError: Integer);
    procedure NewRdpNetworkStatusChanged(ASender: TObject;
      qualityLevel: Cardinal; bandwidth, rtt: Integer);
    procedure NewRdpWarning(ASender: TObject; warningCode: Integer);
    procedure FormShow(Sender: TObject);
    procedure ComboKompDropDown(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
  private
    function createRDP(NamePC,Domain,UserName,Passwd:string):bool;
  public
    { Public declarations }
  end;

var
  FormRDP: TFormRDP;
  MyRDPWin: array [0..100] of  TMsRdpClient9;
  Ntab:     array [0..100] of TTabSheet;
 // Ntab: TTabSheet;
implementation
uses umain;
{$R *.dfm}


function TFormRDP.createRDP(NamePC,Domain,UserName,Passwd:string):bool;
var
i:integer;
begin
try
 i:=0;
 i:=PageControl1.PageCount;
 Ntab[i]:=TTabSheet.Create(PageControl1);
 with Ntab[i] do
  begin
    PageControl := PageControl1;
    Name:='Ntab'+inttostr(i);
    Caption := ComboKomp.Text;
    tag:=i;
  end;

// if PageControl1.FindComponent('MyRDPWin'+inttostr(i))=nil then
 //begin
 MyRDPWin[i]:=TMsRdpClient9.Create(Ntab[i]);
 MyRDPWin[i].Parent:=Ntab[i];
 MyRDPWin[i].Name:='MyRDPWin'+inttostr(i);
 MyRDPWin[i].Align:=alClient;
 MyRDPWin[i].OnAuthenticationWarningDismissed:=NewRdpAuthenticationWarningDismissed;
 MyRDPWin[i].OnAuthenticationWarningDisplayed:=NewRdpAuthenticationWarningDisplayed;
 MyRDPWin[i].OnAutoReconnected:=NewRdpAutoReconnected;
 MyRDPWin[i].OnAutoReconnecting:=NewRdpAutoReconnecting;
 MyRDPWin[i].OnAutoReconnecting2:=NewRdpAutoReconnecting2;
 MyRDPWin[i].OnChannelReceivedData:=NewRdpChannelReceivedData;
 MyRDPWin[i].OnConfirmClose:=NewRdpConfirmClose;
 MyRDPWin[i].OnConnected:=NewRdpConnected;
 MyRDPWin[i].OnConnecting:=NewRdpConnecting;
 MyRDPWin[i].OnDisconnected:=NewRdpDisconnected;
 MyRDPWin[i].OnExit:=NewRdpExit;
 MyRDPWin[i].OnFatalError:=NewRdpFatalError;
 MyRDPWin[i].OnLoginComplete:=NewRdpLoginComplete;
 MyRDPWin[i].OnLogonError:=NewRdpLogonError;
 MyRDPWin[i].OnNetworkStatusChanged:=NewRdpNetworkStatusChanged;
 MyRDPWin[i].OnWarning:=NewRdpWarning;
 //end;

 MyRDPWin[i].Server:=NamePC;
 MyRDPWin[i].Domain:=Domain;
 if LabeledEdit1.Text<>'' then MyRDPWin[i].UserName:=UserName;
 if LabeledEdit2.Text<>'' then MyRDPWin[i].AdvancedSettings9.ClearTextPassword:=Passwd;
 MyRDPWin[i].DesktopWidth:=Ntab[i].Width;
 MyRDPWin[i].DesktopHeight:=Ntab[i].Height;
 MyRDPWin[i].AdvancedSettings9.RedirectDrives:=true;     ///��������� �������������� �����
 MyRDPWin[i].AdvancedSettings9.RedirectPrinters:=true;   /// ��������� �������������� ��������
 MyRDPWin[i].AdvancedSettings9.RelativeMouseMode:=true;  ///
 MyRDPWin[i].AdvancedSettings9.RedirectClipboard:=true;  ///
 MyRDPWin[i].AdvancedSettings9.RedirectDevices:=true;    /// ��������������� ������ ��������� PnP
 MyRDPWin[i].AdvancedSettings9.RedirectPorts:=true;      /// ��������������� ������
 MyRDPWin[i].AdvancedSettings9.EnableAutoReconnect:=true;  //// ��������� ��������������� ��� ��������� ����
 MyRDPWin[i].AdvancedSettings9.ConnectToAdministerServer:=true; /// ��������� ������ �� ������� ���������� ActiveX �������� ������������ � ������� � ���������������� �����.
 MyRDPWin[i].AdvancedSettings9.AudioCaptureRedirectionMode:=true;  ///  ���������, ���������������� �� ���������� ���������� �� ��������� �� ������� � ���������� ������
 MyRDPWin[i].AdvancedSettings9.EnableSuperPan:=true;              ///��������� ������������ ����� ������������ �� ���������� �������� ����� � ������������� ������, ����� ������� ���������� �������� ����� ������ �������� �������� ���� �������. ������ ����, ����� ������������ ������ ��������� ��� ��������� �� �������� �����, ������������ ����� ������� �� ������� ����, � ��������� ������� ���� ����� ������������� �������������� � ���� �����������. SuperPan �� ������������ ����� ������ ��������.
 MyRDPWin[i].AdvancedSettings9.BandwidthDetection:=true;          ///����������, ����� �� ������������� ���������� ��������� ���������� �����������.
 MyRDPWin[i].AdvancedSettings9.AuthenticationLevel:=2;            ///������ ������� �������� ����������� ��� �����������.  (0.1.2)
 MyRDPWin[i].AdvancedSettings9.EnableCredSspSupport:=true;      ///��������� ��� ���������, ������� �� ��� ����� �����������  ��������� ������ ������������ ������� ������ (CredSSP)
 MyRDPWin[i].SecuredSettings3.KeyboardHookMode:=1;
 PageControl1.ActivePageIndex:=i;
 MyRDPWin[i].Connect;
 MyRDPWin[i].Tag:=i;
 result:=true;
 except
 on E:Exception do
  begin
   result:=false;
   Memo1.Lines.Add('������ ��� ����������� � ���������� '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.SpeedButton1Click(Sender: TObject); //// ����������� �� RDP
begin
try
if (ComboKomp.Text='')  then
begin
ShowMessage('�� ������� ��� ����������');
ComboKomp.SetFocus; /// ��������� ����� �� ���������
exit;
end;


if PageControl1.TabIndex<>-1 then  /// ���� �� ��������� ��������
begin
if PageControl1.ActivePage.FindComponent('MyRDPWin'+inttostr(PageControl1.TabIndex))<>nil then
 begin
 MyRDPWin[PageControl1.TabIndex].Connect;
 end;
end
else //// ���� ������� ��� �� ������� ������
 begin
 if not createRDP(ComboKomp.Text,CurrentDomainName,labelededit1.Text,labelededit2.Text) then
 ShowMessage('������ �������� ����������� ��� '+ComboKomp.Text);
 end;


except
 on E:Exception do
  begin
   Memo1.Lines.Add('������ ��� ����������� � ���������� '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;

end;

procedure TFormRDP.SpeedButton2Click(Sender: TObject); //// ���������� RDp �� �������
begin
try

if PageControl1.TabIndex<>-1 then  /// ���� �� ��������� ��������
begin
if PageControl1.ActivePage.FindComponent('MyRDPWin'+inttostr(PageControl1.TabIndex))<>nil then
 begin
 //if MyRDPWin[PageControl1.TabIndex].Tag=PageControl1.TabIndex then
 MyRDPWin[PageControl1.TabIndex].Disconnect;
 end;
end;

except
 on E:Exception do
  begin
   Memo1.Lines.Add('������ ��� ���������� �� ���������� '+PageControl1.ActivePage.Caption+' - '+e.Message)
  end;
end;
end;

procedure TFormRDP.SpeedButton3Click(Sender: TObject); /// �������� �������
var
i:integer;
begin
if ComboKomp.Text='' then
  begin
  ShowMessage('�� ������� ��� ����������');
  ComboKomp.SetFocus; /// ��������� ����� �� ���������
  exit;
  end;
for I := 0 to PageControl1.PageCount-1 do
begin
  if ComboKomp.Text=PageControl1.Pages[i].Caption then
  begin
  PageControl1.ActivePageIndex:=i;
  ShowMessage('����������� ��� ����������');
  exit;
  end;
end;



if PageControl1.PageCount<100 then /// ��� ��� ��������� � ������� ������ 100 �� ������ 100 ������� �� ���������
 begin
 if not createRDP(ComboKomp.Text,CurrentDomainName,labelededit1.Text,labelededit2.Text) then
 ShowMessage('������ �������� ����������� ��� '+ComboKomp.Text);
 end;
end;

procedure TFormRDP.SpeedButton4Click(Sender: TObject); /// ������� �������
begin
if PageControl1.ActivePageIndex<>-1 then
 begin
  if PageControl1.FindComponent('Ntab'+inttostr(PageControl1.TabIndex))<>nil then
   begin
    if PageControl1.ActivePage.FindComponent('MyRDPWin'+inttostr(PageControl1.TabIndex))<>nil then
     freeandnil(MyRDPWin[PageControl1.TabIndex]);
     freeandnil(Ntab[PageControl1.TabIndex]);
   end;
 end;
end;


procedure TFormRDP.ComboKompDropDown(Sender: TObject);
var   //// ���������� ������� ���� ����������� ��� ��������� ������
z:integer;
begin
try
for z := 0 to comboKomp.Items.Count-1 do
if comboKomp.Items[z]=frmDomainInfo.ComboBox2.Items[z] then
comboKomp.ItemsEx[z].ImageIndex:=frmDomainInfo.ComboBox2.ItemsEx[z].ImageIndex;
Except
   on E: Exception do
      begin
        memo1.Lines.add('������ ���������� ������� �����������:'+E.Message);
      end;
   end;
end;

procedure TFormRDP.FormShow(Sender: TObject);
var
i:integer;
begin
ComboKomp.Clear;
for I := 0 to frmDomainInfo.ComboBox2.Items.Count-1 do
begin
ComboKomp.Items.Add(frmDomainInfo.ComboBox2.Items[i]);
ComboKomp.ItemsEx[i].ImageIndex:=frmDomainInfo.ComboBox2.ItemsEx[i].ImageIndex;
end;

LabeledEdit1.Text:=frmDomainInfo.LabeledEdit1.Text;
LabeledEdit2.Text:=frmDomainInfo.LabeledEdit2.Text;
end;

procedure TFormRDP.NewRdpAuthenticationWarningDismissed(Sender: TObject);
begin
Memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - �������������� ��������������');
end;

procedure TFormRDP.NewRdpAuthenticationWarningDisplayed(Sender: TObject);
 const
  AutType: array [0..4] of string =('No authentication is used.',
  '������������ �������� ����������� �����������','������������ �������� ����������� Kerberos',
  '������������ ��� ���� �������� �����������, ����������� � �� ���������','����������� ���������');
begin


if PageControl1.TabIndex<>-1 then  /// ���� ������ �� ��������� ��������
begin
if PageControl1.ActivePage.FindComponent('MyRDPWin'+inttostr(PageControl1.TabIndex))<>nil then
 begin
 if (MyRDPWin[PageControl1.TabIndex].AdvancedSettings9.AuthenticationType<>0) and
 (MyRDPWin[PageControl1.TabIndex].AdvancedSettings9.AuthenticationType<5) then
 memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - �������� �����������. - '+ AutType[MyRDPWin[PageControl1.TabIndex].AdvancedSettings9.AuthenticationType]);
 end;
end;



end;

procedure TFormRDP.NewRdpAutoReconnected(Sender: TObject);
begin
Memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ��������������� � ���������� ����������.- ');
end;

procedure TFormRDP.NewRdpAutoReconnecting(ASender: TObject; disconnectReason,
  attemptCount: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - �������������� ��������������� ������ ��������� ������� ������.');
end;

procedure TFormRDP.NewRdpAutoReconnecting2(ASender: TObject;
  disconnectReason: Integer; networkAvailable: WordBool; attemptCount,
  maxAttemptCount: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ���������� ������� ���������� ���������������� � ���������� ������');
end;

procedure TFormRDP.NewRdpChannelReceivedData(ASender: TObject; const chanName,
  data: WideString);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ������ �������� ������ �� ����������� ������ � ������������ ��������');
end;

procedure TFormRDP.NewRdpConfirmClose(Sender: TObject);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ���������� ������ ����� ���������� �������� �����');
end;

procedure TFormRDP.NewRdpConnected(Sender: TObject);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ��������� ����������...');
end;

procedure TFormRDP.NewRdpConnecting(Sender: TObject);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ������������� ����������...');
end;

procedure TFormRDP.NewRdpDisconnected(ASender: TObject; discReason: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ���������� �� ������� ����� ������� ��������� ������� ������');
end;

procedure TFormRDP.NewRdpExit(Sender: TObject);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - �����');
end;

procedure TFormRDP.NewRdpFatalError(ASender: TObject; errorCode: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ���������� ������������ ������');
end;

procedure TFormRDP.NewRdpLoginComplete(Sender: TObject);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - �������� ���� � �������');
end;

procedure TFormRDP.NewRdpLogonError(ASender: TObject; lError: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ������ ����� � ������� - '+ SysErrorMessage(lError));
end;

procedure TFormRDP.NewRdpNetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption+' : RDP - ��������� ��������� ��������� ����');
end;

procedure TFormRDP.NewRdpWarning(ASender: TObject; warningCode: Integer);
begin
memo1.Lines.Add(PageControl1.ActivePage.Caption
+' : RDP - OnWarning ���������� �� ��������� ������ - '
+ SysErrorMessage(warningCode));
end;





end.
