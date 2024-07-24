unit OfficeActivation;

interface

uses Winapi.Windows, System.SysUtils, System.Classes, Vcl.Graphics, Vcl.Forms,
  Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.ComCtrls,
  ActiveX,System.Variants,ComObj, Vcl.Dialogs, Vcl.ExtDlgs,IdIcmpClient;

type
  TActivateOffice = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    ListView1: TListView;
    Button2: TButton;
    Button3: TButton;
    button4: TButton;
    Label2: TLabel;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListView1CustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure ListView1DblClick(Sender: TObject);
  private
  Function ViewLicensing:string;
  Function Ping(s:string):boolean;
  Function ActivateProduct:string;
  function creatDetalForm(s:string):boolean;
  procedure FormClose(Sender: TObject; var Action: TCloseAction);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ActivateOffice: TActivateOffice;

const
 LicStatus: array [0..6] of string =('Нет лицензии','Лицензия активна',
  'Льготный период','Допустимое отклонение от льготного периода','Льготный период для неподлинных версий','Предупреждение','Расширенное льготное время');


implementation
uses umain,MyDM;
var

MyIdIcmpClient: TIdIcmpClient;

{$R *.dfm}
 function GetObject(const objectName: String): IDispatch;
var
  chEaten: Integer;
  BindCtx: IBindCtx;//for access to a bind context
  Moniker: IMoniker;//Enables you to use a moniker object
begin
  OleCheck(CreateBindCtx(0, bindCtx));
  OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));//Converts a string into a moniker that identifies the object named by the string
  OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));//Binds to the specified object
end;

function DescriptionLicStatus(num:string):string;
var   /// передаем код ошибки и из таблицы получаем описание этой ошибки
z:integer;
he,s,A:string;
begin
try
a:=num;
if pos('OLE error ',num)<>0 then
begin
s:=copy(num,11,length(num));
a:=s;
end;
//https://support.microsoft.com/ru-ru/windows/%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BA%D0%B0-%D0%BF%D0%BE-%D0%BE%D1%88%D0%B8%D0%B1%D0%BA%D0%B0%D0%BC-%D0%B0%D0%BA%D1%82%D0%B8%D0%B2%D0%B0%D1%86%D0%B8%D0%B8-windows-09d8fb64-6768-4815-0c30-159fa7d89d85
if TryStrToInt(a,z) then
begin
he:=IntToHex(z,1);
end
else he:=a;
Datam.FDSelectReadSoft.Transaction.StartTransaction;
try
Datam.FDSelectReadSoft.SQL.clear;
Datam.FDselectReadSoft.SQL.Text:='SELECT * FROM LIC_ERROR WHERE CODE='''+he+'''';
Datam.FDselectReadSoft.Open;
result:=Datam.FDselectReadSoft.FieldByName('DESCRIPTION').AsString;
if result='' then result:='Unknown';
finally
Datam.FDSelectReadSoft.Transaction.Commit;
end;
except on E: Exception do
begin
 result:='Unknown';
end;
end;
end;
{function DescriptionLicStatus(num:string):string;
var
z:integer;
he,s,A:string;
begin
try  // "OLE error C004F050"
a:=num;
if pos('OLE error ',num)<>0 then
begin
s:=copy(num,11,length(num));
a:=s;
end;
//https://support.microsoft.com/ru-ru/windows/%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BA%D0%B0-%D0%BF%D0%BE-%D0%BE%D1%88%D0%B8%D0%B1%D0%BA%D0%B0%D0%BC-%D0%B0%D0%BA%D1%82%D0%B8%D0%B2%D0%B0%D1%86%D0%B8%D0%B8-windows-09d8fb64-6768-4815-0c30-159fa7d89d85
if TryStrToInt(a,z) then
begin
he:=IntToHex(z,1);
end
else he:=a;
if he='0' then result:= 'Продукт на этом устройстве активирован с помощью ключа.';
if he='4004F401' then result:= 'Сервер активации сообщил, что на компьютере есть действующая Цифровая лицензия. ('+he+')';
if he='4004F040' then result:= 'Служба лицензирования программного обеспечения сообщила, что продукт был активирован, но владелец должен подтвердить права на использование продукта.('+he+')';
if he='C004C001' then result:= 'Сервер активации определил, что указанный ключ продукта недопустим ('+he+')';
if he='C004C003' then result:= 'Сервер активации определил, что указанный ключ продукта заблокирован ('+he+')';
if he='C004C017' then result:= 'Сервер активации определил, что указанный ключ продукта был заблокирован для этого географического расположения. ('+he+')';
if he='C004B100' then result:= 'Сервер активации обнаружил, что для данного компьютера не удалось выполнить активацию ('+he+')';
if he='C004D307' then result:= 'Превышено максимально допустимое число сбросов таймера активации. Чтобы повторить попытку сброса таймера активации, переустановите ОС. ('+he+')';
if he='C004F00F' then result:= 'Привязка идентификатора оборудования выходит за рамки допустимого. Замена M /Board или другие аппаратные изменения('+he+')';
if he='C004F014' then result:= 'Служба лицензирования программного обеспечения сообщила, что ключ продукта недоступен ('+he+')';
if he='C004C001' then result:= 'Сервер активации определил, что указан недопустимый ключ продукта. ('+he+')';
if he='C004C003' then result:= 'Сервер активации определил, что указанный ключ продукта заблокирован. ('+he+')';
if he='C004B100' then result:= 'Сервер активации определил, что этот компьютер не может быть активирован. ('+he+')';
if he='C004C008' then result:= 'Сервер активации определил, что указанный ключ продукта невозможно использовать. ('+he+')';
if he='C004C020' then result:= 'Сервер активации вернул сообщение о превышении предела для многопользовательского ключа активации (MAK). ('+he+')';
if he='C004C021' then result:= 'Сервер активации вернул сообщение о превышении предела расширения для многопользовательского ключа активации. ('+he+')';
if he='C004F009' then result:= 'Служба лицензирования программного обеспечения вернула сообщение об истечении льготного периода. ('+he+')';
if he='C004F00F' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о выходе идентификатора привязки оборудования за границы допустимого отклонения. ('+he+')';
if he='C004F014' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о недоступности ключа продукта. ('+he+')';
if he='C004F02C' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о неправильности формата данных автономной активации. ('+he+')';
if he='C004F035' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о невозможности активации компьютера с помощью ключа многократной установки. Используйте ключ другого типа. ('+he+')';
if he='C004F038' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о невозможности активации компьютера. Счетчик, содержащийся в отчете службы управления ключами (KMS), имеет недостаточное значение. Обратитесь к своему системному администратору. ('+he+')';
if he='C004F039' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о невозможности активации компьютера. Служба управления ключами (KMS) отключена. ('+he+')';
if he='C004F041' then result:= 'Служба лицензирования программного обеспечения обнаружила, что служба управления ключами (KMS) не активирована. Необходимо активировать службу KMS. ('+he+')';
if he='C004F042' then result:= 'Служба лицензирования программного обеспечения обнаружила, что указанная служба управления ключами (KMS) не может быть использована.  ('+he+')';
if he='C004F050' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о недопустимости этого ключа продукта. ('+he+')';
if he='C004F051' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о блокировке этого ключа продукта. ('+he+')';
if he='C004F064' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о завершении льготного периода для неподлинной версии. ('+he+')';
if he='C004F065' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о работе приложения в рамках допустимого льготного периода для неподлинной версии. ('+he+')';
if he='C004F066' then result:= 'Служба лицензирования программного обеспечения вернула сообщение, что единицы хранения продукта не найдены. ('+he+')';
if he='C004F069' then result:= 'Служба лицензирования программного обеспечения вернула сообщение о невозможности активации компьютера. Служба лицензирования программного обеспечения (KMS) определила, что отметка времени для запроса является недопустимой. ('+he+')';
if he='C004F06B' then result:= 'Служба лицензирования программного обеспечения определила, что она выполняется на виртуальном компьютере. Служба управления ключами (KMS) в этом режиме не поддерживается ('+he+')';
if he='C004F074' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Служба управления ключами (KMS) недоступна ('+he+')';
if he='C004F075' then result:= 'Служба лицензирования программного обеспечения сообщила, что операцию невозможно выполнить, так как служба останавливается ('+he+')';
if he='C004F304' then result:= 'Служба лицензирования программного обеспечения сообщила, что требуемая лицензия не найдена. ('+he+')';
if he='C004F305' then result:= 'Служба лицензирования программного обеспечения сообщила о том, что в системе не найдены сертификаты, позволяющие активировать продукт. ('+he+')';
if he='C004F30A' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Сертификат не соответствует условиям лицензии. ('+he+')';
if he='C004F30D' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Недействительный отпечаток. ('+he+')';
if he='C004F30E' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Не удается найти сертификат, соответствующий отпечатку. ('+he+')';
if he='C004F30F' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Сертификат не соответствует условиям, указанным в лицензии выдачи. ('+he+')';
if he='C004F310' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Сертификат не соответствует идентификатору точки доверия (TPID), указанному в лицензии выдачи. ('+he+')';
if he='C004F311' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Программный токен не может быть использован для активации. ('+he+')';
if he='C004F312' then result:= 'Служба лицензирования программного обеспечения сообщила, что для данного компьютера не удалось выполнить активацию. Сертификат не может быть использован, так как его закрытый ключ является экспортируемым. ('+he+')';
if he='80070005' then result:= 'Нет доступа. Для выполнения запрашиваемого действия требуются повышенные права. ('+he+')';
if he='8007232A' then result:= 'Ошибка DNS-сервера. ('+he+')';
if he='8007232B' then result:= 'DNS-имя не существует. ('+he+')';
if he='800706BA' then result:= 'Сервер RPC недоступен. ('+he+')';
if he='8007251D' then result:= 'Записи не найдены для данного запроса DNS. ('+he+')';
if he='C004F056' then  result:='Служба лицензирования программного обеспечения сообщила, что продукт не может быть активирован с помощью службы управления ключами. ('+he+')';
if he='C004F034' then  result:='Указан недопустимый ключ продукта или ключ продукта для другой версии Windows. ('+he+')';
if he='C004F057' then  result:='Не удалось установить доказательство покупки из таблицы ACPI. ('+he+')';
if he='4004FC05' then  result:='Служба лицензирования программного обеспечения сообщила, что приложение имеет бессрочный льготный период.. ('+he+')';
if he='4004F00D' then result:='Ошибка активации продукта. ('+he+')';
if result='' then  result:=he;
except on E: Exception do  result:=num;
end;
end; }

{LicenseStatus	DWORD	4	Состояние лицензирования
0 - Нелицензировано
1 - Лицензировано (активировано)
2 – Льготный период OOB
3 – Льготный период OOT
4 - NonGenuineGrace
}

function refreshOfficeLic:boolean;     /// обновление статуса лицензии офис
 var
  FSWbemLocator                           : OLEVariant;
  FWMIService                             : OLEVariant;
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  E                                       : integer;
  iValue        : LongWord;
Begin
try
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT version FROM OfficeSoftwareProtectionService','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
         while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
            begin
            FWbemObject.RefreshLicenseStatus();
            end;
    except
    on E:Exception do
     begin
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet);
     VariantClear(FWMIService);
     VariantClear(FSWbemLocator);
     end;
    end;
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
end;

function TActivateOffice.ActivateProduct:string;
var
FWMIService         : OLEVariant;
ObjectSet,FWbemObject      : OLEVariant;
oEnum                : IEnumvariant;
ErrorActivate: integer;
FSWbemLocator : OLEVariant;
iValue        : LongWord;
begin
try
 listview1.Clear;
 memo1.Lines.Add('Поиск продукта, для активации...');
//FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
 FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
 FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
 FWMIService.Security_.impersonationlevel:=3;
 FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
 ObjectSet:= FWMIService.ExecQuery('SELECT ProductKeyID,ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description,LicenseStatusReason, Name, LicenseStatus FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey <> null');
 oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
     FWbemObject.Activate();
     refreshOfficeLic;
     with listview1.Items.Add do
                 begin
                 if FWbemObject.Name<>null then Caption:= vartostr(FWbemObject.Name)
                  else Caption:='';
                 if FWbemObject.Description<>null then SubItems.add( vartostr(FWbemObject.Description))
                  else SubItems.add('');
                 if FWbemObject.LicenseStatus<>null then SubItems.add(LicStatus[(integer(FWbemObject.LicenseStatus))])
                  else SubItems.add('');
                 if FWbemObject.PartialProductKey<>null then SubItems.add(vartostr(FWbemObject.PartialProductKey))
                  else SubItems.add('');
                 if FWbemObject.ProductKeyID<>null then SubItems.add(vartostr(FWbemObject.ProductKeyID))
                  else SubItems.add('');
                 if FWbemObject.ID<>null then SubItems.add(vartostr(FWbemObject.ID))
                  else SubItems.add('');
                 if FWbemObject.LicenseStatusReason<>null then SubItems.add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason)))
                  else SubItems.add('');
                 end;
     memo1.Lines.Add('Продукт прошел активацию!');
     end;
    VariantClear(FWMIService);
    VariantClear(ObjectSet);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;
 except
    on E:EOleException do
      begin
      result:= 'Ошибка Активации - "'+DescriptionLicStatus(E.Message)+'"';
      VariantClear(FWMIService);
      VariantClear(ObjectSet);
      VariantClear(FWbemObject);
      VariantClear(FSWbemLocator);
      oEnum:=nil;
      end;
    end;
end;


function TActivateOffice.ping(s:string):boolean;
var
z:integer;
begin
try
//////////////////////////////////////
Myidicmpclient:=TIdIcmpClient.Create;
Myidicmpclient.IPVersion:=frmDomainInfo.IdIcmpClient1.IPVersion;
Myidicmpclient.PacketSize:=32;
Myidicmpclient.Port:=0;
Myidicmpclient.Protocol:=1;
Myidicmpclient.ReceiveTimeout:=3000;
////////////////////////////////////////
result:=false;
Myidicmpclient.host:=s;
Myidicmpclient.Port:= 0;
MyIdIcmpClient.Ping;
if (MyIdIcmpClient.ReplyStatus.Msg<>'Echo') or
  (MyIdIcmpClient.ReplyStatus.FromIpAddress='0.0.0.0') then
      begin
        memo1.Lines.Add(s+' - Превышен интервал ожидания запроса, узел не доступен');
        result:=false;
      end
    else
      begin
      Memo1.Lines.Add('IP хоста - '+MyIdIcmpClient.ReplyStatus.FromIpAddress+'  Время='+inttostr(MyIdIcmpClient.ReplyStatus.MsRoundTripTime)+'мс'+'  TTL='+inttostr(MyIdIcmpClient.ReplyStatus.TimeToLive));
      result:=true;
      end;
  except
  on E: Exception do
    begin
      memo1.Lines.add(s+' - Узел не доступен.');
      result:=false;
    end;
  end;
  if Assigned(Myidicmpclient) then freeandnil(Myidicmpclient);
end;
//////////////////////////////////////////////////////////////////
procedure TActivateOffice.Button2Click(Sender: TObject);
var
  FWMIService         : OLEVariant;
  ObjectSet,FWbemObject: OLEVariant;
  oEnum                : IEnumvariant;
  StringKey            :string;
  FSWbemLocator : OLEVariant;
  iValue        : LongWord;
Begin
try
   if not InputQuery('Ключ продукта', 'Key-', StringKey)
    then exit;
  if (StringKey='') then
  begin
   showmessage('Вы не ввели ключ продукта');
   exit;
  end;
   if ping(MyPS)=false then   exit;
  memo1.Lines.Add('Установка ключа продукта...');
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWMIService.Security_.impersonationlevel:=3;
  FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
  ObjectSet:= FWMIService.ExecQuery('SELECT Version FROM OfficeSoftwareProtectionService');
  oEnum:= IUnknown(ObjectSet._NewEnum) as IEnumVariant;
   while oEnum.Next(1, FWbemObject, iValue) = 0 do
     begin
      FWbemObject.InstallProductKey(StringKey); //// установка ключа windows
      //memo1.Lines.Add('Начало активации продукта');
     // memo1.Lines.Add(ActivateProduct); ///// функция активация продукта
     // memo1.Lines.Add('Обновление статуса лицензии');
     // FWbemObject.RefreshLicenseStatus(); /// Обновления статуса лицензирования продукта
      memo1.Lines.Add('Установка ключа продукта завершена');
      memo1.Lines.Add(ViewLicensing);
     end;
 except
   on E:EOleException do
   begin
    ShowMessage('Общая ошибка установки ключа  "'+DescriptionLicStatus(E.Message)+'"');
    memo1.Lines.Add('Общая ошибка установки ключа  "'+DescriptionLicStatus(E.Message)+'"');
    VariantClear(ObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;
   end;
end;
    VariantClear(ObjectSet);
    VariantClear(FWMIService);
    VariantClear(FWbemObject);
    VariantClear(FSWbemLocator);
    oEnum:=nil;

End;
///////////////////////////////////////////////////////////////////////////////////////
procedure TActivateOffice.Button3Click(Sender: TObject);
begin
memo1.Lines.Add(ActivateProduct);
end;

procedure TActivateOffice.button4Click(Sender: TObject);
 var
  FSWbemLocator                           : OLEVariant;
  FWMIService                             : OLEVariant;
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  E                                       : integer;
  iValue        : LongWord;
Begin
if (listview1.Items.Count=0) or (listview1.SelCount=0) then
begin
showmessage('Не выбран продукт');
exit;
end;

e:=MessageDlg('Вы действительно хотите удалить ключ продукта?',mtConfirmation,[mbYes,mbCancel], 0);
if e=mrCancel then exit;
try
  if ping(MyPS)=false then exit;
  memo1.Lines.Add('Процесс удаления ключа продукта...');
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM OfficeSoftwareProtectionProduct WHERE ProductKeyID='+'"'+listview1.Items[listview1.ItemIndex].SubItems[3]+'"','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
  begin
  FWbemObject.UninstallProductKey();  /// удаление ключа продукта
  memo1.Lines.Add('Ключ удален');
  end;
    except
    on E:Exception do
     begin
     ShowMessage('Ошибка удаления ключа продукта  "'+E.Message+'"');
     memo1.Lines.Add('Ошибка удаления ключа продукта  "'+E.Message+'"');
     VariantClear(FWbemObject);
     oEnum:=nil;
     VariantClear(FWbemObjectSet);
     VariantClear(FWMIService);
     VariantClear(FSWbemLocator);
     end;
    end;
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
end;
procedure TActivateOffice.FormShow(Sender: TObject);
begin
listview1.Clear;
memo1.Clear;
if myps='' then myps:='localhost';

end;

procedure TActivateOffice.ListView1CustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
if  item.subitems[3]<>'' then  ListView1.Canvas.Font.Color:=clBlue;
end;

procedure TActivateOffice.FormClose(Sender: TObject; var Action: TCloseAction);
begin
(sender as TForm).FreeOnRelease;
end;

function TActivateOffice.creatDetalForm(s:string):boolean;
var
DetalForm:Tform;
DetalMemo:Tmemo;
begin
DetalForm:=TForm.Create(ActivateOffice);
DetalForm.Caption:=s;
DetalForm.Height:=250;
DetalForm.Width:=400;
DetalForm.Position:=poOwnerFormCenter;
DetalForm.BorderStyle:=bsDialog;
DetalForm.OnClose:=FormClose;
DetalMemo:=Tmemo.Create(DetalForm);
DetalMemo.parent:=DetalForm;
DetalMemo.Align:=alClient;
DetalMemo.ScrollBars:=ssVertical;
DetalMemo.ScrollBars:=ssHorizontal;
DetalMemo.Clear;
DetalMemo.Lines.Add('Продукт: '+ListView1.Selected.Caption);
DetalMemo.Lines.Add('Описание: '+ListView1.Selected.SubItems[0]);
DetalMemo.Lines.Add('Статус лицензии: '+ListView1.Selected.SubItems[1]);
DetalMemo.Lines.Add('Ключ продукта: '+ListView1.Selected.SubItems[2]);
DetalMemo.Lines.Add('Код: '+ListView1.Selected.SubItems[3]);
DetalMemo.Lines.Add('Идентификатор: '+ListView1.Selected.SubItems[4]);
DetalMemo.Lines.Add('Примечание: '+ListView1.Selected.SubItems[5]);
DetalForm.ShowModal;
end;

procedure TActivateOffice.ListView1DblClick(Sender: TObject);
begin
if ListView1.Selected.Caption<>'' then creatDetalForm(ListView1.Selected.Caption);

end;

///////////////////////////////////////////////////////////////////////////////////////
Function TActivateOffice.ViewLicensing:string;
  var
  FSWbemLocator                           : OLEVariant;
  FWMIService                             : OLEVariant;
  FWbemObjectSet                          : OLEVariant;
  FWbemObject                             : OLEVariant;
  oEnum                                   : IEnumvariant;
  iValue        : LongWord;
Begin
try
  listview1.Clear;
  memo1.Lines.Add('Обновляю список продуктов, ожидайте...');
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT Name,Description,LicenseStatus,PartialProductKey,ProductKeyID,ID,LicenseStatusReason FROM OfficeSoftwareProtectionProduct','WQL',wbemFlagForwardOnly);
  oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do     //// Определяем архитектуру
  begin
   with listview1.Items.Add do
   begin
    if (FWbemObject.Name<>null) then Caption:= vartostr(FWbemObject.Name)
    else Caption:='';
    if FWbemObject.Description<>null then SubItems.add( vartostr(FWbemObject.Description))
    else SubItems.add('');
    if FWbemObject.LicenseStatus<>null then SubItems.add(LicStatus[(integer(FWbemObject.LicenseStatus))])
    else SubItems.add('');
    if FWbemObject.PartialProductKey<>null then SubItems.add(vartostr(FWbemObject.PartialProductKey))
    else SubItems.add('');
    if FWbemObject.ProductKeyID<>null then SubItems.add(vartostr(FWbemObject.ProductKeyID))
    else SubItems.add('');
    if FWbemObject.ID<>null then SubItems.add(vartostr(FWbemObject.ID))
    else SubItems.add('');
    if FWbemObject.LicenseStatusReason<>null then SubItems.add(DescriptionLicStatus(vartostr(FWbemObject.LicenseStatusReason)))
    else SubItems.add('');
   end;
  end;

except
 on E:Exception do
  begin
  // ShowMessage('Общая ошибка  "'+E.Message+'"');
   memo1.Lines.Add('Общая ошибка  "'+E.Message+'"');
   memo1.Lines.Add('ОС не поддерживается, попробуйте через активацию Windows!');
   //Result:='Общая ошибка  "'+E.Message+'"';
  end;
end;
    Result:='Список обновлен';
    VariantClear(FWbemObject);
    oEnum:=nil;
    VariantClear(FWbemObjectSet);
    VariantClear(FWMIService);
    VariantClear(FSWbemLocator);
End;



procedure TActivateOffice.Button1Click(Sender: TObject);
begin
ViewLicensing;
end;

end.
