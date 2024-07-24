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
PrintCapabilities :array [0..21] of string =('Неизвестно','Другое','Цветная печать'
,'Двухсторонняя печать','Копирование','Сопоставление','Сшивание','Прозрачная печать'
,'Перфорация','Печать на обложках','Брошуровщик','Черно-белая печать','Односторонний'
,'Двухсторонний, печать по длинному краю','Двухсторонний, печать по короткому краю'
,'Портрет','Пейзаж','Обратный портрет','Обратный пейзаж','Высокое качество',
'Нормальное качество','Низкое качество');
PrintAvailability :array [1..21] of string =('Другое','Неизвестно','Полная мощность'
,'Предупреждение','В тесте','Не применимо','Выключение','Фффлайн режим','Выключен'
,'Ухудшенный','Не установленно','Ошибка установки','Энергосбережение-неизвестно'
,'Энергосбережение-Режим низкого энергопотребления','Энергосбережение - режим ожидания'
,'Цикл питания','Энергосбережение - Предупреждение','Приостановлено','Не готов'
,'Не настроен','Quiesced (Тихое устройство)');
PrintConfigManagerErrorCode :array [0..31] of string =('Устройство работает нормально'
,'Устройство настроено неправильно','Нет драйверов для этого устройства'
,'Возможно, драйвер для этого устройства поврежден, или в вашей системе недостаточно памяти или других ресурсов'
,'Это устройство не работает должным образом. Один из драйверов или ваш реестр могут быть повреждены'
,'Драйвер для этого устройства нуждается в ресурсе, которым Windows не может управлять'
,'Конфигурация загрузки для устройства конфликтует с другими устройствами'
,'Не удается отфильтровать',
'Загрузчик драйверов для устройства отсутствует'
,'Устройство не работает должным образом. Управляющая прошивка неправильно сообщает о ресурсах для устройства'
,'Устройство не может запуститься','Устройство не удалось','Устройство не может найти достаточно свободных ресурсов для использования'
,'Windows не может проверить ресурсы устройства','Устройство не может работать должным образом, пока компьютер не будет перезагружен'
,'Устройство не работает должным образом из-за возможной проблемы с повторным перечислением'
,'Windows не может определить все ресурсы, которые использует устройство.'
,'Устройство запрашивает неизвестный тип ресурса'
,'Переустановите драйверы для этого устройства','Сбой при использовании загрузчика VxD'
,'Реестр может быть поврежден','Сбой системы: попробуйте изменить драйвер для этого устройства.'
,'Устройство отключено','Сбой системы: попробуйте изменить драйвер для этого устройства'
,'Устройство отсутствует, работает неправильно или не установлены все его драйверы'
,'Windows все еще настраивает устройство','Windows все еще настраивает устройство.'
,'Устройство не имеет допустимой конфигурации журнала','Драйверы устройств не установлены'
,'Устройство отключено. Прошивка устройства не обеспечила требуемые ресурсы'
,'Устройство использует ресурс IRQ, который использует другое устройство'
,'Устройство не работает должным образом. Windows не может загрузить необходимые драйверы устройств.'
);
PrintDetectedErrorState :array [0..11] of string =('Неизвестно','Другое','ОК'
,'Мало бумаги',
'Нет бумаги','Мало тонера','Нет тонера','Открыта дверца','Застряла бумага'
,'Оффлайн  режим','Необходимо обслуживание','Выходной лоток заполнен');
printExtendedDetectedErrorState :array [0..15] of string =('Неизвестная ошибка'
,'Другое','Ошибок нет','Мало бумаги','Нет бумаги','Мало тонера','Нет тонера'
,'Открыта дверца принтера','Замятие','Service Requested','Выходной лоток заполнен'
,'Проблема с бумагой','Не удается распечатать страницу','Требуется вмешательство пользователя'
,'Недостаточно памяти','Сервер неизвестен');
PrintExtendedPrinterStatus :array [1..18] of string = ('Неизвестно','Неизвестно'
,'холостой (статус Idle)','Печть','Разогрев','Stopped Printing','Offline','Печать приостановлена'
,'Ошибка','Занят','Не подключен','Режим ожидания','Обработка','Инициализация'
,'Энергосбережение','В ожидании удаления','I/O Active','Ручная подача');
PrintLanguagesSupported :array [1..50] of string =('Other','Unknown','PCL ','HPGL'
,'PJL','PS','PSPrinter','IPDS','PPDS','EscapeP','Epson','DDIF','Interpress'
,'ISO6429','Line Data','MODCA','REGIS','SCS','SPDL','TEK4014','PDS','IGP','CodeV'
,'DSCDSE','WPS','LN03','CCITT','QUIC','CPAP','DecPPL','Simple Text','NPAP','DOC','imPress'
,'Pinwriter','NPDL','NEC201PL','Automatic','Pages','LIPS','TIFF','Diagnostic'
,'CaPSL','EXCL','LCDS','XES','MIME','XPS','HPGL2','PCLXL');
PrintMarkingTechnology :array [1..27] of string=('Other','Неизвестный','Электрофотографический светодиод','Лазер электрофотографический'
,'Электрофотографический','Матричный на 9pin','Матричный на 24pin','Матрица точечных движущихся головок с ударной головкой'
,'Ударная движущаяся головка ','Impact Band','Impact Other','Струйный водный','Струйный твердый','Струйное Другое'
,'Pen','Термическая передача','Термочувствительный','Термодиффузия','Тепловое Другое','Электроэрозия','Электростатический'
,'Photographic Microfiche','Photographic Imagesetter','Photographic Other','Ионное осаждение','eBeam','Typesetter');
PrintPrinterStatus  : array [1..7] of string =('Other','Неизвестный режим','Режим Idle','Режим печати','Разогрев','Остановка печати','Не в сети');


implementation
uses umain,PrintSecuretyEdit,AddPrintLan;



{$R *.dfm}
function VarToInt(AVariant: variant): integer;
begin
  //*** Если NULL или не числовое, то вернем значение по умолчанию
  Result := 0;
  if VarIsNull(AVariant) then
    Result := 0
  else
    {//*** Если числовое, то вернем значение}
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
FWMIService.Security_.authenticationLevel := 6; // 6 или WbemAuthenticationLevelPktPrivacy
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
           frmDomainInfo.memo1.Lines.Add('Ошибка операции SharedNetwork "'+E.Message);
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
Label1.Caption:='Порт:';
Label2.Caption:='Драйвер:';
Label3.Caption:='Диспетчер очереди печати:';
Label4.Caption:='Цветная печать:';
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
FormDriver.Caption:='Подробно: '+ DriverForSelectprinter;
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
      MemoDriver.Lines.Add('Имя драйвера: '+string(FWbemObject.Name));
      if FWbemObject.Version<>null then
     MemoDriver.Lines.Add('Версия драйвера: '+string(FWbemObject.Version));
      if FWbemObject.SupportedPlatform<>null then
      MemoDriver.Lines.Add('Платформа: '+string(FWbemObject.SupportedPlatform));
      if FWbemObject.OEMUrl<>null then
      MemoDriver.Lines.Add('OEMUrl: '+string(FWbemObject.OEMUrl));
      if FWbemObject.ConfigFile<>null then
      MemoDriver.Lines.Add('Файл конфигурации: '+string(FWbemObject.ConfigFile));
      if FWbemObject.DataFile<>null then
      MemoDriver.Lines.Add('Файл данных: '+string(FWbemObject.DataFile));
      if FWbemObject.DriverPath<>null then
     MemoDriver.Lines.Add('DriverPath: '+string(FWbemObject.DriverPath));
      if FWbemObject.FilePath<>null then
      MemoDriver.Lines.Add('FilePath: '+string(FWbemObject.FilePath));
      if FWbemObject.HelpFile<>null then
      MemoDriver.Lines.Add('HelpFile: '+string(FWbemObject.HelpFile));
      if VarIsArray(FWbemObject.DependentFiles) then
            begin
            massArray:=FWbemObject.DependentFiles;
            MemoDriver.Lines.Add('Зависимые файлы__________________________________');
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
             frmDomainInfo.memo1.Lines.Add('Ошибка операции '+E.Message);
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
        if vartoint(FWbemObject.Color)=1 then  label4.Caption:='Цветная печать: Нет'
        else label4.Caption:='Цветная печать: Да';
        end;
     { if FWbemObject.DriverVersion<>null then
        begin
        memo1.Lines.add('Версия драйвера FWbemObject.DriverVersion: '+string(FWbemObject.DriverVersion));
        memo1.Lines.add('------------------------------------------------------');
        end;}
      if FWbemObject.duplex<>null then
        begin
        if boolean(FWbemObject.duplex) then memo1.Lines.add('Двехсторонняя печать: Да')
        else memo1.Lines.add('Двухсторонняя печать: Нет');
        end;
       if (FWbemObject.HorizontalResolution<>null)and (FWbemObject.VerticalResolution<>null) then
        begin
        memo1.Lines.add('Разрешение печати (точек/дюйм): '+string(FWbemObject.HorizontalResolution)+'x'+string(FWbemObject.VerticalResolution));
        memo1.Lines.add('------------------------------------------------------');
        end;
        //////////////////////////////////////////////////////////
        if VarIsNumeric(FWbemObject.ICMMethod) then
        begin
          case vartoint(FWbemObject.ICMMethod) of
          1:memo1.Lines.add('Метод управления цветом: Отключен ');
          2:memo1.Lines.add('Метод управления цветом: Windows ');
          3:memo1.Lines.add('Метод управления цветом: На уровне драйвера устройства ');
          4:memo1.Lines.add('Метод управления цветом: Устройство ');
          else memo1.Lines.add('Метод управления цветом: Неизвестно ');
          end;

        end;
       if VarIsNumeric(FWbemObject.ICMIntent) then
          begin
          case vartoint(FWbemObject.ICMIntent) of
          1:memo1.Lines.add('Метод сопоставления цветов: Насышение ');
          2:memo1.Lines.add('Метод сопоставления цветов: Контраст');
          3:memo1.Lines.add('Метод сопоставления цветов: Точный цвет');
          else memo1.Lines.add('Метод управления цветом: Неизвестно ');
          end;
        end;
       ///////////////////////////////////////////////////////////
         if VarIsNumeric(FWbemObject.MediaType) then
        begin
          case vartoint(FWbemObject.MediaType) of
          1:memo1.Lines.add('Тип носителя, на котором печатает принтер: Стандарт ');
          2:memo1.Lines.add('Тип носителя, на котором печатает принтер: Transparency ');
          3:memo1.Lines.add('Тип носителя, на котором печатает принтер: Глянцевый');
          else memo1.Lines.add('Метод управления цветом: Неизвестно ');
          end;
        end;
       //////////////////////////////////////////////////////////////////////
       if VarIsNumeric(FWbemObject.Orientation) then
        begin
          case vartoint(FWbemObject.Orientation) of
          1:memo1.Lines.add('Ориентация бумаги при печати: Портрет ');
          2:memo1.Lines.add('Ориентация бумаги при печати: Пейзаж');
          else memo1.Lines.add('Ориентация бумаги при печати: ХЗ, криво как то! ');
          end;
        end;
       ///////////////////////////////////////////////////////////////////////////////
     FWbemObject:=Unassigned;
      end;
end;
  except
  on E:Exception do
           begin
             frmDomainInfo.memo1.Lines.Add('Ошибка операции "'+E.Message);
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
        else LinkLabel1.Caption:='Неизвестно';
        ////////////////////////////////////////////
         if FWbemObject.PortName<>null then
        begin
         Label1.Caption:='Порт: '+string(FWbemObject.PortName);
        end
        else Label1.Caption:='Порт: Неизвестно';
              //////////////////////////////////////////////////////
        if FWbemObject.PrintProcessor<>null then
        begin
         Label3.Caption:='Диспетчер очереди печати: '+string(FWbemObject.PrintProcessor);
        end else Label3.Caption:='Диспетчер очереди печати: Неизвестно';
      ////////////////////////////////////////////////////////
     if (FWbemObject.ServerName<>null)  then
      memo1.Lines.add('Имя сервера, который управляет принтером (FWbemObject.ServerName): '+string(FWbemObject.ServerName));
      /////////////////////////////////////////////////////////////////////////////////
     {    if FWbemObject.Hidden<>null then //Если TRUE , принтер скрыт от пользователей сети.
     if boolean(FWbemObject.Hidden)=false then
     begin
      memo1.Lines.add('Принтер доступен как общесетевой ресурс (FWbemObject.Hidden):'+vartostr(FWbemObject.Hidden));
      memo1.Lines.add('------------------------------------------------------');
     end; }
     ////////////////////////////////////////////////////////////////////////
      //Shared принтер доступен как сетевой ресурс
      CheckBox1.Checked:=boolean(FWbemObject.Shared);
      LabeledEdit1.Enabled:=CheckBox1.Checked;
      CheckBox2.enabled:=CheckBox1.Checked;
      ////////////////////////////////////////////////////////////////////////////
      if (FWbemObject.ShareName<>null)  then
     begin
     // memo1.Lines.add('Имя устройства как общесетевого ресурса (FWbemObject.ShareName): '+string(FWbemObject.ShareName));
     // memo1.Lines.add('------------------------------------------------------');
      LabeledEdit1.Text:=string(FWbemObject.ShareName);
     end;
      ///////////////////////////////////////////////////////////////
     { if (FWbemObject.status<>null)  then
      memo1.Lines.add('Текущий статус принтера (FWbemObject.status): '+string(FWbemObject.status));
       /// ///////////////////////////////////////////////////
     if (FWbemObject.network<>null)  then
      memo1.Lines.add('Сетевой принтер (FWbemObject.network): '+booltostr(FWbemObject.network));
     if ((FWbemObject.local)<>null) then
      memo1.Lines.add('FWbemObject.local: '+booltostr(FWbemObject.local));
      memo1.Lines.add('------------------------------------------------------');}
     /////////////////////////////////////////////////////////////////////
     if FWbemObject.Published<>null then
     begin
      CheckBox2.Checked:= boolean(FWbemObject.Published);
     end;
////////////////////////////////////////////////////////////// расположение
     if FWbemObject.Location<>null then LabeledEdit2.Text:=string(FWbemObject.Location)
     else LabeledEdit2.Text:='';
       //////////////////////////////////////////////////////////
      { masvar:=FWbemObject.Capabilities;
       PrintPropertyString:='';
       for i := 0 to VarArrayHighBound(masvar, 1) do
        begin
        PrintPropertyString := PrintPropertyString + PrintCapabilities[vartoint(masvar[i])] + ', ';
        end;
       memo1.lines.add('Поддерживаемые функции: '+PrintPropertyString);
       VarClear(masvar);
       memo1.Lines.add('------------------------------------------------------'); }
      //////////////////////////////////////////////
      {if FWbemObject.Availability<>null then
         memo1.lines.add('Доступность: '+(PrintAvailability[integer(FWbemObject.Availability)]));
         memo1.Lines.add('------------------------------------------------------'); }
      ////////////////////////////////////////////////
      if FWbemObject.ConfigManagerErrorCode<>null then
         memo1.lines.add('Ошибки: '+PrintConfigManagerErrorCode[integer(FWbemObject.ConfigManagerErrorCode)]);
         /////////////////////////////////////////////
       if (FWbemObject.DetectedErrorState<>null)  then
         begin
          if integer(FWbemObject.DetectedErrorState)<>0 then
          memo1.lines.add('Состояние принтера: '+PrintDetectedErrorState[integer(FWbemObject.DetectedErrorState)]);
         end;
        //////////////////////////////////////////////////
        if (FWbemObject.ExtendedDetectedErrorState<>null)  then
          begin
           if (integer(FWbemObject.ExtendedDetectedErrorState)<>0) then
           memo1.lines.add('Информация об ошибках: '+PrintExtendedDetectedErrorState[integer(FWbemObject.ExtendedDetectedErrorState)]);
          end;
          /////////////////////////////////////////////////
          if (FWbemObject.ExtendedPrinterStatus<>null)  then
          begin
          if (integer(FWbemObject.ExtendedPrinterStatus)<>0) then
           memo1.lines.add('Информация о текущем состоянии: '+PrintExtendedPrinterStatus[integer(FWbemObject.ExtendedPrinterStatus)]);
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
            memo1.lines.add('Поддерживаемые языки печати: '+PrintPropertyString);
            VarClear(masvar);
           end;
          ///////////////////////////////////////
          if (FWbemObject.MarkingTechnology<>null)  then
          begin
           if (integer(FWbemObject.MarkingTechnology)<>0) then
           memo1.lines.add('Технология печати: '+PrintMarkingTechnology[integer(FWbemObject.MarkingTechnology)]);
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
            memo1.lines.add('Поддерживаемые форматы мумаги: '+PrintPropertyString);
            memo1.Lines.Add('_________________________________________________');
            VarClear(masvar);
            end;
          ////////////////////////////////////////////////////////////
         if (FWbemObject.HorizontalResolution<>null) and (FWbemObject.VerticalResolution<>null) then
          begin
          memo1.lines.add('Разрешение принтера: '+string(FWbemObject.HorizontalResolution)+'x'+string(FWbemObject.VerticalResolution));
          end;
          /// /////////////////////////////////////////////////////////
          if FWbemObject.MaxCopies<>null then
          memo1.lines.add('Максимальное количество копий (FWbemObject.MaxCopies): '+(vartostr(FWbemObject.MaxCopies)));
         ////////////////////////////////////////////////////////////////
          if FWbemObject.MaxNumberUp<>null then
          memo1.lines.add('Максимальное количество страниц на листе (FWbemObject.MaxNumberUp): '+(vartostr(FWbemObject.MaxNumberUp)));
        end;
    FWbemObject:=Unassigned;
end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка операции:'+E.Message);
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
frmDomainInfo.memo1.Lines.Add('Результат операции PrinterNetworkShared: '+i);
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
  memo1.Lines.Add('Цикл VarIsArray(VarArray['+inttostr(i)+'])');
  Win32ACE:=varArray[i];

  ACETrustee:=Win32ACE.Trustee;
  Win32TrusteeDecompose(ACETrustee);
  if not(VarIsNull(Win32ACE.TIME_CREATED)) then
  Memo1.Lines.Add('Время создания:'+string(Win32ACE.TIME_CREATED));
  if not(VarIsNull(Win32ACE.AccessMask)) then
  Memo1.Lines.Add('Маска доступа (сумма прав):'+string(Win32ACE.AccessMask));
  if not(VarIsNull(Win32ACE.AceFlags)) then
  Memo1.Lines.Add('Наследование Прав:'+string(Win32ACE.AceFlags));
  if not(VarIsNull(Win32ACE.AceType)) then
  Memo1.Lines.Add('Доступ 0/1/2:'+string(Win32ACE.AceType));
  if not(VarIsNull(Win32ACE.GuidInheritedObjectType)) then
  Memo1.Lines.Add('Guid родителя:'+string(Win32ACE.GuidInheritedObjectType));
  if not(VarIsNull(Win32ACE.GuidObjectType)) then
  Memo1.Lines.Add('Guid объекта:'+string(Win32ACE.GuidObjectType));



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
        else Memo1.Lines.Add('PrintSecurityDescriptor.dacl не массив');

      if VarIsArray(PrintSecurityDescriptor.sacl) then
      begin
      SDACESACLArray:=PrintSecurityDescriptor.sacl;
      if Win32ACEDecompose(SDACESACLArray)then memo1.Lines.add('Win32ACEDecompose(SDACESACLArray)');
      end
        else Memo1.Lines.Add('PrintSecurityDescriptor.sacl не массив');

      SDTrusteeGroup:=PrintSecurityDescriptor.group;
      if Win32TrusteeDecompose(SDTrusteeGroup) then Memo1.Lines.Add('Выгружены данные SDTrusteeGroup');

      SDTrusteeOwner:=PrintSecurityDescriptor.owner;
      if Win32TrusteeDecompose(SDTrusteeOwner) then Memo1.Lines.Add('Выгружены данные SDTrusteeOwner');


      ///BinarySDToSecurityDescriptor(NewSecurityDescriptor,PrintSecurityDescriptor,
     /// pchar(MyPS),pchar(MyUser),pchar(MyPasswd),0);
     // GetSecurityDescriptorDacl(NewSecurityDescriptor,true,null,true);
      end;
   memo1.Lines.Add('Результат операции SysErrorMessage: '+  SysErrorMessage(SDescriptorError));

end;
except
on E:Exception do
         begin

           memo1.Lines.Add('Ошибка операции SysErrorMessage: '+  SysErrorMessage(SDescriptorError));
           memo1.Lines.Add('Ошибка операции PrintSecurityDescriptor: '+E.Message);
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
