unit SMARTStat;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.Buttons, ComObj, Vcl.OleCtrls, ActiveX, Vcl.ComCtrls, Vcl.ImgList,
  Vcl.ExtCtrls, Vcl.ButtonGroup, Vcl.Samples.Spin,
   Vcl.Menus, Vcl.Imaging.pngimage, System.TimeSpan,
  Vcl.Samples.Gauges, System.ImageList;

type
  TForm8 = class(TForm)
    ListSmart: TListView;
    GroupBox2: TGroupBox;
    ImageSMART: TImageList;
    Memo1: TMemo;
    GroupDisk: TButtonGroup;
    ImageHDD: TImageList;
    GroupBox1: TGroupBox;
    GroupBox5: TGroupBox;
    SpeedButton11: TSpeedButton;
    Memo2: TMemo;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton4: TSpeedButton;
    EnableDisableHFP: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Button1: TButton;
    EnableDisableFPP: TPopupMenu;
    Button2: TButton;
    N3: TMenuItem;
    N4: TMenuItem;
    Button3: TButton;
    AllowPerformanHit: TPopupMenu;
    N5: TMenuItem;
    N6: TMenuItem;
    PopupTest: TPopupMenu;
    Button4: TButton;
    N0Offlinetestroutin1: TMenuItem;
    N1Shortselftestroutin1: TMenuItem;
    N2Extendedselftestroutin1: TMenuItem;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox6: TGroupBox;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    StaticText6: TStaticText;
    StaticText7: TStaticText;
    GroupBox7: TGroupBox;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    Image9: TImage;
    StaticText8: TStaticText;
    StaticText10: TStaticText;
    StaticText11: TStaticText;
    StaticText12: TStaticText;
    StaticText13: TStaticText;
    StaticText14: TStaticText;
    StaticText15: TStaticText;
    GroupBox8: TGroupBox;
    StaticText9: TStaticText;
    StaticText16: TStaticText;
    StaticText17: TStaticText;
    StaticText18: TStaticText;
    StaticText19: TStaticText;
    StaticText20: TStaticText;
    ProgressBar1: TProgressBar;
    Procedure infodisk;
    procedure ComboBox1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GroupDiskButtonClicked(Sender: TObject; Index: Integer);
    procedure SpeedButton6Click(Sender: TObject);
    Function AllowPerformanceHit(onOff: bool): string;
    Function EnableDisableFailurePredictionPolling(onOff: bool): string;
    function EnableDisableHardwareFailurePrediction(onOff: bool): string;
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure EnableOfflineDiags;
    procedure SmartSelect;
    procedure SpeedButton11Click(Sender: TObject);
    procedure RefreshDisk;
    procedure PhysicalDisk;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    function Obrez(s:string):string;
    function starttest(s,b:string):string;
    function imageitems(Item: TListItem):boolean;
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure OnMyMenuItemClick(Sender: TObject);
    procedure N0Offlinetestroutin1Click(Sender: TObject);
    procedure N1Shortselftestroutin1Click(Sender: TObject);
    procedure N2Extendedselftestroutin1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form8: TForm8;
  Viewimage: Integer;
  smartprioryty: bool;
  // HDDmass:array [0..31] of integer;
  selectHDD,SelectHDDIndex: string;
  Typesmart,SizeLBA: Integer;
  Alllifehdd:integer;
  Atribute241,Atribute242,Atribute243:int64;
const
  wbemFlagForwardOnly = $00000020;

implementation

uses
  uMain;
{$R *.dfm}

function TForm8.Obrez(s:string):string;
  var
  x,y:integer;
  begin
  result:='';
  delete(s,1,pos('\',s));
  delete(s,1,pos('\',s));
  result:=s;
  //result:=copy(s,1,pos('\',s)-1);
  //Memo2.Lines.Add(result);
  end;

function Dectobin(id, col: Integer): string;
var
  delimoe: Integer;
  StrBin: String;
begin
  delimoe := id;
  StrBin := '';
  repeat
    StrBin := IntToStr(delimoe mod 2) + StrBin;
    delimoe := delimoe div 2;
  until (delimoe = 0);
  while Length(StrBin) <> col do
    StrBin := '0' + StrBin;

  Result := StrBin;
end;

function BIN2DEC(BIN: string): LONGINT;
var
  J: LONGINT;
  Error: Boolean;
  DEC: LONGINT;
begin
  DEC := 0;
  Error := False;
  for J := 1 to Length(BIN) do
  begin
    if (BIN[J] <> '0') and (BIN[J] <> '1') then
      Error := True;
    if BIN[J] = '1' then
      DEC := DEC + (1 shl (Length(BIN) - J));
    { (1 SHL (Length(BIN) - J)) = 2^(Length(BIN)- J) }
  end;
  if Error then
    BIN2DEC := 0
  else
    BIN2DEC := DEC;
end;

function SelfTestStat(id: Integer): string;
begin
  case id of
    0:
      Result := 'Предыдущая процедура самопроверки завершена без ошибок.';
    1:
      Result := 'Самопроверка была прервана хостом!';
    2:
      Result := 'Самопроверка была прервана хостом с аппаратным или программным сбросом!!!';
    3:
      Result := 'Неустранимая ошибка или неизвестная ошибка теста произошла во время выполнения устройством процедуры самопроверки и устройство не смогло завершить процедуру самопроверки.';
    4:
      Result := 'The previous self-test completed having a test element that failed and the test element that failed is not known.';
    5:
      Result := 'Предыдущее самотестирование завершилось с отказом электрического элемента испытания.';
    6:
      Result := 'Предыдущее самотестирование завершилось с отказом сервопривода и/или элемента теста поиска теста.';
    7:
      Result := 'Предыдущее самотестирование завершилось с ошибкой элемента read теста. (The previous self-test completed having the read element of the test failed.)';
    8:
      Result := 'Предыдущее самодиагностирование завершено с отказавшим тестовым элементом, возможно устройство имеет повреждения!!!!!';
    15:
      Result := 'Производится самопроверка (Self-test).';
  else
    Result := 'Неизвестный од ошибка - ' + IntToStr(id);
  end;
end;

function OffCollStatus(id: Integer): string;
begin
  case id of
    0:
      Result := 'Автономный сбор данных не производился';
    2:
      Result := 'Автономный сбор данных завершен без ошибок';
    3:
      Result := 'Работа в автономном режиме';
    4:
      Result := 'Автономный сбор данных был приостановлен командой прерывания от хоста';
    5:
      Result := 'Операция сбора данных в автономном режиме была прервана командой от хоста';
    6:
      Result := 'Автономный сбор данных был прерван устройством с фатальной ошибкой';
    128:
      Result := 'Автономный сбор данных не производился';
    130:
      Result := 'Автономный сбор данных завершен без ошибок';
    132:
      Result := 'Автономный сбор данных был приостановлен командой прерывания от хоста';
    133:
      Result := 'Операция сбора данных в автономном режиме была прервана командой от хоста';
    134:
      Result := 'Автономный сбор данных был прерван устройством с фатальной ошибкой';
  else
    Result := 'Неизвестный код статуса - ' + IntToStr(id);
  end;
end;

/// /////////////////////функция возвращающая имя параметра
function smartdscptn(z: Integer): string;

begin
  smartprioryty := False;
  Viewimage := -1;
  case z of
    1:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Ошибки аппаратного чтения (Read Error Rate)';
      end;
    2:
      begin
        Viewimage := 1;
        smartprioryty := true;
        Result := 'Пропускная способность жесткого диска';
      end;
    3:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := '	Среднее время раскрутки шпинделя (млск) (Spin-up Time)';
      end;
    4:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Подсчет циклов запуска / остановки шпинделя (Start/Stop Count)';
      end;
    5:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Количество перераспределенных секторов (Re-allocated Sector Count)';
      end;
    6:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Маржа канала при чтении данных';
      end;
    7:
      begin
        Viewimage := 1;
        smartprioryty := false;
        Result := 'Частота ошибок при позиционировании блока магнитных головок (Seek Error Rate)';
      end;
    8:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Производительность операций поиска магнитных головок';
      end;
    9:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Отработано часов (Power-on Hours Count)';
      end;
    10:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Количество повторных попыток запуска вращения (Spin-up Retry Count)';
      end;
    11:
      begin
        Viewimage := 0;
        smartprioryty := false;
        Result := 'Счетчик повторений калибровки привода (Drive Calibration Retry Count)';
      end;
    12:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Кол-во включений (Drive Power Cycle Count)';
      end;
    13:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Программные ошибки чтения (Soft Read Error Rate)';
      end;
    22:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Уровень Гелия в HGST';
      end;
    100:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Gigabytes Erased (Count*64Gb) Kingston SSD';
      end;
      148:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Total Erase Count SLC (single-level cell)';
      end;
      149:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Maximum Erase Count SLC (single-level cell)';
      end;
      150:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Minimum Erase Count SLC (single-level cell)';
      end;
      151:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Average Erase Count SLC (single-level cell)';
      end;
      160:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Количество некорректных секторов. Uncorrectable Sector Count';
      end;
      161:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Количество запасных блоков. Valid Spare Blocks';
      end;
      163:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Недействительные блоки. Initial Invalid Blocks';
      end;
      164:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Total Erase Count TLC (Three Level Cell)';
      end;
      165:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Maximum Erase Count TLC (Three Level Cell)';
      end;
      166:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Minimum Erase Count TLC (Three Level Cell)';
      end;
       167:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Average Erase Count TLC (Three Level Cell)';
      end;
        168:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'SATA PHY Error Count';
      end;
       169:
      begin
        Viewimage := 0;
        smartprioryty := true;
        Result := 'Remain Life Percentage';
      end;
    170:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Доступное зарезервированное пространство (Available Reserved Space)';
      end;
    171:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество ошибок Frimware SSD (Program Fail Count)';
      end;
    172:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество сбоев флэш-памяти SSD (Erase Fail Count)';
      end;
    173:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Максимальное количество худших стираний в любом блоке (SSD)';
      end;
    174:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Неожиданное отключение питания';
      end;
    175:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Результат последнего теста (Power Loss Protection Failure)';
      end;
    176:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Количество сбоев команды быстрого стирания';
      end;
    177:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := '	Дельта между наиболее изношенными и наименее изношенными блоками Flas memory в накопителях SSD';
      end;
    179:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Общее количество использованных зарезервированных блоков';
      end;
    180:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Всего неиспользованных зарезервированных блоков';
      end;
    181:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Общее количество сбоев работы Firmware Flash (Program Fail Count)';
      end;
    182:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Число ошибок стирания (Erase Fail Count)';
      end;
    183:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'SATA Downshift Error Count or Runtime Bad Block';
      end;
    184:
      begin
        Viewimage := 1;
        smartprioryty := True;
        Result := 'Количество ошибок четности (Reported I/O Error Detection Code Errors )';
      end;
    185:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Head Stability';
      end;
    186:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Обнаружение вибрации';
      end;
    187:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Кол-во неисправимых ошибок (Reported Uncorrectable Errors (URAISE))';
      end;
    188:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Количество прерванных операций из-за тайм-аута жесткого диска!!!';
      end;
    189:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'High Fly write';
      end;
    190:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Температура воздушного потока. C°.';
      end;
    191:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Ошибоки, вызванные внешними воздействиями удара или вибрации.';
      end;
    192:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Счетчик аварийного отключения питания (Power-Off Retract Count (Unsafe Shutdown Count))';
      end;
    193:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Подсчет циклов загрузки / выгрузки';
      end;
    194:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Температура. C°';
      end;
    195:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Количество ошибок скорректированные ECC (Hardware ECC Recovered)/Дополительно смотрите 1й атрибут.';
      end;
    196:
      begin
        Viewimage := 1;
        smartprioryty := True;
        Result := 'Количество операций переназначения секторов!!! (Reallocation Event Count)/Дополнительно смотрите 5й атрибут.';
      end;
    197:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Ожидающих секторов для переназначения!!! (Current Pending Sector Count)';
      end;
    198:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Количество неисправимых ошибок при Short Extended self-test routin!!! (Uncorrectable Sector Count)';
      end;
    199:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Ошибок при передаче данных через интерфейсный кабель (SATA R-Errors Error count)';
      end;
    200:
      begin
        Viewimage := 1;
        smartprioryty := True;
        Result := 'Количество ошибок, найденных при записи (Write Error Rate)'; //// ноль плохо
      end;
    201:
      begin
        Viewimage := 0;
        smartprioryty := True;
        Result := 'Количество ошибок чтения по вине программного обеспечения (Soft Read Error Rate)';
      end;
    202:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Количество ошибок метки адреса данных (Data Address Mark Error)'; ///  0 плохо
      end;
    203:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество ошибок, вызванных неверной CRC при исправлении ошибки';
      end;
    204:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Коррекция ECC (Soft ECC Correction Rate (UECC))';
      end;
    205:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Подсчет ошибок из-за высокой температуры';
      end;
    206:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Высота головок над поверхностью диска.';
      end;
    207:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Величина импульсного тока, используемого для вращения привода';
      end;
    208:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Spin Buzz';
      end;
    209:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Производительность поиска диска во время его внутренних тестов';
      end;
    210:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Вибрация во время записи';
      end;
    211:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Вибрация во время записи';
      end;
    212:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Запись шока, обнаруженного во время операций записи.';
      end;
    218:
      begin
       Viewimage := 5;
       smartprioryty := False;
       Result := 'CRC Error Count (Kingston SSD)';
      end;
    220:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Сдвиг диска (обычно из-за удара или температуры).';
      end;
    221:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество ошибок, вызванных внешними воздействиями удара и вибрации.';
      end;
    222:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Время работы диска под нагрузкой.';
      end;
    223:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Количество раз голова меняет положение.';
      end;
    224:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Сопротивление, вызванное трением в механических деталях во время работы.';
      end;
    225:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Общее количество циклов нагрузки';
      end;
    226:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Общее время загрузки головок (Timed Workload Media Wear)';
      end;
    227:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество попыток компенсировать колебания скорости диска.';
      end;
    228:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество выключений или входа в режим сна';
      end;
    230:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'GMR Head Amplitude, Drive Life Protection Status';
      end;
    231:
      begin
        Viewimage := 1;
        smartprioryty := True;
        Result := 'Указывает приблизительный оставшийся срок службы SSD!!! (SSD Life)';
      end;
    232:
      begin
        Viewimage := 1;
        smartprioryty := True;
        Result := 'Available Reserved Space (Gb)';
      end;
    233:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Индикатор износа носителя (SSD).';
      end;
    234:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Average erase count AND Maximum Erase Count (Среднее и максимальное число стираний)';
      end;
    235:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Good Block Count AND System(Free) Block Count';
      end;
    240:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Время, затрачиваемое на позиционирование БМГ.';
      end;
    241:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := '	Общее количество написанных LBA.';
      end;
    242:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Общее количество прочитанных LBA.';
      end;
    243:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Всего написано LBA. Расширено. Добавляется к 241 атрибуту.';
      end;
    244:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result :='Average Erase Count (Kingston SSD)';
        //Result := 'Всего LBA Прочитано. Расширенное. Добавляется к 242 атрибуту.';
      end;
     245:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Max Erase Count (Kingston SSD)';
      end;
     246:
      begin
       Viewimage := 5;
        smartprioryty := False;
        Result := 'Total Erase Count (Kingston SSD)';
      end;
    249:
      begin
        Viewimage := 5;
        smartprioryty := False;
        Result := 'Значение сообщает о количестве записей в NAND с шагом 1 ГБ.';
      end;
    250:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество ошибок при чтении с диска!!!';
      end;
    251:
      begin
        Viewimage := 1;
        smartprioryty := False;
        Result := 'Minimum Spares Remaining. Количество оставшихся запасных блоков в %.';
      end;
    252:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Общее количество сбойных флэш-блоков';
      end;
    254:
      begin
        Viewimage := 0;
        smartprioryty := False;
        Result := 'Количество обнаруженных событий «Свободное падение» (Free Fall Event Count).';
      end;
  else
    Result := 'Vendor Specific';
  end

end;

function namedisk(s: string): string;
begin
  if s <> '' then
  begin
    delete(s, 1, pos('\', s));
    delete(s, pos('__', s), Length(s));
    Result := s;
  end
  else
    Result := s;
  s := '';
end;

procedure TForm8.ComboBox1Change(Sender: TObject);
begin
  Memo1.Clear;
  ListSmart.Clear;
  GroupBox2.Caption := 'S.M.A.R.T - ?';
  infodisk;
end;

procedure TForm8.FormShow(Sender: TObject);
begin
  //ProgressBar1.Brush.Color:= clRed;
  //PostMessage(ProgressBar1.Handle, $0409, 0, clGreen);
  // ComboBox1.Clear;
  // ComboBox1.Text:='Обновите список дисков';
  Alllifehdd:=0; //// если 0 тодиск живой на 100%
  ProgressBar1.Position:=0;
  StaticText17.Caption:='Здоровье %';
  StaticText9.Caption:='Текущая - C°';
  StaticText16.Caption:='Минимальная - C°';
  StaticText18.Caption:='Отработано -  (час)';
  StaticText19.Caption:='Всего прочитано - Гб';
  StaticText20.Caption:='Всего записано - Гб';
  GroupBox2.Caption := 'S.M.A.R.T - ?';
  GroupBox3.Caption:='Подробно о диске';
  memo2.Clear;
  Memo1.Clear;
  ListSmart.Clear;
  GroupDisk.Items.Clear;
  RefreshDisk;
end;

procedure TForm8.GroupDiskButtonClicked(Sender: TObject; Index: Integer);
begin
  if GroupDisk.Items[index].ImageIndex=1 then Alllifehdd:=100 /// если оценка диска ОС говорит все плохо то диск не жилец
  else  Alllifehdd:=0; //// если 0 то диск живой на 100%
  ProgressBar1.Position:=0;
  StaticText17.Caption:='Здоровье %';
  StaticText9.Caption:='Текущая - C°';
  StaticText16.Caption:='Минимальная - C°';
  StaticText18.Caption:='Отработано -  (час)';
  StaticText19.Caption:='Всего прочитано - Гб';
  StaticText20.Caption:='Всего записано - Гб';
  selectHDD:='';
  Memo1.Clear;
  memo2.Clear;
  ListSmart.Clear;
  GroupBox2.Caption := 'S.M.A.R.T - ?';
  selectHDD := GroupDisk.Items[Index].Hint;
  SelectHDDIndex:=copy(GroupDisk.Items[Index].Caption,8,1);
  GroupBox3.Caption:='Подробно о диске - '+GroupDisk.Items[Index].Caption;
  infodisk; /// инфо о диске
  GroupBox2.Caption:=GroupDisk.Items[Index].Caption;
  SmartSelect; /// SMART диска
  end;

Procedure TForm8.infodisk;
var
/// / сканирование информации о выбранном диске
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObjectSetPart: OLEVariant;
  oEnum: IEnumvariant;
  oEnumPart: IEnumvariant;
  FWbemObject: OLEVariant;
  FWbemObjectPart: OLEVariant;
  iValue: LongWord;
  iValuePart: LongWord;
  masvar: variant;
  i: Integer;
  CapabilityDes: string;
  idHard: Integer;
  /// //////////////////////////////////////////////////////////////////////////
  function hddsn(iid: string): string;
  /// // S/N HDD
  var
    FWbemObjectSetSN: OLEVariant;
    oEnumSN: IEnumvariant;
    FWbemObjectSN: OLEVariant;
    iValueSN: LongWord;
  begin
  try
    iid:=copy(iid,5,length(IiD));
    Result := '';
    FWbemObjectSetSN := FWMIService.ExecQuery
      ('SELECT * FROM Win32_PhysicalMedia WHERE Tag LIKE "%'+IiD+'%"', 'WQL',wbemFlagForwardOnly);
    oEnumSN := IUnknown(FWbemObjectSetSN._NewEnum) as IEnumvariant;
    while oEnumSN.Next(1, FWbemObjectSN, iValueSN) = 0 do
    begin
        Result := vartostr(FWbemObjectSN.SerialNumber);
    end;
    oEnumSN := nil;
    VariantClear(FWbemObjectSN);
    VariantClear(FWbemObjectSetSN);
  except
  on E: Exception do
  begin
  memo1.Lines.Add('Ошибка получения SN - '+e.Message);
  end;
  end;
  end;
///////////////////////////////////////////////////////////////////////////////
 function hddBukva(iid: string): string;
  var
    FWbemObjectSetbu: OLEVariant;
    oEnumbukva: IEnumvariant;
    FWbemObjectbu: OLEVariant;
    iValueSN: LongWord;
    s:string;
  begin
  try
    Result := '';
    s:='';
    FWbemObjectSetbu := FWMIService.ExecQuery //  ('SELECT Antecedent,Dependent FROM Win32_LogicalDiskToPartition', 'WQL',wbemFlagForwardOnly);
      ('SELECT Antecedent,Dependent FROM Win32_LogicalDiskToPartition WHERE Antecedent LIKE "%'+iid+'%"', 'WQL',wbemFlagForwardOnly);
    oEnumbukva := IUnknown(FWbemObjectSetbu._NewEnum) as IEnumvariant;
    while oEnumbukva.Next(1, FWbemObjectbu, iValueSN) = 0 do
    begin
        memo1.Lines.add(vartostr(FWbemObjectbu.Dependent));
        s:=vartostr(FWbemObjectbu.Dependent);
        s := copy(s,pos('=',s),length(s));
        if s<>'' then result:=s
        else result:='Unknown';
    end;
    oEnumbukva := nil;
    VariantClear(FWbemObjectbu);
    VariantClear(FWbemObjectSetbu);
  except
  on E: Exception do
  begin
  memo1.Lines.Add('Ошибка получения ,буквы диска - '+e.Message);
  end;
  end;
  end;
/// ///////////////////////////////////////////////////////////////////////
begin
  idHard := 0;
  try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    /// ////////////////////////////////////////////////////////////////////
    /// информация о дисках
    /// /////////////////////////////////////////////////////////////////////
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,'root\CIMV2', frmDomainInfo.LabeledEdit1.Text,frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery ('SELECT * FROM Win32_DiskDrive WHERE index='+SelectHDDIndex+'', 'WQL',wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      CapabilityDes := '';
      masvar := FWbemObject.CapabilityDescriptions;
      for i := 0 to VarArrayHighBound(masvar, 1) do
      begin
        if i = VarArrayHighBound(masvar, 1) then
          CapabilityDes := CapabilityDes + vartostr(masvar[i])
        else
          CapabilityDes := CapabilityDes + vartostr(masvar[i]) + ', ';
      end;
      VarClear(masvar);
      Memo1.Lines.Add('Модель - ' + vartostr(FWbemObject.model));
      Memo1.Lines.Add('ID - ' + vartostr(FWbemObject.DeviceID));
      Memo1.Lines.Add('S/N - ' + trim(hddsn(vartostr(FWbemObject.DeviceID))));
      Memo1.Lines.Add('Firmware Revision - ' +
        vartostr(FWbemObject.FirmwareRevision));
      Memo1.Lines.Add('Тип - ' + vartostr(FWbemObject.MediaType));
      Memo1.Lines.Add('Поддерживаемые функции - ' + CapabilityDes);
      /// ///////////////////////////////////////////////////////////////////////////
      if FWbemObject.CompressionMethod <> null then
        Memo1.Lines.Add('Сжатие - ' + vartostr(FWbemObject.CompressionMethod))
      else
        Memo1.Lines.Add('Сжатие - Неизвестно');
      if FWbemObject.ErrorDescription <> null then
        Memo1.Lines.Add('Ошибка - ' + vartostr(FWbemObject.ErrorDescription));
      if FWbemObject.InstallDate <> null then
        Memo1.Lines.Add('Дата установки - ' +
          vartostr(FWbemObject.InstallDate));
      Memo1.Lines.Add('Объем - ' + IntToStr(FWbemObject.Size / 1024 / 1024)
        + ' Мб.');
      Memo1.Lines.Add('Кол-во байт на сектор - ' +
        vartostr(FWbemObject.BytesPerSector));
       SizeLBA:=FWbemObject.BytesPerSector;
      Memo1.Lines.Add('Кол-во цилиндров на диске - ' +
        IntToStr(FWbemObject.TotalCylinders));
      Memo1.Lines.Add('Кол-во дорожек в каждом цилиндре - ' +
        IntToStr(FWbemObject.TracksPerCylinder));
      Memo1.Lines.Add('Кол-во головок на диске - ' +
        IntToStr(FWbemObject.TotalHeads));
      Memo1.Lines.Add('Кол-во секторов на диске - ' +
        IntToStr(FWbemObject.TotalSectors));
      Memo1.Lines.Add('Кол-во дорожек на диске - ' +
        IntToStr(FWbemObject.TotalTracks));
      Memo1.Lines.Add('Статус (оценка ОС)  - ' + vartostr(FWbemObject.status));
      /// ////////////////////////////////////////////// логические диски
      FWbemObjectSetPart := FWMIService.ExecQuery
        ('SELECT * FROM Win32_DiskPartition WHERE DiskIndex=' +
        IntToStr(FWbemObject.Index) + '', 'WQL', wbemFlagForwardOnly);
      oEnumPart := IUnknown(FWbemObjectSetPart._NewEnum) as IEnumvariant;
      while oEnumPart.Next(1, FWbemObjectPart, iValuePart) = 0 do
      begin
        Memo1.Lines.Add('------------------ Диск - ' +IntToStr(FWbemObjectPart.Diskindex) + '---- Раздел - '
         +IntToStr(FWbemObjectPart.Index) + '-----------------------');
        //if FWbemObjectPart.caption <> null then memo1.lines.Add('Буква диска - '+ hddbukva(vartostr(FWbemObjectPart.caption)));
        if FWbemObjectPart.status <> null then Memo1.Lines.Add(vartostr(FWbemObjectPart.status));
        if FWbemObjectPart.Bootable then  Memo1.Lines.Add('Загрузочный раздел');
        if FWbemObjectPart.BootPartition then  Memo1.Lines.Add('Активный раздел');
        if FWbemObjectPart.PrimaryPartition then  Memo1.Lines.Add('Основной раздел');
        Memo1.Lines.Add('Объем - ' + IntToStr(FWbemObjectPart.Size / 1024 /1024) + ' Мб');
        Memo1.Lines.Add('Размер блока - ' + IntToStr(FWbemObjectPart.BlockSize));
        Memo1.Lines.Add('Кол-во последовательных блоков - ' + IntToStr(FWbemObjectPart.NumberOfBlocks));
        Memo1.Lines.Add('Начальное смещение раздела - ' + IntToStr(FWbemObjectPart.StartingOffset));
      end;
      VariantClear(FWbemObjectPart);
      VariantClear(FWbemObjectSetPart);
      oEnumPart := nil;
      inc(idHard);
      FWbemObject := Unassigned;
      /// ///////////////////////////////////////////Логические диски
    end;

  except
    on E: Exception do
    begin
      Memo1.Lines.Add('ошибка получения подробных данных о диске- "' +
        E.Message + '"');
      exit;
    end;
  end;

  /// ////////////////////////////////////////////////////////////////////////
  
  /// //////////////////////////////////////////////////////////////////////
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  VariantClear(FWbemObject);
end;



function TForm8.imageitems(Item: TListItem):boolean;
var
  i,LifeHDD: Integer;
  s:string;
const
  mass: array [0 .. 19] of string = ('1', '3', '5', '7', '10', '11', '179',
    '184', '187', '188', '190', '194', '196', '197', '198', '200', '201', '231',
    '232', '233');
 function exceptAtribute(z:integer):bool;
 var
 i:integer;
 const
  mass: array [0 .. 2] of integer = (190,194,199);
 begin
 result:=false;
 for I := 0 to Length(mass) do
   begin
     if mass[i]=z then
     begin
       result:=true;
       break;
     end;
   end;
 end;

  function ocenka(th:integer):integer;
  begin
  result:=0;
    case th of
    0:  result:=5;
    1:  result:=5;
    2:  result:=6;
    3:  result:=7;
    4:  result:=8;
    5:  result:=9;
    6:  result:=10;
    7:  result:=11;
    8:  result:=13;
    9,10: result:=round(th+((th/100)*60));
    11..20:result:=round(th+((th/100)*50));
    21..30:result:=round(th+((th/100)*45));
    31..40:result:=round(th+((th/100)*40));
    41..50: result:=round(th+((th/100)*35));
    51..70:result:=round(th+((th/100)*30));
    71..90:result:=round(th+((th/100)*25));
    91..110:result:=round(th+((th/100)*20));
    111..140:result:=round(th+((th/100)*15));
    else result:=round(th+((th/100)*10));
    end;
  end;
begin
 try
  lifehdd:=0;
  case strtoint(item.Caption) of
/////////////////////////////////////////////////////////////////////////////////////////////////////
  1:
    begin

    if (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) or
     (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) then Item.ImageIndex:=6;

    if(strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) or
    (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))
      then item.ImageIndex:=3
     else Item.ImageIndex:=2;

    if (Item.SubItems[4]='0') then begin item.ImageIndex:=2; end; /// если текущее значение равно 0 то все ок

    if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
    if Item.ImageIndex=6 then lifehdd:=lifehdd+10;

    end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
   2: begin
       if(strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) or (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))
        then item.ImageIndex:=3 else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
     end;
   4:
     begin
       if(strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) or (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))
        then item.ImageIndex:=3 else item.ImageIndex:=2;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
     end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
  3:
  begin
   //if(strtoint(item.SubItems[2])>=strtoint(item.SubItems[3])) or (strtoint(item.SubItems[1])>=strtoint(item.SubItems[3]))
   //     then item.ImageIndex:=3 else item.ImageIndex:=2;
   item.ImageIndex:=2;
  end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
   5:
     begin
     if (Item.SubItems[4]='0') then begin item.ImageIndex:=2; end;
     if item.SubItems[3]<>'' then
        begin
        if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
        (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
             else
             if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])+15)) or
                 (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3])+15)) then Item.ImageIndex:=6
              else
             if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])+10)) or
                 (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3])+10)) then Item.ImageIndex:=7
              else Item.ImageIndex:=2;
        end
          else Item.ImageIndex:=4; /// если значение  item.SubItems[3] =  пусто то надо бы обновит
      if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+20;
      end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
   7:
      begin
      if (Item.SubItems[4]='0')or (Item.SubItems[4]='') then begin item.ImageIndex:=2; end;
      if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
          else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   8:      /// содержит среднюю производительность операций позиционирования головок
   begin
      if (Item.SubItems[4]='0')or (Item.SubItems[4]='') then begin item.ImageIndex:=2; end;
      if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
          else item.ImageIndex:=2;
      if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
   9:
      begin ///// количество отработанных часов
        Item.ImageIndex:=2;
        if strtoint(Item.SubItems[1])<=35 then item.ImageIndex:=7;
        if strtoint(Item.SubItems[1])<=25 then  item.ImageIndex:=6;
        if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3])+10)or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])+10) then
         item.ImageIndex:=3;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+8;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+12;
     end;
///////////////////////////////////////////////////////////////////////////////////////    
    10:  ///Количество повторных попыток запуска вращения.
        begin
     if item.SubItems[3]<>'' then
        begin
        if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
        (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
             else
             if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) or
                 (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then Item.ImageIndex:=6
              else
             if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])+10)) or
                 (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3])+10)) then Item.ImageIndex:=7
              else Item.ImageIndex:=2;
        end;
       if (Item.SubItems[4]='0') then item.ImageIndex:=2;

       if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
       if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
       if Item.ImageIndex=7 then lifehdd:=lifehdd+20;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////
    11,12:  //11-Повторная калибровка или калибровка// 12-  включения / выключения жесткого диска
      begin
       if (Item.SubItems[4]='0')or (Item.SubItems[4]='') then begin item.ImageIndex:=2; end;
       if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
          else item.ImageIndex:=2;
       if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
      end;
////////////////////////////////////////////////////////////////////////////////////////////////////
    13,100: if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=7
          else item.ImageIndex:=2;
///////////////////////////////////////////////////////////////////////////////////////////////////////
    148..151,160,161,163..168:
      begin
      if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) or
      (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then Item.ImageIndex:=6
      else Item.ImageIndex:=2;
      end;
/////////////////////////////////////////////////////////////////////////////////////////////////
    169:   ///Remain Life Percentage SSD
      begin
        try
        if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) or
      (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then Item.ImageIndex:=6
      else Item.ImageIndex:=2;

         except
         Item.ImageIndex:=4;
        end;
      if (Item.ImageIndex=6)or(Item.ImageIndex=4) then lifehdd:=lifehdd+10;
      end;
///////////////////////////////////////////////////////////////////////////////////////////////////
    170:      ////параметр аналогичный ID 8
         begin
         if item.SubItems[3]<>'' then
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
            (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=3
                 else
                 if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) or
                     (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then Item.ImageIndex:=6
                     else Item.ImageIndex:=2;
            end
              else
              Item.ImageIndex:=4; /// если значение  item.SubItems[3] =  пусто то надо бы обновит
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          if (Item.ImageIndex=6)or(Item.ImageIndex=4) then lifehdd:=lifehdd+5;
          end;
 //////////////////////////////////////////////////////////////////////////////////////////////////
      171,172,173,174:
           begin
       if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=7
          else item.ImageIndex:=2;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
          end;
/////////////////////////////////////////////////////////////////////////////////////////////////
      175: item.ImageIndex:=4;
/////////////////////////////////////////////////////////////////////////////////////////////////
      176: if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=7
          else item.ImageIndex:=2;
////////////////////////////////////////////////////////////////////////////////////////////////
      177,179,180:
      begin
       if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=7
          else item.ImageIndex:=2;
         if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
       end;
///////////////////////////////////////////////////////////////////////////////////////////////
      181,182,183:
      begin
       if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
          (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then item.ImageIndex:=7
          else item.ImageIndex:=2;
         if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
      end;
//////////////////////////////////////////////////////////////////////////////////////////
      184:
        begin   //// критический параметр   End-to-End error / IOEDC
        if item.SubItems[3]='' then begin item.ImageIndex:=4; end;
         if item.SubItems[4]='0' then item.ImageIndex:=2  else
          begin
           if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
            (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
            item.ImageIndex:=3
            else
             if  strtoint(item.SubItems[4])>=1 then  item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
         if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
         if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
/////////////////////////////////////////////////////////////////////////////////////////
       185:
          begin
           if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
          (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
          item.ImageIndex:=6
          else item.ImageIndex:=2;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          end;
//////////////////////////////////////////////////////////////////////////////////////////
       187:
          begin
          if (item.SubItems[4]='0') then begin item.ImageIndex:=2; end;
          if (item.SubItems[4]='') then begin item.ImageIndex:=4; end;//// 4 - это знак вопроса
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
          (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
          item.ImageIndex:=3
          else item.ImageIndex:=2;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=4 then lifehdd:=lifehdd+10;
          end;
/////////////////////////////////////////////////////////////////////////////////////////
        188:
           begin
            if (item.SubItems[4]='0') then begin item.ImageIndex:=2; end;
            if (item.SubItems[4]='') then begin item.ImageIndex:=4; end;//// 4 - это знак вопроса
            if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
                item.ImageIndex:=6
            else item.ImageIndex:=2;
             if (strtoint(item.SubItems[1])<=(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=(strtoint(item.SubItems[3]))) then
                item.ImageIndex:=3;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+20;
           end;
////////////////////////////////////////////////////////////////////////////////////////////////
        189:
          begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
        190:
          begin
          item.ImageIndex:=2;
          if item.SubItems[4]<>'' then
            begin
            if (strtoint(item.SubItems[4])>=45) then item.ImageIndex:=7
            else if (strtoint(item.SubItems[4])>=50) then item.ImageIndex:=6
            else if (strtoint(item.SubItems[4])>=60) then item.ImageIndex:=3
            else item.ImageIndex:=2;
            end
          else item.ImageIndex:=4; //' C°'
          if Item.ImageIndex=7 then Item.SubItems[5]:='HDD работает при повышной температуре';
          if Item.ImageIndex=6 then Item.SubItems[5]:='HDD работает при высокой температуре';
          if Item.ImageIndex=3 then Item.SubItems[5]:='HDD работает при критической температуре';
          if Item.ImageIndex=2 then Item.SubItems[5]:='OK';
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
        191:
          begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
               item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
        192:
          begin
           if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then
                item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
        193:
          begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then
               item.ImageIndex:=6
            else item.ImageIndex:=2;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        194:
          begin
          item.ImageIndex:=2;
          if item.SubItems[4]<>'' then
            begin
            if (strtoint(item.SubItems[4])>=45) then item.ImageIndex:=7
            else if (strtoint(item.SubItems[4])>=50) then item.ImageIndex:=6
            else if (strtoint(item.SubItems[4])>=60) then item.ImageIndex:=3
            else item.ImageIndex:=2;
            end
          else item.ImageIndex:=4;
          if Item.ImageIndex=7 then Item.SubItems[5]:='HDD работает при повышной температуре';
          if Item.ImageIndex=6 then Item.SubItems[5]:='HDD работает при высокой температуре';
          if Item.ImageIndex=3 then Item.SubItems[5]:='HDD работает при критической температуре';
          if Item.ImageIndex=2 then Item.SubItems[5]:='OK';
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        195:
          begin
            if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) then
               item.ImageIndex:=6;
             if (strtoint(item.SubItems[1])<=(strtoint(item.SubItems[3]))) then
               item.ImageIndex:=3
            else item.ImageIndex:=2;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        196:
          begin
          if (item.SubItems[4]='0') then begin item.ImageIndex:=2; end;
          if (item.SubItems[4]='') then  item.ImageIndex:=4 /// 4 - это знак вопроса
           else
           begin
           if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) then
                item.ImageIndex:=6;
          if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3])) then
                item.ImageIndex:=3
             else item.ImageIndex:=2;
           end;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
         197:
            begin
            item.ImageIndex:=2;
            if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) then
                 item.ImageIndex:=3
                else
              if (item.SubItems[4]<>'') then
               begin
               if strtoint(item.SubItems[4])<=9 then  item.ImageIndex:=2;
               if strtoint(item.SubItems[4])>=10 then  item.ImageIndex:=7;
               if strtoint(item.SubItems[4])>=50 then  item.ImageIndex:=6;
               if strtoint(item.SubItems[4])>=100 then  item.ImageIndex:=3;
               end
              else  item.ImageIndex:=4;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
          198:
             begin
             item.ImageIndex:=2;
            if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3]))) then
                 item.ImageIndex:=3
                else
              if (item.SubItems[4]<>'') then
               begin
               if strtoint(item.SubItems[4])<=9 then  item.ImageIndex:=2;
               if strtoint(item.SubItems[4])>=10 then  item.ImageIndex:=7;
               if strtoint(item.SubItems[4])>=50 then  item.ImageIndex:=6;
               if strtoint(item.SubItems[4])>=100 then  item.ImageIndex:=3;
               end
              else  item.ImageIndex:=4;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         199:
         begin
          if (item.SubItems[4]='') then item.ImageIndex:=4 /// 4 - это знак вопроса
           else
             begin
               if (item.SubItems[4]='0') then item.ImageIndex:=2;
               if strtoint(item.SubItems[4])<=19 then  item.ImageIndex:=2;
               if strtoint(item.SubItems[4])>=20 then  item.ImageIndex:=7;
               if strtoint(item.SubItems[4])>=50 then  item.ImageIndex:=6;
               end;
              if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
                 (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then
                item.ImageIndex:=3;
            if Item.ImageIndex=7 then Item.SubItems[5]:='Проверьте интерфейсный кабель';
            if Item.ImageIndex=6 then Item.SubItems[5]:='Проверьте интерфейсный кабель'; /// заменить кабель
            if Item.ImageIndex=3 then Item.SubItems[5]:='Замените интерфейсный кабель'; /// заменить кабель
            if Item.ImageIndex=2 then Item.SubItems[5]:='100 %';
         end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
         200: /// содержит частоту возникновения ошибок при записи
            begin
          if (item.SubItems[4]='') then item.ImageIndex:=4 /// 4 - это знак вопроса
           else
             begin
              if (item.SubItems[4]='0') then item.ImageIndex:=2;
              end;
             if (item.SubItems[4]<>'') then
               begin
               if strtoint(item.SubItems[4])<=5 then  item.ImageIndex:=2;
               if strtoint(item.SubItems[4])>=6 then  item.ImageIndex:=7;
               if strtoint(item.SubItems[4])>=30 then  item.ImageIndex:=6
               else  item.ImageIndex:=2;
               end;
              if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
                 (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then
                item.ImageIndex:=3;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+10; //
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15; ///
         end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
         201:  //содержит частоту возникновения ошибок чтения, произошедших по вине программного обеспечения
            begin
            begin
             item.ImageIndex:=2;
            if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then
                 item.ImageIndex:=3
                else
              if (item.SubItems[4]<>'') then
               begin
               if strtoint(item.SubItems[4])<=9 then  item.ImageIndex:=2;
               if strtoint(item.SubItems[4])>=10 then  item.ImageIndex:=7;
               if strtoint(item.SubItems[4])>=50 then  item.ImageIndex:=6;
               if strtoint(item.SubItems[4])>=100 then  item.ImageIndex:=3;
               end
              else  item.ImageIndex:=4;
            if Item.ImageIndex=7 then lifehdd:=lifehdd+30;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+40;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
            end;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////
         202: //  содержание атрибута — загадка, но проанализировав различные диски, могу констатировать, что ненулевое значение — это плохо
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
             if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
         203:   //содержит количество ошибок ECC
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////
         204:  //Количество ошибок, исправленных с помощью программного обеспечения для внутреннего исправления ошибок.
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
         205: /// Подсчет ошибок из-за высокой температуры.
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
          206: //Высота головок над поверхностью диска. Если оно слишком низкое, вероятность падения головы выше; если слишком высокий, ошибки чтения / записи более вероятны.
            begin
             if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
          207:  // Величина импульсного тока, используемого для вращения привода
            begin
            if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
          208:  // Количество циклов, необходимых для раскрутки
                //Количество повторных попыток во время ускорения из-за низкого доступного тока.
            begin
             if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10; /// проверить блок питания
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
          209:  /// Производительность диска во время автономных операций
                  //Производительность поиска накопителя во время внутренних самопроверок.
            begin
             if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////
          210:  /// вибрации во время записи
            begin
              if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
               if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////
       211: //Вибрация, возникающая во время операций записи.
        begin
          if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
       212: //Шок, возникший во время операций записи.
        begin
         if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        218:  // CRC Error Count (kingston SSD)
          begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////
       220:  //Расстояние диска сместилось относительно шпинделя (обычно из-за удара или температуры).
        begin
         if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+15;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       221: //Параметр SMART Disk Shift указывает, что расстояние диска смещено относительно шпинделя,
            // что может быть вызвано механическим ударом или высокой температурой.
        begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
            if Item.ImageIndex=3 then lifehdd:=lifehdd+10;
        end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
      222: ///  Количество отработанных часов .   Жесткие диски, поддерживающие этот атрибут
            //Samsung, Seagate, IBM (Hitachi), Fujitsu (не все модели), Maxtor, Western Digital (не все модели)
         begin
         Item.ImageIndex:=2;
          if strtoint(Item.SubItems[1])<=35 then item.ImageIndex:=7;
          if strtoint(Item.SubItems[1])<=25 then  item.ImageIndex:=6;
          if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3])+10)or
            (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])+10) then
           item.ImageIndex:=3;
           if Item.ImageIndex=7 then lifehdd:=lifehdd+5;
           if Item.ImageIndex=6 then lifehdd:=lifehdd+8;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+12;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
      223:  //Параметр SMART Load / Unload Retry Count указывает количество головок дисков, которые входят / выходят из зоны
        begin
        if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
       224: //Параметр SMART Load Friction указывает скорость трения между механическими частями жесткого диска.
            // Увеличение значения означает, что существует проблема с механической подсистемой привода.
        begin
         if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+15;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       225: ///Load / Unload Cycle Count SMART параметр указывает общее количество циклов загрузки.
        begin
         if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////
       226:  ///Общее время нагрузки на привод магнитных головок (время, не проведенное в зоне парковки)
        begin
          if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
       227: // указывает на количество попыток скомпенсировать изменения скорости вращения диска.
        begin
        if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////
       228://Число циклов выключения, которые подсчитываются, когда происходит «событие втягивания» и головки загружаются с носителя, например, когда машина выключена, переведена в спящий режим или находится в режиме ожидания.
        begin
         if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
            if Item.ImageIndex=6 then lifehdd:=lifehdd+5;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
       230:   // Head Amplitude GMR указывает расстояние перемещения головки между операциями   samsung, Seagate, IBM (Hitachi), Fujitsu (не все модели), Maxtor, Western Digital (не все модели)
        begin  /// для ssd критический параметр с worst=90  наилучший =100
        if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3]))or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])) then  item.ImageIndex:=6
               else item.ImageIndex:=2;
         if Item.ImageIndex=6 then lifehdd:=lifehdd+10;
        end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        231:  //Указывает приблизительный оставшийся срок службы SSD
         item.ImageIndex:=4;
          {begin
          if (strtoint(item.SubItems[4]))<=30 then item.ImageIndex:=7;
          if (strtoint(item.SubItems[4]))<=20 then item.ImageIndex:=6;
          if (strtoint(item.SubItems[4]))<=10 then item.ImageIndex:=3
           else  item.ImageIndex:=2;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+10;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
          end;}
//////////////////////////////////////////////////////////////////////////////////////////////////////////
        232:  //Число циклов физического стирания, выполненных на твердотельном накопителе, в процентах от максимального количества циклов физического стирания, которое рассчитан накопитель.
          begin
          if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3])+5)or
               (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])+5) then  item.ImageIndex:=3
               else item.ImageIndex:=2;
           if Item.ImageIndex=3 then lifehdd:=lifehdd+30;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
         233: //Индикатор износа носителя (SSD) или часы включения    Lifetime Nand Writes
           begin
             if item.SubItems[3]='' then
                begin
                if (strtoint(item.SubItems[1])<=35) or (strtoint(item.SubItems[2])<=35) then
                 item.ImageIndex:=6;
                if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
                 item.ImageIndex:=7
                   else item.ImageIndex:=2;
                end
             else
                begin
                   if (strtoint(item.SubItems[1])<=strtoint(item.SubItems[3])+5)or
                   (strtoint(item.SubItems[2])<=strtoint(item.SubItems[3])+5) then  item.ImageIndex:=3
                    else item.ImageIndex:=2;
                end;
          if Item.ImageIndex=7 then lifehdd:=lifehdd+10;
          if Item.ImageIndex=6 then lifehdd:=lifehdd+30;
          if Item.ImageIndex=3 then lifehdd:=lifehdd+50;
           end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
         234:
          begin  //Среднее количество стираемых И максимальное количество стираемых
           if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////////
         235:// Good Block Count AND System(Free) Block Count
          begin
            if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////
        240:  //Время, затрачиваемое на позиционирование приводных головок
          begin
            if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////
         241: //Всего написано LBA
           begin
           if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
           end;
/////////////////////////////////////////////////////////////////////////////////////////////////
         242: // Всего прочитанных LBA
          begin
           if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
         243:  //Всего написано LBA Расширено
          begin
           if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
///////////////////////////////////////////////////////////////////////////////////////////////////////
         244: // Всего прочитанных LBA Расширено  or Average Erase Count (Kingston SSD)
          begin
          if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
          end;
/////////////////////////////////////////////////////////////////////////////////////////////
         245: ///Max Erase Count (kingston)
          begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
////////////////////////////////////////////////////////////////////////////////////////////
          246: ///Total Erase Count (kingston)
          begin
          if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
               (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then item.ImageIndex:=6
            else item.ImageIndex:=2;
          end;
///////////////////////////////////////////////////////////////////////////////////////
          249: //Необработанное значение сообщает о количестве записей в NAND с шагом 1 ГБ
           begin
           if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
             item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+10;

           end;
/////////////////////////////////////////////////////////////////////////////////////////////////
          250:  ///Количество ошибок при чтении с диска.
            begin
            if item.SubItems[3]<>'' then
               begin
              if (strtoint(item.SubItems[1])<=ocenka(strtoint(item.SubItems[3])))or
                 (strtoint(item.SubItems[2])<=ocenka(strtoint(item.SubItems[3]))) then  item.ImageIndex:=3
                 else item.ImageIndex:=2;
               end
             else
              begin
                if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
              end;
            if Item.ImageIndex=3 then  lifehdd:=lifehdd+20;
            if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
/////////////////////////////////////////////////////////////////////////////////////////////////////
          251:  //Атрибут Minimum Spares Remaining указывает количество оставшихся запасных блоков в процентах от общего количества доступных запасных блоков.
            begin
              if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////
           252: //Недавно добавленная блокировка сбойной флэш-памяти указывает общее количество сбойных флэш-блоков, обнаруженных накопителем с момента его первоначальной инициализации в процессе производства.
            begin
               if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
               if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
          254: //Количество обнаруженных событий «Свободное падение»
            begin
            if (strtoint(item.SubItems[1])<=10) or (strtoint(item.SubItems[2])<=10) then
                item.ImageIndex:=6
               else item.ImageIndex:=2;
             if Item.ImageIndex=6 then  lifehdd:=lifehdd+15;
            end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////
          else item.ImageIndex:=5;
  end;
   except
  item.ImageIndex:=4;
  end;

//////////////////// столбец ресурс
if not exceptAtribute(strtoint(item.Caption)) then
 begin
  case item.ImageIndex of
   7: item.SubItems[5]:='~75 %';
   6: item.SubItems[5]:='~25-50 %';
   3: item.SubItems[5]:='~0-25 %';
   2: item.SubItems[5]:='100 %';
  end;
 end;
/////////////////////////////////
  AllLifehdd:=AllLifehdd+lifehdd;
end;


procedure TForm8.N1Click(Sender: TObject);
begin
  EnableDisableHardwareFailurePrediction(True);
end;

procedure TForm8.N0Offlinetestroutin1Click(Sender: TObject);
begin
starttest('0','Off-line test routin');
end;

procedure TForm8.N1Shortselftestroutin1Click(Sender: TObject);
begin
starttest('1', 'Short self-test routin');
end;

procedure TForm8.N2Click(Sender: TObject);
begin
  EnableDisableHardwareFailurePrediction(False);
end;

procedure TForm8.N2Extendedselftestroutin1Click(Sender: TObject);
begin
starttest('2','Extended self-test routin');
end;

procedure TForm8.N3Click(Sender: TObject);
begin
EnableDisableFailurePredictionPolling(True);
end;

procedure TForm8.N4Click(Sender: TObject);
begin
EnableDisableFailurePredictionPolling(False);
end;

procedure TForm8.N5Click(Sender: TObject);
begin
  AllowPerformanceHit(True);
end;

procedure TForm8.N6Click(Sender: TObject);
begin
  AllowPerformanceHit(false);
end;

procedure TForm8.SpeedButton11Click(Sender: TObject);
begin
EnableOfflineDiags;
end;




procedure TForm8.RefreshDisk;
var
/// /список дисков
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum: IEnumvariant;
  FWbemObject: OLEVariant;
  iValue: LongWord;

  // CapabilityDes   :string;
  // masvar          :variant;
 // id, i, idsmart: Integer;
begin
  GroupDisk.Items.Clear; // очистить список дисков

  try
   // id := 0;
   // idsmart := 0;

    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    /// ////////////////////////////////////////////////////////////////////
    /// список дисков
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\CIMV2', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    //FWbemObjectSet := FWMIService.ExecQuery('SELECT index,status,caption,PNPDeviceID,MediaType FROM Win32_DiskDrive WHERE MediaType=''Fixed hard disk media''',
     FWbemObjectSet := FWMIService.ExecQuery('SELECT index,status,caption,PNPDeviceID,MediaType FROM Win32_DiskDrive', 'WQL', wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
          begin
        with GroupDisk.Items.Add do
        begin
          Caption := 'Диск - ' + IntToStr(FWbemObject.Index) + '   ' +
            vartostr(FWbemObject.Caption)+' /'+vartostr(FWbemObject.status)+'/';
            if vartostr(FWbemObject.status)<>'OK' then ImageIndex := 1
            else ImageIndex := 0;
          Hint := vartostr(FWbemObject.PNPDeviceID);
        end;
      end;
    end;
    FWbemObject := Unassigned;
  except
    on E: Exception do
    begin
      Memo1.Lines.Add('Ошибка получения данных о дисках- "' + E.Message + '"');
      exit;
    end;
  end;
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  VariantClear(FWbemObject);
end;


procedure TForm8.PhysicalDisk;
var   ///// данные о диске начиная с Windows 2012
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  oEnum: IEnumvariant;
  FWbemObject: OLEVariant;
  iValue: LongWord;
begin
  try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'ROOT\Microsoft\Windows\Storage', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery('SELECT * FROM MSFT_PhysicalDisk',
      'WQL', wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
     memo2.Lines.Add('Description-'+FWbemObject.Description);
     memo2.Lines.Add('FriendlyName-'+FWbemObject.FriendlyName);
     memo2.Lines.Add('DeviceId-'+FWbemObject.DeviceId);
     //memo2.Lines.Add('OperationalDetails-'+FWbemObject.OperationalDetails); //not supportes
     memo2.Lines.Add('PhysicalLocation-'+FWbemObject.PhysicalLocation);

    end;
    FWbemObject := Unassigned;
  except
    on E: Exception do
    begin
      Memo1.Lines.Add('ROOT\Microsoft\Windows\Storage - "' + E.Message + '"');
      exit;
    end;
  end;
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  VariantClear(FWbemObject);
end;

procedure TForm8.OnMyMenuItemClick(Sender: TObject);
begin
//memo2.Lines.add(PopupTest.Items.Caption);
end;


procedure TForm8.SmartSelect;
var // SMART Выбранного диска
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet,FWbemObjectSmart,FWbemObjectFPS: OLEVariant;
  FWbemObjectVendorTresh,FWbemObject1: OLEVariant;
  oEnum,oEnumSM,oEnumVT,oEnumFP: IEnumvariant;
  FWbemObject,FWbemObjectSM,FWbemObjectVT: OLEVariant;
  iValue: LongWord;
  // SmartMAS :variant;
  SmartMAS: array of byte;
  ThresholdsArr: array of byte;
  Vendorspec4: array of byte;
  i,z,atribute190,atribute194,Atribute9: Integer;
  predID, SelfStat, selfStatProcent: Integer; //  Atribute241 - написано LBA/ Atribute242- прочитано LBA/Atribute243- дополнительно к написано LBA
  flags, testImpl, SrtCapab,vendorspecific4,SpinUpTime,Atribute234,Atribut235,
  Atribute240,Atribute169,Atribute175: string;
  AllTime:TTimeSpan;

//////////////////////////////////////////////////////////////////////
  function bubblesort(s: string):bool; /// функция сортировки и расчета миним и максим температуры
  var
  A: array[0..3] of integer;
  B: array[0..3] of integer;
  N,i,j,p,z : integer;
  begin
  for z := 0 to 3 do
    begin
      a[z]:=strtoint(copy(s,1,pos(',',s)-1));
      delete(s,1,pos(',',s));
    end;
  for I :=1 to 3 do
  begin
    for j := 1 to 3-i do
    begin
     if A[j]>A[j+1] then
       begin
       p:=A[j];
       A[j]:=A[j+1];
       A[j+1]:=P;
       end;
    end;
  end;
  s:='';
  p:=0;
   for z := 0 to 3 do
   begin
   if (a[z]>0)and(a[z]<80) then
     begin
     s:=s+inttostr(a[z])+',';
     inc(p);
     end;
   end;
   if p>0 then
     begin
       case p of
       1:
       begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          StaticText16.Caption:='Минимальная - '+copy(s,1,pos(',',s)-1)+' °C'
           else StaticText16.Caption:='Максимальная - '+copy(s,1,pos(',',s)-1)+' °C'
       end;
       2:
       begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          StaticText16.Caption:='MIN - '+copy(s,1,pos(',',s)-1)+' °C / ';
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          StaticText16.Caption:=StaticText16.Caption+'MAX - '+copy(s,1,pos(',',s)-1)+' °C';
       end;
        3:
         begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          StaticText16.Caption:='MIN - '+copy(s,1,pos(',',s)-1)+' °C';
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
          delete(s,1,pos(',',s));//// удаляем вторую цифру и знак запятой т.к. нам надо MAX и MIN
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          StaticText16.Caption:=StaticText16.Caption+'/MAX - '+copy(s,1,pos(',',s)-1)+' °C';
         end;
        4:
         begin
          if (strtoint(copy(s,1,pos(',',s)-1))<Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))<Atribute194) then
          StaticText16.Caption:='MIN - '+copy(s,1,pos(',',s)-1)+' °C';
          delete(s,1,pos(',',s));//// удаляем первую цифру и знак запятой
          delete(s,1,pos(',',s));//// удаляем вторую цифру и знак запятой т.к. нам надо MAX и MIN
          delete(s,1,pos(',',s));//// удаляем третью цифру и знак запятой т.к. нам надо MAX и MIN
           if (strtoint(copy(s,1,pos(',',s)-1))>Atribute190)or(strtoint(copy(s,1,pos(',',s)-1))>Atribute194) then
          StaticText16.Caption:=StaticText16.Caption+'/MAX - '+copy(s,1,pos(',',s)-1)+' °C';
         end;
       end;
     end;
  end;
///////////////////////////////////////////////////////////////////////////////////
  function delflags(s: string): string;
  begin
    if s <> '' then
    begin
      delete(s, Length(s) - 2, Length(s));
      Result := s;
    end
    else
      Result := '';
  end;

begin
  try
    fillchar(SmartMAS,sizeof(SmartMAS),0);
    fillchar(ThresholdsArr,sizeof(ThresholdsArr),0);
    fillchar(Vendorspec4,sizeof(Vendorspec4),0);
    flags := '';
    ListSmart.Clear;
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
  ////////////////////////////////////////////////////////////////////////////////////////
  Typesmart := 0;
  FWbemObjectSmart := FWMIService.ExecQuery
    ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
    wbemFlagForwardOnly);
  oEnumSM := IUnknown(FWbemObjectSmart._NewEnum) as IEnumvariant;
  while oEnumSM.Next(1, FWbemObjectSM, iValue) = 0 do
  begin
      FWbemObjectSM.GetFailurePredictionCapability(Typesmart);
      case Typesmart of
        0:          Memo2.Lines.Add('S.M.A.R.T -  Not Supported');
        1:          Memo2.Lines.Add('S.M.A.R.T -  ioctl Based');
        2:          Memo2.Lines.Add('S.M.A.R.T -  IDE SMART');
        3:          Memo2.Lines.Add('S.M.A.R.T -  SCSI SMART');
      end;
    FWbemObjectSM := Unassigned;
  end;
    FWbemObjectSmart:= Unassigned;
    oEnumSM:=nil;
    VariantClear(FWbemObjectSM);
    VariantClear(FWbemObjectSmart);

    /// /////////////////////////////////////////////////////////////////////////////
    FWbemObjectVendorTresh := FWMIService.ExecQuery
    ///// заполняем массив с пороговыми значениями
      ('SELECT * FROM MSStorageDriver_FailurePredictThresholds WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',wbemFlagForwardOnly);
    oEnumVT := IUnknown(FWbemObjectVendorTresh._NewEnum) as IEnumvariant;
    while oEnumVT.Next(1, FWbemObjectVT, iValue) = 0 do
    begin
        SetLength(ThresholdsArr,VarArrayHighBound(FWbemObjectVT.properties_.Item('VendorSpecific',0).value, 1)); /// задаем длинну массива
        for i := 0 to VarArrayHighBound(FWbemObjectVT.properties_.Item('VendorSpecific', 0).value, 1) do
        begin
          ThresholdsArr[i] := FWbemObjectVT.properties_.Item('VendorSpecific',0).value[i];
        end;
      FWbemObjectVT := Unassigned;
    end;
    oEnumVT := nil;
    FWbemObjectVendorTresh := Unassigned;
    FWbemObjectVT:= Unassigned;
    VariantClear(FWbemObjectVT);
    VariantClear(FWbemObjectVendorTresh);

    /// //////////////////////////////////////////////////////////////
    FWbemObjectFPS := FWMIService.ExecQuery
    ///// SMART статус из  MSStorageDriver_FailurePredictStatus
      ('SELECT * FROM MSStorageDriver_FailurePredictStatus WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnumFP := IUnknown(FWbemObjectFPS._NewEnum) as IEnumvariant;
    while oEnumFP.Next(1,FWbemObject1, iValue) = 0 do
    begin
        if FWbemObject1.PredictFailure then
          GroupBox2.Caption := GroupBox2.Caption+' - S.M.A.R.T - WARNING (Оценка операционной системы).'
        else
          GroupBox2.Caption :=GroupBox2.Caption+ ' - S.M.A.R.T - OK (Оценка операционной системы).';
    FWbemObject1 := Unassigned;
    end;
    oEnumFP := nil;
    FWbemObject1:= Unassigned;
    FWbemObjectFPS := Unassigned;
    VariantClear(FWbemObjectFPS);
    VariantClear(FWbemObject1);
    /// ////////////////////////////////////////////////////////////////

    /// /// непосредственно данные SMART
    FWbemObjectSet := Unassigned;
    VariantClear(FWbemObjectSet);
    oEnum:=nil;
    FWbemObjectSet := FWMIService.ExecQuery ('SELECT * FROM MSStorageDriver_ATAPISmartData WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL', wbemFlagForwardOnly);   // MSStorageDriver_ATAPISmartData или MSStorageDriver_FailurePredictData
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
     // memo2.Lines.Add('selectHDD - '+selectHDD);
    //  Memo2.Lines.Add('Объект - '+vartostr(FWbemObject.InstanceName));
        if copy(Dectobin(FWbemObject.ErrorLogCapability, 8), 8, 1) = '1' then
          Memo2.Lines.Add
            ('Устройство поддерживает регистрацию ошибок (S.M.A.R.T)' ) ///+Dectobin(FWbemObject.ErrorLogCapability, 8)
        else
          begin
          Memo2.Lines.Add('Устройство не поддерживает регистрацию ошибок (S.M.A.R.T).  ');// + Dectobin(FWbemObject.ErrorLogCapability, 8));
          Memo2.Lines.Add('УПС....');
          oEnum := nil;
          VariantClear(FWbemObject);
          VariantClear(FWbemObjectSet);
          VariantClear(FWMIService);
          VariantClear(FSWbemLocator);
          exit;
          end;
        /////////////////////////////////////////////////////////////////////////////////////////////
        Memo2.Lines.Add('Время расширенной процедуры самопроверки (Extended self-test routin) - ' + vartostr(FWbemObject.ExtendedPollTimeInMinutes) + ' мин.');
        Memo2.Lines.Add('Время короткой процедуры самопроверки (Short self-test routin) - ' + vartostr(FWbemObject.ShortPollTimeInMinutes) + ' мин.');
        Memo2.Lines.Add('Время до завершения автономного сбора данных - ' +IntToStr(FWbemObject.totalTime) + ' сек.');
        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////
        SrtCapab := copy(Dectobin(FWbemObject.SmartCapability, 16), 16, 1);
        if SrtCapab = '1' then
          Memo2.Lines.Add ('Устройство сохраняет SMART данные до перехода в режим ожидания.')
        else
          Memo2.Lines.Add ('Устройство не сохраняет SMART данные при переходе в режим ожидания.');
        SrtCapab := copy(Dectobin(FWbemObject.SmartCapability, 16), 15, 1);
        if SrtCapab = '1' then
          Memo2.Lines.Add ('Устройство поддерживает атрибут SMART ENABLE/DISABLE')
        else
          Memo2.Lines.Add ('Устройство не поддерживает атрибут SMART ENABLE/DISABLE');
        /// /////////////////////////////////////////////////////////////////////////////////////////////////
        SelfStat := BIN2DEC(copy(Dectobin(FWbemObject.SelfTestStatus,8), 1, 4));
        /// 4,,7 бит это статус
        Memo2.Lines.Add('Состояние выполнения самотестирования - ' + SelfTestStat(SelfStat));
        selfStatProcent := BIN2DEC(copy((Dectobin(FWbemObject.SelfTestStatus, 8)), 5, 8));
        /// 0,,3 бит это процент выполнения теста,
        case selfStatProcent of
          0:     Memo2.Lines.Add('Процедура самотестирования завершена.');
          1 .. 9:Memo2.Lines.Add('Процесс тестирования - ' + IntToStr(100 - (selfStatProcent * 10)) + '%');
        end;
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////
        Memo2.Lines.Add('Статус сбора данных в автономном режиме - ' + OffCollStatus(FWbemObject.OfflineCollectionStatus));
        /// /////////////////////////////////////////////////////////////////////////////
        if FWbemObject.OfflineCollectCapability <> 0 then
        begin
          testImpl := Dectobin(FWbemObject.OfflineCollectCapability, 8);
         // Memo2.Lines.Add('OfflineCollectCapability - ' +Dectobin(FWbemObject.OfflineCollectCapability, 8));
          /// /8ми битовое значение не более 255
          if copy(testImpl, 8, 1) = '0' then
          /// 0й бит, 1й бит Vendor Specific  / если не равно 1 то жесткий не поддерживает команды самотестирования
          begin
            Memo2.Lines.Add ('Жесткий диск не поддерживает команды для самотетирования (S.M.A.R.T)');
          end;
          if copy(testImpl, 6, 1) <> '0' then
          /// 2й бит, 7
          begin
            Memo2.Lines.Add ('Устройство поддерживает остановку и возобновление самотестирования (ABORT/RESTART OFF-LINE BY HOST)');
          end;
          if copy(testImpl, 5, 1) <> '0' then
          /// 3й бит, 7
          begin
            Memo2.Lines.Add ('Поддерживает процедуру Автономного самотестирования ( Off-line test routin)');
          end;
          if copy(testImpl, 4, 1) <> '0' then
          /// 4й бит,
          begin
            Memo2.Lines.Add
              ('Поддерживает Расширенные процедуры самотестирования (Short and Extended SELF-TEST)');
           // ComboBox1.Items.Add('127 - Abort off-line mode self-test routine');
             //ComboBox1.Items.Add('129 - Short self-test routin (captive mode)');
          //  ComboBox1.Items.Add('130 - Extended self-test routin (captive mode)');
             end;
          if copy(testImpl, 3, 1) <> '0' then  /// 5й бит,
          begin
            Memo2.Lines.Add ('Поддерживает процедуру самотестирования транспортировки (CONVEYANCE SELF-TEST)');
           // ComboBox1.Items.Add('3 - Conveyance self-test routin');
           // ComboBox1.Items.Add('130 - Conveyance self-test routin (captive mode)');
          end;
          if copy(testImpl, 2, 1) <> '0' then  /// 6й бит, 7й бит зарезервирован и не используется
          begin
            Memo2.Lines.Add ('Поддерживает выборочную процедуру самотестирования (SELECTIVE SELF-TEST)');
           // ComboBox1.Items.Add('4 - Selective self-test routin');
           // ComboBox1.Items.Add('132 - Selective self-test routin (captive mode)');
          end;
          {SetLength(Vendorspec4, VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific4',0).value, 1));
        for i := 0 to VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific4', 0).value, 1) do
        begin
          Vendorspec4[i] := FWbemObject.properties_.Item('VendorSpecific4',0).value[i]
        end;
        for I := 0 to length(Vendorspec4) do vendorspecific4:=vendorspecific4+','+inttostr(Vendorspec4[i]);
        Memo2.Lines.Add(vendorspecific4);
        memo2.Lines.Add('Длинна - '+inttostr(length(Vendorspec4)));}
         //// данные SMART
        // задаем длинну массива
        SetLength(SmartMAS, VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific',0).value, 1));
        for i := 0 to VarArrayHighBound(FWbemObject.properties_.Item('VendorSpecific', 0).value, 1) do
        begin
          SmartMAS[i] := FWbemObject.properties_.Item('VendorSpecific',0).value[i]
        end;
        /// /////////////////////////////////////////////////////////////////////////////////
          { memo1.Lines.Add('Off-line data collection status - '+inttostr(smartmas[362]));  /// значение может менятся в зависимости от состояния
            memo1.Lines.Add('Self-test execution status byte - '+inttostr(smartmas[363]));  //// значение зависит от поставщика
            memo1.Lines.Add('Off-line data collection capability - '+inttostr(smartmas[367]));  //// значение фиксированное
            memo1.Lines.Add('SMART capability - '+inttostr(smartmas[368])+'/'+inttostr(smartmas[369]));  //// значение фиксированное
            memo1.Lines.Add('Short self-test routine recommended polling time (минуты)- '+inttostr(smartmas[372]));  //// значение фиксированное
            memo1.Lines.Add('Extended self-test routine recommended polling time in minutes. If FFh, use  bytes 375 and 376 for the polling time (минуты)- '+inttostr(smartmas[373]));  //// значение фиксированное
            memo1.Lines.Add('Extended self-test routine recommended polling time (375 and 376 bytes) (минуты)- '+inttostr(smartmas[375])+'/'+inttostr(smartmas[376]));  //// значение фиксированное
            memo1.Lines.Add('Conveyance self-test routine recommended polling time (минуты)- '+inttostr(smartmas[374]));  //// значение фиксированное
          }
          predID := 0;
          /// / предыдущий ID необходим для того чтобы не выводились непонятные ID после последнего максимального
          for i := 0 to 29 do  // почему 29 , данные располагаются с 0 го по 361 байт
          begin
            if (SmartMAS[i * 12 + 2] <> 0) and (SmartMAS[i * 12 + 2] > predID) then /// если ID больше 0 и текущий ID больше предыдущего
            begin
              with ListSmart.Items.Add do
              begin
                predID := SmartMAS[i * 12 + 2];
                Caption := IntToStr(SmartMAS[i * 12 + 2]); // ID
                SubItems.Add(smartdscptn(SmartMAS[i * 12 + 2])); // описание параметра
                SubItems.Add(IntToStr(SmartMAS[i * 12 + 5]));   // 4 текущее value get 6th column where actual normalized data is stored
                SubItems.Add(IntToStr(SmartMAS[i * 12 + 6]));   // 5 наихудшее значение worst get 7th column where worst normalized data is stored

//////////////////////////////////////////////////////////////////////////////////// Threshold Пороговое занчение
                //if ThresholdsArr[i * 12 + 2] = SmartMAS[i * 12 + 2] then memo2.Lines.add(IntToStr(ThresholdsArr[i * 12 + 3]));
                if ThresholdsArr[i * 12 + 2] <> SmartMAS[i * 12 + 2] then
                  begin
                    for z := 0 to 29 do
                    begin
                     if ThresholdsArr[i * 12 + 2] = SmartMAS[z * 12 + 2] then
                      begin
                      SubItems.Add(IntToStr(ThresholdsArr[z * 12 + 3]));
                      break;
                      end
                      else if z=29 then SubItems.Add('0');
                      end;
                  end
                else  SubItems.Add(IntToStr(ThresholdsArr[i * 12 + 3]));
//////////////////////////////////////////////////////////////////////////////////////////////////////////


                case SmartMAS[i * 12 + 2] of
                // 16 в 8й степени          6й степени                 4й степени            2й степени
                  1,2,4 .. 8: SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  3:  begin
                      SpinUpTime:='';
                      if SmartMAS[i * 12 + 11]<>0 then SpinUpTime:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then SpinUpTime:=SpinUpTime+IntToStr(SmartMAS[i * 12 + 7]);
                      if SpinUpTime='' then SpinUpTime:='0';
                       SubItems.Add(SpinUpTime);
                      end;
                  9:begin //// отработаных часов
                    atribute9:=0;
                    atribute9:=(((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                    SubItems.Add(inttostr(Atribute9));
                    AllTime:=TtimeSpan.FromHours(atribute9);
                    StaticText18.Caption:='Отработано - '+
                    (Format('%d (день) %d (ч)',[Alltime.days,AllTime.Hours]));
                    end;
                  10 .. 13: SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  22: SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                 // 100: SubItems.Add(inttostr(BIN2DEC(copy((Dectobin(SmartMAS[i * 12 + 7],6)),4,3))));    ????
                  100:SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]* 65536) + (SmartMAS[i * 12 + 8] * 256) + (SmartMAS[i * 12 + 7])));
                  169:
                     begin     //// оставшийся процент жизни SSD
                     Atribute169:='';
                     if SmartMAS[i * 12 + 11]<>0 then Atribute169:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then Atribute169:=Atribute169+IntToStr(SmartMAS[i * 12 + 7]);
                      SubItems.Add(Atribute169);
                     end;
                  170 .. 174:SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  175 :///// 175 расчитывается по байтам 0-1 последний результат теста в микросекундах, 2-3: минуты с момента последнего теста, 4-5: количество тестов в течение жизни
                  begin
                   Atribute175:='';
                   if SmartMAS[i * 12 + 7]<>0 then  Atribute175:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                   if SmartMAS[i * 12 + 8]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                   if SmartMAS[i * 12 + 9]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                   if SmartMAS[i * 12 + 10]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                   if SmartMAS[i * 12 + 11]<>0 then  Atribute175:= Atribute175+IntToStr(SmartMAS[i * 12 + 7]);
                   SubItems.Add( Atribute175);
                  end;

                  176 .. 189:SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  191 .. 193:SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  195 ..233,236..239,244..254:SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                  190:begin
                      atribute190:=0;
                      SubItems.Add(IntToStr(SmartMAS[i * 12 + 7])); // температура
                      atribute190:=(SmartMAS[i * 12 + 7]);
                      StaticText9.Caption:='Текущая - '+IntToStr(atribute190)+' °C';
                      end;
                  194: begin
                        atribute194:=0;
                        atribute194:=(SmartMAS[i * 12 + 7]);
                        SubItems.Add(IntToStr(SmartMAS[i * 12 + 7])); // температура
                        bubblesort(inttostr(SmartMAS[i * 12 + 8])+','+(IntToStr(SmartMAS[i * 12 + 9]))+','+(IntToStr(SmartMAS[i * 12 + 10]))+','+(IntToStr(SmartMAS[i * 12 + 11]))+',');
                        StaticText9.Caption:='Текущая - '+IntToStr(atribute194)+' °C';
                       end;
                  234: begin //Decoded as: byte 0-1-2 = average erase count (big endian) and byte 3-4-5 = max erase count (big endian)
                       Atribute234:='';
                       Atribute234:=inttostr((SmartMAS[i * 12 + 7])+(SmartMAS[i * 12 + 8]*256)+(SmartMAS[i * 12 + 9]+65536))+'/';
                       Atribute234:=Atribute234+inttostr((SmartMAS[i * 12 + 10]+16777216)+(SmartMAS[i * 12 + 11]*4294967296)+(SmartMAS[i * 12 + 12]+1099511627776));
                       SubItems.Add(Atribute234);
                       end;
                  235: begin
                       Atribut235:='';
                       Atribut235:=inttostr((SmartMAS[i * 12 + 7])+(SmartMAS[i * 12 + 8]*256)+(SmartMAS[i * 12 + 9]+65536))+'/';
                       Atribut235:=Atribut235+inttostr((SmartMAS[i * 12 + 10]+16777216)+(SmartMAS[i * 12 + 11]*4294967296));
                       SubItems.Add(Atribut235);
                       end;
                  240:  begin
                      Atribute240:='';
                      if SmartMAS[i * 12 + 11]<>0 then Atribute240:=IntToStr(SmartMAS[i * 12 + 11]) +'/';
                      if SmartMAS[i * 12 + 10]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 10]) +'/';
                      if SmartMAS[i * 12 + 9]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 9]) +'/';
                      if SmartMAS[i * 12 + 8]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 8]) +'/';
                      if SmartMAS[i * 12 + 7]<>0 then Atribute240:=Atribute240+IntToStr(SmartMAS[i * 12 + 7]);
                       SubItems.Add(Atribute240);
                      end;
                   241:  /// написано LBA
                       begin
                       Atribute241:=0;
                       Atribute241:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       SubItems.Add(inttostr(Atribute241));
                       if sizeLBA<>0 then
                        begin
                       // Atribute241:=(((Atribute241) div 1024) div 1024)div 1024;
                         Atribute241:=(((Atribute241*SizeLBA) div 1024) div 1024)div 1024;
                        end;
                       StaticText20.Caption:='Всего записано - '+inttostr(Atribute241) +' Гб';
                       end;
                   242:  // Прочитано LBA
                       begin
                       Atribute242:=0;
                       Atribute242:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       SubItems.Add(inttostr(Atribute242));
                       if sizeLBA<>0 then
                        begin
                        //Atribute242:=(((Atribute242)div 1024)div 1024)div 1024;
                        Atribute242:=(((Atribute242*SizeLBA)div 1024)div 1024)div 1024;
                        end;
                       StaticText19.Caption:='Всего прочитано - '+inttostr(Atribute242)+' Гб';
                       end;
                   243:  // Дополнительно написано LBA
                       begin
                       Atribute243:=0;
                       Atribute243:=((SmartMAS[i * 12 + 11] * 4294967296) + (SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7]));
                       SubItems.Add(inttostr(Atribute243));
                       if sizeLBA<>0 then
                        begin
                         Atribute243:=(((Atribute243*SizeLBA) div 1024)div 1024)div 1024;
                        // Atribute243:=(((Atribute243) div 1024)div 1024)div 1024;
                         Atribute243:=Atribute243+Atribute241;
                         StaticText20.Caption:='Всего записано - '+inttostr(Atribute243)+' Гб';
                        end;
                       end;

                else
                  SubItems.Add(IntToStr((SmartMAS[i * 12 + 11] * 4294967296) +(SmartMAS[i * 12 + 10] * 16777216) + (SmartMAS[i * 12 + 9]*65536) + (SmartMAS[i * 12 + 8] * 256) +(SmartMAS[i * 12 + 7])));
                end;
                /// /////////////////////////////////////////////////////////
                SubItems.Add(''); /// столбец для оценки ресурса
               //////////////////////////////////////////////////////////////
                if SmartMAS[i * 12 + 3] mod 2 = 0 then
                  SubItems.Add('Pre-Failure')
                else
                  SubItems.Add('Advisory');   // flags
                flags := '';
                if (SmartMAS[i * 12 + 3] and 1) = 0 then
                  flags := 'PF - '; // Pre-Failure (PF, 01h) - при достижении порогового значения данного типа атрибутов диск требует замены
                if (SmartMAS[i * 12 + 3] and 2) = 0 then
                  flags := flags + 'OC - ';  // Online test (OC, 02h)– атрибут обновляет значение при выполнении off-line/on-line встроенных тестов SMART;
                if (SmartMAS[i * 12 + 3] and 4) = 0 then
                  flags := flags + 'PR - ';  //Perfomance Related (PE или PR , 04h)– атрибут характеризует производительность ;
                if (SmartMAS[i * 12 + 3] and 8) = 0 then
                  flags := flags + 'ER - ';  // Error Rate (ER , 08h )– атрибут отражает счетчики ошибок оборудования;
                if (SmartMAS[i * 12 + 3] and 16) = 0 then
                  flags := flags + 'EC - ';  ///Event Counts (EC, 10h ) – атрибут представляет собой счетчик событий;
                if (SmartMAS[i * 12 + 3] and 32) = 0 then
                  flags := flags + 'SP - ';  //Self Preserving (SP, 20h ) – самосохраняющися атрибут;
                flags := delflags(flags);
                SubItems.Add(flags); // flags
///////////////////////////////////////////////////////////////////////////////////
                ListSmart.Items[ListSmart.Items.Count - 1].ImageIndex :=
                  Viewimage;
                imageitems(ListSmart.Items[ListSmart.Items.Count - 1]);
 ////////////////////////////////////////////////////////////////////////////////////////
              end;               /// with listsmart.Items.Add do
            end;/// if (smartMas[i*12+2]<>0) and (smartmas[i*12+2]>predID) then
          end;  /// for I := 0 to 29 do
        end; ///  if FWbemObject.OfflineCollectCapability <> 0 then

      FWbemObject := Unassigned;
      SmartMAS := nil;
      ThresholdsArr := nil;  ///???
      Vendorspec4:=nil;     /// ???
      break;                /// ???
  end;
 ////////////////////////////////////////////////////////
  if Alllifehdd>=100 then   /////// вывод жизни диска
    begin
     ProgressBar1.Position:=5;
     Alllifehdd:=5;
     StaticText17.Caption:='Здоровье - 5 %';
     end
  else
    begin
    Alllifehdd:=100-Alllifehdd;
    ProgressBar1.Position:=Alllifehdd;
    StaticText17.Caption:='Здоровье - '+inttostr(Alllifehdd)+' %';
    end;

 ///////////////////////////////////////////////////////////////
  FWbemObject := Unassigned;
  VariantClear(FWbemObject);
  oEnum := nil;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('ошибка загрузки данных SMART - "' + E.Message + '"');
      oEnum := nil;
      VariantClear(FWbemObject);
      VariantClear(FWbemObjectSet);;
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      exit;
    end;
  end;
  oEnum := nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);;
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  SmartMAS := nil;
  ThresholdsArr := nil;
  Vendorspec4:=nil;
end;

function TForm8.starttest(s,b:string):string;
var
/// запуск офлайн теста
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: Integer;
const
  resmas: array [0 .. 2] of string = ('Successful Completion',
    'Captive Mode Required', 'Unsuccessful Completion');
begin
  try   ///7Fh (127) Abort off-line mode self-test routine
    if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE  InstanceName LIKE "%'+Obrez(selectHDD)+'%"','WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
        FWbemObject.ExecuteSelfTest(s,outparam);
        //Memo2.Lines.Add('Тест ' + (copy(ComboBox1.Text, 1, (pos('-',ComboBox1.Text))-2)
       // + ' Запустился на - '+vartostr(FWbemObject.InstanceName)));
        result:=resmas[outparam];
        Memo2.Lines.Add('Запуск теста '+b+' - '+result);
      FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TForm8.SpeedButton4Click(Sender: TObject);
begin
PhysicalDisk;
end;

procedure TForm8.SpeedButton1Click(Sender: TObject);
var
/// остановка офлайн теста
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: Integer;
const
  resmas: array [0 .. 2] of string = ('Successful Completion',
    'Captive Mode Required', 'Unsuccessful Completion');
begin
  try   ///7Fh (127) Abort off-line mode self-test routine
    if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE  InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// /
    begin
        FWbemObject.ExecuteSelfTest(127,outparam);
        Memo2.Lines.Add('Остановка тестирования на - '+ (FWbemObject.InstanceName));
        Memo2.Lines.Add(resmas[outparam]);
        FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TForm8.SpeedButton2Click(Sender: TObject);
var
///  ReadLogSectors method
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  inLogAddress,inSectorCount,outLength,outLogSectors:integer;
begin
  try
    inLogAddress:=20;
    inSectorCount:=11;
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE  InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// /
    begin
        FWbemObject.ReadLogSectors(inLogAddress,inSectorCount,outLength,outLogSectors);
        Memo2.Lines.Add('ReadLogSectors на - '+ (FWbemObject.InstanceName));
        Memo2.Lines.Add('outLength - '+ inttostr(outLength));
        Memo2.Lines.Add('outLogSectors - '+ inttostr(outLogSectors));
        FWbemObject := Unassigned;
     end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TForm8.EnableOfflineDiags;
var
/// включение автономной диагностики
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: bool;
begin
  try
    if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;
    Memo2.Lines.Add('Включение автономной дианостики S.M.A.R.T');
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE  InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
        FWbemObject.EnableOfflineDiags(outparam);
        if outparam then
          Memo2.Lines.Add('Операция успешно завершена')
        else
          Memo2.Lines.Add('Ошибка при выполнении команды');
      FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

procedure TForm8.SpeedButton5Click(Sender: TObject);
begin
/// отключить снижение производительности
  AllowPerformanceHit(False);
end;

procedure TForm8.SpeedButton6Click(Sender: TObject);
begin
/// включить снижение производительности
  AllowPerformanceHit(True);
end;

procedure TForm8.SpeedButton7Click(Sender: TObject);
begin
  EnableDisableFailurePredictionPolling(False);
end;

procedure TForm8.SpeedButton8Click(Sender: TObject);
begin
  EnableDisableFailurePredictionPolling(True);
end;

procedure TForm8.SpeedButton9Click(Sender: TObject);
begin
  EnableDisableHardwareFailurePrediction(False);
end;

Function TForm8.AllowPerformanceHit(onOff: bool): string;
var
/// разрешить снижение производительности
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: variant;
  s: string;
  const
  resmas: array [0 .. 2] of string = ('Successful Completion',
    'Captive Mode Required', 'Unsuccessful Completion');
begin
  try
    if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;

    if onOff then
      s := 'Включить'
    else
      s := 'Выключить';
    Memo2.Lines.Add(s + ' снижение производительности');
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
        outparam := FWbemObject.AllowPerformanceHit(onOff);
        //Memo2.Lines.Add(s+' на диске - ' + vartostr(FWbemObject.InstanceName));
        Memo2.Lines.Add('Операция успешно завершена. '+ vartostr(outparam));
        break;
      FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

Function TForm8.EnableDisableFailurePredictionPolling(onOff: bool): string;
var
/// влючить или выключить систему опроса для прогназирования сбоев
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: variant;
  s: string;
begin
  try
    if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;
    if onOff then
      s := 'Включить'
    else
      s := 'Выключить';
    Memo2.Lines.Add(s + ' систему опроса для прогназирования сбоев ');
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction WHERE  InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// /
    begin
        outparam := FWbemObject.EnableDisableFailurePredictionPolling
          (30000, onOff);
        //Memo2.Lines.Add(s + ' на - ' + vartostr(FWbemObject.InstanceName));
        //Memo2.Lines.Add(vartostr(outparam));
        Memo2.Lines.Add('Операция успешно завершена. '+vartostr(outparam));
      FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

function TForm8.EnableDisableHardwareFailurePrediction(onOff: bool): string;
var
/// влючить систему опроса для прогназирования сбоев на аппаратном уровне
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  outparam: variant;
  s: string;
begin
  try
   if selectHDD='' then  begin showmessage('Выберите диск'); exit; end;
    if onOff then
      s := 'Включение'
    else
      s := 'Выключение';
    Memo2.Lines.Add(s + ' прогнозирования сбоев на аппаратном уровне ');
    OleInitialize(nil);
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(frmDomainInfo.ComboBox2.Text,
      'root\WMI', frmDomainInfo.LabeledEdit1.Text,
      frmDomainInfo.LabeledEdit2.Text);
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM MSStorageDriver_FailurePredictFunction  WHERE InstanceName LIKE "%'+Obrez(selectHDD)+'%"', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// /
    begin
        outparam := FWbemObject.EnableDisableHardwareFailurePrediction(onOff);
       // Memo2.Lines.Add(s + ' на - ' + vartostr(FWbemObject.InstanceName));
       // Memo2.Lines.Add(vartostr(outparam));
         Memo2.Lines.Add('Операция успешно завершена. '+ vartostr(outparam));
        FWbemObject := Unassigned;
    end;
  except
    on E: Exception do
    begin
      Memo2.Lines.Add('Ошибка - ' + E.Message)
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

end.
