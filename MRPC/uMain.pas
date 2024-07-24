unit uMain;

interface

uses
  Windows, Messages, System.SysUtils, System.Variants, System.Classes, Graphics,
  Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, ActiveX, ComObj, CommCtrl, Vcl.Buttons,
    Math,
  Vcl.Menus, Vcl.OleServer, MSTSCLib_TLB, Vcl.OleCtrls, IniFiles, Vcl.ImgList,
  Vcl.Imaging.pngimage, Vcl.Imaging.jpeg, RDPCOMAPILib_TLB,
  System.Win.ScktComp, Clipbrd, DateUtils, ShellApi, Vcl.ExtDlgs, Vcl.Grids,
  Vcl.DBGrids, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS,
  FireDAC.Phys.Intf, FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt,
  FireDAC.Comp.Client, Data.DB, FireDAC.Comp.DataSet, FireDAC.UI.Intf,
  FireDAC.VCLUI.Wait, FireDAC.Comp.UI, FireDAC.VCLUI.Error, Vcl.DBCtrls,
  Vcl.AppEvnts, Themes, Vcl.ButtonGroup, JvExExtCtrls, JvSplitter,
  JvNetscapeSplitter,
  IdTCPClient, IdIcmpClient, System.ImageList, Vcl.Mask, IdRawBase, IdRawClient,
  IdBaseComponent, IdComponent;


const
  netapi32lib = 'netapi32.dll';
  NERR_Success = NO_ERROR;
  wbemFlagForwardOnly = $00000020;
  wbemFlagReturnImmediately = 16;

type


  // Структура для получения информации о рабочей станции WKSTA_INFO_100
  PWkstaInfo100 = ^TWkstaInfo100;

  TWkstaInfo100 = record
    wki100_platform_id: DWORD;
    wki100_computername: PWideChar;
    wki100_langroup: PWideChar;
    wki100_ver_major: DWORD;
    wki100_ver_minor: DWORD;
  end;

  // Структура для получения информации о рабочей станции WKSTA_INFO_102
  PWkstaInfo102 = ^TWkstaInfo102;

  TWkstaInfo102 = record
    wki102_platform_id: DWORD;
    wki102_computername: PWideChar;
    wki102_langroup: PWideChar;
    wki102_ver_major: DWORD;
    wki102_ver_minor: DWORD;
    wki102_lanroot: PWideChar;
    wki102_logged_on_users: DWORD;
  end;

  // Cтруктура для определения DNS имени контролера домена
  TDomainControllerInfoA = record
    DomainControllerName: LPSTR;
    DomainControllerAddress: LPSTR;
    DomainControllerAddressType: ULONG;
    DomainGuid: TGUID;
    DomainName: LPSTR;
    DnsForestName: LPSTR;
    Flags: ULONG;
    DcSiteName: LPSTR;
    ClientSiteName: LPSTR;
  end;

  PDomainControllerInfoA = ^TDomainControllerInfoA;

  // Структура для отображения групп
  PNetDisplayGroup = ^TNetDisplayGroup;

  TNetDisplayGroup = record
    grpi3_name: LPWSTR;
    grpi3_comment: LPWSTR;
    grpi3_group_id: DWORD;
    grpi3_attributes: DWORD;
    grpi3_next_index: DWORD;
  end;

  // Структура для отображения рабочих станций
  PNetDisplayMachine = ^TNetDisplayMachine;

  TNetDisplayMachine = record
    usri2_name: LPWSTR;
    usri2_comment: LPWSTR;
    usri2_flags: DWORD;
    usri2_user_id: DWORD;
    usri2_next_index: DWORD;
  end;

  // Структура для отображения пользователей
  PNetDisplayUser = ^TNetDisplayUser;

  TNetDisplayUser = record
    usri1_name: LPWSTR;
    usri1_comment: LPWSTR;
    usri1_flags: DWORD;
    usri1_full_name: LPWSTR;
    usri1_user_id: DWORD;
    usri1_next_index: DWORD;
  end;

  // Структура для отображения пользователей принадлежащих группе
  // или групп в которые входит пользователь
  PGroupUsersInfo0 = ^TGroupUsersInfo0;

  TGroupUsersInfo0 = record
    grui0_name: LPWSTR;
  end;

  TStringArray = array of string;
  /// /////////////////////////////////////////////
  /// структура для получения информации о пользователях на компьютерах

  THostInfo = record
    username: PWideChar;
    logon_domain: PWideChar;
    other_domains: PWideChar;
    logon_server: PWideChar;
  end;

  WKSTA_USER_INFO_0 = packed record
    wkui0_username: PWideChar;
  end;

  PWKSTA_USER_INFO_0 = ^WKSTA_USER_INFO_0;
  /// /////////////////////////////////////////////////////////

  TfrmDomainInfo = class(TForm)
    gbCurrent: TGroupBox;
    SpeedButton2: TSpeedButton;
    IdIcmpClient1: TIdIcmpClient;
    SpeedButton1: TSpeedButton;
    StatusBar1: TStatusBar;
    Memo1: TMemo;
    Button1: TButton;
    GroupBox1: TGroupBox;
    Timer1: TTimer;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    PopupMenu2: TPopupMenu;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    MainMenu1: TMainMenu;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    PopupMenuAutoload: TPopupMenu;
    N18: TMenuItem;
    N19: TMenuItem;
    SpeedButton3: TSpeedButton;
    MSIProgram: TPopupMenu;
    N20: TMenuItem;
    N21: TMenuItem;
    PopupMenu5: TPopupMenu;
    N22: TMenuItem;
    SpeedButton4: TSpeedButton;
    DiskDrive: TPopupMenu;
    Chkdsk1: TMenuItem;
    N23: TMenuItem;
    Dismount1: TMenuItem;
    Format1: TMenuItem;
    Mount1: TMenuItem;
    N24: TMenuItem;
    N25: TMenuItem;
    N26: TMenuItem;
    N27: TMenuItem;
    SpeedButton5: TSpeedButton;
    PrintDriver: TPopupMenu;
    N28: TMenuItem;
    N29: TMenuItem;
    N30: TMenuItem;
    N31: TMenuItem;
    N32: TMenuItem;
    N33: TMenuItem;
    N34: TMenuItem;
    N35: TMenuItem;
    UserProfile: TPopupMenu;
    N36: TMenuItem;
    N37: TMenuItem;
    pcRes: TPageControl;
    TabSheet2: TTabSheet;
    lvWorkStation: TListView;
    TabSheet3: TTabSheet;
    ListView1: TListView;
    TabSheet1: TTabSheet;
    ListView2: TListView;
    TabSheet4: TTabSheet;
    ListView3: TListView;
    TabSheet5: TTabSheet;
    ListView4: TListView;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    ListView5: TListView;
    TabSheet9: TTabSheet;
    N38: TMenuItem;
    N39: TMenuItem;
    N40: TMenuItem;
    TabSheet10: TTabSheet;
    ListView6: TListView;
    TabSheet11: TTabSheet;
    ListView7: TListView;
    NetworkInterface: TPopupMenu;
    N43: TMenuItem;
    G1: TMenuItem;
    N44: TMenuItem;
    ImageList1: TImageList;
    N45: TMenuItem;
    N46: TMenuItem;
    msi1: TMenuItem;
    N47: TMenuItem;
    N48: TMenuItem;
    N49: TMenuItem;
    N50: TMenuItem;
    N51: TMenuItem;
    N52: TMenuItem;
    ListView8: TListView;
    ComboBox1: TComboBox;
    Ping1: TMenuItem;
    ADIP1: TMenuItem;
    N53: TMenuItem;
    N55: TMenuItem;
    ImageList2: TImageList;
    BrendImageList3: TImageList;
    ScrollBox2: TScrollBox;
    N62: TMenuItem;
    N63: TMenuItem;
    N64: TMenuItem;
    N65: TMenuItem;
    N54: TMenuItem;
    msi3: TMenuItem;
    msi4: TMenuItem;
    N66: TMenuItem;
    Microsoft1: TMenuItem;
    Office1: TMenuItem;
    TabSheet12: TTabSheet;
    GroupBox4: TGroupBox;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    LabeledEdit6: TLabeledEdit;
    SpeedButton9: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    Panel1: TPanel;
    WakeOnLan1: TMenuItem;
    TabSheet13: TTabSheet;
    ListView9: TListView;
    GroupBox3: TGroupBox;
    GroupBox5: TGroupBox;
    ButtonedEdit1: TButtonedEdit;
    Label2: TLabel;
    WindowsHotFix: TPopupMenu;
    N56: TMenuItem;
    Label3: TLabel;
    N57: TMenuItem;
    N59: TMenuItem;
    MainPage: TPageControl;
    TabSheet14: TTabSheet;
    TabSheet15: TTabSheet;
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    ImageHardWare: TImageList;
    PCInvent: TPageControl;
    TabSheet16: TTabSheet;
    TreeView1: TTreeView;
    FDQueryDelete: TFDQuery;
    ImageList3: TImageList;
    FDUpdateSQLMAIN_PC: TFDUpdateSQL;
    FDUpdateSQLCONFIG_PC: TFDUpdateSQL;
    FDTransactionWrite1: TFDTransaction;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    GroupBox7: TGroupBox;
    GroupBox8: TGroupBox;
    StatusBarInv: TStatusBar;
    GroupBox9: TGroupBox;
    Splitter2: TSplitter;
    FindPC: TLabeledEdit;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    N60: TMenuItem;
    CheckBox1: TCheckBox;
    GroupBox6: TGroupBox;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    GroupBox10: TGroupBox;
    TreeView2: TTreeView;
    SpeedButton21: TSpeedButton;
    SpeedButton22: TSpeedButton;
    FDQueryExcelExport: TFDQuery;
    SpeedButton23: TSpeedButton;
    FDGUIxErrorDialogMRPC: TFDGUIxErrorDialog;
    TabSheet17: TTabSheet;
    GroupBox11: TGroupBox;
    GroupBox12: TGroupBox;
    DBGrid3: TDBGrid;
    SpeedButton24: TSpeedButton;
    SpeedButton25: TSpeedButton;
    GroupBox13: TGroupBox;
    StatusBarSoft: TStatusBar;
    FindPCSoft: TLabeledEdit;
    Splitter3: TSplitter;
    SpeedButton26: TSpeedButton;
    SpeedButton27: TSpeedButton;
    SpeedButton28: TSpeedButton;
    SpeedButton29: TSpeedButton;
    ListViewSoft: TListView;
    ImageList22x22: TImageList;
    TabSheet18: TTabSheet;
    GroupBox14: TGroupBox;
    ComboBox4: TComboBox;
    Label6: TLabel;
    DBGrid4: TDBGrid;
    CheckBox2: TCheckBox;
    SpeedButton30: TSpeedButton;
    Edit1: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    ComboBox3: TComboBox;
    ComboBox5: TComboBox;
    SpeedButton32: TSpeedButton;
    SpeedButton34: TSpeedButton;
    SpeedButton31: TSpeedButton;
    SpeedButton35: TSpeedButton;
    StatusBarFind: TStatusBar;
    ComboBox6: TComboBox;
    ComboBox7: TComboBox;
    SpeedButton36: TSpeedButton;
    SpeedButton33: TSpeedButton;
    SpeedButton37: TSpeedButton;
    SpeedButton38: TSpeedButton;
    PopupMenuTray: TPopupMenu;
    N61: TMenuItem;
    ImageListIcon: TImageList;
    Splitter4: TSplitter;
    TrayIcon1: TTrayIcon;
    ApplicationEvents1: TApplicationEvents;
    N67: TMenuItem;
    N68: TMenuItem;
    ImageListMainMenu: TImageList;
    SMARTHDD1: TMenuItem;
    WakeOnLan2: TMenuItem;
    N69: TMenuItem;
    PopupForListPC: TPopupMenu;
    N70: TMenuItem;
    N71: TMenuItem;
    ComboBox8: TComboBox;
    GroupBox15: TGroupBox;
    SpeedButton39: TSpeedButton;
    SpeedButton40: TSpeedButton;
    TabSheet19: TTabSheet;
    GroupBox16: TGroupBox;
    Panel2: TPanel;
    N72: TMenuItem;
    N73: TMenuItem;
    N74: TMenuItem;
    N75: TMenuItem;
    N76: TMenuItem;
    N77: TMenuItem;
    N78: TMenuItem;
    N79: TMenuItem;
    PopupDriverPnP: TPopupMenu;
    N80: TMenuItem;
    N81: TMenuItem;
    N82: TMenuItem;
    Panel3: TPanel;
    GroupBox2: TGroupBox;
    TreeView4: TTreeView;
    Panel4: TPanel;
    JvNetscapeSplitter1: TJvNetscapeSplitter;
    Panel5: TPanel;
    SpeedButton41: TSpeedButton;
    SpeedButton42: TSpeedButton;
    SpeedButton43: TSpeedButton;
    SpeedButton44: TSpeedButton;
    PagePropertiesPC: TPageControl;
    TabSheet20: TTabSheet;
    TabSheet21: TTabSheet;
    ListViewSoftinPC: TListView;
    Panel6: TPanel;
    SpeedButton45: TSpeedButton;
    SpeedButton46: TSpeedButton;
    SpeedButton47: TSpeedButton;
    GroupBox17: TGroupBox;
    SpeedButton57: TSpeedButton;
    SpeedButton58: TSpeedButton;
    N83: TMenuItem;
    N84: TMenuItem;
    N85: TMenuItem;
    JvNetscapeSplitter2: TJvNetscapeSplitter;
    GroupBox23: TGroupBox;
    Panel7: TPanel;
    ListViewDisk: TListView;
    ImageListDrive: TImageList;
    N86: TMenuItem;
    JvNetscapeSplitter3: TJvNetscapeSplitter;
    Panel8: TPanel;
    ListViewPrint: TListView;
    ImageListPrint: TImageList;
    ImageListLan: TImageList;
    SpeedButton78: TSpeedButton;
    N87: TMenuItem;
    N88: TMenuItem;
    N89: TMenuItem;
    N90: TMenuItem;
    TabSheet22: TTabSheet;
    ListViewShare: TListView;
    ImageListShareFolder: TImageList;
    PopupShareFolder: TPopupMenu;
    N91: TMenuItem;
    N92: TMenuItem;
    GetAccessMask1: TMenuItem;
    SetShareInfo1: TMenuItem;
    N93: TMenuItem;
    N94: TMenuItem;
    N95: TMenuItem;
    N96: TMenuItem;
    N97: TMenuItem;
    N98: TMenuItem;
    N99: TMenuItem;
    GroupBoxShare: TGroupBox;
    CheckBoxShare: TCheckBox;
    ComboBox2: TComboBoxEx;
    TabSheet23: TTabSheet;
    GroupBoxEvent: TGroupBox;
    ListViewEvent: TListView;
    GroupBoxEventProperties: TGroupBox;
    Splitter1: TSplitter;
    ComboBoxLogFile: TComboBox;
    Label7: TLabel;
    ComboBoxSourceEvent: TComboBox;
    Label8: TLabel;
    SpeedButton81: TSpeedButton;
    SpeedButton82: TSpeedButton;
    LabeledEdit8: TLabeledEdit;
    ImageListEventWin: TImageList;
    RadioGroupEventType: TRadioGroup;
    MemoEventInfo: TMemo;
    Panel9: TPanel;
    LabelLogName: TLabel;
    LabelLogSource: TLabel;
    LabelEventCode: TLabel;
    LabelEventType: TLabel;
    LabelEventUser: TLabel;
    LabelEventCodeOper: TLabel;
    LabelEventDate: TLabel;
    LabelEventComp: TLabel;
    Splitter5: TSplitter;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    SpeedButton83: TSpeedButton;
    PopupEvent: TPopupMenu;
    N100: TMenuItem;
    DatePickerTo: TDateTimePicker;
    CheckBox3: TCheckBox;
    TimePickerFrom: TDateTimePicker;
    Label18: TLabel;
    TimePickerTo: TDateTimePicker;
    DatePickerFrom: TDateTimePicker;
    N101: TMenuItem;
    N102: TMenuItem;
    N103: TMenuItem;
    N104: TMenuItem;
    RDP1: TMenuItem;
    uRDM1: TMenuItem;
    N105: TMenuItem;
    LabeledEdit7: TEdit;
    LabeledEdit1: TEdit;
    LabeledEdit2: TEdit;
    LabeledEdit3: TEdit;
    GroupRDPWin: TGroupBox;
    N106: TMenuItem;
    N107: TMenuItem;
    N108: TMenuItem;
    N109: TMenuItem;
    N110: TMenuItem;
    N111: TMenuItem;
    N112: TMenuItem;
    N113: TMenuItem;
    N114: TMenuItem;
    N115: TMenuItem;
    N116: TMenuItem;
    Uninstall1: TMenuItem;
    N117: TMenuItem;
    N118: TMenuItem;
    N119: TMenuItem;
    TabSheet24: TTabSheet;
    Panel10: TPanel;
    TaskListView: TListView;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    PopupMenuTask: TPopupMenu;
    N120: TMenuItem;
    N121: TMenuItem;
    N122: TMenuItem;
    N123: TMenuItem;
    N124: TMenuItem;
    N125: TMenuItem;
    Button6: TButton;
    Button7: TButton;
    ImageButton: TImageList;
    N126: TMenuItem;
    N127: TMenuItem;
    N128: TMenuItem;
    N129: TMenuItem;
    N130: TMenuItem;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Button11: TButton;
    Button13: TButton;
    Button14: TButton;
    Button15: TButton;
    Button16: TButton;
    Button17: TButton;
    Button18: TButton;
    Button20: TButton;
    Button21: TButton;
    Button22: TButton;
    SpeedButton85: TButton;
    Button24: TButton;
    Button25: TButton;
    Button12: TButton;
    Button23: TButton;
    Button19: TButton;
    N131: TMenuItem;
    N132: TMenuItem;
    N133: TMenuItem;
    N134: TMenuItem;
    N135: TMenuItem;
    N136: TMenuItem;
    N137: TMenuItem;
    N138: TMenuItem;
    N139: TMenuItem;
    N140: TMenuItem;
    N141: TMenuItem;
    SpeedButton48: TSpeedButton;
    TabSheet25: TTabSheet;
    Panel11: TPanel;
    LVMicrosoft: TListView;
    Button26: TSpeedButton;
    Button27: TSpeedButton;
    Button28: TSpeedButton;
    SpeedButton49: TSpeedButton;
    SpeedButton50: TSpeedButton;
    SpeedButton51: TSpeedButton;
    PopupWinOffice: TPopupMenu;
    N41: TMenuItem;
    N42: TMenuItem;
    N142: TMenuItem;
    N143: TMenuItem;
    N144: TMenuItem;
    N145: TMenuItem;
    StatusInvMicrosoft: TStatusBar;
    SpeedButton52: TSpeedButton;
    LabeledEdit9: TLabeledEdit;
    PageControl1: TPageControl;
    TabSheet26: TTabSheet;
    TabSheet27: TTabSheet;
    Panel12: TPanel;
    SpeedButton53: TSpeedButton;
    ComboBox9: TComboBox;
    ComboBox10: TComboBox;
    SpeedButton54: TSpeedButton;
    TabSheet28: TTabSheet;
    Panel13: TPanel;
    SpeedButton55: TSpeedButton;
    LVAntivirus: TListView;
    LabeledEdit10: TLabeledEdit;
    SpeedButton56: TSpeedButton;
    SpeedButton59: TSpeedButton;
    SpeedButton60: TSpeedButton;
    SpeedButton61: TSpeedButton;
    TabSheet29: TTabSheet;
    TabSheet30: TTabSheet;
    Panel14: TPanel;
    Panel15: TPanel;
    ListViewMicLic: TListView;
    ListViewAV: TListView;
    SpeedButton62: TSpeedButton;
    SpeedButton63: TSpeedButton;
    SpeedButton64: TSpeedButton;
    SpeedButton65: TSpeedButton;
    SpeedButton66: TSpeedButton;
    SpeedButton67: TSpeedButton;
    CopyClipBoard: TMenuItem;
    N146: TMenuItem;
    N147: TMenuItem;
    N148: TMenuItem;
    N149: TMenuItem;
    N150: TMenuItem;
    N151: TMenuItem;
    N152: TMenuItem;
    N153: TMenuItem;
    N154: TMenuItem;
    N155: TMenuItem;
    N156: TMenuItem;
    N157: TMenuItem;
    ICMP1: TMenuItem;
    WMI1: TMenuItem;
    StatusBarLoadList: TStatusBar;
    SpeedButton68: TSpeedButton;
    SpeedButton69: TSpeedButton;
    Button29: TButton;
    N58: TMenuItem;
    ClientSocket1: TClientSocket;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure lvWorkStationSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure PopupMenu1Change(Sender: TObject; Source: TMenuItem;
      Rebuild: Boolean);
    procedure PopupMenuItemsClick(Sender: TObject);
    /// / приоритет процесса
    procedure PopupMenu2ItemsClick(Sender: TObject);
    /// тип  запуска служы
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure pcResChange(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N13Click(Sender: TObject);

    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure Chkdsk1Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure Dismount1Click(Sender: TObject);
    procedure Mount1Click(Sender: TObject);
    procedure Format1Click(Sender: TObject);
    procedure N25Click(Sender: TObject);
    procedure N26Click(Sender: TObject);
    procedure N27Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure N28Click(Sender: TObject);
    procedure N35Click(Sender: TObject);
    procedure N29Click(Sender: TObject);
    procedure N31Click(Sender: TObject);
    procedure N32Click(Sender: TObject);
    procedure N33Click(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure N34Click(Sender: TObject);
    procedure N36Click(Sender: TObject);
    procedure N37Click(Sender: TObject);
    procedure ReadDB;
    /// / Чтение из базы данных
    function PutInvNumber(pcname: string; NewDescription: string): string;
    /// установка инвентарного номера
    function PutInvNumberToDataBase(pcname: string;
  NewDescription: string): string; // запись инвентарника в базу данных
    function refreshinfoPC(NamePC, user, pass: string): bool;
    /// / инвентаризация выбранного компа
    // procedure SpeedButton6Click(Sender: TObject);
    procedure MsRdpClient91AuthenticationWarningDismissed(Sender: TObject);
    procedure MsRdpClient91AutoReconnected(Sender: TObject);
    procedure MsRdpClient91AutoReconnecting(ASender: TObject;
      disconnectReason, attemptCount: Integer);
    procedure MsRdpClient91AuthenticationWarningDisplayed(Sender: TObject);
    procedure MsRdpClient91AutoReconnecting2(ASender: TObject;
      disconnectReason: Integer; networkAvailable: WordBool;
      attemptCount, maxAttemptCount: Integer);
    procedure MsRdpClient91ChannelReceivedData(ASender: TObject;
      const chanName, Data: WideString);
    procedure MsRdpClient91ConfirmClose(Sender: TObject);
    procedure MsRdpClient91Connected(Sender: TObject);
    procedure MsRdpClient91Connecting(Sender: TObject);
    procedure MsRdpClient91Disconnected(ASender: TObject; discReason: Integer);
    procedure MsRdpClient91FatalError(ASender: TObject; errorCode: Integer);
    procedure MsRdpClient91LoginComplete(Sender: TObject);
    procedure MsRdpClient91LogonError(ASender: TObject; lError: Integer);
    procedure MsRdpClient91NetworkStatusChanged(ASender: TObject;
      qualityLevel: Cardinal; bandwidth, rtt: Integer);
    procedure MsRdpClient91Warning(ASender: TObject; warningCode: Integer);
    procedure N38Click(Sender: TObject);
    procedure N39Click(Sender: TObject);
    procedure lvWorkStationCustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure N40Click(Sender: TObject);
    procedure ListView6CustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure ListView1CustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure N43Click(Sender: TObject);
    procedure ListView7CustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure G1Click(Sender: TObject);
    procedure N44Click(Sender: TObject);
    procedure ListView5CustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure MsRdpClient91Exit(Sender: TObject);
    procedure ListView8SelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure ListView8CustomDrawSubItem(Sender: TCustomListView;
      Item: TListItem; SubItem: Integer; State: TCustomDrawState;
      var DefaultDraw: Boolean);
    procedure ListView8DblClick(Sender: TObject);
    procedure ListView8ColumnClick(Sender: TObject; Column: TListColumn);
    procedure N51Click(Sender: TObject);
    procedure N48Click(Sender: TObject);
    procedure N49Click(Sender: TObject);
    procedure N50Click(Sender: TObject);
    procedure N52Click(Sender: TObject);
    procedure Ping1Click(Sender: TObject);
    procedure ComboBox1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox23KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox1Select(Sender: TObject);
    procedure N54Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ScrollBox2MouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure ScrollBox2MouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure ScrollBox2MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure FlowPanelPrintMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure TabSheet2Show(Sender: TObject);
    procedure TabSheet3Show(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure TabSheet5Show(Sender: TObject);
    procedure TabSheet8Show(Sender: TObject);
    procedure TabSheet10Show(Sender: TObject);
    procedure TabSheet11Show(Sender: TObject);
    procedure TabSheet4Show(Sender: TObject);
    procedure TabSheet7Show(Sender: TObject);
    procedure TabSheet6Show(Sender: TObject);
    procedure TabSheet9Show(Sender: TObject);
    procedure N62Click(Sender: TObject);
    procedure N64Click(Sender: TObject);
    procedure N65Click(Sender: TObject);
    procedure msi3Click(Sender: TObject);
    procedure N63Click(Sender: TObject);
    procedure msi4Click(Sender: TObject);
    procedure Microsoft1Click(Sender: TObject);
    procedure ComboBox23Select(Sender: TObject);
    procedure ComboBox2Select2(Sender: TObject);
    procedure Office1Click(Sender: TObject);
    // procedure Button3Click(Sender: TObject);
    procedure RDPSession1Error(ASender: TObject; ErrorInfo: OleVariant);
    procedure ClientSocket1Disconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Read(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocket1Connect(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocket1Connecting(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Error(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var errorCode: Integer);
    procedure ClientSocket1Write(Sender: TObject; Socket: TCustomWinSocket);
    procedure Code(var text: WideString; password: string;
      /// / процедура кодирования и декодирования файла
      decode: Boolean);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure RDPViewer1ControlLevelChangeRequest(ASender: TObject;
      const pAttendee: IDispatch; RequestedLevel: TOleEnum);
    procedure SpeedButton8Click(Sender: TObject);
    procedure RDPViewer1AttendeeConnected(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPViewer1AttendeeDisconnected(ASender: TObject;
      const pDisconnectInfo: IDispatch);
    procedure RDPViewer1ConnectionAuthenticated(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure RDPViewer1ConnectionEstablished(Sender: TObject);
    procedure RDPViewer1ConnectionFailed(Sender: TObject);
    procedure RDPViewer1ConnectionTerminated(ASender: TObject;
      discReason, ExtendedInfo: Integer);
    procedure RDPViewer1GraphicsStreamPaused(Sender: TObject);
    procedure RDPViewer1GraphicsStreamResumed(Sender: TObject);
    procedure RDPViewer1FocusReleased(ASender: TObject; iDirection: Integer);
    procedure RDPViewer1SharedDesktopSettingsChanged(ASender: TObject;
      width, height, colordepth: Integer);
    procedure RDPViewer1SharedRectChanged(ASender: TObject;
      left, top, right, bottom: Integer);
    procedure RDPViewer1StartDrag(Sender: TObject; var DragObject: TDragObject);
    procedure RDPViewer1WindowClose(ASender: TObject; const pWindow: IDispatch);
    procedure RDPViewer1WindowOpen(ASender: TObject; const pWindow: IDispatch);
    procedure RDPViewer1WindowUpdate(ASender: TObject;
      const pWindow: IDispatch);
    procedure RDPViewer1AttendeeUpdate(ASender: TObject;
      const pAttendee: IDispatch);
    procedure RDPViewer1ChannelDataReceived(ASender: TObject;
      const pChannel: IInterface; lAttendeeId: Integer;
      const bstrData: WideString);
    procedure RDPViewer1ChannelDataSent(ASender: TObject;
      const pChannel: IInterface; lAttendeeId, BytesSent: Integer);

    procedure N55Click(Sender: TObject);
    procedure WakeOnLan1Click(Sender: TObject);
    procedure ListView9ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView4ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView3ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView1ColumnClick(Sender: TObject; Column: TListColumn);
    procedure Button3Click(Sender: TObject);
    procedure ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ListView9DblClick(Sender: TObject);
    procedure N56Click(Sender: TObject);
    procedure N57Click(Sender: TObject);
    procedure ListView8MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure N59Click(Sender: TObject);
    procedure SpeedButton13Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure SpeedButton15Click(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure SpeedButton16Click(Sender: TObject);
    procedure FindPCKeyPress(Sender: TObject; var Key: Char);
    procedure SpeedButton17Click(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure SpeedButton20Click(Sender: TObject);
    procedure SpeedButton18Click(Sender: TObject);
    Function createtreeView(FDQueryinfo: TFDQuery; treeView: TTreeView): bool;
    /// построение деревьев для просмотра свойств компа в списках омпов
    Function createClearTreeView(treeView: TTreeView): bool;
    /// построение пустово дерева
    Procedure comparisonTV; // сравнение деревьева
    procedure ReadDBSoft;
    procedure SpeedButton19Click(Sender: TObject);
    procedure SpeedButton21Click(Sender: TObject);
    procedure SpeedButton22Click(Sender: TObject);
    procedure SpeedButton23Click(Sender: TObject);
    procedure SpeedButton24Click(Sender: TObject);
    procedure SpeedButton25Click(Sender: TObject);
    procedure DBGrid3TitleClick(Column: TColumn);
    procedure SpeedButton26Click(Sender: TObject);
    procedure FindPCSoftKeyPress(Sender: TObject; var Key: Char);
    procedure SpeedButton28Click(Sender: TObject);
    procedure SpeedButton27Click(Sender: TObject);
    procedure DBGrid3CllClick(Column: TColumn);
    procedure SpeedButton29Click(Sender: TObject);
    procedure ListViewSoftColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView8Resize(Sender: TObject);
    procedure TabSheet9Resize(Sender: TObject);
    procedure ComboBox4Select(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SpeedButton30Click(Sender: TObject);
    procedure ComboBox4DropDown(Sender: TObject);
    procedure DBGrid4TitleClick(Column: TColumn);
    procedure ComboBox3Select(Sender: TObject);
    procedure ComboBox3DropDown(Sender: TObject);
    procedure SpeedButton34Click(Sender: TObject);
    procedure ComboBox5DropDown(Sender: TObject);
    procedure ComboBox5Select(Sender: TObject);
    procedure SpeedButton32Click(Sender: TObject);
    procedure SpeedButton33Click(Sender: TObject);
    procedure FDQuerySelectSortAfterOpen(DataSet: TDataSet);
    procedure ComboBox3KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox5KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SpeedButton31Click(Sender: TObject);
    procedure SpeedButton36Click(Sender: TObject);
    procedure SpeedButton38Click(Sender: TObject);
    procedure SpeedButton37Click(Sender: TObject);
    procedure SpeedButton35Click(Sender: TObject);
    procedure ListView8Change(Sender: TObject; Item: TListItem;
      Change: TItemChange);

    procedure N61Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TrayIcon1Animate(Sender: TObject);
    procedure ApplicationEvents1Minimize(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure N67Click(Sender: TObject);
    procedure SMARTHDD1Click(Sender: TObject);
    procedure WakeOnLan2Click(Sender: TObject);

    procedure N69Click(Sender: TObject);
    function GetDomainController(const DomainName: String): String;
    /// функция извлечения имени контроллера домена
    function GetSID(const SecureObject: String): String;
    procedure N71Click(Sender: TObject);
    /// /получаем SID Объекта
    // procedure scanInfo;
    procedure readinfoforpcDB;
    /// чтение информации о компьютерах из списка из базы данных
    procedure SpeedButton39Click(Sender: TObject);
    procedure SpeedButton40Click(Sender: TObject);
    /// чтение пользователей, ОС,  MAC, домен в listview8
    procedure startscanlistpc;
    procedure ComboBox8KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N73Click(Sender: TObject);
    procedure ComboBox8Select(Sender: TObject);
    procedure ListView8KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N72Click(Sender: TObject);
    procedure N75Click(Sender: TObject);
    procedure N76Click(Sender: TObject);
    procedure N77Click(Sender: TObject);
    procedure N78Click(Sender: TObject);
    procedure N79Click(Sender: TObject);
    procedure N80Click(Sender: TObject);
    procedure N81Click(Sender: TObject);
    procedure N82Click(Sender: TObject);
    procedure N74Click(Sender: TObject);
    procedure ListView8Click(Sender: TObject);
    /// / запуск сканирования списка компов
    function ping(s: string): bool;
    function ConnectWMI(NamePC, user, pass: string): Boolean;
    procedure SpeedButton41Click(Sender: TObject);
    procedure SpeedButton42Click(Sender: TObject);
    procedure SpeedButton43Click(Sender: TObject);
    procedure SpeedButton44Click(Sender: TObject);
    procedure JvNetscapeSplitter1Maximize(Sender: TObject);
    procedure SpeedButton45Click(Sender: TObject);
    procedure SpeedButton46Click(Sender: TObject);
    procedure SpeedButton47Click(Sender: TObject);
    procedure N68Click(Sender: TObject);
    procedure ListViewSoftinPCColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListViewallChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure ComboBoxAllClick(Sender: TObject);
    procedure findElement;
    procedure ScanR(Sender: TObject);
    procedure ScanT(Sender: TObject);
    procedure N84Click(Sender: TObject);
    procedure N85Click(Sender: TObject);
    procedure SpeedButton66Click(Sender: TObject);
    procedure ListViewDiskColumnClick(Sender: TObject; Column: TListColumn);
    procedure N86Click(Sender: TObject);
    procedure LabeledEdit1Change(Sender: TObject);
    procedure LabeledEdit2Change(Sender: TObject);
    procedure ListView5ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView6ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView7ColumnClick(Sender: TObject; Column: TListColumn);
    procedure ListView2ColumnClick(Sender: TObject; Column: TListColumn);
    procedure SpeedButton78Click(Sender: TObject);
    procedure N89Click(Sender: TObject);
    procedure N88Click(Sender: TObject);
    procedure N90Click(Sender: TObject);
    function EnumAllTrustedDomains: Boolean;
    function EnumAllWorkStation: Boolean;
    // добаляет список групп домена в combobox
    function EnumAllWorkStationToLit(var comboboxList: TComboBox): Boolean;
    // добавляет список групп домена в TStringList
    function EnumAllGroups: Boolean;
    function EnumAllGroupsToList(GroupList: TstringList): Boolean;
    function GetCurrentComputerName: String;
    function EnumAllUsers: Boolean;
    procedure N91Click(Sender: TObject);
    procedure N92Click(Sender: TObject);
    procedure GetAccessMask1Click(Sender: TObject);
    procedure LoadNetworkShare;
    // procedure SpeedButton80Click(Sender: TObject); // загрузка сетевых ресурсов
    function createListpcForCheck(s: string): TstringList;
    function findForListView(ListViewIn: TListView): bool;
    procedure lvWorkStationMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView2MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView3MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView4MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView6MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ListView9MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure N93Click(Sender: TObject);
    procedure N94Click(Sender: TObject);
    procedure N95Click(Sender: TObject);
    procedure N96Click(Sender: TObject);
    procedure N97Click(Sender: TObject);
    procedure N98Click(Sender: TObject);
    procedure N99Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AppMessage(var Msg: TMSG; var Handled: Boolean);
    procedure ComboBox2Select(Sender: TObject);
    procedure ComboBox2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2KeyUp2(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox2DropDown(Sender: TObject);
    procedure SpeedButton81Click(Sender: TObject);
    procedure ComboBoxLogFileSelect(Sender: TObject);
    procedure SpeedButton82Click(Sender: TObject);
    procedure ListViewEventChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure ListViewEventSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure ListViewEventColumnClick(Sender: TObject; Column: TListColumn);
    procedure N100Click(Sender: TObject);
    procedure SpeedButton83Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    // procedure ComboBox2Change(Sender: TObject); /// функция создания списков компьютеров из Listview8 c выделенным checkbox
    function BrendProg(Name: String): Integer;
    procedure N101Click(Sender: TObject);
    procedure ComboBoxLogFileKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBoxSourceEventKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBoxSourceEventDropDown(Sender: TObject);
    procedure ComboBoxLogFileKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    function GetAllUserGroups(username: String): TstringList;
    procedure FormDestroy(Sender: TObject);
    procedure ConnectRDPOsherWindows(Sender: TObject);
    procedure N105Click(Sender: TObject);
    procedure RDP1Click(Sender: TObject);
    /// контекстное меню, подключение по RDP из списка компьютеров
    procedure RDP2Click(Sender: TObject);
    procedure uRDM1Click(Sender: TObject);
    procedure N106Click(Sender: TObject);
    procedure N107Click(Sender: TObject);
    procedure N108Click(Sender: TObject);
    procedure N111Click(Sender: TObject);
    procedure N112Click(Sender: TObject);
    procedure N115Click(Sender: TObject);
    procedure N116Click(Sender: TObject);
    procedure Uninstall1Click(Sender: TObject);
    procedure N117Click(Sender: TObject);
    procedure N118Click(Sender: TObject);
    procedure MainPageChange(Sender: TObject);
    procedure N119Click(Sender: TObject);
    function GetAllGroupPC(const GroupName: String): TstringList;
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure TaskListViewDblClick(Sender: TObject);
    procedure N123Click(Sender: TObject);
    procedure N124Click(Sender: TObject);
    procedure TabSheet24Show(Sender: TObject);
    procedure N122Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure N120Click(Sender: TObject);
    procedure N121Click(Sender: TObject);
    procedure N125Click(Sender: TObject);
    procedure N127Click(Sender: TObject);
    function NewNameTask(s: string): string;
    procedure GroupRDPWinClick(Sender: TObject); // введите имя новой задачи
    function popupListViewSaveAs(ListView: TListView;
      SaveDlgTitle, filename: string): bool;
    function SaveTreeView(TV: TTreeView; NamePC: string): Boolean;
    // функция сохранения свойств компьютера из TreeView  в txt
    procedure N129Click(Sender: TObject);
    procedure N130Click(Sender: TObject);
    /// передаем listview для открытия openDialog  последующего сохранения в excel или txt
    procedure PowerOn; // включение текущего компьютера
    procedure CloseSession; // завершение сеанса на текущем компьютере
    procedure ResetCurPC; // Перезагрузить текущий комп
    procedure PowerOff; // Выключить текущий комп
    procedure ForceCloseSession; // форсировать завершение сеанса
    procedure rebuutSelectPC; // форсировать перезагрузку
    procedure ForcePowerOff; // Форсировать завершение работы
    procedure ActivationWindowsPC; // активация windows
    procedure ActivationOffice; // активация офис
    procedure InventoryHardware; // инвентаризация оборудования
    procedure InventorySoftware; // bинвентаризация софта
    Procedure InventoryMicrosoft; // Инвентаризация продуктов Microsoft
    Function InventoryMicrosoftPC(NamePC, user, pass: string): Boolean;
    // Инвентаризация продуктов Microsoft Для выбранного компьютера
    procedure EditNamePC; // Переименовать компьютер
    procedure AddPCDomain; // Добавить в домен
    procedure RemovePCDomain; // Вывести из домена
    procedure OpenFormSMATR; // Открыть форму для  SMART
    procedure PropertiesForPC; // свойства компа
    procedure RenewInventNumber; // присвоить инвентарник
    procedure OpenRestoreForm; // открыть форму с точками восстановления
    procedure EditUserGroup; // редактирование пользователей и групп
    procedure NewRDPFormForPC; // открыть новое окно RDP
    procedure OpenFormForuRDM; // Открыть форму для uRDM
    procedure OpenFormExplorer; // Открыть проводник
    procedure RunTaskForPC; // запуск задачи для текущего компа
    procedure CreateNewTaskForPC; // создание задачи для текущего ПК
    procedure PowerForListPC; // включить питание на группе компов
    procedure PowerOffForListPC; // Завершение работы на группе компов
    procedure ResetForListPC; // перезагрузка доя группы компов
    procedure LogOutForListPC; // Завершение сеанса на группе компьютеров
    procedure AddPrinterForListPC; // подключить принтер для списка компов
    procedure AddDriverPrintForListPC;
    // установитьт драйвер принтера для списка компов
    procedure InstallMSIForListPC; // Установка программы MSI на группу компов
    procedure DeleteProgramMSIForListPC;
    // Удаление msi программ на группе компов
    procedure RunProcessForListPC; // запуск процесса на группе компьютеров
    procedure KillProcessForListPC; // завершить процесс на группе компьютеров
    procedure FindProcessForListPC; // Поиск процесса на группе компов
    procedure OpenRDPForListPC; // Открыть отдельный рдп клиент
    procedure CopyDelFFForListPC;
    // управление файлами и каталогами на ггруппе компьютеров
    procedure RunTaskForListPC; // запуск задачи для группы компов
    procedure CreateNewTaskForListPC;
    // создание новой задачи для группы компьютеров
    procedure ActivationKMS; // Форма для KMS активации
    function readSoftRorSelectPC(s: string; LV: TListView): bool;
    /// чтение списка программ для выбранного компа
    function readMicrosoftLic(NamePC: string; LV: TListView): bool;
    // Чтение списка лицензий продуктов Microsoft
    function readAntivirusStatus(NamePC: string; LV: TListView): bool;
    // чтение списка антивирусов и их статусов
    procedure N128Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button22Click(Sender: TObject);
    procedure Button24Click(Sender: TObject);
    procedure SpeedButton85Click(Sender: TObject);
    procedure Button25Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button23Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure N131Click(Sender: TObject);
    procedure N132Click(Sender: TObject);
    procedure N133Click(Sender: TObject);
    procedure N134Click(Sender: TObject);
    procedure N135Click(Sender: TObject);
    procedure N136Click(Sender: TObject);
    procedure N137Click(Sender: TObject);
    procedure N138Click(Sender: TObject);
    procedure N139Click(Sender: TObject);
    procedure N140Click(Sender: TObject);
    procedure N141Click(Sender: TObject);
    function inventorySoftForSelectPC(NamePC, user, pass: string): bool;
    procedure SpeedButton48Click(Sender: TObject);
    procedure Button27Click(Sender: TObject);
    procedure Button26Click(Sender: TObject);
    procedure UpdateWindowsOffice;
    // обновление списка продуктов Microsoft и лицензий
    procedure UpdateAntivirusProduct;
    // обновление списка антивирусного обеспечения
    procedure SpeedButton49Click(Sender: TObject);
    procedure LVMicrosoftDblClick(Sender: TObject);
    /// функция запуска инвентаризации ПО для выбранного компа
    procedure AllFormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonNoClose(Sender: TObject); // закрытие формы по кнопке
    procedure CloseFormEsc(Sender: TObject; var Key: Char);
    procedure Button28Click(Sender: TObject);
    procedure PCInventChange(Sender: TObject);
    procedure SpeedButton51Click(Sender: TObject);
    procedure SpeedButton50Click(Sender: TObject);
    procedure N42Click(Sender: TObject);
    procedure N142Click(Sender: TObject);
    procedure N143Click(Sender: TObject);
    procedure N41Click(Sender: TObject);
    procedure N144Click(Sender: TObject);
    procedure N145Click(Sender: TObject);
    procedure LVMicrosoftColumnClick(Sender: TObject; Column: TListColumn);
    procedure SpeedButton52Click(Sender: TObject);
    procedure LabeledEdit9KeyPress(Sender: TObject; var Key: Char);
    procedure SpeedButton53Click(Sender: TObject);
    procedure SpeedButton54Click(Sender: TObject);
    procedure ComboBox9DropDown(Sender: TObject);
    procedure ComboBox9KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComboBox9Select(Sender: TObject);
    procedure SpeedButton55Click(Sender: TObject);
    procedure SpeedButton56Click(Sender: TObject);
    procedure SpeedButton59Click(Sender: TObject);
    procedure SpeedButton60Click(Sender: TObject);
    procedure SpeedButton61Click(Sender: TObject);
    procedure LVAntivirusColumnClick(Sender: TObject; Column: TListColumn);
    procedure LVAntivirusDblClick(Sender: TObject);
    procedure PagePropertiesPCChange(Sender: TObject);
    procedure SpeedButton62Click(Sender: TObject);
    procedure SpeedButton64Click(Sender: TObject);
    procedure SpeedButton63Click(Sender: TObject);
    procedure SpeedButton65Click(Sender: TObject);
    procedure SpeedButton67Click(Sender: TObject);
    procedure LabeledEdit10KeyPress(Sender: TObject; var Key: Char);
    procedure ListViewSoftinPCDblClick(Sender: TObject);
    procedure ListViewMicLicDblClick(Sender: TObject);
    procedure ListViewAVDblClick(Sender: TObject);
    procedure CopyClipBoardClick(Sender: TObject);
    procedure WMI1Click(Sender: TObject);
    function GetAllGroupUsersInList(GroupName: String): TstringList;
    /// / выдает список пользователей для группы
    function GetAllGroupUsersInComboBox(GroupName: String): Boolean;
    // спсок пользователей выбранной группы в ComboBox в гланом окне программы
    function GetPCForGroupToLisViewComboBox(const GroupName: String): Boolean;
    function GetAllGroupUsers(const GroupName: String): Boolean;
    /// / заполняет списком компьютеров listview выбранной группы
    procedure ADIP1Click(Sender: TObject);
    procedure SpeedButton68Click(Sender: TObject);
    procedure SpeedButton69Click(Sender: TObject);
        /// ///////////////////////////////////////////////////////////////
  private
    // function GetCurrentUserName: String;
    procedure WMCopyData(var MessageData: TWMCopyData); message WM_COPYDATA;
    /// процедура для взаимодействия с приложениями
    procedure createviewerMain;
    function scanInfo(nextPC, UserScan, PassScan: string): Integer;
    function GetInfoComputerInAD(pcname: string): String;
    function GetDNSDomainName(const DomainName: String): String;
    function Log_write(level: Integer; fname, text: string): string;
    // запись в лог файл
    function GetCurrentComputerOS(s: PWideChar): string;
    function EnumNetUsers(HostName: WideString): THostInfo;
    Function ListToExcel(ListView: TListView; filename: String;
      FileFormat: string): bool; // сохраняем ListView В excel
    function ListToTxT(ListView: TListView; NameFile: string): bool;
    /// сохраняем ListView в TXT файй
    // function GetAllUserGroups(const UserName: String): Boolean;
    function DBgrid4Columns(NameFields, CaptionColumns,
      NameTable: string): bool;
    function transactionSort(s: Boolean): bool;
    procedure WMNotify(var AMessage: TWMNotify); message WM_NOTIFY;
    function killscancomp: boolean;
    // функция закрытия потока сканирования компьютеров
    function KillscanSelectGroup: bool;
    function PClistDelAdd(s: string; oper: bool): bool;
    /// функция удаления компьютера из списка сканирования
    function createListpc(s: string): TstringList;
    /// функция создания списков компьютеров из Listview8
    function OpenRegEdit(RootKey, Path, NamePC: string): Boolean;
    // открыть редактор реестра или открыть путь в редакторе реестра
    function win10Befo: Integer; // определение версии Windows
  end;

  // Функции которые предоставят нам возможность получения информации
  // function NetWkstaUserEnum(ServerName: PWideChar; Level: DWORD;
  // Bufptr: Pointer; prefmaxlen:DWORD ;entriesread:LPDWORD;
  // totalentries:LPDWORD; resumehandle:LPDWORD):DWORD; stdcall;
  // external netapi32lib;
function NetWkstaUserEnum(servername: PWideChar;
  // Указатель на строку, которая указывает DNS или NetBIOS-имя
  // удаленный сервер, на котором должна выполняться функция.
  // Если этот параметр nil, используется локальный компьютер.
  level: DWORD;
  // Level = 0: вернуть имена пользователей, которые в данный момент вошли на рабочую станцию.
  var bufptr: Pointer; // Указатель на буфер, который получает данные
  prefmaxlen: DWORD;
  // Указывает предпочтительную максимальную длину возвращаемых данных в байтах.
  var entriesread: PDWord;
  // Указатель на значение, которое получает количество фактически перечисленных элементов.
  var totalentries: PDWord; // общее количество записей
  var resumehandle: PDWord)
// содержит дескриптор резюме, который используется для продолжения существующего поиска
  : Longint; stdcall; external 'netapi32.dll' Name 'NetWkstaUserEnum';

function NetApiBufferFree(Buffer: Pointer): DWORD; stdcall;
  external netapi32lib;
function NetWkstaGetInfo(servername: PWideChar; level: DWORD; bufptr: Pointer)
  : DWORD; stdcall; external netapi32lib;

  function NetGetDCName(servername: PWideChar; DomainName: PWideChar;
  var bufptr: PWideChar): DWORD; stdcall; external netapi32lib;

function DsGetDcName(ComputerName, DomainName: PChar; DomainGuid: PGUID;
  SiteName: PChar; Flags: ULONG;
  var DomainControllerInfo: PDomainControllerInfoA): DWORD; stdcall;
  external netapi32lib name 'DsGetDcNameA';

function NetQueryDisplayInformation(servername: PWideChar; level: DWORD;
  Index: DWORD; EntriesRequested: DWORD; PreferredMaximumLength: DWORD;
  var ReturnedEntryCount: DWORD; SortedBuffer: Pointer): DWORD; stdcall;
  external netapi32lib;
function NetGroupGetUsers(servername: PWideChar; GroupName: PWideChar;
  level: DWORD; var bufptr: Pointer; prefmaxlen: DWORD; var entriesread: DWORD;
  var totalentries: DWORD; resumehandle: PDWord): DWORD; stdcall;
  external netapi32lib;
function NetUserGetGroups(servername: PWideChar; username: PWideChar;
  level: DWORD; var bufptr: Pointer; prefmaxlen: DWORD; var entriesread: DWORD;
  var totalentries: DWORD): DWORD; stdcall; external netapi32lib;
function NetEnumerateTrustedDomains(servername: PWideChar;
  DomainNames: PWideChar): DWORD; stdcall; external netapi32lib;
function CustomSortProc(Item1, Item2: TListItem; ParamSort: Integer)
  : Integer; stdcall;

// function ConvertStringSidToSid(StringSid: PChar; out Sid: PSID):BOOL; stdcall;
// external advapi32 name {$IFDEF UNICODE} 'ConvertStringSidToSidW' {$ELSE}
// 'ConvertStringSidToSidA'{$ENDIF};

function ConvertSidToStringSid(Sid: PSID; out StringSid: PChar): bool; stdcall;
  external advapi32 name {$IFDEF UNICODE} 'ConvertSidToStringSidW' {$ELSE}
  'ConvertSidToStringSidA'{$ENDIF};


var
  frmDomainInfo: TfrmDomainInfo;

  MyExcel: OleVariant;
  /// Объект excel
  StopProc, TypeNewProc: bool;
  HeaderID, MyPriority: Integer;
  MyMemory: real;
  SumTime, SumMemory: int64;
  invNumber, MyPS, MyUser, MyPasswd, SelectServic, TypeRunService, NewProcMyPS,
    NewProgramMyPS, NewNamePC, PasswordDomain, UserDomain: string;
  MyCommandLine, MyCurrentDirectory, CurrentDC: String;
  /// запуск процесса
  ThrInvPC, ThrInvSoft, SelectedPCPing,ThreadMO: TThread;
  CpuLogPr, CpuNum, ActionServic: byte;
  NewNamePrint, CurrentDomainName: string;
  myShutdown, FormatClusterSize: Integer;
  listDelProg: TstringList;
  SelectedHDD, SelectedPrint, FormatFileSystem, FormatLabel, AddNewNetPrint,
    UserSID, NewUserName, SelectedDriver, NetworkInterfaceID: string;
  FixErrors, VigorousIndexCheck, SkipFolderCycle, ForceDismount,
    RecoverBadSectors, OkToRunAtBootUp, FormatQuickFormat,
    FormatEnableCompression: bool;
    FRegisteredSessionNotification,EventUserLogin:boolean; // признак токго что приложение зарегистрировано как получатель системных сообщений
  SetPeak, mem: Integer;
  MyRDPClient: TMsRdpClient9;
  SelectedPC, PingPCList, SelectedPCkillProc, SelectedPCShutDown, ListPCConf,
    ListPCWO: TstringList;
  GroupPC, OutForPing: bool;
  levelCntrl, UnloadFileSetting, AccessSettingLevel, ConnectionEnable: Boolean;
  /// / уровень привелегий при подключении к рабочему столу,загрузка файла настрое
  IpBroadCast, OSVersion: string;
  sort, PCCheck, PcInLan, PCCol: Integer;
  /// сортировка в TlistView
  SortLV8: bool;
  databaseName, databaseDriverID, databaseUserName, databasePassword,
    databaseProtocol, databaseServer, databaseport: string;
  InvWarning, InvSoftWarning, GroupselectProc: string;
  /// строка для учета нарушения конфигурации, используется в настройках и потоке инвентаризации оборудования
  ListSoftID, ListMACAddress, SelectProc: TstringList;
  sm: TStyleManager;
  InventConf, databaseconnected, inventScan, SolveExitInvScan, InventSoft, SolveExitInvConf, SolveExitInvSoft,
    Traymin, InventMicrosoft, SolveExitInvMicrosoft: bool;
  ListViewMousPoint: TPoint;
  MRDUViewerMain: TRDPViewer;
  ///
  captionMainForms,captionMainFormsPro: string;
  startscanT, startscanR: TTimer;
  /// таймеры сканирования сети
  pingtimeout,PingType: Integer;
  urdmport: string;
  uRDMScan, RPCport: Boolean;
  HM: THandle;
  /// / мьютекс для копии приложения
  NameRunTASK: String; // Передаем в поток имя таблицы для задач

implementation

uses unit2, unit3, unit4, unit5, unit6, unit7, unit9, unit10, NewServic,
  PCShutdown, Newstartup,
  DlgNewInstall, DeleteProg, UninstallOtherProg, DlgChkdsk, DefragHDD, Dismount,
  HDDMount, DlgFormatHDD, DlgJoinDomainOrWorkGroup, DlgPCRename,
  DlgUnjoinDomainOrWorkgroup, NetworkAddPrint, AddPrintDriverDialog,
  PrintCancelAllJob, PrintTestPage, PrintRenamePrintDialog, PrintResumePrint,
  PrintPausePrint,
  PrinterSetDefaultPrinter, NewUserProfileDialog, PrintDeleteThread,
  PcProperties,
  UserProfileDelete, DriverPnPDisableThread, DriverPnPEnableThread,
  PropertiesNetworkDialog, NetworkAdapterDisable, NetworkAdapterEnable,
  NetworkConfiguration, SelectedPCShotDownThread, GeneralPingPC, IPAddress,
  SelectedPCKillProcess, SelectedPCDeleteProgramDialog, WindowsActivation,
  OfficeActivation, RemoteDesktopSettingDialog, MRDUNew, SettingsProgramForm,
  WOLThread,
  MyInventoryPC, MyDM, ExportExcel, MyInventorySoft, SmartStat, PrintProperty,
  OneStartForm, FormFSShare, ShareFS, FormFindProcess, EventWindows,
  LogFileEventProperties, LoadCurencyList, PropertyAccount, AccountLocal,
  NewRDPWin, RemoteExplorer, FormCopyFF, FormForCopyPCList,
  FormNewProcess, TaskEdit, TaskSelect, MainPopup, RegEditForm, RestoreWindows,
  KMS, InventoryMicrosoftProduct, InventoryWindowsOffice,MessageSystem,PingForMain,EditProcMsi,My_Proc;
{ //FWMIService.Security_.Privileges.AddAsString('SeCreateTokenPrivilege',true); // Требуется для создания основного объекта токена.
  //FWMIService.Security_.Privileges.AddAsString('SeAssignPrimaryTokenPrivilege',true);  // Требуется для замены токена уровня процесса.
  //FWMIService.Security_.Privileges.AddAsString('SeLockMemoryPrivilege',true);         //Требуется для блокировки страниц в памяти.
  //FWMIService.Security_.Privileges.AddAsString('SeIncreaseQuotaPrivilege',true);     //Требуется настроить квоты памяти для процесса.
  //FWMIService.Security_.Privileges.AddAsString('SeMachineAccountPrivilege',true);   // Требуется для добавления рабочих станций в домен.
  //FWMIService.Security_.Privileges.AddAsString('SeTcbPrivilege',true);             //Требуется действовать как часть операционной системы. Держатель является частью надежной компьютерной базы.
  FWMIService.Security_.Privileges.AddAsString('SeSecurityPrivilege',true);        //Требуется для управления аудитом и журналом безопасности NT.
  //FWMIService.Security_.Privileges.AddAsString('SeTakeOwnershipPrivilege',true);  //Требуется принять права собственности на файлы или другие объекты, не имея записи контроля доступа (ACE) в списке контроля доступа по усмотрению (DACL).
  //FWMIService.Security_.Privileges.AddAsString('SeLoadDriverPrivilege',true);     //Требуется для загрузки или выгрузки драйвера устройства.
  //FWMIService.Security_.Privileges.AddAsString('SeSystemProfilePrivilege',true);     //Требуется для сбора профильной информации о производительности системы.
  //FWMIService.Security_.Privileges.AddAsString('SeSystemtimePrivilege',true);      // Требуется изменить системное время.
  //FWMIService.Security_.Privileges.AddAsString('SeProfileSingleProcessPrivilege',true); //Требуется для сбора информации профиля для одного процесса.
  //FWMIService.Security_.Privileges.AddAsString('SeIncreaseBasePriorityPrivilege',true);  // Требуется для увеличения приоритета планирования
  //FWMIService.Security_.Privileges.AddAsString('SeCreatePagefilePrivilege',true);       // Требуется для создания файла подкачки.
  //FWMIService.Security_.Privileges.AddAsString('SeCreatePermanentPrivilege',true);      //Требуется для создания постоянных общих объектов.
  //FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege',true);               //Требуется для резервного копирования файлов и каталогов, независимо от списка ACL, указанного для файла.
  //FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',true);            //Требуется для восстановления файлов и каталогов независимо от списка ACL, указанного для файла.
  //FWMIService.Security_.Privileges.AddAsString('SeShutdownPrivilege',true);         //   Требуется выключить локальную систему
  //FWMIService.Security_.Privileges.AddAsString('SeDebugPrivilege',true);             //Требуется для отладки и настройки памяти процесса, принадлежащего другой учетной записи.
  //FWMIService.Security_.Privileges.AddAsString('SeAuditPrivilege',true);             //Требуется для создания записей аудита в журнале безопасности NT. Только защищенные серверы должны иметь эту привилегию.
  //FWMIService.Security_.Privileges.AddAsString('SeSystemEnvironmentPrivilege',true); // Требуется для изменения энергонезависимой оперативной памяти систем, которые используют этот тип памяти для хранения данных конфигурации.
  //FWMIService.Security_.Privileges.AddAsString('SeChangeNotifyPrivilege',true);     //  Требуется для получения уведомлений об изменениях файлов или каталогов и обхода проверок доступа. Эта привилегия включена по умолчанию для всех пользователей.
  //FWMIService.Security_.Privileges.AddAsString('SeRemoteShutdownPrivilege',true);   //Требуется выключить удаленный компьютер.
  //FWMIService.Security_.Privileges.AddAsString(' SeUndockPrivilege',true);           //Требуется снять ноутбук с док-станции.
  //FWMIService.Security_.Privileges.AddAsString('SeSyncAgentPrivilege',true);         // Требуется для синхронизации данных службы каталогов.
  //FWMIService.Security_.Privileges.AddAsString('SeEnableDelegationPrivilege',true);   // Требуется, чтобы учетные записи компьютеров и пользователей были доверенными для делегирования.
  //FWMIService.Security_.Privileges.AddAsString('SeManageVolumePrivilege',true);       //Требуется для выполнения задач объемного обслуживания. }
{$R *.dfm}

/// ///////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.WMCopyData(var MessageData: TWMCopyData);
var
  CDS: TCopyDataStruct;
  AllStr: TstringList;
  i: Integer;
const
  SEND_SETLISTTEXT = 1;
  /// отправить
  RECEIVE_REQ_LISTTEXT = 2;
  /// запрос на получение списка
  RECEIVE_REQ_USPASS = 3;
  /// запрос пароля и логина
BEGIN
  // Устанавливаем свойства метки, если заданная команда совпадает
  if MessageData.CopyDataStruct.dwData = RECEIVE_REQ_LISTTEXT then
    Begin
      ComboBox2DropDown(ComboBox2);
      /// обновление статуса компов в комбобоксе, после из него возьмем список со статусами
      AllStr := TstringList.Create;
      for i := 0 to ComboBox2.Items.Count - 1 do
      /// получаем список компов со статусами
        begin
          AllStr.Add(ComboBox2.Items[i] + '=' +
            inttostr(ComboBox2.ItemsEx[i].ImageIndex));
        end;
      CDS.dwData := SEND_SETLISTTEXT;
      /// устанавливаем признак отправки текста
      CDS.cbData := Length(AllStr.CommaText) + 1;
      // Устанавливаем длину передаваемых данных
      GetMem(CDS.lpData, CDS.cbData);
      // Выделяем память буфера для передачи данных
      try
        // Копируем данные в буфер
        StrPCopy(CDS.lpData, AnsiString(AllStr.CommaText));
        // Отсылаем сообщение в окно с заголовком StringReceiver
        SendMessage(FindWindow('TFormRDP', nil), WM_COPYDATA, Handle,
          Integer(@CDS));
      finally
        AllStr.Free;
        // Высвобождаем буфер
        FreeMem(CDS.lpData, CDS.cbData);
      end;
    End;

  if MessageData.CopyDataStruct.dwData = RECEIVE_REQ_USPASS then
    Begin
      AllStr := TstringList.Create;
      AllStr.Add('Log=' + LabeledEdit1.text);
      AllStr.Add('Pass=' + LabeledEdit2.text);
      AllStr.Add('Dom=' + LabeledEdit3.text);
      CDS.dwData := RECEIVE_REQ_USPASS;
      /// устанавливаем признак отправки текста
      CDS.cbData := Length(AllStr.CommaText) + 1;
      // Устанавливаем длину передаваемых данных
      GetMem(CDS.lpData, CDS.cbData);
      // Выделяем память буфера для передачи данных
      try
        // Копируем данные в буфер
        StrPCopy(CDS.lpData, AnsiString(AllStr.CommaText));
        // Отсылаем сообщение в окно с заголовком StringReceiver
        // SendMessage(GetWindow(FindWindow('TFormRDP', nil),GW_OWNER),
        // WM_COPYDATA, Handle, Integer(@CDS));
        SendMessage(FindWindow('TFormRDP', nil), WM_COPYDATA, Handle,
          Integer(@CDS));
      finally
        AllStr.Free;
        // Высвобождаем буфер
        FreeMem(CDS.lpData, CDS.cbData);
      end;
    End;
END;

procedure TfrmDomainInfo.WMI1Click(Sender: TObject);
begin
  if (ListView8.Items.Count <> 0) and (ListView8.SelCount = 1) then
    begin
      if ping(ListView8.Items[ListView8.ItemFocused.Index].SubItems[0]) then
        ConnectWMI(ListView8.Items[ListView8.ItemFocused.Index].SubItems[0],
          LabeledEdit1.text, LabeledEdit2.text);
    end;

end;

/// ////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.EnumNetUsers(HostName
  : WideString { ; Users: TStrings } ): THostInfo;
const
  STR_ERROR_ACCESS_DENIED =
    'Пользователь не имеет доступа к запрашиваемой информации.';
  STR_ERROR_MORE_DATA =
    'Укажите достаточно большой буфер для приема всех записей.';
  STR_ERROR_INVALID_LEVEL = 'Недопустимый параметр уровня.';
var
  Info: Pointer;
  ElTotal: PDWord;
  ElCount: PDWord;
  Resume: PDWord;
  Error: Longint;
const
  MAX_PREFERRED_LENGTH = DWORD(-1);
begin
  try
    Resume := 0;
    Error := NetWkstaUserEnum(PWideChar(HostName),
      /// имякомпа
      1, // 0-WKSTA_USER_INFO_0 , 1- WKSTA_USER_INFO_1
      Info, // Указатель на буфер, который получает данные -- вся инфа получаемая с компов
      MAX_PREFERRED_LENGTH,
      // Если вы укажете MAX_PREFERRED_LENGTH, функция выделяет объем памяти, необходимый для данных  //256 * Integer(ElTotal)
      ElCount,
      /// Указатель на значение, которое получает количество фактически перечисленных элементов.
      ElTotal,
      /// Указатель на значение, которое получает общее количество записей, которые могли быть перечислены с текущей позиции возобновления.
      Resume); // Указатель на значение, которое содержит дескриптор возобновления, который используется для продолжения существующего поиска. Дескриптор должен быть нулевым при первом вызове и оставлен без изменений для последующих вызовов.

    case Error of
      ERROR_ACCESS_DENIED:
        Result.username := STR_ERROR_ACCESS_DENIED;
      ERROR_MORE_DATA:
        Result.username := STR_ERROR_MORE_DATA;
      ERROR_INVALID_LEVEL:
        Result.username := STR_ERROR_INVALID_LEVEL
    else
      if Info <> nil then
        begin
          Result := THostInfo(Info^);
        end
      else
        begin
          Result.username := PChar(SysErrorMessage(Error));
          Result.logon_domain := PChar(SysErrorMessage(Error));
          Result.other_domains := PChar(SysErrorMessage(Error));
          Result.logon_server := PChar(SysErrorMessage(Error));
        end;
    end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add(HostName + ': Ошибка NetWkstaUserEnum  :' + E.Message);
      end;
  end;
end;

/// ===========Получаем SID по имени пользователя
/// //////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.GetSID(const SecureObject: String): String;
var
  Sid: PSID;
  StringSid: PChar;
  ReferencedDomain: String;
  cbSid, cbReferencedDomain: DWORD;
  peUse: SID_NAME_USE;
begin
  cbSid := 128;
  cbReferencedDomain := 16;
  GetMem(Sid, cbSid);
  try
    SetLength(ReferencedDomain, cbReferencedDomain);
    if LookupAccountName(PChar(''),
      /// PChar(ledDNSName.Text)  /// DNS Имя
      PChar(SecureObject), Sid, cbSid, @ReferencedDomain[1], cbReferencedDomain,
      peUse) then
      begin
        ConvertSidToStringSid(Sid, StringSid);
        // StringSid:=(StringSid);
        if StringSid = '' then
          StringSid := 'Unknown';
        Result := (StringSid);
      end;
  finally
    FreeMem(Sid);
  end;
end;

procedure TfrmDomainInfo.GroupRDPWinClick(Sender: TObject);
begin

end;

// Довольно простая функция, возвращает только имена пользователей принадлезжащих группе
// =============================================================================
function TfrmDomainInfo.GetAllGroupUsersInList(GroupName: String): TstringList;
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries, resumehandle: DWORD;
  i: Integer;
  yesBool: Boolean;
begin
  // На вход подается список который мы будем заполнять
  Result := TstringList.Create;
  // Обязательная инициализация
  resumehandle := 0;
  prefmaxlen := DWORD(-1);
  // Выполняем
  yesBool := NetGroupGetUsers
    (StringToOleStr(GetDomainController(CurrentDomainName)),
    StringToOleStr(GroupName), 0, Pointer(Info), prefmaxlen, entriesread,
    totalentries, @resumehandle) = NERR_Success;
  // Смотрим результат...
  if yesBool then
    try
      Tmp := Info;
      for i := 0 to entriesread - 1 do
        begin
          if pos('$', Tmp^.grui0_name) <> Length(Tmp^.grui0_name) then
          // если последний символ <> $ то это имя пользователя, иначе это имя компа
            Result.Add(Tmp^.grui0_name); // выводим результат из структуры
          Inc(Tmp);
        end;
    finally
      // Не забываем, ибо может быть склероз :)
      NetApiBufferFree(Info);
    end;
end;

/// /////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.GetAllGroupUsersInComboBox(GroupName: String): Boolean;
// спсок пользователей выбранной группы в ComboBox в гланом окне программы
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries, resumehandle: DWORD;
  i: Integer;
  yesBool: Boolean;
  ProcBar: TprogressBar;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка пользователей группы ' +
    GroupName;
  try
    ComboBox8.Clear;
    resumehandle := 0;
    prefmaxlen := DWORD(-1);
    yesBool := NetGroupGetUsers
      (StringToOleStr(GetDomainController(CurrentDomainName)),
      StringToOleStr(GroupName), 0, Pointer(Info), prefmaxlen, entriesread,
      totalentries, @resumehandle) = NERR_Success;
    if yesBool then
      try
        Tmp := Info;
        ProcBar.Max := entriesread;
        for i := 0 to entriesread - 1 do
          begin
            Application.ProcessMessages;
            ProcBar.Position := i;
            if pos('$', Tmp^.grui0_name) <> Length(Tmp^.grui0_name) then
            // если последний символ <> $ то это имя пользователя, иначе это имя компа
              ComboBox8.Items.Add(Tmp^.grui0_name);
            // выводим результат из структуры
            Inc(Tmp);
          end;
      finally
        NetApiBufferFree(Info);
      end;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
try
Memo1.Lines.Add('Всего пользователей в группе '+GroupName+': '+inttostr(ComboBox8.Items.Count));
except
exit;
end;
end;

// Возвращает группы безопасности в которые входит пользователь(заметьте - структура таже)
// =============================================================================
function TfrmDomainInfo.GetAllUserGroups(username: String): TstringList;
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries: DWORD;
  i: Integer;
  yesBool: Boolean;
begin
  Result := TstringList.Create;
  prefmaxlen := DWORD(-1);
  yesBool := NetUserGetGroups
    (StringToOleStr(GetDomainController(CurrentDomainName)),
    StringToOleStr(username), 0, Pointer(Info), prefmaxlen, entriesread,
    totalentries) = NERR_Success;
  if yesBool then
    try
      Tmp := Info;
      for i := 0 to entriesread - 1 do
        begin
          Result.Add(Tmp^.grui0_name);
          Inc(Tmp);
        end;
    finally
      NetApiBufferFree(Info);
    end;
end;

/// ///////////////////////////////////////////////////////
// ========================================================
/// //Получаем имя пользователя + домен по SID   (домен\пользователь)
function SidToAcountName(SecureObject, s: string): String;
var
  Sid: PSID;
  peUse: DWORD;
  cchDomain: DWORD;
  cchName: DWORD;
  Name: array of Char;
  Domain: array of Char;
begin
  Result := '';
  Sid := nil;
  // First convert String SID to SID
  Win32Check(ConvertStringSidToSid(PChar(SecureObject), Sid));
  cchName := 0;
  cchDomain := 0;
  // Get Length             // nil это локальный комп
  if (not LookupAccountSid( { Pchar(s) или } nil, Sid, nil, cchName, nil,
    cchDomain, peUse)) and (GetLastError = ERROR_INSUFFICIENT_BUFFER) then
    begin
      SetLength(Name, cchName);
      SetLength(Domain, cchDomain);
      if LookupAccountSid(nil, Sid, @Name[0], cchName, @Domain[0], cchDomain,
        peUse) then
        begin
          // note: cast to PChar because LookupAccountSid returns zero terminated string
          Result := PChar(Domain) + '\' + PChar(Name);
        end;
    end;
end;

// Данная функция получает информацию о всех пользователях присутствующих в домене
// =============================================================================
function TfrmDomainInfo.EnumAllUsers: Boolean;
var
  Tmp, Info: PNetDisplayUser;
  i, CurrIndex, EntriesRequest, PreferredMaximumLength, ReturnedEntryCount
    : Cardinal;
  Error: DWORD;
  ProcBar: TprogressBar;
  z: Integer;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  ProcBar.Max := 100;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка пользователей домена ' +
    CurrentDomainName;
  try
    CurrIndex := 0;
    ComboBox8.Clear;
    repeat
      Info := nil;
      // NetQueryDisplayInformation возвращает информацию только о 100-а записях
      // для того чтобы получить всю информацию используется третий параметр,
      // передаваемый функции, который определяет с какой записи продолжать
      // вывод информации
      ProcBar.Position := 0;
      z := 0;
      EntriesRequest := 100;
      PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayUser);
      ReturnedEntryCount := 0;
      // Для выполнения функции, в нее нужно передать DNS имя контролера домена
      // (или его IP адрес), с которого мы хочем получить информацию
      // Для получения информации о пользователях используется структура NetDisplayUser
      // и ее идентификатор 1 (единица) во втором параметре
      Error := NetQueryDisplayInformation
        (StringToOleStr(GetDomainController(CurrentDomainName)), 1, CurrIndex,
        EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
      // При безошибочном выполнении фунции будет результат либо
      // 1. NERR_Success - все записи возвращены
      // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
      if Error in [NERR_Success, ERROR_MORE_DATA] then
        try
          Tmp := Info;
          // Выводим информацию которую вернула функция в структуру
          for i := 0 to ReturnedEntryCount - 1 do
            begin
              Application.ProcessMessages;
              Inc(z);
              ProcBar.Position := z;
              ComboBox8.Items.Add(Tmp^.usri1_name); // Имя пользователя
              // SubItems.Add(Tmp^.usri1_comment);    // Комментарий
              // SubItems.Add(GetSID(Caption));       // Его SID
              // Запоминаем индекс с которым будем вызывать повторно функцию (если нужно)
              CurrIndex := Tmp^.usri1_next_index;
              Inc(Tmp);
            end;
        finally
          // Грохаем выделенную при вызове NetQueryDisplayInformation память
          NetApiBufferFree(Info);
        end;
      // Если результат выполнения функции ERROR_MORE_DATA
      // (т.е. есть еще данные) - вызываем функцию повторно
    until Error in [NERR_Success, ERROR_ACCESS_DENIED];
    // Ну и возвращаем результат всего что мы тут накодили
    Result := Error = NERR_Success;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
end;

/// ///////////////////////////////////////////////////////////////////////////////////
// Данная функция получает информацию о всех группах присутствующих в домене
// =============================================================================
function TfrmDomainInfo.EnumAllGroups: Boolean;
var
  Tmp, Info: PNetDisplayGroup;
  i, CurrIndex, EntriesRequest, PreferredMaximumLength, ReturnedEntryCount
    : Cardinal;
  Error: DWORD;
  ProcBar: TprogressBar;
  z: Integer;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  ProcBar.Max := 100;
  StatusBarLoadList.Panels[2].text :=
    'Загрузка списка групп безопасности домена ' + CurrentDomainName;
  try
    CurrIndex := 0;
    repeat
      Info := nil;
      // NetQueryDisplayInformation возвращает информацию только о 100-а записях
      // для того чтобы получить всю информацию используется третий параметр,
      // передаваемый функции, который определяет с какой записи продолжать
      // вывод информации
      EntriesRequest := 100;
      z := 0;
      PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayGroup);
      ReturnedEntryCount := 0;
      // Для выполнения функции, в нее нужно передать DNS имя контролера домена
      // (или его IP адрес), с которого мы хочем получить информацию
      // Для получения информации о группах используется структура NetDisplayGroup
      // и ее идентификатор 3 (тройка) во втором параметре
      Error := NetQueryDisplayInformation
        (StringToOleStr(GetDomainController(CurrentDomainName)), 3, CurrIndex,
        EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
      // При безошибочном выполнении фунции будет результат либо
      // 1. NERR_Success - все записи возвращены
      // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
      if Error in [NERR_Success, ERROR_MORE_DATA] then
        try
          Tmp := Info;
          for i := 0 to ReturnedEntryCount - 1 do
            begin
              Application.ProcessMessages;
              Inc(z);
              ProcBar.Position := z;
              with ComboBox1.Items do
                begin
                  Add(Tmp^.grpi3_name); // Имя группы
                  // Запоминаем индекс с которым будем вызывать повторно функцию (если нужно)
                  CurrIndex := Tmp^.grpi3_next_index;
                end;
              Inc(Tmp);
            end;
          /// ////////////////////////////////////////////////////
        finally
          // Чтобы небыло утечки ресурсов, освобождаем память занятую функцией под структуру
          NetApiBufferFree(Info);
        end;
      // Если результат выполнения функции ERROR_MORE_DATA - вызываем функцию повторно
    until Error in [NERR_Success, ERROR_ACCESS_DENIED];
    // Ну и возвращаем результат всего что мы тут накодили
    Result := Error = NERR_Success;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
end;

/// //////////////////////////////////////////////
function TfrmDomainInfo.EnumAllGroupsToList(GroupList: TstringList): Boolean;
// список всех групп домена в StringList
var
  Tmp, Info: PNetDisplayGroup;
  i, CurrIndex, EntriesRequest, PreferredMaximumLength, ReturnedEntryCount
    : Cardinal;
  Error: DWORD;
begin
  CurrIndex := 0;
  repeat
    Info := nil;
    EntriesRequest := 100;
    PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayGroup);
    ReturnedEntryCount := 0;
    Error := NetQueryDisplayInformation
      (StringToOleStr(GetDomainController(CurrentDomainName)), 3, CurrIndex,
      EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
    if Error in [NERR_Success, ERROR_MORE_DATA] then
      try
        Tmp := Info;
        for i := 0 to ReturnedEntryCount - 1 do
          begin
            with GroupList do
              begin
                Add(Tmp^.grpi3_name); // Имя группы
                CurrIndex := Tmp^.grpi3_next_index;
              end;
            Inc(Tmp);
          end;
      finally
        NetApiBufferFree(Info);
      end;
  until Error in [NERR_Success, ERROR_ACCESS_DENIED];
  Result := Error = NERR_Success;
end;

// Получаем DNS имя контроллера домена
// =============================================================================
function TfrmDomainInfo.GetDNSDomainName(const DomainName: String): String;
const
  DS_IS_FLAT_NAME = $00010000;
  DS_RETURN_DNS_NAME = $40000000;
var
  GUID: PGUID;
  DomainControllerInfo: PDomainControllerInfoA;
begin
  GUID := nil;
  // Для большинства операций нам потребуется IP адрес контроллера домена
  // или его DNS имя, которое мы получим вот так:
 if DsGetDcName(nil, PChar(CurrentDomainName), GUID, nil, DS_IS_FLAT_NAME or
   DS_RETURN_DNS_NAME, DomainControllerInfo) = NERR_Success then
   // https://docs.microsoft.com/en-us/windows/win32/api/dsgetdc/nf-dsgetdc-dsgetdcnamea
    // Параметры которые мы передаем означают:
    // DS_IS_FLAT_NAME - передаем просто имя домена
    // DS_RETURN_DNS_NAME - ждем получения DNS имени
    try
      Result := DomainControllerInfo^.DomainControllerName;
      // Результат собсно тут...
    finally
      // Склероз это болезнь, ее нужно лечить...
      NetApiBufferFree(DomainControllerInfo);
    end;
end;

// =============================================================================
// Ну тут без комментариев - просто получаем имя основного контроллера домена
// =============================================================================
function TfrmDomainInfo.GetDomainController(const DomainName: String): String;
var
  Domain: WideString;
  Server: PWideChar;
begin
  Domain := StringToOleStr(DomainName);
  if NetGetDCName(nil, @Domain[1], Server) = NERR_Success then
    try
    if CurrentDC='' then Result := Server // если домен не указан в настройках то используем основной контроллер который определили
    else result:='\\'+CurrentDC;               // иначе использем тот который указан в настройках
    finally
      NetApiBufferFree(Server);
    end;

end;

procedure TfrmDomainInfo.JvNetscapeSplitter1Maximize(Sender: TObject);
/// / если скрываем панель совойств
begin
  if JvNetscapeSplitter1.Maximized then
    begin
      createClearTreeView(TreeView4);
      ListViewSoftinPC.Clear;
      GroupBox2.Caption := '';
    end;
end;

procedure TfrmDomainInfo.Edit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
    SpeedButton30.Click;
end;

function TfrmDomainInfo.GetInfoComputerInAD(pcname: string): String;
var
  Info: PWkstaInfo102;
  Error: DWORD;
  CompPWC: PWideChar;
begin
  try
    begin
      CompPWC := PWideChar(WideString(pcname));
      Error := NetWkstaGetInfo(CompPWC, 102, @Info);
      if Error <> 0 then
        raise Exception.Create(SysErrorMessage(Error));
      Result := inttostr(Info^.wki102_ver_major) + '.' +
        inttostr(Info^.wki102_ver_minor);
    end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка GetInfoComputerInAD  :' + E.Message);
      end;
  end;
end;

/// ///////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.GetCurrentComputerName: String;
/// получение информации о компьютере в сети и домене
var
  Info: PWkstaInfo100;
  Error: DWORD;
begin
  try
    // А для этого мы воспользуемся следующей функцией
    Error := NetWkstaGetInfo(nil, 100, @Info);
    if Error <> 0 then
      raise Exception.Create(SysErrorMessage(Error));
    // Как видно, вызов который возвращает обычную структуру, из которой и прочитаем, все что нужно :)

    // А именно имя компьютера в сети
    Result := Info^.wki100_computername; // в результат имя компа
    // И где он находиться
    CurrentDomainName := Info^.wki100_langroup;
    // в переменную имя текущего домена
  Except
    on E: Exception do
      begin
        // memo1.Lines.add('Ошибка загрузки списка компьютеров  :'+E.Message);
      end;
  end;
end;

/// ////////////////////////////////////////////////////  Версия ОС (5.0,6.0,6.1 и т.д)
function TfrmDomainInfo.GetCurrentComputerOS(s: PWideChar): String;
var
  Info: PWkstaInfo100;
  Error: DWORD;
begin
  // А для этого мы воспользуемся следующей функцией
  try
    Error := NetWkstaGetInfo((s), 101, @Info);
    /// s имя компьютера
    if Error <> 0 then
      raise Exception.Create(SysErrorMessage(Error));
    Result := (inttostr(Info^.wki100_ver_major) + '.' +
      inttostr(Info^.wki100_ver_minor));
  except
    on E: Exception do
      Memo1.Lines.Add(E.Message);
  end;
end;

/// ////////////////////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.EnumAllTrustedDomains: Boolean;
var
  Tmp, DomainList: PWideChar;
begin
  // Используем недокументированную функцию NetEnumerateTrustedDomains
  // (только не пойму, с какого перепуга она не документирована?)
  // Тут все очень просто, на вход имя контролера домена, ны выход - список доверенных доменов
  Result := NetEnumerateTrustedDomains
    (StringToOleStr(GetDomainController(CurrentDomainName)), @DomainList)
    = NERR_Success;
  // Если вызов функции успешен, то...
  if Result then
    try
      Tmp := DomainList;
      while Length(Tmp) > 0 do
        begin
          // memTrustedDomains.Lines.Add(Tmp); // Банально выводим список на экран
          Tmp := Tmp + Length(Tmp) + 1;
        end;
    finally
      // Не забываем про память
      NetApiBufferFree(DomainList);
    end;
end;

// Данная функция получает информацию о всех рабочих станциях присутствующих в домене
// Вообщето так делать немного не верно, дело в том что рабочие станции могут
// присутствовать в списке не только те, которые завел сисадмин (но для демки сойдет и так)
// =============================================================================
function TfrmDomainInfo.EnumAllWorkStation: Boolean;
var
  Tmp, Info: PNetDisplayMachine;
  i, CurrIndex, EntriesRequest, PreferredMaximumLength, ReturnedEntryCount
    : Cardinal;
  Error: DWORD;
  ProcBar: TprogressBar;
  z: Integer;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  ProcBar.Max := 100;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка всех компьютеров домена '
    + CurrentDomainName;
  try
    CurrIndex := 0;
    repeat
      Info := nil;
      // NetQueryDisplayInformation возвращает информацию только о 100-а записях
      // для того чтобы получить всю информацию используется третий параметр,
      // передаваемый функции, который определяет с какой записи продолжать
      // вывод информации
      EntriesRequest := 100;
      ProcBar.Position := 0;
      z := 0;
      PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayMachine);
      ReturnedEntryCount := 0;
      // Для выполнения функции, в нее нужно передать DNS имя контролера домена
      // (или его IP адрес), с которого мы хочем получить информацию
      // Для получения информации о рабочих станциях используется структура NetDisplayMachine
      // и ее идентификатор 2 (двойка) во втором параметре
      Error := NetQueryDisplayInformation
        (StringToOleStr(GetDomainController(CurrentDomainName)), 2, CurrIndex,
        EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
      // При безошибочном выполнении фунции будет результат либо
      // 1. NERR_Success - все записи возвращены
      // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
      if Error in [NERR_Success, ERROR_MORE_DATA] then
        try
          Tmp := Info;
          // Выводим информацию которую вернула функция в структуру
          for i := 0 to ReturnedEntryCount - 1 do
            begin
              if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
                Ansiuppercase(frmDomainInfo.Caption)) <> 0) and
                (ComboBox2.Items.Count >= 30) then
                break;
              Application.ProcessMessages;
              Inc(z);
              ProcBar.Position := z;
              with ComboBox2.Items do
                begin
                  Add(copy(Tmp^.usri2_name, 1, Length(Tmp^.usri2_name) - 1));
                  // combobox2.ItemsEx[ComboBox2.Items.Count-1].ImageIndex:=0;
                  CurrIndex := Tmp^.usri2_next_index;
                end;
              Inc(Tmp);
            end;
        finally
          // Дабы небыло утечек
          NetApiBufferFree(Info);
        end;
      // Если результат выполнения функции ERROR_MORE_DATA
      // (т.е. есть еще данные) - вызываем функцию повторно
      if pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
        Ansiuppercase(frmDomainInfo.Caption)) <> 0 then
        break;
    until Error in [NERR_Success, ERROR_ACCESS_DENIED];
    // Ну и возвращаем результат всего что мы тут накодили
    Result := Error = NERR_Success;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
end;

/// //////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.EnumAllWorkStationToLit(var comboboxList: TComboBox)
  : Boolean; // все рабочие станции в combobox
var
  Tmp, Info: PNetDisplayMachine;
  i, CurrIndex, EntriesRequest, PreferredMaximumLength, ReturnedEntryCount
    : Cardinal;
  Error: DWORD;
  ProcBar: TprogressBar;
  z: Integer;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  ProcBar.Max := 100;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка всех компьютеров домена '
    + CurrentDomainName;
  try
    comboboxList.Clear;
    CurrIndex := 0;
    repeat
      Info := nil;
      // NetQueryDisplayInformation возвращает информацию только о 100-а записях
      // для того чтобы получить всю информацию используется третий параметр,
      // передаваемый функции, который определяет с какой записи продолжать
      // вывод информации
      EntriesRequest := 100;
      ProcBar.Position := 0;
      z := 0;
      PreferredMaximumLength := EntriesRequest * SizeOf(TNetDisplayMachine);
      ReturnedEntryCount := 0;
      // Для выполнения функции, в нее нужно передать DNS имя контролера домена
      // (или его IP адрес), с которого мы хочем получить информацию
      // Для получения информации о рабочих станциях используется структура NetDisplayMachine
      // и ее идентификатор 2 (двойка) во втором параметре
      Error := NetQueryDisplayInformation
        (StringToOleStr(GetDomainController(CurrentDomainName)), 2, CurrIndex,
        EntriesRequest, PreferredMaximumLength, ReturnedEntryCount, @Info);
      // При безошибочном выполнении фунции будет результат либо
      // 1. NERR_Success - все записи возвращены
      // 2. ERROR_MORE_DATA - записи возвращены, но остались еще и нужно вызывать функцию повторно
      if Error in [NERR_Success, ERROR_MORE_DATA] then
        try
          Tmp := Info;
          // Выводим информацию которую вернула функция в структуру
          for i := 0 to ReturnedEntryCount - 1 do
            begin
              if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
                Ansiuppercase(frmDomainInfo.Caption)) <> 0) and (i >= 30) then
                break;
              Application.ProcessMessages;
              Inc(z);
              ProcBar.Position := z;
              with comboboxList.Items do
                begin
                  Add(copy(Tmp^.usri2_name, 1, Length(Tmp^.usri2_name) - 1));
                  CurrIndex := Tmp^.usri2_next_index;
                end;
              Inc(Tmp);
            end;
        finally
          // Дабы небыло утечек
          NetApiBufferFree(Info);
        end;
      if pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
        Ansiuppercase(frmDomainInfo.Caption)) <> 0 then
        break;
      // Если результат выполнения функции ERROR_MORE_DATA
      // (т.е. есть еще данные) - вызываем функцию повторно
    until Error in [NERR_Success, ERROR_ACCESS_DENIED];
    // Ну и возвращаем результат всего что мы тут накодили
    Result := Error = NERR_Success;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
  Result := true;
end;

/// ///////////////////////////////////////////////////////////////////////////
// Довольно простая функция, возвращает только имена компьютеров принадлезжащих группе
// =============================================================================
function TfrmDomainInfo.GetAllGroupUsers(const GroupName: String): Boolean;
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries, resumehandle: DWORD;
  i: Integer;
  ProcBar: TprogressBar;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка компьютеров группы ' +
    GroupName;
  try
    try
      // На вход подается список который мы будем заполнять
      // lbInfo.Items.Clear;
      ListView8.Clear;
      // Обязательная инициализация
      resumehandle := 0;
      prefmaxlen := DWORD(-1);
      // Выполняем
      Result := NetGroupGetUsers
        (StringToOleStr(GetDomainController(CurrentDomainName)),
        StringToOleStr(GroupName), 0, Pointer(Info), prefmaxlen, entriesread,
        totalentries, @resumehandle) = NERR_Success;
      // Смотрим результат...
      if Result then
        try
          Tmp := Info;
          ProcBar.Max := entriesread;
          for i := 0 to entriesread - 1 do
            begin
              if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
                Ansiuppercase(frmDomainInfo.Caption)) <> 0) and
                (ListView8.Items.Count >= 30) then
                break;
              Application.ProcessMessages;
              ProcBar.Position := i;
              // Банально выводим результат из структуры
              if pos('$', (Tmp^.grui0_name)) <> 0 then
                with ListView8.Items.Add do
                  begin
                    Caption := '';
                    SubItems.Add(copy(Tmp^.grui0_name, 1,
                      Length(Tmp^.grui0_name) - 1));
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                  end;
              Inc(Tmp);
            end;
        finally
          // Не забываем, ибо может быть склероз :)
          NetApiBufferFree(Info);
        end;
      StatusBar1.Panels[0].text := 'Всего ПК: ' +
        inttostr(ListView8.Items.Count);
      PCCheck := 0;
      /// выделено компов
      StatusBar1.Panels[2].text := 'Выбрано ПК: ' + '0';
      PcInLan := 0;
      /// Компов в сети
      StatusBar1.Panels[1].text := 'ПК в сети : ' + inttostr(PcInLan);
    except
      begin
        Memo1.Lines.Add
          ('Ограничение загрузки списка компьютеров. Возможно у Вас нет лицензии??? ');
      end;
    end;
  finally
    ProcBar.Free;
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
  end;
end;

/// ////////////////////////////////////////////////////////////////////////////////////
// добавляет список компьютеров выбранной группы в Listview в разделе Компьютера и в Combobox в разделе Управление
function TfrmDomainInfo.GetPCForGroupToLisViewComboBox(const GroupName
  : String): Boolean;
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries, resumehandle: DWORD;
  i: Integer;
  ProcBar: TprogressBar;
begin
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  StatusBarLoadList.Panels[2].text := 'Загрузка списка компьютеров группы ' +
    GroupName;
  try
    try
      // На вход подается группа безопасности
      ListView8.Clear;
      ComboBox2.Clear;
      // Обязательная инициализация
      resumehandle := 0;
      prefmaxlen := DWORD(-1);
      // Выполняем
      Result := NetGroupGetUsers
        (StringToOleStr(GetDomainController(CurrentDomainName)),
        StringToOleStr(GroupName), 0, Pointer(Info), prefmaxlen, entriesread,
        totalentries, @resumehandle) = NERR_Success;
      // Смотрим результат...
      if Result then
        try
          Tmp := Info;
          ProcBar.Max := entriesread;
          for i := 0 to entriesread - 1 do
            begin
              if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),
                Ansiuppercase(frmDomainInfo.Caption)) <> 0) and
                (ListView8.Items.Count >= 30) then
                break;
              Application.ProcessMessages;
              ProcBar.Position := i;
              // Банально выводим результат из структуры
              if pos('$', (Tmp^.grui0_name)) <> 0 then
                Begin
                  with ListView8.Items.Add do
                    begin
                      Caption := '';
                      SubItems.Add(copy(Tmp^.grui0_name, 1,
                        Length(Tmp^.grui0_name) - 1));
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      SubItems.Add('');
                      // SubItems.Add(Tmp^.grui0_name);
                    end;
                  ComboBox2.Items.Add(copy(Tmp^.grui0_name, 1,
                    Length(Tmp^.grui0_name) - 1));
                  // ComboBox2.ItemsEx[ComboBox2.Items.Count-1].ImageIndex:=0;
                End;
              Inc(Tmp);
            end;
        finally
          // Не забываем, ибо может быть склероз :)
          NetApiBufferFree(Info);
        end;
      StatusBar1.Panels[0].text := 'Всего ПК: ' +
        inttostr(ListView8.Items.Count);
      PCCheck := 0;
      /// выделено компов
      StatusBar1.Panels[2].text := 'Выбрано ПК: ' + '0';
      PcInLan := 0;
      /// Компов в сети
      StatusBar1.Panels[1].text := 'ПК в сети : ' + inttostr(PcInLan);
    except
      on E: Exception do
        begin
          Memo1.Lines.Add
            ('Ограничение загрузки списка компьютеров. Возможно у Вас нет лицензии??? ');
        end;
    end;
  finally
    StatusBarLoadList.Panels[2].text := '';
    StatusBarLoadList.Visible := false;
    ProcBar.Free;
  end;
end;

/// //////////////////////////////////////////////////////////////////////////////////////////////////////
/// выдает список компьютеров принадлежащих группе в Stringlist
function TfrmDomainInfo.GetAllGroupPC(const GroupName: String): TstringList;
var
  Tmp, Info: PGroupUsersInfo0;
  prefmaxlen, entriesread, totalentries, resumehandle: DWORD;
  i: Integer;
  Res: Boolean;
begin
  try
    Result := TstringList.Create;
    // На вход подается список который мы будем заполнять
    // lbInfo.Items.Clear;
    // Обязательная инициализация
    resumehandle := 0;
    prefmaxlen := DWORD(-1);
    // Выполняем
    Res := NetGroupGetUsers
      (StringToOleStr(GetDomainController(CurrentDomainName)),
      StringToOleStr(GroupName), 0, Pointer(Info), prefmaxlen, entriesread,
      totalentries, @resumehandle) = NERR_Success;
    // Смотрим результат...
    if Res then
      try
        Tmp := Info;
        for i := 0 to entriesread - 1 do
          begin
            if pos('$', (Tmp^.grui0_name)) <> 0 then
            // Банально выводим результат из структуры
              Result.Add(copy(Tmp^.grui0_name, 1, Length(Tmp^.grui0_name) - 1));
            Inc(Tmp);
          end;
      finally
        // Не забываем, ибо может быть склероз :)
        NetApiBufferFree(Info);
      end;
  except
    on E: Exception do
      begin
        Memo1.Lines.Add
          ('Ограничение загрузки списка компьютеров. Возможно у Вас нет лицензии??? ');
      end;
  end;
end;

/// //////////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.readinfoforpcDB;
var
  i,z: Integer;
  ProcBar: TprogressBar;
  IniSettings:TMeminiFile;
  Rendering:boolean;
begin
  IniSettings:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  try
  Rendering:=IniSettings.Readbool('Scan','Rendering',false); // обновлять или нет listview при блокировке пользователя
  finally
  IniSettings.Free;
  end;
  StatusBarLoadList.Visible := true;
  ProcBar := TprogressBar.Create(frmDomainInfo);
  ProcBar.Parent := StatusBarLoadList;
  ProcBar.height := StatusBarLoadList.height;
  ProcBar.left := StatusBarLoadList.Panels[0].width;
  ProcBar.width := StatusBarLoadList.Panels[1].width;
  ProcBar.Step := 1;
  ProcBar.Position := 0;
  ProcBar.Max := ListView8.Items.Count;
  StatusBarLoadList.Panels[2].text := 'Загрузка данных';
  try
    if (ListView8.Items.Count <> 0) then
      begin
        for i := 0 to ListView8.Items.Count - 1 do
          begin
            try
              Application.ProcessMessages;
              ProcBar.Position := i;
              Datam.FDQueryLoadListPC.SQL.Clear;
              Datam.FDQueryLoadListPC.SQL.text :=
                'SELECT * FROM MAIN_PC WHERE PC_NAME=''' + ListView8.Items[i].SubItems[0] + '''';
              Datam.FDQueryLoadListPC.Open;
              With ListView8.Items[i] do
                begin
                  if not OutForPing then // если сканирование остановлено то обновляем image
                  begin
                  if Datam.FDQueryLoadListPC.FieldByName('EXPT_PC').Value <> null then
                   if TryStrToInt(Datam.FDQueryLoadListPC.FieldByName('EXPT_PC').AsString,z) then
                   begin
                   if z<>12 then z:=0;  //
                    ImageIndex:=z;
                   end
                    else ImageIndex:= 0;
                  end
                  else // иначе запущено сканирование, то image обновляем если только Rendering=false
                   begin
                    if not Rendering then // если в настройках указано не обновлять Listview при блокировке пользователя то данные читаем из базы
                    begin
                      if Datam.FDQueryLoadListPC.FieldByName('EXPT_PC').Value <> null then
                     if TryStrToInt(Datam.FDQueryLoadListPC.FieldByName('EXPT_PC').AsString,z) then
                     begin
                     if z=1 then z:=0; // для предыдущих версии базы, если столо 1 то ставим 0
                     ImageIndex:=z;
                     end
                     else ImageIndex:= 0;
                    end;
                   end;

                  if Datam.FDQueryLoadListPC.FieldByName('ANSWER_MAC').Value <> null then
                    SubItems[2] := Datam.FDQueryLoadListPC.FieldByName('ANSWER_MAC').AsString
                  else SubItems[2] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('PC_OS').Value <> null then
                    SubItems[3] := Datam.FDQueryLoadListPC.FieldByName('PC_OS').AsString
                  else SubItems[3] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('CUR_USER_NAME').Value<> null then
                    SubItems[4] := Datam.FDQueryLoadListPC.FieldByName('CUR_USER_NAME').AsString
                  else SubItems[4] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('CUR_DOMAIN').Value <> null then
                    SubItems[5] :=Datam.FDQueryLoadListPC.FieldByName('CUR_DOMAIN').AsString
                  else SubItems[5] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('OTHER_NAME').Value <> null then
                    SubItems[6] :=Datam.FDQueryLoadListPC.FieldByName('OTHER_NAME').AsString
                  else SubItems[6] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('CUR_IP_ADRS').Value <> null then
                    SubItems[7] :=Datam.FDQueryLoadListPC.FieldByName('CUR_IP_ADRS').AsString
                  else SubItems[7] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('HDD_MY_SMART').Value<> null then
                    SubItems[8] :=Datam.FDQueryLoadListPC.FieldByName('HDD_MY_SMART').AsString
                  else SubItems[8] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('URDM_CLIENT').Value <> null then
                    SubItems[9]:=Datam.FDQueryLoadListPC.FieldByName('URDM_CLIENT').AsString
                  else SubItems[9] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('STATWINLIC').Value <> null then
                    SubItems[10] :=Datam.FDQueryLoadListPC.FieldByName('STATWINLIC').AsString
                  else SubItems[10] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('SSTATOFLIC').Value <> null then
                    SubItems[11] :=Datam.FDQueryLoadListPC.FieldByName('SSTATOFLIC').AsString
                  else SubItems[11] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('ANTIVIRUS_PRODUCT').Value <> null then
                    SubItems[12] :=Datam.FDQueryLoadListPC.FieldByName('ANTIVIRUS_PRODUCT').AsString
                  else SubItems[12] := ('');

                  if Datam.FDQueryLoadListPC.FieldByName('ANTIVIRUS_STATUS').Value <> null then
                    SubItemImages[12] :=Datam.FDQueryLoadListPC.FieldByName('ANTIVIRUS_STATUS').AsInteger
                  else SubItemImages[12] := -1;
                end;
              Datam.FDQueryLoadListPC.SQL.Clear;
              Datam.FDQueryLoadListPC.Close;
            except
              Memo1.Lines.Add('Ошибка загрузки данных ' + ListView8.Items[i]
                .SubItems[0]);
            end;
          end;
      end;
    Datam.FDQueryLoadListPC.SQL.Clear;
    Datam.FDQueryLoadListPC.Close;
  finally
    StatusBarLoadList.Panels[2].text := '';
    ProcBar.Free;
    StatusBarLoadList.Visible := false;
  end;
end;

// меняет ширину выпадающего списка комбобокса ----------------------------------
procedure DropDownWidth(Sender: TObject);
var
  CBox: TComboBox;
  width: Integer;
  i, TextLen: Longint;
  lf: LOGFONT;
  f: HFONT;
begin
  CBox := (Sender as TComboBox);
  width := CBox.width;
  FillChar(lf, SizeOf(lf), 0);
  StrPCopy(lf.lfFaceName, CBox.Font.Name);
  lf.lfHeight := CBox.Font.height;
  lf.lfWeight := FW_NORMAL;
  if fsBold in CBox.Font.Style then
    lf.lfWeight := lf.lfWeight or FW_BOLD;

  f := CreateFontIndirect(lf);
  if (f <> 0) then
    try
      CBox.Canvas.Handle := GetDC(CBox.Handle);
      SelectObject(CBox.Canvas.Handle, f);
      try
        for i := 0 to CBox.Items.Count - 1 do
          begin
            TextLen := CBox.Canvas.TextWidth(CBox.Items[i]);
            if CBox.Items.Count - 1 > CBox.DropDownCount then
              begin
                if TextLen > width - 25 then
                  width := TextLen + 25;
              end
            else if CBox.Items.Count - 1 <= CBox.DropDownCount then
              begin
                if TextLen > width - 5 then
                  width := TextLen + 8;
              end;
          end;
      finally
        ReleaseDC(CBox.Handle, CBox.Canvas.Handle);
      end;
    finally
      DeleteObject(f);
    end;
  SendMessage(CBox.Handle, CB_SETDROPPEDWIDTH, width, 0);
end;

/// ///////////////////////////////////////////////////////////////////////////////
function GridSelectAll(Grid: TDBGrid): Longint; // выделить все строки в DBGrid
begin
  if Grid.DataSource.DataSet.RecordCount = 0 then
    exit;
  Result := 0;
  Grid.SelectedRows.Clear;
  with Grid.DataSource.DataSet do
    begin
      First;
      DisableControls;
      try
        while not EOF do
          begin
            Grid.SelectedRows.CurrentRowSelected := true;
            Inc(Result);
            Next;
          end;
      finally
        EnableControls;
      end;
    end;
end;

/// ///////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.refreshinfoPC(NamePC, user, pass: string): bool;
begin
  try
    if InventConf then
      begin
        showmessage('Дождитесь завершения текущей инвентаризации!');
        exit;
      end;
    if not ping(NamePC) then
      exit;
    ListPCConf := TstringList.Create;
    ListPCConf.Add(NamePC);
    MyUser := user;
    MyPasswd := pass;
    InventConf := true;
    ThrInvPC := MyInventoryPC.InventoryConfig.Create(true);
    ThrInvPC.FreeOnTerminate := true;
    ThrInvPC.Start;
    Result := true;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка: ' + E.Message);
        // frmDomainInfo.Log_write('error',datetostr(date)+'/'+timetostr(time)+ ' Start refresh infoPC in Property PC - ' +e.Message);
        Result := false;
      end;
  end;
end;

/// //////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.Log_write(level: Integer; fname, text: string): string;
var
  f: TstringList;
  /// функция записи в лог файл
begin
  try
    if not DirectoryExists('log') then
      CreateDir('log');
    f := TstringList.Create;
    try
      if FileExists(ExtractFilePath(Application.ExeName) + 'log\' + fname +
        '.log') then
        f.LoadFromFile(ExtractFilePath(Application.ExeName) + 'log\' + fname
          + '.log');
      case level of
        0:
          f.Insert(0, DateTimeToStr(Now) + chr(9) + 'Info' + StringOfChar(' ',
            19) + text);
        1:
          f.Insert(0, DateTimeToStr(Now) + chr(9) + 'Warning' +
            StringOfChar(' ', 13) + text);
        2:
          f.Insert(0, DateTimeToStr(Now) + chr(9) + 'Error' + StringOfChar(' ',
            17) + text);
        3:
          f.Insert(0, DateTimeToStr(Now) + chr(9) + 'Critical Error' +
            StringOfChar(' ', 5) + text);
      end;
      while f.Count > 1000 do
        f.Delete(1000);
      f.SaveToFile(ExtractFilePath(Application.ExeName) + 'log\' + fname
        + '.log');
    finally
      f.Destroy;
    end;
  except
    exit;
  end;
end;

/// /////////////////////////////////////////////////////////////////////////////
/// /////////////////////////////////////////////////////////
/// /Функция сортировки в TListView
function CustomSortProc(Item1, Item2: TListItem; ParamSort: Integer)
  : Integer; stdcall;
begin
  if ParamSort = 0 then
    Result := CompareText(Item1.Caption, Item2.Caption)
  else if Item1.SubItems.Count > ParamSort - 1 then
    begin
      if Item2.SubItems.Count > ParamSort - 1 then
        case SortLV8 of
          true:
            Result := CompareText(Item1.SubItems[ParamSort - 1],
              Item2.SubItems[ParamSort - 1]);
          false:
            Result := CompareText(Item2.SubItems[ParamSort - 1],
              Item1.SubItems[ParamSort - 1]);
        end
      else
        Result := 1;
    end
  else
    Result := -1;
end;

function SortThirdSubItemAsDate(Item1, Item2: TListItem; ParamSort: Integer)
  : Integer; stdcall;
begin
  Result := 0;
  if StrToDate(Item1.SubItems[4]) > StrToDate(Item2.SubItems[4]) then
    Result := ParamSort
  else if StrToDate(Item1.SubItems[4]) < StrToDate(Item2.SubItems[4]) then
    Result := -ParamSort;
end;

function SortThirdSubItemAsDate3(Item1, Item2: TListItem; ParamSort: Integer)
  : Integer; stdcall;
var
  date1, date2: TdateTime;
begin
  Result := 0;
  if (trystrtodate(Item1.SubItems[3], date1) and trystrtodate(Item2.SubItems[3],
    date2)) then
    begin
      if date1 > date2 then
        Result := ParamSort
      else if date1 < date2 then
        Result := -ParamSort;
    end;
end;

function SortThirdSubItemAsDate2(Item1, Item2: TListItem; ParamSort: Integer)
  : Integer; stdcall;
begin
  Result := 0;
  if StrToDate(Item1.SubItems[2]) > StrToDate(Item2.SubItems[2]) then
    Result := ParamSort
  else if StrToDate(Item1.SubItems[2]) < StrToDate(Item2.SubItems[2]) then
    Result := -ParamSort;
end;

/// ////////////////////////////////////////////////////////////////
function TfrmDomainInfo.PutInvNumberToDataBase(pcname: string;
  NewDescription: string): string;
  begin
  try
  DataM.FDTransactionWriteEXPT.StartTransaction;
  DataM.FDQueryWriteEXPT_PC.SQL.Clear;
  DataM.FDQueryWriteEXPT_PC.SQL.Text:='UPDATE CONFIG_PC SET INV_NUMBER=:a WHERE PC_NAME='''+pcname+'''';
  DataM.FDQueryWriteEXPT_PC.ParamByName('a').AsString:=NewDescription;
  DataM.FDQueryWriteEXPT_PC.ExecSQL;
  DataM.FDQueryWriteEXPT_PC.SQL.Clear;
  DataM.FDQueryWriteEXPT_PC.Close;
  DataM.FDTransactionWriteEXPT.Commit;
  Memo1.Lines.Add('Обновление инвентарного номера в базе произведено успешно')
  except on E: Exception do
  begin
  DataM.FDTransactionWriteEXPT.Rollback;
  Memo1.Lines.Add('Ошибка обновления инвентарного номера в базе :'+e.Message)
  end;
  end;
  end;
/// ///////////////////////////////////////////////////////////////////////////////
Function TfrmDomainInfo.PutInvNumber(pcname: string;
  NewDescription: string): string;
/// функция установки инвентарного номера
var
  FWMIService: OleVariant;
  FSWbemLocator: OleVariant;
  FWbemObject: OleVariant;
Begin
  try
    OleInitialize(nil);
    /// //
    // FWMIService   := GetObject('winmgmts:{impersonationlevel=Impersonate, authenticationLevel=pktPrivacy}!\\'+MyPS+'\root\CIMV2');
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer(ComboBox2.text, 'root\CIMV2',
      LabeledEdit1.text, LabeledEdit2.text, '', '', 128);
    FWMIService.Security_.impersonationlevel := 3;
    FWMIService.Security_.authenticationLevel := 6;
    // 6 или WbemAuthenticationLevelPktPrivacy
    FWbemObject := FWMIService.get('Win32_OperatingSystem').SpawnInstance_;
    FWbemObject.Description := NewDescription;
    Result := FWbemObject.put_();
  except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка "' + E.Message + '"');
        VariantClear(FWbemObject);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
      end;
  end;
  if Result <> '' then
    Memo1.Lines.Add(pcname + ' Инв № ' + NewDescription +
      ' успешно установлен');
  FWbemObject := Unassigned;
  VariantClear(FWbemObject);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
End;

/// ////////////////////////////////////////////////////////////////////////////////
/// работа с базой данных
function kolitem(s: string): TStringArray;
/// // функция создания массива данных
var
  i, z: Integer;
begin
  z := 0;
  for i := 1 to Length(s) do
    begin
      if (s[i] = '/') and (s[i + 1] = '*') and (s[i + 2] = '*') then
        Inc(z);
      /// количество вхождений '/'
    end;
  SetLength(Result, z);
  /// количество элементов в массиве
  for i := 0 to z - 1 do
  /// заполняем массив
    begin
      if copy(s, 1, pos('/', s) - 1) <> '' then
        begin
          Result[i] := copy(s, 1, pos('/', s) - 1);
        end;
      Delete(s, 1, pos('/**/', s) + 3);
    end;
end;

Function SummaDDR(mas:TStringArray):string; // расчет общего объема Оперативной памяти
var
i,z,s:integer;
begin
try
z:=0;s:=0;
 for i := 0 to Length(mas) - 1 do
 begin
  if TryStrToInt(mas[i],z) then
   s:=s+z;
 end;
 result:=inttostr(s);
except
result:='';
end;
end;

function kolitemSoft(s: string): TStringArray;
/// // функция создания массива данных
var
  i, z: Integer;
begin
  z := 0;
  for i := 1 to Length(s) do
    begin
      if s[i] = ',' then
        Inc(z);
      /// количество вхождений '/'
    end;
  SetLength(Result, z);
  /// количество элементов в массиве
  for i := 0 to z - 1 do
  /// заполняем массив
    begin
      if copy(s, 1, pos(',', s) - 1) <> '' then
        begin
          Result[i] := copy(s, 1, pos(',', s) - 1);
        end;
      Delete(s, 1, pos(',', s));
    end;
end;

/// /////////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.killscancomp: boolean;
begin
  try
     if not SolveExitInvScan then // если запущено сканирование то  останавливаем
       begin
       Memo1.Lines.Add('Остановка сканирования сети, ожидайте...');
       OutForPing := false;
       while not SolveExitInvScan do Application.ProcessMessages;  /// ожидать Остановки инвентаризации
        try
        freeandnil(SelectedPCPing);
        except on E: Exception do Memo1.Lines.add('Kill thread scan '+e.Message);
        end;
        SpeedButton39.Visible := true;
        SpeedButton40.Visible := false;
        Memo1.Lines.Add('Остановка сканирования сети завершена');
        result:=true;
       end
      else
        begin
        SpeedButton39.Visible := true;
        SpeedButton40.Visible := false;
        OutForPing := false;
        result:=true;
        Memo1.Lines.Add('Остановка сканирования сети завершена');
        end;

  except
    on E: Exception do
      begin
        Result := false;
        Memo1.Lines.Add('Ошибка остановки сканирования "' + E.Message + '"');
      end;
  end;
end;


function TfrmDomainInfo.KillscanSelectGroup: bool;
begin
  try
    if OutForPing then // если скарирование запущено, пытаемся остановить
      begin
         if not SolveExitInvScan then // если запущено сканирование то  останавливаем
         begin
         Memo1.Lines.Add('Остановка сканирования сети, ожидайте...');
         OutForPing := false;
         while not SolveExitInvScan do Application.ProcessMessages;  /// ожидать Остановки инвентаризации
          try
          freeandnil(SelectedPCPing);
          except on E: Exception do Memo1.Lines.add('Kill thread scan '+e.Message);
          end;
          SpeedButton39.Visible := true;
          SpeedButton40.Visible := false;
          Memo1.Lines.Add('Остановка сканирования сети завершена');
          result:=true;
         end
        else
          begin
          SpeedButton39.Visible := true;
          SpeedButton40.Visible := false;
          OutForPing := false;
          result:=true;
          Memo1.Lines.Add('Остановка сканирования сети завершена');
          end;
        end
    else
      Result := true; // если скарирование не запущено то сразу истина в ответ
  except
    on E: Exception do
      begin
        Result := false;
        Memo1.Lines.Add('Ошибка остановки сканирования "' + E.Message + '"');
      end;
  end;
end;


/// ///////////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.createviewerMain;
/// / создание вьювера в основном окне программы
begin
  if not Assigned(MRDUViewerMain) then
    begin
      MRDUViewerMain := TRDPViewer.Create(Panel1);
      MRDUViewerMain.Parent := Panel1;
      MRDUViewerMain.Name := 'RDPViewerMain';
      MRDUViewerMain.Align := alClient;
      MRDUViewerMain.OnAttendeeConnected := RDPViewer1AttendeeConnected;
      MRDUViewerMain.OnAttendeeDisconnected := RDPViewer1AttendeeDisconnected;
      MRDUViewerMain.OnAttendeeUpdate := RDPViewer1AttendeeUpdate;
      MRDUViewerMain.OnChannelDataReceived := RDPViewer1ChannelDataReceived;
      MRDUViewerMain.OnChannelDataSent := RDPViewer1ChannelDataSent;
      MRDUViewerMain.OnConnectionAuthenticated :=
        RDPViewer1ConnectionAuthenticated;
      MRDUViewerMain.OnConnectionEstablished := RDPViewer1ConnectionEstablished;
      MRDUViewerMain.OnConnectionFailed := RDPViewer1ConnectionFailed;
      MRDUViewerMain.OnConnectionTerminated := RDPViewer1ConnectionTerminated;
      MRDUViewerMain.OnControlLevelChangeRequest :=
        RDPViewer1ControlLevelChangeRequest;
      MRDUViewerMain.OnFocusReleased := RDPViewer1FocusReleased;
      MRDUViewerMain.OnGraphicsStreamPaused := RDPViewer1GraphicsStreamPaused;
      MRDUViewerMain.OnGraphicsStreamResumed := RDPViewer1GraphicsStreamResumed;
      MRDUViewerMain.OnSharedDesktopSettingsChanged :=
        RDPViewer1SharedDesktopSettingsChanged;
      MRDUViewerMain.OnSharedRectChanged := RDPViewer1SharedRectChanged;
      MRDUViewerMain.OnStartDrag := RDPViewer1StartDrag;
      MRDUViewerMain.OnWindowClose := RDPViewer1WindowClose;
      MRDUViewerMain.OnWindowOpen := RDPViewer1WindowOpen;
      MRDUViewerMain.OnWindowUpdate := RDPViewer1WindowUpdate;
    end;
end;

/// ///////////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.pcResChange(Sender: TObject);
begin
  if ListView1.Items.Count <> 0 then
    begin
      if pcRes.ActivePage = TabSheet2 then
        ListView1.MultiSelect := false;
      ListView1.Items[0].Selected := true;
    end;
  case pcRes.ActivePageIndex of
    0:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Caption := 'Процессы';
      end;
    1:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Caption := 'Службы';
      end;
    2:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Caption := 'Автозагрузка';
      end;
    3:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := false;
        SpeedButton3.Visible := true;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
      end;
    4:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'Все программы';
      end;
    5:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton4.Visible := true;
      end;
    6:
      begin
        SpeedButton1.Visible := false;
        SpeedButton2.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := true;
      end;
    7:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'Профили';
      end;
    8:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'Устройства';
      end;
    9:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'Network';
      end;
    10:
      begin
        SpeedButton1.Visible := true;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := false;
      end;
    11:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := false;
      end;
    12:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'HotFix';
      end;
    13:
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := true;
        SpeedButton2.Caption := 'Сетевые ресурсы';
      end;
    14:
    /// // вкладка журналы событий
      begin
        SpeedButton1.Visible := false;
        SpeedButton3.Visible := false;
        SpeedButton4.Visible := false;
        SpeedButton5.Visible := false;
        SpeedButton2.Visible := false;
      end;

  end;
end;



function TfrmDomainInfo.ping(s:string):bool;
var
ipadr:string;
available:boolean;
begin
case PingType of
1:
begin
if PingIdIcmp(s,pingtimeout) then
  begin
  result:=true;
  available:=true;
  end
else
  begin
  result:=false;
  available:=false;
  Memo1.Lines.Add('----------------------------------------------');
  Memo1.Lines.Add('Компьютер '+s+ ' не доступен.');
  showmessage('Компьютер '+s+' не доступен.');
  end;
end;
2:
begin
if PingGetaddrinfo(ipadr,s,pingtimeout) then
  begin
  result:=true;
  Memo1.Lines.Add('----------------------------------------------');
  Memo1.Lines.Add('Компьютер  - '+s);
  Memo1.Lines.Add('IP хоста  - '+ipadr);
  available:=true;
  end
else
  begin
  result:=false;
  Memo1.Lines.Add('----------------------------------------------');
  Memo1.Lines.Add('Компьютер '+s+ ' не доступен.');
  showmessage('Компьютер '+s+' не доступен.');
  available:=false;
  end;
end;
3:
begin
 if PingGetHostByName(ipadr,s,pingtimeout) then
  begin
  result:=true;
  Memo1.Lines.Add('----------------------------------------------');
  Memo1.Lines.Add('Компьютер  - '+s);
  Memo1.Lines.Add('IP хоста  - '+ipadr);
  available:=true;
  end
else
  begin
  result:=false;
  Memo1.Lines.Add('----------------------------------------------');
  Memo1.Lines.Add('Компьютер '+s+ ' не доступен.');
  showmessage('Компьютер '+s+' не доступен.');
  available:=false;
  end;
end;

END;

try
if available then
begin
 if (ComboBox2.ItemIndex <> -1) and (ComboBox2.text = s) then
  if (ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex <> 6) then ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex := 4;
  if (ListView8.ItemIndex <> -1) then
  if (ListView8.Items[ListView8.ItemIndex].SubItems[0] = s)
   and (ListView8.Items[ListView8.ItemIndex].ImageIndex <> 6) then  ListView8.Items[ListView8.ItemIndex].ImageIndex := 4;
end
else
begin
   if (ComboBox2.ItemIndex <> -1) and (ComboBox2.text = s) then ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex := 5;
  if (ListView8.ItemIndex <> -1) then
  if ListView8.Items[ListView8.ItemIndex].SubItems[0] = s then ListView8.Items[ListView8.ItemIndex].ImageIndex := 5;
end;

except on E: Exception do
Memo1.Lines.Add('Ошибка обновления статуса компьютера '+e.Message);
end;
end;

// проверка доступности WMI
function TfrmDomainInfo.ConnectWMI(NamePC, user, pass: string): Boolean;
var
  FSWbemLocatorPing: OleVariant;
  FWMIServicePing: OleVariant;
begin
  Result := false;
  try
    try
      FSWbemLocatorPing := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIServicePing := FSWbemLocatorPing.ConnectServer(NamePC, 'root\CIMV2',user, pass, '', '', 128);
      FWMIServicePing.Security_.impersonationlevel := 3;
      FWMIServicePing.Security_.authenticationLevel := 6;
      if (ComboBox2.ItemIndex <> -1) and (ComboBox2.text = NamePC) then
        ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex := 4;
      if (ListView8.ItemIndex <> -1) then
        if (ListView8.Items[ListView8.ItemIndex].SubItems[0] = NamePC) then
          ListView8.Items[ListView8.ItemIndex].ImageIndex := 4;
      Memo1.Lines.Add('----------------------------------------------');
      Memo1.Lines.Add(NamePC + ' Сервер RPC доступен');
      Result := true;
    except
      on E: Exception do
        begin
          Memo1.Lines.Add('----------------------------------------------');
          Memo1.Lines.Add(Format(NamePC + ' Ошибка  %s', [E.Message]));
          Result := false;
          if (ComboBox2.ItemIndex <> -1) and (ComboBox2.text = NamePC) then
            ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex := 6;
          if (ListView8.ItemIndex <> -1) then
            if ListView8.Items[ListView8.ItemIndex].SubItems[0] = NamePC then
              ListView8.Items[ListView8.ItemIndex].ImageIndex := 6;
          showmessage(E.Message);
        end;
    end;
  finally
    VariantClear(FWMIServicePing);
    VariantClear(FSWbemLocatorPing);
  end;
end;

procedure TfrmDomainInfo.PopupMenu1Change(Sender: TObject; Source: TMenuItem;
  Rebuild: Boolean);
// Var
// ItemsNew:TMenuItem;
begin
  // ItemsNew := TMenuItem.Create(PopupMenu1);
  // Popupmenu1.Items[1].Add(ItemsNew);
  // Popupmenu1.Items[1].Items[0].Caption:='Пункт 1';
end;

procedure TfrmDomainInfo.PopupMenuItemsClick(Sender: TObject);
var
  MyProcPriority: TThread;
begin
  /// ///////////////////////////////////         попуп меню для приоритета процесса
  with Sender as TMenuItem do
    begin
      case Tag of
        0:
          MyPriority := 256;
          /// /   Realtime (256) Реального времени
        1:
          MyPriority := 128;
          /// High Priority (128) Высокий
        2:
          MyPriority := 32768;
          /// /   Above Normal (32768) Выше среднего
        3:
          MyPriority := 32;
          /// /  Normal (32) Средний
        4:
          MyPriority := 16384;
          /// /   Below Normal (16384) Ниже среднего
        5:
          MyPriority := 64;
          /// // idie (64) Низкий
      end;
      MyPS := ComboBox2.text;
      MyProcPriority := unit4.Priority.Create(true);
      MyProcPriority.FreeOnTerminate := true;
      MyProcPriority.Start;
    end;

end;

procedure TfrmDomainInfo.ListViewallChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  try
    if Item.Index > 30 then
      Item.Delete;
    StatusBar1.Panels[0].text := 'Всего ПК: ' + inttostr(ListView8.Items.Count);
  finally

  end;
end;

procedure TfrmDomainInfo.ListViewAVDblClick(Sender: TObject);
var
  liststr: TstringList;
begin
  if ListViewAV.SelCount = 0 then
    exit;
  liststr := TstringList.Create;
  try
    liststr.Add(ListViewAV.Selected.Caption);
    liststr.Add(ListViewAV.Selected.SubItems[0]);
    liststr.Add(ListViewAV.Selected.SubItems[1]);
    itemLisToMemo(liststr, GroupBox2.Caption, frmDomainInfo);
  finally
    liststr.Free;
  end;
end;

procedure TfrmDomainInfo.ListViewDiskColumnClick(Sender: TObject;
  Column: TListColumn);
/// сортировка дисков
begin
  SortLV8 := not SortLV8;
  ListViewDisk.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.PopupMenu2ItemsClick(Sender: TObject);
var
  Myservice: TThread;
begin
  /// ///////////////////////////////////         попуп меню для типа запуска службы
  if ListView1.Items.Count <> 0 then
  /// /
    begin
      SelectServic := ListView1.Selected.SubItems[0];
      with Sender as TMenuItem do
        begin
          case Tag of
            0:
              TypeRunService := 'Automatic';
              /// /   Автоматический запуск
            1:
              TypeRunService := 'Manual';
              /// Вручную
            2:
              TypeRunService := 'Disabled';
              /// /   Отключена
          end;
          ActionServic := 4;
          MyPS := ComboBox2.text;
          Myservice := unit6.Service.Create(true);
          Myservice.FreeOnTerminate := true;
          Myservice.Start;
        end;
    end;

end;

procedure TfrmDomainInfo.ApplicationEvents1Minimize(Sender: TObject);
begin
  if Traymin then
    begin
      TrayIcon1.Visible := true;
      // Убираем с панели задач
      ShowWindow(Handle, SW_HIDE); // Скрываем программу
      ShowWindow(Application.Handle, SW_HIDE); // Скрываем кнопку с TaskBar'а
      // SetWindowLong(Application.Handle, GWL_EXSTYLE,
      // GetWindowLong(Application.Handle, GWL_EXSTYLE) or (not WS_EX_APPWINDOW));
      if (InventConf) or (InventSoft) then
        TrayIcon1.Animate := true;
    end;

end;

procedure TfrmDomainInfo.Button11Click(Sender: TObject);
begin
  OpenFormSMATR;
end;

procedure TfrmDomainInfo.Button12Click(Sender: TObject);
begin
  OpenRDPForListPC;
end;

procedure TfrmDomainInfo.Button1Click(Sender: TObject);

begin
  // FreeAndNil(MyRDPClient);
  // MyProcent:=unit2.Procent.Create(false);
  // MyProcent.terminate; MyProcPriority
end;

procedure TfrmDomainInfo.Button22Click(Sender: TObject);
// пользователи и группы
begin
  EditUserGroup;
end;

procedure TfrmDomainInfo.Button23Click(Sender: TObject);
begin
  CopyDelFFForListPC;
end;

procedure TfrmDomainInfo.SpeedButton85Click(Sender: TObject); // окно RDP
begin
  NewRDPFormForPC;
end;

procedure TfrmDomainInfo.Button24Click(Sender: TObject); // окно uRDM
begin
  OpenFormForuRDM;
end;

procedure TfrmDomainInfo.Button25Click(Sender: TObject); // открыть проводник
begin
  OpenFormExplorer;
end;

procedure TfrmDomainInfo.Button26Click(Sender: TObject);
begin
  InventoryMicrosoft;
  // процедура запуска инвентаризации продуктов Windows и Office
end;

procedure TfrmDomainInfo.Button27Click(Sender: TObject);
var
  i: Integer;
begin
  if SolveExitInvMicrosoft then exit // если false то поток можно запустить, на кой хуй его останавливать
  else
    begin
      i := MessageDlg('Остановить инвентаризацию? ', mtConfirmation,
        [mbYes, mbCancel], 0);
      if i = IDCancel then exit;
      InventMicrosoft := false;
      // признак остановки инвентаризации Windows и Office
      Memo1.Lines.Add('Остановка инвентаризации  продуктов Windows и Office... ожидайте');
    end;
end;

procedure TfrmDomainInfo.Button28Click(Sender: TObject);
begin
  UpdateWindowsOffice;
end;




function statusAntivirus(AntStatus, AntUpdate: string): Integer;
var
  AntivirusIdStatus: Integer;
begin
  try
    if (AntStatus = 'Включен') and (AntUpdate = 'Ok') then
      AntivirusIdStatus := 18 // статус ОК
    else if (AntStatus = 'Включен') and (AntUpdate = 'Базы устарели') then
      AntivirusIdStatus := 21 // не совсем ОК
    else if (AntStatus = 'Выключен') and (AntUpdate = 'Ok') then
      AntivirusIdStatus := 19 // ваще не гуд
    else if (AntStatus = 'Выключен') and (AntUpdate = 'Базы устарели') then
      AntivirusIdStatus := 20 // хреного
    else
      AntivirusIdStatus := 19; // ваще жопа
    Result := AntivirusIdStatus;
  except
    on E: Exception do
      Result := 19;
  end;
end;

procedure TfrmDomainInfo.UpdateAntivirusProduct;
begin
  LVAntivirus.Items.BeginUpdate;
  try
    try
      LVAntivirus.Clear;
      Datam.FDSelectReadSoft.SQL.Clear;
      Datam.FDSelectReadSoft.SQL.text := 'SELECT * FROM ANTIVIRUSPRODUCT';
      Datam.FDSelectReadSoft.Open;
      Datam.FDSelectReadSoft.First;
      while not Datam.FDSelectReadSoft.EOF do
        begin
          With LVAntivirus.Items.Add do
            begin
              ImageIndex := 0;
              if Datam.FDSelectReadSoft.FieldByName('NAMEPC').Value <> null then
                Caption :=
                  (Datam.FDSelectReadSoft.FieldByName('NAMEPC').AsString)
              else
                Caption := ('');
              if Datam.FDSelectReadSoft.FieldByName('OS_NAME').Value <> null
              then // 0
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('OS_NAME').AsString)
              else
                SubItems.Add('');

              if Datam.FDSelectReadSoft.FieldByName('ANTIVIRUS').Value <> null
              then // 3
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTIVIRUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('ANTIVIRUS_STATUS').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTIVIRUS_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('ANTIVIRUS_UPDATE').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTIVIRUS_UPDATE').AsString)
              else
                SubItems.Add('');
              SubItemImages[1] := statusAntivirus(SubItems[2], SubItems[3]);

              if Datam.FDSelectReadSoft.FieldByName('FIREWALL').Value <> null
              then // 6
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('FIREWALL').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('FIREWALL_STATUS').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('FIREWALL_STATUS').AsString)
              else
                SubItems.Add('');
              SubItemImages[4] := statusAntivirus(SubItems[5], 'Ok');

              if Datam.FDSelectReadSoft.FieldByName('ANTISPYWARE').Value <> null
              then // 8
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTISPYWARE').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('ANTISPYWARE_STATUS').Value
                <> null then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTISPYWARE_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('ANTISPYWARE_UPDATE').Value
                <> null then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('ANTISPYWARE_UPDATE').AsString)
              else
                SubItems.Add('');
              SubItemImages[6] := statusAntivirus(SubItems[7], SubItems[8]);

              if Datam.FDSelectReadSoft.FieldByName('WIN_DEF').Value <> null
              then // 0
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('WIN_DEF').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('WIN_DEF_STATUS').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('WIN_DEF_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('WIN_DEF_UPDATE').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('WIN_DEF_UPDATE').AsString)
              else
                SubItems.Add('');
              if (SubItems[10] <> '') and (SubItems[11] <> '') then
                SubItemImages[9] := statusAntivirus(SubItems[10], SubItems[11]);

            end;
          Datam.FDSelectReadSoft.Next;
        end;
      Datam.FDSelectReadSoft.SQL.Clear;
      Datam.FDSelectReadSoft.Close;
    finally
      LVAntivirus.Items.EndUpdate;
      // StatusInvMicrosoft.Panels[0].Text:='Всего записей - '+inttostr(LVMicrosoft.Items.Count-1);
    end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка обновления Антивирусных продуктов ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.UpdateWindowsOffice;
begin
  LVMicrosoft.Items.BeginUpdate;
  try
    try
      LVMicrosoft.Clear;
      Datam.FDSelectReadSoft.SQL.Clear;
      Datam.FDSelectReadSoft.SQL.text := 'SELECT * FROM MICROSOFT_LIC';
      Datam.FDSelectReadSoft.Open;
      Datam.FDSelectReadSoft.First;
      while not Datam.FDSelectReadSoft.EOF do
        begin
          With LVMicrosoft.Items.Add do
            begin
              Caption := Datam.FDSelectReadSoft.FieldByName('ID').AsString;
              if Datam.FDSelectReadSoft.FieldByName('NAMEPC').Value <> null then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('NAMEPC').AsString)
              else
                SubItems.Add('');
              SubItemImages[0] := 0;
              if Datam.FDSelectReadSoft.FieldByName('WINPRODUCT').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('WINPRODUCT').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('STATWINLIC').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('STATWINLIC').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('DESCRIPT_LIC_WIN').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('DESCRIPT_LIC_WIN').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('KEYWIN').Value <> null then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('KEYWIN').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('TYPE_LIC_WIN').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('TYPE_LIC_WIN').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('SOFFICEPRODUCT').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('SOFFICEPRODUCT').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('SSTATOFLIC').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('SSTATOFLIC').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('DESCRIPT_LIC_OFFICE').Value
                <> null then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('DESCRIPT_LIC_OFFICE').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('SKEYOFC').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('SKEYOFC').AsString)
              else
                SubItems.Add('');
              if Datam.FDSelectReadSoft.FieldByName('TYPE_LIC_OFFICE').Value <> null
              then
                SubItems.Add(Datam.FDSelectReadSoft.FieldByName
                  ('TYPE_LIC_OFFICE').AsString)
              else
                SubItems.Add('');
            end;
          Datam.FDSelectReadSoft.Next;
        end;
      Datam.FDSelectReadSoft.SQL.Clear;
      Datam.FDSelectReadSoft.Close;
    finally
      LVMicrosoft.Items.EndUpdate;
      StatusInvMicrosoft.Panels[0].text := 'Всего записей - ' +
        inttostr(LVMicrosoft.Items.Count - 1);
    end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка обновления продуктов Microsoft ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.ActivationKMS; // Активация KMS
begin
  if ping(ComboBox2.text) then
    FormKMS.showmodal;
end;

procedure TfrmDomainInfo.OpenRestoreForm;
begin
  if ping(ComboBox2.text) then
    FormRestoreWin.showmodal;
end;

procedure TfrmDomainInfo.Button2Click(Sender: TObject); // Новая задача
var
  nametask: string;
begin
  nametask := NewNameTask('');
  if nametask = '' then
    begin
      exit;
    end;
  EditTask.ListView2.Clear; // очистка списка компов
  EditTask.ListView1.Clear; // очистка списка заданий
  EditTask.Caption := nametask;
  EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая
  EditTask.Button8.Enabled := false;
  // кнопка запуска не активна , активируется после сохранения задачи
  if EditTask.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    EditTask.WindowState := wsNormal
  else // иначе показываем
    EditTask.Show;
end;


procedure TfrmDomainInfo.Button3Click(Sender: TObject); // удалить задачу
var
  i, Step: Integer;
  NameTable: string;
  delTab: Boolean;
  FDQueryDeleteTask: TFDQuery;
  TransactionDeleteTask: TFDTransaction;
begin
  if TaskListView.SelCount = 1 then
    begin
      try
        if EditTask.StatusStartStopTask
          (TaskListView.Items[TaskListView.ItemIndex].Caption) then
        // если задача запущена
          begin
            showmessage('Перед удалением остановите задачу');
            exit;
          end;
        delTab := false;
        i := MessageDlg('Вы действительно хотите удалить задачу ' +
          TaskListView.Items[TaskListView.ItemIndex].Caption, mtConfirmation,
          [mbYes, mbCancel], 0);
        if i = IDCancel then
          exit;
        TransactionDeleteTask := TFDTransaction.Create(nil);
        TransactionDeleteTask.Connection := Datam.ConnectionDB;
        TransactionDeleteTask.Options.Isolation := xiReadCommitted;
        /// xiReadCommitted; //xiSnapshot; xiUnspecified
        TransactionDeleteTask.Options.AutoCommit := false;
        TransactionDeleteTask.Options.AutoStart := false;
        TransactionDeleteTask.Options.AutoStop := false;
        FDQueryDeleteTask := TFDQuery.Create(nil);
        FDQueryDeleteTask.Transaction := TransactionDeleteTask;
        FDQueryDeleteTask.Connection := Datam.ConnectionDB;

        TransactionDeleteTask.StartTransaction; // стартуем
        NameTable := EditTask.nametableForDescription
          (TaskListView.Items[TaskListView.ItemIndex].Caption); // имя таблицы
        FDQueryDeleteTask.SQL.Clear;
        FDQueryDeleteTask.SQL.text :=
          'UPDATE TABLE_TASK SET STATUS_TASK=:a, DESCRIPTION_TASK=:b  WHERE NAME_TABLE='''
          + NameTable + '''';
        FDQueryDeleteTask.ParamByName('a').AsString := 'delete';
        FDQueryDeleteTask.ParamByName('b').AsString := '  '; // удаляем описание
        FDQueryDeleteTask.ExecSQL;
        TransactionDeleteTask.commit;

        TransactionDeleteTask.StartTransaction; // стартуем
        FDQueryDeleteTask.SQL.Clear;
        FDQueryDeleteTask.SQL.text := 'DROP TABLE ' + NameTable;
        FDQueryDeleteTask.ExecSQL;
        TransactionDeleteTask.commit;
        delTab := true;
        if delTab then
          begin
            TransactionDeleteTask.StartTransaction; // стартуем
            FDQueryDeleteTask.SQL.Clear;
            FDQueryDeleteTask.SQL.text :=
              'DELETE FROM TABLE_TASK WHERE NAME_TABLE=''' + NameTable + '''';
            FDQueryDeleteTask.ExecSQL;
            TaskListView.Items[TaskListView.ItemIndex].Delete;
            TransactionDeleteTask.commit;
          end;
        FDQueryDeleteTask.SQL.Clear; // очистить
        FDQueryDeleteTask.Close;
        /// закрыть нах
        FDQueryDeleteTask.Free;
        TransactionDeleteTask.Free;
      except
        on E: Exception do
          begin
            if Assigned(FDQueryDeleteTask) then
              FDQueryDeleteTask.Free;
            if Assigned(TransactionDeleteTask) then
              begin
                TransactionDeleteTask.Rollback;
                TransactionDeleteTask.Free;
              end;
            Button4.Click; // обновляем список задач
            Log_write(1, 'TASK', timetostr(Now) + ' Ошибка удаления задачи - ' +
              E.Message);
            Memo1.Lines.Add('Ошибка удаления задачи, повторите операцию позже');
          end;
      end;
    end;
end;

procedure TfrmDomainInfo.TabSheet24Show(Sender: TObject);
begin
  Button4.Click;
end;

procedure TfrmDomainInfo.Button4Click(Sender: TObject);
// обновление списка задач
var
  FDQueryRead: TFDQuery;
  TransactionRead: TFDTransaction;
begin
  try
    TransactionRead := TFDTransaction.Create(nil);
    TransactionRead.Connection := Datam.ConnectionDB;
    TransactionRead.Options.Isolation := xiReadCommitted;
    /// xiReadCommitted; //xiSnapshot;
    TransactionRead.Options.AutoCommit := false;
    TransactionRead.Options.AutoStart := false;
    TransactionRead.Options.AutoStop := false;
    TransactionRead.Options.ReadOnly := true; // только чтение
    FDQueryRead := TFDQuery.Create(nil);
    FDQueryRead.Transaction := TransactionRead;
    FDQueryRead.Connection := Datam.ConnectionDB;
    TransactionRead.StartTransaction; // стартуем
    FDQueryRead.SQL.text :=
      'SELECT * FROM TABLE_TASK WHERE STATUS_TASK<>''delete''';
    FDQueryRead.Open;
    TaskListView.Clear;
    while not FDQueryRead.EOF do
      begin
        with TaskListView.Items.Add do
          begin
            if strtobool(FDQueryRead.FieldByName('START_STOP').AsString) then
              ImageIndex := 0
            else
              ImageIndex := 1;
            Caption := FDQueryRead.FieldByName('DESCRIPTION_TASK').AsString;
            // описание задачи
            if (not FDQueryRead.FieldByName('WORKS_THREAD').AsBoolean) and
              (FDQueryRead.FieldByName('STATUS_TASK')
              .AsString = 'Остановка задачи') then
              SubItems.Add('Остановлена')
              // если поток утановил заначение  false то задача остановлена или завершена
            else
              SubItems.Add(FDQueryRead.FieldByName('STATUS_TASK').AsString);
            // статус задачи (выполняется/остановлена/остановка)

            SubItems.Add(FDQueryRead.FieldByName('CURRENT_TASK').AsString);
            // Текущее задание
            SubItems.Add(FDQueryRead.FieldByName('PC_RUN').AsString);
            // на каком компьютере сейчас выполняется
            SubItems.Add(FDQueryRead.FieldByName('COUNT_PC').AsString);
            // всего компьютеров в задаче
          end;
        FDQueryRead.Next;
      end;
    FDQueryRead.SQL.Clear;
    FDQueryRead.Close;
    TransactionRead.commit;
    FDQueryRead.Free;
    TransactionRead.Free;
  except
    on E: Exception do
      begin
        if Assigned(FDQueryRead) then
          FDQueryRead.Free;
        if Assigned(TransactionRead) then
          TransactionRead.Free;
        Memo1.Lines.Add('Ошибка обновления статуса задач' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.Button5Click(Sender: TObject);
begin
  if TaskListView.SelCount = 1 then
    begin
      EditTask.Caption := TaskListView.Selected.Caption;
      // имя задачи в заголовок
      EditTask.TaskName.Caption := TaskListView.Selected.Caption;
      // имя задачи в label т.к. происходит редактирование задачи
      // EditTask.ReadTableSelectedTask(EditTask.TaskName.Caption); // запускаем чтение таблицы и в label меняем описание таблицы на ее имя
      EditTask.ReadResulTask(EditTask.TaskName.Caption);
      // новая функция чтения таблицы
      EditTask.Button8.Enabled := true;
      if EditTask.WindowState = wsMinimized then
      // если форма свернута то восстанавливаем ее
        EditTask.WindowState := wsNormal
      else // иначе показываем
        EditTask.Show;
    end;

end;

procedure TfrmDomainInfo.Button6Click(Sender: TObject);
/// Запуск задачи
begin
  if TaskListView.SelCount = 1 then
    begin
      if EditTask.StatusStartStopTask(TaskListView.Selected.Caption) then
      // если задача запущена (true)
        begin
          showmessage('Задача ' + TaskListView.Selected.Caption +
            ' уже запущена');
          exit;
        end;
      if EditTask.StatrTask(TaskListView.Selected.Caption) then
        Memo1.Lines.Add('Запуск задачи ' + TaskListView.Selected.Caption +
          ' успешно выполнен');
    end;
  Button4.Click; // обновляем список задач
end;

procedure TfrmDomainInfo.Button7Click(Sender: TObject); // остановка задачи
begin
  if TaskListView.SelCount = 1 then
    begin
      if not EditTask.StatusStartStopTask(TaskListView.Selected.Caption) then
      // если задача остановлена (false)
        begin
          showmessage('Задача ' + TaskListView.Selected.Caption +
            ' уже остановлена');
          exit;
        end;
      if EditTask.StopTask(TaskListView.Selected.Caption) then
        frmDomainInfo.Memo1.Lines.Add('Операция, остановка задачи ' +
          TaskListView.Selected.Caption + ' успешно выполнена');
    end;
  Button4.Click; // обновляем список задач
end;

function createprogressbarforitem(statB: TStatusBar; MaxCount: Integer;
  NubItem: Integer): TprogressBar;
var
  PB: TprogressBar;
  // pbRect : TRect;
begin
  PB := TprogressBar.Create(statB);
  PB.Parent := statB;
  PB.left := statB.Panels[0].width + statB.Panels[1].width +
    statB.Panels[2].width;
  PB.width := statB.Panels[NubItem].width;
  PB.Min := 0;
  PB.Step := 1;
  PB.parentshowhint := false;
end;

procedure TfrmDomainInfo.CopyDelFFForListPC;
/// копирование файлов для группы компьютеров
var
  i: Integer;
  listPC: TstringList;
  function DlgFFF(s: string; var FCFolder: bool): Integer;
  /// фонкция создания диалогового окна TTaskDialog
  var
    Button: TTaskDialogButtonItem;
    // - кнопки в виде link,  TTaskDialogBaseButtonItem; //- просто кнопки
  begin
    with TTaskDialog.Create(self) do
      try
        Title := 'Выберите тип объекта для копирования';
        Caption := 'Выбор типа объекта файловой системы';
        // Text := Format('Are you sure that you want to remove the book file named "%s"?', [FNameOfBook]);
        CommonButtons := [];
        VerificationText := 'Проверять и создавать каталог назначения';
        Flags := [tfUseCommandLinks, tfVerificationFlagChecked];
        // тип кнопок в виде link
        with TTaskDialogButtonItem(Buttons.Add) do
          begin
            Caption := 'Файлы';
            CommandLinkHint := 'Выбор и копирование только фалов.';
            ModalResult := mrOk;
            Default := true;
          end;
        with TTaskDialogButtonItem(Buttons.Add) do
          begin
            Caption := 'Каталоги';
            CommandLinkHint := 'Выбор и копирование только каталогов.';
            ModalResult := mrNo;
          end;
        with TTaskDialogButtonItem(Buttons.Add) do
          begin
            Caption := 'Отмена';
            CommandLinkHint := 'Отменить операцию и выйти.';
            ModalResult := mrCancel;
          end;
        MainIcon := tdiNone;
        if Execute then
          begin
            FCFolder := tfVerificationFlagChecked in Flags;
            Result := ModalResult;
          end;
      finally
        Free;
      end
  end;

begin
  try
    listPC := TstringList.Create;
    for i := 0 to ListView8.Items.Count - 1 do
      begin
        if (ListView8.Items[i].ImageIndex <> 12) and (ListView8.Items[i].Checked)
        then
          listPC.Add(ListView8.Items[i].SubItems[0] + '=' +
            inttostr(ListView8.Items[i].ImageIndex));
        // имя компа+индекс картинки
      end;
    if listPC.Count = 0 then
      begin
        Memo1.Lines.Add('Не выбран список компьютеров');
        showmessage('Не выбран список компьютеров');
        listPC.Free;
        exit;
      end;
    FormCopyFFFGropPC.LVPC.Clear; // очистка списка компьютеров
    FormCopyFFFGropPC.LVFileFolder.Clear; // Очистка списка файлов и каталогов
    for i := 0 to listPC.Count - 1 do
      with FormCopyFFFGropPC.LVPC.Items.Add do
        begin
          Caption := listPC.Names[i]; // имя компьютера
          Checked := true;
          ImageIndex := strtoint(listPC.ValueFromIndex[i]); // индекс картинки
        end;
    FormCopyFFFGropPC.EditUser.text := LabeledEdit1.text;
    FormCopyFFFGropPC.EditPaswd.text := LabeledEdit2.text;
    FormCopyFFFGropPC.showmodal;
  except
    on E: Exception do
      Memo1.Lines.Add(Format('Ошибка открытия диалога:  %s', [E.Message]));
  end;
  listPC.Free;
end;

procedure TfrmDomainInfo.ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i: Integer;
  s: ShortString;
begin

  for i := 0 to ListView9.Items.Count - 1 do
    begin
      s := Ansiuppercase(System.copy(ListView9.Items[i].SubItems[0], 1,
        Length(ButtonedEdit1.text)));
      if Ansiuppercase(ButtonedEdit1.text) = s then
        begin
          ListView9.Items[i].Selected := true;
          ListView9.ItemIndex := i;
          ListView9.ItemFocused := ListView9.Items[i];
          break;
        end;
    end;
  if ListView9.ItemFocused <> nil then
    begin
      ListView9.ItemFocused.MakeVisible(false);

    end;
end;

procedure TfrmDomainInfo.LoadNetworkShare;
var
/// //////////////// загрузка списка сетевых ресурсов
  FSWbemLocator: OleVariant;
  FWMIService: OleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject: OleVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  function typeshare(z: Integer): string;
  begin
    case z of
      0:
        Result := 'Общий сетевой ресурс';
      1:
        Result := 'Очередь печати';
      2:
        Result := 'Устройство';
      3:
        Result := 'МПК';
      2147483648:
        Result := 'Ресурс для административных целей';
      2147483649:
        Result := 'Очереди печати, Admin';
      2147483650:
        Result := 'Устройство, Admin';
      2147483651:
        Result := 'Удаленный IPC';
    else
      Result := 'Unknown';
    end;
  end;

  function nameinimage(s: string): Integer;
  begin
    if s <> '' then
      begin
        if pos('$', s) = 2 then
          Result := 0
        else
          Result := 1;
      end
    else
      Result := 1;
  end;

begin
  try
    OleInitialize(nil);
    ListViewShare.Clear;
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    if CheckBoxShare.Checked then
      FWMIService := FSWbemLocator.ConnectServer(ComboBox2.text,
        'root\CIMV2', '', '')
      /// WbemUser, WbemPassword
    else
      FWMIService := FSWbemLocator.ConnectServer(ComboBox2.text, 'root\CIMV2',
        LabeledEdit1.text, LabeledEdit1.text);
    /// WbemUser, WbemPassword
    // FWMIService.Security_.impersonationlevel:=3;
    // FWMIService.Security_.authenticationLevel := 6;
    FWbemObjectSet := FWMIService.ExecQuery('SELECT * FROM Win32_Share', 'WQL',
      wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// / принтер
      begin
        if FWbemObject.Caption <> null then
          begin
            with ListViewShare.Items.Add do
              begin
                Caption := '';
                /// ////////////////////////////////// имя ресурса
                if not VarIsNull(FWbemObject.Name) then
                  begin
                    ImageIndex := nameinimage(string(FWbemObject.Name));
                    SubItems.Add(string(FWbemObject.Name))
                  end
                else
                  SubItems.Add('');

                /// //////////////////////////////////путь
                if not VarIsNull(FWbemObject.Path) then
                  SubItems.Add(string(FWbemObject.Path))
                else
                  SubItems.Add('');
                /// //////////////////////////////////////////////  ограничения
                if FWbemObject.AllowMaximum <> null then
                  begin
                    if FWbemObject.AllowMaximum then
                      SubItems.Add('Нет')
                    else
                      SubItems.Add('Да')
                  end;
                /// ////////////////////////////////////////   статус
                if not VarIsNull(FWbemObject.status) then
                  SubItems.Add(string(FWbemObject.status))
                else
                  SubItems.Add('');
                /// //////////////////////////////////////////////////// описание
                if not VarIsNull(FWbemObject.Description) then
                  SubItems.Add(string(FWbemObject.Description))
                else
                  SubItems.Add('');
                /// //////////////////////////////////////////////
                if VarIsNumeric(FWbemObject.type) then
                  SubItems.Add(typeshare(Integer(FWbemObject.type)))
                else
                  SubItems.Add('');
              end;
          end;
        FWbemObject := Unassigned;
      end;

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка загрузки сетевых ресурсов: ' + E.Message);
        VariantClear(FWbemObject);
        oEnum := nil;
        VariantClear(FWbemObjectSet);
      end;
  end;
  oEnum := nil;
  VariantClear(FWbemObject);
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize();
end;

/// //////////////////////////////////////////// uRDM
function Crypt(varStr: WideString): WideString;
var
  k: Integer;
  s: WideString;
begin
  RandSeed := 100;
  s := varStr;
  for k := 1 to Length(s) do
    s[k] := chr(ord(s[k]) xor (Random(127) + 1));
  Crypt := s;
end;

procedure TfrmDomainInfo.Code(var text: WideString; password: string;
  /// / процедура кодирования и декодирования файла
  decode: Boolean);
var
  i, PasswordLength: Integer;
  sign: shortint;
begin
  PasswordLength := Length(password);
  if PasswordLength = 0 then
    exit;
  if decode then
    sign := -1
  else
    sign := 1;
  for i := 1 to Length(text) do
    text[i] := chr(ord(text[i]) + sign *
      ord(password[i mod PasswordLength + 1]));
end;

procedure TfrmDomainInfo.ClientSocket1Connect(Sender: TObject;
  Socket: TCustomWinSocket);
var
  s: WideString;

begin
  Memo1.Lines.Add('Соединение установлено, отправляю запрос на сервер');
  s := WideString('<Client>' + LabeledEdit4.text + '<passwd>' +
    LabeledEdit5.text);
  // memo1.Lines.Add('WideString s - '+s);
  s := Crypt(s);
  /// / шифрация текста
  // memo1.Lines.Add('CODE s - '+s);
  ClientSocket1.Socket.SendText(s);
  /// отправляем зашифрованное сообщение
  ConnectionEnable := true;
end;

procedure TfrmDomainInfo.ClientSocket1Read(Sender: TObject;
  Socket: TCustomWinSocket);
/// / получаем сообщение
var
  s: WideString;
  UnloadFile: TMemInifile;
  UnloadList: TstringList;
  user: WideString;
  z, Y: Integer;
  ss: string;
begin
  try
    z := 0;
    Y := 0;
    user := WideString(LabeledEdit4.text);
    Memo1.Lines.Add('Пришел ответ от сервера - ' + MyPS);
    s := (Socket.ReceiveText);
    /// ответ от сервера,  строка соединения
    s := Crypt(s);
    /// / дешифрация ответа
    if s = WideString('<ErrorUser>') then
    /// ответ от сервера если ошибка
      begin
        Memo1.Lines.Add('Неверный пароль или имя пользователя');
        ClientSocket1.Active := false;
        MRDUViewerMain.Disconnect;
        exit;
      end;
    if s = WideString('<unloadlist>') then
    /// / признак получения файла настроек
      begin
        UnloadFileSetting := true;
        exit;
      end;

    if UnloadFileSetting then
    /// / получаем файл настроек
      begin
        UnloadList := TstringList.Create;
        UnloadList.SetText(PChar(s));
        UnloadFile := TMemInifile.Create(ChangeFileExt(Application.ExeName,
          'UnLF.ini'));
        UnloadFile.SetStrings(UnloadList);
        RemoteDesktopSetting.Memo1.Clear;
        UnloadFile.GetStrings(RemoteDesktopSetting.Memo1.Lines);
        UnloadList.Free;
        UnloadFile.Free;
        UnloadFileSetting := false;
        AccessSettingLevel := true;
        exit;
      end;

    z := pos('RemoteComputer', string(s));
    if (z <> 0) then
      begin
        Memo1.Lines.Add('Подключаю удаленный рабочий стол - ' + MyPS);
        MRDUViewerMain.Connect(s, user, WideString('ADMIN'));
        MRDUViewerMain.SmartSizing := true;
        ConnectionEnable := true;
      end;

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка получения данных от сервера  :' + E.Message);
        if ConnectionEnable then
          begin
            ClientSocket1.Close;
            MRDUViewerMain.Disconnect;
            if Assigned(MRDUViewerMain) then
              FreeAndNil(MRDUViewerMain);
            ConnectionEnable := false;
          end;
      end;
  end;

end;

procedure TfrmDomainInfo.ClientSocket1Connecting(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('Устанавливаю соединение');
end;

procedure TfrmDomainInfo.ClientSocket1Disconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  try
    Memo1.Lines.Add('Отключение от сервера');
    if ConnectionEnable then
      begin
        ClientSocket1.Close;
        ConnectionEnable := false;
      end;
    AccessSettingLevel := false;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при отключении от сервера  :' + E.Message);
      end;
  end;

end;

procedure TfrmDomainInfo.ClientSocket1Error(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent; var errorCode: Integer);
begin
  try
    Memo1.Lines.Add('Ошибка подключения к серверу: ' +
      SysErrorMessage(errorCode));
    errorCode := 0;
    if ConnectionEnable then
      begin
        ClientSocket1.Close;
        /// // отключение
        AccessSettingLevel := false;
        ConnectionEnable := false;
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при подключении к серверу  :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.ClientSocket1Write(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('Отправляю данные на сервер - ' + MyPS);
end;

procedure TfrmDomainInfo.RDP2Click(Sender: TObject);
/// / Контекстное меню списка компьютеров, подключение по RDP
var
  i: Integer;
  listPC: TstringList;
begin
  try
    if ((ListView8.Items.Count = 0) or (ListView8.SelCount = 0)) then
      exit;

    listPC := TstringList.Create;
    listPC := createListpcForCheck('');
    if (listPC.Count = 0) and (ListView8.SelCount = 1) then
      listPC.Add(ListView8.ItemFocused.SubItems[0]);
    for i := 0 to listPC.Count - 1 do
      begin
        if GroupRDPWin.ControlCount >= 1 then
          begin
            Memo1.Lines.Add
              ('Ограничено число одновременных подключений, необходимо приобрести лицензию');
            break;
          end;
        RDPWin.OtherWinForRDPClient(SpeedButton85.Tag, listPC[i],
          LabeledEdit3.text, LabeledEdit1.text, LabeledEdit2.text,
          frmDomainInfo);
        SpeedButton85.Tag := SpeedButton85.Tag + 1;
      end;
    listPC.Free;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка создания подключения для :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.RDP1Click(Sender: TObject);
/// / Контекстное меню списка компьютеров, подключение по RDP
var
  i: Integer;
  listPC: TstringList;
begin
  try
    if ((ListView8.Items.Count = 0) or (ListView8.SelCount = 0)) then
      exit;
    listPC := TstringList.Create;
    listPC := createListpcForCheck('');
    if (listPC.Count = 0) and (ListView8.SelCount = 1) then
      listPC.Add(ListView8.ItemFocused.SubItems[0]);
    for i := 0 to listPC.Count - 1 do
      begin
        RDPWin.OtherWinForRDPClient(SpeedButton85.Tag, listPC[i],
          LabeledEdit3.text, LabeledEdit1.text, LabeledEdit2.text,
          frmDomainInfo);
        SpeedButton85.Tag := SpeedButton85.Tag + 1;
      end;
    listPC.Free;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка создания подключения для :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.uRDM1Click(Sender: TObject);
/// / Контекстное меню списка компьютеров, подключение по uRDM
begin
  if ((ListView8.Items.Count = 0) or (ListView8.SelCount <> 1)) then
    exit;
  MRDForm.LabeledEdit1.text := ListView8.Items[ListView8.ItemFocused.Index]
    .SubItems[0];
  MRDForm.EditUser.text := LabeledEdit4.text;
  MRDForm.EditPass.text := LabeledEdit5.text;
  MRDForm.Editport.text := LabeledEdit6.text;
  MRDForm.Show;
end;

procedure TfrmDomainInfo.RDPSession1Error(ASender: TObject;
  ErrorInfo: OleVariant);
begin
  Memo1.Lines.Add('Ошибка сессии - ' + vartostr(ErrorInfo));
end;

procedure TfrmDomainInfo.RDPViewer1AttendeeConnected(ASender: TObject;
  const pAttendee: IDispatch);
begin
  Memo1.Lines.Add('Подключаюсь к сеансу');
  ConnectionEnable := true;
end;

procedure TfrmDomainInfo.RDPViewer1AttendeeDisconnected(ASender: TObject;
  const pDisconnectInfo: IDispatch);
var
  myclient: IRDPSRAPIAttendeeDisconnectInfo;
begin
  /// // какой пользователь отключился от сеанса
  pDisconnectInfo.QueryInterface(IID_IRDPSRAPIAttendeeDisconnectInfo, myclient);
  // if LabelEdEdit4.Text<>string(Myclient.Attendee.RemoteName) then
  Memo1.Lines.Add(string(myclient.Attendee.RemoteName) +
    ' отключился от сеанса');

end;

procedure TfrmDomainInfo.RDPViewer1AttendeeUpdate(ASender: TObject;
  const pAttendee: IDispatch);
begin
  // memo1.Lines.Add('изменяется одно из значений свойств для участника');
end;

procedure TfrmDomainInfo.RDPViewer1ChannelDataReceived(ASender: TObject;
  const pChannel: IInterface; lAttendeeId: Integer; const bstrData: WideString);
begin
  // memo1.Lines.Add('данные принимаются от участника');
end;

procedure TfrmDomainInfo.RDPViewer1ChannelDataSent(ASender: TObject;
  const pChannel: IInterface; lAttendeeId, BytesSent: Integer);
begin

  // memo1.Lines.Add('данные отправляются клиенту');
end;

procedure TfrmDomainInfo.RDPViewer1ConnectionAuthenticated(Sender: TObject);
begin
  // memo1.Lines.Add('соединение аутентифицировано');
end;

procedure TfrmDomainInfo.RDPViewer1ConnectionEstablished(Sender: TObject);
begin
  Memo1.Lines.Add('Соединение с сервером установлено');
  ConnectionEnable := true;
end;

procedure TfrmDomainInfo.RDPViewer1ConnectionFailed(Sender: TObject);
begin
  Memo1.Lines.Add('клиент не может подключиться к серверу');
  try
    if ConnectionEnable then
      begin
        ConnectionEnable := false;
        MRDUViewerMain.Disconnect;
        if Assigned(MRDUViewerMain) then
          FreeAndNil(MRDUViewerMain);
        ClientSocket1.Close;
        AccessSettingLevel := false
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка подключения к серверу  :' + E.Message);
      end;
  end;

end;

procedure TfrmDomainInfo.RDPViewer1ConnectionTerminated(ASender: TObject;
  discReason, ExtendedInfo: Integer);
begin
  try
    case discReason of
      0:
        Memo1.Lines.Add('Соединение закрыто - Информайия отсутствует');
      1:
        Memo1.Lines.Add('Соединение закрыто - Локальное отключение');
      2:
        Memo1.Lines.Add
          ('Соединение закрыто - Удаленное отключение пользователем');
      3:
        Memo1.Lines.Add('Соединение закрыто - Удаленное отключение сервером');
      260:
        Memo1.Lines.Add('Соединение закрыто - Ошибка поиска имени DNS');
      262:
        Memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
      264:
        Memo1.Lines.Add('Соединение закрыто - Время соединения истекло.');
      516:
        Memo1.Lines.Add('Соединение закрыто - Сбой подключения к Windows.');
      518:
        Memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
      520:
        Memo1.Lines.Add('Соединение закрыто - Ошибка хост не найден.');
      772:
        Memo1.Lines.Add('Соединение закрыто - Windows Sockets ошибка вызова.');
      774:
        Memo1.Lines.Add('Соединение закрыто - Недостаточно памяти.');
      776:
        Memo1.Lines.Add('Соединение закрыто - Указан неверный IP-адрес');
      1028:
        Memo1.Lines.Add('Соединение закрыто - Ошибка вызова Windows Sockets');
      1030:
        Memo1.Lines.Add
          ('Соединение закрыто - Недействительные данные безопасности.');
      1032:
        Memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка.');
      1286:
        Memo1.Lines.Add
          ('Соединение закрыто - Указан недопустимый метод шифрования.');
      1288:
        Memo1.Lines.Add('Соединение закрыто - Ошибка поиска DNS.');
      1540:
        Memo1.Lines.Add
          ('Соединение закрыто - Ошибка сокета Windows Sockets gethostbyname.');
      1542:
        Memo1.Lines.Add
          ('Соединение закрыто - Недействительные данные безопасности сервера..');
      1544:
        Memo1.Lines.Add('Соединение закрыто - Ошибка внутреннего таймера.');
      1796:
        Memo1.Lines.Add('Соединение закрыто - Произошел тайм-аут.');
      1798:
        Memo1.Lines.Add
          ('Соединение закрыто - Не удалось распечатать сертификат сервера.');
      2052:
        Memo1.Lines.Add('Соединение закрыто - Указан неверный IP-адрес.');
      2056:
        Memo1.Lines.Add
          ('Соединение закрыто - Согласование лицензии не удалось.');
      2310:
        Memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка безопасности.');
      2308:
        Memo1.Lines.Add('Соединение закрыто - Server Winsock закрыт.');
      2312:
        Memo1.Lines.Add('Соединение закрыто - Тайм-аут лицензирования.');
      2566:
        Memo1.Lines.Add('Соединение закрыто - Внутренняя ошибка безопасности.');
      2822:
        Memo1.Lines.Add('Соединение закрыто - Ошибка шифрования.');
      3078:
        Memo1.Lines.Add('Соединение закрыто - Ошибка расшифровки.');
      3080:
        Memo1.Lines.Add('Соединение закрыто - Ошибка декомпрессии.');
    end;
    case ExtendedInfo of
      1:
        Memo1.Lines.Add('Причина - Приложение инициировало отключение.');
      2:
        Memo1.Lines.Add('Причина - Приложение инициировало выход клиента.');
      3:
        Memo1.Lines.Add
          ('Причина - Сервер отключил клиент, потому что клиент простаивал в течение периода времени, превышающего указанный период тайм-аута.');
      4:
        Memo1.Lines.Add
          ('Причина - Сервер отключил клиент, потому что клиент превысил период, указанный для подключения.');
      5:
        Memo1.Lines.Add
          ('Причина - Соединение клиента было заменено другим соединением.');
      6:
        Memo1.Lines.Add('Причина - Нет памяти.');
      7:
        Memo1.Lines.Add('Причина - Сервер отказал в подключении.');
      8:
        Memo1.Lines.Add
          ('Причина - Сервер отказал в соединении по соображениям безопасности.');
      9:
        Memo1.Lines.Add
          ('Причина - Соединение было отклонено, поскольку учетная запись пользователя не авторизована для удаленного входа в систему.');
      11:
        Memo1.Lines.Add('Причина - Ошибка внутреннего лицензирования.');
      12:
        Memo1.Lines.Add('Причина - Сервер лицензий недоступен.');
      13:
        Memo1.Lines.Add
          ('Причина - Лицензия на программное обеспечение не была доступна.');
      14:
        Memo1.Lines.Add
          ('Причина - Удаленный компьютер получил недопустимое лицензионное сообщение.');
      15:
        Memo1.Lines.Add
          ('Причина - Идентификатор оборудования не совпадает с идентификатором, указанным в лицензии на программное обеспечение.');
      16:
        Memo1.Lines.Add('Причина - Ошибка лицензии клиента.');
      17:
        Memo1.Lines.Add
          ('Причина - Сетевые проблемы возникли во время протокола лицензирования.');
      18:
        Memo1.Lines.Add
          ('Причина - Клиент преждевременно завершил протокол лицензирования.');
      20:
        Memo1.Lines.Add
          ('Причина - Лицензионное сообщение было зашифровано неправильно.');
      21:
        Memo1.Lines.Add
          ('Причина - Лицензия на доступ к локальному компьютеру не может быть обновлена ​​или обновлена.');
      22:
        Memo1.Lines.Add
          ('Причина - Удаленный компьютер не имеет лицензии на прием удаленных подключений.');
      23:
        Memo1.Lines.Add
          ('Причина - Соединение было отклонено, поскольку продавец не смог аутентифицировать зрителя.');
      24:
        Memo1.Lines.Add('Причина - внутренние ошибки протокола');
    end;
    if ConnectionEnable then
      begin
        ClientSocket1.Close;
        MRDUViewerMain.Disconnect;
        if Assigned(MRDUViewerMain) then
          FreeAndNil(MRDUViewerMain);
        AccessSettingLevel := false;
        ConnectionEnable := false;
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при подключениик к рабочему столу  :' +
          E.Message);
      end;
  end;

end;

procedure TfrmDomainInfo.RDPViewer1ControlLevelChangeRequest(ASender: TObject;
  const pAttendee: IDispatch; RequestedLevel: TOleEnum);
begin
  /// // если сервер изменил уровень привелегий
  // memo1.Lines.Add('Уровень привилегий - CTRL_LEVEL_INTERACTIVE - '+ string(RequestedLevel)) ;
  // memo1.Lines.Add('Уровень привилегий - CTRL_LEVEL_VIEW - '+ string(RequestedLevel)) ;

end;

procedure TfrmDomainInfo.RDPViewer1FocusReleased(ASender: TObject;
  iDirection: Integer);
begin
  // memo1.Lines.Add('RDPViewer1FocusReleased');
end;

procedure TfrmDomainInfo.RDPViewer1GraphicsStreamPaused(Sender: TObject);
begin
  // memo1.Lines.Add('RDPViewer1GraphicsStreamPaused');

end;

procedure TfrmDomainInfo.RDPViewer1GraphicsStreamResumed(Sender: TObject);
begin
  // memo1.Lines.Add('RDPViewer1GraphicsStreamResumed');
end;

procedure TfrmDomainInfo.RDPViewer1SharedDesktopSettingsChanged
  (ASender: TObject; width, height, colordepth: Integer);
begin
  // memo1.Lines.Add('RDPViewer1SharedDesktopSettingsChanged');
end;

procedure TfrmDomainInfo.RDPViewer1SharedRectChanged(ASender: TObject;
  left, top, right, bottom: Integer);
begin
  // memo1.Lines.Add('RDPViewer1SharedRectChanged');
end;

procedure TfrmDomainInfo.RDPViewer1StartDrag(Sender: TObject;
  var DragObject: TDragObject);
begin
  // memo1.Lines.Add('RDPViewer1StartDrag');
end;

procedure TfrmDomainInfo.RDPViewer1WindowClose(ASender: TObject;
  const pWindow: IDispatch);
begin
  // memo1.Lines.Add('RDPViewer1WindowClose');
end;

procedure TfrmDomainInfo.RDPViewer1WindowOpen(ASender: TObject;
  const pWindow: IDispatch);
begin
  // memo1.Lines.Add('RDPViewer1WindowOpen');
end;

procedure TfrmDomainInfo.RDPViewer1WindowUpdate(ASender: TObject;
  const pWindow: IDispatch);
begin
  // memo1.Lines.Add('RDPViewer1WindowUpdate');
end;

{ procedure TfrmDomainInfo.Button3Click(Sender: TObject);
  var
  StAuthString,StGroupName,StPassword:string;
  stream:Olevariant;
  begin
  try
  levelCntrl:=false; /// уровень привелегий CTRL_LEVEL_VIEW  , если true то  CTRL_LEVEL_INTERACTIVE
  MyPS:=Combobox2.Text;
  if myps='' then begin showmessage('Укажите имя компьютера'); exit; end;
  ClientSocket1.Port:=strtoint(LabelEdEdit6.Text);   /// порт клиента
  ClientSocket1.Host:=MyPS;   //// имя сервера
  ClientSocket1.Active:=true;  /// активация клиента
  UnloadFileSetting:=false;
  Except
  on E: Exception do
  begin
  memo1.Lines.add('Ошибка при подключениик к рабочему столу  :'+E.Message);
  end;
  end;
  end; }

procedure TfrmDomainInfo.InventorySoftware;
begin
  inventorySoftForSelectPC(ComboBox2.text, LabeledEdit1.text,
    LabeledEdit2.text);
end;

procedure TfrmDomainInfo.EditNamePC;
begin
  if ping(ComboBox2.text) then
    begin
      OKRightDlg1234567.Caption := ComboBox2.text;
      OKRightDlg1234567.showmodal;
    end;
end;

procedure TfrmDomainInfo.AddPCDomain;
begin
  if ping(ComboBox2.text) then
    begin
      OKRightDlg123456.showmodal;
    end;
end;

procedure TfrmDomainInfo.RemovePCDomain;
begin
/// удалить из домена
  OKRightDlg12345678.showmodal;
end;

procedure TfrmDomainInfo.OpenFormSMATR;
begin
  if ping(ComboBox2.text) then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      Form8.showmodal;
    end;
end;

procedure TfrmDomainInfo.PropertiesForPC;
begin
  MyPS := ComboBox2.text;
  OKRightDlg12345678910111213.showmodal;
end;

procedure TfrmDomainInfo.SpeedButton78Click(Sender: TObject);
begin
  try
    if Assigned(SelectedPCPing) then
      begin
        if SelectedPCPing.Started then
          begin
            OutForPing := false;
            SolveExitInvScan:=true;
            if TerminateThread(SelectedPCPing.Handle, 0) then
              Memo1.Lines.Add
                ('Принудительная остановка сканирования сети завершена успешно');
            SpeedButton39.Visible := true;
            SpeedButton40.Visible := false;

          end;
      end
    else
      begin
        SpeedButton39.Visible := true;
        SpeedButton40.Visible := false;
        OutForPing := false;
        SolveExitInvScan:=true;
      end;
    SpeedButton39.Visible := true;
    SpeedButton40.Visible := false;
    OutForPing := false;
    SolveExitInvScan:=true;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка KillScan  :' + E.Message);
        SpeedButton39.Visible := true;
        SpeedButton40.Visible := false;
        OutForPing := false;
        SolveExitInvScan:=true;
      end;
  end;
end;

procedure TfrmDomainInfo.RenewInventNumber;
var
  NewDescription: string;
begin
  if ComboBox2.text = '' then
    begin
      showmessage('Не выбран компьютер');
      exit;
    end;
  if not InputQuery('Описание компьютера в свойствах (Инв№)', 'Описание',
    NewDescription) then
    exit;
  PutInvNumberToDataBase(ComboBox2.text,NewDescription);
  PutInvNumber(ComboBox2.text, NewDescription);
end;

procedure TfrmDomainInfo.SpeedButton7Click(Sender: TObject);
begin
  try
    if Assigned(MRDUViewerMain) then
      begin
        if levelCntrl then
        /// /// изменение уровня привелегий
          begin
            MRDUViewerMain.RequestControl(CTRL_LEVEL_VIEW);
            levelCntrl := false;
          end
        else
          begin
            MRDUViewerMain.RequestControl(CTRL_LEVEL_MAX);
            levelCntrl := true;
          end;
      end
    else
      showmessage('Установите соединение с сервером!');
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при смене уровня привелегий  :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.FindProcessForListPC;
begin
  Form12.showmodal;
end;

/// ////////////////////////////////////////////////////////////////////
/// ///////////////////////////////////////////////////////// работа с событиями Windows
procedure TfrmDomainInfo.SpeedButton81Click(Sender: TObject);
var
  i: Integer;
begin
/// / загрузка списка имен лог файлов
  try
    if ping(ComboBox2.text) then
      begin
        ComboBoxLogFile.Clear;
        ComboBoxLogFile.Items.AddStrings(readlistLogsFileEvent(ComboBox2.text,
          LabeledEdit1.text, LabeledEdit2.text));
        if ComboBoxLogFile.Items.Count > 0 then
          begin
            for i := 0 to ComboBoxLogFile.Items.Count - 1 do
              if ComboBoxLogFile.Items[i] = 'System' then
                begin
                  ComboBoxLogFile.ItemIndex := i;
                  ComboBoxLogFileSelect(self);
                  /// загрузка событий журнала system
                  break;
                end;
          end;
      end;

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.N100Click(Sender: TObject);
begin
  if ListViewEvent.Items.Count <> 0 then
    popupListViewSaveAs(ListViewEvent, 'Сохранение списка событий', '');
end;

procedure TfrmDomainInfo.ListViewEventColumnClick(Sender: TObject;
  /// сортировка
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  ListViewEvent.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.CheckBox3Click(Sender: TObject);
begin
  DatePickerFrom.Enabled := CheckBox3.Checked;
  DatePickerTo.Enabled := CheckBox3.Checked;
  TimePickerFrom.Enabled := CheckBox3.Checked;
  TimePickerTo.Enabled := CheckBox3.Checked;
end;

procedure clearinfoEvent;
begin
  frmDomainInfo.Label9.Caption := '';
  frmDomainInfo.Label10.Caption := '';
  frmDomainInfo.Label11.Caption := '';
  frmDomainInfo.Label12.Caption := '';
  frmDomainInfo.Label13.Caption := '';
  frmDomainInfo.Label14.Caption := '';
  frmDomainInfo.Label15.Caption := '';
  frmDomainInfo.Label16.Caption := '';
  frmDomainInfo.MemoEventInfo.Clear;
end;

procedure TfrmDomainInfo.SpeedButton82Click(Sender: TObject);
var
  FilterType, NewTread: Integer;
  treadID: LongWord;
  Typebool, FilterLogFile, FilterSourse: bool;
begin
  try
    if not ping(ComboBox2.text) then
      exit;

    if ComboBoxLogFile.text = 'Обновите список журналов' then
      begin
        Memo1.Lines.Add('Обновляю список журналов ... ');
        SpeedButton81.Click;
      end;
    SpeedButton82.Enabled := false;
    clearinfoEvent;
    /// /очистка лейбов и мемо
    FilterLogFile := false;
    FilterSourse := false;
    FilterType := RadioGroupEventType.ItemIndex - 1;
    /// тип фильтруемых событий (0-5)
    if FilterType <> -1 then
      Typebool := true
    else
      Typebool := false;
    /// если выбран тип фильтрации
    if (ComboBoxLogFile.text <> '') or (pos(' ', ComboBoxLogFile.text) = 1) then
      FilterLogFile := true;
    if (ComboBoxSourceEvent.text <> '') or
      (pos(' ', ComboBoxSourceEvent.text) = 1) then
      FilterSourse := true;

    ListViewEvent.Clear;
    /// очистка листбокса
    EventCall.NamePC := ComboBox2.text;
    EventCall.user := LabeledEdit1.text;
    EventCall.pass := LabeledEdit2.text;
    EventCall.FileLog := ComboBoxLogFile.text;
    EventCall.SourseLog := ComboBoxSourceEvent.text;
    EventCall.DateTimeFrom := (FormatDateTime('yyyymmdd', DatePickerFrom.Date))
      + (FormatDateTime('hhmmss', TimePickerFrom.time));
    EventCall.DateTimeTo := (FormatDateTime('yyyymmdd', DatePickerTo.Date)) +
      (FormatDateTime('hhmmss', TimePickerTo.time));
    EventCall.TypeEvent := FilterType;
    EventCall.CounEvent := strtoint(LabeledEdit8.text);
    EventCall.FilterSourse := FilterSourse; // CheckBox9.Checked;
    EventCall.FilterType := Typebool;
    EventCall.FilterDate := CheckBox3.Checked;
    EventCall.FilterLogFile := FilterLogFile; // CheckBox4.Checked;
    EventCall.ListEvent := ListViewEvent;
    GroupBoxEventProperties.Caption := 'Загрузка событий, ожидайте...';
    NewTread := BeginThread(nil, 0, addr(readlistViewEvent), addr(EventCall), 0,
      treadID);
    /// поток для заполнения списка событий
    CloseHandle(NewTread);
    { FuncReadlistViewEvent(          /// функция заполенения списка событий
      combobox2.Text, /// компьютер
      ComboBoxLogFile.Text              /// лог файл
      ,ComboBoxSourceEvent.Text         /// источник событий
      ,(FormatDateTime('yyyymmdd',DatePickerFrom.Date))+(FormatDateTime('hhmmss',TimePickerFrom.time))//// от даты и времени
      ,(FormatDateTime('yyyymmdd',DatePickerTo.Date))+(FormatDateTime('hhmmss',TimePickerTo.time))      /// до даты и времени
      ,FilterType                       /// выбранный тип событий ( Сведения -0,Ошибка - 1,Предупреждение-2,Сведения -3 ,Аудит успеха- 4 .Аудит отказа - 5
      ,strtoint(LabelEdEdit8.Text)      /// количество загружаемых событий
      ,CheckBox9.Checked               /// фильтровать ли по источнику событий
      ,Typebool                        /// фильтровать по типу событий
      ,CheckBox3.Checked
      ,listViewEvent); }
    /// ListView куда записать данные
  Except
    on E: Exception do
      begin
        GroupBoxEventProperties.Caption := '';
        SpeedButton82.Enabled := true;
        Memo1.Lines.Add('Ошибка :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton83Click(Sender: TObject);
begin
/// / свойства журнала событий
  if not ping(ComboBox2.text) then
    exit;

  if ComboBoxLogFile.ItemIndex = -1 then
    begin
      showmessage('Выберите журнал из списка');
      exit;
    end;
  PropertiesFileEvent.Caption := ComboBox2.text;
  ///
  PropertiesFileEvent.GroupBox1.Caption := ComboBoxLogFile.text;
  PropertiesFileEvent.showmodal;
end;

procedure TfrmDomainInfo.EditUserGroup;
begin
  if ping(ComboBox2.text) then
    begin
      FormLocalAccount.Caption := 'Пользователи и группы - ' + ComboBox2.text;
      FormLocalAccount.showmodal;
    end;
end;

procedure TfrmDomainInfo.ConnectRDPOsherWindows(Sender: TObject);
var
  i: Integer;
begin
  try
    if GroupRDPWin.ControlCount >= 2 then
      begin
        Memo1.Lines.Add
          ('Ограничено число одновременных подключений, необходимо приобрести лицензию');
        exit;
      end;
    if not ping(ComboBox2.text) then
      begin
        i := MessageBox(self.Handle,
          PChar('Компьютер ' + ComboBox2.text +
          ' не доступен, продолжить подключение?'), PChar(ComboBox2.text),
          MB_YESNO + MB_ICONQUESTION);
        if i = IDNO then
          exit;
      end;

    RDPWin.OtherWinForRDPClient(SpeedButton85.Tag, ComboBox2.text,
      LabeledEdit3.text, LabeledEdit1.text, LabeledEdit2.text, frmDomainInfo);
    SpeedButton85.Tag := SpeedButton85.Tag + 1;

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка создания подключения для ' + ComboBox2.text +
          ' :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.NewRDPFormForPC;
var
  i: Integer;
begin
  try
    if not ping(ComboBox2.text) then
      begin
        i := MessageBox(self.Handle,
          PChar('Компьютер ' + ComboBox2.text +
          ' не доступен, продолжить подключение?'), PChar(ComboBox2.text),
          MB_YESNO + MB_ICONQUESTION);
        if i = IDNO then
          exit;
      end;

    RDPWin.OtherWinForRDPClient(SpeedButton85.Tag, ComboBox2.text,
      LabeledEdit3.text, LabeledEdit1.text, LabeledEdit2.text, frmDomainInfo);
    SpeedButton85.Tag := SpeedButton85.Tag + 1;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка создания подключения для ' + ComboBox2.text +
          ' :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.OpenRDPForListPC;
var
  patch: string;
  EXEOpenpatch: TOpenDialog;
begin
  patch := ExtractFilePath(Application.ExeName);
  if FileExists(patch + 'MRRDP.exe') then
    begin
      ShellExecute(0, PChar('open'), PChar(patch + 'MRRDP.exe'), nil, nil,
        SW_RESTORE);
    end
  else
    begin
      EXEOpenpatch := TOpenDialog.Create(self);
      EXEOpenpatch.Name := 'MRRDP';
      EXEOpenpatch.Title := 'Файл RDP клиента - MRRDP.exe';
      EXEOpenpatch.InitialDir := ExtractFilePath(Application.ExeName);
      EXEOpenpatch.Filter := '|*.exe';
      if EXEOpenpatch.Execute then
        begin
          ShellExecute(0, PChar('open'), PChar(EXEOpenpatch.filename), nil, nil,
            SW_RESTORE);
          Memo1.Lines.Add('Запуск приложения : ' + EXEOpenpatch.filename)
        end;
      EXEOpenpatch.Destroy;
    end;

end;

procedure TfrmDomainInfo.OpenFormExplorer;
var
  i: Integer;
begin
  if not ping(ComboBox2.text) then
    exit;
  MRPCExplorer.ComboBox2.Clear;
  for i := 0 to ComboBox2.Items.Count - 1 do
    begin
      MRPCExplorer.ComboBox2.Items.Add(ComboBox2.Items[i]);
      MRPCExplorer.ComboBox2.ItemsEx[i].ImageIndex := ComboBox2.ItemsEx[i]
        .ImageIndex;
    end;
  MRPCExplorer.ComboBox2.text := ComboBox2.text;
  /// имя компа
  MRPCExplorer.EditUser.text := LabeledEdit1.text;
  /// пользователь
  MRPCExplorer.EditPass.text := LabeledEdit2.text; // пароль
  if MRPCExplorer.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    MRPCExplorer.WindowState := wsNormal
  else // иначе показываем
    MRPCExplorer.Show;
end;

procedure TfrmDomainInfo.ComboBoxLogFileKeyPress(Sender: TObject;
  var Key: Char);
/// раскрывает список журналов событий
begin
  ComboBoxLogFile.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBoxLogFileKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  { if ComboBoxLogFile.Text='' then
    begin
    ShowMessage('Вы не указали журнал событий! Загрузка событий может занять '+#10#13+'много времени, укажите временной интервал.');
    end; }
end;

procedure TfrmDomainInfo.ComboBoxSourceEventDropDown(Sender: TObject);
begin
/// / ширина списка combobox
  DropDownWidth(ComboBoxSourceEvent);
end;

procedure TfrmDomainInfo.ComboBoxSourceEventKeyPress(Sender: TObject;
  var Key: Char); // раскрывает список источников событий
begin
  ComboBoxSourceEvent.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBoxLogFileSelect(Sender: TObject);
begin
/// загрузка списка источников событий лог файла
  try
    ComboBoxSourceEvent.Clear;
    ComboBoxSourceEvent.text := 'Загрузка событий';
    ComboBoxSourceEvent.Items.AddStrings(readlistSourceEvent(ComboBox2.text,
      ComboBoxLogFile.Items[ComboBoxLogFile.ItemIndex], LabeledEdit1.text,
      LabeledEdit2.text));
    ComboBoxSourceEvent.text := '';
    if not CheckBox3.Checked then
      begin
        DatePickerFrom.Date := Yesterday;
        /// установка вчерашней даты в поле от
        DatePickerTo.Date := Date;
        /// установка сегодняшней даты в поле до
        TimePickerTo.time := time;
        /// установка текущего времени
      end;
    // если выбрали журнал безопасности то включем фильтрацию по дате  иначе будет долго загружать события
    if ComboBoxLogFile.text = 'Security' then
      CheckBox3.Checked := true;

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.ListViewEventChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  GroupBoxEventProperties.Caption := 'Всего записей - ' +
    inttostr(ListViewEvent.Items.Count);
end;

procedure TfrmDomainInfo.ListViewEventSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
begin
  try
    if Selected then
      begin
      /// ///////////////////////выбор события в листбокс
        if (ListViewEvent.Items.Count > 0) and (ListViewEvent.SelCount = 1) then
          begin
            Datam.FDQueryTempTabl.SQL.Clear;
            /// чтение конфигурации первого вхождения и передача в treeview
            Datam.FDQueryTempTabl.SQL.text :=
              'SELECT * FROM EVENT WHERE PC_NAME=''' + ComboBox2.text +
              ''' AND LOG_FILE=''' + ListViewEvent.Items
              [ListViewEvent.ItemIndex].SubItems[4]
            /// или {ComboboxLogFile.Items[ComboboxLogFile.ItemIndex]}
              + ''' AND RECORD_NUMBER=''' + ListViewEvent.Items
              [ListViewEvent.ItemIndex].Caption + '''';
            Datam.FDQueryTempTabl.Open;
            /// / запускаем выборку
            if vartostr(Datam.FDQueryTempTabl.FieldByName('MESSAGE_EVENT')
              .Value) <> '' then
              MemoEventInfo.text :=
                vartostr(Datam.FDQueryTempTabl.FieldByName
                ('MESSAGE_EVENT').Value);
            MemoEventInfo.Lines.Add
              ('-----------------------Подробно----------------------');
            if vartostr(Datam.FDQueryTempTabl.FieldByName('INSERT_STR').Value)
              <> '' then
              MemoEventInfo.Lines.Add
                (vartostr(Datam.FDQueryTempTabl.FieldByName
                ('INSERT_STR').Value));

            if vartostr(Datam.FDQueryTempTabl.FieldByName('LOG_FILE').Value) <> ''
            then
              Label9.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('LOG_FILE').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('LOG_SOURCE').Value)
              <> '' then
              Label10.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('LOG_SOURCE').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('EVENT_CODE').Value)
              <> '' then
              Label11.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('EVENT_CODE').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('EVENT_TYPE').Value)
              <> '' then
              Label12.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('EVENT_TYPE').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('CUR_USER').Value) <> ''
            then
              Label13.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('CUR_USER').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('RECORD_NUMBER')
              .Value) <> '' then
              Label14.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName
                ('RECORD_NUMBER').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('TIME_GEN').Value) <> ''
            then
              Label15.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName('TIME_GEN').Value);

            if vartostr(Datam.FDQueryTempTabl.FieldByName('PC_NAME_LOG').Value)
              <> '' then
              Label16.Caption :=
                vartostr(Datam.FDQueryTempTabl.FieldByName
                ('PC_NAME_LOG').Value);

            if MemoEventInfo.Lines.Count > 2 then
              MemoEventInfo.Perform(WM_VScroll, SB_TOP, 0);
            /// перевод курсора в мемо в начало
            // readInfoSelEvent(ComboBox2.Text,ComboboxLogFile.Items[ComboboxLogFile.ItemIndex],ListViewEvent.Items[ListViewEvent.ItemIndex].Caption);
            Datam.FDQueryTempTabl.Close;
          end;
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.ListViewMicLicDblClick(Sender: TObject);
var
  liststr: TstringList;
begin
  if ListViewMicLic.SelCount = 0 then
    exit;
  liststr := TstringList.Create;
  try
    liststr.Add(ListViewMicLic.Selected.Caption);
    liststr.Add(ListViewMicLic.Selected.SubItems[0]);
    liststr.Add(ListViewMicLic.Selected.SubItems[1]);
    liststr.Add(ListViewMicLic.Selected.SubItems[2]);
    liststr.Add(ListViewMicLic.Selected.SubItems[3]);
    itemLisToMemo(liststr, GroupBox2.Caption, frmDomainInfo);
  finally
    liststr.Free;
  end;
end;

/// ////////////////////////////////////////////////////////////////////////////////////

procedure TfrmDomainInfo.SpeedButton8Click(Sender: TObject);
begin
  try
  /// / окно настройки привелегий
    if not ClientSocket1.Active then
      begin
        showmessage('Установите соединение с сервером!');
        exit;
      end;
    if AccessSettingLevel = true then
      begin
        RemoteDesktopSetting.Tag := 1;
        RemoteDesktopSetting.showmodal
      end
    else
      showmessage('Недостаточно привелегий' + #13#10 + 'учетной записи');
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка :' + E.Message);
      end;
  end;
end;

function TfrmDomainInfo.NewNameTask(s: string): string;
var
  nametask: string;
  SelectedPC: TstringList;
  i: Integer;
  WaitName: Boolean;
begin
  Button4.Click; // обновление списка задач
  WaitName := false;
  while not WaitName do
    Begin
      if not InputQuery('Введите имя Задачи', 'Имя задачи', nametask) then
        break;
      if (nametask = '') then
        break;
      if TaskListView.Items.Count > 0 then
        begin
          for i := 0 to TaskListView.Items.Count - 1 do
            begin
              if nametask = TaskListView.Items[i].Caption then
                begin
                  showmessage('Задача с таким именем уже существует');
                  WaitName := false;
                  nametask := '';
                  break;
                end
              else
                WaitName := true;
            end;
        end
      else
        WaitName := true;
    End;
  Result := nametask;
end;

procedure TfrmDomainInfo.N127Click(Sender: TObject);
var // добавление новой задачи для выделенного списка компьютеров из popup меню
  nametask: string;
  SelectedPC: TstringList;
  i: Integer;
  WaitName: Boolean;
begin
  try
    SelectedPC := TstringList.Create;
    SelectedPC := frmDomainInfo.createListpcForCheck('');
    if SelectedPC.Count = 0 then
      begin
        if ListView8.SelCount = 1 then
          SelectedPC.Add(ListView8.Selected.SubItems[0])
        else
          begin
            showmessage('Не выделен список компьютеров!');
            SelectedPC.Free;
            exit;
          end;
      end;
    nametask := NewNameTask('');
    if nametask = '' then
      begin
        SelectedPC.Free;
        exit;
      end;

    EditTask.ListView2.Clear; // очистка списка компов
    EditTask.ListView1.Clear; // очистка списка заданий
    EditTask.Caption := nametask; // имя взаоловок
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая

    for i := 0 to SelectedPC.Count - 1 do // заполнение списка компов
      begin
        with EditTask.ListView2.Items.Add do
          begin
            ImageIndex := 0;
            Caption := inttostr(EditTask.ListView2.Items.Count);
            SubItems.Add(SelectedPC[i]);
          end;
      end;
    EditTask.Button8.Enabled := false;
    // кнопка запуска не активна , активируется после сохранения задачи
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая
    if EditTask.WindowState = wsMinimized then
    // если форма свернута то восстанавливаем ее
      EditTask.WindowState := wsNormal
    else // иначе показываем
      EditTask.Show;
    Button7.Enabled := true;
    SelectedPC.Free;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка создания новой задачи :' + E.Message);
  end;
end;

procedure TfrmDomainInfo.N128Click(Sender: TObject);
begin
  RunTaskForListPC;
end;

procedure TfrmDomainInfo.N129Click(Sender: TObject);
var
  nametask: string;
  oldname: string;
begin
  if TaskListView.SelCount = 1 then
    begin
      oldname := TaskListView.Selected.Caption;
      if EditTask.StatusStartStopTask(oldname) then
      // если задача запущена (true)
        begin
          showmessage('Задача ' + oldname + 'запущена,' + #10#13 +
            ' остановите задачу и повторите попытку.');
          exit;
        end;
      nametask := NewNameTask('');
      if nametask = '' then
        begin
          exit;
        end;

      if EditTask.RenameTableForTask(EditTask.nametableForDescription(oldname),
        nametask) then
        begin
          Memo1.Lines.Add('Операция переименование задачи ' + oldname + ' -> ' +
            nametask + ' успешно выполнена');
          Button4.Click;
        end;

    end;
end;

procedure TfrmDomainInfo.CreateNewTaskForListPC;
var // добавление новой задачи для выделенного списка компьютеров
  nametask: string;
  SelectedPC: TstringList;
  i: Integer;
  WaitName: Boolean;
begin
  try
    SelectedPC := TstringList.Create;
    SelectedPC := frmDomainInfo.createListpcForCheck('');
    if SelectedPC.Count = 0 then
      begin
        if ListView8.SelCount = 1 then
          SelectedPC.Add(ListView8.Selected.SubItems[0])
        else
          begin
            showmessage('Не выделен список компьютеров!');
            SelectedPC.Free;
            exit;
          end;
      end;
    nametask := NewNameTask('');
    if nametask = '' then
      begin
        SelectedPC.Free;
        exit;
      end;

    EditTask.ListView2.Clear; // очистка списка компов
    EditTask.ListView1.Clear; // очистка списка заданий
    EditTask.Caption := nametask; // имя взаоловок
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая

    for i := 0 to SelectedPC.Count - 1 do // заполнение списка компов
      begin
        with EditTask.ListView2.Items.Add do
          begin
            ImageIndex := 0;
            Caption := inttostr(EditTask.ListView2.Items.Count);
            SubItems.Add(SelectedPC[i]);
          end;
      end;
    EditTask.Button8.Enabled := false;
    // кнопка запуска не активна , активируется после сохранения задачи
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая
    if EditTask.WindowState = wsMinimized then
    // если форма свернута то восстанавливаем ее
      EditTask.WindowState := wsNormal
    else // иначе показываем
      EditTask.Show;
    Button7.Enabled := true;

    SelectedPC.Free;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка создания новой задачи :' + E.Message);
  end;
end;

procedure TfrmDomainInfo.CreateNewTaskForPC;
var // добавление новой задачи для одного текущего компьютера
  nametask: string;
begin
  nametask := frmDomainInfo.NewNameTask('');
  if nametask = '' then
    begin
      exit;
    end;

  try
    EditTask.ListView2.Clear; // очистка списка компов
    EditTask.ListView1.Clear; // очистка списка заданий
    EditTask.Caption := nametask; // имя взаоловок
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая
    with EditTask.ListView2.Items.Add do
      begin
        ImageIndex := 0;
        Caption := inttostr(EditTask.ListView2.Items.Count);
        SubItems.Add(ComboBox2.text);
      end;
    EditTask.Button8.Enabled := false;
    // кнопка запуска не активна , активируется после сохранения задачи
    EditTask.TaskName.Caption := ''; // очищаем label т.к. задача новая
    if EditTask.WindowState = wsMinimized then
    // если форма свернута то восстанавливаем ее
      EditTask.WindowState := wsNormal
    else // иначе показываем
      EditTask.Show;
    Button7.Enabled := true;
  finally

  end;
end;

procedure TfrmDomainInfo.RunTaskForListPC;
// запуск задачи для выбранного списка компьютеров
var
  SelectedPC: TstringList;
begin
  SelectedPC := TstringList.Create;
  SelectedPC := frmDomainInfo.createListpcForCheck('');
  if SelectedPC.Count = 0 then
    begin
      if ListView8.SelCount = 1 then
        SelectedPC.Add(ListView8.Selected.SubItems[0])
      else
        begin
          showmessage('Не выделен список компьютеров!');
          SelectedPC.Free;
          exit;
        end;
    end;
  SelectedPC.Free;
  SelectTask.Tag := 1; // признак того что задача запускается для группы компов
  if SelectTask.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    SelectTask.WindowState := wsNormal
  else // иначе показываем
    SelectTask.Show;

end;

procedure TfrmDomainInfo.RunTaskForPC; // запукск задачи для текущего компьютера
begin
  SelectTask.Tag := 0;
  // признак того что задача запускается для одного выбранного компьютера
  if SelectTask.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    SelectTask.WindowState := wsNormal
  else // иначе показываем
    SelectTask.Show;
end;

procedure TfrmDomainInfo.SpeedButton9Click(Sender: TObject);
begin
  try
    if not ConnectionEnable then
      begin
        levelCntrl := false;
        /// уровень привелегий CTRL_LEVEL_VIEW  , если true то  CTRL_LEVEL_INTERACTIVE
        MyPS := ComboBox2.text;
        if MyPS = '' then
          begin
            showmessage('Укажите имя компьютера');
            exit;
          end;
        ClientSocket1.Port := strtoint(LabeledEdit6.text);
        /// порт клиента
        ClientSocket1.host := MyPS;
        /// / имя сервера
        ClientSocket1.Active := true;
        /// активация клиента
        UnloadFileSetting := false;
        ConnectionEnable := true;
        createviewerMain;
        /// процедура создания вьювера
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при подключении  :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.ResetForListPC;
var
  i: Integer;
  /// /// групповая перезагрузка
  NewSelectedPCShotDown: TThread;
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  i := MessageDlg('Выполнить перезагрузку на выбранных компьютерах? ',
    mtWarning, [mbYes, mbCancel], 0);
  if i = mrYes then
    begin
      if ListView8.Items.Count > 0 then
        begin
          SelectedPCShutDown := TstringList.Create;
          SelectedPCShutDown := createListpcForCheck('');
          /// функция создания списка компов
          if SelectedPCShutDown.Count <> 0 then
            begin
              myShutdown := 2;
              NewSelectedPCShotDown :=
                SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
              NewSelectedPCShotDown.FreeOnTerminate := true;
              NewSelectedPCShotDown.Start;
            end
          else
            showmessage('Не выбран не один компьютер!!!');
        end;
    end;

end;

procedure TfrmDomainInfo.PowerOff;
var
  MyShutdownPC: TThread;
begin // выключение/перезагрузка/завершение сейнса
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 4;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.LogOutForListPC;
var
  i: Integer;
  /// /// групповое завершение сеанса
  NewSelectedPCShotDown: TThread;
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  i := MessageDlg('Завершить сеанс на выбранных компьютерах? ', mtWarning,
    [mbYes, mbCancel], 0);
  if i = mrYes then
    begin
      if ListView8.Items.Count > 0 then
        begin
          SelectedPCShutDown := TstringList.Create;
          SelectedPCShutDown := createListpcForCheck('');
          if SelectedPCShutDown.Count <> 0 then
            begin
              myShutdown := 0;
              NewSelectedPCShotDown :=
                SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
              NewSelectedPCShotDown.FreeOnTerminate := true;
              NewSelectedPCShotDown.Start;
            end
          else
            showmessage('Не выбран не один компьютер!!!');
        end;
    end;
end;

procedure TfrmDomainInfo.AllFormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  (Sender as TForm).Release;
end;

procedure TfrmDomainInfo.ButtonNoClose(Sender: TObject);
begin
  ((Sender as TButton).Parent as TForm).Close;
end;

procedure TfrmDomainInfo.CloseFormEsc(Sender: TObject; var Key: Char);
begin
  if Sender is TForm then
    if Key = #27 then
      (Sender as TForm).Close;
end;

procedure TfrmDomainInfo.LVAntivirusColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  LVAntivirus.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.LVAntivirusDblClick(Sender: TObject);
var
  liststr: TstringList;
begin
  if LVAntivirus.SelCount <> 1 then
    exit;
  liststr := TstringList.Create;
  try
    liststr.Add('Антивирус : ' + LVAntivirus.Selected.SubItems[1] + ' - ' +
      LVAntivirus.Selected.SubItems[2] + ' - ' +
      LVAntivirus.Selected.SubItems[3]);
    liststr.Add('Файрвол : ' + LVAntivirus.Selected.SubItems[4] + ' - ' +
      LVAntivirus.Selected.SubItems[5]);
    liststr.Add('Антишпион : ' + LVAntivirus.Selected.SubItems[6] + ' - ' +
      LVAntivirus.Selected.SubItems[7] + ' - ' +
      LVAntivirus.Selected.SubItems[8]);
    liststr.Add('Microsoft Defender : ' + LVAntivirus.Selected.SubItems[9] +
      ' - ' + LVAntivirus.Selected.SubItems[10]);
    itemLisToMemo(liststr, LVAntivirus.Selected.Caption, frmDomainInfo);
  finally
    liststr.Free;
  end;

end;

procedure TfrmDomainInfo.LVMicrosoftColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  LVMicrosoft.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.LVMicrosoftDblClick(Sender: TObject);
var
  ListOffice, LicStat, LicDes, ListKey, LicType, liststr: TstringList;
  i: Integer;
begin
  if LVMicrosoft.SelCount <> 1 then
    exit;
  liststr := TstringList.Create;
  ListOffice := TstringList.Create;
  LicStat := TstringList.Create;
  LicDes := TstringList.Create;
  ListKey := TstringList.Create;
  LicType := TstringList.Create;
  liststr.Add('Компьютер: ' + LVMicrosoft.Selected.SubItems[0]);
  liststr.Add('Операционная система: ' + LVMicrosoft.Selected.SubItems[1]);
  liststr.Add('Статус лицензии: ' + LVMicrosoft.Selected.SubItems[2]);
  liststr.Add('Описание статуса лицензии: ' + LVMicrosoft.Selected.SubItems[3]);
  liststr.Add('Ключ Windows: ' + LVMicrosoft.Selected.SubItems[4]);
  liststr.Add('Тип лицензии Windows: ' + LVMicrosoft.Selected.SubItems[5]);
  try
    ListOffice.CommaText := LVMicrosoft.Selected.SubItems[6];
    LicStat.CommaText := LVMicrosoft.Selected.SubItems[7];
    LicDes.CommaText := LVMicrosoft.Selected.SubItems[8];
    ListKey.CommaText := LVMicrosoft.Selected.SubItems[9];
    LicType.CommaText := LVMicrosoft.Selected.SubItems[10];
    liststr.Add('Продукты Office:');
    for i := 0 to ListOffice.Count - 1 do
      begin
        liststr.Add(ListOffice[i] + ' : ' + LicStat[i] + ' (' + LicDes[i] +
          ') :' + ListKey[i] + ' (' + LicType[i] + ')')
      end;
    itemLisToMemo(liststr, LVMicrosoft.Selected.SubItems[0], frmDomainInfo);
  finally
    ListOffice.Free;
    LicStat.Free;
    LicDes.Free;
    ListKey.Free;
    LicType.Free;
    liststr.Free;
  end;

end;

procedure TfrmDomainInfo.PowerOn;
var
  setini: TMemInifile;
  BRaddress: string;
  WOLPower: TThread;
begin
  // if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
  begin
    setini := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
      '\Settings.ini', TEncoding.Unicode);
    BRaddress := setini.readString('ConfLAN', 'broadcast', '192.168.0.255');
    IpBroadCast := InputBox('Broadcast адрес сети', 'IP-', BRaddress);
    if IpBroadCast = '' then
      begin
        showmessage('Вы не ввели broadcast адрес сети!');
        exit;
      end;
    if LabeledEdit7.text = '' then
      begin
        showmessage('Не указан MAC адрес!');
        exit;
      end;
    // if Assigned (ListMACAddress)=false then
    ListMACAddress := TstringList.Create;
    ListMACAddress.Add(LabeledEdit7.text);
    WOLPower := WOLThread.WOL.Create(true);
    WOLPower.FreeOnTerminate := true;
    /// / под вопросом самоуничтожение потока
    WOLPower.Start;
    setini.WriteString('ConfLAN', 'broadcast', IpBroadCast);
    setini.UpdateFile;
    setini.Free;
  end;
end;

procedure TfrmDomainInfo.ForceCloseSession;
var
  MyShutdownPC: TThread;
begin // принудительное завершение сеанса пользователя
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 1;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.ForcePowerOff;
var
  MyShutdownPC: TThread;
begin
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 5;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.SpeedButton60Click(Sender: TObject);
begin
  if LVAntivirus.Items.Count <> 0 then
    popupListViewSaveAs(LVAntivirus, 'Сохранение списка антивирусных продуктов',
      'Антивирусное ПО');
end;

procedure TfrmDomainInfo.SpeedButton61Click(Sender: TObject);
begin
  LVAntivirus.SelectAll;
end;

procedure TfrmDomainInfo.SpeedButton62Click(Sender: TObject);
begin
  readMicrosoftLic(GroupBox2.Caption, ListViewMicLic);
end;

procedure TfrmDomainInfo.SpeedButton63Click(Sender: TObject);
begin
  if ListViewMicLic.Items.Count <> 0 then
    begin
      popupListViewSaveAs(ListViewMicLic, 'Сохранение списка лицензий', '');
    end
  else
    showmessage('Список пуст!');
end;

procedure TfrmDomainInfo.SpeedButton64Click(Sender: TObject);
begin
  readAntivirusStatus(GroupBox2.Caption, ListViewAV);
end;

procedure TfrmDomainInfo.SpeedButton65Click(Sender: TObject);
begin
  if ListViewAV.Items.Count <> 0 then
    begin
      popupListViewSaveAs(ListViewAV, 'Сохранение списка Антивирусного ПО', '');
    end
  else
    showmessage('Список пуст!');
end;

procedure TfrmDomainInfo.SpeedButton66Click(Sender: TObject);
/// инвентаризация лицензий Microsoft из свойств компа
begin
  frmDomainInfo.InventoryMicrosoftPC(GroupBox2.Caption, LabeledEdit1.text,
    LabeledEdit2.text);
end;

procedure TfrmDomainInfo.SpeedButton67Click(Sender: TObject);
begin //
  if ping(GroupBox2.Caption) then
    ReadAdnWriteAntivirusStaus(GroupBox2.Caption, LabeledEdit1.text,
      LabeledEdit2.text, ListViewAV);
end;

procedure TfrmDomainInfo.SpeedButton68Click(Sender: TObject);
begin
if ping(ComboBox2.Text) then  showmessage('Узел доступен');
end;

procedure TfrmDomainInfo.SpeedButton69Click(Sender: TObject);
begin
if ConnectWMI(ComboBox2.Text,labeledEdit1.Text,labeledEdit2.Text) then showmessage('Сервер RPC доступен');
end;

procedure TfrmDomainInfo.ActivationWindowsPC;
begin
  MyPS := ComboBox2.text;
  if ping(MyPS) then
    begin
      ActivationWindows.showmodal;
    end;
end;

procedure TfrmDomainInfo.ActivationOffice;
begin
  MyPS := ComboBox2.text;
  if ping(MyPS) then
    begin
      ActivateOffice.showmodal;
    end;
end;

Function TfrmDomainInfo.InventoryMicrosoftPC(NamePC, user,
  pass: string): Boolean;
var
  i: Integer;
  ThreadMO: TThread;
begin
  try
    if not ping(NamePC) then
      exit;
    if not SolveExitInvMicrosoft then
      begin
        showmessage('Дождитесь завершения текущей инвентаризации...');
        exit;
        { i:=MessageDlg('Инвентаризация продуктов Windows и Office уже запущена.'#10#13+'Остановть текущую инвентаризацию?', mtConfirmation,[mbYes,mbCancel],0);
          if i=IDCancel then   exit
          else
          begin
          InventMicrosoft:=false; // признак остановки инвентаризации продуктов windows и office
          memo1.Lines.Add('Ожидайте завершения инвентаризации  продуктов Windows и Office.');
          while not SolveExitInvMicrosoft do ////если false то инвентаризация еще запущена логические переменные нужны для четкого определения остановки инвентаризации
          begin
          Application.ProcessMessages;
          end;
          memo1.Lines.Add('Инвентаризация  продуктов Windows и Office остановлена');
          end; }
      end;
    InventMicrosoft := true; // разрешить выполнение инвентаризации
    ListPCWO := TstringList.Create;
    ListPCWO.Add(NamePC);
    MyUser := user;
    MyPasswd := pass;
    Memo1.Lines.Add('Запускаю инвентаризацию  продуктов Windows и Office');
    ThreadMO := InventoryWindowsOffice.InventoryMicrosoft.Create(true);
    ThreadMO.FreeOnTerminate := true;
    ThreadMO.Start;
    Result := true;
  except
    on E: Exception do
      begin
        Result := false;
        Memo1.Lines.Add
          ('Ошибка запуска инвентаризации продуктов Windows и Office - "' +
          E.Message + '"');
      end;
  end;
end;

Procedure TfrmDomainInfo.InventoryMicrosoft;
var
  i: Integer;
begin
  try
    if not SolveExitInvMicrosoft then
      begin
        showmessage('Дождитесь завершения текущей инвентаризации...');
        exit;
      end;
    InventMicrosoft := true; // разрешить выполнение инвентаризации
    ListPCWO := TstringList.Create;
    ListPCWO := createListpcForCheck('');
    /// / создание списка выделенных компьютеров
    if ListPCWO.Count = 0 then
      ListPCWO := createListpc(''); // если выделенных нет то выделяем все
    if ListPCWO.Count <> 0 then
      begin
        MyUser := LabeledEdit1.text;
        MyPasswd := LabeledEdit2.text;
        Memo1.Lines.Add('Запускаю инвентаризацию  продуктов Windows и Office');
        ThreadMO := InventoryWindowsOffice.InventoryMicrosoft.Create(true);
        ThreadMO.FreeOnTerminate := true;
        ThreadMO.Start;
      end;
  except
    on E: Exception do
      begin
        Memo1.Lines.Add
          ('Ошибка запуска инвентаризации продуктов Windows и Office - "' +
          E.Message + '"');
      end;
  end;
end;

procedure TfrmDomainInfo.InventoryHardware;
begin
  if not Datam.ConnectionDB.Connected then
    begin
      showmessage('Для инвентаризации подключите базу данных!!!');
      exit;
    end;
  if ComboBox2.text <> '' then
    begin
      refreshinfoPC(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text);
    end
  else
    showmessage('Не выбран компьютер!');
end;

procedure TfrmDomainInfo.SpeedButton6Click(Sender: TObject);
begin
  try
    if not ClientSocket1.Active then
      begin
        showmessage('Установите соединение с сервером!');
        exit;
      end;
    MRDUViewerMain.SmartSizing := not MRDUViewerMain.SmartSizing;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Повторите подключение, ошибка :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton10Click(Sender: TObject);
begin
  try // отключение от сервера
    if ConnectionEnable then
      begin
        ClientSocket1.Active := false;
        ClientSocket1.Close;
        MRDUViewerMain.Disconnect;
        MRDUViewerMain.SmartSizing := false;
        if Assigned(MRDUViewerMain) then
          FreeAndNil(MRDUViewerMain);
        AccessSettingLevel := false;
        ConnectionEnable := false;
      end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка при отключении  :' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton11Click(Sender: TObject);
var
/// / отправить сочетание клавиш Alt+Ctrl+Del
  s: string;
begin
  if ClientSocket1.Active then
    begin
      Memo1.Lines.Add('Диспетчер задач');
      s := WideString('<Shift+Ctrl+Esc>');
      s := Crypt(s);
      /// / шифрация текста
      ClientSocket1.Socket.SendText(s);
      /// отправляем зашифрованное сообщение
    end
  else
    showmessage('установите соединение с сервером!');

end;

procedure TfrmDomainInfo.OpenFormForuRDM;
var
  i: Integer;
begin
  MRDForm.LabeledEdit1.Clear;
  for i := 0 to ComboBox2.Items.Count - 1 do
    begin
      MRDForm.LabeledEdit1.Items.Add(ComboBox2.Items[i]);
      MRDForm.LabeledEdit1.ItemsEx[i].ImageIndex := ComboBox2.ItemsEx[i]
        .ImageIndex;
    end;
  MRDForm.LabeledEdit1.text := ComboBox2.text;
  MRDForm.EditUser.text := LabeledEdit4.text;
  MRDForm.EditPass.text := LabeledEdit5.text;
  MRDForm.Editport.text := LabeledEdit6.text;
  if MRDForm.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    MRDForm.WindowState := wsNormal
  else
    MRDForm.Show;
end;

/// ////////////////////////////////////////////////////////uRDM конец

/// //////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.SpeedButton13Click(Sender: TObject);
var
  i, Res: Integer;
  /// / инвентаризация оборудования
begin
  try
    // if listview8.Items.Count>0 then
    begin
      if not SolveExitInvConf then
        begin
          showmessage('Дождитесь завершения текущей инвентаризации...');
          exit;
        end;
      ListPCConf := TstringList.Create;
      for i := 0 to ListView8.Items.Count - 1 do
        begin
          if ListView8.Items[i].Checked = true then
            ListPCConf.Add(ListView8.Items[i].SubItems[0])
        end;
      if ListPCConf.Count = 0 then
        begin
          if ListView8.Items.Count = 0 then
            begin
              MainPage.ActivePageIndex := 0;
              showmessage('Нет списка компьютеров');
              ListPCConf.Free;
              exit;
            end;
          ListPCConf := createListpc('');
        end;
      Memo1.Lines.Add('Запускаю инвентаризаци оборудования');
      InventConf := true;
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      ThrInvPC := MyInventoryPC.InventoryConfig.Create(true);
      ThrInvPC.FreeOnTerminate := true; // поток уничтожится сам после его остановки
      ThrInvPC.Start;
    end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Start invent Config - ' + E.Message)
      end;
  end;
end;

/// /////////////////////////////////////////////////////////////
function delstr(s: string): string;
begin
  Delete(s, 1, 4);
  Delete(s, (Length(s)), 1);
  Result := s;
end;

procedure TfrmDomainInfo.Chkdsk1Click(Sender: TObject); // check disk
begin
  if (ListViewDisk.Items.Count <> 0) and (ListViewDisk.SelCount <> 0) then
    begin
      if ListViewDisk.Items[ListViewDisk.ItemIndex].SubItems[5] <> '' then
        begin
          SelectedHDD := delstr(ListViewDisk.Items[ListViewDisk.ItemIndex]
            .SubItems[5]);
          MyPS := ComboBox2.text;
          MyUser := LabeledEdit1.text;
          MyPasswd := LabeledEdit2.text;
          OKRightDlg123.showmodal;
        end;
    end;
end;

procedure TfrmDomainInfo.ComboBox1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox1.DroppedDown := true;
end;

/// ////////////////////////////////////////

procedure TfrmDomainInfo.ComboBox23KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox2.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBox2Select(Sender: TObject);
var
  i: Integer;
begin
  // if Combobox2.style<>csExDropDown then Combobox2.style:=csExDropDown;
  MyPS := ComboBox2.text;
  LabeledEdit7.text := '';
  Label3.Caption := '';
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if ComboBox2.text = ListView8.Items[i].SubItems[0] then
        begin
          LabeledEdit7.text := ListView8.Items[i].SubItems[2];
          Label3.Caption := ListView8.Items[i].SubItems[3];
          ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex :=
            ListView8.Items[i].ImageIndex;
          break;
        end;
    end;
end;

procedure TfrmDomainInfo.ComboBoxAllClick(Sender: TObject);
var
  i: Integer;
begin
  // if ComboBox2.Style<>csExDropDownList then  ComboBox2.Style:=csExDropDownList;
  if ComboBox2.Items.Count > 30 then
    begin
      for i := ComboBox2.Items.Count - 1 downto 30 do
        ComboBox2.Items.Delete(i);
    end;
end;

procedure TfrmDomainInfo.ComboBox2Select2(Sender: TObject);
var
  i: Integer;
begin
  // if ComboBox2.Style<>csExDropDownList then  ComboBox2.Style:=csExDropDownList;
  MyPS := ComboBox2.text;
  LabeledEdit7.text := '';
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if ComboBox2.text = ListView8.Items[i].SubItems[0] then
        begin
          LabeledEdit7.text := ListView8.Items[i].SubItems[2];
          ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex :=
            ListView8.Items[i].ImageIndex;
          break;
        end;
    end;
end;

procedure TfrmDomainInfo.ComboBox2KeyUp2(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
    begin
      MyPS := ComboBox2.text;
      // ComboBox2Select2(self);
    end
  else
    ComboBox2.text := 'Ограничение - 30 ПК';
end;

procedure TfrmDomainInfo.ComboBox2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
    begin
      MyPS := ComboBox2.text;
      ComboBox2Select(self);
    end;
end;

procedure TfrmDomainInfo.ComboBox23Select(Sender: TObject);
var
  i: Integer;
begin
  MyPS := ComboBox2.text;
  LabeledEdit7.text := '';
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if ComboBox2.text = ListView8.Items[i].SubItems[0] then
        LabeledEdit7.text := ListView8.Items[i].SubItems[2]
    end;
end;

procedure TfrmDomainInfo.ComboBox2DropDown(Sender: TObject);
var
/// / обновление статуса всех компьютеров при раскрытии списка
  i, z: Integer;
begin
  try
    if ComboBox2.Tag = 1 then
      exit;
    /// не выполняется обновление списка если выбирать управление из списка компьютеров, иначе глюки
    for z := 0 to ComboBox2.Items.Count - 1 do
      for i := 0 to ListView8.Items.Count - 1 do
        begin
          if ComboBox2.Items[z] = ListView8.Items[i].SubItems[0] then
            begin
              ComboBox2.ItemsEx[z].ImageIndex := ListView8.Items[i].ImageIndex;
              break;
            end;
        end;
  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка обновления статуса компьютеров:' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.ComboBox2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  // ComboBox2.DroppedDown:=true;
end;

/// /////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.Dismount1Click(Sender: TObject);
var
  NewHDDDismount: TThread;
begin
  if (ListViewDisk.Items.Count <> 0) and (ListViewDisk.SelCount <> 0) then
    begin
      if ListViewDisk.Items[ListViewDisk.ItemIndex].SubItems[5] <> '' then
        begin
          SelectedHDD := delstr(ListViewDisk.Items[ListViewDisk.ItemIndex]
            .SubItems[5]);
          MyPS := ComboBox2.text;
          MyUser := LabeledEdit1.text;
          MyPasswd := LabeledEdit2.text;
          OKRightDlg123.showmodal;
          NewHDDDismount := Dismount.DismountVolume.Create(true);
          NewHDDDismount.FreeOnTerminate := true;
          NewHDDDismount.Start;
        end;
    end;
  // SelectedHDD:=delstr(SelectHDD[DiskDrive.PopupComponent.Tag]);
  // NewHDDDismount:=Dismount.DismountVolume.Create(false);
end;

procedure TfrmDomainInfo.WMNotify(var AMessage: TWMNotify);

begin
  if AMessage.NMHdr^.idFrom = HeaderID then
    if AMessage.NMHdr^.Code = HDN_ITEMSTATEICONCLICK then
      // begin
      // for I := 0 to lvWorkStation.Items.Count-1 do
      // begin
      // lvWorkStation.Items[i].Checked:=true;
      // button1.Caption:='Снять';
      // end;
      // end;
      // ShowMessage('Тык на чекбоксе в заголовке');
      inherited;
end;

procedure TfrmDomainInfo.FlowPanelPrintMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ScrollBox2.SetFocus;
end;

procedure TfrmDomainInfo.Format1Click(Sender: TObject);
begin
  if (ListViewDisk.Items.Count <> 0) and (ListViewDisk.SelCount <> 0) then
    begin
      if ListViewDisk.Items[ListViewDisk.ItemIndex].SubItems[5] <> '' then
        begin
          SelectedHDD := delstr(ListViewDisk.Items[ListViewDisk.ItemIndex]
            .SubItems[5]);
          MyPS := ComboBox2.text;
          MyUser := LabeledEdit1.text;
          MyPasswd := LabeledEdit2.text;
          OKRightDlg12345.showmodal;
        end;
    end;
end;

procedure TfrmDomainInfo.FormClose(Sender: TObject; var Action: TCloseAction);
var
  IniSettings: TMemInifile;
begin
  IniSettings := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
    '\Settings.ini', TEncoding.Unicode);
  IniSettings.WriteBool('AutoCheck', 'Check', CheckBox1.Checked);
  IniSettings.UpdateFile;
  IniSettings.Free;

end;



procedure TfrmDomainInfo.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  i: byte;
  coltask,z: Integer;
begin
  if OutForPing then // если скарирование запущено, пытаемся остановить
    begin
      Memo1.Lines.Add('Остановка сканирования сети, ожидайте...');
      OutForPing:=false; // признак остановки сканирования
      while not SolveExitInvScan do
        begin
         Application.ProcessMessages; // ожидаем остановки сканирования/
        end;
       try
        freeandnil(SelectedPCPing);
       except on E: Exception do Memo1.Lines.add('Kill thread scan '+e.Message);
       end;
      Memo1.Lines.Add('Остановка сканирования сети завершена');
      SpeedButton39.Visible := true;
      SpeedButton40.Visible := false;
    end;
  coltask := 0;
  coltask := EditTask.ThereAnyRunStopTask(true); // узнаем количество запущенных задач
  if coltask <> 0 then
    begin
      i := MessageDlg('В данный момент выполняются задачи (' + inttostr(coltask)
        + ')' + #10#13 +
        'Для выхода из программы остановите выполнение задач!!!' + #10#13 +
        'Вы уверены что хотите закрыть программу???', mtWarning,
        [mbCancel, mbYes], 0);
      if i = mrCancel then
        begin
          CanClose := false;
          exit;
        end;
    end;

  if (SolveExitInvConf) and (SolveExitInvSoft) and (SolveExitInvMicrosoft) then  // если инвентаризация не запущена.
    Begin
    Timer1.Enabled:=false;
    CloseHandle(HM);// убиваем мьютекс для разрешения копий
    CanClose := true; // закрыть программу
    End
  else
    Begin // иначе открыть диалог для завершения инвентаризации
       z:=MainFormPopup.DlgCloseForm('Остановка инвентаризации перед закрытием программы','Остановить','Остановить принудительно','Отмена');
       case z of
       1:  //завершить
         begin
            InventConf := false; // признак остановки потока
            InventSoft := false; // признак остановки потока
            InventMicrosoft := false; // признак остановки потока
            Memo1.Lines.Add('Ожидайте остановки инвентаризации...');
            while (not SolveExitInvConf) and (not SolveExitInvSoft) and
              (not SolveExitInvMicrosoft) do // логические переменные нужны для четкого определения остановки инвентаризации
              begin
                Application.ProcessMessages;
              end;
          Timer1.Enabled:=false;
          CloseHandle(HM);// убиваем мьютекс для разрешения копий
          CanClose := true;//разрешить закрыть форму
         end;
       3: // принудительно завершить
         begin
          try
            InventConf := false; // признак остановки потока
            InventSoft := false; // признак остановки потока
            InventMicrosoft := false; // признак остановки потока
            SolveExitInvConf:=true;
            SolveExitInvSoft:=true;
            SolveExitInvMicrosoft:=true;
            if not SolveExitInvConf then FreeAndNil(ThrInvPC);
            if not SolveExitInvSoft then freeandnil(ThrInvSoft);
            if not SolveExitInvMicrosoft then freeandnil(ThreadMO);
            except on E: Exception do
             Memo1.Lines.Add('Kill thread inventory '+e.Message);
            end;
          Timer1.Enabled:=false;
          CloseHandle(HM);// убиваем мьютекс для разрешения копий
          CanClose := true;//разрешить закрыть форму
         end;
       2:  //отмена
         begin
         CanClose := false /// запретить закрывать форму
         end;
       end;
    End;
end;

function loadlistpc(s: string): Boolean;
var
/// функция загрузки списка компьютеров из файла при загрузке программы
  listPC: TstringList;
  i: Integer;
begin
  try
    listPC := TstringList.Create;
    listPC.LoadFromFile(s);
    for i := 0 to listPC.Count - 1 do
      begin
        frmDomainInfo.ComboBox2.Items.Add(listPC[i]);
        frmDomainInfo.ComboBox2.ItemsEx[i].ImageIndex := 0;
        with frmDomainInfo.ListView8.Items.Add do
          begin
            Caption := '';
            SubItems.Add(listPC[i]);
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
            SubItems.Add('');
          end;
      end;
    frmDomainInfo.ComboBox1.Clear;
    frmDomainInfo.ComboBox1.Enabled := false;
    Result := true;
  except
    Result := false;
  end;
  listPC.Free;
end;

function TfrmDomainInfo.createListpc(s: string): TstringList;
var
  i: Integer;
begin
  Result := TstringList.Create;
  Result.Clear;
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if ListView8.Items[i].ImageIndex <> 12 then
        Result.Add(ListView8.Items[i].SubItems[0]);
    end;
end;

function TfrmDomainInfo.createListpcForCheck(s: string): TstringList;
var
  i: Integer;
begin
  Result := TstringList.Create;
  Result.Clear;
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if (ListView8.Items[i].ImageIndex <> 12) and (ListView8.Items[i].Checked)
      then
        Result.Add(ListView8.Items[i].SubItems[0]);
    end;
end;

/// ///////////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.findElement;
var
  i: Integer;
  CurentNamePC: string;
begin
  frmDomainInfo.Caption := captionMainForms;
  EditTask.Button7.OnClick:= EditTask.SaveTask;
  {if ListView8.Items.Count > 30 then
    begin
      for i := ListView8.Items.Count - 1 downto 30 do
        ListView8.Items[i].Delete;
    end;
  ListView8.OnChange := ListViewallChange;
  CurentNamePC := ComboBox2.text;
  /// чтобы имя пк в cjmbobox не менялся    , запоминаем, потом обратно вставим
  for i := ComboBox2.Items.Count - 1 downto 0 do
  /// чистим combobox
    ComboBox2.Items.Delete(i);

  for i := 0 to ListView8.Items.Count - 1 do
  /// / заполняем combobox
    begin
      ComboBox2.Items.Add(ListView8.Items[i].SubItems[0]);
      ComboBox2.ItemsEx[i].ImageIndex := 0;
    end;
  SpeedButton85.OnClick := ConnectRDPOsherWindows;
  RDP1.OnClick := RDP2Click;
  ComboBox2.text := CurentNamePC;
  /// а вот тут вставляем еэто имя
  ComboBox2.OnClick := ComboBoxAllClick;
  ComboBox2.OnSelect := ComboBox2Select2;
  ComboBox2.OnKeyUp := ComboBox2KeyUp2;
  ComboBox1.Tag := 1;
  ComboBox1.Enabled := false;
  ComboBox8.Enabled := false;}
end;
/// ////////////////////////////////////////////////////////////////////////////////////

procedure TfrmDomainInfo.FormCreate(Sender: TObject);
var
  i: Integer;
  setini: TMemInifile;
  s: WideString;
  confLAN, confIP, confmask, confFile, SelectGroupAD: string;
  ItemsNew, ItemsNew1: Array [1 .. 8] of TMenuItem;
  StartScan, UserGroup, PCGroup: bool;
begin
  // ReportMemoryLeaksOnShutdown := True;
  frmDomainInfo.width := screen.width;
  frmDomainInfo.height := screen.height;
  MainPage.TabIndex := 0;
  /// / призапуске активна вкладка управление
  pcRes.ActivePageIndex := 0;
  PCInvent.TabIndex := 0;
  PagePropertiesPC.TabIndex := 0;
  sort := -1;
  /// / сортировка в ListView
  ListSoftID := TstringList.Create;
  /// список ID установленных программ

  if FileExists(ExtractFilePath(Application.ExeName) + '\Settings.ini') then
    begin
      setini := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
        '\Settings.ini', TEncoding.Unicode);
      try
        TStyleManager.TrySetStyle(setini.readString('Style', 'ST',
          'Silver'), true);
      except
        TStyleManager.TrySetStyle('Windows', true);
      end;
      TabSheet2.TabVisible := true;
      Traymin := setini.ReadBool('Tray', 'ico', false);
      /// / сворачивание в трей
      TabSheet3.TabVisible := setini.ReadBool('View', 'Services', true);
      TabSheet1.TabVisible := setini.ReadBool('View', 'Autoload', true);
      TabSheet4.TabVisible := setini.ReadBool('View', 'ProgramMSI', true);
      TabSheet5.TabVisible := setini.ReadBool('View', 'AllProgram', true);
      TabSheet6.TabVisible := setini.ReadBool('View', 'Drive', true);
      TabSheet7.TabVisible := setini.ReadBool('View', 'Printer', true);
      TabSheet8.TabVisible := setini.ReadBool('View', 'Profiles', true);
      TabSheet10.TabVisible := setini.ReadBool('View', 'DriverPnP', true);
      TabSheet11.TabVisible := setini.ReadBool('View', 'Network', true);
      TabSheet9.TabVisible := setini.ReadBool('View', 'RDP', true);
      TabSheet12.TabVisible := setini.ReadBool('View', 'uRDM', true);
      TabSheet13.TabVisible := setini.ReadBool('View', 'HotFix', true);
      TabSheet22.TabVisible := setini.ReadBool('View', 'Share', true);
      TabSheet23.TabVisible := setini.ReadBool('View', 'EventWindows', true);

      LabeledEdit6.text := setini.readString('MRD', 'port', '48999');
      ClientSocket1.Port := setini.ReadInteger('MRD', 'LocalPort', 48998);
      LabeledEdit1.text := setini.readString('ADM', 'user', '');
      s := WideString(setini.readString('ADM', 'passwd', ''));
      /// / расщифровка пароля
      Code(s, '1234', true);
      /// / расщифровка пароля
      LabeledEdit2.text := s;
      LabeledEdit4.text := setini.readString('MRD', 'user', 'Administrator');
      s := WideString(setini.readString('MRD', 'passwd', ''));
      /// / расщифровка пароля
      Code(s, '1234', true);
      /// / расщифровка пароля
      LabeledEdit5.text := s;
      ///// контроллер домена
      CurrentDC:=''; CurrentDC:=setini.readString('ADM', 'DC', '');
      /// / переменные для подключения к БД
      databaseProtocol := setini.readString('DB', 'Protocol', 'local');
      databaseServer := setini.readString('DB', 'Server', 'localhost');
      databaseport := setini.readString('DB', 'Port', '3050');
      databaseName := setini.readString('DB', 'Patch',
        ExtractFilePath(Application.ExeName) + 'DB.FDB');
      databaseDriverID := setini.readString('DB', 'DrID', 'FB');
      databaseUserName := setini.readString('DB', 'user', 'SYSDBA');
      s := WideString(setini.readString('DB', 'pass', 'masterkey'));
      /// / расщифровка пароля
      Code(s, '1234', true);
      /// / расщифровка пароля
      databasePassword := s;
      databaseconnected := setini.ReadBool('db', 'Connected', true);
      s := '';
      /// ////////////////////////////////
      StartScan := setini.ReadBool('Scan', 'ScanLan', true);
      // если скарируем сеть со старта  программы
      pingtimeout := setini.ReadInteger('Scan', 'timeout', 3000);
      pingtype:= setini.ReadInteger('Scan', 'type', 1);
      /// таймаут для пинга
      RPCport := setini.ReadBool('Scan', '135', false);
      urdmport := setini.readString('MRD', 'port', '48999');
      /// порт uRDM сервера
      uRDMScan := setini.ReadBool('MRD', 'scan', true);
      /// сканировать на наличие uRDM
      SelectGroupAD := setini.readString('ConfLAN', 'Group', '');
      setini.WriteBool('InvWarning', 'Proc', true);
      /// процессор всегда учитывается при инвентаризации
      if setini.ReadBool('InvWarning', 'MAC', false) then
        InvWarning := InvWarning + 'NETWORK_MAC,';
      if setini.ReadBool('InvWarning', 'Inv№', true) then
        InvWarning := InvWarning + 'INV_NUMBER,';
      if setini.ReadBool('InvWarning', 'OS', true) then
        InvWarning := InvWarning + 'OS_DATEINSTALL,';
      if setini.ReadBool('InvWarning', 'Mon', true) then
        InvWarning := InvWarning + 'MONITOR_NAME,';
      if setini.ReadBool('InvWarning', 'Net', false) then
        InvWarning := InvWarning + 'NETWORKINTERFACE,';
      if setini.ReadBool('InvWarning', 'Video', true) then
        InvWarning := InvWarning + 'VIDEOCARD,';
      if setini.ReadBool('InvWarning', 'HDD', true) then
        InvWarning := InvWarning + 'HDD_NAME,HDD_SN,';
      if setini.ReadBool('InvWarning', 'RAM', true) then
        InvWarning := InvWarning + 'SUMM_MEM_SIZE,';
      if setini.ReadBool('InvWarning', 'M/B', true) then
        InvWarning := InvWarning + 'MONTERBOARD,MONTERBOARD_SN,';
      if setini.ReadBool('InvWarning', 'Proc', true) then
        InvWarning := InvWarning + 'PROCESSOR,';
      if setini.ReadBool('InvWarning', 'Name', true) then
        InvWarning := InvWarning + 'PC_NAME,';
      if setini.ReadBool('InvWarning', 'SMART.OS', true) then
        InvWarning := InvWarning + 'HDD_SMART_OS,';
      Delete(InvWarning, Length(InvWarning), 1);
      // удаляем последнюю запятую в строке
      if setini.ReadBool('InvWarning', 'SoftVer', true) then
        InvSoftWarning := InvSoftWarning + 'SOFT_VERSION,';
      if setini.ReadBool('InvWarning', 'SoftManufacture', true) then
        InvSoftWarning := InvSoftWarning + 'MANUFACTURE,';
      if setini.ReadBool('InvWarning', 'SoftName', true) then
        InvSoftWarning := InvSoftWarning + 'SOFT_NAME,';
      Delete(InvSoftWarning, Length(InvSoftWarning), 1);
      // удаляем последнюю запятую в строке

      // загружаем список компьютеров который был открыт до закрытия проги
      confLAN := setini.readString('ConfLAN', 'type', 'AD');
      UserGroup := setini.ReadBool('ConfLAN', 'UserGroup', true);    // Выгружать только пользователей выбранной группы безопасности домена
      PCGroup := setini.ReadBool('ConfLAN', 'PCGroup', true);     // Выгружать только компьютеры выбранной группы безопасности домена
      if confLAN = 'IP' then
        begin
          confIP := setini.readString('ConfLAN', 'IP', '192.168.0.0');
          confmask := setini.readString('ConfLAN', 'subnet', '255.255.255.0');
        end;
      if confLAN = 'file' then
        confFile := setini.readString('ConfLAN', 'file', '');
      CheckBox1.Checked := setini.ReadBool('AutoCheck', 'Check', true);
      /// / Автоматическое считывание конфигурации в инвентаризации компьютеров
    end
  else
    begin
      setini := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
        '\Settings.ini', TEncoding.Unicode);
      setini.WriteBool('View', 'Services', true);
      setini.WriteBool('View', 'Autoload', true);
      setini.WriteBool('View', 'ProgramMSI', true);
      setini.WriteBool('View', 'AllProgram', true);
      setini.WriteBool('View', 'Drive', true);
      setini.WriteBool('View', 'Printer', true);
      setini.WriteBool('View', 'Profiles', true);
      setini.WriteBool('View', 'DriverPnP', true);
      setini.WriteBool('View', 'Network', true);
      setini.WriteBool('View', 'RDP', true);
      setini.WriteBool('View', 'uRDM', true);
      setini.WriteBool('ConfLAN', 'UserGroup', true);
      // Выгружать только пользователей выбранной группы безопасности домена
      setini.WriteBool('ConfLAN', 'PCGroup', true);
      // Выгружать только компьютеры выбранной группы безопасности домена
      setini.WriteInteger('MRD', 'port', 48999);
      setini.WriteBool('MRD', 'scan', true);
      setini.WriteString('MRD', 'user', 'Administrator');
      setini.WriteString('MRD', 'passwd', '');
      setini.WriteString('ADM', 'user', '');
      setini.WriteString('ADM', 'passwd', '');
      ClientSocket1.Port := 48998;
      LabeledEdit6.text := inttostr(48999);
      setini.WriteBool('Scan', 'ScanLan', true);
      setini.WriteString('Style', 'ST', 'Silver');
      setini.WriteBool('SMART', 'read', true);
      setini.WriteInteger('Scan', 'timeout', 3000);
      /// база данных
      setini.WriteString('DB', 'Patch', ExtractFilePath(Application.ExeName) +
        'DB.FDB');
      setini.WriteString('DB', 'DrID', 'FB');
      setini.WriteBool('db', 'Connected', true);
      setini.WriteString('DB', 'user', 'SYSDBA');
      s := WideString('masterkey'); // шифрация пароля
      Code(s, '1234', false);
      setini.WriteString('DB', 'pass', string(s));
      setini.UpdateFile;
      s := '';
      /// / записали в файл, считали настройки по умолчанию для БД
      /// //// переменные для подключения к БД
      databaseProtocol := setini.readString('DB', 'Protocol', 'local');
      databaseServer := setini.readString('DB', 'Server', 'localhost');
      databaseport := setini.readString('DB', 'Port', '3050');
      databaseName := setini.readString('DB', 'Patch',
        ExtractFilePath(Application.ExeName) + 'DB.FDB');
      databaseDriverID := setini.readString('DB', 'DrID', 'FB');
      databaseUserName := setini.readString('DB', 'user', 'SYSDBA');
      s := WideString(setini.readString('DB', 'pass', ''));
      /// / расщифровка пароля
      Code(s, '1234', true);
      /// / расщифровка пароля
      databasePassword := s;
      databaseconnected := setini.ReadBool('db', 'Connected', true);
      s := '';
      /// ///////////////////////////////////////////////
      pingtimeout := setini.ReadInteger('Scan', 'timeout', 3000);
      pingtype:= setini.ReadInteger('Scan', 'type', 1);
      RPCport := setini.ReadBool('Scan', '135', false);
      urdmport := setini.readString('MRD', 'port', '48999');
      StartScan := setini.ReadBool('Scan', 'ScanLan', false);
      uRDMScan := setini.ReadBool('MRD', 'scan', true);
      try
        TStyleManager.TrySetStyle(setini.readString('Style', 'ST',
          'Windows'), true);
      except
        TStyleManager.TrySetStyle('Windows', true);
      end;
      /// //////////////////////////////////////
      setini.WriteBool('InvWarning', 'MAC', false);
      setini.WriteBool('InvWarning', 'Inv№', true);
      setini.WriteBool('InvWarning', 'OS', true);
      setini.WriteBool('InvWarning', 'Mon', true);
      setini.WriteBool('InvWarning', 'Net', false);
      setini.WriteBool('InvWarning', 'Video', true);
      setini.WriteBool('InvWarning', 'HDD', true);
      setini.WriteBool('InvWarning', 'RAM', true);
      setini.WriteBool('InvWarning', 'M/B', true);
      setini.WriteBool('InvWarning', 'Proc', true);
      setini.WriteBool('InvWarning', 'Name', true);
      setini.WriteBool('InvWarning', 'SoftVer', true);
      setini.WriteBool('InvWarning', 'SoftManufacture', true);
      setini.WriteBool('InvWarning', 'SoftName', true);
      // запись в настройки тип сети AD, IP или file
      setini.WriteString('ConfLAN', 'type', 'AD');

      if setini.ReadBool('InvWarning', 'MAC', false) then
        InvWarning := InvWarning + 'ANSWER_MAC,';
      if setini.ReadBool('InvWarning', 'Inv№', true) then
        InvWarning := InvWarning + 'INV_NUMBER,';
      if setini.ReadBool('InvWarning', 'OS', true) then
        InvWarning := InvWarning + 'OS_DATEINSTALL,';
      if setini.ReadBool('InvWarning', 'Mon', true) then
        InvWarning := InvWarning + 'MONITOR_NAME,';
      if setini.ReadBool('InvWarning', 'Net', false) then
        InvWarning := InvWarning + 'NETWORKINTERFACE,';
      if setini.ReadBool('InvWarning', 'Video', true) then
        InvWarning := InvWarning + 'VIDEOCARD,';
      if setini.ReadBool('InvWarning', 'HDD', true) then
        InvWarning := InvWarning + 'HDD_NAME,HDD_SN,';
      if setini.ReadBool('InvWarning', 'RAM', true) then
        InvWarning := InvWarning + 'SUMM_MEM_SIZE,';
      if setini.ReadBool('InvWarning', 'M/B', true) then
        InvWarning := InvWarning + 'MONTERBOARD,MONTERBOARD_SN,';
      if setini.ReadBool('InvWarning', 'Proc', true) then
        InvWarning := InvWarning + 'PROCESSOR,';
      if setini.ReadBool('InvWarning', 'Name', true) then
        InvWarning := InvWarning + 'PC_NAME,';
      if setini.ReadBool('InvWarning', 'SMART.OS', true) then
        InvWarning := InvWarning + 'HDD_SMART_OS,';
      Delete(InvWarning, Length(InvWarning), 1);
      // удаляем последнюю запятую в строке
      if setini.ReadBool('InvWarning', 'SoftVer', true) then
        InvSoftWarning := InvSoftWarning + 'SOFT_VERSION,';
      if setini.ReadBool('InvWarning', 'SoftManufacture', true) then
        InvSoftWarning := InvSoftWarning + 'MANUFACTURE,';
      if setini.ReadBool('InvWarning', 'SoftName', true) then
        InvSoftWarning := InvSoftWarning + 'SOFT_NAME,';
      Delete(InvSoftWarning, Length(InvSoftWarning), 1);
      // удаляем последнюю запятую в строке
      StartScan := setini.ReadBool('Scan', 'ScanLan', true);
      UserGroup := setini.ReadBool('ConfLAN', 'UserGroup', true);
      // Выгружать только пользователей выбранной группы безопасности домена
      PCGroup := setini.ReadBool('ConfLAN', 'PCGroup', true);
      // Выгружать только компьютеры выбранной группы безопасности домена
      setini.UpdateFile;
      /// обновляем ini файл и сохраняем
    end;
  setini.Free;

  /// ///////////////////////////////////////////////
  // Просто вызываем все функции подряд (не делал проверок на результат функций)
  // ledCompName.Text := GetCurrentComputerName;
  // ledDomainName.Text := CurrentDomainName;
  // MyDCName.Text := GetDomainController(CurrentDomainName);
  // Единственно, если нет контролера домена, то дальше определять бесполезно
  // если последний раз открывали список компьютеров из AD то открываем его
  if (GetDomainController(CurrentDomainName) <> '') and (confLAN = 'AD') then
    begin
     // Memo1.Lines.add('--------------'+ GetDNSDomainName(CurrentDomainName));
      // EnumAllTrustedDomains; // Все доверенные домены, нах не надо
      GetCurrentComputerName;
      // информация о компьютере  в результате должно быть имя, но главное получваем в переменную CurrentDomainName имя домена
      LabeledEdit3.text := CurrentDomainName;
      // имя домена в LabelEdEdit на главную форму
      EnumAllGroups; // получаем список всех групп домена
      if SelectGroupAD = '' then
        SelectGroupAD := 'Компьютеры домена';
      // если группа безопасности в ini файле не указана то назначаем ему группу по уморлчанию
      if UserGroup then
        GetAllGroupUsersInComboBox(SelectGroupAD)
        // если выгружать пользователей только выбранной группы
      else
        EnumAllUsers;
      /// / иначе получаем список всех пользователей домена

      if PCGroup then
      // если выбрано загружать только компьютеры выбранной группы безопасности
        begin
          GetPCForGroupToLisViewComboBox(SelectGroupAD);
          ComboBox1.text := SelectGroupAD;
        end
      else // иначе выгружаем все ПК в Combobox на вкладке Управление и список ПК выбранной группы в ListView  на вкладке Компьютеры
        if ComboBox1.Items.Count <> 0 then
          BEGIN
            EnumAllWorkStation;
            // передаем все рабочие станции в combobox в разделе Компьютеры
            for i := 0 to ComboBox1.Items.Count do
            // выгружаем список компьютеров выбранной группы безопасности
              Begin
                if ComboBox1.Items[i] = SelectGroupAD then
                  begin
                    ComboBox1.text := SelectGroupAD;
                    ComboBox1.OnSelect(ComboBox1);
                    /// onSelect -процедура, загружает список компов
                    break;
                  end;
                if ComboBox1.Items[i] = 'Domain Computers' then
                  begin
                    ComboBox1.text := 'Domain Computers';
                    ComboBox1.OnSelect(ComboBox1);
                    break;
                  end;
              End;
          END;
    end;
  // если последний раз открывали список ip адресов
  if confLAN = 'IP' then
    begin
      if not loadlistpc(ExtractFilePath(Application.ExeName) + '\listIP.txt')
      then
        showmessage('Что то пошло не так при загрузке' + #10#13 +
          ' предыдущего списка компьютеров!');
    end;
  // если последний раз открывали список компьютеров из файла
  if confLAN = 'file' then
    begin
      if not loadlistpc(ExtractFilePath(Application.ExeName) + '\listPC.txt')
      then
        showmessage('Что то пошло не так при загрузке' + #10#13 +
          ' предыдущего списка компьютеров!');
    end;

  for i := 1 to 6 do
  /// // создание подпунктов popupmenu
    begin
      ItemsNew[i] := TMenuItem.Create(PopupMenu1);
      PopupMenu1.Items[1].Add(ItemsNew[i]);
      PopupMenu1.Items[1].Items[i - 1].Caption := 'Пункт' + inttostr(i);
      PopupMenu1.Items[1].Items[i - 1].Tag := i - 1;
      PopupMenu1.Items[1].Items[i - 1].OnClick := PopupMenuItemsClick;
    end;
  PopupMenu1.Items[1].Items[0].Caption := 'Реального времени';
  PopupMenu1.Items[1].Items[1].Caption := 'Высокий';
  PopupMenu1.Items[1].Items[2].Caption := 'Выше среднего';
  PopupMenu1.Items[1].Items[3].Caption := 'Средний';
  PopupMenu1.Items[1].Items[4].Caption := 'Ниже среднего';
  PopupMenu1.Items[1].Items[5].Caption := 'Низкий';
  /// ////////////////////////////////////////////////////////////////////////
  for i := 0 to 2 do
  /// // создание подпунктов popupmenu
    begin
      ItemsNew1[i] := TMenuItem.Create(PopupMenu2);
      PopupMenu2.Items[2].Add(ItemsNew1[i]);
      PopupMenu2.Items[2].Items[i].Caption := 'Пункт' + inttostr(i);
      PopupMenu2.Items[2].Items[i].Tag := i;
      PopupMenu2.Items[2].Items[i].OnClick := PopupMenu2ItemsClick;
    end;
  PopupMenu2.Items[2].Items[0].Caption := 'Автоматически';
  PopupMenu2.Items[2].Items[1].Caption := 'Вручную';
  PopupMenu2.Items[2].Items[2].Caption := 'Отключена';
  Memo1.Clear;
  SpeedButton3.left := SpeedButton2.left;
  SpeedButton3.top := SpeedButton2.top;
  SpeedButton3.height := SpeedButton2.height;
  SpeedButton3.width := SpeedButton2.width;
  SpeedButton3.Visible := false;

  SpeedButton4.left := SpeedButton2.left;
  SpeedButton4.top := SpeedButton2.top;
  SpeedButton4.height := SpeedButton2.height;
  SpeedButton4.width := SpeedButton2.width;
  SpeedButton4.Visible := false;

  SpeedButton5.left := SpeedButton2.left;
  SpeedButton5.top := SpeedButton2.top;
  SpeedButton5.height := SpeedButton2.height;
  SpeedButton5.width := SpeedButton2.width;
  SpeedButton5.Visible := false;

  SpeedButton1.left := SpeedButton2.left;
  SpeedButton1.top := SpeedButton2.top;
  SpeedButton1.height := SpeedButton2.height;
  SpeedButton1.width := SpeedButton2.width;
  SpeedButton1.Visible := false;
  AccessSettingLevel := false;
  Application.Title := 'Management Remote PC 4.1';
  captionMainForms := Application.Title + ' Free (skrblog.ru)';
  captionMainFormsPro:=Application.Title+' Pro (skrblog.ru)';
  InventConf := false; // признак того что инвентариация не запущена
  InventMicrosoft := false; // признак того что инвентариация не запущена
  InventSoft := false; // признак того что инвентариация не запущена
  SolveExitInvMicrosoft := true;
  // признак того что поток инвентаризации Windows и  Office не запущен
  SolveExitInvSoft := true;
  // признак того что поток инвентаризации программ не запущен
  SolveExitInvConf := true;
  // признак того что поток инвентаризации оборудования не запущен
  /// ////////////////////////////////
  OutForPing:=false; //// сканирование остановлено
  SolveExitInvScan:=true; //// признак того то поток сканирования не запущен
  /// признак того что поток GeneralPingPC не запущен?, значит его можно запустить


  /// /startscanT,startscanR
  if not Assigned(startscanR) then    /// таймер для генерации активациии
    begin
      startscanR := TTimer.Create(frmDomainInfo);
      startscanR.Name := 'startscanR';
      startscanR.Interval := 20000;
      startscanR.OnTimer := ScanR;
      startscanR.Enabled := true;
    end;

  if (StartScan = true) and (not Assigned(startscanT)) then
    begin
      startscanT := TTimer.Create(frmDomainInfo);
      startscanT.Name := 'startscanT';
      startscanT.Interval := 10000;
      startscanT.OnTimer := ScanT;
      startscanT.Enabled := true;
    end;


  try
  EventUserLogin:=true; // признак изменения сеанса пользователя
  FRegisteredSessionNotification := RegisterSessionNotification(Application.Handle, NOTIFY_FOR_THIS_SESSION); // регистрируемся для получения системных сообщений о сотоянии сеанса пользователя
  if FRegisteredSessionNotification then Application.OnMessage := AppMessage;  // если зарегались то присваиваем небходимую процедуру
  except on E: Exception do
  Log_write(2,'error','Ошибка регистрации приема системных сообщений '+e.Message);
  end;

end;

procedure TfrmDomainInfo.FormDestroy(Sender: TObject);
begin
  ListSoftID.Free; // удаляем список с программамит
  CloseHandle(HM); // убиваем мьютекс для разрешения копий
  if FRegisteredSessionNotification then UnRegisterSessionNotification(Handle); // если мы зарегались для получения сообщений то надо разрегаться
end;


procedure TfrmDomainInfo.AppMessage(var Msg: TMSG; var Handled: Boolean); // собственно эта процедура и получает уведомление о состоянии сеанса пользователя
var
  IniSettings:TMeminiFile;
  StopLockUser:boolean;
begin
  Handled := False;
  if Msg.Message = WM_WTSSESSION_CHANGE then
  begin
   IniSettings:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
   try
   StopLockUser:=IniSettings.Readbool('Scan','StopLockUser',false); // остановка сканирования при блокировке пользователя
   finally
   IniSettings.Free;
   end;
     case Msg.wParam of
       WTS_CONSOLE_CONNECT: //Сеанс, обозначенный lParam, был подключен к консольному терминалу или сеансу RemoteFX.
         begin
         Memo1.Lines.Add( datetimetostr(now)+'  - Консольный сеанс - WTS_CONSOLE_CONNECT (' + IntToStr(msg.Lparam)+')');
         EventUserLogin:=true; // пользователь в системе
         end;
       WTS_CONSOLE_DISCONNECT:
         begin //Сеанс, обозначенный lParam, был отключен от консольного терминала или сеанса RemoteFX.
         EventUserLogin:=false; //пользователь отключился, признак изменения саенса пользователя для обновления listview или только вненсения в базу
         Memo1.Lines.Add( datetimetostr(now)+'  - Завершение консольного сеанса - WTS_CONSOLE_DISCONNECT (' + IntToStr(msg.Lparam)+')');
         if StopLockUser and OutForPing then // если необходимо останавливать сканирование и если запущено сканирование то  останавливаем
         SpeedButton40.Click;
         end;
       WTS_REMOTE_CONNECT: //Сеанс, обозначенный lParam, был подключен к удаленному терминалу.
           begin
           Memo1.Lines.Add( datetimetostr(now)+'  - Подключение удаленного сеанса - WTS_REMOTE_CONNECT (' + IntToStr(msg.Lparam)+')');
           EventUserLogin:=true; // пользователь в системе
           end;
       WTS_REMOTE_DISCONNECT:
          begin  //Сеанс, указанный lParam, был отключен от удаленного терминала.
           EventUserLogin:=false; //пользователь отключился, признак изменения саенса пользователя для обновления listview или только вненсения в базу
           Memo1.Lines.Add( datetimetostr(now)+'  - Отключение удаленного сеанса - WTS_REMOTE_DISCONNECT (' + IntToStr(msg.Lparam)+')');
           if StopLockUser and OutForPing then // если необходимо останавливать сканирование и если запущено сканирование то  останавливаем
           SpeedButton40.Click;
          end;
       WTS_SESSION_LOGON:  //Пользователь вошел в сеанс, указанный lParam .
           begin
           Memo1.Lines.Add( datetimetostr(now)+'  -  Пользователь вошел в систему - WTS_SESSION_LOGON (' + IntToStr(msg.Lparam)+')');
           EventUserLogin:=true; //пользователь в системе, признак изменения саенса пользователя
           end;
       WTS_SESSION_LOGOFF:
          begin //Пользователь вышел из сеанса, указанного lParam .
           EventUserLogin:=false; //пользователь отключился, признак изменения саенса пользователя для обновления listview или только вненсения в базу
           Memo1.Lines.Add( datetimetostr(now)+'  -  Завершение сеанса пользователя - WTS_SESSION_LOGOFF (' + IntToStr(msg.Lparam)+')');
           if StopLockUser and OutForPing then // если необходимо останавливать сканирование и если запущено сканирование то  останавливаем
           SpeedButton40.Click
          end;
       WTS_SESSION_LOCK:
           begin //Сеанс, указанный lParam , заблокирован.
           EventUserLogin:=false; //пользователь отключился, признак изменения саенса пользователя для обновления listview или только вненсения в базу
           Memo1.Lines.Add( datetimetostr(now)+'  -  Пользователь заблокировал сеанс - WTS_SESSION_LOCK (' + IntToStr(msg.Lparam)+')');
           if StopLockUser and OutForPing then // если необходимо останавливать сканирование и если запущено сканирование то  останавливаем
           SpeedButton40.Click
           end;
       WTS_SESSION_UNLOCK: //Сеанс, указанный lParam , разблокирован.
          begin
          Memo1.Lines.Add( datetimetostr(now)+'  -  Пользователь разблокировал сеанс - WTS_SESSION_UNLOCK (' + IntToStr(msg.Lparam)+')');
          EventUserLogin:=true; //пользователь в системе, признак изменения саенса пользователя
          end;
       WTS_SESSION_REMOTE_CONTROL:
           begin
           //Сеанс, идентифицированный lParam , изменил свой статус удаленного управления.
           // Чтобы определить статус, вызовите GetSystemMetrics и проверьте метрику SM_REMOTECONTROL .
            if  GetSystemMetrics(SM_REMOTECONTROL)<>0 then
            begin
            EventUserLogin:=true; // признак изменения саенса пользователя
            Memo1.Lines.Add( datetimetostr(now)+'  -  Удаленный терминальный сеанс - WTS_SESSION_REMOTE_CONTROL (' + IntToStr(msg.Lparam)+')');
             {Эта системная метрика используется в среде служб терминалов,
             чтобы определить, управляется ли текущий сеанс сервера терминалов
              удаленно. Его значение отлично от нуля то текущий сеанс
               управляется удаленно; в противном случае 0.}
            end;
            end;
        End; /// case
  end;
end;

procedure TfrmDomainInfo.FormShow(Sender: TObject);
///
var
  setini: TMemInifile;
  i: Integer;
begin
  try
    MainPage.TabIndex := 0;
    MainPage.TabIndex := 1;
    MainPage.TabIndex := 2;
    MainPage.TabIndex := 0;
    setini := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
      '\Settings.ini', TEncoding.Unicode);
    try
      if setini.readString('ConfLAN', 'type', 'AD') = 'FirstStart' then
        FormOneStart.Show
      else if setini.ReadBool('ConfLAN', 'Load', false) then
        FormOneStart.Show;

      { else
        BEGIN
        SelectGroupAD:=SetIni.readstring('ConfLAN','Group','');
        UserGroup:=setini.ReadBool('ConfLAN','UserGroup',true);
        PCGroup:=setini.ReadBool('ConfLAN','PCGroup',true);
        //////////////////////////////////////////////////
        //Просто вызываем все функции подряд (не делал проверок на результат функций)
        //ledCompName.Text := GetCurrentComputerName;
        //ledDomainName.Text := CurrentDomainName;
        //MyDCName.Text := GetDomainController(CurrentDomainName);
        //Единственно, если нет контролера домена, то дальше определять бесполезно
        // если последний раз открывали список компьютеров из AD то открываем его
        if (GetDomainController(CurrentDomainName) <> '') and (confLAN='AD') then
        begin
        //ledDNSName.Text := GetDNSDomainName(CurrentDomainName);
        // EnumAllTrustedDomains; // Все доверенные домены, нах не надо
        GetCurrentComputerName;  //информация о компьютере  в результате должно быть имя, но главное получваем в переменную CurrentDomainName имя домена
        LabeledEdit3.Text := CurrentDomainName; // имя домена в LabelEdEdit на главную форму
        EnumAllGroups;  // получаем список всех групп домена
        if SelectGroupAD='' then SelectGroupAD:='Компьютеры домена'; //если группа безопасности в ini файле не указана то назначаем ему группу по уморлчанию
        if UserGroup then GetAllGroupUsersInComboBox(SelectGroupAD) // если выгружать пользователей только выбранной группы
        else    EnumAllUsers; //// иначе получаем список всех пользователей домена

        if PCGroup then // если выбрано загружать только компьютеры выбранной группы безопасности
        begin
        GetPCForGroupToLisViewComboBox(SelectGroupAD);
        combobox1.Text:=SelectGroupAD;
        end
        else  // иначе выгружаем все ПК в Combobox на вкладке Управление и список ПК выбранной группы в ListView  на вкладке Компьютеры
        if ComboBox1.Items.Count<>0 then
        BEGIN
        EnumAllWorkStation; // передаем все рабочие станции в combobox в разделе Компьютеры
        for I := 0 to ComboBox1.Items.Count do //выгружаем список компьютеров выбранной группы безопасности
        Begin
        if ComboBox1.Items[i]=SelectGroupAD then
        begin
        combobox1.Text:=SelectGroupAD;
        Combobox1.OnSelect(Combobox1); /// onSelect -процедура, загружает список компов
        Break;
        end;
        if ComboBox1.Items[i]='Domain Computers' then
        begin
        combobox1.Text:='Domain Computers';
        Combobox1.OnSelect(Combobox1);
        Break;
        end;
        End;
        END;
        end;
        // если последний раз открывали список ip адресов
        if confLAN='IP' then
        begin
        if not loadlistpc(extractfilepath(application.ExeName)+'\listIP.txt') then
        ShowMessage('Что то пошло не так при загрузке'
        +#10#13+' предыдущего списка компьютеров!');
        end;
        // если последний раз открывали список компьютеров из файла
        if confLAN='file' then
        begin
        if not loadlistpc(extractfilepath(application.ExeName)+'\listPC.txt') then
        ShowMessage('Что то пошло не так при загрузке'
        +#10#13+' предыдущего списка компьютеров!');
        end;
        END; }
    finally
      setini.Free;
    end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка загрузки списков компьютеров ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.ScanT(Sender: TObject);
begin
  if ListView8.Items.Count > 1 then
  /// / если список  компов не пуст
    begin
      if SpeedButton39.Visible = true then
        begin
          SpeedButton39.Click;
        end;
      startscanT.Enabled := false;
      FreeAndNil(startscanT);
    end;
end;

procedure TfrmDomainInfo.ScanR(Sender: TObject);
begin
  TrLic := unit9.Lcheck.Create(true);
  TrLic.FreeOnTerminate := true;
  TrLic.Start;
  startscanR.Enabled := false;
  FreeAndNil(startscanR);
end;

procedure TfrmDomainInfo.ListView1CustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  if cdsSelected in State then
    Sender.Canvas.Font.Color := clRed;
end;

procedure TfrmDomainInfo.ListView1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListView2ColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  ListView2.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView2MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListView5ColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  sort := -sort;
  SortLV8 := not SortLV8;
  ListView5.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView5CustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  try
    if Item.SubItems[4] = 'True' then
      ListView5.Canvas.Font.Color := clBlue;
    if (Item.SubItems[4] = 'True') and (Item.SubItems[5] = 'False') then
      ListView5.Canvas.Font.Color := clFuchsia;
    if Item.SubItems[6] <> 'Ok' then
      ListView5.Canvas.Font.Color := clRed;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка : ' + E.Message);
  end;

end;

procedure TfrmDomainInfo.ListView6ColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  ListView6.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView6CustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  if Item.SubItems[2] = 'Error' then
    ListView6.Canvas.Font.Color := clRed;
  if Item.SubItems[2] = 'Unknown' then
    ListView6.Canvas.Font.Color := clRed;
  if Item.SubItems[2] = 'Pred Fail' then
    ListView6.Canvas.Font.Color := clRed;
  if Item.SubItems[2] = 'Stressed' then
    ListView6.Canvas.Font.Color := clRed;
end;

procedure TfrmDomainInfo.ListView6MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListView7ColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  ListView7.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView7CustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  if Item.SubItems[3] = 'Подключено' then
    ListView7.Canvas.Font.Color := clBlue;
  if Item.SubItems[3] = 'Отключено' then
    ListView7.Canvas.Font.Color := clRed;
end;

procedure TfrmDomainInfo.ListView8Change(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  StatusBar1.Panels[0].text := 'Всего ПК: ' + inttostr(ListView8.Items.Count);
  // StatusBar1.Panels[1].Text:='ПК в сети : '+inttostr(PcInLan);

end;

procedure TfrmDomainInfo.ListView8Click(Sender: TObject);
begin
  if JvNetscapeSplitter1.Maximized = false then
  /// / если панель свойст раскрыта то читаем инфу
    if (ListView8.Items.Count <> 0) and (ListView8.SelCount = 1) then
      begin
        GroupBox2.Caption := ListView8.Selected.SubItems[0];
        // передаем имя компа в caption для других функций
        ListViewSoftinPC.Clear; // очистка списка
        ListViewMicLic.Clear; // очистка списка
        ListViewAV.Clear; // очистка списка
        try
          // читаем список оборудования
          Datam.FDQueryRead2.SQL.Clear;
          /// чтение конфигурации первого вхождения и передача в treeview
          Datam.FDQueryRead2.SQL.text :=
            'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
            (ListView8.Selected.SubItems[0]) + '''' + ' ORDER BY DATE_INV DESC';
          Datam.FDQueryRead2.Open;
          createtreeView(Datam.FDQueryRead2, TreeView4);
          // Функция постороения дерева
          Datam.FDQueryRead2.SQL.Clear; // очистить
          Datam.FDQueryRead2.Close;
          /// закрыть нах после чтения
          // читаем список установленных программ
          readSoftRorSelectPC(ListView8.Selected.SubItems[0], ListViewSoftinPC);
          // читаем список лицензий
          readMicrosoftLic(ListView8.Selected.SubItems[0], ListViewMicLic);
          // читаем список антивирусных продуктов
          readAntivirusStatus(ListView8.Selected.SubItems[0], ListViewAV);
        except
          on E: Exception do
            Datam.FDQueryRead2.Close;
          // Memo1.Lines.Add('Ошибка : '+e.Message);
        end;
      end;
end;

procedure TfrmDomainInfo.ListView8ColumnClick(Sender: TObject;
  Column: TListColumn);
var
  i, inLan: Integer; // PCCheck

begin
  if Column = ListView8.Columns[0] then
  /// если жмякнули на 0ю колонку то проводим выделение или снятие
    begin
      if (ListView8.Columns[0].ImageIndex = 1) and (PCCheck <> 0) then
      /// снимаем все чекбоксы в списке компьютеров
        begin
          for i := 0 to ListView8.Items.Count - 1 do
            if ListView8.Items[i].Checked then
              ListView8.Items[i].Checked := false;

          ListView8.Columns[0].ImageIndex := 1;
          PCCheck := 0;
          StatusBar1.Panels[2].text := 'Выбрано ПК: ' + inttostr(PCCheck);
          exit;
        end;
      /// ////////////////////////////////////////////////////////////////
      if ListView8.Columns[0].ImageIndex = 1 then
        begin
          inLan := MessageDlg('Выделить доступные ПК или все?', mtConfirmation,
            [mbYes, mbAll, mbCancel], 0);
          if inLan = mrCancel then
            exit; // если отмена то выход
          PCCheck := 0;
          if inLan = mrAll then
          /// выделить все компы
            for i := 0 to ListView8.Items.Count - 1 do
              begin
                if ListView8.Items[i].ImageIndex <> 12 then
                /// если компьютер не исключен из списка сканирования
                  begin
                    ListView8.Items[i].Checked := true;
                    PCCheck := PCCheck + 1;
                  end;
              end;
          if inLan = mrYes then
          /// выделить только доступные ПК
            for i := 0 to ListView8.Items.Count - 1 do
              if ListView8.Items[i].ImageIndex = 4 then
                begin
                  ListView8.Items[i].Checked := true;
                  PCCheck := PCCheck + 1;
                end;
          ListView8.Columns[0].ImageIndex := 2;
          StatusBar1.Panels[2].text := 'Выбрано ПК: ' + inttostr(PCCheck);
          exit;
        end;
      if ListView8.Columns[0].ImageIndex = 2 then
        begin
          for i := 0 to ListView8.Items.Count - 1 do
            begin
              if ListView8.Items[i].Checked then
                begin
                  PCCheck := PCCheck - 1;
                  ListView8.Items[i].Checked := false;
                end;
              ListView8.Columns[0].ImageIndex := 1;
            end;
          StatusBar1.Panels[2].text := 'Выбрано ПК: ' + inttostr(PCCheck);
          exit;
        end;
    end
  else
  /// нжали на другю колонку и проводим сортировку
    begin
      SortLV8 := not SortLV8;
      ListView8.CustomSort(@CustomSortProc, Column.Index);
    end;
end;

procedure TfrmDomainInfo.ListView8CustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin

  begin
    // if (item.Selected)then item.ImageIndex:=0
    // else item.ImageIndex:=-1;

  end;
  // item.SubItemImages[0]:=0;
  // subitemImages[0]:=1;
end;

procedure TfrmDomainInfo.ListView8DblClick(Sender: TObject);
begin
  // if (ListView8.Items.Count<>0)and(ListView8.ItemFocused.Checked) then
  if (ListView8.Items.Count <> 0) and
    (ListView8.Items[ListView8.ItemIndex].Focused) then
    begin
      // ListView8.Items[ListView8.ItemIndex].Checked:=true;
      ComboBox2.text := ListView8.Items[ListView8.ItemIndex].SubItems[0];
      LabeledEdit7.text := ListView8.Items[ListView8.ItemIndex].SubItems[2];
      MyPS := ComboBox2.text;
      // MainPage.TabIndex:=1; // переход на страницу Управление
      // if listview8.Items[ListView8.ItemIndex].Selected then
      // ListView8.Items[ListView8.ItemIndex].ImageIndex:=0;
    end;

end;

procedure TfrmDomainInfo.ListView8KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  StringFind: String;
begin
  if Key = 116 then
  /// / если нажали F5 то обновляем
    begin
      readinfoforpcDB;
    end;
  /// ////////////////////////////////////////////////////////////////////////////////
  if Key = 13 then
  /// если жмякнули Enter то это тоже самое что и клик
    begin
      try
        if JvNetscapeSplitter1.Maximized = false then
        /// / если панель свойст раскрыта то читаем инфу
          if (ListView8.Items.Count <> 0) and (ListView8.SelCount = 1) then
            begin

              // читаем оборудование
              Datam.FDQueryRead2.SQL.Clear;
              /// чтение конфигурации первого вхождения и передача в treeview
              Datam.FDQueryRead2.SQL.text :=
                'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
                (ListView8.Items[ListView8.ItemFocused.Index].SubItems[0]) +
                '''' + ' ORDER BY DATE_INV DESC';
              Datam.FDQueryRead2.Open;
              createtreeView(Datam.FDQueryRead2, TreeView4);
              // Функция постороения дерева
              Datam.FDQueryRead2.SQL.Clear; // очистить
              Datam.FDQueryRead2.Close;
              /// закрыть нах после чтения
              GroupBox2.Caption := ListView8.Items[ListView8.ItemFocused.Index]
                .SubItems[0];
              /// ////////////////////////
              // читаем список установленных программ
              readSoftRorSelectPC(ListView8.Items[ListView8.ItemFocused.Index]
                .SubItems[0], ListViewSoftinPC);
            end;
      except
        on E: Exception do
          Memo1.Lines.Add('Ошибка : ' + E.Message);
      end;
    end;
  /// //////////////////////////////////////////////////////////
  if (ssCtrl in Shift) AND (Key = 70) then // если Ctrl+F поиск
    if (ListView8.Items.Count <> 0) and (ListView8.ItemIndex <> 0) then
      begin
        findForListView(ListView8);
      end;
end;

procedure TfrmDomainInfo.ListView8MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
// const                             // Считаем выделенные чекбоксы
// States: array [Boolean] of string = ('не отмечен', 'отмечен');
var
  Item: TListItem;
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
  if Button = mbLeft then
    with Sender as TListView do
      begin
        Item := GetItemAt(X, Y);
        if Assigned(Item) then
          if htOnStateIcon in GetHitTestInfoAt(X, Y) then
            begin
              if Item.Checked then
                PCCheck := PCCheck + 1
              else
                PCCheck := PCCheck - 1;
              StatusBar1.Panels[2].text := 'Выбрано ПК: ' + inttostr(PCCheck);
              // memo1.Lines.Add (Format('Индекс элемента: %d, состояние: %s', [Item.Index, States[Item.Checked]]));
            end;

      end
end;

procedure TfrmDomainInfo.ListView8Resize(Sender: TObject);
begin
  // if TabSheet9.FindComponent('MyRDPClient')<>nil then
  // MyRDPClient.Reconnect(tabSheet9.Width,tabSheet9.Height);
end;

procedure TfrmDomainInfo.ListView8SelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  if ListView8.SelCount > 1 then
    begin
      if (Item.Selected) and (Item.Checked <> true) then
        begin
          Item.Checked := true;
          PCCheck := PCCheck + 1
        end;
      StatusBar1.Panels[2].text := 'Выбрано ПК: ' + inttostr(PCCheck);

    end;
end;

procedure TfrmDomainInfo.lvWorkStationCustomDrawSubItem(Sender: TCustomListView;
  Item: TListItem; SubItem: Integer; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  if (SubItem = 4) and (Item.SubItems[3] <> '0') then
    begin
      lvWorkStation.Canvas.Font.Color := clRed;
      lvWorkStation.Canvas.Font.Style := [fsBold];
      // sender.Canvas.Font.Color:=clred;
    end;
  if (SubItem = 6) then
    begin
      lvWorkStation.Canvas.Font.Color := clBlack;
    end;

end;

procedure TfrmDomainInfo.lvWorkStationMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.lvWorkStationSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
begin
  Item.Checked := true;
  if Selected = false then
    Item.Checked := false;

end;

procedure TfrmDomainInfo.MainPageChange(Sender: TObject);
begin

  if MainPage.ActivePageIndex = 1 then
  /// если активна вкладка управление переносим flowpanel на нее
    begin
      GroupRDPWin.Parent := Panel7;
    end;
  if MainPage.ActivePageIndex = 0 then
  // если активна вкладка компьютеры переносим flowpanel на нее
    begin
      GroupRDPWin.Parent := Panel8;
    end;
end;

procedure TfrmDomainInfo.Microsoft1Click(Sender: TObject);
begin // активация windows
  ActivationWindowsPC;
end;

procedure TfrmDomainInfo.Mount1Click(Sender: TObject);
var
  NewMountHDD: TThread;
begin
  if (ListViewDisk.Items.Count <> 0) and (ListViewDisk.SelCount <> 0) then
    begin
      if ListViewDisk.Items[ListViewDisk.ItemIndex].SubItems[5] <> '' then
        begin
          SelectedHDD := delstr(ListViewDisk.Items[ListViewDisk.ItemIndex]
            .SubItems[5]);
          MyPS := ComboBox2.text;
          MyUser := LabeledEdit1.text;
          MyPasswd := LabeledEdit2.text;
          NewMountHDD := HDDMount.MountHdd.Create(true);
          NewMountHDD.FreeOnTerminate := true;
          NewMountHDD.Start;
        end;
    end;
end;

procedure TfrmDomainInfo.msi3Click(Sender: TObject);
/// групповая установка программ
begin
  InstallMSIForListPC;
end;

procedure TfrmDomainInfo.msi4Click(Sender: TObject);
/// /// групповое удаление программы (msi)
begin
  DeleteProgramMSIForListPC;
end;

procedure TfrmDomainInfo.N11Click(Sender: TObject);
begin
/// /  завершение сеанса
  CloseSession;
end;

procedure TfrmDomainInfo.N120Click(Sender: TObject);
begin // запустить задачу
  Button6.Click;
end;

procedure TfrmDomainInfo.N121Click(Sender: TObject);
begin // остановить задачу
  Button7.Click;
end;

procedure TfrmDomainInfo.N122Click(Sender: TObject);
begin // удалить задачу
  Button3.Click;
end;

procedure TfrmDomainInfo.N124Click(Sender: TObject); // создание новой задачи
begin
  Button2.Click;
end;

procedure TfrmDomainInfo.N125Click(Sender: TObject);
// обновить список задач popupmenu
begin
  Button4.Click;
end;

procedure TfrmDomainInfo.N123Click(Sender: TObject); // свойство залдачи в popup
begin
  Button5.Click;
end;

procedure TfrmDomainInfo.N12Click(Sender: TObject);
begin
/// / принудительное завершение сеанса
  ForceCloseSession;
end;

procedure TfrmDomainInfo.N13Click(Sender: TObject);
begin
/// перезагрузка
  ResetCurPC;
end;

function TfrmDomainInfo.win10Befo: Integer;
var
  ver: Integer;
begin
  if TryStrToInt(copy(OSVersion, 1, 2), ver) then
    Result := ver
  else if pos('10', OSVersion) = 1 then
    Result := 10
  else
    begin
      TryStrToInt(copy(OSVersion, 1, 1), ver);
      Result := ver;
    end;
end;

procedure TfrmDomainInfo.N140Click(Sender: TObject); // отключить PnP Device
var
  reb: Boolean;
  Res: Integer;
begin
  if not ping(ComboBox2.text) then
    exit;
  if win10Befo < 10 then
    begin
      showmessage('Операция доступна для версий не ниже Windows 10 ');
      exit;
    end;
  if ListView6.SelCount = 1 then
    if ListView6.Selected.SubItems[3] <> '' then
      begin
        Res := EnableDisableDevicePnP(ComboBox2.text, LabeledEdit1.text,
          LabeledEdit2.text, ListView6.Selected.SubItems[3], 0, reb);
        Memo1.Lines.Add('Выключение ' + ListView6.Selected.SubItems[3] + ' : ' +
          SysErrorMessage(Res));
        if reb then
          Memo1.Lines.Add
            ('Чтобы изменения вступили в силу необходима перезагрузка компьютера');
        if Res = 0 then
          begin
            Memo1.Lines.Add('Загрузка списка устройств PnP');
            LoadDevicePnP(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text,
              ListView6);
            Memo1.Lines.Add('Загрузка списка устройств PnP завершена.');
          end;
      end;
end;

procedure TfrmDomainInfo.N141Click(Sender: TObject); // Включить PnP Device
var
  reb: Boolean;
  Res, ver: Integer;
begin
  if not ping(ComboBox2.text) then
    exit;
  if win10Befo < 10 then
    begin
      showmessage('Операция доступна для версий не ниже Windows 10 ');
      exit;
    end;
  if ListView6.SelCount = 1 then
    if ListView6.Selected.SubItems[3] <> '' then
      begin
        Res := EnableDisableDevicePnP(ComboBox2.text, LabeledEdit1.text,
          LabeledEdit2.text, ListView6.Selected.SubItems[3], 1, reb);
        Memo1.Lines.Add('Включение ' + ListView6.Selected.SubItems[3] + ' : ' +
          SysErrorMessage(Res));
        if reb then
          Memo1.Lines.Add
            ('Чтобы изменения вступили в силу необходима перезагрузка компьютера');
        if Res = 0 then
          begin
            Memo1.Lines.Add('Загрузка списка устройств PnP');
            LoadDevicePnP(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text,
              ListView6);
            Memo1.Lines.Add('Загрузка списка устройств PnP завершена.');
          end;

      end;

end;

procedure TfrmDomainInfo.N142Click(Sender: TObject);
begin
  InventoryMicrosoft;
  // процедура запуска инвентаризации продуктов Windows и Office
end;

procedure TfrmDomainInfo.N143Click(Sender: TObject);
begin
  Button27.Click;
end;

procedure TfrmDomainInfo.N144Click(Sender: TObject);
begin
  if LVMicrosoft.Items.Count <> 0 then
    popupListViewSaveAs(LVMicrosoft,
      'Сохранение списка продуктов Windows и Office',
      'Лицензии Windows и Office');

end;

procedure TfrmDomainInfo.N145Click(Sender: TObject);
begin
  LVMicrosoftDblClick(LVMicrosoft);
end;

procedure TfrmDomainInfo.CopyClipBoardClick(Sender: TObject);
var
  s: string;
  Caller: TObject;
  i, z: Integer;
  function CopyToClip(s: string): Boolean;
  begin
    Clipboard.aStext := s;
  end;

begin
  s := '';
  Caller := ((Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent;
  if (Caller is TListView) then
    Begin
      if (Caller as TListView).SelCount = 0 then
        exit;
      for z := 0 to (Caller as TListView).Items.Count - 1 do
        if (Caller as TListView).Items[z].Selected then
          begin
            if (Caller as TListView).Items[z].Caption <> '' then
              s := s + (Caller as TListView).Columns[0].Caption + ' - ' +
                (Caller as TListView).Items[z].Caption + ' * ';
            for i := 0 to (Caller as TListView).Columns.Count - 2 do
              begin
                s := s + (Caller as TListView).Columns[i + 1].Caption + ' - ' +
                  (Caller as TListView).Items[z].SubItems[i] + ' * ';
              end;
            s := s + #10;
          end;
      CopyToClip(s);
    End;

end;

procedure TfrmDomainInfo.rebuutSelectPC;
var
  MyShutdownPC: TThread;
begin
  MyPS := ComboBox2.text;
  if ping(ComboBox2.text) then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 3;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.N15Click(Sender: TObject);
begin
/// завершение работы
  PowerOff;
end;

procedure TfrmDomainInfo.N16Click(Sender: TObject);
begin
/// принудительное завершение работы
  ForcePowerOff;
end;

procedure TfrmDomainInfo.N17Click(Sender: TObject);
var
  MyShutdownPC: TThread;
begin
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 6;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;

end;

/// //////////////////////////////////////////////////////// реестр

procedure TfrmDomainInfo.N18Click(Sender: TObject);
/// // добавление в автозагрузку
begin
  // if (listview2.ItemIndex=-1) then exit;
  // if (ListView2.Items.Count<>0) and (listview2.Selected.SubItems[2]<>'Startup')
  // then
  begin
    OKRightDlg2.showmodal;
    MyPS := ComboBox2.text;
  end;
end;

function hk(s: string): Integer;
begin
  Delete(s, 5, Length(s) - 4);
  if s = 'HKCR' then
    Result := 1;
  if s = 'HKCU' then
    Result := 2;
  if s = 'HKLM' then
    Result := 3;
  if s = 'HKU\' then
    Result := 4;
  if s = 'HKCC' then
    Result := 5;
end;

function hkpath(s: string): string;
var
  i: byte;
begin
  i := pos('\', s);
  Delete(s, 1, i);
  Result := s;
end;

procedure TfrmDomainInfo.N19Click(Sender: TObject);
/// // удалить запись из автозагрузки
var
  FSWbemLocator: OleVariant;
  FWMIService: OleVariant;
  FWbemObjectSet: OleVariant;
  hDefKey: Integer;
begin
  if ListView2.ItemIndex = -1 then
    exit;

  if pos('Startup', ListView2.Selected.SubItems[2]) = 0 then
    Begin
      try
        OleInitialize(nil);
        MyPS := ComboBox2.text;
        FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
        FWMIService := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser,
          MyPasswd, '', '', 128);
        FWMIService.Security_.impersonationlevel := 3;
        FWMIService.Security_.authenticationLevel := 6;
        FWMIService.Security_.Privileges.AddAsString('SeBackupPrivilege', true);
        // Требуется для резервного копирования файлов и каталогов, независимо от списка ACL, указанного для файла.
        FWMIService.Security_.Privileges.AddAsString('SeRestorePrivilege',
          true); // Требуется для восстановления файлов и каталогов независимо от списка ACL, указанного для файла.
        FWbemObjectSet := FWMIService.get('StdRegProv');
        case hk(ListView2.Selected.SubItems[2]) of
          1:
            hDefKey := HKEY_CLASSES_ROOT;
          2:
            hDefKey := HKEY_CURRENT_USER;
          3:
            hDefKey := HKEY_LOCAL_MACHINE;
          4:
            hDefKey := HKEY_USERS;
          5:
            hDefKey := HKEY_CURRENT_CONFIG;
        end;
        if FWbemObjectSet.DeleteValue(hDefKey,
          hkpath(ListView2.Selected.SubItems[2]),
          ListView2.Selected.SubItems[0]) = 0 then
          ListView2.Selected.Delete;
        VariantClear(FWbemObjectSet);
        VariantClear(FWMIService);
        VariantClear(FSWbemLocator);
        OleUnInitialize;
      except
        on E: Exception do
          begin
            frmDomainInfo.Memo1.Lines.Add('Ошибка удаления автозагрузки "' +
              E.Message + '"');
            OleUnInitialize;
            exit;
          end;
      end;
    End
  else // иначе открываесм в проводнике
    N108.Click;

end;

procedure TfrmDomainInfo.N101Click(Sender: TObject);
begin
/// / очистка списка процессов и остановка таймера сканирования списка процесссов
  if lvWorkStation.Items.Count <> 0 then
    begin
      Timer1.Enabled := not Timer1.Enabled;
      StatusBar1.Panels[3].text := '';
      /// загрузка CPU
      StatusBar1.Panels[4].text := '';
      /// Память
      StatusBar1.Panels[5].text := '';
      /// Количество процессов
    end;
end;

procedure TfrmDomainInfo.N105Click(Sender: TObject);
var
  i: Integer;
begin
  if (ListView8.Items.Count = 0) or (ListView8.SelCount <> 1) then
    exit;

  if ComboBox2.Style = csExDropDown then
    begin
      ComboBox2.text := ListView8.Selected.SubItems[0];
      /// добавляем текст
      ComboBox2.Tag := 1;
      ComboBox2.DroppedDown := true;
      /// раскрываем список компов для того чтобы выбрался компьютер из списка для указания индекса строки
      ComboBox2.DroppedDown := false;
      /// / скрываем список
      ComboBox2.Tag := 0;
      ComboBox2.ItemsEx[ComboBox2.ItemIndex].ImageIndex :=
        ListView8.Selected.ImageIndex;
      /// обновляем картинку в списке combobox в соответствии с индексом
      MainPage.TabIndex := 1;
      /// открываем вкладку управления
      MyPS := ComboBox2.text;
      /// /
    end
  else
    begin
      for i := 0 to ComboBox2.Items.Count - 1 do
        if ListView8.Selected.SubItems[0] = ComboBox2.Items[i] then
          begin
            ComboBox2.ItemIndex := i;
            ComboBox2.ItemsEx[i].ImageIndex := ListView8.Selected.ImageIndex;
            MainPage.TabIndex := 1;
            MyPS := ComboBox2.text;
            break;
          end;
    end;
end;

function DecomposeFilePatch(str: string): string;
// функция обрабатывает передаваемый в нее путь с ключами запуска, выдает путь до файла
var
  Path: string;
  Y, X: Integer;
begin
  try
    Path := str;
    if Ansipos('"', Path) = 1 then
    // еслим первый символ кавычки берем строку между ними
      begin
        Y := Ansipos('"', Path);
        Path := copy(Path, Y + 1, Length(Path));
        X := Ansipos('"', Path);
        Path := copy(Path, 1, X - 1);
      end;
    if Ansipos('"', Path) <> 0 then
      Path := StringReplace(Path, '"', '', [rfReplaceAll]);
    Path := copy(Path, Ansipos(':\', Path) - 1, Length(Path));
    if Ansipos(' -', Path) <> 0 then // ищем ключи запуска через дефис
      begin
        Path := copy(Path, 1, Ansipos(' -', Path) - 1);
      end;

    if Ansipos('/', Path) <> 0 then
      begin
        Path := copy(Path, 1, Ansipos('/', Path) - 1);
        // ищем обратный слеш, признак ключа запуска
      end;
    if ExtractFileDir(Path) <> '' then
      Result := Path
    else
      Result := str;

  except
    on E: Exception do
      begin
        frmDomainInfo.Memo1.Lines.Add('Ошибка формирования строки - "' +
          E.Message + '"');
        Result := str;
      end;
  end;

end;

procedure TfrmDomainInfo.N106Click(Sender: TObject);
// открыть место хранения процесса в проводнике
var
  strpath, FileStr: string;
begin
  if (lvWorkStation.Items.Count = 0) or (lvWorkStation.SelCount = 0) then
    exit;
  if lvWorkStation.SelCount = 1 then
    begin
      strpath := DecomposeFilePatch(lvWorkStation.Selected.SubItems[5]);
      // получаем строку без ключей запуска и все такое
      if not InputQuery('Открыть в проводнике', 'Файл:', strpath) then
        exit;
      FileStr := ExtractFileName(strpath); // имя файла
      strpath := ExtractFileDir(strpath); // каталог расположения файла
      MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
      MRPCExplorer.Show; // показываем окно
      MRPCExplorer.EditPath.text := strpath; // путь до каталога с файлом
      MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
      // переходим в каталог
      MRPCExplorer.findFFinListview(FileStr); // пытаемся найти файл

    end;

end;

procedure TfrmDomainInfo.N107Click(Sender: TObject);
// открыть в проводнике каталог с фалом службы
var
  strpath, FileStr: string;
begin
  if (ListView1.Items.Count = 0) or (ListView1.SelCount = 0) then
    exit;
  if ListView1.SelCount = 1 then
    begin
      if ExtractFileDir(ListView1.Selected.SubItems[5]) <> '' then
        begin
          strpath := DecomposeFilePatch(ListView1.Selected.SubItems[5]);
          if not InputQuery('Открыть в проводнике', 'Файл:', strpath) then
            exit;
          FileStr := ExtractFileName(strpath); // имя файла
          strpath := ExtractFileDir(strpath); // каталог расположения файла
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := strpath; // путь до каталога с файлом
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
          MRPCExplorer.findFFinListview(FileStr); // пытаемся найти файл
        end;
    end;
end;

procedure TfrmDomainInfo.N108Click(Sender: TObject);
// открыть в проводнике автозагрузку
var
  SIDforUser, DirUs: string;
  Res: Integer;
begin
  try
    if (ListView2.Items.Count = 0) or (ListView2.SelCount = 0) then
      exit;
    if ListView2.SelCount = 1 then
      begin

        if (pos('Startup', ListView2.Selected.SubItems[2]) = 0) then
        // если элемент автозагузки из реестра а не из  Startup
          begin // открываем в реестра
            Res := MessageBox(self.Handle, PChar('Открыть значение в реестре?'),
              PChar('Автозагрузка'), MB_YESNO);
            if Res = IDYes then
              begin
                N134.Click;
                exit;
              end;
          end;

        if ListView2.Selected.SubItems[3] = '' then
        // если не указано имя пользователя
          begin
            Res := MessageBox(self.Handle,
              PChar('Не указано имя пользователя.' + #10#13 +
              'Найти профиль текущего активного пользователя?'),
              PChar('Нет имени пользователя'), MB_YESNO);
            if Res = IDNO then
              exit;
          end;
        if ListView2.Selected.SubItems[3] <> '' then
          SIDforUser := GetSID(ListView2.Selected.SubItems[3])
          // узнаем SID пользователя
        else
          SIDforUser := '';
        // иначе сид пустой, поробуем извлеч путь активного пользоателя
        // Memo1.Lines.Add(SIDforUser);  //выводим sid  пользователя
        DirUs := GetUserDirectory(ComboBox2.text, // комп
          LabeledEdit1.text, // пользователь
          LabeledEdit2.text, // пароль
          ListView2.Selected.SubItems[3],
          // имя пользователя (либо имя либо сид, что наедем)
          SIDforUser); // сид пользователя
        Memo1.Lines.Add('Каталог профиля пользователя - ' + DirUs);
        DirUs := DirUs + '\Главное меню\Программы\startup';
        if not InputQuery('Открыть в проводнике', 'Каталог:', DirUs) then
          exit;
        MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
        MRPCExplorer.Show; // показываем окно
        MRPCExplorer.EditPath.text := DirUs; // путь до каталога
        MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
        // переходим в каталог
      end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка открытия автозагрузки в проводнике - ' +
        E.Message);
  end;
end;

procedure TfrmDomainInfo.N111Click(Sender: TObject);
// открыть в проводнике директорию установки MSI программы
var
  Path: string;
begin
  if (ListView3.Items.Count = 0) or (ListView3.SelCount = 0) then
    exit;
  if ListView3.SelCount = 1 then
    begin
      if ListView3.Selected.SubItems[1] <> '' then
        begin
          // path:=DecomposeFilePatch(ListView3.Selected.SubItems[1]);
          Path := ListView3.Selected.SubItems[1];
          if copy(Path, Length(Path), 1) = '\' then
            Path := copy(Path, 1, Length(Path) - 1);
          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N112Click(Sender: TObject);
// открыть в проводнике источник установки MSI программы
var
  Path: string;
begin
  if (ListView3.Items.Count = 0) or (ListView3.SelCount = 0) then
    exit;
  if ListView3.SelCount = 1 then
    begin
      if ListView3.Selected.SubItems[2] <> '' then
        begin
          // path:=DecomposeFilePatch(ListView3.Selected.SubItems[2]);
          Path := ListView3.Selected.SubItems[2];
          if copy(Path, Length(Path), 1) = '\' then
            Path := copy(Path, 1, Length(Path) - 1);
          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N115Click(Sender: TObject);
// открыть в проводнике директорию установки Все программы
var
  Path: string;
begin
  if (ListView4.Items.Count = 0) or (ListView4.SelCount = 0) then
    exit;
  if ListView4.SelCount = 1 then
    begin
      if ListView4.Selected.SubItems[1] <> '' then
        begin
          Path := ListView4.Selected.SubItems[1];
          if copy(Path, Length(Path), 1) = '\' then
            Path := copy(Path, 1, Length(Path) - 1);

          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N116Click(Sender: TObject);
// открыть в проводнике источник установки Все программы
var
  Path: string;
begin
  if (ListView4.Items.Count = 0) or (ListView4.SelCount = 0) then
    exit;
  if ListView4.SelCount = 1 then
    begin
      if ListView4.Selected.SubItems[2] <> '' then
        begin
          Path := ListView4.Selected.SubItems[2];
          if copy(Path, Length(Path), 1) = '\' then
            Path := copy(Path, 1, Length(Path) - 1);

          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N117Click(Sender: TObject);
// открыть в проводнике расположение профиля пользователя
var
  Path: string;
begin
  if (ListView5.Items.Count = 0) or (ListView5.SelCount = 0) then
    exit;
  if ListView5.SelCount = 1 then
    begin
      if ListView5.Selected.SubItems[2] <> '' then
        begin
          // path:=DecomposeFilePatch(ListView5.Selected.SubItems[2]);
          Path := ListView5.Selected.SubItems[2];
          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N118Click(Sender: TObject);
/// открыть в проводнике расшаренный (сетевой ресурс) каталог
var
  Path: string;
begin
  if (ListViewShare.Items.Count = 0) or (ListViewShare.SelCount = 0) then
    exit;
  if ListViewShare.SelCount = 1 then
    begin
      if ListViewShare.Selected.SubItems[1] <> '' then
        begin
          // path:=DecomposeFilePatch(ListViewShare.Selected.SubItems[1]);
          Path := (ListViewShare.Selected.SubItems[1]);
          if not InputQuery('Открыть в проводнике', 'Каталог:', Path) then
            exit;
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := Path; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
        end;
    end;
end;

procedure TfrmDomainInfo.N119Click(Sender: TObject);
/// открыть проводник из контекстного меню списка компьютеров
var
  i: Integer;
begin
  if (ListView8.Items.Count = 0) or (ListView8.SelCount = 0) then
    exit;
  if ListView8.SelCount <> 1 then
    exit;
  if not ping(ListView8.Selected.SubItems[0]) then
    exit;

  MRPCExplorer.ComboBox2.Clear;
  for i := 0 to ComboBox2.Items.Count - 1 do
    begin
      MRPCExplorer.ComboBox2.Items.Add(ComboBox2.Items[i]);
      MRPCExplorer.ComboBox2.ItemsEx[i].ImageIndex := ComboBox2.ItemsEx[i]
        .ImageIndex;
    end;

  MRPCExplorer.ComboBox2.text := ListView8.Selected.SubItems[0];
  /// имя компа
  MRPCExplorer.EditUser.text := LabeledEdit1.text;
  /// пользователь
  MRPCExplorer.EditPass.text := LabeledEdit2.text; // пароль
  if MRPCExplorer.WindowState = wsMinimized then
  // если форма свернута то восстанавливаем ее
    MRPCExplorer.WindowState := wsNormal
  else // иначе показываем
    MRPCExplorer.Show;
end;

procedure TfrmDomainInfo.Uninstall1Click(Sender: TObject);
// открыть в проводнике Uninstall установки Все программы
var
  Path, pathDir, patchFile: string;
begin
  if (ListView4.Items.Count = 0) or (ListView4.SelCount = 0) then
    exit;
  if ListView4.SelCount = 1 then
    begin
      if ListView4.Selected.SubItems[5] <> '' then
        begin
          Path := DecomposeFilePatch(ListView4.Selected.SubItems[5]);
          if not InputQuery('Открыть в проводнике', 'Файл:', Path) then
            exit;
          pathDir := ExtractFileDir(Path);
          patchFile := ExtractFileName(Path);
          MRPCExplorer.ComboBox2.text := ComboBox2.text; // имя компа
          MRPCExplorer.Show; // показываем окно
          MRPCExplorer.EditPath.text := pathDir; // путь до каталога
          MRPCExplorer.SpeedButton2Click(MRPCExplorer.SpeedButton2);
          // переходим в каталог
          MRPCExplorer.findFFinListview(patchFile); // пытаемся найти файл
        end;
    end;
end;

procedure TfrmDomainInfo.N1Click(Sender: TObject);
/// //// завершение процесса
var
  i: Integer;
  allProc: string;
  MyKillProcess: TThread;
begin
  try
    allProc := '';
    if (lvWorkStation.Items.Count <> 0) and (lvWorkStation.SelCount <> 0) then
      begin
        i := MessageBox(self.Handle,
          PChar('Завершить ' + inttostr(lvWorkStation.SelCount) +
          ' процесс(а/ов)?'), PChar('Завершение процессов'),
          MB_YESNO + MB_ICONQUESTION);
        if i = IDNO then
          exit;

        GroupPC := false;
        SelectProc := TstringList.Create;
        NewProcMyPS := ComboBox2.text;
        for i := 0 to lvWorkStation.Items.Count - 1 do
          begin
            if lvWorkStation.Items[i].Selected then
              begin
                allProc := allProc + lvWorkStation.Items[i].SubItems[0] + '=' +
                  lvWorkStation.Items[i].SubItems[2] + ',';
              end;
          end;
        Delete(allProc, Length(allProc), 1);
        /// удаляем последнюю запятую
        SelectProc.CommaText := allProc;
        /// список процессов для завершения
        MyKillProcess := unit3.KillProcess.Create(true);
        MyKillProcess.FreeOnTerminate := true;
        MyKillProcess.Start;
      end;

  except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка завершения процесса - ' + E.Message);
      end;
  end;

end;

procedure TfrmDomainInfo.N20Click(Sender: TObject);
begin
  // if ListView3.Items.Count<>0 then
  begin
    GroupPC := false;
    NewProgramMyPS := ComboBox2.text;
    OKRightDlg1.showmodal;
  end;
end;

procedure TfrmDomainInfo.N21Click(Sender: TObject);
/// // удаление программы
var
  i: Integer;
  MyDelProgram: TThread;
begin
  try
    if (ListView3.Items.Count <> 0) and (ListView3.SelCount <> 0) then
      begin
        MyPS := ComboBox2.text;
        listDelProg := TstringList.Create;
        for i := 0 to ListView3.Items.Count - 1 do
          begin
            if ListView3.Items[i].Selected then
              listDelProg.Add(ListView3.Items[i].SubItems[0])
          end;
        i := MessageDlg('Вы действительно хотите удалить ' +
          inttostr(listDelProg.Count) + ' програм...?', mtConfirmation,
          [mbYes, mbCancel], 0);
        if i = mrYes then
          begin
            MyDelProgram := DeleteProg.DeleteProgram.Create(true);
            MyDelProgram.FreeOnTerminate := true;
            /// / под вопросом самоуничтожение потока
            MyDelProgram.Start;
          end;
        if i = mrCancel then
          begin
            exit;
            listDelProg.Free;
          end;
      end;
  except
    on E: Exception do
      begin
        frmDomainInfo.Memo1.Lines.Add('Ошибка удаления программ - "' +
          E.Message + '"');
      end;
  end;
end;

procedure TfrmDomainInfo.N22Click(Sender: TObject);
var
  s: string;
begin
  if ListView4.Items.Count <> 0 then
    begin
      NewProcMyPS := ComboBox2.text;
      MyPS := ComboBox2.text;
      if pos('MsiExec.exe', ListView4.Selected.SubItems[5]) <> 0 then // /qn
        s := ListView4.Selected.SubItems[5] + ' /quiet'
      else
        s := ListView4.Selected.SubItems[5];
      if pos('MsiExec.exe', s) <> 0 then
        begin // MsiExec.exe /I{517492F5-28F7-4A7C-ACFD-51A86E95610D}
          if pos(Ansiuppercase('MsiExec.exe /I'), Ansiuppercase(s)) <> 0 then
            begin
              s := StringReplace(s, 'MsiExec.exe /I', 'MsiExec.exe /X',
                [rfReplaceAll]);
            end;
        end;
      OKRightDlg12.LabeledEdit1.text := s;
      OKRightDlg12.showmodal;
    end;
end;

procedure TfrmDomainInfo.N23Click(Sender: TObject);
begin
  if (ListViewDisk.Items.Count <> 0) and (ListViewDisk.SelCount <> 0) then
    begin
      if ListViewDisk.Items[ListViewDisk.ItemIndex].SubItems[5] <> '' then
        begin
          SelectedHDD := delstr(ListViewDisk.Items[ListViewDisk.ItemIndex]
            .SubItems[5]);
          MyPS := ComboBox2.text;
          MyUser := LabeledEdit1.text;
          MyPasswd := LabeledEdit2.text;
          OKRightDlg1234.showmodal;
        end;
    end;
end;

procedure TfrmDomainInfo.N25Click(Sender: TObject);
begin
/// добавить в домен или группу
  AddPCDomain;
end;

procedure TfrmDomainInfo.N26Click(Sender: TObject);
begin
/// // переименовать комп
  EditNamePC;
end;

procedure TfrmDomainInfo.N27Click(Sender: TObject);
begin
/// удалить из домена
  OKRightDlg12345678.showmodal;
end;

procedure TfrmDomainInfo.N28Click(Sender: TObject);
/// /добавить сетевой принтер на текущий компьютер
begin
  GroupPC := false;
  MyPS := ComboBox2.text;
  OKRightDlg123456789.showmodal;
end;

procedure TfrmDomainInfo.N29Click(Sender: TObject);
/// / очистить очередь пкечати
var
  SelectPrintCancelAllJob: TThread;
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      SelectPrintCancelAllJob := PrintCancelAllJob.PrintCancelAllJobThread.
        Create(true);
      SelectPrintCancelAllJob.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      SelectPrintCancelAllJob.Start;
    end;
end;

procedure TfrmDomainInfo.N2Click(Sender: TObject);
begin
/// /изменить приоритет процесса
  if (lvWorkStation.Items.Count <> 0) and (lvWorkStation.SelCount <> 0) then
    GroupselectProc := lvWorkStation.Selected.SubItems[2];
end;

procedure TfrmDomainInfo.N30Click(Sender: TObject);
/// /// приостановить печать
var
  NewPrintPause: TThread;
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      NewPrintPause := PrintPausePrint.PrintPause.Create(true);
      NewPrintPause.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      NewPrintPause.Start;
    end;
end;

procedure TfrmDomainInfo.N31Click(Sender: TObject);
/// // печать тестовой страницы
var
  NewPrintTestPage: TThread;
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      NewPrintTestPage := PrintTestPage.PrinterTestPageThread.Create(true);
      NewPrintTestPage.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      NewPrintTestPage.Start;
    end;
end;

procedure TfrmDomainInfo.N32Click(Sender: TObject);
/// переименовать принтер
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      MyPS := ComboBox2.text;
      OKRightDlg1234567891011.showmodal;
    end;
end;

procedure TfrmDomainInfo.N33Click(Sender: TObject);
/// /возобновить печать
var
  NewPrintResume: TThread;
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      MyPS := ComboBox2.text;
      NewPrintResume := PrintResumePrint.ResumePrintThread.Create(true);
      NewPrintResume.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      NewPrintResume.Start;
    end;
end;

procedure TfrmDomainInfo.N34Click(Sender: TObject);
/// / использовать по умолчанию
var
  NewSetdefaultPrinter: TThread;
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      MyPS := ComboBox2.text;
      NewSetdefaultPrinter := PrinterSetDefaultPrinter.DefaultPrinter.
        Create(true);
      NewSetdefaultPrinter.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      NewSetdefaultPrinter.Start;
    end;
end;

procedure TfrmDomainInfo.N35Click(Sender: TObject);
/// установить драйвер на текущий компьютер
begin
  MyPS := ComboBox2.text;
  GroupPC := false;
  OKRightDlg12345678910.showmodal;
  /// / открываем AddPrintDriverDialog
end;

function SIDtoUserName(NewSid: string): string;
/// имя пользователя по SID
var
  FSWbemLocator: OleVariant;
  FWMIService: OleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject: OleVariant;
  username: string;
  oEnum: IEnumvariant;
  iValue: LongWord;
begin
  Result := '';
  try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer('localhost',
      'root\CIMV2', '', '');
    FWbemObjectSet := FWMIService.ExecQuery
      ('SELECT * FROM Win32_UserAccount WHERE SID =' + '"' + NewSid + '"',
      'WQL', wbemFlagForwardOnly);
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    /// / службы
      begin
        Result := (FWbemObject.Name);
        frmDomainInfo.Memo1.Lines.Add(username);
      end;
  except
    on E: Exception do
      begin
        Result := 'Unknown';
        frmDomainInfo.Memo1.Lines.Add('Имя пользователя, ошибка - "' +
          E.Message + '"');
        exit;
      end;
  end;
  FSWbemLocator := Unassigned;
end;

procedure TfrmDomainInfo.N36Click(Sender: TObject);
/// / сменить владельца профиля
var
  i: Integer;
begin
  if (ListView5.Items.Count <> 0) and (ListView5.SelCount <> 0) then
  /// / имя пользователя
    begin
      i := 0;
      if ListView5.Items[ListView5.ItemIndex].SubItems[4] = 'True' then
        i := MessageBox(self.Handle,
          PChar('В текущий момент профиль активен, завершите сеанс или' +
          ' перезагрузите удаленный компьютер.' + #10#13 +
          ' Продолжить  выполнение операции?'),
          PChar(ListView5.Selected.SubItems[1]), MB_YESNO + MB_ICONQUESTION);
      if i = IDNO then
        exit;

      begin
        MyPS := ComboBox2.text;
        UserSID := ListView5.Selected.SubItems[0];
        Memo1.Lines.Add('-------------------------------------');
        Memo1.Lines.Add('Выбран профиль ' + ListView5.Selected.SubItems[1]);
        OKRightDlg123456789101112.showmodal;
      end;
    end;
end;

procedure TfrmDomainInfo.N37Click(Sender: TObject);
begin
  if (ListView5.Items.Count <> 0) and (ListView5.SelCount <> 0) then
  /// / имя пользователя
    begin
      FormUserAccount.SidLab.Caption := '';
      FormUserAccount.SidLab.Caption := ListView5.Items[ListView5.ItemIndex]
        .SubItems[0];
      FormUserAccount.showmodal;
    end;
end;

procedure TfrmDomainInfo.N69Click(Sender: TObject); // свойства принтера
begin
  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      PrinterProperty.Caption := 'Сведения - ' + SelectedPrint;
      PrinterProperty.showmodal;
    end;
end;

procedure TfrmDomainInfo.N38Click(Sender: TObject);
/// / удалить принтер
var
  i: Integer;
  NewDeletePrint: TThread;
begin

  if (ListViewPrint.Items.Count <> 0) and (ListViewPrint.SelCount <> 0) then
    begin
      SelectedPrint := ListViewPrint.Items[ListViewPrint.ItemIndex].Caption;
      i := MessageDlg('Удалить выбранный принтер? ' + SelectedPrint,
        mtConfirmation, [mbYes, mbCancel], 0);
      if i = mrYes then
        begin
          NewDeletePrint := PrintDeleteThread.PrintDeletePrint.Create(true);
          NewDeletePrint.FreeOnTerminate := true;
          /// / под вопросом самоуничтожение потока
          NewDeletePrint.Start;
        end;
      if i = mrCancel then
        exit;
    end;

end;

procedure TfrmDomainInfo.N39Click(Sender: TObject);
begin
  MyPS := ComboBox2.text;
  OKRightDlg12345678910111213.showmodal;
end;

procedure TfrmDomainInfo.N3Click(Sender: TObject);
begin
  GroupPC := false;
  NewProcMyPS := ComboBox2.text;
  NewProcForm.showmodal;
end;

procedure TfrmDomainInfo.N40Click(Sender: TObject);
var
  i: Integer;
  NewUserProfileDelete: TThread;
begin
  if (ListView5.Items.Count <> 0) and (ListView5.SelCount <> 0) then
  /// / имя пользователя
    begin
      UserSID := ListView5.Selected.SubItems[0];
      i := MessageDlg('Удаление профиля -' + ListView5.Selected.SubItems[2] +
        #10 + #13 + 'ВНИМАНИЕ удаляются все файлы пользователя!!!', mtWarning,
        [mbYes, mbCancel], 0);
      if i = mrYes then
        begin
          MyPS := ComboBox2.text;
          Memo1.Lines.Add('-------------------------------------');
          Memo1.Lines.Add('Удаляю профиль ' + ListView5.Selected.SubItems[1]);
          NewUserProfileDelete := UserProfileDelete.UserProfileDeleteThread.
            Create(true);
          NewUserProfileDelete.FreeOnTerminate := true;
          /// / под вопросом самоуничтожение потока
          NewUserProfileDelete.Start;
        end;
      if i = mrCancel then
        exit;
    end;
end;

procedure TfrmDomainInfo.N41Click(Sender: TObject);
begin
  SpeedButton49.Click;
end;

procedure TfrmDomainInfo.N42Click(Sender: TObject);
begin
  UpdateWindowsOffice;
end;

procedure TfrmDomainInfo.N43Click(Sender: TObject);
/// / сведения сетевого интерфейса
begin
  if (ListView7.Items.Count <> 0) and (ListView7.SelCount <> 0) then
  /// /
    begin

      MyPS := ComboBox2.text;
      NetworkInterfaceID := ListView7.Selected.SubItems[5];
      OKRightDlg1234567891011121314.showmodal;
    end;
end;

procedure TfrmDomainInfo.N44Click(Sender: TObject);
/// / свойсва TCP/IP
begin
  if (ListView7.Items.Count <> 0) and (ListView7.SelCount <> 0) then
    begin
      MyPS := ComboBox2.text;
      NetworkInterfaceID := ListView7.Selected.SubItems[5];
      OKRightDlg123456789101112131415.showmodal;
    end;
end;

procedure TfrmDomainInfo.G1Click(Sender: TObject);
var
  i: Integer;
  NewNetworkAdapterDisable, NewNetworkAdapterEnable: TThread;
begin
/// / отключить устройство , сетевой интерфейс

  if (ListView7.Items.Count <> 0) and (ListView7.SelCount <> 0) then
    begin
      MyPS := ComboBox2.text;
      NetworkInterfaceID := ListView7.Selected.SubItems[5];
      if ListView7.Selected.SubItems[3] = 'Отключено' then
        begin
          NewNetworkAdapterEnable :=
            NetworkAdapterEnable.NetworkAdapterEnableThread.Create(true);
          NewNetworkAdapterEnable.FreeOnTerminate := true;
          NewNetworkAdapterEnable.Start;
        end
      else
        begin
          i := MessageBox(self.Handle,
            PChar('Вы действительно хотите отключить адаптер?'),
            PChar(ListView7.Selected.SubItems[0]), MB_YESNO + MB_ICONQUESTION);
          if i = IDNO then
            exit;
          if i = IDYes then
            begin
              NewNetworkAdapterDisable :=
                NetworkAdapterDisable.NetworkAdapterDisableThread.Create(true);
              NewNetworkAdapterDisable.FreeOnTerminate := true;
              NewNetworkAdapterDisable.Start;
            end;
        end;
    end;
end;

procedure TfrmDomainInfo.N48Click(Sender: TObject);
/// /// групповая перезагрузка
begin
  ResetForListPC;
end;

procedure TfrmDomainInfo.N49Click(Sender: TObject);
var
  i: Integer;
  /// /// групповая принудительная перезагрузка
  NewSelectedPCShotDown: TThread;
begin
  if ListView8.SelCount = 0 then
    begin
      showmessage('Не выбран список компьютеров');
      exit;
    end;
  i := MessageDlg('Выполнить перезагрузку на выбранных компьютерах? ',
    mtWarning, [mbYes, mbCancel], 0);
  if i = mrYes then
    begin
      if ListView8.Items.Count > 0 then
        begin
          SelectedPCShutDown := TstringList.Create;
          for i := 0 to ListView8.Items.Count - 1 do
            begin
              if ListView8.Items[i].Checked = true then
                SelectedPCShutDown.Add(ListView8.Items[i].SubItems[0])
            end;
          if SelectedPCShutDown.Count <> 0 then
            begin
              myShutdown := 3;
              NewSelectedPCShotDown :=
                SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
              NewSelectedPCShotDown.FreeOnTerminate := true;
              NewSelectedPCShotDown.Start;
            end
          else
            showmessage('Не выбран не один компьютер!!!');
        end;
    end;
end;

procedure TfrmDomainInfo.N4Click(Sender: TObject);
var
  i, z, NewTread: Integer;
  p: TPoint;
  treadID: LongWord;
begin
  if (lvWorkStation.Items.Count <> 0) then
    begin
      if ListView1.Items.Count = 0 then
      /// если список служб пуст
        begin
          z := MessageBox(self.Handle, PChar('Загрузить список служб?'),
            PChar('Список служб пуст.'), MB_YESNO + MB_ICONQUESTION);
          if z = IDNO then
            exit
          else
            begin
              LoadListServices(ComboBox2.text, LabeledEdit1.text,
                LabeledEdit2.text, ListView1);
            end;
        end;
      ListView1.MultiSelect := false;
      ListView1.MultiSelect := true;
      GroupselectProc := lvWorkStation.Selected.SubItems[2];
      pcRes.ActivePage := TabSheet3;
      SpeedButton2.Caption := 'Службы';
      for i := 0 to ListView1.Items.Count - 1 do
        begin
          if ListView1.Items[i].SubItems[1] = GroupselectProc then
            begin
              ListView1.Items[i].Selected := true;
              ListView1.ItemIndex := i;
              ListView1.ItemFocused := ListView1.Items[i];
              p := ListView1.Items.Item[i].Position;
              ListView1.Scroll(p.X, p.Y);
            end;
        end;
    end;
end;

procedure TfrmDomainInfo.N50Click(Sender: TObject);
/// /// групповое завершение сеанса
begin
  LogOutForListPC;
  // SpeedButton62.Click;
end;

procedure TfrmDomainInfo.N51Click(Sender: TObject);
var
  i: Integer;
  /// /// групповое принудительное завершение работы
  NewSelectedPCShotDown: TThread;
begin
  if ListView8.SelCount = 0 then
    begin
      showmessage('Не выбран список компьютеров');
      exit;
    end;
  i := MessageDlg('Завершить работу на выбранных компьютерах? ', mtWarning,
    [mbYes, mbCancel], 0);
  if i = mrYes then
    begin
      if ListView8.Items.Count > 0 then
        begin
          SelectedPCShutDown := TstringList.Create;
          for i := 0 to ListView8.Items.Count - 1 do
            begin
              if ListView8.Items[i].Checked = true then
                SelectedPCShutDown.Add(ListView8.Items[i].SubItems[0])
            end;
          if SelectedPCShutDown.Count <> 0 then
            begin
              myShutdown := 5;
              NewSelectedPCShotDown :=
                SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
              NewSelectedPCShotDown.FreeOnTerminate := true;
              NewSelectedPCShotDown.Start;
            end
          else
            showmessage('Не выбран не один компьютер!!!');
        end;
    end;

end;

procedure TfrmDomainInfo.Ping1Click(Sender: TObject);
var
  i: Integer;
  /// // групповой пинг
begin
  // OutForPing:=false;
  if ListView8.Items.Count > 0 then
    begin
      PingPCList := TstringList.Create;
      for i := 0 to ListView8.Items.Count - 1 do
        begin
          ListView8.Items[i].ImageIndex := 0;
          if ListView8.Items[i].Checked = true then
            PingPCList.Add(ListView8.Items[i].SubItems[0]);
        end;
      if PingPCList.Count <> 0 then
        begin
          OutForPing := true;
          SelectedPCPing := GeneralPingPC.GeneralPing.Create(true);
          SelectedPCPing.FreeOnTerminate := true;
          SelectedPCPing.Start;
        end
      else
        showmessage('Не выбран не один компьютер!!!');
    end;
end;


procedure TfrmDomainInfo.N52Click(Sender: TObject);
var
  i: Integer;
  /// /// групповое принудительное завершение сеанса
  NewSelectedPCShotDown: TThread;
begin
  if ListView8.SelCount = 0 then
    begin
      showmessage('Не выбран список компьютеров');
      exit;
    end;
  i := MessageDlg('Завершить сеанс на выбранных компьютерах? ', mtWarning,
    [mbYes, mbCancel], 0);
  if i = mrYes then
    begin
      if ListView8.Items.Count > 0 then
        begin
          SelectedPCShutDown := TstringList.Create;
          for i := 0 to ListView8.Items.Count - 1 do
            begin
              if ListView8.Items[i].Checked = true then
                SelectedPCShutDown.Add(ListView8.Items[i].SubItems[0])
            end;
          if SelectedPCShutDown.Count <> 0 then
            begin
              myShutdown := 1;
              NewSelectedPCShotDown :=
                SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
              NewSelectedPCShotDown.FreeOnTerminate := true;
              NewSelectedPCShotDown.Start;
            end
          else
            showmessage('Не выбран не один компьютер!!!');
        end;
    end;

end;

procedure TfrmDomainInfo.N54Click(Sender: TObject);
/// /// групповое завершение работы
begin
  PowerOffForListPC;
  // SpeedButton56.Click;
end;

procedure TfrmDomainInfo.N55Click(Sender: TObject);
begin
  SettingsProgram.showmodal;
end;

procedure TfrmDomainInfo.N56Click(Sender: TObject);
var
  s: string;
  q: Integer;
  OpenFile: TOpenDialog;
begin
/// удаление обновлений Windows
  if (ListView9.Items.Count <> 0) and (ListView9.SelCount = 1) then
    begin
      NewProcMyPS := ComboBox2.text;
      MyPS := ComboBox2.text;
      s := ListView9.Selected.SubItems[0];
      q := MessageDlg('Удалить обновление - ' + s + '?', mtWarning,
        [mbYes, mbNo], 0);
      if q = mrCancel then
        exit;
      if copy(OSVersion, 1, 2) = '10' then
      /// // если стоит Windows 10
        Begin
          showmessage
            ('Для Windows 10 укажите доступный сетевой ресурс с файлом обновлений '
            + s);
          OpenFile := TOpenDialog.Create(frmDomainInfo);
          try
            OpenFile.Name := 'OpenFileMSU';
            OpenFile.Filter := 'Файлы обновлений |*.msu|';
            OpenFile.Title := 'Укажите файл обновлений ' + s +
              ' для удаления (только для Windows 10/Server 2016)';
            if OpenFile.Execute then
              begin
                s := OpenFile.filename;
                MyCommandLine := 'wusa.exe /uninstall "' + s +
                  '" /quiet /norestart';
                // 'DISM.exe /Online /Remove-Package /PackageName:'+s+' /quiet /norestart';
                /// /  wusa.exe /uninstall /KB:'+s+' /norestart'
              end
            else
              exit;
          finally
            OpenFile.Free;
          end;
        End
      else
        begin
        /// / если другая
          s := StringReplace(s, 'KB', '', [rfReplaceAll]);
          // Delete(s,1,2);
          MyCommandLine := 'wusa.exe /uninstall /KB:' + s +
            ' /quiet /norestart';
          // memo1.Lines.Add('Запускаю процесс для удаления обновления - '+MyCommandLine);
        end;
      if InputQuery('Запуск процесса', 'Выполнить:', MyCommandLine) then
        begin
          UnProg := unit5.NewProcess.Create(true);
          UnProg.FreeOnTerminate := true;
          UnProg.Start;
        end;
    end;
end;

procedure TfrmDomainInfo.N57Click(Sender: TObject);
var
  OpenFile: TOpenDialog;
begin
/// установка обновлений Windows
  if ping(ComboBox2.text) then
    begin
      NewProcMyPS := ComboBox2.text;
      MyPS := ComboBox2.text;
      showmessage('Укажите доступный сетевой ресурс с файлом обновлений.');
      OpenFile := TOpenDialog.Create(frmDomainInfo);
      try
        OpenFile.Name := 'OpenFileMSU';
        OpenFile.Filter := 'Файлы обновлений |*.msu|';
        OpenFile.Title := 'Для установки обновлений. Укажите файл обновлений';
        if OpenFile.Execute then
          begin
            MyCommandLine := OpenFile.filename;
            MyCommandLine := 'wusa.exe "' + MyCommandLine +
              '" /quiet /norestart';
            if InputQuery('Запуск процесса', 'Выполнить:', MyCommandLine) then
              begin
                // 'DISM.exe /Online /Remove-Package /PackageName:'+s+' /quiet /norestart';
                /// /  wusa.exe /uninstall /KB:'+s+' /norestart'
                Memo1.Lines.Add('Запускаю процесс для установки обновления - ' +
                  MyCommandLine);
                UnProg := unit5.NewProcess.Create(true);
                UnProg.FreeOnTerminate := true;
                UnProg.Start;
              end;
          end;
      finally
        FreeAndNil(OpenFile);
      end;
    end;
end;

procedure TfrmDomainInfo.N59Click(Sender: TObject);
begin
  if lvWorkStation.Items.Count <> 0 then
    begin
      popupListViewSaveAs(lvWorkStation, 'Сохранение списка процессов',
        'Процессы ' + ComboBox2.text);
    end;
end;

procedure TfrmDomainInfo.N61Click(Sender: TObject);
begin
/// / закрыть форму из трея
  frmDomainInfo.Close;
end;

procedure TfrmDomainInfo.N62Click(Sender: TObject);
/// /// групповой запуск процесса
begin
  RunProcessForListPC;
end;

procedure TfrmDomainInfo.N63Click(Sender: TObject);
/// /// групповое завершение процесса
begin
  KillProcessForListPC;
end;

procedure TfrmDomainInfo.N64Click(Sender: TObject);
/// /// групповая установка драйверов принтера
begin
  AddDriverPrintForListPC;
end;

procedure TfrmDomainInfo.N65Click(Sender: TObject);
/// //// групповое добавление принтеров
begin
  AddPrinterForListPC;
end;

procedure TfrmDomainInfo.N67Click(Sender: TObject);
begin
/// инвентаризация оборудования
  InventoryHardware;
end;

procedure TfrmDomainInfo.N68Click(Sender: TObject);
begin
  inventorySoftForSelectPC(ComboBox2.text, LabeledEdit1.text,
    LabeledEdit2.text);
end;

function TfrmDomainInfo.inventorySoftForSelectPC(NamePC, user,
  pass: string): bool;
begin
/// / иныентаризация выбранного компа
  if not Datam.ConnectionDB.Connected then
    begin
      showmessage('Для инвентаризации подключите базу данных!!!');
      exit;
    end;
  if InventSoft then
    begin
      showmessage('Дождитесь завершения текущей инвентаризации');
      exit;
    end;
  if NamePC = '' then
    begin
      showmessage('Не указано имя компьютера');
      exit;
    end;
  ListPCConf := TstringList.Create;
  ListPCConf.Add(NamePC);
  InventSoft := true;
  MyUser := user;
  MyPasswd := pass;
  ThrInvSoft := MyInventorySoft.InventorySoft.Create(true);
  ThrInvSoft.FreeOnTerminate := true;
  ThrInvSoft.Start;
end;

procedure TfrmDomainInfo.N5Click(Sender: TObject);
var
  Myservice: TThread;
begin
  if (ListView1.Items.Count <> 0) and (ListView1.SelCount <> 0) then
  /// / запуск службы
    begin
      MyPS := ComboBox2.text;
      SelectServic := ListView1.Selected.SubItems[0];
      ActionServic := 1;
      Myservice := unit6.Service.Create(true);
      Myservice.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      Myservice.Start;
    end;
end;

procedure TfrmDomainInfo.N6Click(Sender: TObject);
var
  Myservice: TThread;
begin
  if (ListView1.Items.Count <> 0) and (ListView1.SelCount <> 0) then
  /// / остановка службы
    begin
      MyPS := ComboBox2.text;
      SelectServic := ListView1.Selected.SubItems[0];
      ActionServic := 2;
      Myservice := unit6.Service.Create(true);
      Myservice.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      Myservice.Start;
    end;
end;

procedure TfrmDomainInfo.N71Click(Sender: TObject);
begin
  if (ListView8.Items.Count <> 0) and (ListView8.SelCount <> 0) then
    begin
      if ping(ListView8.Items[ListView8.ItemFocused.Index].SubItems[0]) then
        begin
          if ListView8.Items[ListView8.ItemFocused.Index].ImageIndex <> 12 then
            ListView8.Items[ListView8.ItemFocused.Index].ImageIndex := 4
        end
      else
        begin
          if ListView8.Items[ListView8.ItemFocused.Index].ImageIndex <> 12 then
            ListView8.Items[ListView8.ItemFocused.Index].ImageIndex := 5;
        end;
    end;
end;

/// ///////////////////////////////////////////////////////
procedure TfrmDomainInfo.LabeledEdit10KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    SpeedButton56.Click;
end;

procedure TfrmDomainInfo.LabeledEdit1Change(Sender: TObject);
begin
  MyUser := LabeledEdit1.text;
end;

procedure TfrmDomainInfo.LabeledEdit2Change(Sender: TObject);
begin
  MyPasswd := LabeledEdit2.text;
end;

procedure TfrmDomainInfo.LabeledEdit9KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    SpeedButton52.Click;
end;

Function TfrmDomainInfo.ListToExcel(ListView: TListView; filename: string;
  FileFormat: string): bool;
var
/// / экспорт в Excel
  i, z: Integer;
  Myworksheets: Variant;
  /// лист в ексель
  ProgressExport: TprogressBar;
begin
  try
    if (ListView.Items.Count = 0) then
      exit;
    try
      if not Assigned(ProgressExport) then
        begin
          ProgressExport := TprogressBar.Create(self);
          ProgressExport.Parent := StatusBar1;
          ProgressExport.Name := 'ProgExp';
        end;
      ProgressExport.left := StatusBar1.Panels[0].width + StatusBar1.Panels[1]
        .width + StatusBar1.Panels[2].width + StatusBar1.Panels[3].width;
      ProgressExport.width := StatusBar1.Panels[4].width;
      ProgressExport.Min := 0;
      ProgressExport.Max := ListView.Items.Count - 1;
      ProgressExport.Step := 1;
    except
      on E: Exception do
        if Assigned(ProgressExport) then
          FreeAndNil(ProgressExport);
    end;

    AllRunExcel;
    /// запуск Excel
    MyExcel.WorkBooks[1].WorkSheets[1].Name := 'Список';
    /// имя листа
    Myworksheets := MyExcel.WorkBooks[1].WorkSheets[1];
    // MyWorkSheets.range['A:AK'].Columns.AutoFit; /// авто ширина ячеек
    // myworksheets.columns['A:AK'].EntireColumn.AutoFit;  /// авто ширина колонок
    Myworksheets.Columns['A:AK'].columnwidth := 20; // ширина колонки
    Myworksheets.Columns['A:AK'].wraptext := true;
    /// Перенос по словам
    Myworksheets.Columns['A:AK'].numberFormat := '@';
    /// текстовый формат
    for i := 0 to ListView.Columns.Count - 1 do
      begin
        Myworksheets.cells[1, i + 1] := ListView.Columns[i].Caption;
      end;
    for i := 0 to ListView.Items.Count - 1 do
      begin
        if Assigned(ProgressExport) then
          ProgressExport.Position := i; // Позиция прогрессбара
        Myworksheets.cells[i + 2, 1] := ListView.Items[i].Caption;
        /// запись первого столбца, т.к. он caption , следующие идут subitems[]
        for z := 1 to ListView.Columns.Count - 1 do
          begin
            Myworksheets.cells[i + 2, z + 1] := ListView.Items[i]
              .SubItems[z - 1];
          end;
      end;
    MyExcel.Visible := true; // показать Excel
    if FileFormat = '.xls' then
      Result := SaveWorkBook(filename, 1, '56');
    /// Имя файла, номер листа и тип файла  // 6-csv/56-xls/44-html/60-ods/51-xlsx
    if FileFormat = '.xlsx' then
      Result := SaveWorkBook(filename, 1, '51');
    /// Имя файла, номер листа и тип файла
    if FileFormat = '.html' then
      Result := SaveWorkBook(filename, 1, '44');
    /// Имя файла, номер листа и тип файла
    if FileFormat = '.csv' then
      Result := SaveWorkBook(filename, 1, '6');
    if FileFormat = '.ods' then
      Result := SaveWorkBook(filename, 1, '60');
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export Excel File - ' + E.Message);
        Result := false;
        VarClear(Myworksheets);
        VariantClear(MyExcel);
        VariantClear(MyExcel);
        if Assigned(ProgressExport) then
          FreeAndNil(ProgressExport);
      end;
  end;
  if Assigned(ProgressExport) then
    FreeAndNil(ProgressExport);
  VarClear(Myworksheets);
  VariantClear(MyExcel);
  VariantClear(MyExcel);
end;

/// ///////////////////////////////////////////////////////
function TfrmDomainInfo.ListToTxT(ListView: TListView; NameFile: string): bool;
var
  i, z: Integer;
  Mylist: TstringList;
Begin
  try
    if ListView.Items.Count <> 0 then
      begin
        Mylist := TstringList.Create;
        for i := 0 to ListView.Items.Count - 1 do
          begin
            Mylist.Add('');
            for z := 0 to ListView.Columns.Count - 2 do
              begin
                Mylist[i] := Mylist[i] + ListView.Columns[z + 1].Caption + ': '
                  + ListView.Items[i].SubItems[z] + '; ';
              end;
          end;
        Mylist.SaveToFile(NameFile);
        FreeAndNil(Mylist);
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export txt file - ' + E.Message);
        Result := false;

      end;
  end;
  if Assigned(Mylist) then
    FreeAndNil(Mylist);
End;

/// //////////////////////////////////////////////////////
function TfrmDomainInfo.SaveTreeView(TV: TTreeView; NamePC: string): Boolean;
// функция сохранения свойств компьютера из TreeView  в txt
var
  Mylist: TstringList;
  i, ParentItem: Integer;
  SaveFileDlg: TSaveTextFileDialog;
begin
  Mylist := TstringList.Create;
  try
    try
      for i := 0 to TV.Items.Count - 1 do
        begin
          if i <> 0 then
            begin
              if TV.Items[i].ImageIndex <> 0 then
                begin
                  Mylist.Add(TV.Items[i].text);
                  ParentItem := Mylist.Count - 1;
                end
              else
                Mylist[ParentItem] := Mylist[ParentItem] + '\' +
                  TV.Items[i].text;
            end
          else
            Mylist.Add(TV.Items[i].text);
        end;
    except
      on E: Exception do
        begin
          Mylist.Clear;
          for i := 0 to TV.Items.Count - 1 do
            Mylist.Add(TV.Items[i].text);
        end;
    end;
    SaveFileDlg := TSaveTextFileDialog.Create(nil);
    begin
      SaveFileDlg.Title := 'Конфигурация компьютера ' + NamePC;
      SaveFileDlg.InitialDir := ExtractFilePath(Application.ExeName) +
        'reports\';
      SaveFileDlg.Filter := 'Текстовый файл(*.txt)|*.TXT';
      SaveFileDlg.FilterIndex := 0;
      SaveFileDlg.filename := 'Конфигурация ' + NamePC + '.txt';
      if SaveFileDlg.Execute then
        begin
          Mylist.SaveToFile(SaveFileDlg.filename);
        end;
      try
        ShellApi.ShellExecute(0, 'Open', PChar(SaveFileDlg.filename), nil, nil,
          SW_SHOWNORMAL);
      except
        showmessage('Ошибка при открытии приложения')
      end;
    end;
  finally
    Mylist.Free;
    SaveFileDlg.Free;
  end;

end;

/// ///////////////////////////////////////////////////////
function TfrmDomainInfo.popupListViewSaveAs(ListView: TListView;
  SaveDlgTitle, filename: string): bool;
var
  SaveFileDlg: TSaveTextFileDialog;
begin
  if ListView.Items.Count <> 0 then
    begin
      SaveFileDlg := TSaveTextFileDialog.Create(nil);
      SaveFileDlg.Title := SaveDlgTitle;
      SaveFileDlg.InitialDir := ExtractFilePath(Application.ExeName) +
        'reports\';
      SaveFileDlg.Filter := 'Книга Excel 97-2003 (*.XLS)|*.XLS|' +
        'Книга Open XML (*.xlsx)|*.xlsx|Формат HTML (*.html)|*.html|Файл csv (*.csv)|*.csv|'
        + 'Электронная таблица OpenDocument (*.ods)|*.ods|Текстовый файл(*.txt)|*.TXT';
      SaveFileDlg.FilterIndex := 1;
      SaveFileDlg.filename := filename;
      if SaveFileDlg.Execute then
        begin
          case SaveFileDlg.FilterIndex of
            1:
              begin
                SaveFileDlg.DefaultExt := '.xls';
                ListToExcel(ListView, SaveFileDlg.filename,
                  SaveFileDlg.DefaultExt);
              end;
            2:
              begin
                SaveFileDlg.DefaultExt := '.xlsx';
                ListToExcel(ListView, SaveFileDlg.filename,
                  SaveFileDlg.DefaultExt);
              end;
            3:
              begin
                SaveFileDlg.DefaultExt := '.html';
                ListToExcel(ListView, SaveFileDlg.filename,
                  SaveFileDlg.DefaultExt);
              end;
            4:
              begin
                SaveFileDlg.DefaultExt := '.csv';
                ListToExcel(ListView, SaveFileDlg.filename,
                  SaveFileDlg.DefaultExt);
              end;
            5:
              begin
                SaveFileDlg.DefaultExt := '.ods';
                ListToExcel(ListView, SaveFileDlg.filename,
                  SaveFileDlg.DefaultExt);
              end;
            6:
              begin
                SaveFileDlg.DefaultExt := '.txt';
                ListToTxT(ListView, SaveFileDlg.filename +
                  SaveFileDlg.DefaultExt);
              end;
              /// сохраняем в txt  ListToTxT - функция сохранения TlistView в текстовый файл
          end;
        end;
      FreeAndNil(SaveFileDlg);
    end;
  if Assigned(SaveFileDlg) then
    FreeAndNil(SaveFileDlg);
  Result := true;
end;

/// ///////////////////////////////////////////////////////
procedure TfrmDomainInfo.N72Click(Sender: TObject);
begin
  if ListView8.Items.Count <> 0 then
    begin
      popupListViewSaveAs(ListView8, 'Сохранение списка компьютеров', '');
    end;
end;

function TfrmDomainInfo.findForListView(ListViewIn: TListView): bool;
var
  p: TPoint;
  i, z: Integer;
  StringFind: string;
  scrpos: SCROLLINFO;
begin
  try
    if (ListViewIn.Items.Count <> 0) and (ListViewIn.ItemIndex <> 0) then
      Begin

        ListViewMousPoint.X := ListViewMousPoint.X - ListViewIn.left;
        ListViewMousPoint.Y := ListViewMousPoint.Y - ListViewIn.top;
        scrpos.fMask := SIF_TRACKPOS;
        GetScrollInfo(ListViewIn.Handle, SB_HORZ, scrpos);
        /// получаем информацию о положении горизонтального скрола
        ListViewMousPoint.X := ListViewMousPoint.X + scrpos.nTrackPos -
          (ListViewIn.Column[0].width);
        // отнимаем размер первой колонки и панели с лева от листа
        for i := 1 to ListViewIn.Columns.Count - 1 do
        // проходим по цыклу все колонки
          begin
            if ListViewIn.Column[i].width < ListViewMousPoint.X then
              dec(ListViewMousPoint.X, ListViewIn.Column[i].width)
            else
              begin
                if not InputQuery('Поиск по колонке',
                  ListViewIn.Column[i].Caption + ':', StringFind) then
                  exit;
                if (StringFind = '') then
                  exit;
                for z := 0 to ListViewIn.Items.Count - 1 do
                  begin
                    if pos(Ansiuppercase(StringFind),
                      Ansiuppercase(ListViewIn.Items[z].SubItems[i - 1])) <> 0
                    then
                      begin
                        ListViewIn.Items[z].Selected := true;
                        ListViewIn.ItemIndex := z;
                        ListViewIn.ItemFocused := ListViewIn.Items[z];
                        p := ListViewIn.Items.Item[z].Position;
                        ListViewIn.Scroll(p.X, p.Y);
                        break;
                      end;
                  end;
                break;
              end;
          end;
      End;
    if ListViewIn.ItemFocused <> nil then
      ListViewIn.ItemFocused.MakeVisible(false);
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка поиска: ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.N73Click(Sender: TObject);
begin
  findForListView(ListView8);
end;

procedure TfrmDomainInfo.N74Click(Sender: TObject);
begin
  if (ListView8.Items.Count <> 0) and (ListView8.SelCount <> 0) then
    begin
      try
        JvNetscapeSplitter1.Maximized := false;
        /// раскрыть панель свойств компьютера
        GroupBox2.Caption := ListView8.Selected.SubItems[0];
        // передаем имя компа в caption для других функций
        ListViewSoftinPC.Clear; // очистка списка
        ListViewMicLic.Clear; // очистка списка
        ListViewAV.Clear; // очистка списка

        // читаем список оборудования
        Datam.FDQueryRead2.SQL.Clear;
        /// чтение конфигурации первого вхождения и передача в treeview
        Datam.FDQueryRead2.SQL.text :=
          'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
          (ListView8.Selected.SubItems[0]) + '''' + ' ORDER BY DATE_INV DESC';
        Datam.FDQueryRead2.Open;
        createtreeView(Datam.FDQueryRead2, TreeView4);
        // Функция постороения дерева
        Datam.FDQueryRead2.SQL.Clear; // очистить
        Datam.FDQueryRead2.Close;
        /// закрыть нах после чтения

        // читаем список установленных программ
        readSoftRorSelectPC(ListView8.Selected.SubItems[0], ListViewSoftinPC);

        // читаем список лицензий
        readMicrosoftLic(ListView8.Selected.SubItems[0], ListViewMicLic);
        // читаем список антивирусных продуктов
        readAntivirusStatus(ListView8.Selected.SubItems[0], ListViewAV);

      except
        on E: Exception do
          Memo1.Lines.Add('Ошибка : ' + E.Message);
      end;
    end;
end;

procedure TfrmDomainInfo.N75Click(Sender: TObject);
begin
  if ListView1.Items.Count <> 0 then
    popupListViewSaveAs(ListView1, 'Сохранение списка служб',
      'Службы ' + ComboBox2.text);

end;

procedure TfrmDomainInfo.N76Click(Sender: TObject);
begin
  if ListView2.Items.Count <> 0 then
    popupListViewSaveAs(ListView2, 'Сохранение списка автозагрузки',
      'Автозагрузка ' + ComboBox2.text);

end;

procedure TfrmDomainInfo.N77Click(Sender: TObject);
begin
  if ListView3.Items.Count <> 0 then
    popupListViewSaveAs(ListView3,
      'Сохранение списка установленных программ (msi)',
      'Программы msi ' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N78Click(Sender: TObject);
begin
  if ListView4.Items.Count <> 0 then
    popupListViewSaveAs(ListView4,
      'Сохранение списка всех установленных программ',
      'Программы ' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N79Click(Sender: TObject);
begin
  if ListView5.Items.Count <> 0 then
    popupListViewSaveAs(ListView5, 'Сохранение списка профилей пользователей',
      'Профили ' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N80Click(Sender: TObject);
begin
  if ListView6.Items.Count <> 0 then
    popupListViewSaveAs(ListView6, 'Сохранение списка PnP устройств',
      'Устройства ' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N81Click(Sender: TObject);
begin
  if ListView7.Items.Count <> 0 then
    popupListViewSaveAs(ListView7, 'Сохранение списка сетевых интерфейсов',
      'Сетевые интерфейсы ' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N82Click(Sender: TObject);
begin
  if ListView9.Items.Count <> 0 then
    popupListViewSaveAs(ListView9, 'Сохранение списка обновлений',
      'Windows Update' + ComboBox2.text);
end;

procedure TfrmDomainInfo.N84Click(Sender: TObject);
begin
  Form10.showmodal;
end;

procedure TfrmDomainInfo.N85Click(Sender: TObject);
begin
  try
    ShellApi.ShellExecute(0, 'Open', PChar('https://skrblog.ru/manual/'), nil,
      nil, SW_SHOWNORMAL);
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка при открытии приложения : ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.N130Click(Sender: TObject);
begin
  try
    ShellApi.ShellExecute(0, 'Open', PChar('https://skrblog.ru/buy/'), nil, nil,
      SW_SHOWNORMAL);
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка при открытии приложения : ' + E.Message);
  end;
end;

function TfrmDomainInfo.OpenRegEdit(RootKey, Path, NamePC: string): Boolean;
var
  i: Integer;
begin
  try
    RegEdit.ComboBox2.Clear;
    for i := 0 to ComboBox2.Items.Count - 1 do
      begin
        RegEdit.ComboBox2.Items.Add(ComboBox2.Items[i]);
        RegEdit.ComboBox2.ItemsEx[i].ImageIndex := ComboBox2.ItemsEx[i]
          .ImageIndex;
      end;
    RegEdit.ComboBox2.text := NamePC;
    /// имя компа
    RegEdit.EditUser.text := LabeledEdit1.text;
    /// пользователь
    RegEdit.EditPass.text := LabeledEdit2.text; // пароль
    if RegEdit.WindowState = wsMinimized then
    // если форма свернута то восстанавливаем ее
      RegEdit.WindowState := wsNormal
    else // иначе показываем
      RegEdit.Show;
    if (RootKey <> '') and (Path <> '') then
    // если директории не пустые то переходим в нужную
      begin
        RegEdit.MemoLog.Lines.Add(RootKey + ':' + Path);
        RegEdit.EditRooT.text := RootKey;
        RegEdit.JumpTreeView(RootKey, Path);
      end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка открытия реестра : ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.N131Click(Sender: TObject);
// открыть в реестре профиль пользователя
begin
  try
    if ListView5.SelCount = 1 then // если выделена одна строка
      if not ping(ComboBox2.text) then
        exit;
    if ListView5.Selected.SubItems[0] <> '' then // если строка с SID не пуста
      begin
        OpenRegEdit('HKEY_LOCAL_MACHINE',
          'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' +
          ListView5.Selected.SubItems[0], ComboBox2.text);
      end
    else
      showmessage('Не указан SID пользователя');
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка открытия реестра : ' + E.Message);
  end
end;

procedure TfrmDomainInfo.N132Click(Sender: TObject);
// открыть в реестре путь из списка Все программы
begin
  try
    if ListView4.SelCount = 1 then // если выделена одна строка
      if not ping(ComboBox2.text) then
        exit;
    if ListView4.Selected.SubItems[7] <> '' then
    // если строка со значением не пуста
      OpenRegEdit('HKEY_LOCAL_MACHINE',
        'Software\Microsoft\Windows\CurrentVersion\Uninstall\' +
        ListView4.Selected.SubItems[7], ComboBox2.text)
    else
      OpenRegEdit('HKEY_LOCAL_MACHINE',
        'Software\Microsoft\Windows\CurrentVersion\Uninstall', ComboBox2.text)
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка открытия реестра : ' + E.Message);
  end
end;

procedure TfrmDomainInfo.N133Click(Sender: TObject);
begin
  if ListView8.SelCount <> 1 then
    exit;
  if not ping(ListView8.Selected.SubItems[0]) then
    exit;
  OpenRegEdit('', '', ListView8.Selected.SubItems[0]);
  // передаем пустые переменные чтобы просто открыть реестр
end;

procedure TfrmDomainInfo.N134Click(Sender: TObject);
// открыть в реестре путь из автозагрузки
var
  hDefKey: string;
  sNameValue: string;
  Res: Integer;
begin
  case hk(frmDomainInfo.ListView2.Selected.SubItems[2]) of
    1:
      hDefKey := 'HKEY_CLASSES_ROOT';
    2:
      hDefKey := 'HKEY_CURRENT_USER';
    3:
      hDefKey := 'HKEY_LOCAL_MACHINE';
    4:
      hDefKey := 'HKEY_USERS';
    5:
      hDefKey := 'HKEY_CURRENT_CONFIG';
  end;
  try
    if ListView2.SelCount = 1 then // если выделена одна строка
      begin
        if not ping(ComboBox2.text) then
          exit;
        if (pos('Startup', ListView2.Selected.SubItems[2]) <> 0) then
        // если элемент автозагузки из реестра а не из  Startup
          begin // открываем в реестра
            Res := MessageBox(self.Handle,
              PChar('Открыть значение в проводнике?'), PChar('Автозагрузка'),
              MB_YESNO);
            if Res = IDYes then
              begin
                N108.Click;
                exit;
              end;
          end;
        if (ListView2.Selected.SubItems[2] <> '') and
          (pos('Startup', ListView2.Selected.SubItems[2]) = 0) then
        // если строка со значением не пуста
          begin
            sNameValue := copy(ListView2.Selected.SubItems[2],
              pos('\', ListView2.Selected.SubItems[2]) + 1,
              Length(ListView2.Selected.SubItems[2]));
            // ShowMessage(hDefKey+':'+sNameValue);
            OpenRegEdit(hDefKey, sNameValue, ComboBox2.text);
          end;
      end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка открытия реестра : ' + E.Message);
  end

end;

procedure TfrmDomainInfo.N135Click(Sender: TObject); // открыть службу в реестре
begin
  if not ping(ComboBox2.text) then
    exit;
  if ListView1.SelCount = 1 then
    begin
      OpenRegEdit('HKEY_LOCAL_MACHINE', 'SYSTEM\CurrentControlSet\services\' +
        ListView1.Selected.SubItems[0], ComboBox2.text);
    end;
end;

procedure TfrmDomainInfo.N136Click(Sender: TObject);
// открыть принтер в реестре
begin
  if not ping(ComboBox2.text) then
    exit;
  if ListViewPrint.SelCount = 1 then
    begin
      OpenRegEdit('HKEY_LOCAL_MACHINE',
        'SYSTEM\CurrentControlSet\Control\Print\Printers\' +
        ListViewPrint.Selected.Caption, ComboBox2.text);
    end;
end;

procedure TfrmDomainInfo.N137Click(Sender: TObject);
// открыть в реестре сетевой интерфейс
begin
  if not ping(ComboBox2.text) then
    exit;
  if ListView7.SelCount = 1 then
    begin
      if ListView7.Selected.SubItems[7] <> '' then
        OpenRegEdit('HKEY_LOCAL_MACHINE',
          'SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces\' +
          ListView7.Selected.SubItems[7], ComboBox2.text)
      else
        OpenRegEdit('HKEY_LOCAL_MACHINE',
          'SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces',
          ComboBox2.text)
    end;
end;

procedure TfrmDomainInfo.N138Click(Sender: TObject);
// открыть драйвер в реестре
begin
  if not ping(ComboBox2.text) then
    exit;
  if ListView6.SelCount = 1 then
    begin
      if ListView6.Selected.SubItems[3] <> '' then
        OpenRegEdit('HKEY_LOCAL_MACHINE', 'SYSTEM\CurrentControlSet\Enum\' +
          ListView6.Selected.SubItems[3], ComboBox2.text)
      else
        OpenRegEdit('HKEY_LOCAL_MACHINE', 'SYSTEM\CurrentControlSet\Enum',
          ComboBox2.text)
    end;
end;

procedure TfrmDomainInfo.N139Click(Sender: TObject); //
begin
  if not ping(ComboBox2.text) then
    exit;
  if ListViewShare.SelCount = 1 then
    begin
      OpenRegEdit('HKEY_LOCAL_MACHINE',
        'SYSTEM\CurrentControlSet\Services\LanmanServer\Shares',
        ComboBox2.text);
    end;
end;

procedure TfrmDomainInfo.Button19Click(Sender: TObject);
/// просто открыть реестр
var
  i: Integer;
begin
  if not ping(ComboBox2.text) then
    exit;
  OpenRegEdit('', '', ComboBox2.text);
  // передаем пустые переменные чтобы просто открыть реестр
end;

procedure TfrmDomainInfo.N86Click(Sender: TObject);
begin
  if ListViewDisk.Items.Count <> 0 then
    begin
      popupListViewSaveAs(ListViewDisk, 'Сохранение списка дисков',
        'Локальные диски ' + ComboBox2.text);
    end;
end;

procedure TfrmDomainInfo.PagePropertiesPCChange(Sender: TObject);
begin
  case PagePropertiesPC.ActivePageIndex of
    0:
      if TreeView4.Items.Count < 13 then
        begin
          // читаем список оборудования
          Datam.FDQueryRead2.SQL.Clear;
          /// чтение конфигурации первого вхождения и передача в treeview
          Datam.FDQueryRead2.SQL.text :=
            'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' + (GroupBox2.Caption) +
            '''' + ' ORDER BY DATE_INV DESC';
          Datam.FDQueryRead2.Open;
          createtreeView(Datam.FDQueryRead2, TreeView4);
          // Функция постороения дерева
          Datam.FDQueryRead2.SQL.Clear; // очистить
          Datam.FDQueryRead2.Close;
          /// закрыть нах после чтения
        end;
    1:
      begin // читаем список установленных программ
        if ListViewSoftinPC.Items.Count = 0 then
          readSoftRorSelectPC(GroupBox2.Caption, ListViewSoftinPC);
      end;
    2:
      begin // читаем список лицензий
        if ListViewMicLic.Items.Count = 0 then
          readMicrosoftLic(GroupBox2.Caption, ListViewMicLic);
      end;
    3:
      begin // читаем список антивирусных продуктов
        if ListViewAV.Items.Count = 0 then
          readAntivirusStatus(GroupBox2.Caption, ListViewAV);
      end;
  end;
end;

procedure TfrmDomainInfo.PCInventChange(Sender: TObject);
begin
  case PCInvent.ActivePageIndex of
    2:
      begin
        if LVMicrosoft.Items.Count = 0 then
          UpdateWindowsOffice; // Обновляем список продуктов Microsoft
      end;
    3:
      begin
        if LVAntivirus.Items.Count = 0 then
          UpdateAntivirusProduct; // Обновляем список антивирусного обеспечения
      end;
  end;
end;

function TfrmDomainInfo.PClistDelAdd(s: string; oper: bool): bool;
var
  i: Integer;
begin
  if not OutForPing then
    exit;
  if Assigned(PingPCList) then
    case oper of
      false:
        begin
          if PingPCList.Count > 0 then
            begin
              for i := 0 to PingPCList.Count - 1 do
                if PingPCList[i] = s then
                  begin
                    PingPCList.Delete(i);
                    break;
                  end;
            end;
        end;
      true:
        begin
          if PingPCList.Count > 0 then
            begin
              for i := 0 to PingPCList.Count - 1 do
                if PingPCList[i] <> s then
                  begin
                    PingPCList.Add(s);
                    break;
                  end;
            end;
        end;
    end;
end;

procedure TfrmDomainInfo.N88Click(Sender: TObject);
/// // включить в список исключения
begin
  if (ListView8.Items.Count <> 0) and (ListView8.SelCount <> 0) then
    begin
      if Datam.EXPTCurrentPC(ListView8.Items[ListView8.ItemFocused.Index]
        .SubItems[0], 0) then
      /// /операция включения компа в список на сканирование
        begin
          Memo1.Lines.Add(ListView8.Items[ListView8.ItemFocused.Index].SubItems
            [0] + ' - включен в список сканирования');
          PClistDelAdd(ListView8.Items[ListView8.ItemFocused.Index].SubItems
            [0], true);
          ListView8.Items[ListView8.ItemFocused.Index].ImageIndex := 0;
        end
      else
        Memo1.Lines.Add(ListView8.Items[ListView8.ItemFocused.Index].SubItems[0]
          + ' - не удалось провести включение');
    end;

end;

procedure TfrmDomainInfo.N89Click(Sender: TObject);
/// // исключить из списка исключений
begin
  if (ListView8.Items.Count <> 0) and (ListView8.SelCount <> 0) then
    begin
      if Datam.EXPTCurrentPC(ListView8.Items[ListView8.ItemFocused.Index]
        .SubItems[0], 12) then
      /// /операция исключения компа в список на сканирование
        begin
          Memo1.Lines.Add(ListView8.Items[ListView8.ItemFocused.Index].SubItems
            [0] + ' - исключен из списка сканирования');
          PClistDelAdd(ListView8.Items[ListView8.ItemFocused.Index].SubItems
            [0], false);
          ListView8.Items[ListView8.ItemFocused.Index].ImageIndex := 12;
        end
      else
        Memo1.Lines.Add(ListView8.Items[ListView8.ItemFocused.Index].SubItems[0]
          + ' - не удалось провести исключение');
    end
end;

/// //////////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.N7Click(Sender: TObject);
begin
  if ListView1.Items.Count <> 0 then
  /// / создать службу
    begin
      MyPS := ComboBox2.text;
      OKRightDlg.showmodal;
    end;
end;

procedure TfrmDomainInfo.N8Click(Sender: TObject);
/// /// тип запуска службы
begin
  if (ListView1.Items.Count <> 0) and (ListView1.SelCount <> 0) then
    begin
      SelectServic := ListView1.Selected.SubItems[0];
      ActionServic := 4;
    end;
end;

procedure TfrmDomainInfo.N90Click(Sender: TObject);
begin
/// // обновить данные для компов из бд
  if (ListView8.Items.Count <> 0) then
    readinfoforpcDB;
end;

/// ///////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.N91Click(Sender: TObject);
begin
/// / открываем форму создания сетевого ресурса
  FormShareFS.showmodal;
end;

procedure TfrmDomainInfo.N92Click(Sender: TObject);
var
  Names: string;
begin
  try
  /// удаление или отлючение сетевого ресурса
    if (ListViewShare.Items.Count <> 0) and (ListViewShare.SelCount = 1) then
      begin
        Names := ListViewShare.Items[ListViewShare.ItemIndex].SubItems[0];
        if DeleteShareFS(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text,
          Names) then
          LoadNetworkShare;
      end;
  except
    on E: Exception do
      frmDomainInfo.Memo1.Lines.Add
        (MyPS + 'Ошибка удаления сетевого ресурса : "' + E.Message + '"');
  end;
end;

procedure TfrmDomainInfo.N93Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(lvWorkStation);
end;

procedure TfrmDomainInfo.N94Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView1);
end;

procedure TfrmDomainInfo.N95Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView2);
end;

procedure TfrmDomainInfo.N96Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView3);
end;

procedure TfrmDomainInfo.N97Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView4);
end;

procedure TfrmDomainInfo.N98Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView6);
end;

procedure TfrmDomainInfo.N99Click(Sender: TObject);
begin
/// поиск по ListView
  findForListView(ListView9);
end;

procedure TfrmDomainInfo.GetAccessMask1Click(Sender: TObject);
/// / GetAccessMaskShareFS
var
  Names: String;
begin
  if (ListViewShare.Items.Count <> 0) and (ListViewShare.SelCount = 1) then
    begin
      Names := ListViewShare.Items[ListViewShare.ItemIndex].SubItems[0];
      GetAccessMaskShareFS(ComboBox2.text, LabeledEdit1.text,
        LabeledEdit2.text, Names);
    end;
end;

/// //////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.N9Click(Sender: TObject);
/// // удалить службу
var
  Myservice: TThread;
var
  buttonSelected: Integer;
begin
  if (ListView1.Items.Count <> 0) and (ListView1.SelCount <> 0) then
    begin
      buttonSelected := MessageDlg('Вы действительно хотите удалить службу' +
        #10#13 + ListView1.Selected.SubItems[0], mtConfirmation,
        [mbYes, mbNo], 0);
      if buttonSelected = mrYes then
        begin
          MyPS := ComboBox2.text;
          SelectServic := ListView1.Selected.SubItems[0];
          ActionServic := 5;
          Myservice := unit6.Service.Create(true);
          Myservice.FreeOnTerminate := true;
          /// / под вопросом самоуничтожение потока
          Myservice.Start;
        end;
      if buttonSelected = mrNo then
        exit;
    end;
end;

procedure TfrmDomainInfo.Office1Click(Sender: TObject);
begin // активация office
  ActivationOffice;
end;

function TfrmDomainInfo.readAntivirusStatus(NamePC: string;
  LV: TListView): bool;
var
/// / чтение списка антивирусных  продуктов
  IDSoft: TStringArray;
begin
  LV.Clear;
  LV.Items.BeginUpdate;
  try
    try
      Datam.FDReadSoftSelectPC.SQL.Clear;
      Datam.FDReadSoftSelectPC.SQL.text :=
        'SELECT * FROM ANTIVIRUSPRODUCT WHERE NAMEPC=''' + NamePC + '''';
      Datam.FDReadSoftSelectPC.Open;

      if Datam.FDReadSoftSelectPC.FieldByName('WIN_DEF').Value <> null then
        if Datam.FDReadSoftSelectPC.FieldByName('WIN_DEF').AsString <> '' then
          With LV.Items.Add do // Windows Defender
            begin
              // NAMEPC,WIN_DEF,WIN_DEF_STATUS,WIN_DEF_UPDATE,ANTIVIRUS,ANTIVIRUS_STATUS,ANTIVIRUS_UPDATE,FIREWALL,FIREWALL_STATUS,ANTISPYWARE,ANTISPYWARE_STATUS,ANTISPYWARE_UPDATE
              Caption := ('Антивирус ' + Datam.FDReadSoftSelectPC.FieldByName
                ('WIN_DEF').AsString);
              if Datam.FDReadSoftSelectPC.FieldByName('WIN_DEF_STATUS').Value <> null
              then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('WIN_DEF_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('WIN_DEF_UPDATE').Value <> null
              then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('WIN_DEF_UPDATE').AsString)
              else
                SubItems.Add('');
              ImageIndex := statusAntivirus(SubItems[0], SubItems[1]);
            end;

      if Datam.FDReadSoftSelectPC.FieldByName('ANTIVIRUS').Value <> null then
        if Datam.FDReadSoftSelectPC.FieldByName('ANTIVIRUS').AsString <> '' then
          With LV.Items.Add do // antivirus
            begin
              Caption := ('Антивирус ' + Datam.FDReadSoftSelectPC.FieldByName
                ('ANTIVIRUS').AsString);
              if Datam.FDReadSoftSelectPC.FieldByName('ANTIVIRUS_STATUS').Value
                <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('ANTIVIRUS_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('ANTIVIRUS_UPDATE').Value
                <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('ANTIVIRUS_UPDATE').AsString)
              else
                SubItems.Add('');
              ImageIndex := statusAntivirus(SubItems[0], SubItems[1]);
            end;

      if Datam.FDReadSoftSelectPC.FieldByName('ANTISPYWARE').Value <> null then
        if Datam.FDReadSoftSelectPC.FieldByName('ANTISPYWARE').AsString <> ''
        then
          With LV.Items.Add do // AntiSpyWare
            begin
              Caption := 'Антишпион ' +
                (Datam.FDReadSoftSelectPC.FieldByName('ANTISPYWARE').AsString);
              if Datam.FDReadSoftSelectPC.FieldByName('ANTISPYWARE_STATUS')
                .Value <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('ANTISPYWARE_STATUS').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('ANTISPYWARE_UPDATE')
                .Value <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('ANTISPYWARE_UPDATE').AsString)
              else
                SubItems.Add('');
              ImageIndex := statusAntivirus(SubItems[0], SubItems[1]);
            end;

      if Datam.FDReadSoftSelectPC.FieldByName('FIREWALL').Value <> null then
        if Datam.FDReadSoftSelectPC.FieldByName('FIREWALL').AsString <> '' then
          With LV.Items.Add do // FireWall
            begin
              Caption := 'Файрвол ' +
                (Datam.FDReadSoftSelectPC.FieldByName('FIREWALL').AsString);
              if Datam.FDReadSoftSelectPC.FieldByName('FIREWALL_STATUS').Value
                <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('FIREWALL_STATUS').AsString)
              else
                SubItems.Add('');
              SubItems.Add('');
              ImageIndex := statusAntivirus(SubItems[0], 'Ok');
            end;
      Datam.FDReadSoftSelectPC.Close;
    except
      on E: Exception do
        begin
          Datam.FDReadSoftSelectPC.Close;
          Memo1.Lines.Add('Ошибка : ' + E.Message);
        end;
    end;
  finally
    LV.Items.EndUpdate;
  end;
end;

function TfrmDomainInfo.readMicrosoftLic(NamePC: string; LV: TListView): bool;
var
/// / чтение списка лицензий для продуктов Microsoft
  IDSoft: TStringArray;
  i: Integer;
  OfficeProduct, OfficeStatus, OfficeKey, OfficeTypeLic, OfficeDes: TstringList;
begin
  LV.Clear;
  LV.Items.BeginUpdate;
  OfficeProduct := TstringList.Create;
  OfficeStatus := TstringList.Create;
  OfficeKey := TstringList.Create;
  OfficeTypeLic := TstringList.Create;
  OfficeDes := TstringList.Create;
  try
    try
      Datam.FDReadSoftSelectPC.SQL.Clear;
      Datam.FDReadSoftSelectPC.SQL.text :=
        'SELECT * FROM MICROSOFT_LIC WHERE NAMEPC=''' + NamePC + '''';
      Datam.FDReadSoftSelectPC.Open;
      if Datam.FDReadSoftSelectPC.FieldByName('WINPRODUCT').Value <> null then
        if Datam.FDReadSoftSelectPC.FieldByName('WINPRODUCT').AsString <> ''
        then
          With LV.Items.Add do
            begin
              // WINPRODUCT,STATWINLIC,KEYWIN,TYPE_LIC_WIN,DESCRIPT_LIC_WIN,SOFFICEPRODUCT,SSTATOFLIC,SKEYOFC,TYPE_LIC_OFFICE,DESCRIPT_LIC_OFFICE
              Caption := (Datam.FDReadSoftSelectPC.FieldByName('WINPRODUCT')
                .AsString);
              if Datam.FDReadSoftSelectPC.FieldByName('STATWINLIC').Value <> null
              then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('STATWINLIC').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('KEYWIN').Value <> null
              then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('KEYWIN').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('TYPE_LIC_WIN').Value <> null
              then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('TYPE_LIC_WIN').AsString)
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('DESCRIPT_LIC_WIN').Value
                <> null then
                SubItems.Add(Datam.FDReadSoftSelectPC.FieldByName
                  ('DESCRIPT_LIC_WIN').AsString)
              else
                SubItems.Add('');

              if Datam.FDReadSoftSelectPC.FieldByName('SOFFICEPRODUCT').Value <> null
              then
                begin
                  OfficeProduct.CommaText :=
                    Datam.FDReadSoftSelectPC.FieldByName
                    ('SOFFICEPRODUCT').AsString;
                  OfficeKey.CommaText := Datam.FDReadSoftSelectPC.FieldByName
                    ('SKEYOFC').AsString;
                  OfficeTypeLic.CommaText :=
                    Datam.FDReadSoftSelectPC.FieldByName
                    ('TYPE_LIC_OFFICE').AsString;
                  OfficeStatus.CommaText := Datam.FDReadSoftSelectPC.FieldByName
                    ('SSTATOFLIC').AsString;
                  OfficeDes.CommaText := Datam.FDReadSoftSelectPC.FieldByName
                    ('DESCRIPT_LIC_OFFICE').AsString;
                end;
            end;
      for i := 0 to OfficeProduct.Count - 1 do
        begin
          With LV.Items.Add do
            begin
              Caption := OfficeProduct[i];
              SubItems.Add(OfficeStatus[i]);
              SubItems.Add(OfficeKey[i]);
              SubItems.Add(OfficeTypeLic[i]);
              SubItems.Add(OfficeDes[i]);
            end;
        end;
      Datam.FDReadSoftSelectPC.Close;
    except
      on E: Exception do
        begin
          Memo1.Lines.Add('Ошибка : ' + E.Message);
        end;
    end;
  finally
    LV.Items.EndUpdate;
    OfficeProduct.Free;
    OfficeStatus.Free;
    OfficeKey.Free;
    OfficeTypeLic.Free;
    OfficeDes.Free;
  end;
end;

function TfrmDomainInfo.readSoftRorSelectPC(s: string; LV: TListView): bool;
var
/// / чтение списка программ для выбранного компа на вкладке компы
  IDSoft: TStringArray;
  i: Integer;
begin
  LV.Clear;
  LV.Items.BeginUpdate;
  try
    try
      Datam.FDReadSoftSelectPC.SQL.Clear;
      Datam.FDReadSoftSelectPC.SQL.text :=
        'SELECT * FROM MAIN_PC_SOFT WHERE PC_NAME=''' + s + '''';
      Datam.FDReadSoftSelectPC.Open;
      IDSoft := kolitemSoft
        (vartostr(Datam.FDReadSoftSelectPC.FieldByName('INST_SOFT').Value));
      for i := 0 to Length(IDSoft) - 1 do
        begin
          // memo1.Lines.add(IDSoft[i]); DataM.FDQueryReadSoft.SQL.Text:='SELECT * FROM MAIN_PC_SOFT ORDER BY PC_NAME';
          Datam.FDReadSoftSelectPC.SQL.Clear;
          Datam.FDReadSoftSelectPC.SQL.text :=
            'SELECT * FROM SOFT_PC WHERE ID=''' + IDSoft[i] + '''';
          Datam.FDReadSoftSelectPC.Open;
          With LV.Items.Add do
            begin
              Caption := inttostr(LV.Items.Count);
              if Datam.FDReadSoftSelectPC.FieldByName('SOFT_NAME').Value <> null
              then
                SubItems.Add
                  (string(Datam.FDReadSoftSelectPC.FieldByName
                  ('SOFT_NAME').Value))
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('MANUFACTURE').Value <> null
              then
                SubItems.Add
                  (string(Datam.FDReadSoftSelectPC.FieldByName
                  ('MANUFACTURE').Value))
              else
                SubItems.Add('');
              if Datam.FDReadSoftSelectPC.FieldByName('SOFT_VERSION').Value <> null
              then
                SubItems.Add
                  (string(Datam.FDReadSoftSelectPC.FieldByName
                  ('SOFT_VERSION').Value))
              else
                SubItems.Add('');
              ImageIndex := BrendProg(LV.Items[i].SubItems[1]);
            end;
        end;

      Datam.FDReadSoftSelectPC.Close;
    except
      on E: Exception do
        begin
          Datam.FDReadSoftSelectPC.Close;
          // Memo1.Lines.Add('Ошибка : '+e.Message);
        end;
    end;
  finally
    LV.Items.EndUpdate;
  end;
end;

/// ///////////////////////////////////////////////////////
function WbemTimeToDateTime(const V: OleVariant): string; // перевод даты
var
  Dt: string;
  FormatSettings, FormatSettings1: TFormatSettings;
begin
  // Result:=0;
  try
    if VarIsNull(V) then
      exit;
    Dt := string(V);
    Insert('/', Dt, 5);
    Insert('/', Dt, 8);
    FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
    FormatSettings.DateSeparator := '/';
    FormatSettings1.ShortDateFormat := 'dd.mm.yyyy';
    FormatSettings1.DateSeparator := ':';
    Result := datetostr(StrToDate(Dt, FormatSettings), FormatSettings1)
  except
    Result := '01.01.2001';
  end;
  // result:=dt;
end;

function WbemTimeToDateAndTime(const V: OleVariant): string;
// перевод даты и времени
var
  Dt: string;
begin
  if VarIsNull(V) then
    exit;
  Dt := string(V);
  Insert('/', Dt, 5);
  Insert('/', Dt, 8);
  Insert('-', Dt, 11);
  Insert(':', Dt, 14);
  Insert(':', Dt, 17);
  Delete(Dt, 20, Length(Dt) - 19);
  Result := Dt;
end;

function DateConvert(s: string): string;
/// / конвертируем дату  из строки
var
/// // с заменой делителя и сменой формата
  FormatSettings, FormatSettings1: TFormatSettings;
begin
  FormatSettings.ShortDateFormat := 'mm-dd-yyyy';
  FormatSettings.DateSeparator := '/';
  FormatSettings1.ShortDateFormat := 'dd.mm.yyyy';
  FormatSettings1.DateSeparator := ':';
  Result := datetostr(StrToDate(s, FormatSettings), FormatSettings1);

end;

/// //////////////////////////////////////////////////////////
function ArrayToString(const Data: array of string): string;
var
  SL: TstringList;
  s: string;
begin
  SL := TstringList.Create;
  try
    for s in Data do
      SL.Add(s);
    Result := SL.text;
  finally
    SL.Free;
  end;
end;

/// ////////////////////////////// получем информацию с машин через WMI
function TfrmDomainInfo.scanInfo(nextPC, UserScan, PassScan: string): Integer;
var
  MyMem, MysystemType: string;
  FSWbemLocator: OleVariant;
  FWMIService: OleVariant;
  FWbemObjectSet: OleVariant;
  FWbemObject: OleVariant;
  oEnum: IEnumvariant;
  NewTread, NewTreadProc: Integer;
  OSName, OSSP: string;
  iValue, treadID, treadIDProc: LongWord;
  ///
begin
  CpuLogPr := 1;
  CpuNum := 1;
  begin
    try
      begin
        OleInitialize(nil);
        /// ////////////////////////////////////////////////////////////////////////////////
        if pcRes.ActivePageIndex = 13 then
          begin
            LoadNetworkShare;
            /// Загрузка списка сетевых ресурсов
            exit;
          end;
        /// //////////////////////////////////////////////////////////////////////////////
        FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
        FWMIService := FSWbemLocator.ConnectServer(nextPC, 'root\CIMV2',
          UserScan, PassScan, '', '', 128);
        try
          FWbemObjectSet := FWMIService.ExecQuery
            ('SELECT OSArchitecture,Caption,CSDVersion,Version FROM Win32_OperatingSystem',
            'WQL', wbemFlagForwardOnly);
          oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
          while oEnum.Next(1, FWbemObject, iValue) = 0 do
          /// / Определяем архитектуру
            begin
              if FWbemObject.OSArchitecture <> null then
                MysystemType := string(FWbemObject.OSArchitecture);
              if FWbemObject.Caption <> null then
                OSName := string(FWbemObject.Caption) + '/';
              if FWbemObject.CSDVersion <> null then
                OSSP := string(FWbemObject.CSDVersion) + '/';
              if FWbemObject.Version <> null then
                OSVersion := string(FWbemObject.Version) + '/';
              Label3.Caption := OSName + ' ' + OSSP + ' ' + OSVersion + ' ' +
                MysystemType;
              VariantClear(FWbemObject);
            end;
          oEnum := nil;
          VariantClear(FWbemObjectSet);
        except
          begin
            MysystemType := '32 bit';
          end;
        end;

        /// //////////////////////////////////////////////////////////////////////////
        try
          FWbemObjectSet := FWMIService.ExecQuery
            ('SELECT NumberOfLogicalProcessors,TotalPhysicalMemory FROM Win32_ComputerSystem',
            'WQL', wbemFlagForwardOnly);
          oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
          while oEnum.Next(1, FWbemObject, iValue) = 0 do
          /// / Определяем количество процессоров и кол-во ядер
            begin
              if FWbemObject.NumberOfLogicalProcessors <> null then
                CpuLogPr := (FWbemObject.NumberOfLogicalProcessors);
              if FWbemObject.TotalPhysicalMemory <> null then
                begin
                  MyMem := vartostr(FWbemObject.TotalPhysicalMemory);
                  MyMemory := strtofloat(MyMem);
                  MyMemory := (Trunc(MyMemory / 1024 / 1024));
                  mem := strtoint(floattostr(MyMemory));
                  VariantClear(FWbemObject);
                end;
            end;
          oEnum := nil;
          VariantClear(FWbemObjectSet);
        except
          begin
            mem := 4096;
          end;
        end;

        /// //////////////////////////////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 0 then
          begin
            if TabSheet2.Tag = 1 then
              begin
                showmessage('Дождитесь предыдущей загрузки списка процессов');
                exit;
              end;
            Timer1.Enabled := false;
            lvWorkStation.Clear;
            try
              CallListProcess.NamePC := nextPC;
              CallListProcess.user := UserScan;
              CallListProcess.pass := PassScan;
              CallListProcess.ListViewProc := lvWorkStation;
              CallListProcess.TimerScan := Timer1;
              TabSheet2.Tag := 1;
              /// /// признак загрузки списка процессов
              NewTreadProc := BeginThread(nil, 0, addr(LoadListProcess),
                addr(CallListProcess), 0, treadIDProc);
              /// поток для заполнения списка событий
              CloseHandle(NewTreadProc);
            except
              on E: Exception do
                begin
                  TabSheet2.Tag := 0;
                  Memo1.Lines.Add('Ошибка загрузки списка процессов "' +
                    E.Message + '"');
                end;
            end;

          end;

        /// //////////////////////////////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 1 then
        /// / загрузка списка служб
          begin
            if TabSheet3.Tag = 1 then
              begin
                showmessage('Дождитесь предыдущей загрузки списка служб');
                exit;
              end;
            try
              CallListServise.NamePC := nextPC;
              CallListServise.user := UserScan;
              CallListServise.pass := PassScan;
              CallListServise.ListViewMsi := ListView1;
              TabSheet3.Tag := 1;
              /// /// признак загрузки списка
              NewTread := BeginThread(nil, 0, addr(LoadListServicesThread),
                addr(CallListServise), 0, treadID);
              /// поток для заполнения списка событий
              CloseHandle(NewTread);
            except
              on E: Exception do
                begin
                  TabSheet3.Tag := 0;
                  Memo1.Lines.Add('Ошибка загрузки списка служб "' +
                    E.Message + '"');
                end;
            end;
            // LoadListServices(nextPC,UserScan,PassScan,Listview1);
          end;
        /// /////////////////////////////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 2 then
        /// загрузка списка автозарузки
          begin
            LoadListAutoStart(nextPC, UserScan, PassScan, MysystemType,
              ListView2);
          end;
        /// /////////////////////////////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 4 then
        /// / загрузка списка всех программ
          begin
            if TabSheet5.Tag = 1 then
              begin
                showmessage('Дождитесь предыдущей загрузки списка программ');
                exit;
              end;
            try
              CallListAllprogram.NamePC := nextPC;
              CallListAllprogram.user := UserScan;
              CallListAllprogram.pass := PassScan;
              CallListAllprogram.ListViewMsi := ListView4;
              TabSheet5.Tag := 1;
              /// /// признак загрузки списка
              NewTread := BeginThread(nil, 0, addr(LoadListAllProgram),
                addr(CallListAllprogram), 0, treadID);
              /// поток для заполнения списка событий
              CloseHandle(NewTread);
            except
              on E: Exception do
                begin
                  TabSheet5.Tag := 0;
                  Memo1.Lines.Add('Ошибка загрузки списка всех программ "' +
                    E.Message + '"');
                end;
            end;
          end;
        /// ////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 7 then
        /// загрузка списка профилей
          begin
            LoadListProfile(nextPC, UserScan, PassScan, ListView5);
          end;
        /// //////////////////////////////////////////////////////////
        if frmDomainInfo.pcRes.ActivePageIndex = 8 then
        /// устройства PnP
          begin
            Memo1.Lines.Add('Загрузка списка устройств PnP');
            LoadDevicePnP(nextPC, UserScan, PassScan, ListView6);
            Memo1.Lines.Add('Загрузка списка устройств PnP завершена.');
          end;
        /// ////////////////////////////////////////////////////////////
        if pcRes.ActivePageIndex = 9 then
        /// загрузка сетевых интерфейсов
          begin
            LoadListSNetworkInterface(nextPC, UserScan, PassScan, ListView7);
          end;
        /// /////////////////////////////////////////////////////////
        if pcRes.ActivePageIndex = 12 then
          begin
            if TabSheet13.Tag = 1 then
              begin
                showmessage('Дождитесь предыдущей загрузки списка обновлений');
                exit;
              end;
            ListView9.Clear;
            Memo1.Lines.Add
              ('Загрузка списка установленных обновлений/исправлений Windows');
            try
              CallListUpdate.NamePC := nextPC;
              CallListUpdate.user := UserScan;
              CallListUpdate.pass := PassScan;
              CallListUpdate.ListViewMsi := ListView9;
              TabSheet13.Tag := 1;
              /// /// признак загрузки списка обновлений
              NewTread := BeginThread(nil, 0, addr(LoadListWindowsUpdate),
                addr(CallListUpdate), 0, treadID);
              /// поток для заполнения списка событий
              CloseHandle(NewTread);
            except
              on E: Exception do
                begin
                  TabSheet13.Tag := 0;
                  Memo1.Lines.Add('Ошибка загрузки списка обновлений Windows "'
                    + E.Message + '"');
                end;
            end;

          end;

      end;
    except
      on E: Exception do
        begin
          VariantClear(FWbemObjectSet);
          VariantClear(FSWbemLocator);
          VariantClear(FWMIService);
          FWbemObject := Unassigned;
          Memo1.Lines.Add('Общая ошибка  "' + E.Message + '"');
          exit;
        end;
    end;
  end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
  OleUnInitialize;
end;

/// ///////////////////////////////////////////////////////////

procedure TfrmDomainInfo.SpeedButton2Click(Sender: TObject);
/// / сканируем
begin
  if ping(ComboBox2.text) then
    scanInfo(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text);
end;

/// /////////////////////////////////////////////////
function TfrmDomainInfo.BrendProg(Name: String): Integer;
var
  i, Res: Integer;
const
  Brend: array [0 .. 40] of string = ('Microsoft Corporation',
    'Adobe Systems Incorporated.',
    'ATI Advanced Micro Devices, Inc. Advanced Micro Devices Inc. Advanced Micro Devices, Inc Advanced Micro Devices Inc AMD ATI ATI Technologies, Inc.',
    'Piriform', 'Cisco', 'Mozilla', 'Google Google Inc. Google LLC',
    'Hewlett Packard HP Hewlett-Packard Co. Hewlett-Packard Enterprise',
    'Kyocera Corporation', 'Nero AG', 'Opera Software ASA',
    'Samsung Electronics Co., Ltd. Samsung Electronics Co. Ltd. Samsung Electronics Co Ltd ',
    'Skype Technologies S.A.', 'TeamViewer', 'VideoLAN Team', 'Visioneer Inc.',
    'Корпорация Майкрософт Microsoft Corporation Microsoft Corporations',
    'ЗАО "Лаборатория Касперского" AVP Касперский Антивирус «Лаборатория Касперского» АО "Лаборатория Касперского"',
    'ООО «Доктор Веб» Dr Web', 'McAfee, Inc. McAfee Inc', 'Mail.Ru Kometa',
    'Apple Inc.', 'AVAST Software', 'Intel Corporation', '1С 1C 1С-Софт',
    'ООО "ДубльГИС" 2Gis 2ГИС', 'Фаматек RAdmin', 'ABBYY', 'Алексей Скрябин',
    'skrblog.ru', 'Yandex Яндекс', 'NVIDIA NVIDIA Corporation', 'win.rar GmbH',
    'D-Link', 'Logitech Logitech Inc.', 'Atheros Atheros Communications Inc.',
    'Qualcomm', 'Autodesk Autodesk, Inc.', 'Corel Corporation',
    'Realtek Semiconductor Corp.', 'other');
begin
  for i := 0 to 40 do
    begin
      Res := pos(AnsiLowerCase(trim(Name)), AnsiLowerCase(Brend[i]));
      if (Res <> 0) then
        begin
          Result := i;
          break;
        end
      else
        Result := 40;
    end;
end;
/// ////////////////////////////////////////////////////

procedure TfrmDomainInfo.SpeedButton3Click(Sender: TObject);
/// список программ msi
var
  NewTread: Integer;
  treadID: LongWord;
begin
  ListView3.Clear;
  MyPS := ComboBox2.text;
  if ping(MyPS) then
    begin
      Timer1.Enabled := false;
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      Application.ProcessMessages;
      Memo1.Lines.Add('Загрузка списка установленных программ (msi)...');
      try
        CallListMSI.NamePC := MyPS;
        CallListMSI.user := MyUser;
        CallListMSI.pass := MyPasswd;
        CallListMSI.ListViewMsi := ListView3;
        SpeedButton3.Enabled := false;
        NewTread := BeginThread(nil, 0, addr(LoadListProgpamMSI),
          addr(CallListMSI), 0, treadID);
        /// поток для заполнения списка событий
        CloseHandle(NewTread);
      except
        on E: Exception do
          begin
            Memo1.Lines.Add('Программы - "' + E.Message + '"');
            SpeedButton3.Enabled := true;
          end;
      end;
    end;
end;

procedure TfrmDomainInfo.SpeedButton4Click(Sender: TObject);
/// загрузка списка дисков
begin
  ListViewDisk.Clear;
  if ping(ComboBox2.text) = true then
    LoadListLocalDisk(ComboBox2.text, LabeledEdit1.text, LabeledEdit2.text,
      ListViewDisk);

end;

procedure TfrmDomainInfo.RunProcessForListPC;
/// /// групповой запуск процесса
begin
  if ListView8.Items.Count = 0 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  GroupPC := true;
  NewProcForm.showmodal;
end;

procedure TfrmDomainInfo.KillProcessForListPC;
/// /// групповое завершение процесса
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  if ListView8.Items.Count > 0 then
    begin
      SelectedPCkillProc := TstringList.Create;
      SelectedPCkillProc := createListpcForCheck('');
      /// / создание списка компьютеров для завершения процесса
      if SelectedPCkillProc.Count <> 0 then
        begin
          GroupPC := true;
          OKRightDlg123456789101112131415161718192021.showmodal;
        end
      else
        showmessage('Не выбран список компьютеров!');
    end;

end;

procedure TfrmDomainInfo.AddDriverPrintForListPC;
/// /// групповая установка драйверов принтера
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  if ListView8.Items.Count > 0 then
    begin
      GroupPC := true;
      OKRightDlg12345678910.showmodal;
    end;
end;

procedure TfrmDomainInfo.AddPrinterForListPC;
/// //// групповое добавление принтеров
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;

  showmessage('Перед установкой принтера, убедитесь что на выбранных' + #10#13 +
    'компьютерах установлены драйвера для устанавливаемого принтера!!!');
  if ListView8.Items.Count > 0 then
    begin
      SelectedPC := TstringList.Create;
      SelectedPC := createListpcForCheck('');
      if SelectedPC.Count <> 0 then
        begin
          GroupPC := true;
          OKRightDlg123456789.showmodal;
        end
      else
        showmessage('Не выбран список компьютеров!');
    end;
end;

procedure TfrmDomainInfo.ADIP1Click(Sender: TObject);
// открыть форму для заполнения спсиска ПК
begin
  if not killscancomp  then
    begin
      Memo1.Lines.Add
        ('Не могу остановить сканирование, повторите чуть позже...');
      exit;
    end;
  FormOneStart.Show;
end;

/// ///////////////////////////////////////////////////////////////////////////WOL
procedure TfrmDomainInfo.PowerForListPC;
var
  i, z: Integer;
  setini: TMemInifile;
  BRaddress: string;
  WOLPower: TThread;
begin
  try
    if (ListView8.Items.Count < 1) then
      begin
        showmessage('В списке нет компьютеров');
        exit;
      end;

    ListMACAddress := TstringList.Create;
    for i := 0 to ListView8.Items.Count - 1 do
      begin
        if (ListView8.Items[i].Checked) and
          (ListView8.Items[i].SubItems[2] <> '') then
          begin
            ListMACAddress.Add(ListView8.Items[i].SubItems[2]);
          end;
      end;

    if ListMACAddress.Count = 0 then
      begin
        showmessage('У выбранных компьютеров не определены MAC адреса');
        ListMACAddress.Free;
        exit;
      end;

    setini := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
      '\Settings.ini', TEncoding.Unicode);
    BRaddress := setini.readString('ConfLAN', 'broadcast', '192.168.0.255');
    IpBroadCast := InputBox('Broadcast адрес сети', 'IP-', BRaddress);
    if IpBroadCast = '' then
      begin
        showmessage('Вы не ввели broadcast адрес сети!');
        ListMACAddress.Free;
        exit;
      end;

    z := MessageDlg('Включить ' + inttostr(ListMACAddress.Count) + '-шт ПК? ',
      mtWarning, [mbYes, mbCancel], 0);
    if z = mrYes then
      begin
        WOLPower := WOLThread.WOL.Create(true);
        WOLPower.FreeOnTerminate := true;
        /// / под вопросом самоуничтожение потока
        WOLPower.Start;
        setini.WriteString('ConfLAN', 'broadcast', IpBroadCast);
        setini.UpdateFile;
      end
    else
      ListMACAddress.Free;
    if Assigned(setini) then
      FreeAndNil(setini);
  except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка WOL - "' + E.Message + '"');
        exit;
      end;
  end;
end;

procedure TfrmDomainInfo.ResetCurPC;
var
  MyShutdownPC: TThread;
begin
/// / перезагрузка
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 2;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.WakeOnLan2Click(Sender: TObject);
// включить группу компьютеров
begin
  PowerForListPC;
  // SpeedButton54.Click;
end;

procedure TfrmDomainInfo.WakeOnLan1Click(Sender: TObject);
begin // WOL
  PowerOn;
end;

/// /////////////////////////////////////////////////////////////////////////////////////

procedure TfrmDomainInfo.PowerOffForListPC;
var
  i: Integer;
  /// /// групповое завершение работы
  NewSelectedPCShotDown: TThread;
begin
  if (ListView8.Items.Count < 1) then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  i := MessageDlg('Завершить работу на выбранных компьютерах? ', mtWarning,
    [mbYes, mbCancel], 0);
  if i = mrCancel then
    exit;
  begin
    if ListView8.Items.Count > 0 then
      begin
        SelectedPCShutDown := TstringList.Create;
        SelectedPCShutDown := createListpcForCheck('');
        if SelectedPCShutDown.Count <> 0 then
          begin
            myShutdown := 4;
            NewSelectedPCShotDown :=
              SelectedPCShotDownThread.SelectedPCShotDown.Create(true);
            NewSelectedPCShotDown.FreeOnTerminate := true;
            NewSelectedPCShotDown.Start;
          end
        else
          showmessage('Не выбран не один компьютер!!!');
      end;
  end;
end;

procedure TfrmDomainInfo.CloseSession;
var
  MyShutdownPC: TThread;
begin
  MyPS := ComboBox2.text;
  if ping(MyPS) = true then
    begin
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      myShutdown := 0;
      MyShutdownPC := PCShutdown.Shutdown.Create(true);
      MyShutdownPC.FreeOnTerminate := true;
      /// / под вопросом самоуничтожение потока
      MyShutdownPC.Start;
    end;
end;

procedure TfrmDomainInfo.SpeedButton50Click(Sender: TObject);
begin
  if LVMicrosoft.Items.Count <> 0 then
    popupListViewSaveAs(LVMicrosoft,
      'Сохранение списка продуктов Windows и Office',
      'Лицензии Windows и Office');
end;

procedure TfrmDomainInfo.SpeedButton51Click(Sender: TObject);
begin
  LVMicrosoft.SelectAll;
end;

procedure TfrmDomainInfo.SpeedButton52Click(Sender: TObject);
var
  i: Integer;
  p: TPoint;
begin
  for i := 0 to LVMicrosoft.Items.Count - 1 do
    begin
      if pos(Ansiuppercase(LabeledEdit9.text),
        Ansiuppercase(LVMicrosoft.Items[i].SubItems[0])) <> 0 then
        begin
          LVMicrosoft.SetFocus;
          LVMicrosoft.Items[i].Selected := true;
          LVMicrosoft.ItemIndex := i;
          LVMicrosoft.ItemFocused := LVMicrosoft.Items[i];
          p := LVMicrosoft.Items.Item[i].Position;
          LVMicrosoft.Scroll(p.X, p.Y);
          // LVMicrosoft.SetFocus;
          break;
        end;
    end;
  if LVMicrosoft.ItemFocused <> nil then
    LVMicrosoft.ItemFocused.MakeVisible(false);
end;

procedure TfrmDomainInfo.SpeedButton53Click(Sender: TObject);
begin
/// / список всех продуктов Windows и Office
  try
    DBGrid4.DataSource.DataSet.DisableControls;
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes;
    /// очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    // очистка перед созданием полей в DBgrid4
    DBgrid4Columns('', 'Имя', 'WIN_OFF');
    /// создание колонок в DBgrid4
    Datam.FDQuerySelectSort.SQL.text :=
      'SELECT * FROM MICROSOFT_PRODUCT ORDER BY DESCRIPTION';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Read Windows Office - ' + E.Message)
      end;
  end;
  DBGrid4.DataSource.DataSet.EnableControls;
end;

procedure TfrmDomainInfo.SpeedButton54Click(Sender: TObject);
var
  i: Integer;
  b: Boolean;
begin
/// / обновление списка продуктов Windows Office
  DBGrid4.DataSource.DataSet.DisableControls;
  try
    Datam.FDQuerySort.Transaction.StartTransaction;
    Datam.FDQuerySort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySort.SQL.Clear;
    Datam.FDQuerySort.SQL.text :=
      'SELECT * FROM MICROSOFT_PRODUCT ORDER BY DESCRIPTION';
    Datam.FDQuerySort.Open;
    Datam.FDQuerySort.First;
    ComboBox9.Clear;
    while not Datam.FDQuerySort.EOF do
      begin
      /// заполняем  ComboBox9 списком Продуктов
        if Datam.FDQuerySort.FieldByName('DESCRIPTION').Value <> null then
          begin
            for i := 0 to ComboBox9.Items.Count - 1 do
              if ComboBox9.Items[i] <> Datam.FDQuerySort.FieldByName
                ('DESCRIPTION').AsString then
                b := false
              else
                begin
                  b := true;
                  break;
                end;
            if not b then
              ComboBox9.Items.Add
                (string(Datam.FDQuerySort.FieldByName('DESCRIPTION').AsString));
          end;
        Datam.FDQuerySort.Next;
      end;
    if ComboBox9.Items.Count > 0 then
      ComboBox9.ItemIndex := 0;
    Datam.FDQuerySort.SQL.Clear;
    Datam.FDQuerySort.Close;
    Datam.FDQuerySort.Transaction.commit;
    /// закрываем транзакцию
  finally
    DBGrid4.DataSource.DataSet.EnableControls;
  end;
end;

procedure TfrmDomainInfo.SpeedButton55Click(Sender: TObject);
begin
  UpdateAntivirusProduct;
end;

procedure TfrmDomainInfo.SpeedButton56Click(Sender: TObject);
var
  i: Integer;
  p: TPoint;
begin
  for i := 0 to LVAntivirus.Items.Count - 1 do
    begin
      if pos(Ansiuppercase(LabeledEdit10.text),
        Ansiuppercase(LVAntivirus.Items[i].Caption)) <> 0 then
        begin
          LVAntivirus.SetFocus;
          LVAntivirus.Items[i].Selected := true;
          LVAntivirus.ItemIndex := i;
          LVAntivirus.ItemFocused := LVAntivirus.Items[i];
          p := LVAntivirus.Items.Item[i].Position;
          LVAntivirus.Scroll(p.X, p.Y);
          break;
        end;
    end;
  if LVAntivirus.ItemFocused <> nil then
    LVAntivirus.ItemFocused.MakeVisible(false);
end;

procedure TfrmDomainInfo.SpeedButton59Click(Sender: TObject);
var
  i: Integer;
  FDQueryAnt: TFDQuery;
  Transaction: TFDTransaction;
begin
  try
    if LVAntivirus.SelCount = 0 then
      exit;
    Transaction := TFDTransaction.Create(nil);
    Transaction.Connection := Datam.ConnectionDB;
    Transaction.Options.Isolation := xiReadCommitted;
    /// xiReadCommitted; //xiSnapshot; xiUnspecified
    FDQueryAnt := TFDQuery.Create(nil);
    FDQueryAnt.Transaction := Transaction;
    FDQueryAnt.Connection := Datam.ConnectionDB;
    Transaction.StartTransaction; // стартуем
    LVAntivirus.Items.BeginUpdate;
    try
      for i := LVAntivirus.Items.Count - 1 downto 0 do
        begin
          if LVAntivirus.Items[i].Selected then
            begin
              FDQueryAnt.SQL.Clear;
              FDQueryAnt.SQL.text :=
                'delete FROM ANTIVIRUSPRODUCT WHERE NAMEPC=''' +
                LVAntivirus.Items[i].Caption + '''';
              FDQueryAnt.ExecSQL;
              LVAntivirus.Items[i].Delete;
            end;
        end;
    finally
      Transaction.commit;
      LVAntivirus.Items.EndUpdate;
      FDQueryAnt.Free;
      Transaction.Free;
    end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка удаления ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.SpeedButton5Click(Sender: TObject);
/// // загрузка списка принтеров
var
  i: byte;
  CaptionPrint: string;
  FSWbemLocator: OleVariant;
  FWMIService: OleVariant;
  FWbemObjectSet: OleVariant;
  oEnum: IEnumvariant;
  FWbemObject: OleVariant;
  iValue: LongWord;
begin
  MyPS := ComboBox2.text;
  MyUser := LabeledEdit1.text;
  MyPasswd := LabeledEdit2.text;
  if ping(MyPS) = true then
    try
      begin
        ListViewPrint.Clear;
        begin
          frmDomainInfo.Memo1.Lines.Add('Загруска списка принтеров...');
          OleInitialize(nil);
          FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
          FWMIService := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser,
            MyPasswd);
          FWMIService.Security_.impersonationlevel := 3;
          FWMIService.Security_.authenticationLevel := 6;
          FWbemObjectSet := FWMIService.ExecQuery('SELECT * FROM Win32_Printer',
            'WQL', wbemFlagForwardOnly);
          oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumvariant;
          while oEnum.Next(1, FWbemObject, iValue) = 0 do
          /// / Определяем архитектуру
            begin
              if (FWbemObject.DeviceID <> null) then
                begin
                  CaptionPrint := string(FWbemObject.DeviceID);
                  /// FWbemObject.caption
                  // if FWbemObject.Local=True then CaptionPrint:=CaptionPrint+'\Local' else CaptionPrint:=CaptionPrint+'\Network';
                  // if boolean(FWbemObject.Shared)  then CaptionPrint:=CaptionPrint+'\SharedNetwork';
                  begin
                    with ListViewPrint.Items.Add do
                      begin
                        Caption := CaptionPrint;
                        if not VarIsNull(FWbemObject.Shared) then
                          begin
                            if Boolean(FWbemObject.Shared) then
                              ImageIndex := 1
                            else
                              ImageIndex := 0;
                          end
                        else
                          ImageIndex := 0;
                      end;
                  end;
                end;
              VariantClear(FWbemObject);
              Inc(i);
            end;
          oEnum := nil;
          OleUnInitialize;
          frmDomainInfo.Memo1.Lines.Add('Загрузка списка принтеров завершена.');
        end;
      end;
    except
      on E: Exception do
        begin
          Memo1.Lines.Add('Ошибка загрузки списка принтеров - "' +
            E.Message + '"');
          exit;
        end;
    end;
  VariantClear(FWbemObject);
  oEnum := nil;
  VariantClear(FWbemObjectSet);
  VariantClear(FWMIService);
  VariantClear(FSWbemLocator);
end;

procedure TfrmDomainInfo.ListView9ColumnClick(Sender: TObject;
  /// сортировка Обновлений Windows
  Column: TListColumn);
begin
  sort := -sort;
  SortLV8 := not SortLV8;
  if Column = ListView9.Columns[5] then
    ListView9.CustomSort(@SortThirdSubItemAsDate, sort)
  else
    ListView9.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView9DblClick(Sender: TObject);
begin
  if ListView9.Items.Count <> 0 then
    ShellExecute(Handle, 'open',
      PChar(ListView9.Items[ListView9.Selected.Index].SubItems[2]), nil,
      nil, SW_SHOW);
end;

procedure TfrmDomainInfo.ListView9MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListViewSoftColumnClick(Sender: TObject;
  Column: TListColumn);
/// / сортировка списка программ
begin
  SortLV8 := not SortLV8;
  ListViewSoft.CustomSort(@CustomSortProc, Column.Index);

end;

procedure TfrmDomainInfo.ListViewSoftinPCColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  SortLV8 := not SortLV8;
  ListViewSoftinPC.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListViewSoftinPCDblClick(Sender: TObject);
var
  liststr: TstringList;
begin
  if ListViewSoftinPC.SelCount = 0 then
    exit;
  liststr := TstringList.Create;
  try
    liststr.Add(ListViewSoftinPC.Selected.SubItems[0]);
    liststr.Add(ListViewSoftinPC.Selected.SubItems[1]);
    liststr.Add(ListViewSoftinPC.Selected.SubItems[2]);
    itemLisToMemo(liststr, GroupBox2.Caption, frmDomainInfo);
  finally
    liststr.Free;
  end;

end;

procedure TfrmDomainInfo.ListView4ColumnClick(Sender: TObject;
  Column: TListColumn);
/// // сортировка всех программ
begin
  sort := -sort;
  SortLV8 := not SortLV8;
  if Column = ListView4.Columns[4] then
    ListView4.CustomSort(@SortThirdSubItemAsDate3, sort)
  else
    ListView4.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView4MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListView3ColumnClick(Sender: TObject;
  Column: TListColumn);
/// сортировка программ MSI
begin
  sort := -sort;
  SortLV8 := not SortLV8;
  if Column = ListView3.Columns[4] then
    ListView3.CustomSort(@SortThirdSubItemAsDate3, sort)
  else
    ListView3.CustomSort(@CustomSortProc, Column.Index);
end;

procedure TfrmDomainInfo.ListView3MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  ListViewMousPoint.X := X;
  ListViewMousPoint.Y := Y;
end;

procedure TfrmDomainInfo.ListView1ColumnClick(Sender: TObject;
  Column: TListColumn);
/// / сортировка служб
begin
  SortLV8 := not SortLV8;
  ListView1.CustomSort(@CustomSortProc, Column.Index);
end;

/// /////////////////////////////////////////////////////////////// RD

procedure TfrmDomainInfo.ScrollBox2MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ScrollBox2.SetFocus;
end;

procedure TfrmDomainInfo.ScrollBox2MouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  ScrollBox2.VertScrollBar.Position := ScrollBox2.VertScrollBar.Position + 10;
end;

procedure TfrmDomainInfo.ScrollBox2MouseWheelUp(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  ScrollBox2.VertScrollBar.Position := ScrollBox2.VertScrollBar.Position - 10;
end;

procedure TfrmDomainInfo.SMARTHDD1Click(Sender: TObject);
begin
/// / форма просмотра статуса SMART HDD
  OpenFormSMATR;
end;

procedure TfrmDomainInfo.SpeedButton1Click(Sender: TObject);
/// // подключение/отключение RDP
begin
  try
    begin
      if SpeedButton1.Caption = 'Подключится' then
        begin
          if (ComboBox2.text <> '') and (ping(ComboBox2.text)) then
            begin
              try
                if TabSheet9.FindComponent('MyRDPClient') = nil then
                  begin
                    MyRDPClient := TMsRdpClient9.Create(TabSheet9);
                    MyRDPClient.Parent := TabSheet9;
                    MyRDPClient.Name := 'MyRDPClient';
                    MyRDPClient.Align := alClient;
                    MyRDPClient.OnAuthenticationWarningDismissed :=
                      MsRdpClient91AuthenticationWarningDismissed;
                    MyRDPClient.OnAuthenticationWarningDisplayed :=
                      MsRdpClient91AuthenticationWarningDisplayed;
                    MyRDPClient.OnAutoReconnected :=
                      MsRdpClient91AutoReconnected;
                    MyRDPClient.OnAutoReconnecting :=
                      MsRdpClient91AutoReconnecting;
                    MyRDPClient.OnAutoReconnecting2 :=
                      MsRdpClient91AutoReconnecting2;
                    MyRDPClient.OnChannelReceivedData :=
                      MsRdpClient91ChannelReceivedData;
                    MyRDPClient.OnConfirmClose := MsRdpClient91ConfirmClose;
                    MyRDPClient.OnConnected := MsRdpClient91Connected;
                    MyRDPClient.OnConnecting := MsRdpClient91Connecting;
                    MyRDPClient.OnDisconnected := MsRdpClient91Disconnected;
                    MyRDPClient.OnExit := MsRdpClient91Exit;
                    MyRDPClient.OnFatalError := MsRdpClient91FatalError;
                    MyRDPClient.OnLoginComplete := MsRdpClient91LoginComplete;
                    MyRDPClient.OnLogonError := MsRdpClient91LogonError;
                    MyRDPClient.OnNetworkStatusChanged :=
                      MsRdpClient91NetworkStatusChanged;
                    MyRDPClient.OnWarning := MsRdpClient91Warning;
                  end;

                MyRDPClient.Server := ComboBox2.text;
                MyRDPClient.Domain := CurrentDomainName;
                if LabeledEdit1.text <> '' then
                  MyRDPClient.username := LabeledEdit1.text;
                if LabeledEdit2.text <> '' then
                  MyRDPClient.AdvancedSettings9.ClearTextPassword :=
                    LabeledEdit2.text;
                MyRDPClient.DesktopWidth := TabSheet9.width;
                MyRDPClient.DesktopHeight := TabSheet9.height;
                MyRDPClient.AdvancedSettings9.RedirectDrives := true;
                /// разрешить перенаправлять диски
                MyRDPClient.AdvancedSettings9.RedirectPrinters := true;
                /// разрешить перенаправлять принтеры
                MyRDPClient.AdvancedSettings9.RelativeMouseMode := true;
                ///
                MyRDPClient.AdvancedSettings9.RedirectClipboard := true;
                ///
                MyRDPClient.AdvancedSettings9.RedirectDevices := true;
                /// перенаправление других устройств PnP
                MyRDPClient.AdvancedSettings9.RedirectPorts := true;
                /// перенаправление портов
                MyRDPClient.AdvancedSettings9.EnableAutoReconnect := true;
                /// / разрешить переподключение при проблемах сети
                MyRDPClient.AdvancedSettings9.ConnectToAdministerServer := true;
                /// указывает должен ли элемент управления ActiveX пытаться подключиться к серверу в административных целях.
                MyRDPClient.AdvancedSettings9.
                  AudioCaptureRedirectionMode := true;
                /// указывает, перенаправляется ли устройство аудиовхода по умолчанию от клиента к удаленному сеансу
                MyRDPClient.AdvancedSettings9.EnableSuperPan := true;
                /// позволяет пользователю легко перемещаться по удаленному рабочему столу в полноэкранном режиме, когда размеры удаленного рабочего стола больше размеров текущего окна клиента. Вместо того, чтобы использовать полосы прокрутки для навигации по рабочему столу, пользователь может указать на границу окна, и удаленный рабочий стол будет автоматически прокручиваться в этом направлении. SuperPan не поддерживает более одного монитора.
                MyRDPClient.AdvancedSettings9.BandwidthDetection := true;
                /// Определяет, будут ли автоматически обнаружены изменения пропускной способности.
                MyRDPClient.AdvancedSettings9.authenticationLevel := 2;
                /// Задает уровень проверки подлинности для подключения.  (0.1.2)
                MyRDPClient.AdvancedSettings9.EnableCredSspSupport := true;
                /// Указывает или извлекает, включен ли для этого подключения  Поставщик службы безопасности учетных данных (CredSSP)
                MyRDPClient.SecuredSettings3.KeyboardHookMode := 1;
                MyRDPClient.Connect;
                SpeedButton1.Caption := 'Отключится';
                exit;
              except
                on E: Exception do
                  begin
                    Memo1.Lines.Add('Ошибка RDP"' + E.Message + '"');
                    Memo1.Lines.Add('---------------------------');
                    if TabSheet9.FindComponent('MyRDPClient') <> nil then
                      FreeAndNil(MyRDPClient);
                    SpeedButton1.Caption := 'Подключится';
                    exit;
                  end;
              end;
            end;
        end;
      if SpeedButton1.Caption = 'Отключится' then
        begin
          try
            // MyRDPClient.RequestClose;
            // memo1.Lines.Add('RDP - завершение сеанса');
            MyRDPClient.Disconnect;
            if TabSheet9.FindComponent('MyRDPClient') <> nil then
              FreeAndNil(MyRDPClient);
            SpeedButton1.Caption := 'Подключится';
          except
            on E: Exception do
              begin
                Memo1.Lines.Add('Ошибка RDP"' + E.Message + '"');
                Memo1.Lines.Add('---------------------------');
                if TabSheet9.FindComponent('MyRDPClient') <> nil then
                  FreeAndNil(MyRDPClient);
                SpeedButton1.Caption := 'Подключится';
                exit;
              end;
          end;
        end;
    end;
  except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка RDP"' + E.Message + '"');
        Memo1.Lines.Add('---------------------------');
        if TabSheet9.FindComponent('MyRDPClient') <> nil then
          FreeAndNil(MyRDPClient);
        SpeedButton1.Caption := 'Подключится';
        exit;
      end;
  end;

end;

procedure TfrmDomainInfo.MsRdpClient91AuthenticationWarningDismissed
  (Sender: TObject);
begin
  Memo1.Lines.Add('RDP - Предупреждение Аутентификации');
end;

procedure TfrmDomainInfo.MsRdpClient91AuthenticationWarningDisplayed
  (Sender: TObject);
const
  AutType: array [0 .. 4] of string = ('No authentication is used.',
    'Используется проверка подлинности сертификата',
    'Используется проверка подлинности Kerberos',
    'Используется оба типа проверки подлинности, сертификата и по протоколу',
    'неизвестное состояние');
begin
  if MyRDPClient.AdvancedSettings9.AuthenticationType <> 0 then
    Memo1.Lines.Add('RDP - Проверка подлинности. - ' +
      AutType[MyRDPClient.AdvancedSettings9.AuthenticationType]);
end;

procedure TfrmDomainInfo.MsRdpClient91AutoReconnected(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - Переподключение к клиентсому компьютеру.- ')
end;

procedure TfrmDomainInfo.MsRdpClient91AutoReconnecting(ASender: TObject;
  disconnectReason, attemptCount: Integer);
begin
  Memo1.Lines.Add
    ('RDP - Автоматическое переподключение сеанса удаленных рабочих столов.');
end;

procedure TfrmDomainInfo.MsRdpClient91AutoReconnecting2(ASender: TObject;
  disconnectReason: Integer; networkAvailable: WordBool;
  attemptCount, maxAttemptCount: Integer);
begin
  Memo1.Lines.Add
    ('RDP - клиентский элемент управления переподключается к удаленному сеансу');
end;

procedure TfrmDomainInfo.MsRdpClient91ChannelReceivedData(ASender: TObject;
  const chanName, Data: WideString);
begin
  Memo1.Lines.Add
    ('RDP - клиент получает данные на виртуальном канале с возможностью сценария');
end;

procedure TfrmDomainInfo.MsRdpClient91ConfirmClose(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - Завершение сеанса служб удаленного рабочего стола');
  if TabSheet9.FindComponent('MyRDPClient') <> nil then
    begin
      FreeAndNil(MyRDPClient);
    end;
  SpeedButton1.Caption := 'Подключится';
end;

procedure TfrmDomainInfo.MsRdpClient91Connected(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - установка соединения...');
end;

procedure TfrmDomainInfo.MsRdpClient91Connecting(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - инициализация соединения...');
end;

procedure TfrmDomainInfo.MsRdpClient91Disconnected(ASender: TObject;
  discReason: Integer);
begin
  Memo1.Lines.Add
    ('RDP - отключение от сервера узела сеансов удаленных рабочих столов');
  // if tabSheet9.FindComponent('MyRDPClient')<>nil then FreeAndNil(MyRDPClient);

  SpeedButton1.Caption := 'Подключится';
end;

procedure TfrmDomainInfo.MsRdpClient91Exit(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - выход');
  if TabSheet9.FindComponent('MyRDPClient') <> nil then
    begin
      FreeAndNil(MyRDPClient);
    end;
  SpeedButton1.Caption := 'Подключится';
end;

procedure TfrmDomainInfo.MsRdpClient91FatalError(ASender: TObject;
  errorCode: Integer);
begin
  Memo1.Lines.Add('RDP - произошщла неустранимая ошибка');
  if TabSheet9.FindComponent('MyRDPClient') <> nil then
    begin
      FreeAndNil(MyRDPClient);
    end;
  SpeedButton1.Caption := 'Подключится';
end;

procedure TfrmDomainInfo.MsRdpClient91LoginComplete(Sender: TObject);
begin
  Memo1.Lines.Add('RDP - успешный вход в систему');
  SpeedButton1.Caption := 'Отключится';
end;

procedure TfrmDomainInfo.MsRdpClient91LogonError(ASender: TObject;
  lError: Integer);
begin
  Memo1.Lines.Add('RDP - ошибка входа в систему');
  // if tabSheet9.FindComponent('MyRDPClient')<>nil then MyRDPClient.Free;
  SpeedButton1.Caption := 'Подключится';
end;

procedure TfrmDomainInfo.MsRdpClient91NetworkStatusChanged(ASender: TObject;
  qualityLevel: Cardinal; bandwidth, rtt: Integer);
begin
  Memo1.Lines.Add('RDP - произошло изменение состояния сети');
end;

procedure TfrmDomainInfo.MsRdpClient91Warning(ASender: TObject;
  warningCode: Integer);
begin
  Memo1.Lines.Add('RDP - OnWarning обнаружена не фатальная ошибка');
end;

/// ////////////////////////////////////////////////////////////////////////

procedure TfrmDomainInfo.TabSheet10Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton2.Caption := 'Драйвера';
end;

procedure TfrmDomainInfo.TabSheet11Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton2.Caption := 'Network';
end;

procedure TfrmDomainInfo.TabSheet1Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Caption := 'Автозагрузка';
end;

procedure TfrmDomainInfo.TabSheet2Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Caption := 'Процесы';
end;

procedure TfrmDomainInfo.TabSheet3Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Caption := 'Службы';
end;

procedure TfrmDomainInfo.TabSheet4Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := false;
  SpeedButton3.Visible := true;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
end;

procedure TfrmDomainInfo.TabSheet5Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton2.Caption := 'Все программы';
end;

procedure TfrmDomainInfo.TabSheet6Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton4.Visible := true;
end;

procedure TfrmDomainInfo.TabSheet7Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton2.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := true;
end;

procedure TfrmDomainInfo.TabSheet8Show(Sender: TObject);
begin
  SpeedButton1.Visible := false;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Visible := true;
  SpeedButton2.Caption := 'Профили';
end;

procedure TfrmDomainInfo.TabSheet9Resize(Sender: TObject);
begin
  if TabSheet9.FindComponent('MyRDPClient') <> nil then
    MyRDPClient.Reconnect(TabSheet9.width, TabSheet9.height);

end;

procedure TfrmDomainInfo.TabSheet9Show(Sender: TObject);
begin
  SpeedButton1.Visible := true;
  SpeedButton3.Visible := false;
  SpeedButton4.Visible := false;
  SpeedButton5.Visible := false;
  SpeedButton2.Visible := false;
end;

procedure TfrmDomainInfo.TaskListViewDblClick(Sender: TObject);
// редактирование задачи
begin
  if TaskListView.SelCount = 1 then
    begin
      EditTask.Caption := TaskListView.Selected.Caption;
      // имя задачи в заголовок
      EditTask.TaskName.Caption := TaskListView.Selected.Caption;
      // имя задачи в label т.к. происходит редактирование задачи
      // EditTask.ReadTableSelectedTask(EditTask.TaskName.Caption); // запускаем чтение таблицы и в label меняем описание таблицы на ее имя
      EditTask.ReadResulTask(EditTask.TaskName.Caption);
      // новая функция чтения таблицы
      if EditTask.WindowState = wsMinimized then
      // если форма свернута то восстанавливаем ее
        EditTask.WindowState := wsNormal
      else // иначе показываем
        EditTask.Show;
    end;
end;

procedure TfrmDomainInfo.Timer1Timer(Sender: TObject);
var
  loadmem, LoadPrivate: real;
  Myprocent: TThread;
begin
  if (lvWorkStation.Items.Count <> 0) and (pcRes.ActivePageIndex = 0) then
    begin
      begin
        loadmem := (SumMemory div 1024) / mem;
        LoadPrivate := (SetPeak div 1024) / mem;
        StatusBar1.Panels[3].text := 'CPU - ' + inttostr(SumTime) + ' %';
        StatusBar1.Panels[4].text := 'Физическая память: Рабочий набор - ' +
          floattostr(round((loadmem) * 100)) + '%' + ' / Пиковый набор - ' +
          floattostr(round((LoadPrivate) * 100)) + '%';
        StatusBar1.Panels[5].text := 'Процессов - ' +
          inttostr(lvWorkStation.Items.Count);
        Myprocent := unit2.Procent.Create(true);
        Myprocent.FreeOnTerminate := true; //самоуничтожение потока
        Myprocent.Start;
        // StatusBar1.Panels[2].Text:='Используется памяти - '+inttostr(SumMemory div 1024)+' Mb/ из - '+ floattostr(MyMemory)+' Mb';
      end;
    end;
end;

/// /////////////////////////////////////////////////////////////////

procedure TfrmDomainInfo.DBGrid1TitleClick(Column: TColumn);
var
/// // сортировка  DBgrid1
  f: string;
begin
  if DBGrid1.DataSource.DataSet.RecordCount <> 0 then
    begin
      f := Column.FieldName;
      if DBGrid1.Tag = 0 then
        begin
          Datam.FDQueryReadPropertiesPC.IndexFieldNames := f + ':A;';
          DBGrid1.Tag := 1;
        end
      else
        begin
          Datam.FDQueryReadPropertiesPC.IndexFieldNames := f + ':D;';
          DBGrid1.Tag := 0;
        end;
    end;
end;

procedure TfrmDomainInfo.ReadDB;
/// чтение из базы данных
begin
  try
    DBGrid1.DataSource := Datam.DataSource1;
    /// меняем датасоурс
    DBGrid1.DataSource.DataSet.DisableControls;

    /// / без него выдает ошибку при обновлении если включена инвентаризация
    // if not inventConf then   ///// если инвентаризация не вкулючена то подсчитываем количество компьютеров
    begin
    /// //  после сортировки по полю подсчет выдает ошибку
      /// 'SELECT PC_NAME,INV_NUMBER,DATE_INV,RESULT_INV,ERROR_INV,count(PC_ID) as countPC FROM MAIN_PC ORDER BY PC_NAME';
      // 'select count(*) as countPC from (SELECT * FROM MAIN_PC ORDER BY PC_NAME)'
      Datam.FDQueryReadPropertiesPC.DeleteIndexes;
      /// обнуляем индексы
      Datam.FDQueryReadPropertiesPC.IndexFieldNames := '';
      /// обнуляем индекс если была сортировка в Dbgrid1
      Datam.FDQueryReadPropertiesPC.SQL.Clear;
      Datam.FDQueryReadPropertiesPC.SQL.text :=
        'select count(PC_ID) as countPC from main_pc';
      Datam.FDQueryReadPropertiesPC.Open;
      StatusBarInv.Panels[0].text := 'Всего записей - ' +
        (Datam.FDQueryReadPropertiesPC.FieldByName('countPC').AsString);

      Datam.FDQueryReadPropertiesPC.SQL.Clear;
      Datam.FDQueryReadPropertiesPC.SQL.text :=
        'select count(PC_ID) as countPC from main_pc where result_inv=:pStat';
      Datam.FDQueryReadPropertiesPC.ParamByName('pStat').Value := 'OK';
      Datam.FDQueryReadPropertiesPC.Open;
      StatusBarInv.Panels[1].text := 'OK - ' +
        (Datam.FDQueryReadPropertiesPC.FieldByName('countPC').AsString);
    end;
    Datam.FDQueryReadPropertiesPC.DeleteIndexes;
    Datam.FDQueryReadPropertiesPC.SQL.Clear;
    Datam.FDQueryReadPropertiesPC.SQL.text :=
      'SELECT * FROM MAIN_PC ORDER BY PC_NAME';
    Datam.FDQueryReadPropertiesPC.Open;
    DBGrid1.DataSource.DataSet.EnableControls;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Update Config PC - ' + E.Message)
      end;
  end;
end;

/// /////////////////////////////////////////////////////////////////////////////////////
Function TfrmDomainInfo.createClearTreeView(treeView: TTreeView): bool;
/// построение пустова дерева
begin
  treeView.Items.Clear;
  treeView.Images := ImageHardWare;
  treeView.Items.Add(nil, 'Дата и время инвентаризации - ');
  treeView.Items.Add(nil, 'Материнская плата').ImageIndex := 1;
  treeView.Items.Add(nil, 'Процессор').ImageIndex := 2;
  treeView.Items.Add(nil, 'Память').ImageIndex := 3;
  treeView.Items.Add(nil, 'Видеокарта').ImageIndex := 4;
  treeView.Items.Add(nil, 'Сетевой интерфейс').ImageIndex := 5;
  treeView.Items.Add(nil, 'Жесткий диск').ImageIndex := 6;
  treeView.Items.Add(nil, 'CD/DVD').ImageIndex := 9;
  treeView.Items.Add(nil, 'Звук').ImageIndex := 10;
  treeView.Items.Add(nil, 'Монитор').ImageIndex := 7;
  treeView.Items.Add(nil, 'Операционная система').ImageIndex := 8;
  treeView.Items.Add(nil, 'Принтер').ImageIndex := 12;
end;

/// ////////////////////////////////////////////////////////////////////////////////////////
Function TfrmDomainInfo.createtreeView(FDQueryinfo: TFDQuery;
  treeView: TTreeView): bool; // построение заполненного дерева
var
  i, j, z, c, s, d, w: Integer;
  MyNode, Level1, Level2: TTreeNode;
  ArrayProc, MyArray, ArrayMem, ArrayVideo, arrayHDD, ArrayNI, NIcountip,
    NIcountGW, NIcountDNS, NIcountDHCP, NIcountWINS, NIarrayIP,
    MonName: TStringArray;
  BiosDes: TStringArray;
begin
  try
    treeView.Items.Clear;
    treeView.Images := ImageHardWare;
    /// ////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Дата и время инвентаризации - ' +
      vartostr(FDQueryinfo.FieldByName('DATE_INV').Value) + '/' +
      vartostr(FDQueryinfo.FieldByName('TIME_INV').Value));
    treeView.Items.Add(nil, 'Материнская плата').ImageIndex := 1;
    MyNode := treeView.Items.Item[treeView.Items.Count - 1];
    MyNode.SelectedIndex := MyNode.ImageIndex;
    if vartostr(FDQueryinfo.FieldByName('MONTERBOARD').Value) <> '' then
      treeView.Items.AddChild(MyNode,
        'Модель - ' + vartostr(FDQueryinfo.FieldByName('MONTERBOARD').Value))
        .ImageIndex := 0;;
    // := vartostr(FDQueryinfo.FieldByName('MONTERBOARD').Value);
    if vartostr(FDQueryinfo.FieldByName('MONTERBOARD_SN').Value) <> '' then
      treeView.Items.AddChild(MyNode,
        'SN - ' + vartostr(FDQueryinfo.FieldByName('MONTERBOARD_SN').Value))
        .ImageIndex := 0;
    if vartostr(FDQueryinfo.FieldByName('MONTERBOARD_MANUFACTURE').Value) <> ''
    then
      treeView.Items.AddChild(MyNode, 'Производитель - ' +
        vartostr(FDQueryinfo.FieldByName('MONTERBOARD_MANUFACTURE').Value))
        .ImageIndex := 0;
    MyNode.Expanded := true;
    /// разворачиваем узел

    /// //////////////////////////////////////////////////////////////////////
    /// BiosDes,BIOSSN,SMBIOSVERSION,BIOSVERSION,BIOSPRIMARY,BIOSMANUF
    BiosDes :=
      kolitem(vartostr(FDQueryinfo.FieldByName('SMBIOSBIOSVERSION').Value));
    if BiosDes <> nil then
      for i := 0 to Length(BiosDes) - 1 do
        begin
          treeView.Items.AddChild(MyNode, 'BIOS - ' + BiosDes[i])
            .ImageIndex := 0;
          /// картинка   BiosDes[i] -Имя BIOS
          Level1 := treeView.Items.Item[treeView.Items.Count - 1];
          // level1.SelectedIndex:=level1.ImageIndex; /// картинка
          // treeView.Items.AddChild(Level1,BiosDes[i]); /// Имя BIOS
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 0;
          /// картинка вложенная
          MyNode.SelectedIndex := MyNode.ImageIndex;
          /// картинка вложенная
          MyArray := kolitem(vartostr(FDQueryinfo.FieldByName('BIOSSN').Value));
          /// //  статус bios
          treeView.Items.AddChild(MyNode, 'SN - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('BIOSSTATUS').Value));
          /// //  статус bios
          treeView.Items.AddChild(MyNode, 'Статус - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('BIOSMANUFAC').Value));
          /// / производитель bios
          treeView.Items.AddChild(MyNode, 'Разработчик - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName
            ('BIOS_DESCRIPTION').Value));
          /// //  версия
          treeView.Items.AddChild(MyNode, 'SMBIOS - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('BIOSPRIMARY').Value));
          /// //  основной или нет
          treeView.Items.AddChild(MyNode, 'Основной BIOS - ' + MyArray[i]);
          /// /////////////////////////////////////////////////////////////////////////
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName
            ('BIOSCHARACTERISTICS').Value));
          if MyArray <> nil then
            begin
              treeView.Items.AddChild(MyNode, 'Поддерживаемые функции')
                .ImageIndex := 0;
              MyNode := treeView.Items.Item[treeView.Items.Count - 1];
              for z := 0 to Length(MyArray) - 1 do
                begin
                  treeView.Items.AddChild(MyNode, ' - ' + MyArray[z]);
                end;
            end;
          /// ////////////////////////////////////////////////////////////////////////
          // MyNode.Expanded:=true; /// разворачиваем узел
        end;

    /// //////////////////////////////////////////////////////////////
    ArrayProc := kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR').Value));
    treeView.Items.Add(nil, 'Процессор').ImageIndex := 2;
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    /// при клике не меняется картинка

    for i := 0 to Length(ArrayProc) - 1 do
      begin
        treeView.Items.AddChild(Level1, ArrayProc[i]).ImageIndex := 2;
        /// картинка
        MyNode := treeView.Items.Item[treeView.Items.Count - 1];
        MyNode.SelectedIndex := MyNode.ImageIndex;
        /// при клике не меняется картинка
        MyArray :=
          kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR_CORE').Value));
        if MyArray <> nil then
          treeView.Items.AddChild(MyNode, 'Ядер - ' + MyArray[i]);
        MyArray :=
          kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR_LOGPROC').Value));
        if MyArray <> nil then
          treeView.Items.AddChild(MyNode, 'Логических процессоров - ' +
            MyArray[i]);
        MyArray :=
          kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR_SPEED').Value));
        if MyArray <> nil then
          treeView.Items.AddChild(MyNode, 'Частота - ' + MyArray[i] + ' МГц');
        MyArray :=
          kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR_ARCH').Value));
        if MyArray <> nil then
          treeView.Items.AddChild(MyNode, 'Архитектура - ' + MyArray[i]);
        MyArray :=
          kolitem(vartostr(FDQueryinfo.FieldByName('PROCESSOR_SOKET').Value));
        if MyArray <> nil then
          treeView.Items.AddChild(MyNode, 'Socket - ' + MyArray[i]);
      end;
    Level1.Expanded := true;
    /// разворачиваем узел
    /// //////////////////////////////////////////////////////////////////////////////////////////////

    ArrayMem := kolitem(vartostr(FDQueryinfo.FieldByName('MEMORY_SIZE').Value));
    MyArray := kolitem(vartostr(FDQueryinfo.FieldByName('MEMORY_TYPE').Value));
    treeView.Items.Add(nil, 'Память ('+SummaDDR(ArrayMem)+' Мб)').ImageIndex := 3;
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    /// картинка
    if ArrayMem <> nil then
      for i := 0 to Length(ArrayMem) - 1 do
        begin
          MyNode := treeView.Items.Item[treeView.Items.Count - 1 - i];
          MyNode.SelectedIndex := MyNode.ImageIndex;
          /// картинка
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, ArrayMem[i] + ' Мб (' + MyArray[i]
              + ' МГц)')
          else
            treeView.Items.AddChild(MyNode, ArrayMem[i] + ' Мб');
          treeView.Items.Item[treeView.Items.Count - 1].ImageIndex :=
            MyNode.ImageIndex;
          /// картинка
          treeView.Items.Item[treeView.Items.Count - 1].SelectedIndex :=
            MyNode.ImageIndex;
          /// картинка
        end;
    Level1.Expanded := true;
    /// /////////////////////////////////////////////
    treeView.Items.Add(nil, 'Видеокарта').ImageIndex := 4;
    ArrayVideo := kolitem(vartostr(FDQueryinfo.FieldByName('VIDEOCARD').Value));
    MyArray :=
      kolitem(vartostr(FDQueryinfo.FieldByName('VIDEOCARD_MEM').Value));
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    /// узел для развертывания
    Level1.SelectedIndex := Level1.ImageIndex;
    /// картинка
    if ArrayVideo <> nil then
      for i := 0 to Length(ArrayVideo) - 1 do
        begin
          MyNode := treeView.Items.Item[treeView.Items.Count - 1 - i];
          MyNode.SelectedIndex := MyNode.ImageIndex;
          /// картинка
          treeView.Items.AddChild(MyNode, ArrayVideo[i] + ' (' + MyArray[i]
            + ' Мб)');
          treeView.Items.Item[treeView.Items.Count - 1].ImageIndex :=
            MyNode.ImageIndex;
          /// картинка
          treeView.Items.Item[treeView.Items.Count - 1].SelectedIndex :=
            MyNode.ImageIndex;
          /// картинка
        end;
    Level1.Expanded := true;
    /// ///////////////////////////////////////////
    treeView.Items.Add(nil, 'Сетевой интерфейс').ImageIndex := 5;
    /// картинка
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    ArrayNI :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORKINTERFACE').Value));
    NIcountip :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_COUNT_IP').Value));
    NIarrayIP := kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_IP').Value));
    NIcountGW :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_COUNT_GATEWAY').Value));
    NIcountDNS :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_COUNT_DNS').Value));
    NIcountDHCP :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_COUNT_DHCP').Value));
    NIcountWINS :=
      kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_COUNT_WINS').Value));
    c := 0;
    s := 0;
    j := 0;
    i := 0;
    d := 0;
    w := 0;
    Level1.SelectedIndex := Level1.ImageIndex;
    /// картинка
    if ArrayNI <> nil then
      for i := 0 to Length(ArrayNI) - 1 do
        begin
          treeView.Items.AddChild(Level1, ArrayNI[i]);
          /// Имя сетевого интерфейса
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 5;
          /// картинка вложенная
          MyNode.SelectedIndex := MyNode.ImageIndex;
          /// картинка вложенная
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_MAC').Value));
          /// / MAC адрес
          treeView.Items.AddChild(MyNode, 'MAC адрес - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_SPEED').Value));
          /// //  Скорость интерфейса
          treeView.Items.AddChild(MyNode, 'Скорость - ' + MyArray[i]);

          Level2 := MyNode;

          if strtoint(NIcountip[i]) <> 0 then
            for j := 0 to strtoint(NIcountip[i]) - 1 do
            /// количество адресов на интерфейсе
              begin
                treeView.Items.AddChild(Level2, 'IP адрес - ' + NIarrayIP[c]);
                MyNode := treeView.Items.Item[treeView.Items.Count - 1];
                MyArray :=
                  kolitem(vartostr(FDQueryinfo.FieldByName
                  ('NETWORK_MASK').Value));
                /// ///   адрес интерфейса
                if MyArray <> nil then
                  treeView.Items.AddChild(MyNode, 'Маска - ' + MyArray[c]);
                Inc(c);
              end;

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_GATEWAY').Value));
          if strtoint(NIcountGW[i]) <> 0 then
            for z := 0 to strtoint(NIcountGW[i]) - 1 do
              begin
                treeView.Items.AddChild(Level2, 'Шлюз - ' + MyArray[z]);
              end;

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_DNS').Value));
          if MyArray <> nil then
            begin
              if strtoint(NIcountDNS[i]) <> 0 then
                for z := 0 to strtoint(NIcountDNS[i]) - 1 do
                  begin
                    treeView.Items.AddChild(Level2,
                      'DNS сервер - ' + MyArray[s]);
                    /// dns сервера
                    Inc(s);
                  end;
            end;

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_DHCP').Value));
          if MyArray <> nil then
            begin
              if strtoint(NIcountDHCP[i]) <> 0 then
                for z := 0 to strtoint(NIcountDHCP[i]) - 1 do
                  begin
                    treeView.Items.AddChild(Level2,
                      'DHCP сервер- ' + MyArray[d]);
                    Inc(d);
                  end;
            end;

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('NETWORK_WINS').Value));
          if MyArray <> nil then
            begin
              if strtoint(NIcountWINS[i]) <> 0 then
                for z := 0 to strtoint(NIcountWINS[i]) - 1 do
                  begin
                    treeView.Items.AddChild(Level2,
                      'WINS сервер- ' + MyArray[d]);
                    Inc(w);
                  end;
            end;

          Level1.Expanded := true;
          /// разворачиваем узел
        end;
    /// //////////////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Жесткий диск').ImageIndex := 6;
    /// картинка
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    arrayHDD := kolitem(vartostr(FDQueryinfo.FieldByName('HDD_NAME').Value));
    if arrayHDD <> nil then
      for i := 0 to Length(arrayHDD) - 1 do
        begin
          treeView.Items.AddChild(Level1, arrayHDD[i]);
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 6;
          MyNode.SelectedIndex := MyNode.ImageIndex;
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_TYPE').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Тип - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_SIZE').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Объем ' + MyArray[i] + ' Гб');
          MyArray := kolitem(vartostr(FDQueryinfo.FieldByName('HDD_SN').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'SN - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName
            ('HDD_INTERFACETYPE').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Интерфейс - ' + MyArray[i]);
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_FIRMWARE').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Firmware - ' + MyArray[i]);

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_SMART_OS').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'S.M.A.R.T (Опер.сист) - ' +
              MyArray[i]);

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_MY_SMART').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Оценка S.M.A.R.T (%) - ' +
              MyArray[i]);

          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('HDD_CUR_TEMP').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Температура (°С) - ' + MyArray[i]);

        end;
    Level1.Expanded := true;
    /// разворачиваем узел
    /// ///////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'CD/DVD').ImageIndex := 9;
    /// картинка
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    arrayHDD := kolitem(vartostr(FDQueryinfo.FieldByName('DVDROM_NAME').Value));
    if arrayHDD <> nil then
      for i := 0 to Length(arrayHDD) - 1 do
        begin
          treeView.Items.AddChild(Level1, arrayHDD[i]);
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 9;
          /// вложенная картинка
          MyNode.SelectedIndex := MyNode.ImageIndex;
          MyArray :=
            kolitem(vartostr(FDQueryinfo.FieldByName('DVDROM_TYPE').Value));
          if MyArray <> nil then
            treeView.Items.AddChild(MyNode, 'Тип - ' + MyArray[i]);
        end;
    Level1.Expanded := true;
    /// разворачиваем узел
    /// /////////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Звук').ImageIndex := 10;
    /// картинка
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    MyArray := kolitem(vartostr(FDQueryinfo.FieldByName('SOUND_NAME').Value));
    if MyArray <> nil then
      for i := 0 to Length(MyArray) - 1 do
        begin
          treeView.Items.AddChild(Level1, MyArray[i]);
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 10;
          MyNode.SelectedIndex := MyNode.ImageIndex;
        end;
    Level1.Expanded := true;
    /// разворачиваем узел
    /// ///////////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Монитор').ImageIndex := 7;
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;

    if (FDQueryinfo.FieldByName('MONITOR_COUNT').Value <> null) then
      begin
        MonName :=
          kolitem(vartostr(FDQueryinfo.FieldByName('MONITOR_NAME').Value));
        for i := 0 to strtoint((FDQueryinfo.FieldByName('MONITOR_COUNT')
          .Value)) - 1 do
          begin
            if MonName <> nil then
              treeView.Items.AddChild(Level1, MonName[i]);
            MyNode := treeView.Items.Item[treeView.Items.Count - 1];
            MyNode.ImageIndex := 7;
            MyNode.SelectedIndex := MyNode.ImageIndex;
            MyArray :=
              kolitem(vartostr(FDQueryinfo.FieldByName('MONITOR_MANUF').Value));
            /// ///   адрес интерфейса
            if MyArray <> nil then
              treeView.Items.AddChild(MyNode, 'Производитель- ' + MyArray[i]);
            MyArray :=
              kolitem(vartostr(FDQueryinfo.FieldByName('MONITOR_HW').Value));
            /// ///   адрес интерфейса
            if MyArray <> nil then
              treeView.Items.AddChild(MyNode, 'Разрешение - ' + MyArray[i]);
            MyArray :=
              kolitem(vartostr(FDQueryinfo.FieldByName('MONITOR_DPI').Value));
            /// ///   адрес интерфейса
            if MyArray <> nil then
              treeView.Items.AddChild(MyNode, 'DPI - ' + MyArray[i]);
          end;
        Level1.Expanded := true;
        /// разворачиваем узел
      end;
    /// ////////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Операционная система').ImageIndex := 8;
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    if vartostr(FDQueryinfo.FieldByName('OS_NAME').Value) <> '' then
      treeView.Items.AddChild(Level1,
        (vartostr(FDQueryinfo.FieldByName('OS_NAME').Value)));
    if (vartostr(FDQueryinfo.FieldByName('OS_SP').Value)) <> '' then
      treeView.Items.AddChild(Level1, 'Service Pack - ' +
        (vartostr(FDQueryinfo.FieldByName('OS_SP').Value)));
    if vartostr(FDQueryinfo.FieldByName('OS_VER').Value) <> '' then
      treeView.Items.AddChild(Level1,
        'Версия - ' + (vartostr(FDQueryinfo.FieldByName('OS_VER').Value)));
    if vartostr(FDQueryinfo.FieldByName('OS_TYPE').Value) <> '' then
      treeView.Items.AddChild(Level1,
        'Тип - ' + (vartostr(FDQueryinfo.FieldByName('OS_TYPE').Value)));
    if vartostr(FDQueryinfo.FieldByName('OS_KEY').Value) <> '' then
      treeView.Items.AddChild(Level1,
        'Serial - ' + (vartostr(FDQueryinfo.FieldByName('OS_KEY').Value)));
    if vartostr(FDQueryinfo.FieldByName('OS_DATEINSTALL').Value) <> '' then
      treeView.Items.AddChild(Level1, 'Дата установки - ' +
        (vartostr(FDQueryinfo.FieldByName('OS_DATEINSTALL').Value)) + ' ' +
        (vartostr(FDQueryinfo.FieldByName('OS_TIMEINSTALL').Value)));
    if (vartostr(FDQueryinfo.FieldByName('INV_NUMBER').Value)) <> '' then
      treeView.Items.AddChild(Level1,
        'Инв № ' + (vartostr(FDQueryinfo.FieldByName('INV_NUMBER').Value)));
    if vartostr(FDQueryinfo.FieldByName('USER_NAME').Value) <> '' then
      treeView.Items.AddChild(Level1, 'Пользователь - ' +
        (vartostr(FDQueryinfo.FieldByName('USER_NAME').Value)));
    Level1.Expanded := true;
    /// разворачиваем узел
    /// /////////////////////////////////////////////////////////////////////////////////
    /// /////////////////////////////////////////////////////////////////////////
    treeView.Items.Add(nil, 'Принтер').ImageIndex := 12;
    /// картинка
    Level1 := treeView.Items.Item[treeView.Items.Count - 1];
    Level1.SelectedIndex := Level1.ImageIndex;
    MyArray := kolitem(vartostr(FDQueryinfo.FieldByName('PRINTER_NAME').Value));
    if MyArray <> nil then
      for i := 0 to Length(MyArray) - 1 do
        begin
          treeView.Items.AddChild(Level1, MyArray[i]);
          MyNode := treeView.Items.Item[treeView.Items.Count - 1];
          MyNode.ImageIndex := 12;
          MyNode.SelectedIndex := MyNode.ImageIndex;
        end;
    Level1.Expanded := true;
    /// разворачиваем узел
    /// ////////////////////////////////////////////////////////////////////////////////////
    treeView.Items.Item[0].Selected := true;
  except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка функции CreatetreeView: ' + E.Message);
        Result := false;
      end;
  end;
end;

procedure TfrmDomainInfo.DBGrid1CellClick(Column: TColumn);
begin
  try
    if (DBGrid1.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid1.SelectedRows.Count = 1) then
    /// если DBgrid не пустой то выбираем из таблицы конфигурацию
      begin
        Datam.FDQueryRead3.SQL.Clear;
        /// // чтение всех конфигураций  и передача DBgrid2
        Datam.FDQueryRead3.SQL.text :=
          'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
          (string(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME').Value)) +
          '''' + ' ORDER BY DATE_INV DESC';
        Datam.FDQueryRead3.Open;
        if CheckBox1.Checked then
          begin
            Datam.FDQueryRead2.SQL.Clear;
            /// чтение конфигурации первого вхождения и передача в treeview
            Datam.FDQueryRead2.SQL.text :=
              'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
              (string(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME').Value))
              + '''' + ' ORDER BY DATE_INV DESC';
            Datam.FDQueryRead2.Open;
            createtreeView(Datam.FDQueryRead2, TreeView1);
            // Функция постороения дерева
            Datam.FDQueryRead2.Close;
          end;
        GroupBox7.Caption := 'Список конфигураций - ' +
          string(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME').Value);
      end;
    /// statusBarInv.Panels[2].Text:=inttostr(Dbgrid1.SelectedRows.Count);  /// колличество выделенных строк
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Selected PC in table CONFIG_PC - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.DBGrid2CellClick(Column: TColumn);
begin
  try
    if (DBGrid2.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid2.SelectedRows.Count = 1) then
    /// если DBgrid не пустой то выбираем из таблицы конфигурацию
      begin
        Datam.FDQueryRead2.SQL.Clear;
        /// чтение выбранной конфигурации и передача в treeview
        Datam.FDQueryRead2.SQL.text :=
          'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
          (string(DBGrid2.DataSource.DataSet.FieldByName('PC_NAME').Value)) +
          ''' AND DATE_INV= ''' +
          (string(DBGrid2.DataSource.DataSet.FieldByName('DATE_INV').Value)) +
          ''' AND TIME_INV= ''' +
          (string(DBGrid2.DataSource.DataSet.FieldByName('TIME_INV')
          .Value)) + '''';
        Datam.FDQueryRead2.Open;
        // createtree;   /// если данные есть запуск процедуры создания дерева
        createtreeView(Datam.FDQueryRead2, TreeView1);
        // Функция постороения дерева
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Selected PC in table CONFIG_PC sort - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton15Click(Sender: TObject);
begin
  ReadDB;
end;

procedure TfrmDomainInfo.SpeedButton14Click(Sender: TObject);
begin
  try
    // DbGrid1.DataSource.DataSet.Locate('PC_NAME;INV_NUMBER',  /// поиск по соответствию 2х полей
    // vararrayof([FindPC.text,FindPC.text]),[loCaseInsensitive, loPartialKey]);
    if DBGrid1.DataSource.DataSet.RecordCount <> 0 then
      begin
        if not DBGrid1.DataSource.DataSet.Locate('PC_NAME', FindPC.text,
          [loCaseInsensitive, loPartialKey]) then
          if not DBGrid1.DataSource.DataSet.Locate('INV_NUMBER', FindPC.text,
            [loCaseInsensitive, loPartialKey]) then
            showmessage('Ничего не найдено!!!');
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find Config - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.FindPCKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    SpeedButton14.Click;

end;

procedure TfrmDomainInfo.FindPCSoftKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    SpeedButton26.Click;
end;

procedure TfrmDomainInfo.SpeedButton16Click(Sender: TObject);
/// удаление записей
var
  i: Integer;
begin
  try
    if (DBGrid1.SelectedRows.Count > 0) and
      (DBGrid1.DataSource.DataSet.RecordCount <> 0) then
    /// / если выделенно больше 0 строк
      begin
        i := MessageBox(self.Handle,
          PChar('Кол-во записей для удаления - ' +
          inttostr(DBGrid1.SelectedRows.Count) +
          '. Удалять конфигурацию компьютера?'),
          PChar('Удаление данных конфигурации'), MB_YESNOCANCEL);
        // i:=MessageDlg('Кол-во записей для удаления - '+inttostr(Dbgrid1.SelectedRows.Count)+ #10#13+'Удалить?', mtConfirmation,[mbYes,mbCancel],0);
        if i = IDYes then
          begin
            DBGrid1.DataSource.DataSet.DisableControls;
            for i := 0 to DBGrid1.SelectedRows.Count - 1 do
              begin
                DBGrid1.DataSource.DataSet.Bookmark :=
                  DBGrid1.SelectedRows.Items[i];
                FDTransactionWrite1.Options.AutoCommit := true;
                FDTransactionWrite1.Options.AutoStart := true;
                FDTransactionWrite1.Options.AutoStop := true;
                FDQueryDelete.Active := false;
                FDQueryDelete.SQL.Clear;
                FDQueryDelete.SQL.text :=
                  'delete FROM CONFIG_PC WHERE PC_NAME=''' +
                  vartostr(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME')
                  .Value) + '''';
                FDQueryDelete.ExecSQL;
                FDQueryDelete.Active := false;
                FDQueryDelete.SQL.Clear;
                FDQueryDelete.SQL.text := 'delete FROM MAIN_PC WHERE PC_ID=''' +
                  vartostr(DBGrid1.DataSource.DataSet.FieldByName('PC_ID')
                  .Value) + '''';
                FDQueryDelete.ExecSQL;
                FDTransactionWrite1.Options.AutoCommit := false;
                FDTransactionWrite1.Options.AutoStart := false;
                FDTransactionWrite1.Options.AutoStop := false;
              end;
          end;
        if i = IDNO then
          begin
            DBGrid1.DataSource.DataSet.DisableControls;
            for i := 0 to DBGrid1.SelectedRows.Count - 1 do
              begin
                DBGrid1.DataSource.DataSet.Bookmark :=
                  DBGrid1.SelectedRows.Items[i];
                FDTransactionWrite1.Options.AutoCommit := true;
                FDTransactionWrite1.Options.AutoStart := true;
                FDTransactionWrite1.Options.AutoStop := true;
                FDQueryDelete.Active := false;
                FDQueryDelete.SQL.Clear;
                FDQueryDelete.SQL.text := 'delete FROM MAIN_PC WHERE PC_ID=''' +
                  vartostr(DBGrid1.DataSource.DataSet.FieldByName('PC_ID')
                  .Value) + '''';
                FDQueryDelete.ExecSQL;
                FDTransactionWrite1.Options.AutoCommit := false;
                FDTransactionWrite1.Options.AutoStart := false;
                FDTransactionWrite1.Options.AutoStop := false;
              end;
          end;
        DBGrid1.DataSource.DataSet.EnableControls;
        DBGrid1.DataSource.DataSet.Close;
        /// закрыть
        DBGrid1.DataSource.DataSet.Open;
        /// /  открыть для обновления грида

      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Delete invent Config - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton20Click(Sender: TObject);
var
/// /////// удаление записи конфигураций
  i: Integer;
begin
  try
    if (DBGrid2.SelectedRows.Count <> 0) and
      (DBGrid2.DataSource.DataSet.RecordCount <> 0) then
    /// / если выделенно больше 0 строк
      begin
        i := MessageDlg('Удалить конфигурацию?', mtConfirmation,
          [mbYes, mbCancel], 0);
        if i = IDYes then
          begin
            DBGrid2.DataSource.DataSet.DisableControls;
            for i := 0 to DBGrid2.SelectedRows.Count - 1 do
              begin
                DBGrid2.DataSource.DataSet.Bookmark :=
                  DBGrid2.SelectedRows.Items[i];
                FDTransactionWrite1.Options.AutoCommit := true;
                FDTransactionWrite1.Options.AutoStart := true;
                FDTransactionWrite1.Options.AutoStop := true;
                FDQueryDelete.Active := false;
                FDQueryDelete.SQL.Clear;
                FDQueryDelete.SQL.text := 'delete FROM CONFIG_PC WHERE PC_ID='''
                  + vartostr(DBGrid2.DataSource.DataSet.FieldByName('PC_ID')
                  .Value) + '''';
                FDQueryDelete.ExecSQL;
                FDTransactionWrite1.Options.AutoCommit := false;
                FDTransactionWrite1.Options.AutoStart := false;
                FDTransactionWrite1.Options.AutoStop := false;
              end;
            DBGrid2.DataSource.DataSet.EnableControls;
            DBGrid2.DataSource.DataSet.Close;
            /// / закрыть
            DBGrid2.DataSource.DataSet.Open;
            /// /// открыть для изменения отображаемых данных
          end;

      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Delete selected sort CONFIG_PC sort - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton21Click(Sender: TObject);
begin
  try
  /// поиск компьютеров с нарушением конфигурации (поиск дублей)
    DBGrid1.DataSource := Datam.DataSourceViolation;
    /// меняем датасоурс
    DBGrid1.DataSource.DataSet.ControlsDisabled;
    Datam.FDQueryFindViolations.DeleteIndexes; // FDQueryFindViolations
    Datam.FDQueryFindViolations.IndexFieldNames := '';
    /// обнуляем индекс если была сортировка в Dbgrid
    Datam.FDQueryFindViolations.SQL.Clear;
    // ниже вариант с поиском дублей в таблице  MAIN_PC
    // FDQueryRead.SQL.Text:='SELECT * FROM MAIN_PC T1 WHERE (select count(*)'
    // +'FROM MAIN_PC T2 WHERE T1.PC_NAME=T2.PC_NAME)>1 ORDER BY PC_NAME';
    /// Ниже рабочий вариант  выдает с дублями из таблицы CONFIG_PC
    // FDQueryRead.SQL.Text:='SELECT PC_NAME,INV_NUMBER,DATE_INV FROM CONFIG_PC T1 WHERE (select count(*)'
    // +'FROM CONFIG_PC T2 WHERE T1.PC_NAME=T2.PC_NAME)>1 ORDER BY PC_NAME';
    /// рабочий вариант, находит дубли в CONFIG_PC и выдает записи из MAIN_PC
    Datam.FDQueryFindViolations.SQL.text :=
      'SELECT * FROM MAIN_PC T1 WHERE PC_NAME=' +
      '(SELECT PC_NAME FROM CONFIG_PC T2 WHERE T1.PC_NAME=T2.PC_NAME' +
      ' group by PC_NAME HAVING COUNT(PC_NAME)>1) ';
    Datam.FDQueryFindViolations.Open;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find violation Config - ' + E.Message)
      end;
  end;
  DBGrid1.DataSource.DataSet.EnableControls;
end;

procedure TfrmDomainInfo.SpeedButton17Click(Sender: TObject);
Var
  i: byte;
begin
  try
    if not SolveExitInvConf then
    /// если запущена инвентаризация то задаем вопрос и останавливаем
      begin
        i := MessageDlg('Остановить инвентаризацию? ', mtConfirmation,
          [mbYes, mbCancel], 0);
        if i = IDCancel then
          exit;
        InventConf := false;
        Memo1.Lines.Add('Остановка инвентаризации оборудования...ожидайте');
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error',datetostr(Date)+'/'+timetostr(time)+' Stop invent Config - '+E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton18Click(Sender: TObject);
begin

  if DBGrid2.SelectedRows.Count <> 2 then
    begin
      showmessage('Необходимо выбрать две конфигурации');
      exit;
    end;

  TreeView1.Align := AlLeft;
  TreeView2.Align := AlRight;
  TreeView2.Visible := true;
  TreeView2.width := GroupBox10.width div 2;
  TreeView1.width := GroupBox10.width div 2;
  TreeView2.left := TreeView1.left + TreeView1.width;
  SpeedButton19.left := GroupBox10.width - SpeedButton19.width - 10;
  SpeedButton19.Visible := true;
  // DBGrid2.Datasource.Dataset.DisableControls;
  // TempBookmark:=DBgrid2.DataSource.DataSet.Bookmark;
  // TempBookmark:=DBgrid2.DataSource.DataSet.Bookmark;
  // setLength(OneConf,fdQueryRead3.FieldCount);
  // setLength(TwoConf,fdQueryRead3.FieldCount);
  // for I := 0 to fdQueryRead3.FieldCount-1 do
  begin
    DBGrid2.DataSource.DataSet.Bookmark := DBGrid2.SelectedRows.Items[0];
    { if
      (vartostr(DBGrid2.DataSource.DataSet.FieldByName('PC_ID').Value)=
      vartostr(FDQueryRead3.FieldByName('PC_ID').Value)) then
      begin
      //OneConf[i]:=vartostr(fdQueryRead3.Fields[i].Value);
      ////memo1.Lines.Add(OneConf[0]);
      end; }
  end;
  createtreeView(Datam.FDQueryRead3, TreeView1); // Функция постороения дерева
  // for I := 0 to fdQueryRead3.FieldCount-1 do
  begin
    DBGrid2.DataSource.DataSet.Bookmark := DBGrid2.SelectedRows.Items[1];
    { if
      (vartostr(DBGrid2.DataSource.DataSet.FieldByName('PC_ID').Value)=
      vartostr(FDQueryRead3.FieldByName('PC_ID').Value)) then
      begin
      //TwoConf[i]:=vartostr(fdQueryRead3.Fields[i].Value);
      // memo1.Lines.Add(TwoConf[0]);
      end; }
  end;
  createtreeView(Datam.FDQueryRead3, TreeView2); // Функция постороения дерева
  comparisonTV; // сравнение деревьев
  DBGrid2.SelectedRows.Clear;
  // DBgrid2.DataSource.DataSet.GotoBookmark(TempBookmark);
  // DBgrid2.DataSource.DataSet.FreeBookmark(TempBookmark);
  // DBGrid2.Datasource.Dataset.EnableControls

end;

procedure TfrmDomainInfo.SpeedButton19Click(Sender: TObject);
begin
  TreeView1.Align := alClient;
  TreeView2.Visible := false;
  SpeedButton19.Visible := false;
end;

Procedure TfrmDomainInfo.comparisonTV;
/// сравнение деревьев
var
  i, z, w: Integer;
begin
  try
    for i := 0 to TreeView1.Items.Count - 1 do
    /// от еденицы , потому что первый узел не сравниваем там дата и время инвентаризации
      begin
        if (TreeView1.Items[i].level = 0) then
          begin
            for z := 0 to TreeView2.Items.Count - 1 do
              begin
                if (TreeView2.Items[z].text = TreeView1.Items[i].text) then
                /// если имена разделов совпадают
                  begin
                    if (TreeView1.Items.Item[i].Count = TreeView2.Items.Item[z]
                      .Count) then
                    // если количество подразделов совпадает  проверяем по тексту в разделах
                      begin
                        for w := 0 to TreeView1.Items.Item[i].Count - 1 do
                          begin
                            if TreeView1.Items.Item[i].Item[w].text <>
                              TreeView2.Items.Item[z].Item[w].text then
                              begin
                              /// если текст не совпадает меняем картинку на ветку
                                TreeView1.Items.Item[i].Item[w]
                                  .ImageIndex := 11;
                                TreeView1.Items.Item[i].Item[w]
                                  .SelectedIndex := 11;
                                TreeView2.Items.Item[z].Item[w]
                                  .ImageIndex := 11;
                                TreeView2.Items.Item[z].Item[w]
                                  .SelectedIndex := 11;
                              end;
                          end;
                      end
                    else
                    /// иначе если количество подразделов не совпадает то меняем картинку на главную ветку
                      begin
                        TreeView1.Items[i].ImageIndex := 11;
                        TreeView1.Items[i].SelectedIndex := 11;
                        TreeView2.Items[z].ImageIndex := 11;
                        TreeView2.Items[z].SelectedIndex := 11;
                      end;
                  end;
              end;
          end;
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Error Comparison of the tree 1 and 2 - ' + E.Message)
      end;
  end;
end;

/// //////////////////////////////////////Excel   все обращения к excel в unit ExportExcel

function delZ(s: string): string;
begin
/// функция удаления из строки /**/
  if (s <> '') then
    begin
      if (pos('/**/', s) <> 0) then
        begin
          Result := '';
          s := StringReplace(s, '/**/', ' ,'#10#13,
            [rfReplaceAll, rfIgnoreCase]);
          Delete(s, Length(s) - 3, 3);
          Result := s;
        end
      else
        Result := s;
    end
  else
    Result := s;
end;

procedure TfrmDomainInfo.SpeedButton22Click(Sender: TObject);
var
/// / экспорт в Excel
  i, z: Integer;
  saveexcel: Tsavedialog;
  Myworksheets: Variant;
  /// лист в ексель
begin
  try
    if (DBGrid1.SelectedRows.Count <= 0) then
      begin
        showmessage('Необходимо выделить список компьютеров');
        exit;
      end;
    if InventConf then
      begin
        showmessage('Дождитесь завершения инвентаризации');
        exit;
      end;
    AllRunExcel;
    /// запуск Excel
    // myExcel.visible:=true; // показать Excel
    MyExcel.WorkBooks[1].WorkSheets[1].Name := 'Инвентаризация';
    /// имя листа
    Myworksheets := MyExcel.WorkBooks[1].WorkSheets[1];
    // MyWorkSheets.range['B1:D1'].Columns.AutoFit; /// авто ширина ячеек
    // myworksheets.columns['E:F'].EntireColumn.AutoFit;  /// авто ширина колонок
    Myworksheets.Columns['A:AL'].columnwidth := 20; // ширина колонки
    Myworksheets.Columns['A:AL'].wraptext := true;
    /// Перенос по словам
    Myworksheets.Columns['A:AK'].numberFormat := '@';
    /// текстовый формат
    Myworksheets.cells[1, 1] := 'Компьютер';
    Myworksheets.cells[1, 2] := 'Инв №';
    Myworksheets.cells[1, 3] := 'Дата инв...';
    Myworksheets.cells[1, 4] := 'M/Board';
    Myworksheets.cells[1, 5] := 'S/N M/Board';
    Myworksheets.cells[1, 6] := 'Производитель M/Board';

    Myworksheets.cells[1, 7] := 'Кол-во процессоров';
    Myworksheets.cells[1, 8] := 'Сокет процессора';
    Myworksheets.cells[1, 9] := 'Процессор';
    Myworksheets.cells[1, 10] := 'Ядер процессора';
    Myworksheets.cells[1, 11] := 'Логических процессоров';
    Myworksheets.cells[1, 12] := 'Частота процессора (Mz)';
    Myworksheets.cells[1, 13] := 'Архитектура процессора';

    Myworksheets.cells[1, 14] := 'Оперативная память (Mb)';
    Myworksheets.cells[1, 15] := 'Всего оперативной памяти (Mb)';
    Myworksheets.cells[1, 16] := 'Частота RAM (Mz)';
    Myworksheets.cells[1, 17] := 'Видеокарта';
    Myworksheets.cells[1, 18] := 'Память видеокарты (Mb)';
    Myworksheets.cells[1, 19] := 'Звук';
    Myworksheets.cells[1, 20] := 'Сетевой интерфейс';
    Myworksheets.cells[1, 21] := 'MAC адрес';
    Myworksheets.cells[1, 22] := 'HDD';
    Myworksheets.cells[1, 23] := 'S/N HDD';
    Myworksheets.cells[1, 24] := 'FirmWare HDD';
    Myworksheets.cells[1, 25] := 'Объем HDD';
    Myworksheets.cells[1, 26] := 'Кол-во HDD';
    Myworksheets.cells[1, 27] := 'CD/DVD Rom';
    Myworksheets.cells[1, 28] := 'Кол-во CD/DVD Rom';
    Myworksheets.cells[1, 29] := 'Монитор';
    Myworksheets.cells[1, 30] := 'Разрешение монитора';
    Myworksheets.cells[1, 31] := 'Кол-во мониторов';
    Myworksheets.cells[1, 32] := 'Операционная система';
    Myworksheets.cells[1, 33] := 'ОС версия';
    Myworksheets.cells[1, 34] := 'ОС SP';
    Myworksheets.cells[1, 35] := 'ОС KEY';
    Myworksheets.cells[1, 36] := 'ОС тип';
    Myworksheets.cells[1, 37] := 'Дата установки';
    Myworksheets.cells[1, 38] := 'Принтер';

    z := 2; // со второй строки начинаем вводить таблицу
    if (DBGrid1.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid1.SelectedRows.Count > 0) then
    /// / если выделенно больше 0 строк
      begin
        DBGrid1.DataSource.DataSet.DisableControls;
        for i := 0 to DBGrid1.SelectedRows.Count - 1 do
          begin
            DBGrid1.DataSource.DataSet.Bookmark :=
              DBGrid1.SelectedRows.Items[i];
            if string(DBGrid1.DataSource.DataSet.FieldByName('RESULT_INV')
              .Value) = 'OK' then
              begin
                StatusBarInv.Panels[2].text := inttostr(i + 1) + ' из ' +
                  inttostr(DBGrid1.SelectedRows.Count) + 'отчет - ' +
                  string(DBGrid1.DataSource.DataSet.FieldByName
                  ('PC_NAME').Value);
                FDQueryExcelExport.SQL.Clear;
                /// // чтение всех конфигураций  и передача DBgrid2
                FDQueryExcelExport.SQL.text :=
                  'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
                  (string(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME')
                  .Value)) + '''' + ' ORDER BY DATE_INV DESC';
                FDQueryExcelExport.Open;
                Myworksheets.cells[z, 1] :=(vartostr(FDQueryExcelExport.FieldByName('PC_NAME').Value));
                Myworksheets.cells[z, 2] :=(vartostr(FDQueryExcelExport.FieldByName('INV_NUMBER').Value));
                Myworksheets.cells[z, 3] :=(vartostr(FDQueryExcelExport.FieldByName('DATE_INV').Value));
                Myworksheets.cells[z, 4] :=(vartostr(FDQueryExcelExport.FieldByName('MONTERBOARD').Value));
                Myworksheets.cells[z, 5] :=(vartostr(FDQueryExcelExport.FieldByName('MONTERBOARD_SN').Value));
                Myworksheets.cells[z, 6] :=(vartostr(FDQueryExcelExport.FieldByName('MONTERBOARD_MANUFACTURE').Value));
                Myworksheets.cells[z, 7] :=(vartostr(FDQueryExcelExport.FieldByName('COUNT_PROC').Value));
                Myworksheets.cells[z, 8] :=delZ(vartostr(FDQueryExcelExport.FieldByName('PROCESSOR_SOKET').Value));
                Myworksheets.cells[z, 9] :=delZ(vartostr(FDQueryExcelExport.FieldByName ('PROCESSOR').Value));
                Myworksheets.cells[z, 10] :=delZ(vartostr(FDQueryExcelExport.FieldByName('PROCESSOR_CORE').Value));
                Myworksheets.cells[z, 11] := delZ(vartostr(FDQueryExcelExport.FieldByName('PROCESSOR_LOGPROC').Value));
                Myworksheets.cells[z, 12] :=delZ(vartostr(FDQueryExcelExport.FieldByName('PROCESSOR_SPEED').Value));
                Myworksheets.cells[z, 13] :=delZ(vartostr(FDQueryExcelExport.FieldByName('PROCESSOR_ARCH').Value));
                Myworksheets.cells[z, 14] :=delZ(vartostr(FDQueryExcelExport.FieldByName('MEMORY_SIZE').Value));
                Myworksheets.cells[z, 15] :=SummaDDR(kolitem(vartostr(FDQueryExcelExport.FieldByName('MEMORY_SIZE').Value)));  // всего оперативной памяти
                Myworksheets.cells[z, 16] :=delZ(vartostr(FDQueryExcelExport.FieldByName('MEMORY_TYPE').Value));
                Myworksheets.cells[z, 17] :=delZ(vartostr(FDQueryExcelExport.FieldByName('VIDEOCARD').Value));
                Myworksheets.cells[z, 18] :=delZ(vartostr(FDQueryExcelExport.FieldByName('VIDEOCARD_MEM').Value));
                Myworksheets.cells[z, 19] :=delZ(vartostr(FDQueryExcelExport.FieldByName('SOUND_NAME').Value));
                Myworksheets.cells[z, 20] :=delZ(vartostr(FDQueryExcelExport.FieldByName('NETWORKINTERFACE').Value));
                Myworksheets.cells[z, 21] :=delZ(vartostr(FDQueryExcelExport.FieldByName('NETWORK_MAC').Value));
                Myworksheets.cells[z, 22] :=delZ(vartostr(FDQueryExcelExport.FieldByName('HDD_NAME').Value));
                Myworksheets.cells[z, 23] :=delZ(vartostr(FDQueryExcelExport.FieldByName('HDD_SN').Value));
                Myworksheets.cells[z, 24] :=delZ(vartostr(FDQueryExcelExport.FieldByName('HDD_FIRMWARE').Value));
                Myworksheets.cells[z, 25] :=delZ(vartostr(FDQueryExcelExport.FieldByName('HDD_SIZE').Value));
                Myworksheets.cells[z, 26] :=(vartostr(FDQueryExcelExport.FieldByName('COUNT_HDD').Value));
                Myworksheets.cells[z, 27] :=delZ(vartostr(FDQueryExcelExport.FieldByName('DVDROM_NAME').Value));
                Myworksheets.cells[z, 28] :=(vartostr(FDQueryExcelExport.FieldByName('DVDROM_COUNT').Value));
                Myworksheets.cells[z, 29] :=delZ(vartostr(FDQueryExcelExport.FieldByName('MONITOR_NAME').Value));
                Myworksheets.cells[z, 30] :=delZ(vartostr(FDQueryExcelExport.FieldByName('MONITOR_HW').Value));
                Myworksheets.cells[z, 31] :=(vartostr(FDQueryExcelExport.FieldByName('MONITOR_COUNT').Value));
                Myworksheets.cells[z, 32] :=(vartostr(FDQueryExcelExport.FieldByName('OS_NAME').Value));
                Myworksheets.cells[z, 33] :=(vartostr(FDQueryExcelExport.FieldByName('OS_VER').Value));
                Myworksheets.cells[z, 34] :=(vartostr(FDQueryExcelExport.FieldByName('OS_SP').Value));
                Myworksheets.cells[z, 35] :=(vartostr(FDQueryExcelExport.FieldByName('OS_KEY').Value));
                Myworksheets.cells[z, 36] :=(vartostr(FDQueryExcelExport.FieldByName('OS_TYPE').Value));
                Myworksheets.cells[z, 37] :=(vartostr(FDQueryExcelExport.FieldByName('OS_DATEINSTALL').Value));
                Myworksheets.cells[z, 38] :=delZ(vartostr(FDQueryExcelExport.FieldByName('PRINTER_NAME').Value));
                Inc(z);
              end;
          end;
      end;
    DBGrid1.DataSource.DataSet.EnableControls;
    MyExcel.Visible := true; // показать Excel
    StatusBarInv.Panels[2].text := 'Экспорт завершен';
    saveexcel := Tsavedialog.Create(self);
    saveexcel.Title := 'Сохранение отчета';
    saveexcel.InitialDir := ExtractFilePath(Application.ExeName) + 'reports\';
    saveexcel.filename := 'Конфигурации ПК ' + datetostr(Now);
    saveexcel.Filter := 'Книга Excel 97-2003 (*.XLS)|*.XLS|' +
      'Книга Open XML (*.xlsx)|*.xlsx|Формат HTML (*.html)|*.html|Файл csv (*.csv)|*.csv|'
      + 'Электронная таблица OpenDocument (*.ods)|*.ods';
    saveexcel.FilterIndex := 1;;
    if saveexcel.Execute then
      begin
        case saveexcel.FilterIndex of
          1:
            begin
              saveexcel.DefaultExt := '.xls';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '56');
            end;
          2:
            begin
              saveexcel.DefaultExt := '.xlsx';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '51');
            end;
          3:
            begin
              saveexcel.DefaultExt := '.html';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '44');
            end;
          4:
            begin
              saveexcel.DefaultExt := '.csv';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '6');
            end;
          5:
            begin
              saveexcel.DefaultExt := '.ods';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '60');
            end;
        end;
        FreeAndNil(saveexcel);
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export Excel Config - ' + E.Message);
        VarClear(Myworksheets);
        VariantClear(MyExcel);
        if Assigned(saveexcel) then
          FreeAndNil(saveexcel);
      end;
  end;
  if Assigned(saveexcel) then
    FreeAndNil(saveexcel);
  VarClear(Myworksheets);
  VariantClear(MyExcel);
end;

procedure TfrmDomainInfo.SpeedButton23Click(Sender: TObject);
var
  i, z: Integer;
  saveexcel: Tsavedialog;
  Myworksheets: Variant;
  /// лист в ексель
begin
  try
    if DBGrid2.DataSource.DataSet.RecordCount = 0 then
      exit; // если нет записей, то нехуй делать
    if (DBGrid2.SelectedRows.Count < 1) then
    // если записи не выделены, то выделяем
      begin
        // showmessage('Необходимо выделить конфигурацию');
        GridSelectAll(DBGrid2);
        // exit;
      end;
    AllRunExcel;
    /// Полный запуск Excel
    // myExcel.visible:=true; // показать Excel
    MyExcel.WorkBooks[1].WorkSheets[1].Name := 'Инвентаризация';
    /// имя листа
    Myworksheets := MyExcel.WorkBooks[1].WorkSheets[1];
    // MyWorkSheets.range['B1:D1'].Columns.AutoFit; /// авто ширина ячеек
    // myworksheets.columns['E:F'].EntireColumn.AutoFit;  /// авто ширина колонок
    Myworksheets.Columns['A:AK'].columnwidth := 20; // ширина колонки
    Myworksheets.Columns['A:AK'].wraptext := true;
    /// Перенос по словам
    Myworksheets.Columns['A:AK'].numberFormat := '@';
    /// текстовый формат
    Myworksheets.cells[1, 1] := 'Компьютер';
    Myworksheets.cells[1, 2] := 'Инв №';
    Myworksheets.cells[1, 3] := 'Дата инв...';
    Myworksheets.cells[1, 4] := 'M/Board';
    Myworksheets.cells[1, 5] := 'S/N M/Board';
    Myworksheets.cells[1, 6] := 'Производитель M/Board';
    Myworksheets.cells[1, 7] := 'Процессор';
    Myworksheets.cells[1, 8] := 'Ядер процессора';
    Myworksheets.cells[1, 9] := 'Логических процессоров';
    Myworksheets.cells[1, 10] := 'Частота процессора (Mz)';
    Myworksheets.cells[1, 11] := 'Архитектура процессора';
    Myworksheets.cells[1, 12] := 'Сокет процессора';
    Myworksheets.cells[1, 13] := 'Кол-во процессоров';
    Myworksheets.cells[1, 14] := 'Оперативная память (Mb)';
    Myworksheets.cells[1, 15] := 'Частота RAM (Mz)';
    Myworksheets.cells[1, 16] := 'Видеокарта';
    Myworksheets.cells[1, 17] := 'Память видеокарты (Mb)';
    Myworksheets.cells[1, 18] := 'Звук';
    Myworksheets.cells[1, 19] := 'Сетевой интерфейс';
    Myworksheets.cells[1, 20] := 'MAC адрес';
    Myworksheets.cells[1, 21] := 'HDD';
    Myworksheets.cells[1, 22] := 'S/N HDD';
    Myworksheets.cells[1, 23] := 'FirmWare HDD';
    Myworksheets.cells[1, 24] := 'Объем HDD';
    Myworksheets.cells[1, 25] := 'Кол-во HDD';
    Myworksheets.cells[1, 26] := 'CD/DVD Rom';
    Myworksheets.cells[1, 27] := 'Кол-во CD/DVD Rom';
    Myworksheets.cells[1, 28] := 'Монитор';
    Myworksheets.cells[1, 29] := 'Разрешение монитора';
    Myworksheets.cells[1, 30] := 'Кол-во мониторов';
    Myworksheets.cells[1, 31] := 'Операционная система';
    Myworksheets.cells[1, 32] := 'ОС версия';
    Myworksheets.cells[1, 33] := 'ОС SP';
    Myworksheets.cells[1, 34] := 'ОС KEY';
    Myworksheets.cells[1, 35] := 'ОС тип';
    Myworksheets.cells[1, 36] := 'Дата установки';
    Myworksheets.cells[1, 37] := 'Принтер';
    z := 2; // со второй строки начинаем вводить таблицу
    if (DBGrid2.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid2.SelectedRows.Count > 0) then
    /// / если выделенно больше 0 строк
      begin
        for i := 0 to DBGrid2.SelectedRows.Count - 1 do
          begin
            DBGrid2.DataSource.DataSet.Bookmark :=
              DBGrid2.SelectedRows.Items[i];
            begin
              StatusBarInv.Panels[2].text := 'отчет-' +
                string(DBGrid1.DataSource.DataSet.FieldByName('PC_NAME').Value);
              FDQueryExcelExport.SQL.Clear;
              /// // чтение всех конфигураций  и передача DBgrid2
              FDQueryExcelExport.SQL.text :=
                'SELECT * FROM CONFIG_PC WHERE PC_ID=''' +
                (string(DBGrid2.DataSource.DataSet.FieldByName('PC_ID')
                .Value)) + '''';
              FDQueryExcelExport.Open;
              Myworksheets.cells[z, 1] :=
                (vartostr(FDQueryExcelExport.FieldByName('PC_NAME').Value));
              Myworksheets.cells[z, 2] :=
                (vartostr(FDQueryExcelExport.FieldByName('INV_NUMBER').Value));
              Myworksheets.cells[z, 3] :=
                (vartostr(FDQueryExcelExport.FieldByName('DATE_INV').Value));
              Myworksheets.cells[z, 4] :=
                (vartostr(FDQueryExcelExport.FieldByName('MONTERBOARD').Value));
              Myworksheets.cells[z, 5] :=
                (vartostr(FDQueryExcelExport.FieldByName
                ('MONTERBOARD_SN').Value));
              Myworksheets.cells[z, 6] :=
                (vartostr(FDQueryExcelExport.FieldByName
                ('MONTERBOARD_MANUFACTURE').Value));
              Myworksheets.cells[z, 7] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR').Value));
              Myworksheets.cells[z, 8] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR_CORE').Value));
              Myworksheets.cells[z, 9] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR_LOGPROC').Value));
              Myworksheets.cells[z, 10] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR_SPEED').Value));
              Myworksheets.cells[z, 11] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR_ARCH').Value));
              Myworksheets.cells[z, 12] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PROCESSOR_SOKET').Value));
              Myworksheets.cells[z, 13] :=
                (vartostr(FDQueryExcelExport.FieldByName('COUNT_PROC').Value));
              Myworksheets.cells[z, 14] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('MEMORY_SIZE').Value));
              Myworksheets.cells[z, 15] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('MEMORY_TYPE').Value));
              Myworksheets.cells[z, 16] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('VIDEOCARD').Value));
              Myworksheets.cells[z, 17] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('VIDEOCARD_MEM').Value));
              Myworksheets.cells[z, 18] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('SOUND_NAME').Value));
              Myworksheets.cells[z, 19] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('NETWORKINTERFACE').Value));
              Myworksheets.cells[z, 20] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('NETWORK_MAC').Value));
              Myworksheets.cells[z, 21] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('HDD_NAME').Value));
              Myworksheets.cells[z, 22] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName('HDD_SN').Value));
              Myworksheets.cells[z, 23] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('HDD_FIRMWARE').Value));
              Myworksheets.cells[z, 24] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('HDD_SIZE').Value));
              Myworksheets.cells[z, 25] :=
                (vartostr(FDQueryExcelExport.FieldByName('COUNT_HDD').Value));
              Myworksheets.cells[z, 26] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('DVDROM_NAME').Value));
              Myworksheets.cells[z, 27] :=
                (vartostr(FDQueryExcelExport.FieldByName
                ('DVDROM_COUNT').Value));
              Myworksheets.cells[z, 28] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('MONITOR_NAME').Value));
              Myworksheets.cells[z, 29] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('MONITOR_HW').Value));
              Myworksheets.cells[z, 30] :=
                (vartostr(FDQueryExcelExport.FieldByName
                ('MONITOR_COUNT').Value));
              Myworksheets.cells[z, 31] :=
                (vartostr(FDQueryExcelExport.FieldByName('OS_NAME').Value));
              Myworksheets.cells[z, 32] :=
                (vartostr(FDQueryExcelExport.FieldByName('OS_VER').Value));
              Myworksheets.cells[z, 33] :=
                (vartostr(FDQueryExcelExport.FieldByName('OS_SP').Value));
              Myworksheets.cells[z, 34] :=
                (vartostr(FDQueryExcelExport.FieldByName('OS_KEY').Value));
              Myworksheets.cells[z, 35] :=
                (vartostr(FDQueryExcelExport.FieldByName('OS_TYPE').Value));
              Myworksheets.cells[z, 36] :=
                (vartostr(FDQueryExcelExport.FieldByName
                ('OS_DATEINSTALL').Value));
              Myworksheets.cells[z, 37] :=
                delZ(vartostr(FDQueryExcelExport.FieldByName
                ('PRINTER_NAME').Value));
              Inc(z);
            end;
          end;
      end;

    MyExcel.Visible := true; // показать Excel
    saveexcel := Tsavedialog.Create(self);
    saveexcel.Title := 'Сохранение отчета';
    saveexcel.InitialDir := ExtractFilePath(Application.ExeName) + 'reports\';
    saveexcel.filename := 'Конфигурация ПК ' + datetostr(Now);
    saveexcel.Filter := 'Книга Excel 97-2003 (*.XLS)|*.XLS|' +
      'Книга Open XML (*.xlsx)|*.xlsx|Формат HTML (*.html)|*.html|Файл csv (*.csv)|*.csv|'
      + 'Электронная таблица OpenDocument (*.ods)|*.ods';
    saveexcel.FilterIndex := 1;;
    if saveexcel.Execute then
      begin
        case saveexcel.FilterIndex of
          1:
            begin
              saveexcel.DefaultExt := '.xls';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '56');
            end;
          2:
            begin
              saveexcel.DefaultExt := '.xlsx';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '51');
            end;
          3:
            begin
              saveexcel.DefaultExt := '.html';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '44');
            end;
          4:
            begin
              saveexcel.DefaultExt := '.csv';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '6');
            end;
          5:
            begin
              saveexcel.DefaultExt := '.ods';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '60');
            end;
        end;
      end;
    StatusBarInv.Panels[2].text := '';
    FreeAndNil(saveexcel);
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export Excel PC in table CONFIG_PC sort - ' + E.Message);
        VarClear(Myworksheets);
        VariantClear(MyExcel);
      end;
  end;
  if Assigned(saveexcel) then
    FreeAndNil(saveexcel);
  VarClear(Myworksheets);
  VariantClear(MyExcel);
end;

procedure TfrmDomainInfo.SpeedButton24Click(Sender: TObject);
var
  i, Res: Integer;
  /// / инвентаризация программ
begin
  try
    if ListView8.Items.Count > 0 then
      begin
        if not SolveExitInvSoft then
          begin
            showmessage('Дождитесь завершения текущей инвентаризации...');
            exit;
          end;

        ListPCConf := TstringList.Create;
        for i := 0 to ListView8.Items.Count - 1 do
          begin
            if ListView8.Items[i].Checked = true then
              ListPCConf.Add(ListView8.Items[i].SubItems[0])
          end;

        if ListPCConf.Count = 0 then
          begin
            if ListView8.Items.Count = 0 then
              begin
                MainPage.ActivePageIndex := 0;
                showmessage('Нет списка компьютеров');
                ListPCConf.Free;
                exit;
              end;
            ListPCConf := createListpc('');
          end;
        Memo1.Lines.Add('Запускаю инвентаризацию программного обеспечения');
        InventSoft := true;
        MyUser := LabeledEdit1.text;
        MyPasswd := LabeledEdit2.text;
        ThrInvSoft := MyInventorySoft.InventorySoft.Create(true);
        ThrInvSoft.FreeOnTerminate := true;
        ThrInvSoft.Start;
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Start invent soft - ' + E.Message)
      end;
  end;
end;

/// ////////////////////////////////////////////////////////////////////////////////////////
/// ////////////////////////////////////////////////////////////////////////////////////////

/// ////////////////////////////////////////////////////////////////////////////////////////////
/// ///////////////////////Инвентаризация программ////////////////////////////////////////
procedure TfrmDomainInfo.ReadDBSoft;
/// чтение из базы данных
begin
  try
    begin
    /// //  после сортировки по полю подсчет выдает ошибку
      DBGrid3.DataSource.DataSet.DisableControls;
      Datam.FDQueryReadSoft.DeleteIndexes;
      /// обнуляем индексы
      Datam.FDQueryReadSoft.IndexFieldNames := '';
      /// обнуляем индекс если была сортировка в Dbgrid1
      Datam.FDQueryReadSoft.SQL.Clear;
      Datam.FDQueryReadSoft.SQL.text :=
        'select count(PC_ID) as countPC from main_pc_soft';
      Datam.FDQueryReadSoft.Open;
      StatusBarSoft.Panels[0].text := 'Всего записей - ' +
        (Datam.FDQueryReadSoft.FieldByName('countPC').AsString);
      Datam.FDQueryReadSoft.SQL.Clear;
      Datam.FDQueryReadSoft.SQL.text :=
        'select count(PC_ID) as countPC from main_pc_soft where resul_inv=:pStat';
      Datam.FDQueryReadSoft.ParamByName('pStat').Value := 'OK';
      Datam.FDQueryReadSoft.Open;
      StatusBarSoft.Panels[1].text := 'OK - ' +
        (Datam.FDQueryReadSoft.FieldByName('countPC').AsString);
    end;
    Datam.FDQueryReadSoft.DeleteIndexes;
    Datam.FDQueryReadSoft.SQL.Clear;
    Datam.FDQueryReadSoft.SQL.text :=
      'SELECT * FROM MAIN_PC_SOFT ORDER BY PC_NAME';
    Datam.FDQueryReadSoft.Open;
    // statusbarInv.Panels[0].Text:='Всего записей - '+inttostr(DBgrid1.DataSource.DataSet.RecordCount);
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Update pc soft - ' + E.Message)
      end;
  end;
  DBGrid3.DataSource.DataSet.EnableControls;
end;

procedure TfrmDomainInfo.SpeedButton25Click(Sender: TObject);
begin
  ReadDBSoft;
end;

procedure TfrmDomainInfo.SpeedButton26Click(Sender: TObject);
begin
/// поиск компьютера в Soft
  try
    if DBGrid3.DataSource.DataSet.RecordCount <> 0 then
      begin
        if not DBGrid3.DataSource.DataSet.Locate('PC_NAME', FindPCSoft.text,
          [loCaseInsensitive, loPartialKey]) then
          showmessage('Ничего не найдено!!!');
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find pc soft - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton27Click(Sender: TObject);
var
/// // удаление записей из таблицы с софтом
  i: Integer;
begin
  try
    if (DBGrid3.SelectedRows.Count > 0) and
      (DBGrid3.DataSource.DataSet.RecordCount <> 0) then
    /// / если выделенно больше 0 строк
      begin
        i := MessageDlg('Кол-во записей для удаления - ' +
          inttostr(DBGrid3.SelectedRows.Count) + #10#13 + 'Удалить?',
          mtConfirmation, [mbYes, mbCancel], 0);
        if i = IDYes then
          begin
            DBGrid3.DataSource.DataSet.DisableControls;
            for i := 0 to DBGrid3.SelectedRows.Count - 1 do
              begin
                DBGrid3.DataSource.DataSet.Bookmark :=
                  DBGrid3.SelectedRows.Items[i];
                FDTransactionWrite1.Options.AutoCommit := true;
                FDTransactionWrite1.Options.AutoStart := true;
                FDTransactionWrite1.Options.AutoStop := true;
                FDQueryDelete.Active := false;
                FDQueryDelete.SQL.Clear;
                FDQueryDelete.SQL.text :=
                  'delete FROM MAIN_PC_SOFT WHERE PC_ID=''' +
                  vartostr(DBGrid3.DataSource.DataSet.FieldByName('PC_ID')
                  .Value) + '''';
                FDQueryDelete.ExecSQL;
                FDTransactionWrite1.Options.AutoCommit := false;
                FDTransactionWrite1.Options.AutoStart := false;
                FDTransactionWrite1.Options.AutoStop := false;
              end;
          end;
        DBGrid3.DataSource.DataSet.EnableControls;
        DBGrid3.DataSource.DataSet.Close;
        /// закрыть
        DBGrid3.DataSource.DataSet.Open;
        /// /  открыть для обновления грида

      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Delete pc invent soft - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton28Click(Sender: TObject);
var
  i: byte;
begin
/// остановка инвентаризации софта
  try

    if not SolveExitInvSoft then
    /// если запущена инвентаризация то задаем вопрос и останавливаем
      begin
        i := MessageDlg('Остановить инвентаризацию? ', mtConfirmation,
          [mbYes, mbCancel], 0);
        if i = IDCancel then
          exit;
        InventSoft := false;
        /// Остановка инвентаризации
        Memo1.Lines.Add
          ('Остановка инвентаризации программного обеспечения... ожидайте');
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Stop invent soft - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton29Click(Sender: TObject);
var
/// / экспорт softa в Excel
  i, X: Integer;
  saveexcel: Tsavedialog;
  IDSoft: TStringArray;
  Myworksheets: Variant;
  /// лист в ексель
  ProgressExport: TprogressBar;
begin
  try
    if (DBGrid3.SelectedRows.Count <= 0) then
      begin
        showmessage('Необходимо выделить список компьютеров');
        exit;
      end;
    if InventSoft then
      begin
        showmessage('Дождитесь завершения текущей инвентаризации!');
        exit;
      end;

    AllRunExcel;
    /// / запуск Excel
    if (DBGrid3.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid3.SelectedRows.Count > 0) then
    /// / если выделенно больше 0 строк
      begin
        DBGrid3.DataSource.DataSet.DisableControls;
        /// чтобы не мигал DBgrid
        /// //////////////////////////////////////////////////////////////
        if not Assigned(ProgressExport) then
          begin
            ProgressExport := TprogressBar.Create(self);
            ProgressExport.Parent := StatusBarSoft;
            ProgressExport.Name := 'ProgSoft';
          end;
        ProgressExport.left := StatusBarSoft.Panels[0].width +
          StatusBarSoft.Panels[1].width + StatusBarSoft.Panels[2].width;
        ProgressExport.width := StatusBarSoft.Panels[3].width;
        ProgressExport.Min := 0;
        ProgressExport.Max := DBGrid3.SelectedRows.Count - 1;
        ProgressExport.Step := 1;

        /// ////////////////////////////////////////////////////////////////
        for i := 0 to DBGrid3.SelectedRows.Count - 1 do
          begin
            DBGrid3.DataSource.DataSet.Bookmark :=
              DBGrid3.SelectedRows.Items[i];
            MyExcel.WorkBooks[1].WorkSheets[1].Name :=
              string(DBGrid3.DataSource.DataSet.FieldByName('PC_NAME').Value);
            /// имя лист
            Myworksheets := MyExcel.WorkBooks[1].WorkSheets[1];
            Myworksheets.Columns['A'].columnwidth := 7; // ширина колонки
            Myworksheets.Columns['B'].columnwidth := 70;
            Myworksheets.Columns['C'].columnwidth := 15;
            Myworksheets.Columns['D'].columnwidth := 40;
            Myworksheets.Columns['A:AD'].wraptext := true;
            /// Перенос по словам
            Myworksheets.Columns['A:AD'].numberFormat := '@';
            /// текстовый формат
            Myworksheets.cells[1, 1] := '№';
            Myworksheets.cells[1, 2] := 'Имя';
            Myworksheets.cells[1, 3] := 'Версия';
            Myworksheets.cells[1, 4] := 'Издатель';
            if (string(DBGrid3.DataSource.DataSet.FieldByName('RESUL_INV')
              .Value) = 'OK') and
              (Datam.FDQueryReadSoft.FieldByName('INST_SOFT').Value <> null)
            then
              begin
                if Assigned(ProgressExport) then
                  ProgressExport.Position := i;
                StatusBarSoft.Panels[2].text := inttostr(i + 1) + ' из ' +
                  inttostr(DBGrid3.SelectedRows.Count) + ' Отчет-' +
                  string(DBGrid3.DataSource.DataSet.FieldByName
                  ('PC_NAME').Value);
                IDSoft := kolitemSoft
                  (vartostr(Datam.FDQueryReadSoft.FieldByName
                  ('INST_SOFT').Value));
                for X := 0 to Length(IDSoft) - 1 do
                  begin
                    FDQueryExcelExport.SQL.Clear;
                    /// // чтение всех конфигураций  и передача DBgrid2
                    FDQueryExcelExport.SQL.text :=
                      'SELECT * FROM SOFT_PC WHERE ID=''' + IDSoft[X] + '''';
                    FDQueryExcelExport.Open;
                    Myworksheets.cells[X + 2, 1] := inttostr(X + 1);
                    Myworksheets.cells[X + 2, 2] :=
                      (vartostr(FDQueryExcelExport.FieldByName
                      ('SOFT_NAME').Value));
                    Myworksheets.cells[X + 2, 3] :=
                      (vartostr(FDQueryExcelExport.FieldByName
                      ('SOFT_VERSION').Value));
                    Myworksheets.cells[X + 2, 4] :=
                      (vartostr(FDQueryExcelExport.FieldByName
                      ('MANUFACTURE').Value));
                  end;
                if i <> DBGrid3.SelectedRows.Count - 1 then
                /// если итерация в цикле не последняя то создаем лист
                  begin
                    MyExcel.WorkBooks[1].WorkSheets.Add(EmptyParam, EmptyParam,
                      EmptyParam);
                    /// вставка листа в начало
                    MyExcel.ActiveWorkBook.Sheets.Item[1].Activate;
                    /// делаем лист активным
                  end;
              end;
          end;
      end;
    if Assigned(ProgressExport) then
      FreeAndNil(ProgressExport);
    DBGrid3.DataSource.DataSet.EnableControls;
    /// чтобы не мигал DBgrid
    MyExcel.Visible := true; // показать Excel
    StatusBarSoft.Panels[2].text := 'Экспорт завершен';
    saveexcel := Tsavedialog.Create(self);
    saveexcel.Title := 'Сохранение отчета';
    saveexcel.InitialDir := ExtractFilePath(Application.ExeName) + 'reports\';
    saveexcel.filename := 'Программное обеспечение ПК ' + datetostr(Now);
    saveexcel.Filter := 'Книга Excel 97-2003 (*.XLS)|*.XLS|' +
      'Книга Open XML (*.xlsx)|*.xlsx|Формат HTML (*.html)|*.html|Файл csv (*.csv)|*.csv|'
      + 'Электронная таблица OpenDocument (*.ods)|*.ods';
    saveexcel.FilterIndex := 1;
    if saveexcel.Execute then
      begin
        case saveexcel.FilterIndex of
          1:
            begin
              saveexcel.DefaultExt := '.xls';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '56');
            end;
          2:
            begin
              saveexcel.DefaultExt := '.xlsx';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '51');
            end;
          3:
            begin
              saveexcel.DefaultExt := '.html';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '44');
            end;
          4:
            begin
              saveexcel.DefaultExt := '.csv';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '6');
            end;
          5:
            begin
              saveexcel.DefaultExt := '.ods';
              SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt, 1, '60');
            end;
        end;
      end;
    StatusBarInv.Panels[2].text := '';
    FreeAndNil(saveexcel);
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export invent soft - ' + E.Message);
        VarClear(Myworksheets);
        VariantClear(MyExcel);
        if Assigned(saveexcel) then
          FreeAndNil(saveexcel);
        if Assigned(ProgressExport) then
          FreeAndNil(ProgressExport);
      end;
  end;
  if Assigned(saveexcel) then
    FreeAndNil(saveexcel);
  VarClear(Myworksheets);
  VariantClear(MyExcel);
end;

procedure TfrmDomainInfo.DBGrid3CllClick(Column: TColumn);
/// // чтение софта выбранного компьютера и передача listview
var
  i: Integer;
  IDSoft: TStringArray;
begin
  try
    if (DBGrid3.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid3.SelectedRows.Count = 1) then
    /// если DBgrid не пустой то выбираем из таблицы конфигурацию
      begin
        DBGrid3.DataSource.DataSet.Bookmark := DBGrid3.SelectedRows.Items[0];
        IDSoft := kolitemSoft
          (vartostr(Datam.FDQueryReadSoft.FieldByName('INST_SOFT').Value));
        ListViewSoft.Clear;
        GroupBox13.Caption := 'Список программ - ' +
          vartostr(Datam.FDQueryReadSoft.FieldByName('PC_NAME').Value);
        for i := 0 to Length(IDSoft) - 1 do
          begin
            Datam.FDSelectReadSoft.SQL.Clear;
            Datam.FDSelectReadSoft.SQL.text :=
              'SELECT * FROM SOFT_PC WHERE ID=''' + IDSoft[i] + '''';
            Datam.FDSelectReadSoft.Open;
            With ListViewSoft.Items.Add do
              begin
                Caption := inttostr(ListViewSoft.Items.Count);
                if Datam.FDSelectReadSoft.FieldByName('SOFT_NAME').Value <> null
                then
                  SubItems.Add
                    (string(Datam.FDSelectReadSoft.FieldByName
                    ('SOFT_NAME').Value))
                else
                  SubItems.Add('');
                if Datam.FDSelectReadSoft.FieldByName('MANUFACTURE').Value <> null
                then
                  SubItems.Add
                    (string(Datam.FDSelectReadSoft.FieldByName
                    ('MANUFACTURE').Value))
                else
                  SubItems.Add('');
                if Datam.FDSelectReadSoft.FieldByName('SOFT_VERSION').Value <> null
                then
                  SubItems.Add
                    (string(Datam.FDSelectReadSoft.FieldByName
                    ('SOFT_VERSION').Value))
                else
                  SubItems.Add('');
              end;
          end;
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз:' + E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' read invent soft - ' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.DBGrid3TitleClick(Column: TColumn);
var
/// // сортировка  DBgrid3
  f: string;
begin
  try
    if DBGrid3.DataSource.DataSet.RecordCount <> 0 then
      begin
        f := Column.FieldName;
        if DBGrid3.Tag = 0 then
          begin
            Datam.FDQueryReadSoft.IndexFieldNames := f + ':A;';
            DBGrid3.Tag := 1;
          end
        else
          begin
            Datam.FDQueryReadSoft.IndexFieldNames := f + ':D;';
            DBGrid3.Tag := 0;
          end;
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз:' + E.Message);
      end;
  end;
end;

/// ///////////////////////////////// выборки  ///////////////////////////////////////////////
/// //////////////////////////////////////////////////////////////////////////////////////
function TfrmDomainInfo.transactionSort(s: Boolean): bool;
begin
  Datam.FDTransactionSort.Options.AutoCommit := s;
  Datam.FDTransactionSort.Options.AutoStart := s;
  Datam.FDTransactionSort.Options.AutoStop := s;
end;

procedure TfrmDomainInfo.TrayIcon1Animate(Sender: TObject);
begin
  if (InventConf) or (InventSoft) then
    begin
      TrayIcon1.BalloonFlags := bfInfo;
      TrayIcon1.Hint := 'Включена инвентаризация';
    end
  else
    begin
      TrayIcon1.Iconindex := 0;
      TrayIcon1.Animate := false;
      TrayIcon1.BalloonFlags := bfInfo;
      TrayIcon1.Hint := 'Management Remote PC 3.0';
      TrayIcon1.BalloonTitle := 'Management Remote PC 3.0';
      TrayIcon1.BalloonHint := 'Инвентаризация завершена';
      TrayIcon1.ShowBalloonHint;
    end;
end;

procedure TfrmDomainInfo.TrayIcon1DblClick(Sender: TObject);
begin
  // ShowWindow(Handle,SW_RESTORE);
  ShowWindow(Application.Handle, SW_RESTORE);
  TrayIcon1.Visible := false;
  SetForegroundWindow(Handle);
end;

function TfrmDomainInfo.DBgrid4Columns(NameFields, CaptionColumns,
  NameTable: string): bool;
begin
  if NameTable = 'CONFIG_PC' then
    begin
    /// если выбоорка из таблицы с конфигурацией  , создаем колонки
      DBGrid4.Columns.Clear;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[0].FieldName := 'PC_NAME';
      DBGrid4.Columns[0].ImeName := 'PC_NAME';
      DBGrid4.Columns[0].Title.Caption := 'Имя компьютера';
      DBGrid4.Columns[0].width := 150;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[1].FieldName := NameFields;
      DBGrid4.Columns[1].ImeName := NameFields;
      DBGrid4.Columns[1].Title.Caption := CaptionColumns;
      DBGrid4.Columns[1].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[2].FieldName := 'MONTERBOARD';
      DBGrid4.Columns[2].ImeName := 'MONTERBOARD';
      DBGrid4.Columns[2].Title.Caption := 'Материнская плата';
      DBGrid4.Columns[2].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[3].FieldName := 'PROCESSOR';
      DBGrid4.Columns[3].ImeName := 'PROCESSOR';
      DBGrid4.Columns[3].Title.Caption := 'Процессор';
      DBGrid4.Columns[3].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[4].FieldName := 'MEMORY_SIZE';
      DBGrid4.Columns[4].ImeName := 'MEMORY_SIZE';
      DBGrid4.Columns[4].Title.Caption := 'RAM';
      DBGrid4.Columns[4].width := 75;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[5].FieldName := 'VIDEOCARD';
      DBGrid4.Columns[5].ImeName := 'VIDEOCARD';
      DBGrid4.Columns[5].Title.Caption := 'Видео карта';
      DBGrid4.Columns[5].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[6].FieldName := 'HDD_NAME';
      DBGrid4.Columns[6].ImeName := 'HDD_NAME';
      DBGrid4.Columns[6].Title.Caption := 'Жесткий диск';
      DBGrid4.Columns[6].width := 150;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[7].FieldName := 'NETWORKINTERFACE';
      DBGrid4.Columns[7].ImeName := 'NETWORKINTERFACE';
      DBGrid4.Columns[7].Title.Caption := 'Сетевой интерфейс';
      DBGrid4.Columns[7].width := 200;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[8].FieldName := 'MONITOR_NAME';
      DBGrid4.Columns[8].ImeName := 'MONITOR_NAME';
      DBGrid4.Columns[8].Title.Caption := 'Монитор';
      DBGrid4.Columns[8].width := 200;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[9].FieldName := 'DVDROM_NAME';
      DBGrid4.Columns[9].ImeName := 'DVDROM_NAME';
      DBGrid4.Columns[9].Title.Caption := 'CD/DVD-ROM';
      DBGrid4.Columns[9].width := 200;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[10].FieldName := 'SOUND_NAME';
      DBGrid4.Columns[10].ImeName := 'SOUND_NAME';
      DBGrid4.Columns[10].Title.Caption := 'Звук';
      DBGrid4.Columns[10].width := 200;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[11].FieldName := 'PRINTER_NAME';
      DBGrid4.Columns[11].ImeName := 'PRINTER_NAME';
      DBGrid4.Columns[11].Title.Caption := 'Принтер';
      DBGrid4.Columns[11].width := 200;
      exit;
    end;
  if NameTable = 'SOFT_PC' then
    begin
      DBGrid4.Columns.Clear;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[0].FieldName := 'PC_NAME';
      DBGrid4.Columns[0].ImeName := 'PC_NAME';
      DBGrid4.Columns[0].Title.Caption := 'Имя компьютера';
      DBGrid4.Columns[0].width := 150;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[1].FieldName := 'DATE_INV';
      DBGrid4.Columns[1].ImeName := 'DATE_INV';
      DBGrid4.Columns[1].Title.Caption := 'Дата инв...';
      DBGrid4.Columns[1].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[2].FieldName := 'TIME_INV';
      DBGrid4.Columns[2].ImeName := 'TIME_INV';
      DBGrid4.Columns[2].Title.Caption := 'Время инв...';
      DBGrid4.Columns[2].width := 100;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[3].FieldName := 'COUNT_SOFT';
      DBGrid4.Columns[3].ImeName := 'COUNT_SOFT';
      DBGrid4.Columns[3].Title.Caption := 'Кол-во программ';
      DBGrid4.Columns[3].width := 100;
      exit;
    end;
  if NameTable = 'ALL_SOFT' then
    begin
      DBGrid4.Columns.Clear;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[0].FieldName := 'SOFT_NAME';
      DBGrid4.Columns[0].ImeName := 'SOFT_NAME';
      DBGrid4.Columns[0].Title.Caption := 'Программа';
      DBGrid4.Columns[0].width := 350;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[1].FieldName := 'SOFT_VERSION';
      DBGrid4.Columns[1].ImeName := 'SOFT_VERSION';
      DBGrid4.Columns[1].Title.Caption := 'Версия';
      DBGrid4.Columns[1].width := 150;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[2].FieldName := 'MANUFACTURE';
      DBGrid4.Columns[2].ImeName := 'MANUFACTURE';
      DBGrid4.Columns[2].Title.Caption := 'Издатель';
      DBGrid4.Columns[2].width := 250;
      exit;
    end;
  if NameTable = 'ALL_HARDWARE' then
    begin
      DBGrid4.Columns.Clear;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[0].FieldName := 'NAME_HARDWARE';
      DBGrid4.Columns[0].ImeName := 'NAME_HARDWARE';
      DBGrid4.Columns[0].Title.Caption := 'Модель';
      DBGrid4.Columns[0].width := 350;
      DBGrid4.Columns.Add;
      DBGrid4.Columns[1].FieldName := 'TYPE_HARDWARE';
      DBGrid4.Columns[1].ImeName := 'TYPE_HARDWARE';
      DBGrid4.Columns[1].Title.Caption := 'Тип';
      DBGrid4.Columns[1].width := 100;
    end;

  if NameTable = 'WIN_OFF' then
    Begin
      DBGrid4.Columns.Clear;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'DESCRIPTION';
          ImeName := 'DESCRIPTION';
          Title.Caption := 'Имя продукта';
          width := 350;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'KEY_PRODUCT';
          ImeName := 'KEY_PRODUCT';
          Title.Caption := 'Ключ продукта';
          width := 100;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'TYPE_LIC';
          ImeName := 'TYPE_LIC';
          Title.Caption := 'Тип лицензии';
          width := 100;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'PRODUCT';
          ImeName := 'PRODUCT';
          Title.Caption := 'Продукт';
          width := 100;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'PARTIALPRODUCTKEY';
          ImeName := 'PARTIALPRODUCTKEY';
          Title.Caption := 'Код продукта';
          width := 250;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'PRODUCT_ID';
          ImeName := 'PRODUCT_ID';
          Title.Caption := 'Идентификатор';
          width := 250;
        end;
    End;

  if NameTable = 'MIC_LIC' then
    Begin
      DBGrid4.Columns.Clear;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'NAMEPC';
          ImeName := 'NAMEPC';
          Title.Caption := 'Компьютер';
          width := 120;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'WINPRODUCT';
          ImeName := 'WINPRODUCT';
          Title.Caption := 'Windows';
          width := 250;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'STATWINLIC';
          ImeName := 'STATWINLIC';
          Title.Caption := 'Статус лицензии';
          width := 200;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'KEYWIN';
          ImeName := 'KEYWIN';
          Title.Caption := 'Ключ';
          width := 80;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'TYPE_LIC_WIN';
          ImeName := 'TYPE_LIC_WIN';
          Title.Caption := 'Тип лицензии';
          width := 150;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'DESCRIPT_LIC_WIN';
          ImeName := 'DESCRIPT_LIC_WIN';
          Title.Caption := 'Описание статуса';
          width := 450;
        end;
    End;

  if NameTable = 'OFF_LIC' then
    Begin
      DBGrid4.Columns.Clear;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'NAMEPC';
          ImeName := 'NAMEPC';
          Title.Caption := 'Компьютер';
          width := 120;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'SOFFICEPRODUCT';
          ImeName := 'SOFFICEPRODUCT';
          Title.Caption := 'Продукты Office';
          width := 250;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'SSTATOFLIC';
          ImeName := 'SSTATOFLIC';
          Title.Caption := 'Статус лицензий';
          width := 200;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'SKEYOFC';
          ImeName := 'SKEYOFC';
          Title.Caption := 'Ключи';
          width := 80;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'TYPE_LIC_OFFICE';
          ImeName := 'TYPE_LIC_OFFICE';
          Title.Caption := 'Тип лицензий';
          width := 150;
        end;
      with DBGrid4.Columns.Add do
        begin
          FieldName := 'DESCRIPT_LIC_OFFICE';
          ImeName := 'DESCRIPT_LIC_OFFICE';
          Title.Caption := 'Описание статуса';
          width := 450;
        end;
    End;

end;

procedure TfrmDomainInfo.DBGrid4TitleClick(Column: TColumn);
var
/// // сортировка  DBgrid4
  f: string;
begin
  if DBGrid4.DataSource.DataSet.RecordCount <> 0 then
    begin
      f := Column.FieldName;
      if DBGrid4.Tag = 0 then
        begin
          Datam.FDQuerySelectSort.IndexFieldNames := f + ':A;';
          DBGrid4.Tag := 1;
        end
      else
        begin
          Datam.FDQuerySelectSort.IndexFieldNames := f + ':D;';
          DBGrid4.Tag := 0;
        end;
    end;
end;

procedure TfrmDomainInfo.ComboBox4Select(Sender: TObject);
begin
/// выбор из списка оборудования и передача списков в ComboBox3
  try
    Label4.Caption := ComboBox4.text;
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySort.SQL.Clear;
    Datam.FDQuerySort.SQL.text :=
      'SELECT * FROM ALL_HARDWARE WHERE TYPE_HARDWARE=''' + ComboBox4.text +
      ''' ORDER BY NAME_HARDWARE';
    Datam.FDQuerySort.Open;
    ComboBox3.Clear;
    while not Datam.FDQuerySort.EOF do
      begin
      /// заполняем  ComboBox3
        if Datam.FDQuerySort.FieldByName('NAME_HARDWARE').Value <> null then
          ComboBox3.Items.Add
            (string(Datam.FDQuerySort.FieldByName('NAME_HARDWARE').Value));
        Datam.FDQuerySort.Next;
      end;
    transactionSort(false); // закрываем транзакцию
    if ComboBox3.Items.Count > 0 then
      begin
        ComboBox3.ItemIndex := 0;
        ActiveControl := ComboBox3;
      end
    else
      begin
        CheckBox2.Checked := true;
        Edit1.Enabled := true;
        SpeedButton30.Enabled := true;
        ActiveControl := Edit1;
      end;
    Datam.FDQuerySort.Close;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Read table ALL_HARDWARE - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.ComboBox3DropDown(Sender: TObject);
begin
  DropDownWidth(ComboBox3);
  /// меняем ширину выпадающего списка
end;

procedure TfrmDomainInfo.ComboBox3KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox3.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBox3Select(Sender: TObject);
var
/// / поиск по выбору в ComboBox3
  notlike: string;
const
  hardware: array [0 .. 12] of string = ('MONTERBOARD', 'PROCESSOR', 'HDD_NAME',
    /// / имена полей для подставки
    'VIDEOCARD', 'SOUND_NAME', 'DVDROM_NAME', 'NETWORKINTERFACE',
    'MONTERBOARD_MANUFACTURE', 'PRINTER_NAME', 'USER_NAME', 'MONITOR_NAME',
    'OS_NAME', 'BIOS_DESCRIPTION');
begin
  try
    if ComboBox6.text = '<>' then
      notlike := ' NOT '
    else
      notlike := '';
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    // очистка перед созданием полей в DBgrid4
    DBgrid4Columns(hardware[ComboBox4.ItemIndex], ComboBox3.text, 'CONFIG_PC');
    Datam.FDQuerySelectSort.SQL.text := 'SELECT * FROM CONFIG_PC WHERE ' +
      notlike + hardware[ComboBox4.ItemIndex] + ' LIKE ' + '''%' +
      ComboBox3.text + '%''';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find table CONFIG_PC - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.CheckBox2Click(Sender: TObject);
begin
/// / включение edit и кнопки поиска
  Edit1.Enabled := CheckBox2.Checked;
  SpeedButton30.Enabled := CheckBox2.Checked;
end;

procedure TfrmDomainInfo.SpeedButton30Click(Sender: TObject);
var
/// // поиск по кнопке
  notlike: string;
const
  hardware: array [0 .. 12] of string = ('MONTERBOARD', 'PROCESSOR', 'HDD_NAME',
    'VIDEOCARD', 'SOUND_NAME', 'DVDROM_NAME', 'NETWORKINTERFACE',
    'MONTERBOARD_MANUFACTURE', 'PRINTER_NAME', 'USER_NAME', 'MONITOR_NAME',
    'OS_NAME', 'BIOS_DESCRIPTION');
begin
  try
    if ComboBox6.text = '<>' then
      notlike := ' NOT '
    else
      notlike := '';
    if ComboBox4.text = 'Оборудование' then
      begin
        showmessage('Выберите тип оборудования');
        exit;
      end;
    Datam.FDQuerySelectSort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    /// очистка перед созданием полей в DBgrid
    DBgrid4Columns(hardware[ComboBox4.ItemIndex], Edit1.text, 'CONFIG_PC');
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.SQL.text := 'SELECT * FROM CONFIG_PC WHERE ' +
      notlike + hardware[ComboBox4.ItemIndex] + ' LIKE ' + '''%' +
      Edit1.text + '%''';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find other table CONFIG_PC  - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton34Click(Sender: TObject);
var
  verprog: string;
begin
/// / обновление списка установленных программ + передача ID программы в ListSoftID
  DBGrid4.DataSource.DataSet.DisableControls;
  try
    Datam.FDQuerySort.Transaction.StartTransaction;
    /// открываем транзакцию
    Datam.FDQuerySort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySort.SQL.Clear;
    Datam.FDQuerySort.SQL.text := 'SELECT * FROM SOFT_PC ORDER BY SOFT_NAME';
    Datam.FDQuerySort.Open;
    ComboBox5.Clear;
    ListSoftID.Clear;
    while not Datam.FDQuerySort.EOF do
      begin
      /// заполняем  ComboBox5 списком программ
        if Datam.FDQuerySort.FieldByName('SOFT_NAME').Value <> null then
          begin
            verprog := '';
            if (Datam.FDQuerySort.FieldByName('SOFT_VERSION').Value <> null)
            then
              begin
                if string(Datam.FDQuerySort.FieldByName('SOFT_VERSION').Value)
                  <> '' then
                  verprog := '  Версия - ' +
                    string(Datam.FDQuerySort.FieldByName('SOFT_VERSION').Value);
              end;
            ComboBox5.Items.Add
              (string(Datam.FDQuerySort.FieldByName('SOFT_NAME').Value)
              + verprog);
            ListSoftID.Add(string(Datam.FDQuerySort.FieldByName('ID').Value));
            /// список ID программ
          end;
        Datam.FDQuerySort.Next;
      end;
    if ComboBox5.Items.Count > 0 then
      ComboBox5.ItemIndex := 0;
    Datam.FDQuerySort.SQL.Clear;
    Datam.FDQuerySort.Close;
    Datam.FDQuerySort.Transaction.commit;
    /// закрываем транзакцию
  finally
    DBGrid4.DataSource.DataSet.EnableControls;
  end;
end;

procedure TfrmDomainInfo.ComboBox5Select(Sender: TObject);
var
/// непоcредственно выборка списка компьютеров на которых установлено
  notlike: string;
begin
  try
    if ComboBox7.text = '<>' then
      notlike := 'NOT'
    else
      notlike := '';
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    /// очистка перед созданием полей в DBgrid
    DBgrid4Columns(ComboBox5.text, ComboBox5.text, 'SOFT_PC');
    /// создание колонок в DBgrid4
    Datam.FDQuerySelectSort.SQL.text := 'SELECT * FROM MAIN_PC_SOFT WHERE ' +
      notlike + ' INST_SOFT' + ' LIKE ' + '''%,' + ListSoftID
      [ComboBox5.ItemIndex] + ',%''';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find table MAIN_PC_SOFT - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.ComboBox8KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox8.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBox8Select(Sender: TObject);
var
  i: Integer;
  p: TPoint;
begin
  for i := 0 to ListView8.Items.Count - 1 do
    begin
      if pos(Ansiuppercase(ComboBox8.text),
        Ansiuppercase(ListView8.Items[i].SubItems[4])) <> 0 then
        begin
          // ListView8.HideSelection := False;
          ListView8.Items[i].Selected := true;
          ListView8.ItemIndex := i;
          ListView8.ItemFocused := ListView8.Items[i];
          p := ListView8.Items.Item[i].Position;
          ListView8.Scroll(p.X, p.Y);
          ListView8.Items[i].Checked := false;
          // Listview8.SetFocus;
          break;
        end;
    end;
  if ListView8.ItemFocused <> nil then
    ListView8.ItemFocused.MakeVisible(false);

end;

procedure TfrmDomainInfo.ComboBox9DropDown(Sender: TObject);
begin
  DropDownWidth(ComboBox9);
end;

procedure TfrmDomainInfo.ComboBox9KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox9.DroppedDown := true;
end;

procedure TfrmDomainInfo.ComboBox9Select(Sender: TObject);
var
/// непоcредственно выборка списка компьютеров на которых установлено
  notlike: string;
begin
  try
    if ComboBox9.text = 'Продукты Windows и Office' then
      exit;
    if ComboBox10.text = '<>' then
      notlike := 'NOT'
    else
      notlike := '';
    Datam.FDQuerySelectSort.Transaction.StartTransaction;
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    /// очистка перед созданием полей в DBgrid
    if pos('Windows', ComboBox9.text) <> 0 then
      begin
        DBgrid4Columns(ComboBox9.text, ComboBox9.text, 'MIC_LIC');
        /// создание колонок в DBgrid4
        Datam.FDQuerySelectSort.SQL.text := 'SELECT * FROM MICROSOFT_LIC WHERE '
          + notlike + ' WINPRODUCT' + ' LIKE ''%' + ComboBox9.text + '%''';
        Datam.FDQuerySelectSort.Open;
      end
    else // иначе это продукты Office  OFF_LIC
      begin
        DBgrid4Columns(ComboBox9.text, ComboBox9.text, 'OFF_LIC');
        /// создание колонок в DBgrid4
        Datam.FDQuerySelectSort.SQL.text := 'SELECT * FROM MICROSOFT_LIC WHERE '
          + notlike + ' SOFFICEPRODUCT' + ' LIKE ''%' + ComboBox9.text + '%''';
        Datam.FDQuerySelectSort.Open;
      end;

    Datam.FDQuerySelectSort.Transaction.commit;
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Find table MICROSOFT_PRODUCT - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.ComboBox4DropDown(Sender: TObject);
begin
/// ширина  списка ComboBox4
  PostMessage(ComboBox4.Handle, CB_SETDROPPEDWIDTH, 150, 0);
end;

procedure TfrmDomainInfo.ComboBox5DropDown(Sender: TObject);
begin
/// ширина  списка ComboBox5
  DropDownWidth(ComboBox5);
  /// меняем ширину выпадающего списка
end;

procedure TfrmDomainInfo.ComboBox5KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  ComboBox5.DroppedDown := true;
end;

procedure TfrmDomainInfo.SpeedButton32Click(Sender: TObject);
begin
/// / список всех программ
  try
    DBGrid4.DataSource.DataSet.DisableControls;
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes;
    /// очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    // очистка перед созданием полей в DBgrid4
    DBgrid4Columns('SOFT_NAME', 'Имя', 'ALL_SOFT');
    /// создание колонок в DBgrid4
    Datam.FDQuerySelectSort.SQL.text :=
      'SELECT * FROM SOFT_PC ORDER BY SOFT_NAME';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Read all software - ' + E.Message)
      end;
  end;
  DBGrid4.DataSource.DataSet.EnableControls;
end;

procedure TfrmDomainInfo.SpeedButton33Click(Sender: TObject);
begin
  try
    DBGrid4.DataSource.DataSet.DisableControls;
    transactionSort(true);
    /// открываем транзакцию
    Datam.FDQuerySelectSort.DeleteIndexes; // очистка индексов
    Datam.FDQuerySelectSort.IndexFieldNames := ''; // очистка индексов
    Datam.FDQuerySelectSort.SQL.Clear;
    // очистка перед созданием полей в DBgrid4
    DBgrid4Columns('NAME_HARDWARE', 'Модель', 'ALL_HARDWARE');
    /// создание колонок в DBgrid4
    Datam.FDQuerySelectSort.SQL.text :=
      'SELECT * FROM ALL_HARDWARE ORDER BY TYPE_HARDWARE';
    Datam.FDQuerySelectSort.Open;
    transactionSort(false);
    /// закрываем транзакцию
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Read all equipment - ' + E.Message)
      end;
  end;
  DBGrid4.DataSource.DataSet.EnableControls;
end;

procedure TfrmDomainInfo.FDQuerySelectSortAfterOpen(DataSet: TDataSet);
begin
  StatusBarFind.Panels[0].text := 'Всего - ' +
    inttostr(Datam.FDQuerySelectSort.RecordCount);
end;

procedure TfrmDomainInfo.SpeedButton31Click(Sender: TObject);
var
  i, z: Integer;
  saveexcel: Tsavedialog;
  Myworksheets: Variant;
  /// лист в ексель
  ProgressExport: TprogressBar;
begin
/// экспорт отсортированного в Excel
  try
    if (DBGrid4.SelectedRows.Count <= 0) then
      begin
        showmessage('Необходимо выделить строки');
        exit;
      end;
    if (InventConf) or (InventSoft) then
      begin
        showmessage('Дождитесь завершения инвентаризации');
        exit;
      end;

    if (DBGrid4.DataSource.DataSet.RecordCount <> 0) and
      (DBGrid4.SelectedRows.Count > 0) then
    /// / если выделенно больше 0 строк
      begin
        AllRunExcel;
        /// Полный запуск Excel
        MyExcel.WorkBooks[1].WorkSheets[1].Name := 'Сортировка' +
          inttostr(Random(50));
        /// имя листа
        Myworksheets := MyExcel.WorkBooks[1].WorkSheets[1];
        Myworksheets.Columns['A:Az'].columnwidth := 20; // ширина колонки
        Myworksheets.Columns['A:Az'].wraptext := true;
        /// Перенос по словам
        Myworksheets.Columns['A:Az'].numberFormat := '@';
        /// текстовый формат
        DBGrid4.DataSource.DataSet.DisableControls;
        /// //////////////////////////////////////////////////////////////
        if not Assigned(ProgressExport) then
          begin
            ProgressExport := TprogressBar.Create(self);
            ProgressExport.Parent := StatusBarFind;
            ProgressExport.Name := 'Prog';
          end;
        ProgressExport.left := StatusBarFind.Panels[0].width +
          StatusBarFind.Panels[1].width;
        ProgressExport.width := StatusBarFind.Panels[2].width;
        ProgressExport.Min := 0;
        ProgressExport.Max := DBGrid4.SelectedRows.Count - 1;
        ProgressExport.Step := 1;
        /// ////////////////////////////////////////////////////////////////
        for i := 0 to DBGrid4.Columns.Count - 1 do
        /// Имена колонок
          begin
            Myworksheets.cells[1, i + 1] := DBGrid4.Columns[i].Title.Caption;
          end;
        for z := 0 to DBGrid4.SelectedRows.Count - 1 do
          begin
            DBGrid4.DataSource.DataSet.Bookmark :=
              DBGrid4.SelectedRows.Items[z];
            for i := 0 to DBGrid4.Columns.Count - 1 do
              begin
                Myworksheets.cells[z + 2, i + 1] :=
                  delZ(vartostr(Datam.FDQuerySelectSort.FieldByName
                  (DBGrid4.Columns[i].FieldName).Value))
              end;
            if Assigned(ProgressExport) then
              ProgressExport.Position := z;
            StatusBarFind.Panels[1].text := inttostr(z) + ' из ' +
              inttostr(DBGrid4.SelectedRows.Count - 1);
          end;
        DBGrid4.DataSource.DataSet.EnableControls;
        if Assigned(ProgressExport) then
          FreeAndNil(ProgressExport);
        /// уничтожение прогресс бара
        MyExcel.Visible := true; // показать Excel
        StatusBarFind.Panels[1].text := 'Экспорт завершен';
        saveexcel := Tsavedialog.Create(self);
        saveexcel.Title := 'Сохранение отчета';
        saveexcel.InitialDir := ExtractFilePath(Application.ExeName) +
          'reports\';
        saveexcel.filename := 'Выборка ' + datetostr(Now);
        saveexcel.Filter := 'Книга Excel 97-2003 (*.XLS)|*.XLS|' +
          'Книга Open XML (*.xlsx)|*.xlsx|Формат HTML (*.html)|*.html|Файл csv (*.csv)|*.csv|'
          + 'Электронная таблица OpenDocument (*.ods)|*.ods';
        saveexcel.FilterIndex := 1;
        if saveexcel.Execute then
          begin
            case saveexcel.FilterIndex of
              1:
                begin
                  saveexcel.DefaultExt := '.xls';
                  SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt,
                    1, '56');
                end;
              2:
                begin
                  saveexcel.DefaultExt := '.xlsx';
                  SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt,
                    1, '51');
                end;
              3:
                begin
                  saveexcel.DefaultExt := '.html';
                  SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt,
                    1, '44');
                end;
              4:
                begin
                  saveexcel.DefaultExt := '.csv';
                  SaveWorkBook(saveexcel.filename +
                    saveexcel.DefaultExt, 1, '6');
                end;
              5:
                begin
                  saveexcel.DefaultExt := '.ods';
                  SaveWorkBook(saveexcel.filename + saveexcel.DefaultExt,
                    1, '60');
                end;
            end;
          end;
        FreeAndNil(saveexcel);
      end;
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', datetostr(Date) + '/' + timetostr(time) +
          ' Export Excel sort - ' + E.Message);
        VarClear(Myworksheets);
        VariantClear(MyExcel);
        if Assigned(saveexcel) then
          FreeAndNil(saveexcel);
        if Assigned(ProgressExport) then
          FreeAndNil(ProgressExport);
      end;
  end;
  if Assigned(saveexcel) then
    FreeAndNil(saveexcel);
  VarClear(Myworksheets);
  VariantClear(MyExcel);
end;

procedure TfrmDomainInfo.SpeedButton35Click(Sender: TObject);
/// использовать список компьютеров
var
  i, z: Integer;
  mac, pcname: Boolean;
begin
  try
    mac := false;
    pcname := false;
    if DBGrid4.SelectedRows.Count <= 0 then
      begin
        showmessage('Необходимо выделить строки');
        exit;
      end;

    /// //////////////////////////////////////////  проверка на сканирование компьютеров
    if not killscancomp then
      begin
        Memo1.Lines.Add
          ('Не могу остановить сканирование, повторите чуть позже...');
        exit;
      end;

    /// ////////////////////////////////////////////////////////
    ///
    for z := 0 to Datam.FDQuerySelectSort.FieldCount - 1 do
    /// проверяем имена столбцов
      begin
      /// / находим есть ли поля в запросе
        if Datam.FDQuerySelectSort.Fields[z].FieldName = 'ANSWER_MAC' then
          mac := true;
        if Datam.FDQuerySelectSort.Fields[z].FieldName = 'PC_NAME' then
          pcname := true;
      end;

    if pcname then // если поля есть то переносим список компов
      begin
        if (DBGrid4.DataSource.DataSet.RecordCount <> 0) and
          (DBGrid4.SelectedRows.Count > 0) then
          begin
            DBGrid4.DataSource.DataSet.DisableControls;
            ListView8.Clear;
            for i := 0 to DBGrid4.SelectedRows.Count - 1 do
              begin
                DBGrid4.DataSource.DataSet.Bookmark :=
                  DBGrid4.SelectedRows.Items[i];
                with ListView8.Items.Add do
                  begin
                    Caption := '';
                    SubItems.Add
                      (vartostr(Datam.FDQuerySelectSort.FieldByName
                      ('PC_NAME').Value));
                    SubItems.Add('');
                    if mac then
                      SubItems.Add
                        (vartostr(Datam.FDQuerySelectSort.FieldByName
                        ('ANSWER_MAC').Value))
                    else
                      SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                    SubItems.Add('');
                  end;
              end;
            DBGrid4.DataSource.DataSet.EnableControls;
          end;
        readinfoforpcDB;
        /// читаем информацию о списке компов из БД
        MainPage.TabIndex := 0;
        /// / Переход к вкладке со списком компьютеров
      end;
    if not pcname then
      showmessage('В списке нет компьютеров');
  except
    on E: Exception do
      begin
        showmessage('Произошла ошибка, повторите еще раз');
        Memo1.Lines.Add(E.Message);
        Log_write(1, 'error', ' Apply list pc - ' + E.Message)
      end;
  end;
end;

procedure TfrmDomainInfo.SpeedButton36Click(Sender: TObject);
begin
  GridSelectAll(DBGrid4);
  /// выделить весь grid
end;

procedure TfrmDomainInfo.SpeedButton37Click(Sender: TObject);
begin
  GridSelectAll(DBGrid3);
  /// выделить весь grid
end;

procedure TfrmDomainInfo.SpeedButton38Click(Sender: TObject);
begin
  GridSelectAll(DBGrid1);
  /// выделить весь grid
end;

/// //////////////////////////////////////////////////////////////////////////////
procedure TfrmDomainInfo.ComboBox1Select(Sender: TObject);
/// выбор группы домена из списка
var
  IniFile: TMemInifile;
  PCGroup, UserGroup: Boolean;
begin
  try
    if not KillscanSelectGroup then
      begin
        Memo1.Lines.Add ('Не могу остановить сканирование, повторите чуть позже...');
        exit;
      end;
    if ComboBox1.text <> '' then
      begin
        if FileExists(ExtractFilePath(Application.ExeName) + '\Settings.ini')
        then
          begin
            IniFile := TMemInifile.Create(ExtractFilePath(Application.ExeName) +
              '\Settings.ini', TEncoding.Unicode);
            try
              IniFile.WriteString('ConfLAN', 'Group',
                ComboBox1.Items[ComboBox1.ItemIndex]);
              UserGroup := IniFile.ReadBool('ConfLAN', 'UserGroup', true);
              // Выгружать только пользователей выбранной группы безопасности домена
              PCGroup := IniFile.ReadBool('ConfLAN', 'PCGroup', true);
              // Выгружать только компьютеры выбранной группы безопасности домена
            finally
              IniFile.UpdateFile;
              FreeAndNil(IniFile);
            end;
          end;

        if PCGroup then
          GetPCForGroupToLisViewComboBox(ComboBox1.text)
          // добавляем список компьютеров в выбранной группе в ListView главном окне программы и в combobox в разделе Управление
        else
          GetAllGroupUsers(ComboBox1.text);
        // иначе добавляем только список компьютеров в выбранной группе в ListView главном окне программы т.к. в comboBox в разделе Управление уже есть полный список компьютеров домена

        if UserGroup then
          GetAllGroupUsersInComboBox(ComboBox1.text);
        // если необходимо считывать пользователей только выбранной группы , читаем список пользователей выбранной группы
        if Datam.ConnectionDB.Connected then
          readinfoforpcDB;
        /// если соединение установленно то читаем пользователей из базы данных
      end;
    /// ////////////////////////////////

  Except
    on E: Exception do
      begin
        Memo1.Lines.Add('Ошибка загрузки списка из базы данных:' + E.Message);
      end;
  end;
end;

procedure TfrmDomainInfo.startscanlistpc;
var
  i: Integer;
begin
  try
    /// ///////////////////////////////////////////////
    if ListView8.Items.Count > 0 then
    /// если список компьютеров не пуст, то запускаем сканирование.
      begin
        PingPCList := TstringList.Create;
        PingPCList := createListpc('');
        SelectedPCPing := GeneralPingPC.GeneralPing.Create(true);
        SelectedPCPing.FreeOnTerminate := false; // завершим поток после остановки
        SelectedPCPing.Priority := tpLower;
        OutForPing := true; // потоку можно работать
        SolveExitInvScan:=false;  // признак что поток запущен необходимо для того если нажать сразу остановить то сначала поток запустится а только потом остановится
        SelectedPCPing.Start;  //пуск сканирования
        SpeedButton39.Visible := false;
        SpeedButton40.Visible := true;
      end;
  except
    on E: Exception do
      showmessage('Ошибка : ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.SpeedButton39Click(Sender: TObject);
/// / кнопка запуска сканирования
begin
  if ListView8.Items.Count > 1 then
    begin
      SpeedButton39.Visible := false;
      SpeedButton40.Visible := true;
      startscanlistpc;
    end
  else
    showmessage('Список компьютеров очень мал');

end;

procedure TfrmDomainInfo.SpeedButton40Click(Sender: TObject);
/// остановить сканирование
begin
  try
   if not killscancomp then
      begin
        Memo1.Lines.Add('Не могу остановить сканирование, повторите чуть позже...');
        exit;
      end;
  except
    on E: Exception do
      showmessage('Ошибка : ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.SpeedButton41Click(Sender: TObject);
begin
  try
    if GroupBox2.Caption = '' then
      begin
        showmessage('Не выбран компьютер');
        exit;
      end;
    Datam.FDQueryRead2.SQL.Clear;
    /// чтение конфигурации первого вхождения и передача в treeview
    Datam.FDQueryRead2.SQL.text := 'SELECT * FROM CONFIG_PC WHERE PC_NAME=''' +
      GroupBox2.Caption + '''' + ' ORDER BY DATE_INV ASC';
    /// /    DESC или ASC
    Datam.FDQueryRead2.Open;
    frmDomainInfo.createtreeView(Datam.FDQueryRead2, TreeView4);
    // Функция постороения дерева
    Datam.FDQueryRead2.SQL.Clear; // очистить
    Datam.FDQueryRead2.Close;
    /// закрыть нах после чтения
  except
    on E: Exception do
      showmessage('Ошибка : ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.SpeedButton42Click(Sender: TObject);
var
  NewDescription: string;
begin
  if GroupBox2.Caption = '' then
    begin
      showmessage('Не выбран компьютер');
      exit;
    end;
  if not InputQuery('Описание компьютера в свойствах (Инв№)', 'Описание',
    NewDescription) then
    exit;
  PutInvNumberToDataBase(GroupBox2.Caption, NewDescription);
  PutInvNumber(GroupBox2.Caption, NewDescription);
end;

procedure TfrmDomainInfo.SpeedButton43Click(Sender: TObject);
begin
  if GroupBox2.Caption = '' then
    begin
      showmessage('Не выбран компьютер');
      exit;
    end;
  if ping(GroupBox2.Caption) then
    begin
      MyPS := GroupBox2.Caption;
      MyUser := LabeledEdit1.text;
      MyPasswd := LabeledEdit2.text;
      OKRightDlg1234567.showmodal;
    end;
end;

procedure TfrmDomainInfo.SpeedButton44Click(Sender: TObject);
var
  Res: Integer;
begin
  if InventConf then
    begin
      showmessage('Дождитесь завершения текущей инвентаризации!');
      exit;
    end;
  SpeedButton44.Tag := SpeedButton44.Tag + 1;
  if SpeedButton44.Tag > RandomRange(5, 10) then
    begin
      SpeedButton44.Tag := 0;
      Res := MessageBox(self.Handle,
        PChar('Запустить инвентаризацию компьютеров???'),
        PChar('Сколько можно на меня тыкать...'), MB_YESNO);
      if Res = IDYes then
        begin
          SpeedButton13.Click;
          exit;
        end;
    end;

  if GroupBox2.Caption = '' then
    begin
      showmessage('Не выбран компьютер');
      exit;
    end;
  refreshinfoPC(GroupBox2.Caption, LabeledEdit1.text, LabeledEdit2.text);

end;

procedure TfrmDomainInfo.SpeedButton45Click(Sender: TObject);
begin
  if GroupBox2.Caption <> '' then
    begin
      if PagePropertiesPC.TabIndex = 1 then
        readSoftRorSelectPC(GroupBox2.Caption, ListViewSoftinPC);
    end
  else
    showmessage('Не выбран компьютер');
end;

procedure TfrmDomainInfo.SpeedButton46Click(Sender: TObject);
begin
  if ListViewSoftinPC.Items.Count <> 0 then
    begin
      popupListViewSaveAs(ListViewSoftinPC, 'Сохранение списка программ', '');
    end
  else
    showmessage('Список пуст!');
end;

procedure TfrmDomainInfo.SpeedButton47Click(Sender: TObject);
begin
  if GroupBox2.Caption <> '' then
    inventorySoftForSelectPC(GroupBox2.Caption, LabeledEdit1.text,
      LabeledEdit2.text)
  else
    showmessage('Не выбран компьютер');
end;

procedure TfrmDomainInfo.SpeedButton48Click(Sender: TObject);
begin
  SaveTreeView(TreeView4, GroupBox2.Caption);
end;

procedure TfrmDomainInfo.SpeedButton49Click(Sender: TObject);
var
  i: Integer;
  FDQueryLic: TFDQuery;
  Transaction: TFDTransaction;
begin
  try
    if LVMicrosoft.SelCount = 0 then
      exit;
    Transaction := TFDTransaction.Create(nil);
    Transaction.Connection := Datam.ConnectionDB;
    Transaction.Options.Isolation := xiReadCommitted;
    /// xiReadCommitted; //xiSnapshot; xiUnspecified
    FDQueryLic := TFDQuery.Create(nil);
    FDQueryLic.Transaction := Transaction;
    FDQueryLic.Connection := Datam.ConnectionDB;
    Transaction.StartTransaction; // стартуем
    LVMicrosoft.Items.BeginUpdate;
    try
      for i := LVMicrosoft.Items.Count - 1 downto 0 do
        begin
          if LVMicrosoft.Items[i].Selected then
            begin
              FDQueryLic.SQL.Clear;
              FDQueryLic.SQL.text := 'delete FROM MICROSOFT_LIC WHERE ID=''' +
                LVMicrosoft.Items[i].Caption + '''';
              FDQueryLic.ExecSQL;
              LVMicrosoft.Items[i].Delete;
            end;
        end;
    finally
      Transaction.commit;
      LVMicrosoft.Items.EndUpdate;
      FDQueryLic.Free;
      Transaction.Free;
    end;
  except
    on E: Exception do
      Memo1.Lines.Add('Ошибка удаления ' + E.Message);
  end;
end;

procedure TfrmDomainInfo.InstallMSIForListPC;
/// / групповая установка программы
begin
  if ListView8.Items.Count = 0 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  GroupPC := true;
  OKRightDlg1.showmodal;
end;

procedure TfrmDomainInfo.DeleteProgramMSIForListPC;
/// /// групповое удаление программы (msi)
begin
  if ListView8.Items.Count < 1 then
    begin
      showmessage('В списке нет компьютеров');
      exit;
    end;
  if ListView8.Items.Count > 0 then
    begin
      GroupPC := true;
      OKRightDlg12345678910111213141516171819202122.showmodal;
    end
end;

end.
