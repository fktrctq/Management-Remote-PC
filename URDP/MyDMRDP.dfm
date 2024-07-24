object DataMod: TDataMod
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 391
  Width = 636
  object FDTransactionUpdate: TFDTransaction
    Connection = ConnectionDB
    Left = 64
    Top = 24
  end
  object ConnectionDB: TFDConnection
    Params.Strings = (
      'OSAuthent=No'
      'Server=localhost'
      'CharacterSet=UTF8'
      'User_Name=sysdba'
      'Password=masterkey'
      'DriverID=FB')
    LoginPrompt = False
    UpdateTransaction = FDTransactionUpdate
    Left = 192
    Top = 24
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    DriverID = 'FB'
    Left = 296
    Top = 24
  end
  object FDTransactionRead: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 64
    Top = 80
  end
  object FDQueryRead: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionRead
    Left = 196
    Top = 80
  end
  object FDQueryWrite: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionWrite
    UpdateTransaction = FDTransactionWrite
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC ORDER BY PC_NAME')
    Left = 198
    Top = 142
  end
  object FDTransactionWrite: TFDTransaction
    Options.EnableNested = False
    Connection = ConnectionDB
    Left = 70
    Top = 142
  end
  object FDGUIxScriptDialog1: TFDGUIxScriptDialog
    Provider = 'Forms'
    Left = 67
    Top = 232
  end
  object FDScriptCreateTabl: TFDScript
    SQLScripts = <
      item
        Name = 'ddd'
        SQL.Strings = (
          'SELECT DISTINCT R.RDB$RELATION_NAME '
          ' FROM RDB$RELATIONS R '
          'WHERE R.RDB$RELATION_NAME = '#39'RDP_SET'#39
          'INTO :ORDSTATNAME; '
          ' IF (ORDSTATNAME IS NULL) THEN'
          ' BEGIN'
          
            ' CREATE_GEN_RDP:='#39'CREATE GENERATOR GEN_RDP_SET_ID START WITH 0 I' +
            'NCREMENT BY 1'#39';'
          ' '
          ' CREATE_TAB_RDP:= '#39' CREATE TABLE RDP_SET ('
          #9'ID_RDP  INTEGER NOT NULL,'
          #9'PC_NAME     VARCHAR(100) ,'
          #9'ColorDepth     integer ,'
          #9'BitmapPeristence     integer ,'
          #9'CachePersistenceActive     integer ,'
          #9'BitmapCacheSize    integer ,'
          #9'VirtualCache16BppSize     integer ,'
          #9'VirtualCache32BppSize     integer ,'
          #9'VirtualCacheSize     integer ,'
          #9'DisableCtrlAltDel     integer ,'
          #9'DoubleClickDetect     integer ,'
          #9'EnableWindowsKey     integer ,'
          #9'MinutesToIdleTimeout     integer ,'
          #9'OverallConnectionTimeout     integer ,'
          #9'RdpPort     integer ,'
          #9'PerformanceFlags     integer ,'
          #9'NetworkConnectionType     integer ,'
          #9'MaxReconnectAttempts     integer ,'
          #9'AudioRedirectionMode     integer ,'
          #9'AuthenticationLevel     integer ,'
          #9'KeyboardHookMode     integer ,'
          
            #9'{DisplayConnectionBar,GrabFocusOnConnect,BandwidthDetection,Ena' +
            'bleAutoReconnect,'
          #9'RedirectDrives,RedirectPrinters,RelativeMouseMode'
          
            #9',RedirectClipboard,RedirectDevices,RedirectPorts,ConnectToAdmin' +
            'isterServer,'
          
            #9'AudioCaptureRedirectionMode,EnableSuperPan,EnableCredSspSupport' +
            '}'
          #9'DisplayConnectionBar   BOOLEAN ,'
          #9'GrabFocusOnConnect    BOOLEAN ,'
          #9'BandwidthDetection    BOOLEAN ,'
          #9'EnableAutoReconnect   BOOLEAN ,'
          #9'RedirectDrives    BOOLEAN ,'
          #9'RedirectPrinters    BOOLEAN ,'
          #9'RelativeMouseMode    BOOLEAN ,'
          #9'RedirectClipboard    BOOLEAN ,'
          #9'RedirectDevices    BOOLEAN ,'
          #9'RedirectPorts    BOOLEAN ,'
          #9'ConnectToAdministerServer    BOOLEAN ,'
          #9'AudioCaptureRedirectionMode    BOOLEAN ,'
          #9'EnableSuperPan    BOOLEAN , //EnableCredSspSupport'
          #9'EnableCredSspSupport    BOOLEAN)'#39';'
          #9
          
            'CREATE_PK_RDP:='#39'ALTER TABLE RDP_SET ADD CONSTRAINT PK_RDP_SET PR' +
            'IMARY KEY (ID_RDP)'#39';'
          ''
          
            'CREATE_TRIG_RDP:='#39'CREATE OR ALTER TRIGGER RDP_SET_BI FOR RDP_SET' +
            ' ACTIVE BEFORE INSERT POSITION 0'
          
            'as begin  if (new.id_rdp is null) then new.id_rdp = gen_id(GEN_R' +
            'DP_SET_ID,1); end'#39';'
          ''
          'EXECUTE STATEMENT CREATE_GEN_RDP;'
          'EXECUTE STATEMENT CREATE_TAB_RDP;'
          'EXECUTE STATEMENT CREATE_PK_RDP;'
          'EXECUTE STATEMENT CREATE_TRIG_RDP;'
          ''
          'END'
          '')
      end>
    Connection = ConnectionDB
    ScriptOptions.AutoPrintParams = True
    ScriptOptions.BreakOnError = True
    ScriptOptions.FileEncoding = ecUTF8
    ScriptOptions.FileEndOfLine = elWindows
    ScriptOptions.DriverID = 'FB'
    ScriptOptions.SQLDialect = 3
    Params = <>
    Macros = <>
    FetchOptions.AssignedValues = [evItems, evAutoClose, evAutoFetchAll]
    FetchOptions.Items = [fiBlobs, fiDetails]
    ResourceOptions.AssignedValues = [rvMacroCreate, rvMacroExpand, rvDirectExecute, rvPersistent]
    ResourceOptions.MacroCreate = False
    ResourceOptions.DirectExecute = True
    Left = 412
    Top = 23
  end
  object FDGUIxWaitCursor1: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 186
    Top = 239
  end
end
