object DataM: TDataM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 703
  Width = 1058
  object FDTransactionSelectRead: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 735
    Top = 232
  end
  object FDTransactionUpdate: TFDTransaction
    Connection = ConnectionDB
    Left = 64
    Top = 24
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    DriverID = 'FB'
    Left = 280
    Top = 24
  end
  object ConnectionDB: TFDConnection
    Params.Strings = (
      'OSAuthent=No'
      'Server=localhost'
      'CharacterSet=UTF8'
      'User_Name=sysdba'
      'Password=masterkey'
      'DriverID=FB'
      'SQLDialect=3')
    LoginPrompt = False
    Transaction = FDTransactionSelectRead
    UpdateTransaction = FDTransactionUpdate
    Left = 168
    Top = 24
  end
  object FDQueryReadSoft: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionReadsoft
    UpdateTransaction = FDTransactionSelectRead
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC, CONFIG_PC ORDER BY PC_NAME')
    Left = 863
    Top = 153
  end
  object DataSourceSoft: TDataSource
    DataSet = FDQueryReadSoft
    Left = 985
    Top = 151
  end
  object FDSelectReadSoft: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionSelectRead
    UpdateTransaction = FDTransactionSelectRead
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC, CONFIG_PC ORDER BY PC_NAME')
    Left = 863
    Top = 233
  end
  object FDQueryLoadListPC: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionReadListPC
    UpdateTransaction = FDTransactionReadListPC
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC ORDER BY PC_NAME')
    Left = 863
    Top = 313
  end
  object FDTransactionReadListPC: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 727
    Top = 312
  end
  object FDReadSoftSelectPC: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionReadSoftSelectPC
    UpdateTransaction = FDTransactionSelectRead
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC, CONFIG_PC ORDER BY PC_NAME')
    Left = 863
    Top = 393
  end
  object FDQueryWriteEXPT_PC: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionWriteEXPT
    UpdateTransaction = FDTransactionWriteEXPT
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC ORDER BY PC_NAME')
    Left = 871
    Top = 473
  end
  object FDTransactionWriteEXPT: TFDTransaction
    Options.EnableNested = False
    Connection = ConnectionDB
    Left = 719
    Top = 472
  end
  object FDTransactionReadsoft: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 735
    Top = 160
  end
  object FDTransactionReadSoftSelectPC: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 719
    Top = 392
  end
  object FDQueryReadPropertiesPC: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionReadPropPC
    UpdateTransaction = FDTransactionSelectRead
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC, CONFIG_PC ORDER BY PC_NAME')
    Left = 871
    Top = 83
  end
  object FDTransactionReadPropPC: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 735
    Top = 88
  end
  object FDQueryFindViolations: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionFindViolat
    UpdateTransaction = FDTransactionSelectRead
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      'SELECT * FROM MAIN_PC, CONFIG_PC ORDER BY PC_NAME')
    Left = 871
    Top = 19
  end
  object FDTransactionFindViolat: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 735
    Top = 24
  end
  object FDTransactionQuery3: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 64
    Top = 168
  end
  object FDQueryRead3: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionQuery3
    Left = 164
    Top = 168
  end
  object DataSource2: TDataSource
    DataSet = FDQueryRead3
    Left = 250
    Top = 168
  end
  object DataSource1: TDataSource
    DataSet = FDQueryReadPropertiesPC
    Left = 983
    Top = 80
  end
  object DataSourceSelectSort: TDataSource
    DataSet = FDQuerySelectSort
    Left = 264
    Top = 239
  end
  object FDTransactionSort: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 64
    Top = 240
  end
  object FDQuerySelectSort: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionSort
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    Left = 168
    Top = 231
  end
  object FDQuerySort: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionSelectSort
    Left = 168
    Top = 303
  end
  object DataSourceSort: TDataSource
    DataSet = FDQuerySort
    Left = 264
    Top = 303
  end
  object FDTransactionSelectSort: TFDTransaction
    Options.ReadOnly = True
    Connection = ConnectionDB
    Left = 64
    Top = 304
  end
  object FDQueryRead2: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionSelectRead
    Left = 168
    Top = 112
  end
  object DataSourceViolation: TDataSource
    DataSet = FDQueryFindViolations
    Left = 975
    Top = 24
  end
  object FDScriptTempTabl: TFDScript
    SQLScripts = <
      item
        SQL.Strings = (
          'SET SQL DIALECT 3;'
          ''
          'CREATE GLOBAL TEMPORARY TABLE PC_EVENT ('
          '    NEW_FIELD  INTEGER NOT NULL,'
          '    PC_NAME    VARCHAR(100) NOT NULL,'
          '    "LOGFILE"  VARCHAR(200),'
          '    LOGSOURCE  VARCHAR(200)'
          ') ON COMMIT PRESERVE ROWS;'
          ''
          ''
          ''
          '')
      end>
    Connection = ConnectionDB
    Transaction = FDTransactionTempTabl
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
    Left = 995
    Top = 544
  end
  object FDTransactionTempTabl: TFDTransaction
    Connection = ConnectionDB
    Left = 735
    Top = 544
  end
  object FDQueryTempTabl: TFDQuery
    Connection = ConnectionDB
    Transaction = FDTransactionTempTabl
    UpdateTransaction = FDTransactionTempTabl
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 10000
    UpdateOptions.AssignedValues = [uvGeneratorName, uvCheckUpdatable]
    SQL.Strings = (
      '')
    Left = 879
    Top = 545
  end
  object FDGUIxScriptDialog1: TFDGUIxScriptDialog
    Provider = 'Forms'
    Left = 571
    Top = 544
  end
end
