object ActivateOffice: TActivateOffice
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1055#1088#1086#1076#1091#1082#1090#1099' Microsoft Office (2010,2013,2016...)'
  ClientHeight = 512
  ClientWidth = 1117
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poOwnerFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 9
    Top = 361
    Width = 454
    Height = 13
    Caption = 
      #1044#1083#1103' '#1072#1082#1090#1080#1074#1072#1094#1080#1080' '#1087#1088#1086#1076#1091#1082#1090#1072' '#1091#1076#1072#1083#1077#1085#1085#1099#1081' '#1082#1086#1084#1087#1100#1102#1090#1077#1088' '#1076#1086#1083#1078#1077#1085' '#1073#1099#1090#1100' '#1087#1086#1076#1082#1083#1102#1095#1077#1085 +
      '  '#1082' '#1080#1085#1090#1077#1088#1085#1077#1090#1091'.'
  end
  object Label1: TLabel
    Left = 8
    Top = 345
    Width = 498
    Height = 13
    Caption = 
      #1050#1083#1102#1095' '#1087#1088#1086#1076#1091#1082#1090#1072' '#1089#1086#1089#1090#1086#1103#1097#1080#1081' '#1080#1079' 25 '#1079#1085#1072#1082#1086#1074', '#1091#1082#1072#1079#1099#1074#1072#1090#1100' '#1074' '#1092#1086#1088#1084#1072#1090#1077'  XXXXX' +
      '-XXXXX-XXXXX-XXXXX-XXXXX '
  end
  object Button1: TButton
    Left = 8
    Top = 382
    Width = 120
    Height = 25
    Caption = #1054#1073#1085#1086#1074#1080#1090#1100' '#1089#1087#1080#1089#1086#1082
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 0
    Top = 423
    Width = 1117
    Height = 89
    Align = alBottom
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object ListView1: TListView
    Left = 0
    Top = 0
    Width = 1117
    Height = 339
    Align = alTop
    Columns = <
      item
        Caption = #1048#1084#1103' '#1087#1088#1086#1076#1091#1082#1090#1072
        Width = 150
      end
      item
        Caption = #1054#1087#1080#1089#1072#1085#1080#1077' '#1083#1080#1094#1077#1085#1079#1080#1080
        Width = 150
      end
      item
        Caption = #1057#1090#1072#1090#1091#1089' '#1083#1080#1094#1077#1085#1079#1080#1080
        Width = 100
      end
      item
        Caption = #1050#1083#1102#1095' '#1087#1088#1086#1076#1091#1082#1090#1072' ('#1087#1086#1089#1083#1077#1076#1085#1080#1077' 5 '#1089#1080#1084#1074#1086#1083#1086#1074')'
        Width = 150
      end
      item
        Caption = #1050#1086#1076' '#1087#1088#1086#1076#1091#1082#1090#1072
        Width = 200
      end
      item
        Caption = #1048#1076#1077#1085#1090#1080#1092#1080#1082#1072#1090#1086#1088
        Width = 200
      end
      item
        AutoSize = True
        Caption = #1054#1087#1080#1089#1072#1085#1080#1077' '#1089#1090#1072#1090#1091#1089#1072' '#1083#1080#1094#1077#1085#1079#1080#1080
      end>
    GridLines = True
    RowSelect = True
    TabOrder = 2
    ViewStyle = vsReport
    OnCustomDrawSubItem = ListView1CustomDrawSubItem
    OnDblClick = ListView1DblClick
  end
  object Button2: TButton
    Left = 152
    Top = 382
    Width = 120
    Height = 25
    Caption = #1059#1089#1090#1072#1085#1086#1074#1082#1072' '#1082#1083#1102#1095#1072
    TabOrder = 3
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 290
    Top = 382
    Width = 120
    Height = 25
    Caption = #1040#1082#1090#1080#1074#1072#1094#1080#1103
    TabOrder = 4
    OnClick = Button3Click
  end
  object button4: TButton
    Left = 432
    Top = 382
    Width = 120
    Height = 25
    Caption = #1059#1076#1072#1083#1077#1085#1080#1077' '#1082#1083#1102#1095#1072
    TabOrder = 5
    OnClick = button4Click
  end
end