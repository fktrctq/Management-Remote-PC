object SelectDelMSITask: TSelectDelMSITask
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  ClientHeight = 114
  ClientWidth = 388
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 263
    Top = 81
    Width = 112
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1079#1072#1076#1072#1095#1091
    Default = True
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 183
    Top = 81
    Width = 74
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object Button1: TButton
    Left = 304
    Top = 50
    Width = 71
    Height = 25
    Hint = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1089#1087#1080#1089#1086#1082' '#1087#1088#1086#1075#1088#1072#1084#1084' '#1089' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1072' '#1074' '#1089#1077#1090#1080
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 2
    OnClick = Button1Click
  end
  object ComboBox1: TComboBox
    Left = 8
    Top = 52
    Width = 287
    Height = 21
    CharCase = ecLowerCase
    DropDownCount = 10
    Sorted = True
    TabOrder = 3
    TextHint = #1047#1072#1075#1088#1091#1079#1080#1090#1077' '#1089#1087#1080#1089#1086#1082' '#1087#1088#1086#1075#1088#1072#1084#1084' '#1089' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1072' '#1074' '#1089#1077#1090#1080
    OnSelect = ComboBox1Select
  end
  object CheckBox2: TCheckBox
    Left = 8
    Top = 31
    Width = 209
    Height = 17
    Caption = #1058#1088#1077#1073#1086#1074#1072#1090#1100' '#1090#1086#1095#1085#1086#1077' '#1089#1086#1074#1087#1072#1076#1077#1085#1080#1077' '#1080#1084#1077#1085#1080
    Checked = True
    State = cbChecked
    TabOrder = 4
    OnClick = CheckBox2Click
  end
  object Edit1: TEdit
    Left = 8
    Top = 8
    Width = 367
    Height = 21
    TabOrder = 5
    TextHint = #1048#1084#1103' '#1080#1083#1080' '#1095#1072#1089#1090#1100' '#1080#1084#1077#1085#1080' '#1087#1088#1086#1075#1088#1072#1084#1084#1099' '#1076#1083#1103' '#1091#1076#1072#1083#1077#1085#1080#1103'.'
  end
end
