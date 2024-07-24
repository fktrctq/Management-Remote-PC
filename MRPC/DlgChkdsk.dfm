object OKRightDlg123: TOKRightDlg123
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = 'Chkdsk'
  ClientHeight = 174
  ClientWidth = 263
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 174
    Top = 139
    Width = 75
    Height = 25
    Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 85
    Top = 139
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object CheckBox1: TCheckBox
    Left = 16
    Top = 6
    Width = 176
    Height = 17
    Caption = #1048#1089#1087#1088#1072#1074#1083#1103#1090#1100' '#1086#1096#1080#1073#1082#1080' '#1085#1072' '#1076#1080#1089#1082#1077
    Checked = True
    State = cbChecked
    TabOrder = 2
  end
  object CheckBox2: TCheckBox
    Left = 16
    Top = 27
    Width = 201
    Height = 17
    Caption = #1040#1082#1090#1080#1074#1085#1072#1103' '#1087#1088#1086#1074#1077#1088#1082#1072' '#1079#1072#1087#1080#1089#1077#1081
    Checked = True
    State = cbChecked
    TabOrder = 3
  end
  object CheckBox3: TCheckBox
    Left = 16
    Top = 49
    Width = 169
    Height = 17
    Caption = #1055#1088#1086#1087#1091#1089#1082' '#1087#1088#1086#1074#1077#1088#1082#1080' '#1094#1080#1082#1083#1086#1074
    Checked = True
    State = cbChecked
    TabOrder = 4
  end
  object CheckBox4: TCheckBox
    Left = 16
    Top = 70
    Width = 193
    Height = 17
    Caption = #1054#1090#1082#1083#1102#1095#1077#1085#1080#1077' '#1090#1086#1084#1072' '#1076#1083#1103' '#1087#1088#1086#1074#1077#1088#1082#1080
    TabOrder = 5
  end
  object CheckBox5: TCheckBox
    Left = 16
    Top = 91
    Width = 249
    Height = 17
    Caption = #1055#1088#1086#1074#1077#1088#1103#1090#1100' '#1080' '#1074#1086#1089#1089#1090#1072#1085#1072#1074#1083#1080#1074#1072#1090#1100' Bad '#1089#1077#1082#1090#1086#1088#1072
    Checked = True
    State = cbChecked
    TabOrder = 6
  end
  object CheckBox6: TCheckBox
    Left = 16
    Top = 113
    Width = 233
    Height = 17
    Caption = #1055#1088#1086#1074#1077#1088#1082#1072' '#1087#1088#1080' '#1089#1083#1077#1076#1091#1102#1097#1077#1081' '#1079#1072#1075#1088#1091#1079#1082#1077' '#1054#1057
    TabOrder = 7
  end
end
