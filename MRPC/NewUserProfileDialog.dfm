object OKRightDlg123456789101112: TOKRightDlg123456789101112
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1057#1084#1077#1085#1072' '#1074#1083#1072#1076#1077#1083#1100#1094#1072' '#1087#1088#1086#1092#1080#1083#1103
  ClientHeight = 118
  ClientWidth = 384
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 18
    Top = 67
    Width = 351
    Height = 13
    Caption = '('#1054#1073#1103#1079#1072#1090#1077#1083#1100#1085#1086', '#1077#1089#1083#1080' '#1087#1088#1086#1092#1080#1083#1100' '#1085#1086#1074#1086#1075#1086' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' '#1091#1078#1077' '#1089#1091#1097#1077#1089#1090#1074#1091#1077#1090')'
  end
  object OKBtn: TButton
    Left = 294
    Top = 86
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 205
    Top = 86
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object LabeledEdit1: TLabeledEdit
    Left = 8
    Top = 24
    Width = 361
    Height = 21
    EditLabel.Width = 197
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1085#1086#1074#1086#1075#1086' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' ('#1074#1083#1072#1076#1077#1083#1100#1094#1072')'
    TabOrder = 2
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 51
    Width = 368
    Height = 14
    Caption = #1047#1072#1084#1077#1085#1080#1090#1100' '#1074#1083#1072#1076#1077#1083#1100#1094#1072' '#1080' '#1091#1076#1072#1083#1080#1090#1100' '#1080#1084#1077#1102#1097#1080#1081#1089#1103' '#1087#1088#1086#1092#1080#1083#1100' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103
    TabOrder = 3
    OnClick = CheckBox1Click
  end
end
