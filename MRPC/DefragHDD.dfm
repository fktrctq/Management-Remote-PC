object OKRightDlg1234: TOKRightDlg1234
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1044#1080#1089#1082' -'
  ClientHeight = 166
  ClientWidth = 201
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 270
    Top = 68
    Width = 137
    Height = 41
    Caption = #1044#1077#1092#1088#1072#1075#1084#1077#1085#1090#1072#1094#1080#1103' '#1076#1080#1089#1082#1072
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 270
    Top = 8
    Width = 137
    Height = 41
    Caption = #1040#1085#1072#1083#1080#1079#1080#1088#1086#1074#1072#1090#1100' '#1076#1080#1089#1082
    OnClick = SpeedButton2Click
  end
  object OKBtn: TButton
    Left = 24
    Top = 130
    Width = 75
    Height = 25
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 105
    Top = 130
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 18
    Top = 28
    Width = 156
    Height = 21
    EditLabel.Width = 85
    EditLabel.Height = 13
    EditLabel.Caption = #1041#1091#1082#1074#1072' '#1076#1080#1089#1082#1072' (C:)'
    TabOrder = 2
  end
  object CheckBox1: TCheckBox
    Left = 18
    Top = 95
    Width = 169
    Height = 17
    Caption = #1048#1085#1076#1077#1082#1089#1072#1094#1080#1103' '#1092#1072#1081#1083#1086#1074' '#1085#1072' '#1076#1080#1089#1082#1077
    TabOrder = 3
  end
  object LabeledEdit2: TLabeledEdit
    Left = 18
    Top = 68
    Width = 156
    Height = 21
    EditLabel.Width = 135
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1090#1086#1084#1072' ('#1076#1086' 32 '#1089#1080#1084#1074#1086#1083#1086#1074')'
    TabOrder = 4
  end
end
