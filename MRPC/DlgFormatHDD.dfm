object OKRightDlg12345: TOKRightDlg12345
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1060#1086#1088#1084#1072#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1090#1086#1084#1072
  ClientHeight = 199
  ClientWidth = 176
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 13
    Top = 5
    Width = 93
    Height = 13
    Caption = #1060#1072#1081#1083#1086#1074#1072#1103' '#1089#1080#1089#1090#1077#1084#1072
  end
  object OKBtn: TButton
    Left = 91
    Top = 158
    Width = 75
    Height = 25
    Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 6
    Top = 158
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object ComboBox1: TComboBox
    Left = 13
    Top = 24
    Width = 145
    Height = 21
    ItemIndex = 0
    TabOrder = 2
    Text = 'NTFS'
    Items.Strings = (
      'NTFS'
      'FAT32'
      'FAT')
  end
  object CheckBox1: TCheckBox
    Left = 13
    Top = 51
    Width = 153
    Height = 17
    Caption = #1041#1099#1089#1090#1088#1086#1077' '#1092#1086#1088#1084#1072#1090#1080#1088#1086#1074#1072#1085#1080#1077
    TabOrder = 3
  end
  object LabeledEdit1: TLabeledEdit
    Left = 13
    Top = 88
    Width = 145
    Height = 21
    EditLabel.Width = 85
    EditLabel.Height = 13
    EditLabel.Caption = #1056#1072#1079#1084#1077#1088' '#1082#1083#1072#1089#1090#1077#1088#1072
    NumbersOnly = True
    TabOrder = 4
    Text = '4096'
  end
  object LabeledEdit2: TLabeledEdit
    Left = 13
    Top = 128
    Width = 145
    Height = 21
    EditLabel.Width = 59
    EditLabel.Height = 13
    EditLabel.Caption = #1052#1077#1090#1082#1072' '#1090#1086#1084#1072
    TabOrder = 5
    Text = #1053#1086#1074#1099#1081' '#1090#1086#1084
  end
end
