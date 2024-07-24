object OKRightDlg2: TOKRightDlg2
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1072#1074#1090#1086#1079#1072#1075#1088#1091#1079#1082#1091
  ClientHeight = 233
  ClientWidth = 269
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 269
    Height = 233
    Align = alClient
    Shape = bsFrame
    ExplicitLeft = -1
    ExplicitWidth = 294
  end
  object Label1: TLabel
    Left = 14
    Top = 89
    Width = 76
    Height = 13
    Caption = #1042#1077#1090#1092#1100' '#1088#1077#1077#1089#1090#1088#1072
  end
  object OKBtn: TButton
    Left = 182
    Top = 195
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 101
    Top = 195
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 14
    Top = 21
    Width = 241
    Height = 21
    EditLabel.Width = 19
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103
    TabOrder = 2
  end
  object LabeledEdit2: TLabeledEdit
    Left = 14
    Top = 63
    Width = 241
    Height = 21
    EditLabel.Width = 152
    EditLabel.Height = 13
    EditLabel.Caption = #1047#1085#1072#1095#1077#1085#1080#1077' ('#1082#1086#1084#1072#1085#1076#1085#1072#1103' '#1089#1090#1088#1086#1082#1072')'
    TabOrder = 3
  end
  object ComboBox1: TComboBox
    Left = 14
    Top = 106
    Width = 241
    Height = 21
    ItemIndex = 0
    TabOrder = 4
    Text = 'HKEY_CLASSES_ROOT'
    Items.Strings = (
      'HKEY_CLASSES_ROOT'
      'HKEY_CURRENT_USER'
      'HKEY_LOCAL_MACHINE'
      'HKEY_USERS'
      'HKEY_CURRENT_CONFIG')
  end
  object LabeledEdit3: TLabeledEdit
    Left = 14
    Top = 160
    Width = 241
    Height = 21
    EditLabel.Width = 80
    EditLabel.Height = 13
    EditLabel.Caption = #1056#1072#1079#1076#1077#1083' '#1088#1077#1077#1089#1090#1088#1072
    TabOrder = 5
  end
end
