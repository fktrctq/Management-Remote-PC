object OKRightDlg123456: TOKRightDlg123456
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1076#1086#1084#1077#1085
  ClientHeight = 257
  ClientWidth = 189
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 106
    Top = 224
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 8
    Top = 224
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 8
    Top = 22
    Width = 169
    Height = 21
    EditLabel.Width = 165
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1076#1086#1084#1077#1085#1072' '#1080#1083#1080' '#1088#1072#1073#1086#1095#1077#1081' '#1075#1088#1091#1087#1087#1099
    TabOrder = 2
  end
  object LabeledEdit2: TLabeledEdit
    Left = 8
    Top = 64
    Width = 169
    Height = 21
    EditLabel.Width = 155
    EditLabel.Height = 13
    EditLabel.Caption = #1044#1086#1084#1077#1085'\'#1040#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088' '#1076#1086#1084#1077#1085#1072
    TabOrder = 3
  end
  object LabeledEdit3: TLabeledEdit
    Left = 8
    Top = 102
    Width = 169
    Height = 21
    EditLabel.Width = 37
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1072#1088#1086#1083#1100
    PasswordChar = #7
    TabOrder = 4
  end
  object ComboBox1: TComboBox
    Left = 8
    Top = 135
    Width = 169
    Height = 21
    ItemIndex = 0
    TabOrder = 5
    Text = 'NETSETUP_JOIN_DOMAIN+NETSETUP_ACCT_CREATE'
    OnSelect = ComboBox1Select
    Items.Strings = (
      'NETSETUP_JOIN_DOMAIN+NETSETUP_ACCT_CREATE'
      'NETSETUP_JOIN_DOMAIN'
      'NETSETUP_ACCT_CREATE'
      'NETSETUP_WIN9X_UPGRADE'
      'NETSETUP_DOMAIN_JOIN_IF_JOINED'
      'NETSETUP_JOIN_UNSECURE'
      #1053#1077#1090' '#1074#1072#1088#1080#1072#1085#1090#1086#1074' (0)')
  end
  object RebootAfter: TCheckBox
    Left = 8
    Top = 194
    Width = 169
    Height = 17
    Caption = #1055#1077#1088#1077#1079#1072#1075#1088#1091#1079#1080#1090#1100' '#1082#1086#1084#1087#1100#1102#1090#1077#1088
    Checked = True
    State = cbChecked
    TabOrder = 6
  end
  object NumOper: TEdit
    Left = 8
    Top = 167
    Width = 169
    Height = 21
    NumbersOnly = True
    TabOrder = 7
    TextHint = 'FJoinOptions'
  end
end
