object OKRightDlg12345678: TOKRightDlg12345678
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1042#1099#1074#1086#1076' '#1080#1079' '#1076#1086#1084#1077#1085#1072
  ClientHeight = 215
  ClientWidth = 198
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 0
    Top = 62
    Width = 198
    Height = 153
    Align = alClient
    TabOrder = 1
    object CancelBtn: TButton
      Left = 8
      Top = 114
      Width = 75
      Height = 25
      Cancel = True
      Caption = #1054#1090#1084#1077#1085#1072
      ModalResult = 2
      TabOrder = 4
      OnClick = CancelBtnClick
    end
    object EditPassAdmin: TLabeledEdit
      Left = 8
      Top = 63
      Width = 176
      Height = 21
      EditLabel.Width = 37
      EditLabel.Height = 13
      EditLabel.Caption = #1055#1072#1088#1086#1083#1100
      PasswordChar = #7
      TabOrder = 1
    end
    object EditUserAdmin: TLabeledEdit
      Left = 8
      Top = 23
      Width = 176
      Height = 21
      EditLabel.Width = 72
      EditLabel.Height = 13
      EditLabel.Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
      TabOrder = 0
    end
    object OKBtn: TButton
      Left = 109
      Top = 114
      Width = 75
      Height = 25
      Caption = 'OK'
      Default = True
      ModalResult = 1
      TabOrder = 3
      OnClick = OKBtnClick
    end
    object RebootAfter: TCheckBox
      Left = 8
      Top = 91
      Width = 178
      Height = 17
      Caption = #1055#1077#1088#1077#1079#1072#1075#1088#1091#1079#1080#1090#1100' '#1082#1086#1084#1087#1100#1102#1090#1077#1088
      Checked = True
      State = cbChecked
      TabOrder = 2
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 0
    Width = 198
    Height = 62
    Align = alTop
    TabOrder = 0
    object CheckBox1: TCheckBox
      Left = 8
      Top = 4
      Width = 176
      Height = 13
      Caption = #1057#1084#1077#1085#1080#1090#1100' '#1080#1084#1103' '#1088#1072#1073#1086#1095#1077#1081' '#1075#1088#1091#1087#1087#1099
      TabOrder = 0
      OnClick = CheckBox1Click
    end
    object EditNameGroup: TLabeledEdit
      Left = 8
      Top = 34
      Width = 176
      Height = 21
      EditLabel.Width = 92
      EditLabel.Height = 13
      EditLabel.Caption = #1053#1086#1074#1086#1077' '#1080#1084#1103' '#1075#1088#1091#1087#1087#1099
      Enabled = False
      TabOrder = 1
    end
  end
end
