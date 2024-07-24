object OKRightDlg: TOKRightDlg
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1053#1086#1074#1072#1103' '#1089#1083#1091#1078#1073#1072
  ClientHeight = 354
  ClientWidth = 331
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 331
    Height = 354
    Align = alClient
    Shape = bsFrame
    ExplicitHeight = 356
  end
  object Label1: TLabel
    Left = 19
    Top = 135
    Width = 60
    Height = 13
    Caption = #1058#1080#1087' '#1089#1083#1091#1078#1073#1099
  end
  object Label2: TLabel
    Left = 19
    Top = 174
    Width = 61
    Height = 13
    Caption = #1058#1080#1087' '#1079#1072#1087#1091#1089#1082#1072
  end
  object OKBtn: TButton
    Left = 237
    Top = 316
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 146
    Top = 316
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 16
    Top = 25
    Width = 296
    Height = 21
    EditLabel.Width = 161
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1085#1086#1074#1086#1081' '#1089#1083#1091#1078#1073#1099' (AdobeFlash)'
    TabOrder = 2
  end
  object LabeledEdit2: TLabeledEdit
    Left = 16
    Top = 66
    Width = 296
    Height = 21
    EditLabel.Width = 200
    EditLabel.Height = 13
    EditLabel.Caption = #1054#1090#1086#1073#1088#1072#1078#1072#1077#1084#1086#1077' '#1080#1084#1103' (Adobe Flash Player)'
    TabOrder = 3
  end
  object LabeledEdit3: TLabeledEdit
    Left = 16
    Top = 108
    Width = 296
    Height = 21
    EditLabel.Width = 207
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1089#1087#1086#1083#1085#1103#1077#1084#1099#1081' '#1092#1072#1081#1083' ('#1085#1077' '#1079#1072#1073#1099#1074#1072#1077#1084' '#1087#1088#1086' " ")'
    TabOrder = 4
  end
  object ComboBox1: TComboBox
    Left = 16
    Top = 152
    Width = 296
    Height = 21
    ItemIndex = 4
    TabOrder = 5
    Text = '16   | Own Process'
    Items.Strings = (
      '1     | Kernel Driver'
      '2     | File System Driver'
      '4     | Adapter'
      '8     | Recognizer Driver'
      '16   | Own Process'
      '32   | Share Process'
      '256 | Interactive Process')
  end
  object ComboBox2: TComboBox
    Left = 16
    Top = 190
    Width = 296
    Height = 21
    ItemIndex = 3
    TabOrder = 6
    Text = 'Manual'
    Items.Strings = (
      'Boot'
      'System'
      'Automatic'
      'Manual'
      'Disabled')
  end
  object CheckBox1: TCheckBox
    Left = 16
    Top = 216
    Width = 296
    Height = 17
    Caption = #1088#1072#1079#1088#1077#1096#1080#1090#1100' '#1074#1079#1072#1080#1084#1086#1076#1077#1081#1089#1090#1074#1080#1077' '#1089' '#1088#1072#1073#1086#1095#1080#1084' '#1089#1090#1086#1083#1086#1084
    TabOrder = 7
  end
  object LabeledEdit4: TLabeledEdit
    Left = 16
    Top = 251
    Width = 185
    Height = 21
    EditLabel.Width = 80
    EditLabel.Height = 13
    EditLabel.Caption = #1059#1095#1077#1090#1085#1072#1103' '#1079#1072#1087#1080#1089#1100
    TabOrder = 8
    Text = '.\LocalSystem'
  end
  object LabeledEdit5: TLabeledEdit
    Left = 16
    Top = 289
    Width = 185
    Height = 21
    EditLabel.Width = 37
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1072#1088#1086#1083#1100
    PasswordChar = #8
    TabOrder = 9
  end
end
