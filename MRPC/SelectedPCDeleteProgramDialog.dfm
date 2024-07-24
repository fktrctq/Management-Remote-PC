object OKRightDlg12345678910111213141516171819202122: TOKRightDlg12345678910111213141516171819202122
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1043#1088#1091#1087#1087#1086#1074#1086#1077' '#1091#1076#1072#1083#1077#1085#1080#1077' '#1087#1088#1086#1075#1088#1072#1084#1084' (msi)'
  ClientHeight = 155
  ClientWidth = 384
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 69
    Width = 86
    Height = 13
    Caption = #1057#1087#1080#1089#1086#1082' '#1087#1088#1086#1075#1088#1072#1084#1084
  end
  object Label3: TLabel
    Left = 170
    Top = 122
    Width = 41
    Height = 19
    Caption = '0%'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    Visible = False
  end
  object OKBtn: TButton
    Left = 300
    Top = 122
    Width = 71
    Height = 25
    Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 223
    Top = 122
    Width = 71
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 8
    Top = 21
    Width = 368
    Height = 21
    EditLabel.Width = 358
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1080#1083#1080' '#1095#1072#1089#1090#1100' '#1080#1084#1077#1085#1080' '#1087#1088#1086#1075#1088#1072#1084#1084#1099' '#1076#1083#1103' '#1091#1076#1072#1083#1077#1085#1080#1103' '#1087#1088#1080' '#1089#1086#1074#1087#1072#1076#1077#1085#1080#1080' '#1080#1084#1077#1085#1080'.'
    TabOrder = 2
  end
  object Button1: TButton
    Left = 301
    Top = 85
    Width = 71
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 3
    OnClick = Button1Click
  end
  object ComboBox1: TComboBox
    Left = 8
    Top = 85
    Width = 287
    Height = 21
    CharCase = ecLowerCase
    DropDownCount = 10
    Sorted = True
    TabOrder = 4
    OnSelect = ComboBox1Select
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 122
    Width = 153
    Height = 17
    Caption = #1047#1072#1087#1091#1089#1090#1080#1090#1100' '#1074' '#1086#1076#1085#1086#1084' '#1087#1086#1090#1086#1082#1077
    Checked = True
    State = cbChecked
    TabOrder = 5
  end
  object CheckBox2: TCheckBox
    Left = 8
    Top = 48
    Width = 217
    Height = 17
    Caption = #1058#1086#1095#1085#1086#1077' '#1089#1086#1074#1087#1072#1076#1077#1085#1080#1077' '#1080#1084#1077#1085#1080' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
    Checked = True
    State = cbChecked
    TabOrder = 6
    OnClick = CheckBox2Click
  end
end
