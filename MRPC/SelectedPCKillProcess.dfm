object OKRightDlg123456789101112131415161718192021: TOKRightDlg123456789101112131415161718192021
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1043#1088#1091#1087#1087#1086#1074#1086#1077' '#1079#1072#1074#1077#1088#1096#1077#1085#1080#1077' '#1087#1088#1086#1094#1077#1089#1089#1072
  ClientHeight = 89
  ClientWidth = 338
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 268
    Top = 17
    Width = 25
    Height = 20
    Caption = '0%'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = 20
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    Visible = False
  end
  object OKBtn: TButton
    Left = 254
    Top = 56
    Width = 75
    Height = 25
    Caption = #1047#1072#1074#1077#1088#1096#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 173
    Top = 56
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 7
    Top = 19
    Width = 242
    Height = 21
    EditLabel.Width = 143
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1079#1072#1074#1077#1088#1096#1072#1077#1084#1086#1075#1086' '#1087#1088#1086#1094#1077#1089#1089#1072
    TabOrder = 2
    Text = 'notepad.exe'
  end
  object CheckBox1: TCheckBox
    Left = 6
    Top = 60
    Width = 165
    Height = 17
    Caption = #1047#1072#1074#1077#1088#1096#1077#1085#1080#1077' '#1074' '#1086#1076#1085#1086#1084' '#1087#1086#1090#1086#1082#1077
    Checked = True
    State = cbChecked
    TabOrder = 3
    OnClick = CheckBox1Click
  end
end
