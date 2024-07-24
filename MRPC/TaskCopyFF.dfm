object TaskCopyDelFF: TTaskCopyDelFF
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
  ClientHeight = 121
  ClientWidth = 415
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object EditSource: TLabeledEdit
    Left = 8
    Top = 17
    Width = 311
    Height = 21
    EditLabel.Width = 112
    EditLabel.Height = 13
    EditLabel.Caption = #1044#1080#1088#1077#1082#1090#1086#1088#1080#1103' '#1080#1089#1090#1086#1095#1085#1080#1082
    TabOrder = 0
  end
  object EditDestination: TLabeledEdit
    Left = 8
    Top = 57
    Width = 311
    Height = 21
    EditLabel.Width = 228
    EditLabel.Height = 13
    EditLabel.Caption = #1044#1080#1088#1077#1082#1090#1086#1088#1080#1103' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1077' ('#1087#1088#1080#1084#1077#1088' - C$\TEMP\)'
    TabOrder = 1
    Text = 'C$\TEMP\'
  end
  object Button1: TButton
    Left = 325
    Top = 15
    Width = 75
    Height = 25
    Caption = #1054#1073#1079#1086#1088
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 325
    Top = 55
    Width = 75
    Height = 25
    Caption = #1054#1073#1079#1086#1088
    TabOrder = 3
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 288
    Top = 89
    Width = 112
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1079#1072#1076#1072#1085#1080#1077
    TabOrder = 4
    OnClick = Button3Click
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 90
    Width = 249
    Height = 17
    Caption = #1055#1088#1086#1074#1077#1088#1103#1090#1100' '#1080' '#1089#1086#1079#1076#1072#1074#1072#1090#1100' '#1082#1072#1090#1072#1083#1086#1075' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1103
    Checked = True
    State = cbChecked
    TabOrder = 5
  end
end
