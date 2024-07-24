object Form12: TForm12
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1055#1086#1080#1089#1082' '#1087#1088#1086#1094#1077#1089#1089#1072
  ClientHeight = 113
  ClientWidth = 340
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 250
    Top = 79
    Width = 75
    Height = 25
    Caption = #1053#1072#1081#1090#1080
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 160
    Top = 78
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    TabOrder = 1
    OnClick = Button2Click
  end
  object LabeledEdit1: TLabeledEdit
    Left = 8
    Top = 18
    Width = 317
    Height = 21
    EditLabel.Width = 154
    EditLabel.Height = 13
    EditLabel.Caption = #1048#1084#1103' '#1080#1083#1080' '#1095#1072#1089#1090#1100' '#1080#1084#1077#1085#1080' '#1087#1088#1086#1094#1077#1089#1089#1072
    TabOrder = 2
  end
  object Button3: TButton
    Left = 250
    Top = 48
    Width = 75
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 3
    OnClick = Button3Click
  end
  object ComboBox1: TComboBox
    Left = 8
    Top = 48
    Width = 227
    Height = 21
    TabOrder = 4
    Text = #1047#1072#1075#1088#1091#1079#1080#1090#1077' '#1089#1087#1080#1089#1086#1082' '#1087#1088#1086#1094#1077#1089#1089#1086#1074
    OnSelect = ComboBox1Select
  end
end
