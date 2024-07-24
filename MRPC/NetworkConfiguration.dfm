object OKRightDlg123456789101112131415: TOKRightDlg123456789101112131415
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1057#1074#1086#1081#1089#1090#1074#1072
  ClientHeight = 342
  ClientWidth = 316
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 8
    Top = 293
    Width = 98
    Height = 25
    Caption = #1044#1086#1087#1086#1083#1085#1080#1090#1077#1083#1100#1085#1086
    OnClick = SpeedButton1Click
  end
  object OKBtn: TButton
    Left = 144
    Top = 293
    Width = 75
    Height = 25
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 230
    Top = 293
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    ModalResult = 2
    TabOrder = 1
  end
  object Panel1: TPanel
    Left = 8
    Top = 45
    Width = 297
    Height = 103
    TabOrder = 2
    object Label1: TLabel
      Left = 7
      Top = 22
      Width = 47
      Height = 13
      Caption = 'IP '#1072#1076#1088#1077#1089':'
    end
    object Label2: TLabel
      Left = 7
      Top = 49
      Width = 80
      Height = 13
      Caption = #1052#1072#1089#1082#1072' '#1087#1086#1076#1089#1077#1090#1080':'
    end
    object Label3: TLabel
      Left = 7
      Top = 76
      Width = 84
      Height = 13
      Caption = #1054#1089#1085#1086#1074#1085#1086#1081' '#1096#1083#1102#1079':'
    end
    object MaskEdit1: TMaskEdit
      Left = 170
      Top = 19
      Width = 114
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      Text = ''
      OnKeyPress = MaskEdit1KeyPress
    end
    object MaskEdit2: TMaskEdit
      Left = 170
      Top = 46
      Width = 114
      Height = 21
      TabOrder = 1
      Text = ''
      OnKeyPress = MaskEdit2KeyPress
    end
    object MaskEdit3: TMaskEdit
      Left = 170
      Top = 73
      Width = 114
      Height = 21
      TabOrder = 2
      Text = ''
      OnKeyPress = MaskEdit3KeyPress
    end
  end
  object CheckBox1: TCheckBox
    Left = 15
    Top = 12
    Width = 201
    Height = 17
    Caption = #1055#1086#1083#1091#1095#1080#1090#1100' IP '#1072#1076#1088#1077#1089' '#1072#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1080
    TabOrder = 3
    OnClick = CheckBox1Click
  end
  object CheckBox3: TCheckBox
    Left = 15
    Top = 166
    Width = 263
    Height = 16
    Caption = #1055#1086#1083#1091#1095#1080#1090#1100' '#1072#1076#1088#1077#1089' DNS-'#1089#1077#1088#1074#1077#1088#1072' '#1072#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1080
    TabOrder = 4
    OnClick = CheckBox3Click
  end
  object Panel2: TPanel
    Left = 8
    Top = 200
    Width = 297
    Height = 70
    TabOrder = 5
    object Label4: TLabel
      Left = 7
      Top = 13
      Width = 150
      Height = 13
      Caption = #1055#1088#1077#1076#1087#1086#1095#1080#1090#1072#1077#1084#1099#1081' DNS-'#1089#1077#1088#1074#1077#1088
    end
    object Label5: TLabel
      Left = 7
      Top = 40
      Width = 149
      Height = 13
      Caption = #1040#1083#1100#1090#1077#1088#1085#1072#1090#1080#1074#1085#1099#1081' DNS-'#1089#1077#1088#1074#1077#1088
    end
    object MaskEdit4: TMaskEdit
      Left = 170
      Top = 10
      Width = 114
      Height = 21
      TabOrder = 0
      Text = ''
      OnKeyPress = MaskEdit4KeyPress
    end
    object MaskEdit5: TMaskEdit
      Left = 170
      Top = 37
      Width = 114
      Height = 21
      BiDiMode = bdLeftToRight
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBiDiMode = False
      ParentFont = False
      TabOrder = 1
      Text = ''
      OnKeyPress = MaskEdit5KeyPress
    end
  end
  object CheckBox4: TCheckBox
    Left = 15
    Top = 188
    Width = 277
    Height = 17
    Caption = #1048#1089#1087#1086#1083#1100#1079#1086#1074#1072#1090#1100' '#1089#1083#1077#1076#1091#1102#1097#1080#1077' '#1072#1076#1088#1077#1089#1072' DNS-'#1089#1077#1088#1074#1077#1088#1086#1074':'
    TabOrder = 6
    OnClick = CheckBox4Click
  end
  object CheckBox2: TCheckBox
    Left = 15
    Top = 35
    Width = 201
    Height = 17
    Caption = #1048#1089#1087#1086#1083#1100#1079#1086#1074#1072#1090#1100' '#1089#1083#1077#1076#1091#1102#1097#1080#1081' IP '#1072#1076#1088#1077#1089':'
    TabOrder = 7
    OnClick = CheckBox2Click
  end
end
