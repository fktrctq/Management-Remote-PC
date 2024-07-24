object FormOneStart: TFormOneStart
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1042#1099#1073#1086#1088' '#1089#1087#1080#1089#1082#1072' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1086#1074
  ClientHeight = 441
  ClientWidth = 356
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object RadioGroup1: TRadioGroup
    Left = 0
    Top = 0
    Width = 356
    Height = 68
    Align = alTop
    Caption = #1058#1080#1087' '#1089#1077#1090#1080
    Items.Strings = (
      #1059' '#1084#1077#1085#1103' '#1089#1077#1090#1100' '#1089' Active Directory'
      #1059' '#1084#1077#1085#1103' '#1086#1076#1085#1086#1088#1072#1085#1075#1086#1074#1072#1103' '#1089#1077#1090#1100)
    TabOrder = 0
    OnClick = RadioGroup1Click
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 230
    Width = 356
    Height = 158
    Align = alTop
    Caption = #1059#1095#1077#1090#1085#1072#1103' '#1079#1072#1087#1080#1089#1100' '#1072#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088#1072
    TabOrder = 1
    object Label2: TLabel
      AlignWithMargins = True
      Left = 5
      Top = 75
      Width = 346
      Height = 78
      Align = alBottom
      Alignment = taCenter
      Caption = 
        #1059#1095#1077#1090#1085#1072#1103' '#1079#1072#1087#1080#1089#1100' '#1040#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088#1072' '#1074' '#1042#1072#1096#1077#1081' '#1083#1086#1082#1072#1083#1100#1085#1086#1081' '#1089#1077#1090#1080'. '#1045#1089#1083#1080' '#1091' '#1042#1072#1089 +
        ' '#1089#1077#1090#1100' '#1089' AD '#1080' '#1087#1088#1086#1075#1088#1072#1084#1084#1072' '#1079#1072#1087#1091#1097#1077#1085#1072' '#1086#1090' '#1080#1084#1077#1085#1080' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' '#1080#1084#1077#1102#1097#1077#1075#1086' '#1087 +
        #1088#1072#1074#1072' '#1072#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088#1072' '#1074#1072#1096#1077#1075#1086' '#1076#1086#1084#1077#1085#1072', '#1076#1072#1085#1085#1099#1077' '#1084#1086#1078#1085#1086' '#1085#1077' '#1091#1082#1072#1079#1074#1072#1090#1100'. '#1042#1089#1077 +
        ' '#1086#1087#1077#1088#1072#1094#1080#1080' '#1073#1091#1076#1091#1090' '#1074#1099#1087#1086#1083#1085#1103#1090#1089#1103' '#1089' '#1087#1088#1072#1074#1072#1084#1080' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103'.  '#1058#1072#1082 +
        #1078#1077' '#1080#1084#1103' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' '#1080' '#1087#1072#1088#1086#1083#1100' '#1084#1086#1078#1085#1086' '#1091#1082#1072#1079#1072#1090#1100' '#1074' '#1085#1072#1089#1090#1088#1086#1081#1082#1072#1093' '#1087#1088#1086#1075#1088#1072#1084#1084 +
        #1099'.'
      Layout = tlCenter
      WordWrap = True
      ExplicitWidth = 344
    end
    object LabeledEdit1: TLabeledEdit
      Left = 10
      Top = 49
      Width = 229
      Height = 21
      EditLabel.Width = 3
      EditLabel.Height = 13
      LabelPosition = lpLeft
      PasswordChar = #7
      TabOrder = 1
      TextHint = #1055#1072#1088#1086#1083#1100
    end
    object LabeledEdit2: TLabeledEdit
      Left = 10
      Top = 22
      Width = 229
      Height = 21
      BiDiMode = bdLeftToRight
      EditLabel.Width = 3
      EditLabel.Height = 13
      EditLabel.BiDiMode = bdLeftToRight
      EditLabel.ParentBiDiMode = False
      LabelPosition = lpLeft
      ParentBiDiMode = False
      TabOrder = 0
      TextHint = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 388
    Width = 356
    Height = 53
    Align = alTop
    TabOrder = 2
    object Button1: TButton
      Left = 240
      Top = 15
      Width = 97
      Height = 25
      Caption = 'OK'
      TabOrder = 0
      OnClick = Button1Click
    end
    object CheckBox1: TCheckBox
      Left = 10
      Top = 6
      Width = 199
      Height = 17
      Caption = #1054#1090#1086#1073#1088#1072#1078#1072#1090#1100' '#1076#1080#1072#1083#1086#1075' '#1087#1088#1080' '#1079#1072#1075#1088#1091#1079#1082#1077
      TabOrder = 1
      OnMouseUp = CheckBox1MouseUp
    end
    object CheckBox2: TCheckBox
      Left = 10
      Top = 26
      Width = 199
      Height = 17
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1080#1079' '#1041#1044
      TabOrder = 2
      OnMouseUp = CheckBox2MouseUp
    end
  end
  object GroupBox3: TGroupBox
    Left = 0
    Top = 68
    Width = 356
    Height = 162
    Align = alTop
    Caption = #1057#1087#1080#1089#1082#1080' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1086#1074
    TabOrder = 3
    object Label3: TLabel
      AlignWithMargins = True
      Left = 10
      Top = 88
      Width = 330
      Height = 39
      Alignment = taCenter
      Caption = 
        #1044#1083#1103' '#1074#1099#1073#1086#1088#1072' '#1089#1087#1080#1089#1082#1072' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1086#1074' '#1091#1082#1072#1078#1080#1090#1077' '#1075#1088#1091#1087#1087#1091' '#1073#1077#1079#1086#1087#1072#1089#1085#1086#1089#1090#1080' Active' +
        ' Directory. '#1042#1053#1048#1052#1040#1053#1048#1045'!!! '#1047#1072#1075#1088#1091#1079#1082#1072' '#1089#1087#1080#1089#1082#1072' '#1080#1079' '#1090#1099#1089#1103#1095' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1086#1074' '#1084#1086#1078 +
        #1077#1090' '#1079#1072#1085#1103#1090#1100' '#1087#1088#1086#1076#1086#1083#1078#1080#1090#1077#1083#1100#1085#1086#1077' '#1074#1088#1077#1084#1103'. '
      WordWrap = True
    end
    object ComboGroup: TComboBox
      Left = 7
      Top = 133
      Width = 335
      Height = 21
      DropDownCount = 15
      Enabled = False
      TabOrder = 0
      TextHint = #1043#1088#1091#1087#1087#1099' '#1073#1077#1079#1086#1087#1072#1089#1085#1086#1089#1090#1080' Active Directory'
      OnKeyDown = ComboGroupKeyDown
    end
    object RadioGroup2: TRadioGroup
      Left = 2
      Top = 15
      Width = 352
      Height = 67
      Align = alTop
      Enabled = False
      Items.Strings = (
        #1044#1080#1072#1087#1072#1079#1086#1085' IP '#1076#1088#1077#1089#1086#1074
        #1057#1087#1080#1089#1086#1082' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1086#1074' '#1080#1079' '#1090#1077#1082#1089#1090#1086#1074#1086#1075#1086' '#1092#1072#1081#1083#1072)
      TabOrder = 1
      OnClick = RadioGroup2Click
    end
  end
end
