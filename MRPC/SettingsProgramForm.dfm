object SettingsProgram: TSettingsProgram
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1072
  ClientHeight = 548
  ClientWidth = 413
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010001001010000001002000680400001600000028000000100000002000
    0000010020000000000000040000C30E0000C30E000000000000000000000000
    0000333333003333330A33333364333333793333337733333377333333773333
    3377333333773333337733333379333333643333330A33333300000000000000
    0000333333003333330933333393333333CA333333C7333333C5333333C43333
    33C4333333C5333333C7333333CA3333339A3333330B33333300000000000000
    0000333333003333330033333304333333093333330C33333369333333973333
    3397333333683333330C33333309333333053333330033333300000000003333
    3311333333463333335633333355333333543333335C333333D9333333FF3333
    33FF333333D93333335C33333354333333553333335633333348333333123333
    33A9333333F8333333FE333333FD333333FD333333FE333333FF333333FF3333
    33FF333333FF333333FE333333FD333333FF333333E9333333B0333333A63333
    33FB333333F9333333F2333333F2333333F2333333F2333333F2343131EF3332
    31F0343131EF343131EF333333F2333333F3333333E2333333C4333333F93333
    33FC3333337D333333342B2E353525293733252936332C3137361781A2790E9B
    C6C60F9AC4C0197B9A6F34302F3633333337333333353333337F333333FC3333
    33F733333347292C3600F8B90039FBBB0040FBBA0045F2B804660DC5F2A900C5
    FFFF00C5FFFF00C5FFB500C9FF0800C9FF003333330033333347333333F73333
    33F7333333478E711B00F1B4006BF1B400F6F1B400F9F5B400E041BFBA8F00C3
    FFFD00C3FFFF00C3FFE700C3FF2A00C3FF003333330033333347333333F73333
    33F733333347856B1D00F1B40033F1B400EDF1B400FFF4B400FAA2B9548900C3
    FFA400C3E59600C3E9A100C3F33C00C3F4003333330033333347333333F73333
    33F73333334751472B00EBB2070BE7B00C8FF1B401CBF1B402BDC1A7385900C2
    998500C38FE100C390E100C3928100C390023139370033333347333333F73333
    33F733333347333334003E6FE400306AF2623D6EE3A8386CE9A42B66FBAB06B7
    9F9D00C390FF00C391FF00C391D300C39117255A4C0033333347333333F73333
    33F733333347333333002C68F8002C68F84E2B68F9F92C68F9FF2D67FAF01598
    C18900C48FEC00C391E100C391DD00C391401B765F0033333347333333F73333
    33F733333347333333002C6AFD002C6AFD1C2C69FBD22C69FBFF2C68FCFB2577
    ED7D00D1913F00D89F1B00D29B2500CD97192952470033333347333333F73333
    33ED33333389333333443333334733343748304A8A7F2F53ABB62F52A6AA3141
    6961362B2B43342F3045352E3044343031453333334433333389333333ED3333
    3383333333EE333333F8333333F7333333F7333230F433322FF233322FF23333
    32F5333333F7333333F7333333F7333333F7333333F8333333EE33333383C003
    0000C0030000E007000000000000000000000000000000000000200C0000200C
    0000200C0000200400003004000030040000300400000000000000000000}
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 413
    Height = 507
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    ExplicitHeight = 483
    object TabSheet1: TTabSheet
      Caption = #1054#1089#1085#1086#1074#1085#1099#1077
      ExplicitLeft = 0
      ExplicitTop = 22
      object Label1: TLabel
        Left = 5
        Top = 319
        Width = 96
        Height = 13
        Caption = #1057#1090#1080#1083#1100' '#1086#1092#1086#1088#1084#1083#1077#1085#1080#1103
      end
      object GroupBox1: TGroupBox
        Left = 0
        Top = 210
        Width = 405
        Height = 103
        Align = alTop
        Caption = 'uRDM'
        TabOrder = 0
        ExplicitTop = 189
        object LabeledEdit2: TLabeledEdit
          Left = 173
          Top = 34
          Width = 150
          Height = 21
          EditLabel.Width = 83
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1072#1088#1086#1083#1100' (Default)'
          PasswordChar = #7
          TabOrder = 0
        end
        object LabeledEdit1: TLabeledEdit
          Left = 3
          Top = 34
          Width = 150
          Height = 21
          EditLabel.Width = 118
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100' (Default)'
          TabOrder = 1
        end
        object LabeledEdit3: TLabeledEdit
          Left = 173
          Top = 75
          Width = 150
          Height = 21
          EditLabel.Width = 130
          EditLabel.Height = 13
          EditLabel.Caption = #1051#1086#1082#1072#1083#1100#1085#1099#1081' '#1087#1086#1088#1090' (Default)'
          NumbersOnly = True
          TabOrder = 2
          Visible = False
        end
        object LabeledEdit4: TLabeledEdit
          Left = 3
          Top = 75
          Width = 150
          Height = 21
          EditLabel.Width = 140
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1086#1088#1090' '#1089#1077#1088#1074#1077#1088#1072' (uRDMServer)'
          NumbersOnly = True
          TabOrder = 3
        end
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 86
        Width = 405
        Height = 124
        Align = alTop
        Caption = #1054#1090#1086#1073#1088#1072#1078#1077#1085#1080#1077' '#1074#1082#1083#1072#1076#1086#1082
        DoubleBuffered = False
        ParentDoubleBuffered = False
        TabOrder = 1
        ExplicitTop = 65
        object CheckBox1: TCheckBox
          Left = 5
          Top = 17
          Width = 97
          Height = 17
          Caption = #1055#1088#1086#1094#1077#1089#1089#1099
          Checked = True
          Enabled = False
          State = cbChecked
          TabOrder = 0
        end
        object CheckBox2: TCheckBox
          Left = 5
          Top = 35
          Width = 97
          Height = 17
          Caption = #1057#1083#1091#1078#1073#1099
          TabOrder = 1
        end
        object CheckBox3: TCheckBox
          Left = 5
          Top = 53
          Width = 97
          Height = 17
          Caption = #1040#1074#1090#1086#1079#1072#1075#1088#1091#1079#1082#1072
          TabOrder = 2
        end
        object CheckBox4: TCheckBox
          Left = 5
          Top = 72
          Width = 97
          Height = 17
          Caption = #1055#1088#1086#1075#1088#1072#1084#1084#1099
          TabOrder = 3
        end
        object CheckBox5: TCheckBox
          Left = 5
          Top = 90
          Width = 97
          Height = 17
          Caption = #1042#1089#1077' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
          TabOrder = 4
        end
        object CheckBox6: TCheckBox
          Left = 148
          Top = 17
          Width = 97
          Height = 17
          Caption = #1044#1080#1089#1082#1080
          TabOrder = 5
        end
        object CheckBox7: TCheckBox
          Left = 148
          Top = 35
          Width = 97
          Height = 17
          Caption = #1055#1088#1080#1085#1090#1077#1088#1099
          TabOrder = 6
        end
        object CheckBox8: TCheckBox
          Left = 148
          Top = 53
          Width = 97
          Height = 17
          Caption = #1055#1088#1086#1092#1080#1083#1080' (Users)'
          TabOrder = 7
        end
        object CheckBox9: TCheckBox
          Left = 148
          Top = 72
          Width = 97
          Height = 17
          Caption = #1044#1088#1072#1081#1074#1077#1088#1072' PnP'
          TabOrder = 8
        end
        object CheckBox10: TCheckBox
          Left = 148
          Top = 90
          Width = 107
          Height = 17
          Caption = 'Network Interface'
          TabOrder = 9
        end
        object CheckBox11: TCheckBox
          Left = 268
          Top = 17
          Width = 45
          Height = 17
          Caption = 'RDP'
          TabOrder = 10
        end
        object CheckBox12: TCheckBox
          Left = 268
          Top = 35
          Width = 53
          Height = 17
          Caption = 'uRDM'
          TabOrder = 11
        end
        object CheckBox13: TCheckBox
          Left = 268
          Top = 54
          Width = 53
          Height = 17
          Caption = 'HotFix'
          TabOrder = 12
        end
        object CheckBox37: TCheckBox
          Left = 268
          Top = 72
          Width = 117
          Height = 17
          Caption = #1057#1077#1090#1077#1074#1099#1077' '#1088#1077#1089#1091#1088#1089#1099
          TabOrder = 13
        end
        object CheckBox38: TCheckBox
          Left = 268
          Top = 90
          Width = 117
          Height = 17
          Caption = #1057#1086#1073#1099#1090#1080#1103' Windows'
          TabOrder = 14
        end
      end
      object GroupBox3: TGroupBox
        Left = 0
        Top = 0
        Width = 405
        Height = 86
        Align = alTop
        Caption = #1040#1076#1084#1080#1085#1080#1089#1090#1088#1080#1088#1086#1074#1072#1085#1080#1077
        TabOrder = 2
        object LabeledEdit5: TLabeledEdit
          Left = 16
          Top = 32
          Width = 193
          Height = 21
          EditLabel.Width = 187
          EditLabel.Height = 13
          EditLabel.Caption = #1048#1084#1103' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' , '#1072#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088#1072
          TabOrder = 0
        end
        object LabeledEdit6: TLabeledEdit
          Left = 223
          Top = 32
          Width = 145
          Height = 21
          EditLabel.Width = 37
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1072#1088#1086#1083#1100
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          PasswordChar = #7
          TabOrder = 1
        end
        object LabeledEdit10: TLabeledEdit
          Left = 223
          Top = 59
          Width = 145
          Height = 21
          Hint = 
            #1055#1088#1080' '#1085#1077#1076#1086#1089#1090#1091#1087#1085#1086#1089#1090#1080' '#1086#1089#1085#1086#1074#1085#1086#1075#1086' DC '#1091#1082#1072#1078#1080#1090#1077' '#1083#1102#1073#1086#1081' '#1076#1086#1089#1090#1091#1087#1085#1099#1081' '#1082#1086#1085#1090#1088#1086#1083#1083#1077 +
            #1088' '#1076#1086#1084#1077#1085#1072
          EditLabel.Width = 205
          EditLabel.Height = 13
          EditLabel.BiDiMode = bdLeftToRight
          EditLabel.Caption = #1054#1087#1088#1072#1096#1080#1074#1072#1090#1100' '#1091#1082#1072#1079#1072#1085#1085#1099#1081' '#1082#1086#1085#1090#1088'.'#1076#1086#1084#1077#1085#1072'   '
          EditLabel.ParentBiDiMode = False
          LabelPosition = lpLeft
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
        end
      end
      object CheckBox30: TCheckBox
        Left = 3
        Top = 359
        Width = 198
        Height = 17
        Caption = #1057#1074#1086#1088#1072#1095#1080#1074#1072#1090#1100' '#1074' '#1090#1088#1077#1081
        TabOrder = 3
        OnClick = CheckBox30Click
      end
      object ComboBox1: TComboBox
        Left = 3
        Top = 334
        Width = 244
        Height = 21
        TabOrder = 4
        Text = 'Windows'
        OnChange = ComboBox1Change
      end
      object CheckBox33: TCheckBox
        Left = 3
        Top = 376
        Width = 251
        Height = 17
        Caption = #1047#1072#1087#1091#1089#1082' '#1089#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1103' '#1087#1088#1080' '#1089#1090#1072#1088#1090#1077' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
        TabOrder = 5
      end
      object CheckBox36: TCheckBox
        Left = 3
        Top = 393
        Width = 382
        Height = 17
        Caption = #1047#1072#1087#1080#1089#1100' '#1086#1096#1080#1073#1086#1082' DCOM '#1074' '#1089#1080#1089#1090#1077#1084#1085#1099#1081' '#1083#1086#1075' ('#1088#1077#1082#1086#1084#1077#1085#1076#1091#1077#1090#1089#1103')'
        TabOrder = 6
        OnClick = CheckBox36Click
      end
      object LinkLabel1: TLinkLabel
        Left = 304
        Top = 366
        Width = 18
        Height = 20
        Caption = '<a> ? </a>'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        OnClick = LinkLabel1Click
      end
      object CheckBox42: TCheckBox
        Left = 3
        Top = 410
        Width = 399
        Height = 17
        Hint = #1055#1088#1080' '#1089#1085#1103#1090#1080#1080' '#1095#1077#1082#1073#1086#1082#1089#1072' '#1073#1091#1076#1091#1090' '#1074#1099#1075#1088#1091#1078#1072#1090#1100#1089#1103' '#1074#1089#1077' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1080' '#1076#1086#1084#1077#1085#1072
        Caption = #1056#1072#1079#1076#1077#1083' "'#1050#1086#1084#1087#1100#1102#1090#1077#1088#1099'". '#1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1080'  '#1074#1099#1073#1088#1072#1085#1085#1086#1081' '#1075#1088#1091#1087#1087#1099' '#1073#1077#1079#1086#1087#1072#1089#1085#1086#1089#1090#1080
        ParentShowHint = False
        ShowHint = True
        TabOrder = 8
      end
      object CheckBox43: TCheckBox
        Left = 3
        Top = 427
        Width = 382
        Height = 17
        Hint = #1055#1088#1080' '#1089#1085#1103#1090#1080#1080' '#1095#1077#1082#1073#1086#1082#1089#1072' '#1073#1091#1076#1091#1090' '#1074#1099#1075#1088#1091#1078#1072#1090#1100#1089#1103' '#1074#1089#1077' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1099' '#1076#1086#1084#1077#1085#1072
        Caption = #1056#1072#1079#1076#1077#1083' "'#1059#1087#1088#1072#1074#1083#1077#1085#1080#1077'". '#1050#1086#1084#1087#1100#1102#1090#1077#1088#1099' '#1074#1099#1073#1088#1072#1085#1085#1086#1081' '#1075#1088#1091#1087#1087#1099' '#1073#1077#1079#1086#1087#1072#1089#1085#1086#1089#1090#1080
        ParentShowHint = False
        ShowHint = True
        TabOrder = 9
      end
      object CheckBox44: TCheckBox
        Left = 3
        Top = 444
        Width = 398
        Height = 17
        Hint = 
          #1055#1088#1080' '#1089#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1080' '#1089#1077#1090#1080' '#1074#1085#1086#1089#1080#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1085#1072' '#1074#1082#1083#1072#1076#1082#1077' "'#1050#1086#1084#1087#1100#1102#1090#1077#1088#1099'" '#1085#1077' ' +
          #1079#1072#1074#1080#1089#1080#1084#1086' '#1086#1090' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1089#1086#1089#1090#1086#1103#1085#1080#1103' '#1089#1077#1072#1085#1089#1072' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103'.'
        Caption = #1057#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1077': '#1086#1090#1088#1080#1089#1086#1074#1082#1072' '#1076#1072#1085#1085#1099#1093' '#1087#1088#1080' '#1080#1079#1084#1077#1085#1077#1085#1080#1080' '#1089#1077#1072#1085#1089#1072' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103
        ParentShowHint = False
        ShowHint = True
        TabOrder = 10
      end
      object CheckBox45: TCheckBox
        Left = 3
        Top = 461
        Width = 399
        Height = 17
        Hint = 
          #1055#1088#1086#1080#1079#1074#1086#1076#1080#1090#1100' '#1086#1089#1090#1072#1085#1086#1074#1082#1091' '#1089#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1103' '#1089#1077#1090#1080' '#1087#1088#1080' '#1080#1079#1084#1077#1085#1077#1085#1080#1080' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1089 +
          #1077#1072#1085#1089#1072' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103'.'
        Caption = #1054#1089#1090#1072#1085#1086#1074#1082#1072' '#1089#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1103' '#1087#1088#1080' '#1080#1079#1084#1077#1085#1077#1085#1080#1080' '#1089#1077#1072#1085#1089#1072' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103' '
        ParentShowHint = False
        ShowHint = True
        TabOrder = 11
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1048#1085#1074#1077#1085#1090#1072#1088#1080#1079#1072#1094#1080#1103
      ImageIndex = 1
      ExplicitHeight = 455
      object GroupBox5: TGroupBox
        Left = 0
        Top = 0
        Width = 405
        Height = 479
        Align = alClient
        TabOrder = 0
        ExplicitHeight = 455
        object GroupBox6: TGroupBox
          Left = 2
          Top = 212
          Width = 401
          Height = 136
          Align = alTop
          Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1085#1072#1088#1091#1096#1077#1085#1080#1103' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080' ('#1091#1095#1080#1090#1099#1074#1072#1090#1100')'
          TabOrder = 0
          object CheckBox14: TCheckBox
            Left = 14
            Top = 16
            Width = 97
            Height = 15
            Caption = #1055#1088#1086#1094#1077#1089#1089#1086#1088
            Checked = True
            Enabled = False
            State = cbChecked
            TabOrder = 0
          end
          object CheckBox15: TCheckBox
            Left = 14
            Top = 30
            Width = 131
            Height = 17
            Caption = #1052#1072#1090#1077#1088#1080#1085#1089#1082#1072#1103' '#1087#1083#1072#1090#1072
            TabOrder = 1
          end
          object CheckBox16: TCheckBox
            Left = 14
            Top = 45
            Width = 123
            Height = 17
            Caption = #1054#1087#1077#1088#1072#1090#1080#1074#1085#1103' '#1087#1072#1084#1103#1090#1100
            TabOrder = 2
          end
          object CheckBox17: TCheckBox
            Left = 14
            Top = 60
            Width = 97
            Height = 17
            Caption = #1046#1077#1089#1090#1082#1080#1081' '#1076#1080#1089#1082
            TabOrder = 3
          end
          object CheckBox18: TCheckBox
            Left = 14
            Top = 75
            Width = 97
            Height = 17
            Caption = #1042#1080#1076#1077#1086#1082#1072#1088#1090#1072
            TabOrder = 4
          end
          object CheckBox19: TCheckBox
            Left = 177
            Top = 45
            Width = 144
            Height = 17
            Caption = #1054#1087#1077#1088#1072#1094#1080#1086#1085#1085#1072#1103' '#1089#1080#1089#1090#1077#1084#1072
            TabOrder = 5
          end
          object CheckBox20: TCheckBox
            Left = 177
            Top = 16
            Width = 136
            Height = 15
            Caption = #1057#1077#1090#1077#1074#1099#1077' '#1080#1085#1090#1077#1088#1092#1077#1081#1089#1099
            TabOrder = 6
          end
          object CheckBox21: TCheckBox
            Left = 177
            Top = 30
            Width = 97
            Height = 17
            Caption = #1052#1086#1085#1080#1090#1086#1088#1099
            TabOrder = 7
          end
          object CheckBox22: TCheckBox
            Left = 177
            Top = 60
            Width = 136
            Height = 17
            Caption = #1048#1085#1074#1077#1085#1090#1072#1088#1085#1099#1081' '#1085#1086#1084#1077#1088
            TabOrder = 8
          end
          object CheckBox23: TCheckBox
            Left = 177
            Top = 77
            Width = 88
            Height = 15
            Caption = 'MAC '#1072#1076#1088#1077#1089
            TabOrder = 9
          end
          object CheckBox24: TCheckBox
            Left = 14
            Top = 90
            Width = 123
            Height = 17
            Caption = #1048#1084#1103' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1072
            TabOrder = 10
          end
          object CheckBox31: TCheckBox
            Left = 14
            Top = 106
            Width = 147
            Height = 17
            Caption = 'S.M.A.R.T ('#1086#1094#1077#1085#1082#1072' '#1054#1057')'
            TabOrder = 11
          end
        end
        object GroupBox7: TGroupBox
          Left = 2
          Top = 59
          Width = 401
          Height = 76
          Align = alTop
          Caption = #1048#1085#1074#1077#1085#1090#1072#1088#1080#1079#1072#1094#1080#1103' '#1087#1088#1086#1075#1088#1072#1084#1084' ('#1091#1095#1080#1090#1099#1074#1072#1090#1100')'
          TabOrder = 1
          object CheckBox27: TCheckBox
            Left = 14
            Top = 14
            Width = 211
            Height = 17
            Caption = #1042#1077#1088#1089#1080#1102' '#1087#1088#1086#1075#1088#1072#1084#1084#1085#1086#1075#1086' '#1086#1073#1077#1089#1087#1077#1095#1077#1085#1080#1103
            TabOrder = 0
          end
          object CheckBox28: TCheckBox
            Left = 14
            Top = 30
            Width = 227
            Height = 17
            Caption = #1055#1088#1086#1080#1079#1074#1086#1076#1080#1090#1077#1083#1103' ('#1080#1079#1076#1072#1090#1077#1083#1103') '#1087#1088#1086#1075#1088#1072#1084#1084
            TabOrder = 1
          end
          object CheckBox29: TCheckBox
            Left = 14
            Top = 46
            Width = 97
            Height = 17
            Caption = #1048#1084#1103' '#1087#1088#1086#1075#1088#1072#1084#1084
            Checked = True
            Enabled = False
            State = cbChecked
            TabOrder = 2
          end
        end
        object GroupBox8: TGroupBox
          Left = 2
          Top = 15
          Width = 401
          Height = 44
          Align = alTop
          Caption = 'SMART HDD'
          TabOrder = 2
          object CheckBox32: TCheckBox
            Left = 14
            Top = 16
            Width = 363
            Height = 17
            Caption = #1055#1088#1086#1080#1079#1074#1086#1076#1080#1090#1100' '#1086#1094#1077#1085#1082#1091' '#1089#1086#1089#1090#1086#1103#1085#1080#1103' HDD '#1087#1086' '#1072#1090#1088#1080#1073#1091#1090#1072#1084'  S.M.A.R.T'
            TabOrder = 0
          end
        end
        object GroupBox11: TGroupBox
          Left = 2
          Top = 135
          Width = 401
          Height = 77
          Align = alTop
          Caption = #1057#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1077' '#1089#1077#1090#1080
          TabOrder = 3
          object Label3: TLabel
            Left = 157
            Top = 15
            Width = 176
            Height = 13
            Caption = #1042#1099#1073#1086#1088' '#1092#1091#1085#1082#1094#1080#1080' ICMP Echo Request'
          end
          object LabeledEdit11: TLabeledEdit
            Left = 14
            Top = 30
            Width = 118
            Height = 21
            EditLabel.Width = 118
            EditLabel.Height = 13
            EditLabel.Caption = #1058#1072#1081#1084#1072#1091#1090' '#1076#1083#1103' '#1087#1080#1085#1075#1072' ('#1084#1089')'
            NumbersOnly = True
            ParentShowHint = False
            ShowHint = False
            TabOrder = 0
          end
          object CheckBox34: TCheckBox
            Left = 157
            Top = 55
            Width = 235
            Height = 17
            Caption = #1055#1088#1086#1074#1077#1088#1103#1090#1100' '#1076#1086#1089#1090#1091#1087#1085#1086#1089#1090#1100' uRDM '#1085#1072' '#1082#1083#1080#1077#1085#1090#1077
            TabOrder = 1
          end
          object CheckBox35: TCheckBox
            Left = 14
            Top = 55
            Width = 137
            Height = 17
            Caption = #1044#1086#1089#1090#1091#1087#1085#1086#1089#1090#1100' 135 '#1087#1086#1088#1090#1072
            TabOrder = 2
          end
          object ComboBox2: TComboBox
            Left = 157
            Top = 30
            Width = 229
            Height = 22
            Style = csOwnerDrawFixed
            ItemIndex = 0
            TabOrder = 3
            Text = 'IdIcmp'
            Items.Strings = (
              'IdIcmp'
              'GetAddrInfo+IcmpSendEcho'
              'GetHostByName+IcmpSendEcho')
          end
        end
        object GroupBox10: TGroupBox
          Left = 2
          Top = 348
          Width = 401
          Height = 128
          Align = alTop
          Caption = #1055#1088#1086#1076#1091#1082#1090#1099' Windows '#1080' Office'
          TabOrder = 4
          object ButtonCodeError: TButton
            Left = 271
            Top = 36
            Width = 122
            Height = 25
            Caption = #1050#1086#1076#1099' '#1080' '#1089#1090#1072#1090#1091#1089#1099
            TabOrder = 0
            OnClick = ButtonCodeErrorClick
          end
          object CheckBox39: TCheckBox
            Left = 14
            Top = 17
            Width = 282
            Height = 17
            Caption = #1057#1090#1072#1074#1080#1090#1100' '#1087#1088#1080#1074#1080#1074#1082#1091' '#1076#1083#1103' '#1087#1088#1086#1076#1091#1082#1090#1086#1074' Windows '#1080' Office'
            TabOrder = 1
            OnClick = CheckBox39Click
            OnMouseUp = CheckBox39MouseUp
          end
          object ComboKMS: TComboBox
            Left = 14
            Top = 39
            Width = 244
            Height = 22
            Style = csOwnerDrawFixed
            ItemIndex = 1
            TabOrder = 2
            Text = #1055#1088#1080#1074#1080#1074#1082#1072' '#1085#1072' 180 '#1076#1085#1077#1081
            Items.Strings = (
              #1055#1088#1080#1074#1080#1074#1082#1072' '#1089' '#1072#1074#1090#1086#1087#1088#1086#1076#1083#1077#1085#1080#1077#1084
              #1055#1088#1080#1074#1080#1074#1082#1072' '#1085#1072' 180 '#1076#1085#1077#1081
              #1059#1076#1072#1083#1080#1090#1100' '#1087#1088#1080#1074#1080#1074#1082#1091)
          end
          object LoadManual: TButton
            Left = 271
            Top = 65
            Width = 122
            Height = 25
            Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1087#1088#1080#1074#1080#1074#1082#1091
            TabOrder = 3
            OnClick = LoadManualClick
          end
          object CheckBox40: TCheckBox
            Left = 14
            Top = 73
            Width = 244
            Height = 17
            Caption = #1059#1076#1072#1083#1103#1090#1100' '#1082#1083#1102#1095#1080' KMS '#1087#1088#1080' '#1091#1076#1072#1083#1077#1085#1080#1080' '#1087#1088#1080#1074#1080#1074#1082#1080
            Enabled = False
            TabOrder = 4
          end
          object CheckBox41: TCheckBox
            Left = 14
            Top = 93
            Width = 363
            Height = 17
            Caption = #1057#1090#1072#1074#1080#1090#1100' '#1087#1088#1080#1074#1080#1082#1091' '#1087#1088#1080' '#1086#1090#1089#1091#1090#1089#1090#1074#1080#1080' '#1082#1083#1102#1095#1077#1081' '#1080' '#1089#1090#1072#1090#1091#1089#1077' "'#1053#1077#1080#1079#1074#1077#1089#1090#1085#1086'" '
            Enabled = False
            TabOrder = 5
          end
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093
      ImageIndex = 2
      ExplicitHeight = 455
      object GroupBox9: TGroupBox
        Left = 0
        Top = 0
        Width = 405
        Height = 479
        Align = alClient
        Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080' '#1087#1086#1076#1082#1083#1102#1095#1077#1085#1080#1103' '#1082' '#1073#1072#1079#1077' '#1076#1072#1085#1085#1099#1093
        TabOrder = 0
        ExplicitHeight = 455
        object Label2: TLabel
          Left = 16
          Top = 15
          Width = 49
          Height = 13
          Caption = #1055#1088#1086#1090#1086#1082#1086#1083
        end
        object SpeedButton1: TSpeedButton
          Left = 354
          Top = 112
          Width = 23
          Height = 23
          Flat = True
          Glyph.Data = {
            C6070000424DC607000000000000360000002800000016000000160000000100
            20000000000090070000C40E0000C40E00000000000000000000FCFCFCFFE1E1
            E1FF909090FF555555FF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D
            4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D4DFF4D4D
            4DFF575757FF949494FFE6E6E6FFFCFCFCFFE1E1E1FF676767FF4D4D4DFF4D4D
            4DFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E
            4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4E4E4EFF4D4D4DFF4D4D
            4DFF6C6C6CFFE6E6E6FF9C9C9CFF4D4D4DFF5A5A5AFF999999FFABABABFFABAB
            ABFFACACACFFABABABFFABABABFFABABABFFABABABFFABABABFFABABABFFABAB
            ABFFABABABFFACACACFFABABABFFABABABFF959595FF585858FF4D4D4DFF9A9A
            9AFF6E6E6EFF4D4D4DFFA4A4A4FFFBFBFBFFFCFCFCFFFDFDFDFFFDFDFDFFFDFD
            FDFFFDFDFDFFFDFDFDFFFDFDFDFFFCFCFCFFFDFDFDFFFCFCFCFFFCFCFCFFFDFD
            FDFFFDFDFDFFFDFDFDFFFBFBFBFF919191FF4D4D4DFF616161FF666666FF4E4E
            4EFFB6B6B6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFA4A4A4FF4E4E4EFF565656FF666666FF4E4E4EFFB6B6B6FFFFFF
            FFFFF7F7F7FF949494FF676767FF676767FF676767FF676767FF676767FF6767
            67FF676767FF676767FF676767FF676767FF686868FFA9A9A9FFFDFDFDFFA4A5
            A5FF4E4E4EFF575757FF666666FF4E4E4EFFB6B6B6FFFFFFFFFFCDCDCDFF5454
            54FF626262FF636363FF636363FF636363FF636363FF636363FF636363FF6363
            63FF636363FF626262FF4F4F4FFF676767FFF4F4F4FFA4A4A4FF4E4E4EFF5656
            56FF666666FF4E4E4EFFB6B6B6FFFFFFFFFFB8B8B8FF575757FF7D7D7DFF8080
            80FF808080FF808080FF808080FF808080FF808080FF808080FF808080FF7979
            79FF4E4E4EFF5F5F5FFFF4F4F4FFA4A5A5FF4E4E4EFF575757FF666666FF4E4E
            4EFFB6B6B6FFFAFAFAFF9F9F9FFF616161FF7E7E7EFF808080FF808080FF8080
            80FF808080FF808080FF808080FF808080FF808080FF707070FF464646FF5B5B
            5BFFF3F3F3FFA4A4A4FF4E4E4EFF575757FF666666FF4E4E4EFFB6B6B6FFE7E7
            E7FF5E5E5EFF6D6D6DFF7F7F7FFF808080FF808080FF808080FF808080FF8080
            80FF808080FF808080FF808080FF636363FF444444FF575757FFF3F3F3FFA4A4
            A4FF4E4E4EFF565656FF666666FF4E4E4EFFB6B6B6FFDBDBDBFF515151FF7676
            76FF808080FF808080FF808080FF808080FF808080FF808080FF808080FF8080
            80FF7B7B7BFF585858FF525252FF545454FFF2F2F2FFA4A4A4FF4E4E4EFF5757
            57FF666666FF4E4E4EFFB6B6B6FFB7B7B7FF585858FF7A7A7AFF808080FF8080
            80FF808080FF808080FF808080FF808080FF808080FF808080FF737373FF4F4F
            4FFF656565FF525252FFF1F1F1FFA4A4A4FF4E4E4EFF565656FF666666FF4E4E
            4EFFB6B6B6FF7D7D7DFF616161FF7D7D7DFF808080FF808080FF7F7F7FFF7B7B
            7BFF7A7A7AFF7A7A7AFF7A7A7AFF797979FF666666FF4F4F4FFF6E6E6EFF4E4E
            4EFFF1F1F1FFA4A4A4FF4E4E4EFF575757FF666666FF4E4E4EFFB6B6B6FF5959
            59FF666666FF7F7F7FFF808080FF808080FF767676FF5A5A5AFF545454FF5454
            54FF545454FF535353FF4B4B4BFF5B5B5BFF737373FF4C4C4CFFF0F0F0FFA4A4
            A4FF4E4E4EFF565656FF666666FF4E4E4EFFB6B6B6FF8F8F8FFF575757FF6868
            68FF6A6A6AFF696969FF5A5A5AFF494949FF5C5C5CFF5C5C5CFF5C5C5CFF5C5C
            5CFF5D5D5DFF6B6B6BFF6F6F6FFF484848FFEFEFEFFFA5A5A5FF4E4E4EFF5757
            57FF666666FF4E4E4EFFB6B6B6FFE4E4E4FF7D7D7DFF5A5A5AFF4D4D4DFF4C4C
            4CFF4B4B4BFF4B4B4BFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F
            4FFF4F4F4FFF4F4F4FFFF0F0F0FFA4A4A4FF4E4E4EFF565656FF666666FF4E4E
            4EFFB6B6B6FFFFFFFFFFFFFFFFFFE0E0E0FFCCCCCCFFCCCCCCFFCCCCCCFFCCCC
            CCFFCCCCCCFFCCCCCCFFCCCCCCFFCCCCCCFFCCCCCCFFCCCCCCFFCCCCCCFFD0D0
            D0FFFBFBFBFFA4A5A5FF4E4E4EFF575757FF666666FF4E4E4EFFB6B6B6FFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA4A4
            A4FF4E4E4EFF575757FF6D6D6DFF4D4D4DFFA9A9A9FFFCFCFCFFFDFDFDFFFDFD
            FDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFD
            FDFFFDFDFDFFFDFDFDFFFDFDFDFFFDFDFDFFFCFCFCFF949494FF4D4D4DFF5F5F
            5FFF989898FF4D4D4DFF5A5A5AFFAAAAAAFFBABABAFFBABBBBFFBABABAFFB9BA
            BAFFBABABAFFB9B9B9FFBABABAFFB9BABAFFBABABAFFBABBBBFFBABABAFFBBBB
            BBFFBABABAFFBABABAFFA6A6A6FF5A5A5AFF4D4D4DFF979797FFDEDEDEFF6060
            60FF4D4D4DFF4E4E4EFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F
            4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F4FFF4F4F
            4FFF4E4E4EFF4D4D4DFF676767FFE3E3E3FFFCFCFCFFDDDDDDFF949494FF6565
            65FF5D5D5DFF5E5E5EFF5D5D5DFF5E5E5EFF5D5D5DFF5D5D5DFF5D5D5DFF5D5D
            5DFF5E5E5EFF5D5D5DFF5E5E5EFF5E5E5EFF5D5D5DFF5E5E5EFF666666FF9696
            96FFDFDFDFFFFCFCFCFF}
          Layout = blGlyphBottom
          OnClick = SpeedButton1Click
        end
        object SpeedButton4: TSpeedButton
          Left = 354
          Top = 154
          Width = 28
          Height = 17
          Flat = True
          Glyph.Data = {
            CA020000424DCA020000000000003600000028000000140000000B0000000100
            18000000000094020000C40E0000C40E00000000000000000000F5F5F5F9F9F9
            F5F5F5F6F6F6F9F9F9F3F3F3C4C4C48989895F5F5F4C4C4C4B4B4B5E5E5E8888
            88C1C1C1F3F3F3F9F9F9F6F6F6F5F5F5F9F9F9F5F5F5F7F7F7F6F6F6F7F7F7F6
            F6F6BABABA5F5F5F404040404040404040404040404040404040404040404040
            5C5C5CB6B6B6F6F6F6F7F7F7F6F6F6F7F7F7F8F8F8F4F4F4E9E9E97777774040
            404D4D4D7C7C7C4646464040404040404040404040404545457C7C7C4E4E4E40
            4040737373E7E7E7F4F4F4F8F8F8F5F5F5D7D7D7555555404040888888E0E0E0
            575757404040404040404040404040404040404040535353DFDFDF8D8D8D4141
            41525252D3D3D3F5F5F5CECECE4A4A4A484848BABABAF5F5F5B5B5B540404040
            4040404040404040404040404040404040404040AEAEAEF5F5F5BFBFBF4A4A4A
            484848C9C9C95D5D5D404040B3B3B3F7F7F7F5F5F59A9A9A3E3E461F1FA41D1D
            AB3D3D4B404040404040404040404040939393F5F5F5F7F7F7BABABA40404055
            5555C1C1C14646464C4C4CC4C4C4F9F9F9ADADAD2424940000FF0000FF1E1EA6
            404040404040404040404040A6A6A6F9F9F9C8C8C84F4F4F444444BCBCBCF8F8
            F8C9C9C94D4D4D424242929292E8E8E839399A0000FF0000FF23239A40404040
            40404040404D4D4DE5E5E59797974343434B4B4BC5C5C5F8F8F8F7F7F7F6F6F6
            E0E0E06A6A6A40404055555584848530308C28288B3F3F434040404040404646
            46838383585858404040666666DDDDDDF6F6F6F7F7F7F5F5F5F9F9F9F5F5F5F4
            F4F4B0B0B0525252404040404040404040404040404040404040404040404040
            505050ACACACF3F3F3F5F5F5F9F9F9F5F5F5F8F8F8F4F4F4F8F8F8F7F7F7F4F4
            F4F4F4F4B5B5B57B7B7B515151414141404040505050797979B2B2B2F2F2F2F4
            F4F4F7F7F7F8F8F8F4F4F4F8F8F8}
          OnMouseDown = SpeedButton4MouseDown
          OnMouseUp = SpeedButton4MouseUp
        end
        object ComboProtocol: TComboBox
          Left = 16
          Top = 32
          Width = 361
          Height = 21
          TabOrder = 0
          Text = 'local'
          OnSelect = ComboProtocolSelect
          Items.Strings = (
            'local'
            'TCPIP')
        end
        object EditDBserver: TLabeledEdit
          Left = 16
          Top = 72
          Width = 202
          Height = 21
          EditLabel.Width = 37
          EditLabel.Height = 13
          EditLabel.Caption = #1057#1077#1088#1074#1077#1088
          TabOrder = 1
          Text = 'localhost'
        end
        object CheckBox25: TCheckBox
          Left = 16
          Top = 174
          Width = 243
          Height = 17
          Caption = #1055#1086#1076#1082#1083#1102#1095#1072#1090#1100' '#1073#1072#1079#1091' '#1087#1088#1080' '#1079#1072#1087#1091#1089#1082#1077' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
          TabOrder = 2
        end
        object CheckBox26: TCheckBox
          Left = 351
          Top = 201
          Width = 16
          Height = 17
          TabOrder = 3
          OnClick = CheckBox26Click
        end
        object LabeledEdit7: TLabeledEdit
          Left = 16
          Top = 113
          Width = 332
          Height = 21
          EditLabel.Width = 299
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1091#1090#1100' '#1082' '#1073#1072#1079#1077' '#1076#1072#1085#1085#1099#1093' ('#1085#1077' '#1079#1072#1074#1080#1089#1080#1084#1086' '#1086#1090' '#1089#1077#1088#1074#1077#1088#1072' '#1080' '#1087#1088#1086#1090#1086#1082#1086#1083#1072')'
          TabOrder = 4
        end
        object LabeledEdit8: TLabeledEdit
          Left = 16
          Top = 150
          Width = 160
          Height = 21
          EditLabel.Width = 72
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
          TabOrder = 5
        end
        object LabeledEdit9: TLabeledEdit
          Left = 188
          Top = 150
          Width = 160
          Height = 21
          EditLabel.Width = 37
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1072#1088#1086#1083#1100
          PasswordChar = #7
          TabOrder = 6
        end
        object EditDBPort: TLabeledEdit
          Left = 224
          Top = 72
          Width = 70
          Height = 21
          EditLabel.Width = 25
          EditLabel.Height = 13
          EditLabel.Caption = #1055#1086#1088#1090
          TabOrder = 7
          Text = '3050'
        end
        object SpeedButton2: TButton
          Left = 150
          Top = 197
          Width = 88
          Height = 25
          Hint = #1057#1086#1079#1076#1072#1090#1100' '#1085#1086#1074#1091#1102' '#1073#1072#1079#1091' '#1076#1072#1085#1085#1099#1093
          Caption = #1057#1086#1079#1076#1072#1090#1100' '#1041#1044
          ImageIndex = 3
          Images = ImageListSettings
          ParentShowHint = False
          ShowHint = True
          TabOrder = 8
          OnClick = SpeedButton2Click
        end
        object SpeedButton3: TButton
          Left = 242
          Top = 197
          Width = 102
          Height = 25
          Hint = #1054#1095#1080#1089#1090#1080#1090#1100' '#1090#1077#1082#1091#1097#1091#1102' '#1073#1072#1079#1091' '#1076#1072#1085#1085#1099#1093
          Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1041#1044
          DropDownMenu = CleanDataBasePopup
          Enabled = False
          ImageIndex = 1
          Images = ImageListSettings
          ParentShowHint = False
          ShowHint = True
          Style = bsSplitButton
          TabOrder = 9
        end
        object Button3: TButton
          Left = 300
          Top = 70
          Width = 77
          Height = 25
          Caption = #1058#1077#1089#1090
          ImageIndex = 0
          Images = ImageListSettings
          TabOrder = 10
          OnClick = Button3Click
        end
        object Button4: TButton
          Left = 16
          Top = 197
          Width = 128
          Height = 25
          Caption = #1054#1073#1085#1086#1074#1080#1090#1100' '#1041#1044
          DropDownMenu = PopupUpdateDB
          ImageIndex = 2
          Images = ImageListSettings
          Style = bsSplitButton
          TabOrder = 11
        end
        object Button5: TButton
          Left = 16
          Top = 315
          Width = 137
          Height = 25
          Caption = #1055#1088#1086#1076#1091#1082#1090#1099' Microsoft'
          TabOrder = 12
          Visible = False
          OnClick = Button5Click
        end
        object Button6: TButton
          Left = 16
          Top = 346
          Width = 137
          Height = 25
          Caption = #1050#1086#1076#1099' '#1086#1096#1080#1073#1086#1082' LIC-ERROR'
          TabOrder = 13
          Visible = False
          OnClick = Button6Click
        end
      end
    end
  end
  object GroupBox4: TGroupBox
    Left = 0
    Top = 507
    Width = 413
    Height = 41
    Align = alBottom
    TabOrder = 1
    ExplicitTop = 483
    object Button1: TButton
      Left = 314
      Top = 6
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 227
      Top = 6
      Width = 75
      Height = 25
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 1
      OnClick = Button2Click
    end
  end
  object FDScript1: TFDScript
    SQLScripts = <
      item
        Name = 'CreateDB'
      end
      item
        Name = 'ClearIndex'
      end>
    Connection = FDConnection1
    ScriptOptions.BreakOnError = True
    ScriptOptions.FileEncoding = ecUTF8
    ScriptOptions.FileEndOfLine = elWindows
    ScriptOptions.DriverID = 'FB'
    ScriptOptions.SQLDialect = 3
    ScriptDialog = FDGUIxScriptDialog1
    Params = <>
    Macros = <>
    FetchOptions.AssignedValues = [evItems, evAutoClose, evAutoFetchAll]
    FetchOptions.AutoClose = False
    FetchOptions.Items = [fiBlobs, fiDetails]
    ResourceOptions.AssignedValues = [rvMacroCreate, rvMacroExpand, rvDirectExecute, rvPersistent]
    ResourceOptions.MacroCreate = False
    ResourceOptions.DirectExecute = True
    Left = 196
    Top = 480
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    Left = 148
    Top = 480
  end
  object FDGUIxScriptDialog1: TFDGUIxScriptDialog
    Provider = 'Forms'
    Caption = #1055#1088#1086#1094#1077#1089#1089' '#1089#1086#1079#1076#1072#1085#1080#1103' '#1085#1086#1074#1086#1081' '#1041#1044
    Options = [ssCallstack, ssConsole]
    Left = 276
    Top = 480
  end
  object FDConnection1: TFDConnection
    Params.Strings = (
      'Server=localhost'
      'CharacterSet=UTF8'
      'User_Name=sysdba'
      'Password=masterkey'
      
        'Database=N:\'#1055#1088#1086#1077#1082#1090#1099' '#1076#1083#1103' '#1088#1072#1073#1086#1090#1099'\'#1055#1088#1086#1077#1082#1090#1099' WMI\DataBASEForMRPC\DB - ' +
        '31.FDB'
      'OSAuthent=No'
      'OpenMode=OpenOrCreate'
      'DriverID=FB')
    LoginPrompt = False
    Left = 28
    Top = 480
  end
  object ImageListSettings: TImageList
    Left = 76
    Top = 480
    Bitmap = {
      494C0101040078007C0010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000002000000001002000000000000020
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000009F336600A25739000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFFECECECFF4646
      46FF030303FF000000FF000000FF000000FF000000FF000000FF000000FF0303
      03FF474747FFEDEDEDFFFFFFFFFF000000000000000000000000000000000000
      0000DEDEDE00979797006969690051525200515252006868680096969600DEDE
      DE000000000000000000000000000000000000000000F6F6F600F5F5F500D3D3
      D300AEAEAE00999999008F8F8F0094949400A5A5A500D9D9D900E2E2E2008484
      84007979790078787700BCBCBC0000000000A73E6D00BE865A00B87C5400A55C
      3B00000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF6E6E6EFF1A1A
      1AFF9D9D9DFFA5A5A5FFA5A5A5FFA5A5A5FFA5A5A5FFA5A5A5FFA5A5A5FF9D9D
      9DFF181818FF6F6F6FFFFFFFFFFF000000000000000000000000F8F9F9009495
      95004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D
      4D0094959500F8F8F8000000000000000000F5F5F500AAAAAA00757575009D9D
      9D00C5C5C500DCDCDC00E6E5E500E1E1E100CFCFCF00D3D3D30071717100DEDE
      DE00D3D3D300F2F2F2008F8F8F00BCBCBC00C0815900C5916600CDA17E00C495
      6E00B5795100AA684400A25C3B009B5033000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF484848FF5F5F
      5FFFFFFFFFFFC3C3C3FF9C9C9CFFFFFFFFFFFFFFFFFF9C9C9CFFC3C3C3FFFFFF
      FFFF5F5F5FFF484848FFFFFFFFFF0000000000000000F8F9F800797979004D4D
      4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D
      4D004D4D4D0079797900F8F8F8000000000091919000B1B1B100F6F6F600F6F6
      F600F6F6F60000000000F6F6F60000000000F6F6F600B1B1B100BABABA00F3F3
      F20081818100EEEEED00F2F2F1007878770000000000C2855F00D1A58200CC9E
      7900CB9D7900C7987200C2916A00BD8C6300A767430000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF0000000000000000949595004D4D4D004D4D
      4D004F4F4F00525252004D4D4D004D4D4D0078787800646565004D4D4D004D4D
      4D004D4D4D004D4D4D0095969600000000007D7D7C00F6F6F600EBEBEB00B9B9
      B90094949400808080007F7F7E007F7E7E00BBBBBB009F9F9F00D1D1D1008E8E
      8E005959590081818100D3D3D3007A7979000000000000000000D0A07A00D7AD
      8E00C9966D00C28E6400BC875A00BF8B6200BB875F00814C3D00000000000000
      00000000000000000000000000000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000DEDEDE004D4E4E004D4D4D004D4D
      4D00BCBCBC00CCCCCC004F4F4F004D4D4D0062626200EAEAEA00C8C8C8005757
      57004D4D4D004D4D4D004D4D4D00DEDEDE007A797900929291007F7F7F00B7B7
      B700DFDFDF00F4F4F40000000000F6F6F600F4F4F400C9C9C900999999000000
      00008E8E8E00F3F3F200DEDEDE00848484000000000000000000D7A58100DCB5
      9800D0A07B00CB997100CFA38100C7987200886A560086868600000000000000
      00000000000000000000000000000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000979797004D4D4D004D4D4D00A0A0
      A0000000000000000000B2B2B2004D4D4D004D4D4D008F8F8F0000000000C1C1
      C1004D4D4D004D4D4D004D4D4D009696960049494900C9CAC900000000000000
      0000F6F6F60000000000F6F6F6000000000000000000F6F6F6008E8E8E009999
      9900D1D1D100BABABA0071717100E2E2E2000000000000000000DDAD8B00E2BD
      A300D8AA8800D9B29300CF9E7800A172580092929200A8A8A800000000000000
      00000000000000000000000000000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000686969004D4D4D007F7F7F00FBFB
      FB000000000000000000FDFDFD00909090004D4D4D006767670000000000F9F9
      F9004D4E4E004D4D4D004D4D4D00686868007D7C7C0000000000F6F6F600D9D9
      D800B4B4B4009E9E9D00969696009A9A9A00ABABAB00CBCBCA00F1F1F100C9C9
      C9009F9F9F00B1B1B100EEEEEE00F6F6F6000000000000000000E3B39200E8C5
      AC00E3BFA500DBAF8E00B1867000000000000000000000000000000000000000
      00008D3C300000000000000000000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000515252004D4D4D004D4D4D004F4F
      4F000000000000000000636363004D4D4D004D4D4D0063636300000000000000
      00004F4F4F004D4D4D004D4D4D005152520080808000B1B1B000757575009797
      9700BFBFBF00D7D7D700E0DFDF00DBDBDB00C9C8C800A7A7A6007B7B7B009191
      9000BBBBBB00C2C2C200F6F6F60000000000000000000000000000000000E8BF
      A200E5BEA200B2978900ADADAD00000000000000000000000000000000009C4F
      3800AC704A00A15D3D008C2E41000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF747474FF323232FFFFFFFFFFFFFFFFFF313131FF757575FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000515252004D4D4D004D4D4D004F4F
      4F000000000000000000636363004D4D4D004D4D4D0063636300000000000000
      00004F4F4F004D4D4D004D4D4D005152520048484800ABABAB00F5F5F5000000
      0000F6F6F60000000000F6F6F6000000000000000000F6F6F60000000000D2D2
      D1005C5C5C00B8B8B800F6F6F600000000000000000000000000000000000000
      0000C9A49800B6B6B600BBBBBB00000000000000000000000000AA634300BA82
      5A00BF8E6500BA895E00984E31000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFF828282FF404040FFFFFFFFFFFFFFFFFF404040FF838383FFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000696969004D4D4D004D4D4D004D4E
      4E00F9F9F90000000000676767004D4D4D008E8E8E00FBFBFB00000000000000
      0000F9F9F9007E7E7E004D4D4D00696969007777770000000000F6F6F600F2F2
      F100D4D4D400BEBEBE00B6B5B500BABAB900CBCBCA00EAEAEA00F6F6F6000000
      0000B4B4B400B8B8B80000000000F6F6F6000000000000000000000000000000
      00000000000000000000000000000000000000000000BE7D5800C9966D00CB9E
      7A00BB845700C2916900A5613C000000000000000000FFFFFFFF474747FF6060
      60FFFFFFFFFFFDFDFDFFF6F6F6FFFFFFFFFFFFFFFFFFF6F6F6FFFDFDFDFFFFFF
      FFFF606060FF474747FFFFFFFFFF00000000969696004D4D4D004D4D4D004D4D
      4D00C0C1C100000000008F8F8F004D4D4D004D4D4D00B2B2B200000000000000
      0000A0A0A0004D4D4D004D4D4D009797970081818100CECECE00818181007E7E
      7E009F9F9F00B7B7B700BFBFBF00BBBBBB00A9A9A9008787870079797900AFAF
      AE00BBBBBB00B8B8B800F6F6F600000000000000000000000000000000000000
      000000000000000000000000000000000000D9AA8700DAB19300D8AF9100CB98
      7000C38F6600C89B7600B17048000000000000000000FFFFFFFF505050FF6A6A
      6AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFF6A6A6AFF505050FFFFFFFFFF00000000DEDEDE004D4E4E004D4D4D004D4D
      4D0057575700C9C9C900EAEAEA00616262004D4D4D004F4F4F00CDCDCD00BCBC
      BC004D4D4D004D4D4D004D4D4D00DEDEDE00505050008D8D8D00E9E9E900F6F6
      F60000000000F6F6F60000000000F6F6F600F6F6F60000000000F4F4F400B4B4
      B4005C5C5C00B8B8B80000000000F6F6F6000000000000000000000000000000
      0000000000000000000000000000DCA48E00E7C0A700E0B99E00D8AB8A00D2A4
      8100CE9C7500D1A58300BD855A000000000000000000CECECEFFA9A9A9FFADAD
      ADFFB8B8B8FFB8B8B8FFB8B8B8FFB8B8B8FFB8B8B8FFB8B8B8FFB8B8B8FFB8B8
      B8FFADADADFFA8A8A8FFCECECEFF0000000000000000949595004D4D4D004D4D
      4D004D4D4D004D4D4D0064646400787878004D4D4D004D4D4D00525252004F4F
      4F004D4D4D004D4D4D0094959500000000006F6F6E00F6F6F600000000000000
      0000F6F6F60000000000F6F6F6000000000000000000F6F6F60000000000F6F6
      F600ACABAB00B8B8B800F6F6F600000000000000000000000000000000000000
      000000000000000000000000000000000000E8BFA300E9C8AF00E5C2A800E1BC
      A100DCB59800D5AA8900D0A38100B170450000000000171717FF000000FF0000
      00FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF0000
      00FF000000FF000000FF171717FF0000000000000000F8F9F900797979004D4D
      4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D
      4D004D4D4D0079797900F8F9F9000000000076757500EBEBEB00F6F6F6000000
      0000F6F6F60000000000F6F6F60000000000F6F6F600F6F6F60000000000F6F6
      F6008B8B8B00D6D6D600F6F6F600000000000000000000000000000000000000
      00000000000000000000000000000000000000000000E7BA9C00E4B59600E0B1
      9100DAAD8E00DCB49700CF9E7800AC43740000000000F5F5F5FFECECECFFECEC
      ECFFECECECFFECECECFFDADADAFF8E8E8EFF8E8E8EFFD9D9D9FFECECECFFECEC
      ECFFECECECFFECECECFFF5F5F5FF000000000000000000000000F8F9F9009495
      95004D4E4E004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4D4D004D4E
      4E0094959500F8F9F8000000000000000000D8D8D80074747400B0B0B000EBEB
      EB0000000000F6F6F60000000000F6F6F600F6F6F600F4F4F400C8C7C7007C7C
      7B00AFAFAF00F6F6F60000000000F6F6F6000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000D5A07F00B45084000000000000000000FFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFB0B0B0FF040404FF040404FFAFAFAFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFF000000000000000000000000000000000000
      0000DEDEDE00969696006969690051525200515252006869690097979700DEDE
      DE000000000000000000000000000000000000000000F3F3F300BABABA008585
      84007C7B7B007F7F7E00807F7F007F7F7E007D7D7D007D7C7C00A5A5A500E7E7
      E7000000000000000000F6F6F60000000000424D3E000000000000003E000000
      2800000040000000200000000100010000000000000100000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000009FFF8001F00F80010FFF8001C0030000
      00FF800180010500807F800180010000C03F800100000210C03F80010C203580
      C03F80010C204000C1F780010C300001E1E180010C3015A1F1C1800104304012
      FF81800104300001FF01800100000A42FE018001800135A1FF00800180011521
      FF808001C0030A02FFF98001F00F800D00000000000000000000000000000000
      000000000000}
  end
  object CleanDataBasePopup: TPopupMenu
    Left = 116
    Top = 480
    object N1: TMenuItem
      Caption = #1041#1072#1079#1091' '#1076#1072#1085#1085#1099#1093
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1047#1072#1073#1083#1086#1082#1080#1088#1086#1074#1072#1085#1085#1099#1077' '#1079#1072#1076#1072#1095#1080
      OnClick = N2Click
    end
    object N6: TMenuItem
      Caption = #1042#1089#1077' '#1079#1072#1076#1072#1095#1080
      OnClick = N6Click
    end
    object N3: TMenuItem
      Caption = #1048#1085#1074#1077#1085#1090#1072#1088#1080#1079#1072#1094#1080#1103' '#1055#1054
      OnClick = N3Click
    end
    object N5: TMenuItem
      Caption = #1048#1085#1074#1077#1085#1090#1072#1088#1080#1079#1072#1094#1080#1103'  '#1086#1073#1086#1088#1091#1076#1086#1074#1072#1085#1080#1103
      OnClick = N5Click
    end
    object N4: TMenuItem
      Caption = #1057#1082#1072#1085#1080#1088#1086#1074#1072#1085#1080#1077' '#1089#1077#1090#1080
      OnClick = N4Click
    end
    object WindowsOffice1: TMenuItem
      Caption = 'Windows '#1080' Office'
      OnClick = WindowsOffice1Click
    end
    object N7: TMenuItem
      Caption = #1040#1085#1090#1080#1074#1080#1088#1091#1089#1085#1099#1077' '#1087#1088#1086#1076#1091#1082#1090#1099
      OnClick = N7Click
    end
    object N8: TMenuItem
      Caption = #1057#1086#1093#1088#1072#1085#1077#1085#1085#1099#1077' '#1082#1083#1102#1095#1080' '#1088#1077#1077#1089#1090#1088#1072
      OnClick = N8Click
    end
  end
  object PopupUpdateDB: TPopupMenu
    Images = frmDomainInfo.ImageListMainMenu
    Left = 172
    Top = 480
    object N40411: TMenuItem
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100' 4.0 -> 4.1'
      ImageIndex = 32
      OnClick = Update4to41
    end
    object N30401: TMenuItem
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100' 3.0 -> 4.0'
      ImageIndex = 32
      OnClick = Button4Click
    end
    object N21301: TMenuItem
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100' 2.1 ->3.0'
      ImageIndex = 32
      OnClick = SpeedButton5Click
    end
  end
  object IdHTTP1: TIdHTTP
    AllowCookies = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = -1
    Request.ContentRangeEnd = -1
    Request.ContentRangeStart = -1
    Request.ContentRangeInstanceLength = -1
    Request.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    Request.Ranges.Units = 'bytes'
    Request.Ranges = <>
    HTTPOptions = [hoForceEncodeParams]
    Left = 356
    Top = 305
  end
end
