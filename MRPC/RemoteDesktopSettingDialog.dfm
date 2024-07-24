object RemoteDesktopSetting: TRemoteDesktopSetting
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080' '#1089#1077#1088#1074#1077#1088#1072
  ClientHeight = 330
  ClientWidth = 286
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 286
    Height = 294
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1054#1089#1085#1086#1074#1085#1099#1077
      object Label1: TLabel
        Left = 24
        Top = 233
        Width = 116
        Height = 26
        Caption = #1054#1090#1086#1073#1088#1072#1078#1072#1090#1100' '#1087#1088#1086#1094#1077#1089#1089', '#1074#1077#1089#1090#1080' '#1083#1086#1075' '#1089#1086#1077#1076#1080#1085#1077#1085#1080#1081'.'
        WordWrap = True
      end
      object Label3: TLabel
        Left = 162
        Top = 242
        Width = 25
        Height = 13
        Caption = 'Color'
        Visible = False
      end
      object Label4: TLabel
        Left = 163
        Top = 209
        Width = 18
        Height = 13
        Caption = 'FPS'
        Visible = False
      end
      object ListView1: TListView
        Left = 0
        Top = 0
        Width = 274
        Height = 185
        Columns = <
          item
            Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
            Width = 150
          end
          item
            Caption = #1055#1088#1072#1074#1072
            Width = 110
          end>
        GridLines = True
        ReadOnly = True
        RowSelect = True
        PopupMenu = PopupMenu1
        TabOrder = 0
        ViewStyle = vsReport
      end
      object LabeledEdit1: TLabeledEdit
        Left = 28
        Top = 206
        Width = 121
        Height = 21
        EditLabel.Width = 69
        EditLabel.Height = 13
        EditLabel.Caption = #1055#1086#1088#1090' '#1089#1077#1088#1074#1077#1088#1072
        Enabled = False
        TabOrder = 1
      end
      object CheckBox1: TCheckBox
        Left = 3
        Top = 208
        Width = 19
        Height = 17
        TabOrder = 2
        OnClick = CheckBox1Click
      end
      object CheckBox2: TCheckBox
        Left = 3
        Top = 238
        Width = 15
        Height = 21
        TabOrder = 3
      end
      object EditRDPFPS: TSpinEdit
        Left = 193
        Top = 206
        Width = 35
        Height = 22
        Hint = 'FPS ('#1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1082#1072#1076#1088#1086#1074'/'#1089#1077#1082')'
        MaxValue = 30
        MinValue = 5
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        Value = 20
        Visible = False
      end
      object ComboBox1: TComboBox
        Left = 193
        Top = 238
        Width = 35
        Height = 22
        Style = csOwnerDrawFixed
        ItemIndex = 2
        TabOrder = 5
        Text = '24'
        Visible = False
        Items.Strings = (
          '8'
          '16'
          '24'
          '25'
          '32')
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Firewall Windows'
      ImageIndex = 1
      object Label2: TLabel
        Left = 9
        Top = 159
        Width = 254
        Height = 78
        Caption = 
          #1056#1072#1079#1088#1077#1096#1072#1090#1100' '#1087#1086#1076#1082#1083#1102#1095#1077#1085#1080#1103' '#1090#1086#1083#1100#1082#1086' '#1089' '#1091#1082#1072#1079#1072#1085#1085#1099#1093' IP '#1072#1076#1088#1077#1089#1086#1074'. '#1060#1086#1088#1084#1072#1090' '#1074#1074#1086#1076 +
          #1072': 192.168.0.1                    '#1055#1086#1076#1089#1077#1090#1100' -  192.168.0.1/24     ' +
          '                              '#1044#1080#1072#1087#1072#1079#1086#1085' - 192.168.0.1-192.168.0.1' +
          '55       '#1047#1072#1087#1080#1089#1080' '#1088#1072#1079#1076#1077#1083#1103#1102#1090#1089#1103' '#1079#1072#1087#1103#1090#1099#1084#1080', '#1076#1083#1103' '#1088#1072#1079#1088#1077#1096#1077#1085#1080#1103' '#1074#1089#1077#1093' '#1072#1076#1088#1077#1089#1086 +
          #1074' '#1091#1082#1072#1079#1072#1090#1100' *'
        WordWrap = True
      end
      object CheckBox3: TCheckBox
        Left = 6
        Top = 6
        Width = 153
        Height = 17
        Caption = #1055#1088#1072#1074#1080#1083#1086' '#1076#1083#1103' uRDMServer'
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object GroupBox1: TGroupBox
        Left = 6
        Top = 55
        Width = 259
        Height = 100
        Caption = #1055#1088#1077#1076#1086#1087#1088#1077#1076#1077#1083#1077#1085#1085#1099#1077' '#1075#1088#1091#1087#1087#1099' '#1087#1088#1072#1074#1080#1083
        TabOrder = 1
        object CheckBox4: TCheckBox
          Left = 9
          Top = 17
          Width = 240
          Height = 17
          Caption = #1054#1073#1097#1080#1081' '#1076#1086#1089#1090#1091#1087' '#1082' '#1092#1072#1081#1083#1072#1084' '#1080' '#1087#1088#1080#1085#1090#1077#1088#1072#1084
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
        object CheckBox6: TCheckBox
          Left = 9
          Top = 54
          Width = 256
          Height = 17
          Caption = #1048#1085#1089#1090#1088#1091#1084#1077#1085#1090#1072#1088#1080#1081' '#1091#1087#1088#1072#1074#1083#1077#1085#1080#1103' Windows (WMI)'
          TabOrder = 1
        end
        object CheckBox7: TCheckBox
          Left = 9
          Top = 72
          Width = 240
          Height = 17
          Caption = #1050#1086#1086#1088#1076#1080#1085#1072#1090#1086#1088' '#1088#1072#1089#1087#1088#1077#1076#1077#1083#1077#1085#1085#1099#1093' '#1090#1088#1072#1085#1079#1072#1082#1094#1080#1081
          TabOrder = 2
        end
        object CheckBox5: TCheckBox
          Left = 9
          Top = 35
          Width = 262
          Height = 17
          Caption = #1044#1080#1089#1090#1072#1085#1094#1080#1086#1085#1085#1086#1077' '#1091#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1088#1072#1073#1086#1095#1080#1084' '#1089#1090#1086#1083#1086#1084
          TabOrder = 3
        end
      end
      object CheckBox8: TCheckBox
        Left = 6
        Top = 24
        Width = 259
        Height = 25
        Caption = #1055#1088#1072#1074#1080#1083#1072' '#1076#1083#1103' '#1089#1080#1089#1090#1077#1084#1085#1099#1093' '#1089#1083#1091#1078#1073' (RPC,WMI, RDP)'
        Checked = True
        State = cbChecked
        TabOrder = 2
        WordWrap = True
      end
      object Edit1: TEdit
        Left = 9
        Top = 243
        Width = 251
        Height = 21
        TabOrder = 3
        Text = '*'
      end
    end
  end
  object Memo1: TMemo
    Left = 328
    Top = 24
    Width = 339
    Height = 233
    Lines.Strings = (
      #1053#1072#1089#1090#1088#1086#1081#1082#1080' '#1086#1090#1089#1091#1090#1089#1090#1074#1091#1102#1090)
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 294
    Width = 286
    Height = 36
    Align = alBottom
    TabOrder = 2
    object OKBtn: TButton
      Left = 119
      Top = 6
      Width = 75
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      Default = True
      TabOrder = 0
      OnClick = OKBtnClick
    end
    object CancelBtn: TButton
      Left = 201
      Top = 6
      Width = 75
      Height = 25
      Cancel = True
      Caption = #1047#1072#1082#1088#1099#1090#1100
      ModalResult = 2
      TabOrder = 1
      OnClick = CancelBtnClick
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 232
    object N1: TMenuItem
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1048#1079#1084#1077#1085#1080#1090#1100
      OnClick = N2Click
    end
    object N3: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      OnClick = N3Click
    end
  end
  object PopupDeptColor: TPopupMenu
    Left = 249
    Top = 121
    object N161: TMenuItem
      Caption = '16'
    end
    object N241: TMenuItem
      Caption = '24'
    end
  end
end
