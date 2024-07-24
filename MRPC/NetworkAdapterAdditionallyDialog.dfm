object OKRightDlg12345678910111213141516: TOKRightDlg12345678910111213141516
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1044#1086#1087#1086#1083#1085#1080#1090#1077#1083#1100#1085#1099#1077' '#1087#1072#1088#1072#1084#1077#1090#1088#1099' TCP/IP'
  ClientHeight = 415
  ClientWidth = 384
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 220
    Top = 383
    Width = 75
    Height = 25
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 301
    Top = 383
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    ModalResult = 2
    TabOrder = 1
  end
  object PageControl1: TPageControl
    Left = 8
    Top = 8
    Width = 368
    Height = 369
    ActivePage = TabSheet1
    TabOrder = 2
    object TabSheet1: TTabSheet
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' IP'
      object GroupBox1: TGroupBox
        Left = 3
        Top = 4
        Width = 354
        Height = 163
        Caption = 'IP-'#1072#1076#1088#1077#1089#1072
        TabOrder = 0
        object ListView1: TListView
          Left = 3
          Top = 20
          Width = 341
          Height = 98
          Columns = <
            item
              Caption = 'IP-'#1072#1076#1088#1077#1089
              Width = 165
            end
            item
              Caption = #1052#1072#1089#1082#1072' '#1087#1086#1076#1089#1077#1090#1080
              Width = 165
            end>
          GridLines = True
          RowSelect = True
          TabOrder = 0
          ViewStyle = vsReport
        end
        object Button1: TButton
          Left = 92
          Top = 124
          Width = 75
          Height = 25
          Caption = #1044#1086#1073#1072#1074#1080#1090#1100
          TabOrder = 1
          OnClick = Button1Click
        end
        object Button2: TButton
          Left = 270
          Top = 124
          Width = 75
          Height = 25
          Caption = #1059#1076#1072#1083#1080#1090#1100
          TabOrder = 2
          OnClick = Button2Click
        end
        object Button7: TButton
          Left = 181
          Top = 124
          Width = 75
          Height = 25
          Caption = #1048#1079#1084#1077#1085#1080#1090#1100
          TabOrder = 3
          OnClick = Button7Click
        end
      end
      object GroupBox2: TGroupBox
        Left = 3
        Top = 171
        Width = 354
        Height = 161
        Caption = #1054#1089#1085#1086#1074#1085#1099#1077' '#1096#1083#1102#1079#1099
        TabOrder = 1
        object ListView2: TListView
          Left = 7
          Top = 16
          Width = 341
          Height = 98
          Columns = <
            item
              Caption = #1064#1083#1102#1079
              Width = 165
            end
            item
              Caption = #1052#1077#1090#1088#1080#1082#1072
              Width = 165
            end>
          GridLines = True
          RowSelect = True
          TabOrder = 0
          ViewStyle = vsReport
        end
        object Button3: TButton
          Left = 92
          Top = 120
          Width = 75
          Height = 25
          Caption = #1044#1086#1073#1072#1074#1080#1090#1100
          TabOrder = 1
          OnClick = Button3Click
        end
        object Button4: TButton
          Left = 270
          Top = 120
          Width = 75
          Height = 25
          Caption = #1059#1076#1072#1083#1080#1090#1100
          TabOrder = 2
          OnClick = Button4Click
        end
        object Button10: TButton
          Left = 181
          Top = 120
          Width = 75
          Height = 25
          Caption = #1048#1079#1084#1077#1085#1080#1090#1100
          TabOrder = 3
          OnClick = Button10Click
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'DNS/WINS'
      ImageIndex = 1
      object Label1: TLabel
        Left = 2
        Top = 10
        Width = 249
        Height = 13
        Caption = #1040#1076#1088#1077#1089#1072' DNS-'#1089#1077#1088#1074#1077#1088#1086#1074', '#1074' '#1087#1086#1088#1103#1076#1082#1077' '#1080#1089#1087#1086#1083#1100#1079#1086#1074#1072#1085#1080#1103
      end
      object Label2: TLabel
        Left = 0
        Top = 168
        Width = 209
        Height = 13
        Caption = 'WINS-'#1072#1076#1088#1077#1089#1072', '#1074' '#1087#1086#1088#1103#1076#1082#1077' '#1080#1089#1087#1086#1083#1100#1079#1086#1074#1072#1085#1080#1103':'
      end
      object ListBox1: TListBox
        Left = 2
        Top = 28
        Width = 354
        Height = 97
        ItemHeight = 13
        MultiSelect = True
        TabOrder = 0
      end
      object Button5: TButton
        Left = 109
        Top = 131
        Width = 75
        Height = 25
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        TabOrder = 1
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 281
        Top = 131
        Width = 75
        Height = 25
        Caption = #1059#1076#1072#1083#1080#1090#1100
        TabOrder = 2
        OnClick = Button6Click
      end
      object ListBox2: TListBox
        Left = 2
        Top = 184
        Width = 354
        Height = 97
        ItemHeight = 13
        MultiSelect = True
        TabOrder = 3
      end
      object Button8: TButton
        Left = 109
        Top = 287
        Width = 75
        Height = 25
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        TabOrder = 4
        OnClick = Button8Click
      end
      object Button9: TButton
        Left = 281
        Top = 287
        Width = 75
        Height = 25
        Caption = #1059#1076#1072#1083#1080#1090#1100
        TabOrder = 5
        OnClick = Button9Click
      end
      object Button11: TButton
        Left = 196
        Top = 131
        Width = 75
        Height = 25
        Caption = #1048#1079#1084#1077#1085#1080#1090#1100
        TabOrder = 6
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 196
        Top = 287
        Width = 75
        Height = 25
        Caption = #1048#1079#1084#1077#1085#1080#1090#1100
        TabOrder = 7
        OnClick = Button12Click
      end
    end
  end
end
