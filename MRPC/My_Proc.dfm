object EditMyProc: TEditMyProc
  Left = 0
  Top = 0
  Caption = #1056#1077#1076#1072#1082#1090#1086#1088' '#1087#1088#1080#1083#1086#1078#1077#1085#1080#1081
  ClientHeight = 311
  ClientWidth = 710
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ListView1: TListView
    Left = 0
    Top = 33
    Width = 710
    Height = 278
    Align = alClient
    Columns = <
      item
        Caption = #8470
        Width = 0
      end
      item
        Caption = #1048#1084#1103
        Width = 150
      end
      item
        Caption = #1060#1072#1081#1083' '#1079#1072#1087#1091#1089#1082#1072
        Width = 250
      end
      item
        Caption = #1040#1088#1075#1091#1084#1077#1085#1090#1099
        Width = 250
      end>
    GridLines = True
    MultiSelect = True
    ReadOnly = True
    RowSelect = True
    PopupMenu = PopupMenu1
    TabOrder = 0
    ViewStyle = vsReport
    OnDblClick = ListView1DblClick
    ExplicitTop = 47
    ExplicitHeight = 270
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 710
    Height = 33
    Align = alTop
    TabOrder = 1
    object Button1: TButton
      Left = 12
      Top = 3
      Width = 75
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 264
      Top = 3
      Width = 75
      Height = 25
      Caption = #1059#1076#1072#1083#1080#1090#1100
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 180
      Top = 3
      Width = 75
      Height = 25
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100
      TabOrder = 2
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 96
      Top = 3
      Width = 75
      Height = 25
      Caption = #1048#1079#1084#1077#1085#1080#1090#1100
      TabOrder = 3
      OnClick = Button4Click
    end
  end
  object PopupMenu1: TPopupMenu
    Images = frmDomainInfo.ImageButton
    Left = 384
    Top = 112
    object N1: TMenuItem
      Caption = #1057#1086#1079#1076#1072#1090#1100
      ImageIndex = 0
      OnClick = Button1Click
    end
    object N2: TMenuItem
      Caption = #1048#1079#1084#1077#1085#1080#1090#1100
      ImageIndex = 1
      OnClick = Button4Click
    end
    object N3: TMenuItem
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100
      ImageIndex = 11
      OnClick = Button3Click
    end
    object N4: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100
      ImageIndex = 5
      OnClick = Button2Click
    end
  end
  object PopupMyCommand: TPopupMenu
    Left = 104
    Top = 112
  end
end
