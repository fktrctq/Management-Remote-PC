object RemoteDesktopSettionTransform: TRemoteDesktopSettionTransform
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
  ClientHeight = 214
  ClientWidth = 202
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 1
    Top = 1
    Width = 198
    Height = 174
    Shape = bsFrame
  end
  object Label1: TLabel
    Left = 16
    Top = 120
    Width = 76
    Height = 13
    Caption = #1055#1088#1072#1074#1072' '#1076#1086#1089#1090#1091#1087#1072
  end
  object OKBtn: TButton
    Left = 110
    Top = 181
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 16
    Top = 181
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 16
    Top = 21
    Width = 169
    Height = 21
    EditLabel.Width = 72
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1100
    TabOrder = 2
  end
  object LabeledEdit2: TLabeledEdit
    Left = 16
    Top = 56
    Width = 169
    Height = 21
    EditLabel.Width = 37
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1072#1088#1086#1083#1100
    PasswordChar = #7
    TabOrder = 3
  end
  object LabeledEdit3: TLabeledEdit
    Left = 16
    Top = 93
    Width = 169
    Height = 21
    EditLabel.Width = 76
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1086#1074#1090#1086#1088' '#1087#1072#1088#1086#1083#1103
    PasswordChar = #7
    TabOrder = 4
  end
  object ComboBox1: TComboBox
    Left = 16
    Top = 135
    Width = 169
    Height = 22
    AutoComplete = False
    Style = csOwnerDrawFixed
    ItemIndex = 0
    TabOrder = 5
    Text = #1055#1086#1083#1085#1099#1081' '#1076#1086#1089#1090#1091#1087
    Items.Strings = (
      #1055#1086#1083#1085#1099#1081' '#1076#1086#1089#1090#1091#1087
      #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077
      #1055#1088#1086#1089#1084#1086#1090#1088)
  end
end
