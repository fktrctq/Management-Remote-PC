object OKRightDlg1234567891011121314151617: TOKRightDlg1234567891011121314151617
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = 'TCP/IP-'#1072#1076#1088#1077#1089
  ClientHeight = 91
  ClientWidth = 259
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 11
    Width = 48
    Height = 13
    Caption = 'IP-'#1072#1076#1088#1077#1089':'
  end
  object Label2: TLabel
    Left = 8
    Top = 35
    Width = 80
    Height = 13
    Caption = #1052#1072#1089#1082#1072' '#1087#1086#1076#1089#1077#1090#1080':'
  end
  object OKBtn: TButton
    Left = 88
    Top = 59
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 176
    Top = 59
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object Edit1: TEdit
    Left = 125
    Top = 8
    Width = 121
    Height = 21
    TabOrder = 2
  end
  object Edit2: TEdit
    Left = 125
    Top = 32
    Width = 121
    Height = 21
    TabOrder = 3
  end
end
