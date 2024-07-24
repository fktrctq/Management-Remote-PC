object OKRightDlg123456789101112131415161718: TOKRightDlg123456789101112131415161718
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1040#1076#1088#1077#1089' '#1096#1083#1102#1079#1072' TCP/IP'
  ClientHeight = 76
  ClientWidth = 192
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 19
    Width = 34
    Height = 13
    Caption = #1064#1083#1102#1079':'
  end
  object OKBtn: TButton
    Left = 109
    Top = 43
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 8
    Top = 43
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object Edit1: TEdit
    Left = 60
    Top = 16
    Width = 121
    Height = 21
    TabOrder = 2
  end
end
