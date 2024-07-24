object OKRightDlg1234567891011: TOKRightDlg1234567891011
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1055#1077#1088#1077#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1088#1080#1085#1090#1077#1088#1072
  ClientHeight = 89
  ClientWidth = 192
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 103
    Height = 13
    Caption = #1053#1086#1074#1086#1077' '#1080#1084#1103' '#1087#1088#1080#1085#1090#1077#1088#1072
  end
  object OKBtn: TButton
    Left = 102
    Top = 51
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
    Top = 51
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object Edit1: TEdit
    Left = 8
    Top = 24
    Width = 169
    Height = 21
    TabOrder = 2
  end
end
