object OKRightDlg12345678910111213141516171819: TOKRightDlg12345678910111213141516171819
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = 'DNS'
  ClientHeight = 73
  ClientWidth = 198
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 11
    Width = 31
    Height = 13
    Caption = 'Label1'
  end
  object OKBtn: TButton
    Left = 108
    Top = 35
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
    Top = 35
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object Edit1: TEdit
    Left = 62
    Top = 8
    Width = 121
    Height = 21
    TabOrder = 2
  end
end
