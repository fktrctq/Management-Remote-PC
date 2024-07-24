object OKRightDlg1234567891011121314151617181920: TOKRightDlg1234567891011121314151617181920
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1055#1086#1076#1089#1077#1090#1100
  ClientHeight = 119
  ClientWidth = 230
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 19
    Top = 8
    Width = 44
    Height = 13
    Caption = 'IP-'#1072#1076#1088#1077#1089
  end
  object Label2: TLabel
    Left = 19
    Top = 51
    Width = 57
    Height = 13
    Caption = #1052#1072#1089#1082#1072' '#1089#1077#1090#1080
  end
  object CancelBtn: TButton
    Left = 143
    Top = 25
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    ModalResult = 2
    TabOrder = 3
  end
  object Edit1: TEdit
    Left = 16
    Top = 27
    Width = 121
    Height = 21
    TabOrder = 0
    Text = '192.168.0.0'
  end
  object Edit2: TEdit
    Left = 16
    Top = 70
    Width = 121
    Height = 21
    TabOrder = 1
    Text = '255.255.255.0'
  end
  object Button1: TButton
    Left = 143
    Top = 68
    Width = 75
    Height = 25
    Caption = 'Ok'
    TabOrder = 2
    OnClick = Button1Click
  end
  object CheckBox1: TCheckBox
    Left = 16
    Top = 97
    Width = 158
    Height = 17
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1080#1079' '#1073#1072#1079#1099
    TabOrder = 4
    OnMouseUp = CheckBox1MouseUp
  end
  object IdNetworkCalculator1: TIdNetworkCalculator
    NetworkAddress.AsString = '192.168.0.1'
    NetworkMask.AsString = '255.255.255.0'
    NetworkMaskLength = 24
    Left = 136
    Top = 65528
  end
end
