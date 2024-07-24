object OKRightDlg1234567: TOKRightDlg1234567
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1048#1084#1103' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1072
  ClientHeight = 194
  ClientWidth = 184
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object OKBtn: TButton
    Left = 96
    Top = 158
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
    Top = 159
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object NamePCEdit: TLabeledEdit
    Left = 11
    Top = 20
    Width = 160
    Height = 21
    EditLabel.Width = 118
    EditLabel.Height = 13
    EditLabel.Caption = #1053#1086#1074#1086#1077' '#1080#1084#1103' '#1082#1086#1084#1087#1100#1102#1090#1077#1088#1072
    TabOrder = 2
  end
  object UserAdmin: TLabeledEdit
    Left = 11
    Top = 62
    Width = 160
    Height = 21
    EditLabel.Width = 79
    EditLabel.Height = 13
    EditLabel.Caption = #1040#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088
    TabOrder = 3
  end
  object PassAdmin: TLabeledEdit
    Left = 8
    Top = 104
    Width = 163
    Height = 21
    EditLabel.Width = 37
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1072#1088#1086#1083#1100
    PasswordChar = #7
    TabOrder = 4
  end
  object RebootAfter: TCheckBox
    Left = 8
    Top = 133
    Width = 163
    Height = 17
    Caption = #1055#1077#1088#1077#1079#1072#1075#1088#1091#1079#1080#1090#1100' '#1082#1086#1084#1087#1100#1102#1090#1077#1088
    TabOrder = 5
  end
end
