object OKRightDlg12: TOKRightDlg12
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1059#1076#1072#1083#1077#1085#1080#1077
  ClientHeight = 75
  ClientWidth = 421
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 44
    Width = 216
    Height = 13
    Caption = '/q, /quiet , /verysilent, /silent, /qn '#1080' '#1076#1088#1091#1075#1080#1077' '
  end
  object Label2: TLabel
    Left = 8
    Top = 58
    Width = 117
    Height = 13
    Caption = #1089#1084#1086#1090#1088#1080#1084' '#1074' '#1080#1085#1090#1077#1088#1077#1085#1077#1090#1077'.'
  end
  object OKBtn: TButton
    Left = 338
    Top = 44
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 253
    Top = 44
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 8
    Top = 17
    Width = 405
    Height = 21
    EditLabel.Width = 390
    EditLabel.Height = 13
    EditLabel.Caption = 
      #1057#1090#1088#1086#1082#1072' Uninstall ('#1087#1088#1080' '#1085#1077#1086#1073#1093#1086#1076#1080#1084#1086#1089#1090#1080' '#1076#1086#1073#1072#1074#1080#1090#1100' '#1082#1083#1102#1095#1080' '#1076#1083#1103' '#1090#1080#1093#1086#1075#1086' '#1091#1076 +
      #1072#1083#1077#1085#1080#1103')'
    TabOrder = 2
  end
end
