object OKRightDlg1234567891011121314: TOKRightDlg1234567891011121314
  Left = 227
  Top = 108
  BorderStyle = bsDialog
  Caption = #1057#1074#1086#1081#1089#1090#1074#1072
  ClientHeight = 265
  ClientWidth = 367
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object CancelBtn: TButton
    Left = 275
    Top = 231
    Width = 75
    Height = 25
    Cancel = True
    Caption = #1047#1072#1082#1088#1099#1090#1100
    ModalResult = 2
    TabOrder = 0
  end
  object ListView1: TListView
    Left = -1
    Top = 0
    Width = 370
    Height = 225
    Columns = <
      item
        Width = 360
      end>
    GridLines = True
    RowSelect = True
    TabOrder = 1
    ViewStyle = vsReport
  end
end
