object Form11: TForm11
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'FormForCopyFF'
  ClientHeight = 207
  ClientWidth = 462
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object TaskDialog1: TTaskDialog
    Buttons = <>
    CommonButtons = [tcbOk, tcbYes, tcbNo, tcbCancel, tcbRetry, tcbClose]
    ExpandButtonCaption = 'EButCaption'
    ExpandedText = 'ExpandedText'
    Flags = [tfEnableHyperlinks, tfAllowDialogCancellation, tfExpandFooterArea, tfExpandedByDefault, tfVerificationFlagChecked, tfShowProgressBar, tfShowMarqueeProgressBar, tfCallbackTimer, tfPositionRelativeToWindow, tfNoDefaultRadioButton, tfCanBeMinimized]
    FooterIcon = 2
    FooterText = 'Fotertext'
    RadioButtons = <>
    Text = 'text'
    Title = 'title'
    VerificationText = #1087#1088#1086#1074#1077#1088#1082#1072' '#1080' '#1089#1086#1079#1076#1072#1085#1080#1077' '#1076#1080#1088#1077#1082#1090#1086#1088#1080#1080' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1103'.'
    OnVerificationClicked = TaskDialog1VerificationClicked
    Left = 400
    Top = 120
  end
end
