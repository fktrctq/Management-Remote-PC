object MainForm: TMainForm
  Left = 260
  Top = 107
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'OS Version information'
  ClientHeight = 65
  ClientWidth = 214
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 4
    Top = 4
    Width = 89
    Height = 13
    Caption = 'Operating system:'
  end
  object Label2: TLabel
    Left = 4
    Top = 22
    Width = 59
    Height = 13
    Caption = 'Architecture'
  end
  object Label3: TLabel
    Left = 4
    Top = 40
    Width = 44
    Height = 13
    Caption = 'Updates:'
  end
  object L3: TLabel
    Left = 98
    Top = 40
    Width = 3
    Height = 13
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object L2: TLabel
    Left = 98
    Top = 22
    Width = 3
    Height = 13
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object L1: TLabel
    Left = 98
    Top = 4
    Width = 3
    Height = 13
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
end
