object frmMain: TfrmMain
  Left = 100
  Top = 100
  Caption = 'Kill Outlook Tasks'
  ClientHeight = 461
  ClientWidth = 852
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object StatusBar1: TStatusBar
    Left = 0
    Top = 442
    Width = 852
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = 'Copyright 2018 John M, Wargo'
  end
  object output: TMemo
    Left = 0
    Top = 0
    Width = 852
    Height = 442
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Consolas'
    Font.Style = []
    ParentFont = False
    ScrollBars = ssVertical
    TabOrder = 1
  end
end
