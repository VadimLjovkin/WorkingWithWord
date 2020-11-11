object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 337
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonReplace: TButton
    Left = 208
    Top = 32
    Width = 401
    Height = 25
    Caption = #1053#1072#1081#1090#1080' '#1080' '#1047#1072#1084#1077#1085#1080#1090#1100' '#1090#1077#1082#1089#1090' '#1074' '#1092#1072#1081#1083#1077' Word'
    TabOrder = 0
    OnClick = ButtonReplaceClick
  end
  object WordApplication1: TWordApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 40
    Top = 16
  end
  object WordDocument1: TWordDocument
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 40
    Top = 72
  end
  object WordFont1: TWordFont
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 40
    Top = 152
  end
  object WordParagraphFormat1: TWordParagraphFormat
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 40
    Top = 216
  end
end
