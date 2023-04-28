object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 169
  ClientWidth = 239
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 32
    Top = 24
    Width = 62
    Height = 13
    Caption = 'Array output'
  end
  object btnProcessArray: TButton
    Left = 32
    Top = 43
    Width = 177
    Height = 25
    Caption = 'Process Array'
    TabOrder = 0
    OnClick = btnProcessArrayClick
  end
  object btnProcessObject: TButton
    Left = 32
    Top = 88
    Width = 177
    Height = 25
    Caption = 'Process Object'
    TabOrder = 1
    OnClick = btnProcessObjectClick
  end
end
