object Form1: TForm1
  Left = 0
  Top = 0
  Caption = '0000'
  ClientHeight = 344
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
  object Label1: TLabel
    Left = 552
    Top = 311
    Width = 63
    Height = 13
    Caption = 'Label1'
  end
  object StringGrid1: TStringGrid
    Left = 8
    Top = 39
    Width = 619
    Height = 252
    TabOrder = 0
  end
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 145
    Height = 25
    Caption = 'Open xlsx/xls  with call back'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 512
    Top = 8
    Width = 115
    Height = 25
    Caption = 'Write Demo xlsx'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 159
    Top = 8
    Width = 162
    Height = 25
    Caption = 'Open xlsx/xls'
    TabOrder = 3
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 32
    Top = 306
    Width = 121
    Height = 25
    Caption = 'DOCX writer test'
    TabOrder = 4
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 159
    Top = 306
    Width = 105
    Height = 25
    Caption = 'RTF writer test'
    TabOrder = 5
    OnClick = Button5Click
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel file|*.xlsx|Old Excel|*.xls'
    Left = 440
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '.xslx'
    Filter = 'Excel File|*.xlsx'
    Left = 480
  end
end
