VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm1bit12 
   Caption         =   "Imagem BMP 1 bit"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm1bit12.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "bm1bit12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

Option Explicit

Private Sub CommandButton1_Click()

    ' Declarações gerais:
    
    Dim HX As String    ' Dados (hexadecimal)
    Dim BT As String    ' Bytes
    Dim i As Integer    ' Índices
    
    ' Primeira estrutura 'Bitmap File Header' contém informações sobre o tipo,
    ' tamanho e layout de um bitmap e ocupa 14 bytes (padrão).

    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"     O tipo de arquivo ("BM").
    HX = HX & "74000000"    ' BitmapFileSize         DOUBLE WORD    00000074 = 14 + 12 + 6 + 84 = 116 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                          Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                          Reservados (0 byte)
    HX = HX & "20000000"    ' BitmapFileOffBits      DOUBLE WORD    00000020 = 14 + 12 + 6 = 32 bytes          O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Core Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 12 bytes (padrão).

    HX = HX & "0C000000"    ' BitmapCoreSize         DOUBLE WORD    0000000C = 12 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "1200"        ' BitmapCoreWidth        WORD           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "1500"        ' BitmapCoreHeight       WORD           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapCorePlanes       WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "0100"        ' BitmapCoreBitCoun      WORD               0001 = 1 bpp        Especifica o número de bits por pixel.
    
    ' Terceira estrutura 'Palette' só será necessária para bitmaps menores que
    ' 24 bits, quando não for possível inserir as cores RGB ou ARGB de cada
    ' pixel diretamente no bitmap e, como nosso bitmap tem 1 bit e utiliza o
    ' cabeçalho Core/RGB, ela ocupa 2 cores * 3 bytes = 6 bytes.
    
    HX = HX & "000000"      ' 0 Black    000000 = RGB(000, 000, 000)
    HX = HX & "FFFFFF"      ' 1 White    FFFFFF = RGB(255, 255, 255)
    
    ' Quarta estrutura 'Bitmap' contém todos os pixels extrudados em uma matriz
    ' de coluna e linha, onde temos linhas de 0 a 20 = 21 de altura e 18 na
    ' largura, em partes de 32 bits, por esse motivo completamos com 0 (zero)
    ' até obter os completos 32 bits, ela ocupa 21 linhas * 4 bytes = 84 bytes.
    
    '      4bits   4bits   4bits   4bits   4bits   4bits   4bits   4bits      byte
    '     ------- ------- ------- ------- ------- ------- ------- -------   --------
    '  0: 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FFFFC000
    '  1: 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FFFFC000
    '  2: 1 1 1 1 1 1 1 0 0 0 0 1 1 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FE1FC000
    '  3: 1 1 1 1 1 0 0 1 1 1 1 0 0 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F9E7C000
    '  4: 1 1 1 1 0 1 1 1 1 1 1 1 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F7FBC000
    '  5: 1 1 1 1 0 1 1 1 1 1 1 1 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F7FBC000
    '  6: 1 1 1 0 1 1 1 1 1 1 1 1 1 1 0 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = EFFDC000
    '  7: 1 1 1 0 1 1 1 1 1 1 1 1 1 1 0 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = EFFDC000
    '  8: 1 1 1 0 1 1 1 1 1 1 1 1 1 1 0 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = EFFDC000
    '  9: 1 1 1 1 0 1 1 0 1 1 0 1 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F6DBC000
    ' 10: 1 1 1 1 1 0 1 0 1 1 0 1 0 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FAD7C000
    ' 11: 1 1 1 1 0 1 0 1 1 1 1 0 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F5EBC000
    ' 12: 1 1 1 0 1 1 1 0 0 0 0 1 1 1 0 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = EE1DC000
    ' 13: 1 1 0 1 1 0 1 1 1 1 1 1 0 1 1 0 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = DBF6C000
    ' 14: 1 1 0 1 1 0 1 1 1 1 1 1 0 1 1 0 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = DBF6C000
    ' 15: 1 1 1 0 0 1 0 0 0 0 0 0 1 0 0 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = E409C000
    ' 16: 1 1 1 1 0 1 1 1 1 1 1 1 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F7FBC000
    ' 17: 1 1 1 1 0 1 1 1 0 0 1 1 1 0 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F73BC000
    ' 18: 1 1 1 1 1 0 0 0 1 1 0 0 0 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = F8C7C000
    ' 19: 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FFFFC000
    ' 20: 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 = FFFFC000

    ' A imagem é convertida de 'bit' para ser armazenada em 'byte' e para isso
    ' foi utilizada a tabela de conversão:

    '   --------- ----------------
    '     4bits  |  4bits (byte)
    '   --------- ----------------
    '    0 0 0 0 |  0 = 0
    '    0 0 0 1 |  1 = 1
    '    0 0 1 0 |  2 = 2
    '    0 0 1 1 |  3 = 3
    '    0 1 0 0 |  4 = 4
    '    0 1 0 1 |  5 = 5
    '    0 1 1 0 |  6 = 6
    '    0 1 1 1 |  7 = 7
    '    1 0 0 0 |  8 = 8
    '    1 0 0 1 |  9 = 9
    '    1 0 1 0 | 10 = A
    '    1 0 1 1 | 11 = B
    '    1 1 0 0 | 12 = C
    '    1 1 0 1 | 13 = D
    '    1 1 1 0 | 14 = E
    '    1 1 1 1 | 15 = F
    '   --------- ---------------

    HX = HX & "FFFFC000"    ' 20:                                     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "FFFFC000"    ' 19:                                     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F8C7C000"    ' 18:           0 0 0     0 0 0           0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F73BC000"    ' 17:         0       0 0       0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F7FBC000"    ' 16:         0                 0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "E409C000"    ' 15:       0 0   0 0 0 0 0 0   0 0       0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "DBF6C000"    ' 14:     0     0             0     0     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "DBF6C000"    ' 13:     0     0             0     0     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "EE1DC000"    ' 12:       0       0 0 0 0       0       0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F5EBC000"    ' 11:         0   0         0   0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "FAD7C000"    ' 10:           0   0     0   0           0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F6DBC000"    '  9:         0     0     0     0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "EFFDC000"    '  8:       0                     0       0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "EFFDC000"    '  7:       0                     0       0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "EFFDC000"    '  6:       0                     0       0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F7FBC000"    '  5:         0                 0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F7FBC000"    '  4:         0                 0         0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "F9E7C000"    '  3:           0 0         0 0           0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "FE1FC000"    '  2:               0 0 0 0               0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "FFFFC000"    '  1:                                     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    HX = HX & "FFFFC000"    '  0:                                     0 0 0 0 0 0 0 0 0 0 0 0 0 0
    
    ' Salvar arquivo bitmap monocromático (*.bmp;*.dib).
    
    Open Project.ThisDocument.Path & "\~$bm1bit12.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm1bit12.bmp")
    
End Sub
