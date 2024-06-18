Attribute VB_Name = "Shipping_Mark_EXCELL"
Sub generador_Shipping_Mark_Excell()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Worksheets("Shipping mark")
    Dim rng As Range: Set rng = ws.Cells(1, 1)
    Dim rngoff As Range: Set rngoff = rng.Offset(7, 3)
    Dim matrz(1 To 12, 149) As String
    Dim cantidadEtiquetas As Integer

    ws.Cells.Clear
    
    cantidadEtiquetas = 0

    Call ponercero(matrz)

    Sheets("Master").Select
    Range("Z132").Activate

    For i = 0 To 149
        Range("Z132").Activate ' nro de parte
        If (StrComp("", ActiveCell.Offset(i, 0), 0) = 0) Then

        Else
        cantidadEtiquetas = cantidadEtiquetas + 1

        End If





    Next i


    For j = 0 To cantidadEtiquetas

        matrz(1, j) = ActiveCell.Offset(j, 0).Value 'nro de parte
        matrz(2, j) = ActiveCell.Offset(j, -2).Value & " / " & ActiveCell.Offset(j, -1).Value ' descripcion en ingles y español
        matrz(3, j) = ActiveCell.Offset(j, 157).Value ' cantidad total
        matrz(4, j) = "1/100" 'numero de carton
        matrz(5, j) = "100"  'Qty Carton
        'matrz(6,j) = PO number
        'matrz(7,j) = net weight
        'matrz(8,j) = gross weight
        'matrz(9,j) = Dimension
        matrz(10, j) = ActiveCell.Offset(j, 163) ' ORIGEN
'        matrz(11, j) = ActiveCell.Offset(j, -15) ' Brand
'        matrz(12, j) = ActiveCell.Offset(j, -14) ' Model
'

    Next j

    Sheets("Shipping mark").Select
    Call formato(rng, rngoff, matrz, cantidadEtiquetas)


'    Worksheet("MASTER").Select



End Sub

'1- numero de parte
'2- Descripcion en ingles y español
'3- cantidad total en el batch
'4- Carton No. : nro de carton
'5- PO number
'6- Qty Carton : cantidad en el carton
'10- Origin

' FORMATO
'Part numer: m1
'English / Spanish Description:     m2      m3
'Chinese Description:           Carton No: m4
'PO Number: m5          Qty Carton: m6
'Net weight:
'Gros weight:
'Dimension:
'Origin:  m7

Sub formato(ByRef celda As Range, ByRef celdaoffset As Range, ByRef m() As String, ByRef cant As Integer)

    Sheets("Shipping mark").Select


    For x = 0 To cant - 2

                celda.Offset(0, 1).Value = Code128(m(1, x)) & Chr(10) & (m(1, x))  'nro de parte CODIGO
                celda.Offset(1, 1).Value = m(2, x)  ' descripcion en ingles y español.Value 'nro de parte
                celda.Offset(1, 2).Value = "Shipment Total Qty " & m(3, x)  ' cantidad total


                celda.Offset(2, 3).Value = Code128(m(4, x)) & Chr(10) & (m(4, x))  'carton No. CODIGO

                celda.Offset(3, 3).Value = Code128(m(5, x)) & Chr(10) & (m(5, x))  'Qty Carton
                celda.Offset(3, 1).Value = "EL06-01YT" & Chr(10) & "EL06-01YT"      'PO Number
                'celda.Offset(4, 1).Value = m(11, x) ' Brand
                'celda.Offset(5, 1).Value = m(12, x) ' Model
                celda.Offset(7, 1).Value = m(10, x)  ' ORIGEN

                Call formato2(celda, celdaoffset)
                'Funcion para generar shipping mark en el excell
                celda.Offset(13, 0).Select
                Set celda = Selection
                
                celdaoffset.Offset(13, 0).Select
                Set celdaoffset = Selection
                
                'Funciones para generar archivos word
                'Call Generador_Shipping_Word1(m(1, x))
                'Call CargarDatosWord(m(1, x))

    Next x





End Sub




Sub formato2(ByRef celda As Range, ByRef celdaoffset As Range)
    Sheets("Shipping mark").Select
    Columns("D:D").ColumnWidth = 15.21
    Columns("C:C").ColumnWidth = 16.13
    Columns("B:B").ColumnWidth = 33
    Columns("A:A").ColumnWidth = 12.71
    
    
                                'Rango 1 - celda
    Range(celda.Address(), celdaoffset.Address()).Select
                    'formato de texto
                        With Selection
                            .WrapText = True
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                           With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideVertical)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlInsideHorizontal)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With

     'celda.Range(celda, celda.Offset(9, 0)).Select
Range(celda.Address(), celdaoffset.Address()).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With


    celda.RowHeight = 53.25  'Part no
    celda.Value = "Part Number:"
    celda.WrapText = True

    'formato de celdas
   Range(celda.Offset(0, 1).Address(), celda.Offset(0, 3).Address()).Select
    With Selection
            .Merge
            

    End With
      Range(celda.Offset(1, 2).Address(), celda.Offset(1, 3).Address()).Select
    With Selection
            .Merge

    End With
      Range(celda.Offset(4, 1).Address(), celda.Offset(4, 3).Address()).Select
    With Selection
            .Merge

    End With
          Range(celda.Offset(5, 1).Address(), celda.Offset(5, 3).Address()).Select
    With Selection
            .Merge

    End With
          Range(celda.Offset(6, 1).Address(), celda.Offset(6, 3).Address()).Select
    With Selection
            .Merge

    End With
          Range(celda.Offset(7, 1).Address(), celda.Offset(7, 3).Address()).Select
    With Selection
            .Merge

    End With
    Range(celda.Offset(8, 1).Address(), celda.Offset(8, 3).Address()).Select
    With Selection
            .Merge

    End With
    Range(celda.Offset(9, 1).Address(), celda.Offset(9, 3).Address()).Select
    With Selection
            .Merge

    End With



    celda.Offset(1, 0).RowHeight = 60 ' Descripcion en ingles y en español
    celda.Offset(1, 0).Value = "English / Spanish Description:"


    celda.Offset(2, 0).RowHeight = 45 ' Chinese description
    celda.Offset(2, 0).Value = "Chinese Description:"

    celda.Offset(3, 0).RowHeight = 76.5  ' PO number
    celda.Offset(3, 0).Value = "PO Number:"

'    celda.Offset(4, 0).RowHeight = 15  ' brand
'    celda.Offset(4, 0).Value = "Brand:"
'
'    celda.Offset(5, 0).RowHeight = 15  ' model
'    celda.Offset(5, 0).Value = "Model:"

    celda.Offset(4, 0).RowHeight = 15  ' Net weight
    celda.Offset(4, 0).Value = "Net Weight:"

    celda.Offset(5, 0).RowHeight = 15      'Gross Weight
    celda.Offset(5, 0).Value = "Gross Weight:"

    celda.Offset(6, 0).RowHeight = 15      'Dimension
    celda.Offset(6, 0).Value = "Dimension:"

    celda.Offset(7, 0).RowHeight = 21      'Origin
    celda.Offset(7, 0).Value = "Origin:"


    celda.Offset(2, 2).Value = "Carton No.:"
    celda.Offset(3, 2).Value = "Qty Carton:"


    
    Range("B1:D1").Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


    Range("D3").Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("D4").Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B4").Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    'Color
    celda.Offset(0, 1).Select  'B1:D1
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
 
    celda.Offset(3, 1).Select 'B4
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
    celda.Offset(2, 3).Select 'D3
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With

    celda.Offset(3, 3).Select 'D4
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
    

    
    'CAMBIAR LAS FUENTES DE DOS MITADES
    
    
    
    'INTENTO DINAMICO 1
    With celda.Offset(0, 1).Characters(Start:=1, Length:=23).Font 'B1:D1
        .Name = "code 128"
        .FontStyle = "Normal"
        .Size = 22
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
        .TintAndShade = 0
    End With
    With celda.Offset(0, 1).Characters(Start:=24, Length:=21).Font 'B1:D1
        .Name = "Century Gothic"
        .FontStyle = "Normal"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    
    End With
    
    With celda.Offset(3, 1).Characters(Start:=1, Length:=10).Font    'B4
        .Name = "Code 128"
        .Size = 22
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
     With celda.Offset(3, 1).Characters(Start:=11, Length:=9).Font    'B4
        .Name = "Century Gothic"
        .Size = 12
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
    With celda.Offset(2, 3).Characters(Start:=1, Length:=9).Font    'D3
        .Name = "Code 128"
        .Size = 17
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
     With celda.Offset(2, 3).Characters(Start:=10, Length:=5).Font    'D3
        .Name = "Century Gothic"
        .Size = 11
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
    
    With celda.Offset(3, 3).Characters(Start:=1, Length:=7).Font    'D4
        .Name = "Code 128"
        .Size = 17
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
     With celda.Offset(3, 3).Characters(Start:=8, Length:=3).Font    'D4
        .Name = "Century Gothic"
        .Size = 11
        .Shadow = False
        .Strikethrough = False
        .OutlineFont = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    
    End With
    
    
    'NO DINAMICO
    
        
'    With Range("B1:D1").Characters(Start:=1, Length:=10).Font
'        .Name = "Code 128"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    With Range("B1:D1").Characters(Start:=15, Length:=15).Font
'        .Name = "Century Gothic"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'
'    With Range("B1:D1").Characters(Start:=28, Length:=2).Font
'        .Name = "Century Gothic"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    With Range("B4").Characters(Start:=1, Length:=10).Font
'        .Name = "code 128"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    With Range("B4").Characters(Start:=11, Length:=9).Font
'        .Name = "Century Gothic"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'        With Range("D4").Characters(Start:=1, Length:=7).Font
'        .Name = "code 128"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    With Range("D4").Characters(Start:=8, Length:=3).Font
'        .Name = "Century Gothic"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'       With Range("D3").Characters(Start:=1, Length:=9).Font
'        .Name = "code 128"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    With Range("D3").Characters(Start:=10, Length:=5).Font
'        .Name = "Century Gothic"
'        .FontStyle = "Normal"
'        .Size = 13
'        .Strikethrough = False
'        .Superscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With


 
Range(celda.Address(), celdaoffset.Address()).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
End With

End Sub

Sub ponercero(ByRef m() As String)

    Dim i As Integer

    For i = 1 To 10

        For j = 0 To 149
                        m(i, j) = "0"
        Next j

    Next i



End Sub



