Attribute VB_Name = "SIMI_Actualizador"
Sub ActualizarPavas()

Dim cantidadpaginas As Long
Dim MatrizTv(1 To 5, 16, 149) As String
Dim cantelementos(1 To 5) As Integer

Call borrarmatriz(MatrizTv)

Call vectorencero2(cantelementos())
Call copiarenmatriz(MatrizTv, cantelementos())



cantidadpaginas = ThisWorkbook.Sheets.Count

Sheets("Master").Select


For i = 4 To cantidadpaginas
    'BATCH 1
    If (i = 4) Then
    Sheets("Batch 1").Select
    Call Borrar
         For j = 0 To 16
        
                For m = 0 To cantelementos(1)
                 If (StrComp("0", MatrizTv(1, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 12 Or j = 13 Or j = 14) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(1, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(1, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
            Next j
    
    End If
    'BATCH 2
    If (i = 5) Then
    Sheets("Batch 2").Select
    Call Borrar
         For j = 0 To 16
        
                For m = 0 To cantelementos(2)
                 If (StrComp("0", MatrizTv(2, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 12 Or j = 13 Or j = 14) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(2, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(2, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
            Next j
    
    End If
    'BATCH 3
    If (i = 6) Then
    Sheets("Batch 3").Select
    Call Borrar
         For j = 0 To 16
        
                For m = 0 To cantelementos(3)
                 If (StrComp("0", MatrizTv(3, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 12 Or j = 13 Or j = 14) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(3, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(3, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
            Next j
    
    End If
    
    'BATCH 4
    If (i = 7) Then
    Sheets("Batch 4").Select
    Call Borrar
         For j = 0 To 16
        
                For m = 0 To cantelementos(4)
                 If (StrComp("0", MatrizTv(4, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 12 Or j = 13 Or j = 14) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(4, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(4, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
            Next j
    
    End If
    'BATCH 5
    If (i = 8) Then
    Sheets("Batch 5").Select
    Call Borrar
         For j = 0 To 16
        
                For m = 0 To cantelementos(5)
                 If (StrComp("0", MatrizTv(5, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 12 Or j = 13 Or j = 14) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(5, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(5, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
            Next j
    
    End If
    
    
    
    
Next i


For i = 4 To cantidadpaginas
    If (i = 4) Then
    Sheets("Batch 1").Select
    Call darformato
    End If
    If (i = 5) Then
    Sheets("Batch 2").Select
    Call darformato
    End If
    If (i = 6) Then
    Sheets("Batch 3").Select
    Call darformato
    End If
    If (i = 7) Then
    Sheets("Batch 4").Select
    Call darformato
    End If
    If (i = 8) Then
    Sheets("Batch 5").Select
    Call darformato
    End If
    

Next i
 
       
End Sub

Sub borrarmatriz(ByRef m() As String)


For Z = 1 To 5

    For x = 0 To 16
    
    
                For j = 0 To 149
                    
                        m(Z, x, j) = "0"
                
                
                Next j
                
    
    Next x
    
Next Z







End Sub

Sub vectorencero(ByRef vector() As String)
    
    For x = 0 To 10
                
                vector(x) = "0"
        Next x

End Sub

Sub vectorencero2(ByRef vector() As Integer)
    
    For x = 1 To 5
                
                vector(x) = "0"
        Next x

End Sub

Sub vectorenfalso(ByRef vexiste() As Boolean)
    
    For x = 1 To 5
                
                vexiste(x) = False
        Next x

End Sub

Sub indicenfalso(ByRef vexiste() As Boolean)
    
    For x = 0 To 50
                
                vexiste(x) = False
        Next x

End Sub

Sub copiarenmatriz(ByRef m() As String, ByRef cantelementos() As Integer)


    Sheets("Master").Select
    
 For x = 0 To 149
            Range("GL132").Activate
            If (StrComp("1", ActiveCell.Offset(x, 0).Value, 0) = 0) Then

                            
                            ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            
                            m(1, 3, cantelementos(1)) = ActiveCell.Offset(0, -170).Value 'Englis
                            m(1, 4, cantelementos(1)) = ActiveCell.Offset(0, -169).Value 'Spanish
                           
                            
                            'm(1, 5, cantelementos(1)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            m(1, 5, cantelementos(1)) = ActiveCell.Offset(0, -11).Value  'Qty
                            m(1, 6, cantelementos(1)) = ActiveCell.Offset(0, -12).Value  'Unit of measurement
                            m(1, 7, cantelementos(1)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            m(1, 8, cantelementos(1)) = ActiveCell.Offset(0, -8).Value  'Fob total
                    
                            m(1, 9, cantelementos(1)) = ActiveCell.Offset(0, -7).Value  'NW unit

      
                            m(1, 10, cantelementos(1)) = ActiveCell.Offset(0, -5).Value  'Origen
                            m(1, 11, cantelementos(1)) = ActiveCell.Offset(0, -4).Value  'NCM
                            m(1, 12, cantelementos(1)) = ActiveCell.Offset(0, -3).Value  'derechos
                            m(1, 13, cantelementos(1)) = ActiveCell.Offset(0, -2).Value  'TE
                            m(1, 14, cantelementos(1)) = ActiveCell.Offset(0, -1).Value  'IVA
                            m(1, 15, cantelementos(1)) = "1"                             'Batch
                            m(1, 16, cantelementos(1)) = ActiveCell.Offset(0, 1).Value  'LICENCIAS
                            
                         
                            
                            m(1, 2, cantelementos(1)) = ActiveCell.Offset(0, -182).Value  'nro parte en lugar de Model

                            m(1, 1, cantelementos(1)) = ActiveCell.Offset(0, -183).Value  'Brand
                            m(1, 0, cantelementos(1)) = cantelementos(1) + 1
                            
                            
                            
                            cantelementos(1) = cantelementos(1) + 1
                            Range("GL132").Activate

            End If
            If (StrComp("2", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                             
                             ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            m(2, 3, cantelementos(2)) = ActiveCell.Offset(0, -170).Value 'Englis
                            m(2, 4, cantelementos(2)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(2, 5, cantelementos(2)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            'm(2, 5, cantelementos(2)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            m(2, 5, cantelementos(2)) = ActiveCell.Offset(0, -11).Value  'Qty
                            m(2, 6, cantelementos(2)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            m(2, 7, cantelementos(2)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            m(2, 8, cantelementos(2)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                           
                            m(2, 9, cantelementos(2)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            m(2, 10, cantelementos(2)) = ActiveCell.Offset(0, -5).Value  'Origen
                            m(2, 11, cantelementos(2)) = ActiveCell.Offset(0, -4).Value  'NCM
                            m(2, 12, cantelementos(2)) = ActiveCell.Offset(0, -3).Value  'derechos
                            m(2, 13, cantelementos(2)) = ActiveCell.Offset(0, -2).Value  'TE
                            m(2, 14, cantelementos(2)) = ActiveCell.Offset(0, -1).Value  'IVA
                            m(2, 15, cantelementos(2)) = "2"                             'BATCH
                            m(2, 16, cantelementos(2)) = ActiveCell.Offset(0, 1).Value   'LICENCIAS
                            
                            m(2, 2, cantelementos(2)) = ActiveCell.Offset(0, -182).Value  'Model
                            m(2, 1, cantelementos(2)) = ActiveCell.Offset(0, -183).Value  'Brand
                            m(2, 0, cantelementos(2)) = cantelementos(2) + 1
                            
                            
                            cantelementos(2) = cantelementos(2) + 1
                            Range("GL132").Activate
            End If
            If (StrComp("3", ActiveCell.Offset(x, 0).Value, 0) = 0) Then

                          ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            
                            m(3, 3, cantelementos(3)) = ActiveCell.Offset(0, -170).Value 'Englis
                            m(3, 4, cantelementos(3)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(3, 5, cantelementos(3)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            'm(2, 5, cantelementos(2)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            m(3, 5, cantelementos(3)) = ActiveCell.Offset(0, -11).Value  'Qty
                            m(3, 6, cantelementos(3)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            m(3, 7, cantelementos(3)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            m(3, 8, cantelementos(3)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            m(3, 9, cantelementos(3)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            m(3, 10, cantelementos(3)) = ActiveCell.Offset(0, -4).Value  'NCM
                            m(3, 11, cantelementos(3)) = ActiveCell.Offset(0, -3).Value  'derechos
                            m(3, 12, cantelementos(3)) = ActiveCell.Offset(0, -2).Value  'TE
                            m(3, 13, cantelementos(3)) = ActiveCell.Offset(0, -1).Value  'IVA
                            m(3, 14, cantelementos(3)) = "3"
                            m(3, 15, cantelementos(3)) = ActiveCell.Offset(0, 1).Value   'LICENCIAS
                            
                            m(3, 2, cantelementos(3)) = ActiveCell.Offset(0, -182).Value  'Model
                            m(3, 1, cantelementos(3)) = ActiveCell.Offset(0, -183).Value  'Brand
                            m(3, 0, cantelementos(3)) = cantelementos(3) + 1
                            
                            cantelementos(3) = cantelementos(3) + 1
                            Range("GL132").Activate
                            
                        
            End If
            If (StrComp("4", ActiveCell.Offset(x, 0).Value, 0) = 0) Then

                            ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            m(4, 3, cantelementos(4)) = ActiveCell.Offset(0, -170).Value 'Englis
                            m(4, 4, cantelementos(4)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(4, 5, cantelementos(4)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            'm(4, 5, cantelementos(4)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            m(4, 5, cantelementos(4)) = ActiveCell.Offset(0, -11).Value  'Qty
                            m(4, 6, cantelementos(4)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            m(4, 7, cantelementos(4)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            m(4, 8, cantelementos(4)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            m(4, 9, cantelementos(4)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            m(4, 10, cantelementos(4)) = ActiveCell.Offset(0, -5).Value  'Origen
                            m(4, 11, cantelementos(4)) = ActiveCell.Offset(0, -4).Value  'NCM
                            m(4, 12, cantelementos(4)) = ActiveCell.Offset(0, -3).Value  'derechos
                            m(4, 13, cantelementos(4)) = ActiveCell.Offset(0, -2).Value  'TE
                            m(4, 14, cantelementos(4)) = ActiveCell.Offset(0, -1).Value  'IVA
                            m(4, 15, cantelementos(4)) = "4"                             'ORIGEN
                            m(4, 16, cantelementos(4)) = ActiveCell.Offset(0, 1).Value   'LICENCIAS
                            
                            
                            m(4, 2, cantelementos(4)) = ActiveCell.Offset(0, -182).Value  'Model
                            m(4, 1, cantelementos(4)) = ActiveCell.Offset(0, -183).Value  'Brand
                            m(4, 0, cantelementos(4)) = cantelementos(4) + 1
                            

                            cantelementos(4) = cantelementos(4) + 1
                            Range("GL132").Activate
                        
                        
            End If
            If (StrComp("5", ActiveCell.Offset(x, 0).Value, 0) = 0) Then

                          ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            m(5, 3, cantelementos(5)) = ActiveCell.Offset(0, -170).Value 'Englis
                            m(5, 4, cantelementos(5)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(5, 5, cantelementos(5)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            'm(5, 5, cantelementos(5)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            m(5, 5, cantelementos(5)) = ActiveCell.Offset(0, -11).Value  'Qty
                            m(5, 6, cantelementos(5)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            m(5, 7, cantelementos(5)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            m(5, 8, cantelementos(5)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            m(5, 9, cantelementos(5)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            m(5, 10, cantelementos(5)) = ActiveCell.Offset(0, -5).Value  'Origen
                            m(5, 11, cantelementos(5)) = ActiveCell.Offset(0, -4).Value  'NCM
                            m(5, 12, cantelementos(5)) = ActiveCell.Offset(0, -3).Value  'derechos
                            m(5, 13, cantelementos(5)) = ActiveCell.Offset(0, -2).Value  'TE
                            m(5, 14, cantelementos(5)) = ActiveCell.Offset(0, -1).Value  'IVA
                            m(5, 15, cantelementos(5)) = "5"
                            m(5, 16, cantelementos(5)) = ActiveCell.Offset(0, 1).Value   'LICENCIAS
                            
                            
                            m(5, 2, cantelementos(5)) = ActiveCell.Offset(0, -182).Value  'Model
                            m(5, 1, cantelementos(5)) = ActiveCell.Offset(0, -183).Value  'Brand
                            m(5, 0, cantelementos(5)) = cantelementos(5) + 1
                            cantelementos(5) = cantelementos(5) + 1
                            
                            
                            Range("GL132").Activate
                            
                            
            End If
            
            
    Next x


End Sub

Sub Borrar()



Range("A2:Q200").Clear


End Sub

Sub darformato()

Application.CutCopyMode = False
  
    Range("A1:Q1").Select
    Range("M1").Activate
    
                        
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("A1:Q99").Select
    Range("Q99").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    

   
    Range("A2:Q268").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
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
    Range("A1:Q1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
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
    Range("N1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
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

    Range("A1:Q99").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
     Range("A1:Q549").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-573
    Range("P1:Q549").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
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
    Range("P1:Q501").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        Range("A1:Q1").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2:Q314").Select
    ActiveWindow.SmallScroll Down:=-360
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns("H:H").ColumnWidth = 22
   
    
    

End Sub






