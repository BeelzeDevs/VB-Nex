Attribute VB_Name = "SIMI_Generador"
Sub Plantilla_Generar_Simi()

Dim Vec(10) As String

Dim MatrizTv(1 To 5, 17, 149) As String
Dim contadordepaginas As Integer
Dim auxiliar As Integer
Dim Vbool(1 To 5) As Boolean
Dim cantelementos(1 To 5) As Integer


auxiliar = 0
contadordepaginas = 0
Call borrarmatriz(MatrizTv)
Call vectorencero(Vec())
Call vectorenfalso(Vbool())
Call vectorencero2(cantelementos())


Sheets("Master").Visible = True
Sheets("Master").Select


 
    For x = 0 To 150
            Range("GL132").Activate
            If (StrComp("1", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                            Vbool(1) = True
                            'MatrizTv(
                            
                            ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            
                            MatrizTv(1, 3, cantelementos(1)) = ActiveCell.Offset(0, -170).Value 'Englis
                            MatrizTv(1, 4, cantelementos(1)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(1, 5, cantelementos(1)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            MatrizTv(1, 5, cantelementos(1)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            MatrizTv(1, 6, cantelementos(1)) = ActiveCell.Offset(0, -11).Value  'Qty
                            MatrizTv(1, 7, cantelementos(1)) = ActiveCell.Offset(0, -12).Value  'Unit of measurement
                            MatrizTv(1, 8, cantelementos(1)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            MatrizTv(1, 9, cantelementos(1)) = ActiveCell.Offset(0, -8).Value  'Fob total
                    
                            MatrizTv(1, 10, cantelementos(1)) = ActiveCell.Offset(0, -7).Value  'NW unit

      
                            MatrizTv(1, 11, cantelementos(1)) = ActiveCell.Offset(0, -5).Value  'Origen
                            MatrizTv(1, 12, cantelementos(1)) = ActiveCell.Offset(0, -4).Value  'NCM
                            MatrizTv(1, 13, cantelementos(1)) = ActiveCell.Offset(0, -3).Value  'derechos
                            MatrizTv(1, 14, cantelementos(1)) = ActiveCell.Offset(0, -2).Value  'TE
                            MatrizTv(1, 15, cantelementos(1)) = ActiveCell.Offset(0, -1).Value  'IVA
                            MatrizTv(1, 16, cantelementos(1)) = "1"                             'Batch
                            MatrizTv(1, 17, cantelementos(1)) = ActiveCell.Offset(0, 3).Value  'LICENCIAS
                            
                         
                            
                            MatrizTv(1, 2, cantelementos(1)) = ActiveCell.Offset(0, -182).Value  'nro parte en lugar de Model

                            MatrizTv(1, 1, cantelementos(1)) = ActiveCell.Offset(0, -183).Value  'Brand
                            MatrizTv(1, 0, cantelementos(1)) = cantelementos(1) + 1
                            
                            
                            
                            cantelementos(1) = cantelementos(1) + 1
                            Range("GL132").Activate

            End If
            If (StrComp("2", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                            Vbool(2) = True
                                                        
                             ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            MatrizTv(2, 3, cantelementos(2)) = ActiveCell.Offset(0, -170).Value 'Englis
                            MatrizTv(2, 4, cantelementos(2)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(2, 5, cantelementos(2)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            MatrizTv(2, 5, cantelementos(2)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            MatrizTv(2, 6, cantelementos(2)) = ActiveCell.Offset(0, -11).Value  'Qty
                            MatrizTv(2, 7, cantelementos(2)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            MatrizTv(2, 8, cantelementos(2)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            MatrizTv(2, 9, cantelementos(2)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                           
                            MatrizTv(2, 10, cantelementos(2)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            MatrizTv(2, 11, cantelementos(2)) = ActiveCell.Offset(0, -5).Value  'Origen
                            MatrizTv(2, 12, cantelementos(2)) = ActiveCell.Offset(0, -4).Value  'NCM
                            MatrizTv(2, 13, cantelementos(2)) = ActiveCell.Offset(0, -3).Value  'derechos
                            MatrizTv(2, 14, cantelementos(2)) = ActiveCell.Offset(0, -2).Value  'TE
                            MatrizTv(2, 15, cantelementos(2)) = ActiveCell.Offset(0, -1).Value  'IVA
                            MatrizTv(2, 16, cantelementos(2)) = "2"                             'BATCH
                            MatrizTv(2, 17, cantelementos(2)) = ActiveCell.Offset(0, 3).Value   'LICENCIAS
                            
                            MatrizTv(2, 2, cantelementos(2)) = ActiveCell.Offset(0, -182).Value  'Model
                            MatrizTv(2, 1, cantelementos(2)) = ActiveCell.Offset(0, -183).Value  'Brand
                            MatrizTv(2, 0, cantelementos(2)) = cantelementos(2) + 1
                            
                            
                            cantelementos(2) = cantelementos(2) + 1
                            Range("GL132").Activate
            End If
            If (StrComp("3", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                            Vbool(3) = True
                        
                          ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            
                            MatrizTv(3, 3, cantelementos(3)) = ActiveCell.Offset(0, -170).Value 'Englis
                            MatrizTv(3, 4, cantelementos(3)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(3, 5, cantelementos(3)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            MatrizTv(2, 5, cantelementos(2)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            MatrizTv(3, 6, cantelementos(3)) = ActiveCell.Offset(0, -11).Value  'Qty
                            MatrizTv(3, 7, cantelementos(3)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            MatrizTv(3, 8, cantelementos(3)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            MatrizTv(3, 9, cantelementos(3)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            MatrizTv(3, 10, cantelementos(3)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            MatrizTv(3, 11, cantelementos(3)) = ActiveCell.Offset(0, -5).Value  'Origen
                            MatrizTv(3, 12, cantelementos(3)) = ActiveCell.Offset(0, -4).Value  'NCM
                            MatrizTv(3, 13, cantelementos(3)) = ActiveCell.Offset(0, -3).Value  'derechos
                            MatrizTv(3, 14, cantelementos(3)) = ActiveCell.Offset(0, -2).Value  'TE
                            MatrizTv(3, 15, cantelementos(3)) = ActiveCell.Offset(0, -1).Value  'IVA
                            MatrizTv(3, 16, cantelementos(3)) = "3"
                            MatrizTv(3, 17, cantelementos(3)) = ActiveCell.Offset(0, 3).Value   'LICENCIAS
                            
                            MatrizTv(3, 2, cantelementos(3)) = ActiveCell.Offset(0, -182).Value  'Model
                            MatrizTv(3, 1, cantelementos(3)) = ActiveCell.Offset(0, -183).Value  'Brand
                            MatrizTv(3, 0, cantelementos(3)) = cantelementos(3) + 1
                            
                            cantelementos(3) = cantelementos(3) + 1
                            Range("GL132").Activate
                            
                        
            End If
            If (StrComp("4", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                            Vbool(4) = True
                        
                            ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            MatrizTv(4, 3, cantelementos(4)) = ActiveCell.Offset(0, -170).Value 'Englis
                            MatrizTv(4, 4, cantelementos(4)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(4, 5, cantelementos(4)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            MatrizTv(4, 5, cantelementos(4)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            MatrizTv(4, 6, cantelementos(4)) = ActiveCell.Offset(0, -11).Value  'Qty
                            MatrizTv(4, 7, cantelementos(4)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            MatrizTv(4, 8, cantelementos(4)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            MatrizTv(4, 9, cantelementos(4)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            MatrizTv(4, 10, cantelementos(4)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            MatrizTv(4, 11, cantelementos(4)) = ActiveCell.Offset(0, -5).Value  'Origen
                            MatrizTv(4, 12, cantelementos(4)) = ActiveCell.Offset(0, -4).Value  'NCM
                            MatrizTv(4, 13, cantelementos(4)) = ActiveCell.Offset(0, -3).Value  'derechos
                            MatrizTv(4, 14, cantelementos(4)) = ActiveCell.Offset(0, -2).Value  'TE
                            MatrizTv(4, 15, cantelementos(4)) = ActiveCell.Offset(0, -1).Value  'IVA
                            MatrizTv(4, 16, cantelementos(4)) = "4"                             'ORIGEN
                            MatrizTv(4, 17, cantelementos(4)) = ActiveCell.Offset(0, 3).Value   'LICENCIAS
                            
                            
                            MatrizTv(4, 2, cantelementos(4)) = ActiveCell.Offset(0, -182).Value  'Model
                            MatrizTv(4, 1, cantelementos(4)) = ActiveCell.Offset(0, -183).Value  'Brand
                            MatrizTv(4, 0, cantelementos(4)) = cantelementos(4) + 1
                            

                            cantelementos(4) = cantelementos(4) + 1
                            Range("GL132").Activate
                        
                        
            End If
            If (StrComp("5", ActiveCell.Offset(x, 0).Value, 0) = 0) Then
                        Vbool(5) = True
                        
                          ActiveCell.Offset(x, 0).Activate
                            'lleno mi matriz con los elementos del batch numerado
                            MatrizTv(5, 3, cantelementos(5)) = ActiveCell.Offset(0, -170).Value 'Englis
                            MatrizTv(5, 4, cantelementos(5)) = ActiveCell.Offset(0, -169).Value 'Spanish
                            'MatrizTv(5, 5, cantelementos(5)) = "placa pared" 'IMAGEN BORRAR LUEGO PROTOTIPO
                            
                            MatrizTv(5, 5, cantelementos(5)) = ActiveCell.Offset(0, -168).Value 'IMAGEN Por nro parte
                            MatrizTv(5, 6, cantelementos(5)) = ActiveCell.Offset(0, -11).Value  'Qty
                            MatrizTv(5, 7, cantelementos(5)) = ActiveCell.Offset(0, -12).Value 'Unit of measurement
                            MatrizTv(5, 8, cantelementos(5)) = ActiveCell.Offset(0, -9).Value  'Punitario
                            MatrizTv(5, 9, cantelementos(5)) = ActiveCell.Offset(0, -8).Value  'Fob total
                            
                            MatrizTv(5, 10, cantelementos(5)) = ActiveCell.Offset(0, -7).Value  'NW PER BOX
                            
                            MatrizTv(5, 11, cantelementos(5)) = ActiveCell.Offset(0, -5).Value  'Origen
                            MatrizTv(5, 12, cantelementos(5)) = ActiveCell.Offset(0, -4).Value  'NCM
                            MatrizTv(5, 13, cantelementos(5)) = ActiveCell.Offset(0, -3).Value  'derechos
                            MatrizTv(5, 14, cantelementos(5)) = ActiveCell.Offset(0, -2).Value  'TE
                            MatrizTv(5, 15, cantelementos(5)) = ActiveCell.Offset(0, -1).Value  'IVA
                            MatrizTv(5, 16, cantelementos(5)) = "5"
                            MatrizTv(5, 17, cantelementos(5)) = ActiveCell.Offset(0, 3).Value   'LICENCIAS
                            
                            
                            MatrizTv(5, 2, cantelementos(5)) = ActiveCell.Offset(0, -182).Value  'Model
                            MatrizTv(5, 1, cantelementos(5)) = ActiveCell.Offset(0, -183).Value  'Brand
                            MatrizTv(5, 0, cantelementos(5)) = cantelementos(5) + 1
                            cantelementos(5) = cantelementos(5) + 1
                            
                            
                            Range("GL132").Activate
                            
                            
            End If
            
            
    Next x
    
    For x = 1 To 5
            If (Vbool(x) = True) Then
                contadordepaginas = contadordepaginas + 1
            End If
    
    Next x
    
    
    
With ThisWorkbook
       
For x = 1 To contadordepaginas
    
    If (x = 1) Then
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Batch 1"
    'pegar lo datos correspondientes y darle formato
    Call darformato
'    For m = 0 To cantelementos(1) - 1
'
'            Call insertarimagen(MatrizTv(1, 5, m), m, x)
'    Next m
    
    For j = 0 To 17
        
                For m = 0 To cantelementos(1)
                      If (StrComp("0", MatrizTv(1, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 13 Or j = 14 Or j = 15) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(1, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(1, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
    Next j
    
    
    
    
End If
    If (x = 2) Then
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Batch 2"
    Call darformato
'    For m = 0 To cantelementos(2) - 1
'
'            Call insertarimagen(MatrizTv(2, 5, m), m, x)
'    Next m
            
    For j = 0 To 17
        
                For m = 0 To cantelementos(2)
                 If (StrComp("0", MatrizTv(2, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 13 Or j = 14 Or j = 15) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(2, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(2, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
    Next j
    
End If
    If (x = 3) Then
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Batch 3"
    Call darformato
'    For m = 0 To cantelementos(3) - 1
'
'            Call insertarimagen(MatrizTv(3, 5, m), m, x)
'    Next m
    For j = 0 To 17
        
                For m = 0 To cantelementos(3)
                 If (StrComp("0", MatrizTv(3, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 13 Or j = 14 Or j = 15) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(3, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(3, j, m)
                      
                      
                      
                        End If
                      
                      End If
                Next m
    Next j
    
End If
    
    If (x = 4) Then
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Batch 4"
    Call darformato
'    For m = 0 To cantelementos(4) - 1
'
'            Call insertarimagen(MatrizTv(4, 5, m), m, x)
'    Next m
    For j = 0 To 17
        
                For m = 0 To cantelementos(4)
                 If (StrComp("0", MatrizTv(4, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 13 Or j = 14 Or j = 15) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(4, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(4, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
    Next j
    
End If
    
    If (x = 5) Then
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Batch 5"
    Call darformato
'    For m = 0 To cantelementos(5) - 1
'
'            Call insertarimagen(MatrizTv(5, 5, m), m, x)
'    Next m
    For j = 0 To 17
        
                For m = 0 To cantelementos(5)
                 If (StrComp("0", MatrizTv(5, j, m), 0) = 0 Or j = 5) Then
                      
                        
                      
                      Else
                      
                        If (j = 13 Or j = 14 Or j = 15) Then
                        
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(5, j, m) * 100 & "%"
                      
                      
                        Else
                        
                        Range("A2").Offset(m, j).Value = MatrizTv(5, j, m)
                      
                      
                      
                        End If
                      
                      End If
                        
                Next m
    Next j
    
End If
    'Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = vector(x)
            
Next x
End With
    
With ThisWorkbook
For x = 1 To contadordepaginas

If (x = 1) Then

    Sheets("Batch 1").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft

End If

If (x = 2) Then
    Sheets("Batch 2").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft


End If

If (x = 3) Then
    Sheets("Batch 3").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft




End If

If (x = 4) Then
    Sheets("Batch 4").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft



End If


If (x = 5) Then

    Sheets("Batch 5").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft



End If





Next x
End With

    
    
    
    
    Sheets("master").Visible = True
'    Sheets("batch 1").Activate
    
    

End Sub


Sub borrarmatriz(ByRef m() As String)


For Z = 1 To 5

    For x = 0 To 17
    
    
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

Sub darformato()

Application.CutCopyMode = False
    
    ActiveCell.FormulaR1C1 = "#"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "BRAND"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "MODEL"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "ENGLISH DESCRIPTION"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "SPANISH DESCRIPTION"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "PICTURE"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "QTY"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "UNIT OF MEASUREMENT"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "$U$ PRICE"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "FOB TOTAL"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "N.W UNIT"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "ORIGIN"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "NCM"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "DERECHOS"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "TE"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "IVA"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "BATCH"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "LA // LNA"
    
    Range("A1:R1").Select
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
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Columns("D:D").ColumnWidth = 61.45
    Columns("E:E").ColumnWidth = 61.45
    
    ActiveWindow.SmallScroll Down:=33
    Range("A1:R99").Select
    Range("R99").Activate
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
    

   
    
    Range("Q10").Select
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A2:R268").Select
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
    Range("A1:R1").Select
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
    Range("M6").Select
    
    Columns("A:A").ColumnWidth = 4.29           'NUMERO
    Columns("D:D").ColumnWidth = 61.45          'ESPAÑOL
    Columns("E:E").ColumnWidth = 34.82          'INGLES
    Columns("H:H").ColumnWidth = 23.74
    Columns("B:B").ColumnWidth = 10
    Columns("C:C").ColumnWidth = 10
    Columns("M:M").ColumnWidth = 16.57
    Columns("R:R").ColumnWidth = 20
    Columns("H:H").ColumnWidth = 22
    Range("A1:R99").Select
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
    Columns("D:D").ColumnWidth = 61.45
    Columns("E:E").ColumnWidth = 61.45
    Range("K9").Select
    Range("A1").Select
    Columns("F:F").ColumnWidth = 27
    ActiveWindow.View = xlPageLayoutView
    ActiveWindow.View = xlNormalView
    
     Range("A1:R549").Select
    Range("R549").Activate
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-573
    Range("P1:R268").Select
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
    Range("Q1:R501").Select
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
        Range("A1:R1").Select
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
    Range("A2:R314").Select
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
     ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Range("O1:R268").Select
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
    ' Ancho de imagen
    Columns("F:F").ColumnWidth = 20
    Columns("C:C").ColumnWidth = 20
    Columns("B:B").ColumnWidth = 16
    'Columns("O:O").ColumnWidth = 15.45
    Columns("A:A").ColumnWidth = 8
    'FORMATO PORCENTAJE
    Columns("N:P").Select
    'Selection.NumberFormat = "0%"
    
    Columns("H:H").ColumnWidth = 22
    
    ' Alto de las filas
    Rows("2:290").Select
    Selection.RowHeight = 25
    Rows("1").Select
    Selection.RowHeight = 37.5
    
    

End Sub








