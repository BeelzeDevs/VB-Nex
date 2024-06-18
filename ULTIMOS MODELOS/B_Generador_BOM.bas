Attribute VB_Name = "B_Generador_BOM"
Sub Generar_BOM()

Dim Vec(10) As String
Dim descripcion(0 To 300) As String
Dim MatrizTv(300, 0 To 8, 0 To 300) As Variant
Dim MatrizInsumos(300, 0 To 8, 0 To 99) As Variant
Dim VecModelos(300) As Variant
Dim contadordepaginas As Integer
Dim auxiliar As Integer
Dim Vbool(0 To 300) As Boolean
Dim cantelementos(300) As Integer
Dim cantelementos2 As Integer
Dim velementos(300) As Integer
Dim producto As String
Dim VecDescrip(300) As String
Dim matrizEAN(300, 0 To 1) As String
Dim rutaimg(300) As String

Application.ScreenUpdating = False
producto = "BICICLETAS"
entero = 0
auxiliar = 0
cantelementos2 = 0
contadordepaginas = 0
Call vectorencero3(VecModelos())
Call borrarmatriz(MatrizTv)
Call borrarmatriz2(MatrizInsumos)
Call vectorencero(Vec())
Call vectorenfalso(Vbool())
Call vectorencero4(cantelementos())
Call vectorencero4(velementos())
Call vectorencero4(cantelementos())
Call descripencero(VecDescrip())
Call descripenceroIMAGEN(rutaimg)
Call ponerceroEAN(matrizEAN())

' ordenar proyecto ''Agregado descrip 25/1/2023
For x = 0 To 300
    descripcion(x) = "vacio"
Next x


Sheets("Master").Select



For Z = 0 To 149
            Range("AC23").Activate
            If (StrComp("", ActiveCell.Offset(0, Z).Value, 0) = 0 Or StrComp("", ActiveCell.Offset(0, Z), 0) = 0) Then
                         Else
            VecModelos(Z) = ActiveCell.Offset(0, Z).Value
            VecDescrip(Z) = ActiveCell.Offset(-1, Z).Value
            contadordepaginas = contadordepaginas + 1
            End If
Next Z




 
For x = 0 To contadordepaginas
            Range("AC23").Activate
            ActiveCell.Offset(0, x).Activate
                cantelementos(x) = 0
                cantelementos2 = 0
               If (StrComp("", ActiveCell.Value, 0) = 0) Then
               
               Else
                 For Z = 0 To 300
                            
                            Range("AC23").Activate
                            ActiveCell.Offset(0, x).Activate
                            MatrizTv(x, 8, cantelementos(x)) = ActiveCell.Offset(0, 0).Value 'Modelo de Producto finalizado
                            
                            If (StrComp("0", ActiveCell.Offset(109 + Z, 0).Value, 0) = 0 Or StrComp("", ActiveCell.Offset(109 + Z, 0), 0) = 0) Then ' si no tiene pedido no lo guarda en la lista
                            
                            Else
                            
'
                            
                            MatrizTv(x, 0, cantelementos(x)) = cantelementos(x) + 1             'Cantidad de elementos
                            MatrizTv(x, 1, cantelementos(x)) = ""  'IMAGEN Por nro parte
                            
                            MatrizTv(x, 7, cantelementos(x)) = ActiveCell.Offset(109 + Z, -x - 6).Value  'Torque de las partes
                            MatrizTv(x, 2, cantelementos(x)) = ActiveCell.Offset(109 + Z, -x - 3).Value 'nro de parte
                            MatrizTv(x, 3, cantelementos(x)) = ActiveCell.Offset(109 + Z, -x - 4).Value 'Spanish
                            MatrizTv(x, 4, cantelementos(x)) = " " 'MatrizTv(x, 4, cantelementos(x)) = ActiveCell.Offset(109 + Z, -x - 3 - 1).Value 'Modelo de Producto finalizado ESPACIO PARA AGREGAR CUALQUIER VARIABLE QUE SE REQUIERA EN COLUMNA "DENOMINACION DE FABRICA"
                            MatrizTv(x, 5, cantelementos(x)) = ActiveCell.Offset(109 + Z, 0).Value 'Qty
                           
                            MatrizTv(x, 6, cantelementos(x)) = ActiveCell.Offset(109 + Z, -x - 3 + 156).Value 'Unit of measurement
                            
                            
                            
                            cantelementos(x) = cantelementos(x) + 1
                            velementos(x) = cantelementos(x)
                            End If
                            Next Z
            End If
            
                            Range("AC23").Activate

            
Next x
       
       
      
For x = 0 To contadordepaginas - 1
Range("AC23").Activate
ActiveCell.Offset(0, x).Activate



 If (StrComp("", ActiveCell.Value, 0) = 0 Or StrComp("0", ActiveCell.Value, 0) = 0) Then
               
               Else
                
               
                cantelementos2 = velementos(x)
                For j = 0 To 99
                            Range("AC23").Activate
                            ActiveCell.Offset(0, x).Activate
                            MatrizInsumos(x, 8, j) = ActiveCell.Offset(0, 0).Value 'Modelo de Producto finalizado
                       
                       If (StrComp("", ActiveCell.Offset(j + 3, 0).Value, 0) = 0 Or StrComp("0", ActiveCell.Offset(j + 3, 0).Value, 0) = 0) Then

                       Else
                       
                       MatrizInsumos(x, 0, j) = cantelementos2 + 1 ' cant elementos
                       
                       MatrizInsumos(x, 1, j) = ""  'IMAGEN Por nro parte
                       'MatrizInsumos(x, 1, j) = ActiveCell.Offset(j + 3, -3 - x).Value ' Nro de imagen
                       
                       MatrizInsumos(x, 2, j) = ActiveCell.Offset(j + 3, -3 - x).Value 'sku nro de parte
                       MatrizInsumos(x, 3, j) = ActiveCell.Offset(j + 3, -4 - x).Value ' denominacion espa
                       MatrizInsumos(x, 4, j) = " " 'MatrizInsumos(x, 4, j) = ActiveCell.Offset(j + 3, -4 - x).Value ' denominacion espa 2, ESPACIO LIBRE PARA AGREGAR VARIABLE
                       MatrizInsumos(x, 5, j) = ActiveCell.Offset(j + 3, 0).Value 'Qty
                       MatrizInsumos(x, 6, j) = ActiveCell.Offset(j + 3, -x + 153).Value 'UNIT OF MEASUREMENT
                       'MatrizInsumos(x, 7, j) = ActiveCell.Offset(j + 3, -x + 153).Value 'UNIT OF MEASUREMENT
                                              
                       velementos(x) = velementos(x) + 1
                       cantelementos2 = cantelementos2 + 1
                       End If
                Next j
            End If
    
Next x
                                'BUSQUEDA DE SKU Y EAN EN LA CODIF DE BICICLETAS
Call abrolibro_SKU(producto)

For pos = 0 To contadordepaginas - 1
ThisWorkbook.Activate
Call BusquedaSKU_EAN(VecModelos(pos), matrizEAN, pos, producto, descripcion) ''Agregado descrip 25/1/2023
     
Next pos
Call cierrolibro_SKU(producto)
 
 
        
                     '' BUSQUEDA de nombre para la ruta de imagenes
For x = 0 To contadordepaginas
    For j = 0 To cantelementos(x)
    If (StrComp(VecModelos(x), MatrizTv(x, 8, j), 0) = 0 Or StrComp(VecModelos(x), MatrizInsumos(x, 8, j), 0) = 0) Then

 Call Busqueda_nombremod_img(VecModelos(x), rutaimg(), x, j)
    End If
    Next j
Next x

'INSERTO LOS PRODUCTOS EN MI BOM
    
With ThisWorkbook
       
For x = 0 To contadordepaginas - 1
    
    
    Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "BOM" & " " & MatrizTv(x, 8, 0)
    
    'version más lenta -> Call Busqueda_SKU_EAN(VecModelos(x), matrizEAN(), x)
    
    'pegar lo datos correspondientes y darle formato
    
    Call darformato(VecModelos(x), producto, VecDescrip(x), descripcion(x))
    Call Formatode(matrizEAN(), x, VecModelos(x))
    'Call autofit
    
    'Insertar imagenes
    For m = 0 To cantelementos(x) - 1
            If (StrComp("BICICLETAS", producto, 0) = 0) Then
            Call insertarimagenPARTES_BICI(MatrizTv(x, 8, 0), m, x, MatrizTv(x, 2, m), rutaimg(x), producto)
            Call insertar_imagen_ensambleproducto(MatrizTv(x, 8, 0), rutaimg(x), producto)
            End If
            If (StrComp("VACUUM CLEANER ROBOT", producto, 0) = 0) Then
            Call insertarimagenPARTES_VACUUM(MatrizTv(x, 8, 0), m, x, MatrizTv(x, 2, m), rutaimg(x), producto)
            End If
            If (StrComp("KETTLE", producto, 0) = 0) Then
            Call insertarimagenPARTES_KETTLE(MatrizTv(x, 8, 0), m, x, MatrizTv(x, 2, m), rutaimg(x), producto)
            End If
    Next m
    
    
    For j = 0 To 7
        
                For m = 0 To cantelementos(x) - 1
                      If (StrComp("", MatrizTv(x, j, m), 0) = 0 Or StrComp("0", MatrizTv(x, j, m), 0) = 0) Then
                      
                        
                      
                      Else
    
                        
                      Range("A9").Offset(m, j).Value = MatrizTv(x, j, m)
                      
                      
                      
                      
                      End If
                        
                Next m
    Next j
    
    'INSERTO LOS INSUMOS EN MI BOM
    
    Call lineadedivision(cantelementos(x))
                
                'Insertar imagenes
                For m = 0 To velementos(x) - cantelementos(x)

                        Call insertarimagenINSUMOS(MatrizInsumos(x, 8, j), m + cantelementos(x) + 1, x, MatrizInsumos(x, 2, m), rutaimg(x), producto)
                Next m
                
                
                For j = 0 To 7
                    
                            For m = 0 To velementos(x) - cantelementos(x)
                                      If (StrComp("", MatrizInsumos(x, j, m), 0) = 0 Or StrComp("0", MatrizInsumos(x, j, m), 0) = 0) Then



                                      Else
                    
                                        
                                        Range("A10").Offset(m + cantelementos(x), j).Value = MatrizInsumos(x, j, m)
                                      
                                      
                                      
                                      
                                      End If
                                    
                            Next m
                Next j
              
    
    
    
    

    'Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = vector(x)
                    
        Next x
        End With
    
    

        
    
    Application.ScreenUpdating = True
    Sheets("master").Visible = True
    'Sheets(6).Select
    'Range("A1").Select
    

End Sub

' Producto en la parte de imagenes del codigo y en la apertura y cierre del sku
Sub borrarmatriz(ByRef m() As Variant)


For Z = 0 To 300

    For x = 0 To 8
    
    
                For j = 0 To 300
                    
                        m(Z, x, j) = 0
                
                
                Next j
                
    
    Next x
    
Next Z







End Sub
Sub borrarmatriz2(ByRef m() As Variant)


For Z = 0 To 300

    For x = 0 To 8
    
    
                For j = 0 To 99
                    
                        m(Z, x, j) = 0
                
                
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
    
    For x = 1 To 1
                
                vector(x) = 0
        Next x

End Sub

Sub vectorencero3(ByRef vector() As Variant)
    For x = 0 To 300
                
                vector(x) = "0"
        Next x

End Sub
Sub vectorenfalso(ByRef vexiste() As Boolean)
    
    For x = 0 To 300
                
                vexiste(x) = False
        Next x

End Sub
Sub vectorencero4(ByRef vexiste() As Integer)

    For x = 0 To 300
        
                vexiste(x) = 0
    
    Next x

End Sub
Sub descripencero(ByRef vector() As String)
    For x = 0 To 300
            
            vector(x) = "0"
        
    Next x

End Sub
Sub descripenceroIMAGEN(ByRef m() As String)

For x = 0 To 300
 '   For j = 0 To 300
            m(x) = ""
   ' Next j
Next x


End Sub

Sub darformato(ByVal modelo As String, ByVal producto As String, ByVal descripcion As String, ByVal descripcion2 As String)

Application.CutCopyMode = False
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "#"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "PICTURE" '"SPANISH DESCRIPTION"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "Nro Parte"           '"PART NUMBER"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "Denominación"          '"UDM"
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "Denominación de fábrica"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "QTY"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "UDM"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "Torque"
    Range("A7:H7").Select
    
                        
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
    
    Range("A7:H400").Select
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
    
    Range("A8:H400").Select
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
    
    Range("A7:H7").Select
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
    '                  // FORMATO 1 de BOM. Formato 2 en funcion Autofit
     
     
    Columns("A:A").ColumnWidth = 10           'NUMERO
    Columns("D:D").ColumnWidth = 65          'ESPAÑOL
    Columns("G:G").ColumnWidth = 17
    Columns("E:E").ColumnWidth = 24          'INGLES
    Columns("B:B").ColumnWidth = 25         'foto
    Columns("G:G").ColumnWidth = 17             'unidad de medida
    Columns("F:F").ColumnWidth = 17          'QTY
    Columns("C:C").ColumnWidth = 25
    Columns("H:H").ColumnWidth = 25

    
   
    Range("A7").Select
    
    

'PORCENTAJES IVA TE DERECHO

        Range("A7:H7").Select
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
    
    Range("A8:H300").Select
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
    
'    Range("O1:R268").Select                                    NIDEAH
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
'    With Selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .Weight = xlThin
'    End With
    
    ' Alto de las filas
    Rows("8:300").Select
    Selection.RowHeight = 50
    Rows("7").Select
    Selection.RowHeight = 37.5
    
    
    ' formato de color y resaltar la primera fila, Titulos
    
    Range("A7:H7").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -6279056
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -6279056
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -6279056
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -6279056
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 10498160
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0.5)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 10498160
        .TintAndShade = 0
    End With
    
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 10498160
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0.5)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 10498160
        .TintAndShade = 0
    End With
    
    
    Range("C1:D1").Merge
    Range("C2:D2").Merge
    Range("C3:D3").Merge
    Range("C4:D4").Merge
    Range("C5:D5").Merge
    
    Range("A1").RowHeight = 34.5
    Rows("8:308").RowHeight = 50
    
    Range("C1").Select
    With Selection.Font
        .Name = "Rockwell"
        .Size = 13
        .Bold = True
    End With
 
    
    Range("C1:D5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
'  Range("C1").Value = producto & " " & modelo & " - " & descripcion
    
    Range("C1").Value = descripcion2    'DSCRIPCION 2
    
    
'    Range("F1:F5").Select
'    With Selection
'        .HorizontalAlignment = xlBottom
'        .VerticalAlignment = xlBottom
'
'    End With
    
    ' otro formato de tablas +
'     Range("A1:G549").Select
'    With Selection.Interior
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorDark1
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
   
    
    

End Sub

Sub Formatode(ByRef MatEAN() As String, ByVal pos As Integer, ByVal codigo As String)
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Producto:"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "EAN:"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "SKU:"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = " " '"Código Actual:"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Código Actual:"
        
    Range("B1:B5").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Confeccionó:"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Fecha:"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Versión"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "Revisó:"
    
    Range("E1:E5").Select
 
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Ingeniería"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "Silva K."
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "v1"
    Range("F5").Select
    
    
    
     Range("F1:F6").Select
    With Selection
        .VerticalAlignment = xlBottom
        .HorizontalAlignment = xlRight
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    
    'degrade en las columnas de las puntas
   
    Range("C2:D2").Value = MatEAN(pos, 1)
    Range("C3:D3").Value = MatEAN(pos, 0)
    
    Range("C2:D2").Select
    Selection.NumberFormat = "0"
    
    'Range("C4:D4").Value = codigo
    Range("C5:D5").Value = codigo
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.RowHeight = 250.5
    Range("C1:E1").Select
    Selection.Merge
    
    Range("A1:A7").Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 180
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -5.09659108249153E-02
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 10498160
        .TintAndShade = 0
    End With
    Range("H1:H7").Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 10498160
        .TintAndShade = 0
    End With
    
    Range("A7:H7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -6279056
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
End Sub

Sub lineadedivision(ByVal posicion As Integer)

    Range("A9:H9").Offset(posicion, 0).Select
    
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
    Selection.Merge
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 10498160
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0.5)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 10498160
        .TintAndShade = 0
    End With
  Range("A9:H9").Offset(posicion, 0).Value = " COMPRAS LOCALES "

End Sub

Sub ponerceroEAN(ByRef m() As String)

For x = 0 To 300
                For j = 0 To 1
                m(x, j) = " "
                
                Next j

Next x
End Sub

Sub autofit()
    Columns("E:E").EntireColumn.autofit
    Columns("D:D").EntireColumn.autofit
    Columns("F:F").EntireColumn.autofit
    Columns("G:G").EntireColumn.autofit
    Columns("A:A").EntireColumn.autofit
    Columns("B:B").EntireColumn.autofit
    Columns("C:C").EntireColumn.autofit
    Columns("D:D").EntireColumn.autofit
    Columns("H:H").EntireColumn.autofit
    
End Sub

