Attribute VB_Name = "r_Mostrar_FILA_Insumos"
Sub Mostrar_FILA_Insumos()
    Sheets("master").Select
    
    For x = 24 To 126
            Cells(x, x).Select
            Selection.RowHeight = 60
            
    Next x

    Range("AA23").Select
End Sub

