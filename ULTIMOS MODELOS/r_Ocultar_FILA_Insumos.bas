Attribute VB_Name = "r_Ocultar_FILA_Insumos"
Sub ocultar_FILA_Insumos()
    Sheets("master").Select
    
    For x = 24 To 126
            Cells(x, x).Select
            Selection.RowHeight = 1
            
    Next x

    
    Range("AA129").Select
End Sub
