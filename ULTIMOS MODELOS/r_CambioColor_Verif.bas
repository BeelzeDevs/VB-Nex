Attribute VB_Name = "r_CambioColor_Verif"
Sub CambioColor()
    Sheets("caracteristicas de los modelos").Select
    Dim rango As Range
    
    

    
    For x = 0 To 150
    Range("D3").Activate
        If (ActiveCell.Offset(x, 0).Value = 1) Then
            'ActiveCell.Offset(x, 0).Select
            ActiveCell.Offset(x, 0).EntireRow.Interior.Color = RGB(30, 200, 15)
            Range("D3").Activate
        Else
        If (ActiveCell.Offset(x, o).Value = 2) Then
        'ActiveCell.Offset(x, 0).Select
            ActiveCell.Offset(x, 0).EntireRow.Interior.Color = RGB(200, 10, 20)
        Range("D3").Activate
                
        End If
        End If
        
    Next x

    
End Sub
