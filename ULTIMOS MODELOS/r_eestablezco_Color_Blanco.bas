Attribute VB_Name = "r_eestablezco_Color_Blanco"
Sub reestablecer1()
    Sheets("caracteristicas de los modelos").Select
    
    Range("D3").Activate
    For x = 0 To 150
            
        ActiveCell.Offset(x, 0).EntireRow.Interior.Color = RGB(255, 255, 255)
            
    
    Next x


End Sub
