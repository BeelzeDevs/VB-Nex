Attribute VB_Name = "Shipping_Mark_Borrar"
Sub shipping_Mark_Borrar()
    Sheets("Shipping mark").Select
    
    With Range("A1:J9999")
    .Delete
    End With
    Range("A1:J9999").Interior.Color = RGB(255, 255, 255)
    '.Interior.Color = RGB(255, 255, 255)
    



End Sub
