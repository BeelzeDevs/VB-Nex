Attribute VB_Name = "Módulo2"
Sub Borrarhojas()
Attribute Borrarhojas.VB_ProcData.VB_Invoke_Func = " \n14"
'
Sheets("Master").Select
Range("Z10").Select
Selection.End(xlToRight).Select
SKU = Selection.Cells.Column


 Application.DisplayAlerts = False
For BOM = SKU - 28 To SKU
    pEST$ = "BOM - " & Cells(10, BOM).Value
    
       
    Sheets(pEST$).Delete
    
    Sheets("Master").Select
Next

Application.DisplayAlerts = True

End Sub
