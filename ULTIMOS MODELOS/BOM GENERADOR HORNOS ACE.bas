Attribute VB_Name = "Módulo1"
Sub BOMgenerator()
Attribute BOMgenerator.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BOMgenerator Macro
'
Sheets("Master").Select

Range("Z10").Select
Selection.End(xlToRight).Select
SKU = Selection.Cells.Column
    
Range("Z10").Select
    Selection.End(xlDown).Select
Parts = Selection.Cells.Row

    For BOM = SKU - 28 To SKU
    pEST$ = "BOM - " & Cells(10, BOM).Value
    Sheets.Add
    ActiveSheet.Name = pEST$
    
    Sheets("Carátula").Select
    Range(Cells(1, 2), Cells(8, 9)).Select
    Selection.Copy
    Worksheets(pEST$).Activate
    
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
   
    
    Worksheets("Master").Activate
    
    
    
    Range(Cells(11, 13), Cells(11, 19)).Select
    Selection.Copy
    Worksheets(pEST$).Activate
    
    ActiveCell.Offset(2, 4).Select
    Selection.Value = pEST$
    ActiveCell.Offset(7, -4).Select
    
    
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
   ActiveCell.Offset(1, 0).Select
    
    
    Worksheets("Master").Activate
        
    For Item = 12 To Parts
    cantidad = Cells(Item, BOM).Value
    
    If cantidad > 0 Then
    
    Range(Cells(Item, 13), Cells(Item, 19)).Select
    Selection.Copy
    Worksheets(pEST$).Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    ActiveCell.Offset(0, 4).Select
    Selection.Value = cantidad
    ActiveCell.Offset(1, -4).Select
    Sheets("Master").Select
    
    End If
    
    Next
    
    Next
    End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
End Sub
