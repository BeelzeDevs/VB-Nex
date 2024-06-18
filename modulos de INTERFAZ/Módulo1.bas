Attribute VB_Name = "Módulo1"
    
Sub Hide_Show_SIMI()

Sheets("MACROS").Shapes("OP_BOM").Visible = msoFalse
Sheets("MACROS").Shapes("OP_SHIP").Visible = msoFalse

If Sheets("MACROS").Shapes("OP_SIMI").Visible = True Then
Sheets("MACROS").Shapes("OP_SIMI").Visible = msoFalse
Else
    Sheets("MACROS").Shapes("OP_SIMI").Visible = msoCTrue
End If


End Sub
   

Sub Hide_Show_BOM()
Sheets("MACROS").Shapes("OP_SIMI").Visible = msoFalse

Sheets("MACROS").Shapes("OP_SHIP").Visible = msoFalse

If Sheets("MACROS").Shapes("OP_BOM").Visible = True Then
Sheets("MACROS").Shapes("OP_BOM").Visible = msoFalse
Else
    Sheets("MACROS").Shapes("OP_BOM").Visible = msoCTrue
End If


End Sub

Sub Hide_Show_SHIP()
Sheets("MACROS").Shapes("OP_SIMI").Visible = msoFalse

Sheets("MACROS").Shapes("OP_BOM").Visible = msoFalse

If Sheets("MACROS").Shapes("OP_SHIP").Visible = True Then
Sheets("MACROS").Shapes("OP_SHIP").Visible = msoFalse
Else
    Sheets("MACROS").Shapes("OP_SHIP").Visible = msoCTrue
End If


End Sub
   

'' dentro del this workbook
'Private Sub Workbook_Open()
'
'Sheets("MACROS").Shapes("OP_SIMI").Visible = msoFalse
'
'Sheets("MACROS").Shapes("OP_BOM").Visible = msoFalse
'Sheets("MACROS").Shapes("OP_SHIP").Visible = msoFalse
'
'End Sub

