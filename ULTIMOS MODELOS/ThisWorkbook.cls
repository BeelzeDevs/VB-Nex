VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
ThisWorkbook.Unprotect "nexina13"

End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
Dim boleano As Boolean

boleano = False
If (Sh.Name = "Master") Then
ThisWorkbook.Protect "nexina13", structure:=True
boleano = True
End If

If (Sh.Name = "Shipping Mark") Then
ThisWorkbook.Protect "nexina13", structure:=True
boleano = True
End If

If (Sh.Name = "Caracteristicas de los modelos") Then
ThisWorkbook.Protect "nexina13", structure:=True
boleano = True
End If

If (Sh.Name = "Macro generador") Then
ThisWorkbook.Protect "nexina13", structure:=True
boleano = True
End If

If (boleano = True) Then
MsgBox "Hoja protegida por el administrador, pongase en contacto con Ignacio", vbDefaultButton1
End If

End Sub


