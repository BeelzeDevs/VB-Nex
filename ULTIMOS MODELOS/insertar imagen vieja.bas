Attribute VB_Name = "Módulo2"
Option Explicit
Sub insertarimagen(nombre As String, ByVal posicion As Integer, ByVal worsh As Integer)

Dim Ws As Worksheet
Dim Pic As Picture
Dim Rng As Range
Dim Ruta As String
        'Condicionales para ver cual es el worksheet
    If (worsh = 1) Then
    Set Ws = Worksheets("Batch 1")
    End If
    If (worsh = 2) Then
    Set Ws = Worksheets("Batch 2")
    End If
    If (worsh = 3) Then
    Set Ws = Worksheets("Batch 3")
    End If
    If (worsh = 4) Then
    Set Ws = Worksheets("Batch 4")
    End If
    If (worsh = 5) Then
    Set Ws = Worksheets("Batch 5")
    End If
    
    Ruta = ThisWorkbook.Path & "\Imagenes\"
    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
    Set Rng = Nothing
    Set Pic = Nothing
    Set Rng = Ws.Range("F2").Offset(posicion, 0)
    
    'Set Pic = Ws.Pictures.Insert(Ruta & nombre & ".png")
    Set Pic = Ws.Pictures.Insert(Ruta & nombre & ".png")
    With Pic
        .ShapeRange.LockAspectRatio = msoFalse
        .Left = Rng.Left                'esquina de la imagen igual al rango
        .Top = Rng.Top                  'esquina igual al superior
        .Width = Rng.Width
        .Height = Rng.Height
        .Placement = xlMoveAndSize
                                            'xlMoveAndSize Mover y Cambiar el tamaño con celdas
                                            'xlMove Mover pero No cambiar tamaño con celdas
                                            'xlFreeFloating No mover Ni cambiar tamaño con celdas
    
    End With
    
End Sub
