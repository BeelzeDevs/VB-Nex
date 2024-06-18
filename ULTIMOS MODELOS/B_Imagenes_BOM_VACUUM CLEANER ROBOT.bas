Attribute VB_Name = "B_Imagenes_BOM"
Option Explicit
Sub insertarimagenPARTES(ByVal nombre As String, ByVal posicion As Integer, ByVal worsh As Integer, ByVal nroparte As String, ByVal rutamodelo As String, ByVal product As String)

Dim ws As Worksheet
'Dim Pic As Picture
Dim rng As Range
Dim ruta As String
Dim mayusc As String
Dim minusc As String


mayusc = UCase(product)
minusc = LCase(product)

        'Condicionales para ver cual es el worksheet
    
    Set ws = Worksheets("BOM" & " " & nombre)
    
    'Batch
'    If (worsh = 2) Then
'    Set Ws = Worksheets("Batch 2")
'    End If
'    If (worsh = 3) Then
'    Set Ws = Worksheets("Batch 3")
'    End If
'    If (worsh = 4) Then
'    Set Ws = Worksheets("Batch 4")
'    End If
'    If (worsh = 5) Then
'    Set Ws = Worksheets("Batch 5")
'    End If
    
    
    
          'BUSQUEDA POR RUTA DESD DOND SE ENCUENTRA EL ARCHIVO
'
'    ruta = ThisWorkbook.Path & "\Imagenes\"
'    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
'    Set Rng = Nothing
'    Set Pic = Nothing
'    Set Rng = ws.Range("B8").Offset(posicion, 0)
'
'    Set Pic = ws.Pictures.Insert(ruta & nombre & ".png")
'    Set Pic = ws.Pictures.Insert(ruta & nroparte & ".png")
'    With Pic
'        .ShapeRange.LockAspectRatio = msoFalse
'        .Left = Rng.Left                'esquina de la imagen igual al rango
'        .Top = Rng.Top                  'esquina igual al superior
'        .Width = Rng.Width
'        .Height = Rng.Height
'        .Placement = xlMoveAndSize
''                                            xlMoveAndSize Mover y Cambiar el tamaño con celdas
''                                            xlMove Mover pero No cambiar tamaño con celdas
''                                            xlFreeFloating No mover Ni cambiar tamaño con celdas
'
'    End With
'
                '    BUSQUEDA POR CARPETA Y MODELOS
                
    Dim app As Application
    Dim shp As Object
    Dim usuario As String
    Dim ObjWshNw As Object
    Set ObjWshNw = CreateObject("WScript.Network")
    ws.Activate
    usuario = ObjWshNw.UserName
    ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\" & mayusc & "\INFORMACION DEL PRODUCTO\INFORMACION NUESTRA\JC\" & rutamodelo & "\" & nroparte & ".png"
    
    Application.ScreenUpdating = False
    
    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
    Set rng = Nothing
'    Set Pic = Nothing
    Set shp = Nothing
    Set rng = ws.Range("B8").Offset(posicion, 0)
    
    ' Set Pic = ws.Pictures.Insert(ruta & nroparte & ".png")
   
    'Set Pic = Ws.Pictures.Insert(Ruta & nombre & ".png")
    With rng.Cells.Worksheet.Shapes.AddPicture(Filename:=ruta, LinkToFile:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, Width:=-1, Height:=-1)
        .LockAspectRatio = msoFalse
        .Left = rng.Left
        .Top = rng.Top
        .Width = rng.Width
        .Height = rng.Height
        .Placement = xlMoveAndSize
    End With
    

'    With Pic
'        .ShapeRange.LockAspectRatio = msoFalse
'        .Left = rng.Left                'esquina de la imagen igual al rango
'        .Top = rng.Top                  'esquina igual al superior
'        .Width = rng.Width
'        .Height = rng.Height
'        .Placement = xlMoveAndSize
'                                           'xlMoveAndSize Mover y Cambiar el tamaño con celdas
'                                            'xlMove Mover pero No cambiar tamaño con celdas
'                                            'xlFreeFloating No mover Ni cambiar tamaño con celdas
'
'    End With

    Application.ScreenUpdating = True
End Sub

Sub insertarimagenINSUMOS(ByVal nombre As String, ByVal posicion As Integer, ByVal worsh As Integer, ByVal nroparte As String, ByVal rutamodelo As String, ByVal product As String)

Dim ws As Worksheet
'Dim Pic As Picture
Dim rng As Range
Dim ruta As String
Set ws = Worksheets("BOM" & " " & nombre)
    Dim app As Application
    Dim shp As Object
    Dim usuario As String
    Dim ObjWshNw As Object
    Set ObjWshNw = CreateObject("WScript.Network")
    Dim mayusc As String
    mayusc = UCase(product)
    
    
    'OPCIONES DE RUTA
    'ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\BICICLETAS\INFORMACION DEL PRODUCTO\INFORMACION NUESTRA\JC\" & rutamodelo & "\" & nroparte & ".png"
    'ruta = ThisWorkbook.Path & "\Imagenes\"
    
    usuario = ObjWshNw.UserName
    ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\" & mayusc & "\INFORMACION DEL PRODUCTO\INFORMACION NUESTRA\JC\" & "INSUMOS LOCALES" & "\" & nroparte & ".png"
    
    Application.ScreenUpdating = False
    
    
        'Condicionales para ver cual es el worksheet
    Set ws = Worksheets("BOM" & " " & nombre)
    
'    If (worsh = 2) Then
'    Set Ws = Worksheets("Batch 2")
'    End If
'    If (worsh = 3) Then
'    Set Ws = Worksheets("Batch 3")
'    End If
'    If (worsh = 4) Then
'    Set Ws = Worksheets("Batch 4")
'    End If
'    If (worsh = 5) Then
'    Set Ws = Worksheets("Batch 5")
'    End If
    
    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
    Set rng = Nothing
    Set shp = Nothing
    Set rng = ws.Range("B9").Offset(posicion, 0)
    

    
    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
    Set rng = Nothing
'    Set Pic = Nothing
    Set shp = Nothing
    Set rng = ws.Range("B8").Offset(posicion, 0)
    
    ' Set Pic = ws.Pictures.Insert(ruta & nroparte & ".png")
   
    With rng.Cells.Worksheet.Shapes.AddPicture(Filename:=ruta, LinkToFile:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, Width:=-1, Height:=-1)
        .LockAspectRatio = msoFalse
        .Left = rng.Left
        .Top = rng.Top
        .Width = rng.Width
        .Height = rng.Height
        .Placement = xlMoveAndSize
    End With
    
    
    
    
'    With Pic
'        .ShapeRange.LockAspectRatio = msoFalse
'        .Left = rng.Left                'esquina de la imagen igual al rango
'        .Top = rng.Top                  'esquina igual al superior
'        .Width = rng.Width
'        .Height = rng.Height
'        .Placement = xlMoveAndSize
'        .ShapeRange.savewithdocumento = True                                 'xlMoveAndSize Mover y Cambiar el tamaño con celdas
'                                            'xlMove Mover pero No cambiar tamaño con celdas
'                                            'xlFreeFloating No mover Ni cambiar tamaño con celdas
'
'    End With

Application.ScreenUpdating = True
End Sub

Sub insertarimagenPARTESVACUUM(ByVal nombre As String, ByVal posicion As Integer, ByVal worsh As Integer, ByVal nroparte As String, ByVal rutamodelo As String, ByVal product As String)

Dim ws As Worksheet
'Dim Pic As Picture
Dim rng As Range
Dim ruta As String
Dim mayusc As String
Dim minusc As String


mayusc = UCase(product)
minusc = LCase(product)

    
    Set ws = Worksheets("BOM" & " " & nombre)
    

                
    Dim app As Application
    Dim shp As Object
    Dim usuario As String
    Dim ObjWshNw As Object
    Set ObjWshNw = CreateObject("WScript.Network")
    Dim ruta2 As String
    ws.Activate
    usuario = ObjWshNw.UserName
    ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\" & mayusc & "\INFORMACION DEL PRODUCTO\INFORMACIÓN NUESTRA\SENA\" & rutamodelo & "\" & nroparte & ".png"
    Application.ScreenUpdating = False
    
    On Error Resume Next ' Luego de esta linea si hay error, no lo genera
    Set rng = Nothing
'    Set Pic = Nothing
    Set shp = Nothing
    Set rng = ws.Range("B8").Offset(posicion, 0)
    
    ' Set Pic = ws.Pictures.Insert(ruta & nroparte & ".png")
   
    'Set Pic = Ws.Pictures.Insert(Ruta & nombre & ".png")
    With rng.Cells.Worksheet.Shapes.AddPicture(Filename:=ruta, LinkToFile:=msoFalse, savewithdocument:=msoCTrue, Left:=0, Top:=0, Width:=-1, Height:=-1)
        .LockAspectRatio = msoFalse
        .Left = rng.Left
        .Top = rng.Top
        .Width = rng.Width
        .Height = rng.Height
        .Placement = xlMoveAndSize
    End With
    

'    With Pic
'        .ShapeRange.LockAspectRatio = msoFalse
'        .Left = rng.Left                'esquina de la imagen igual al rango
'        .Top = rng.Top                  'esquina igual al superior
'        .Width = rng.Width
'        .Height = rng.Height
'        .Placement = xlMoveAndSize
'                                           'xlMoveAndSize Mover y Cambiar el tamaño con celdas
'                                            'xlMove Mover pero No cambiar tamaño con celdas
'                                            'xlFreeFloating No mover Ni cambiar tamaño con celdas
'
'    End With

    Application.ScreenUpdating = True
End Sub



