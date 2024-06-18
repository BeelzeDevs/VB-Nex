Attribute VB_Name = "Busqueda_SKU_EAN"
Sub abrolibro_SKU(ByVal producto As String)
 Dim ws As Workbook
 Dim ruta As String
 Dim app As Application
 Dim usuario As String
 Dim ObjWshNw As Object
 Set ObjWshNw = CreateObject("WScript.Network")
 Dim mayus As String
 Dim minusc As String
 mayusc = UCase(producto)
 minusc = LCase(producto)
 'StrConv primera letra mayus, luego minusc
 
 usuario = ObjWshNw.UserName
 
 ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\" & mayusc & "\CODIFICACION DE PRODUCTO TERMINADO\"
 
 
 Set ws = Workbooks.Open(Filename:=ruta & "Codificacion de " & minusc & ".xlsx")
 ws.Activate



End Sub

Sub cierrolibro_SKU(ByVal producto As String)
 Dim ws As Workbook
 Dim ruta As String
 Dim app As Application
 Dim usuario As String
 Dim ObjWshNw As Object
 Set ObjWshNw = CreateObject("WScript.Network")
 Dim mayus As String
 Dim minusc As String
 mayusc = UCase(producto)
 minusc = LCase(producto)
 'StrConv primera letra mayus, luego minusc
 
 usuario = ObjWshNw.UserName
 
 
 Workbooks("Codificacion de " & minusc & ".xlsx").Close SaveChanges:=False


End Sub
Sub BusquedaSKU_EAN(ByVal m As String, ByRef v() As String, ByVal pos As Integer, ByVal producto As String, ByRef descripcion() As String) '
 Dim ws As Workbook
 Dim ruta As String
 Dim app As Application
 Dim usuario As String
 Dim ObjWshNw As Object
 Set ObjWshNw = CreateObject("WScript.Network")
 Dim mayus As String
 Dim minusc As String
 mayusc = UCase(producto)
 minusc = LCase(producto)
 'StrConv primera letra mayus, luego minusc
 
 usuario = ObjWshNw.UserName
 
 ruta = "C:\Users\" & usuario & "\Dropbox\INGENIERIA\" & mayusc & "\CODIFICACION DE PRODUCTO TERMINADO\" & "Codificacion de " & minusc & ".xlsx"
 
 
 Set ws = Application.Workbooks("Codificacion de " & minusc & ".xlsx")
 ws.Activate
 
 Range("V13").Activate  'BICI
'Range("U13").Activate 'KETTLE'
 For x = 0 To 170 '1250

    If (StrComp(m, ActiveCell.Offset(x, 0).Value, 0) = 0) Then
    v(pos, 0) = ActiveCell.Offset(x, 1).Value '' SKU
    v(pos, 1) = ActiveCell.Offset(x, 5).Value '' EAN
    descripcion(pos) = ActiveCell.Offset(x, 6)
    
    x = 170
    
    End If


 Next x

   

End Sub
Sub Busqueda_nombremod_img(ByVal vmod As String, ByRef ruta() As String, ByVal x As Integer, ByVal j As Integer) '


        
    
                    'OVANTA
                    If (StrComp(vmod, "ZI-MBOB-0021", 0) = 0 Or StrComp(vmod, "ZI-MBOB-0016", 0) = 0 Or StrComp(vmod, "ZI-MBOB-0018", 0) = 0 Or StrComp(vmod, "ZI-MBOB-0020", 0) = 0 Or StrComp(vmod, "ZI-MBOV-0015", 0) = 0 Or StrComp(vmod, "ZI-MBOV-0016", 0) = 0 Or StrComp(vmod, "ZI-MBOV-0018", 0) = 0 Or StrComp(vmod, "ZI-MBOV-0020", 0) = 0 Or StrComp(vmod, "ZI-MBOV-0021", 0) = 0 Or StrComp(vmod, "ZI-MBOA-0016", 0) = 0 Or StrComp(vmod, "ZI-MBOA-0018", 0) = 0 Or StrComp(vmod, "ZI-MBOA-0020", 0) = 0 Or StrComp(vmod, "ZI-MBOA-0021", 0) = 0 Or StrComp(vmod, "ZI-MBOG-0015", 0) = 0 Or StrComp(vmod, "ZI-MBOG-0016", 0) = 0 Or StrComp(vmod, "ZI-MBOG-0018", 0) = 0 Or StrComp(vmod, "ZI-MBOG-0020", 0) = 0 Or StrComp(vmod, "ZI-MBOG-0021", 0) = 0) Then
                        ruta(x) = "ZION OVANTA"
                    End If
                    'BREVA
                    If (StrComp(vmod, "ZI-MBBR-0015", 0) = 0 Or StrComp(vmod, "ZI-MBBR-0016", 0) = 0 Or StrComp(vmod, "ZI-MBBR-0018", 0) = 0 Or StrComp(vmod, "ZI-MBBR-0020", 0) = 0 Or StrComp(vmod, "ZI-MBBV-0015", 0) = 0 Or StrComp(vmod, "ZI-MBBV-0016", 0) = 0 Or StrComp(vmod, "ZI-MBBV-0018", 0) = 0 Or StrComp(vmod, "ZI-MBBV-0020", 0) = 0 Or StrComp(vmod, "ZI-MBBF-0015", 0) = 0 Or StrComp(vmod, "ZI-MBBF-0016", 0) = 0 Or StrComp(vmod, "ZI-MBBF-0018", 0) = 0 Or StrComp(vmod, "ZI-MBBF-0020", 0) = 0) Then
                        ruta(x) = "ZION BREVA"
                    End If
                    'ASPRO
                    If (StrComp(vmod, "ZI-MBAA-0016", 0) = 0 Or StrComp(vmod, "ZI-MBAA-0018", 0) = 0 Or StrComp(vmod, "ZI-MBAA-0020", 0) = 0 Or StrComp(vmod, "ZI-MBAA-0021", 0) = 0 Or StrComp(vmod, "ZI-MBAG-0015", 0) = 0 Or StrComp(vmod, "ZI-MBAG-0016", 0) = 0 Or StrComp(vmod, "ZI-MBAG-0018", 0) = 0 Or StrComp(vmod, "ZI-MBAG-0020", 0) = 0) Then
                        ruta(x) = "ZION ASPRO"
                    End If
                    'STRIX
                    If (StrComp(vmod, "ZI-MBSC-0016", 0) = 0 Or StrComp(vmod, "ZI-MBSC-0018", 0) = 0 Or StrComp(vmod, "ZI-MBSC-0020", 0) = 0 Or StrComp(vmod, "ZI-MBSA-0016", 0) = 0 Or StrComp(vmod, "ZI-MBSA-0018", 0) = 0 Or StrComp(vmod, "ZI-MBSA-0020", 0) = 0 Or StrComp(vmod, "ZI-MBSA-0021", 0) = 0 Or StrComp(vmod, "ZI-MBSR-0015", 0) = 0 Or StrComp(vmod, "ZI-MBSR-0016", 0) = 0 Or StrComp(vmod, "ZI-MBSR-0018", 0) = 0 Or StrComp(vmod, "ZI-MBSR-0020", 0) = 0 Or StrComp(vmod, "ZI-MBSR-0021", 0) = 0) Then
                        ruta(x) = "ZION STRIX"
                    End If
                    'AVRA
                    If (StrComp(vmod, "ZI-GBAC-0017", 0) = 0 Or StrComp(vmod, "ZI-GBAC-0019", 0) = 0 Or StrComp(vmod, "ZI-GBAC-0021", 0) = 0) Then
                        ruta(x) = "ZION AVRA"
                    End If
                    
                    'DIABLO
                    If (StrComp(vmod, "ZI-MBDG-0017", 0) = 0 Or StrComp(vmod, "ZI-MBDG-0019", 0) = 0 Or StrComp(vmod, "ZI-MBDG-0015", 0) = 0) Then
                        ruta(x) = "ZION DIABLO"
                    End If
                    
                    'PATAGONIA
                    If (StrComp(vmod, "XI-BMPA18-000", 0) = 0 Or StrComp(vmod, "XI-BMPA20-000", 0 Or StrComp(vmod, "ZI-MBPN-0016", 0) = 0 Or StrComp(vmod, "ZI-MBPN-0018", 0) = 0 Or StrComp(vmod, "ZI-MBPN-0020", 0) = 0 Or StrComp(vmod, "ZI-MBPG-0015", 0) = 0 Or StrComp(vmod, "ZI-MBPG-0016", 0) = 0 Or StrComp(vmod, "ZI-MBPG-0018", 0) = 0) = 0 Or StrComp(vmod, "ZI-MBPG-0020", 0) = 0) Then
                        ruta(x) = "ZION PATAGONIA"
                    End If
                    'PAMPA
                    If (StrComp(vmod, "ZI-MBPM-0216", 0) = 0 Or StrComp(vmod, "ZI-MBPM-0218", 0) = 0 Or StrComp(vmod, "ZI-MBPM-0220", 0) = 0) Then
                        ruta(x) = "ZION PAMPA"
                    End If
                    '
                    'MESOPOTAMIA
                    If (StrComp(vmod, "ZI-MBME-0016", 0) = 0 Or StrComp(vmod, "ZI-MBME-0018", 0) = 0 Or StrComp(vmod, "ZI-MBME-0020", 0) = 0 Or StrComp(vmod, "", 0) = 0) Then
                        ruta(x) = "ZION MESOPOTAMIA"
                    End If
                    'Lowrider
                    If (StrComp(vmod, "GR-MBLN-0016", 0) = 0 Or StrComp(vmod, "GR-MBLN-0018", 0) = 0 Or StrComp(vmod, "GR-MBLN-0020", 0) = 0 Or StrComp(vmod, "GR-MBLA-0016", 0) = 0 Or StrComp(vmod, "GR-MBLA-0018", 0) = 0 Or StrComp(vmod, "GR-MBLA-0020", 0) = 0) Then
                         ruta(x) = "GRAVITY LOWRIDER"
                    End If
                    If (StrComp(vmod, "GR-MBSV-0016", 0) = 0 Or StrComp(vmod, "GR-MBSV-0018", 0) = 0 Or StrComp(vmod, "GR-MBSV-0020", 0) = 0 Or StrComp(vmod, "GR-MBSR-0016", 0) = 0 Or StrComp(vmod, "GR-MBSR-0018", 0) = 0 Or StrComp(vmod, "GR-MBSR-0020", 0) = 0) Then
                         ruta(x) = "GRAVITY SMASH"
                    End If
                    
                    'FIN
                    
                    'VACUUM CLEANERS
                    If (StrComp(vmod, "DW-RVDE-1KN", 0) = 0) Then
                        ruta(x) = "DW-RVDE-1KN"
                    End If
                    If (StrComp(vmod, "DW-RVDE-1WN", 0) = 0) Then
                        ruta(x) = "DW-RVDE-1WN"
                    End If
                    'FIN
                    
                    If (StrComp(vmod, "EC-KEP18MN", 0) = 0) Then
                        ruta(x) = "EC-KEP18MN"
                    End If
                    If (StrComp(vmod, "EC-KEP18MN ", 0) = 0) Then
                        ruta(x) = "EC-KEP18MN"
                    End If
                    If (StrComp(vmod, "EC-KEP18MN  ", 0) = 0) Then
                        ruta(x) = "EC-KEP18MN"
                    End If
'                    'THUNDERSKY
'                    If (StrComp(vmod, "", 0) = 0 Or StrComp(vmod, "", 0) = 0) Then
'                        ruta = "THUNDER SKY"
'                    End If
'                    'RAGE LYON
'                    If (StrComp(vmod, "", 0) = 0 Or StrComp(vmod, "", 0) = 0) Then
'                        ruta= "RAGE LYON"
'                    End If
             
                    
    


End Sub
Sub ej()

For x = 0 To contadordepaginas
    For Item = 0 To 300
    If (StrComp("", MatrizTv(x, 7, cantelementos(x)), 0) = 0) Then
    Else
 Call Busqueda_nombremod_img(VecModelos(x), rutaimg(x, Item), x)
    End If
    Next Item
Next x
End Sub
