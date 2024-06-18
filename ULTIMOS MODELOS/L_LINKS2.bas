Attribute VB_Name = "L_LINKS2"
Sub links2()

 Dim Cadena(149) As String
        
        Sheets("Inspeccion").Select
        Call ponercero(Cadena)
        Call cargardirecciones(Cadena)
        For x = 0 To 149
        With Sheets("inspeccion")
        .Hyperlinks.Add Anchor:=.Range("T9").Offset(0, x), _
         Address:="" & Cadena(x), _
         ScreenTip:="Link a inspeccion del producto", _
         TextToDisplay:=" LINK "
        End With
        
        Next x
        
        Range("W2").Select

End Sub



Sub ponercero(ByRef cad() As String)
        
            For x = 0 To 149
            
                    cad(x) = "0"
            Next x

    
End Sub

Sub cargardirecciones(ByRef cad() As String)
    Sheets("Inspeccion").Select
    Range("T8").Activate
    
    
    For x = 0 To 149
    cad(x) = ActiveCell.Offset(0, x).Value
    
    
    Next x

End Sub

