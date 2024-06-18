Attribute VB_Name = "L_Links"


Sub links()
        Dim Cadena(149) As String
        Sheets("Master").Select
        
        Call ponercero(Cadena)
        Call cargardirecciones(Cadena)
            For x = 0 To 149
                        
                        If (StrComp(Cadena(x), "", vbBinaryCompare) = "0") Then
                        
                        Else
                        
                        With Sheets("Master")
                        .Hyperlinks.Add Anchor:=.Range("AC20").Offset(0, x), _
                         Address:="" & Cadena(x), _
                         ScreenTip:="Link a INSPRECCION DE PRODUCTO", _
                         TextToDisplay:=" LINK "
                         
                        End With
                        End If
                        
        
            Next x
        
        Range("AC1").Select
        

End Sub


Sub ponercero(ByRef cad() As String)
        
            For x = 0 To 149
            
                    cad(x) = "0"
            Next x

    
End Sub

Sub cargardirecciones(ByRef cad() As String)
    Sheets("Master").Select
    Range("AC19").Activate
    
    
    For x = 0 To 149
    cad(x) = ActiveCell.Offset(0, x).Value
    
    
    Next x

End Sub

