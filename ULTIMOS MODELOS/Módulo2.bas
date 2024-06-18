Attribute VB_Name = "Módulo2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'


'
    
    Sheets ("BOM ZI-MBOV-0018")
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 1133.25, 35.25, 236.25, 48.75). _
        Select
    Selection.ShapeRange.ScaleWidth 1.019047619, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset23
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
        .Solid
    End With
    
    
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset28
    With Selection.ShapeRange.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "VISTA 1"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent6
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 11
    End With
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .NameComplexScript = "Rockwell"
        .NameFarEast = "Rockwell"
        .Name = "Rockwell"
    End With
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 14
    Range("M8").Select
    ActiveSheet.Shapes.Range(Array("Rectangle 1")).Select
    Selection.OnAction = "autofit"
    Range("L9").Select
End Sub
