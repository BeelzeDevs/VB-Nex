Attribute VB_Name = "C_ocultar_Columnas_Master"
Sub ocultar_Columnas_Master()
    
    Sheets("master").Select
    
    Range("AC1:FV1").Select
    Selection.ColumnWidth = 0
    


End Sub




