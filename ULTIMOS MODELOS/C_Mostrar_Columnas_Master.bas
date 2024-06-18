Attribute VB_Name = "C_Mostrar_Columnas_Master"
Sub Mostrar_Columnas_Master()

    Sheets("Master").Select
    Range("AC1:FV1").Select
    Selection.ColumnWidth = 40

End Sub
