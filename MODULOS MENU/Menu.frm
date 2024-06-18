VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "UserForm1"
   ClientHeight    =   8370.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16545
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cerrar_Click()

Dim i As Long
    i = 192
    Do Until i = 0
        DoEvents
        i = i - 1
    Loop

Me.Frame2.Width = 0
Me.Frame1.Left = Me.Frame2.Width
Me.btn_menu.Visible = True

Me.Frame1.Width = Me.Width - 1


End Sub

Private Sub btn_menu_Click()

Dim i As Long
    i = 192
    Do Until i = 0
        DoEvents
        i = i - 1
    Loop
    
Me.Frame2.Width = 192
Me.Frame1.Left = Me.Frame2.Width
'Me.Frame1.Width = Me.Width - Me.Frame2.Width

Me.btn_cerrar.Visible = True
Me.btn_menu.Visible = False

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Call btn_cerrar_Click

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ruta As String

    ruta = wb.Path()

    

    With Frame1
        .Picture = LoadPicture(ruta & "\imagenes\fondo\fondo_principal.jpg")
        .PictureSizeMode = 1
        .Height = Me.Height
        .Width = Me.Width - Me.Frame2.Width
    End With
    
    With Frame3
        .Picture = LoadPicture(ruta & "\imagenes\fondo\fondo_1.jpg")
        .PictureSizeMode = 1

        .Height = Me.Height
        .Width = Me.Width

    End With
    

End Sub
    
    
