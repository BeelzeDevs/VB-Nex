Attribute VB_Name = "HideorUnhide"
Sub Hide()

Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    
End Sub

Sub Unhide()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",true)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True

End Sub
