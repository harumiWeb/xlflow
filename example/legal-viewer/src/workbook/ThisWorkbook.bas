Option Explicit

Private Sub Workbook_Open()
    On Error GoTo ErrHandler

    modAppStartup.AppInitializeWorkbook Me
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "ThisWorkbook.Workbook_Open", Err.description, "", "", ""
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrHandler

    modAppStartup.AppBeforeCloseWorkbook Me, Cancel
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "ThisWorkbook.Workbook_BeforeClose", Err.description, "", "", ""
    Cancel = True
End Sub










