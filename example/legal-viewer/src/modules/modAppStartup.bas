Attribute VB_Name = "modAppStartup"
Option Explicit

Public Sub AppInitializeWorkbook(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    modTempCache.DeleteAllTempCaches targetWorkbook
    modSheetStore.EnsurePersistentSheets targetWorkbook
    modSettings.EnsureDefaultSettings targetWorkbook
    modSheetStore.ApplyManagedSheetVisibility targetWorkbook
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "modAppStartup.AppInitializeWorkbook", Err.description, "", "", "Workbook_Open"
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Sub AppBeforeCloseWorkbook(ByVal targetWorkbook As Workbook, ByRef cancelClose As Boolean)
    On Error GoTo ErrHandler

    modTempCache.DeleteAllTempCaches targetWorkbook
    modSheetStore.ApplyManagedSheetVisibility targetWorkbook
SaveAndExit:
    If Not SaveWorkbookState(targetWorkbook) Then
        cancelClose = True
    End If
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "modAppStartup.AppBeforeCloseWorkbook", Err.description, "", "", "Workbook_BeforeClose"
    cancelClose = True
    Resume SaveAndExit
End Sub

Public Function SaveWorkbookState(ByVal targetWorkbook As Workbook) As Boolean
    On Error GoTo ErrHandler

    Dim originalAlerts As Boolean
    originalAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    targetWorkbook.Save
    SaveWorkbookState = True

Cleanup:
    Application.DisplayAlerts = originalAlerts
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "modAppStartup.SaveWorkbookState", Err.description, "", "", targetWorkbook.Name
    SaveWorkbookState = False
    Resume Cleanup
End Function

Public Function BeforeCloseSaveSmoke() As String
    On Error GoTo ErrHandler

    Dim cancelClose As Boolean
    AppBeforeCloseWorkbook ThisWorkbook, cancelClose
    If cancelClose Then
        Err.Raise vbObjectError + 7981, "modAppStartup.BeforeCloseSaveSmoke", "Before-close processing cancelled workbook close."
    End If
    If Not ThisWorkbook.Saved Then
        Err.Raise vbObjectError + 7982, "modAppStartup.BeforeCloseSaveSmoke", "Workbook was not saved by before-close processing."
    End If

    BeforeCloseSaveSmoke = "ok:" & ThisWorkbook.Name
    Exit Function

ErrHandler:
    modLogger.LogErrorSafe "AppStartup", "modAppStartup.BeforeCloseSaveSmoke", Err.description, "", "", ""
    BeforeCloseSaveSmoke = "error:" & Err.description
End Function
