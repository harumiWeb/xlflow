Attribute VB_Name = "modTempCache"
Option Explicit

Public Sub DeleteAllTempCaches(ByVal targetWorkbook As Workbook)
    On Error GoTo ErrHandler

    modSheetStore.DeleteSheetsWithPrefix targetWorkbook, modSheetStore.TempSheetPrefix()
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "TempCache", "modTempCache.DeleteAllTempCaches", Err.description, "", "", "temporary sheet cleanup"
    Err.Raise Err.Number, "modTempCache.DeleteAllTempCaches", Err.description
End Sub

Public Function DeleteAllTempCachesForThisWorkbook() As String
    On Error GoTo ErrHandler

    DeleteAllTempCaches ThisWorkbook
    DeleteAllTempCachesForThisWorkbook = "ok:" & CStr(ThisWorkbook.Worksheets.Count)
    Exit Function

ErrHandler:
    DeleteAllTempCachesForThisWorkbook = "error:" & Err.description
End Function
