Attribute VB_Name = "App"
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
  Dim apiKey As String
  Dim previousDisplayAlerts As Boolean
  Dim previousEnableEvents As Boolean
  Dim previousScreenUpdating As Boolean
  Dim errorNumber As Long
  Dim errorDescription As String
  Dim errorSource As String

  On Error GoTo ErrHandler

  previousDisplayAlerts = Application.DisplayAlerts
  previousEnableEvents = Application.EnableEvents
  previousScreenUpdating = Application.ScreenUpdating

  Application.DisplayAlerts = False
  Application.EnableEvents = False
  Application.ScreenUpdating = False

  apiKey = Trim$(Environ$("TWELVE_DATA_API_KEY"))
  If Len(apiKey) = 0 Then
    Err.Raise vbObjectError + 700, "App.RunCore", "TWELVE_DATA_API_KEY is not set."
  End If

  MarketDashboard.RefreshDashboard wb, apiKey

Cleanup:
  On Error GoTo CleanupFailed
  Application.DisplayAlerts = previousDisplayAlerts
  Application.EnableEvents = previousEnableEvents
  Application.ScreenUpdating = previousScreenUpdating
  On Error GoTo 0
  If errorNumber <> 0 Then
    Err.Raise errorNumber, errorSource, errorDescription
  End If
  Exit Sub

ErrHandler:
  errorNumber = Err.Number
  errorDescription = Err.Description
  errorSource = Err.Source
  Resume Cleanup

CleanupFailed:
  If errorNumber = 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
  End If
  On Error GoTo 0
  Err.Raise errorNumber, errorSource, errorDescription
End Sub
