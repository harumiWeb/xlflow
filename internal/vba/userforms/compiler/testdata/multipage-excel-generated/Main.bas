Attribute VB_Name = "Main"
Option Explicit
Public Sub RunMultiSentinel()
    On Error GoTo Failed
    Dim form As Object, page As Object, item As Object, result As String
    Set form = VBA.UserForms.Add("MultiTopologyForm")
    result = "issue-885-ok|" & CStr(form.MultiMain.Pages.Count)
    If form.MultiMain.Pages.Count = 0 Then
        result = result & ":-1"
    Else
        result = result & ":" & CStr(form.MultiMain.Value)
    End If
    For Each page In form.MultiMain.Pages
        result = result & "|" & page.Name & ":" & page.Caption & ":" & CStr(page.Enabled) & ":" & CStr(page.Visible)
    Next page
    result = result & "|tabs=" & CStr(form.StripMain.Tabs.Count) & ":" & CStr(form.StripMain.Value)
    For Each item In form.StripMain.Tabs
        result = result & "|" & item.Name & ":" & item.Caption & ":" & CStr(item.Enabled) & ":" & CStr(item.Visible)
    Next item
    If form.MultiMain.Pages.Count > 0 Then
        result = result & "|text=" & form.PageText.Value & "|label=" & form.NestedLabel.Caption
    Else
        result = result & "|text=|label="
    End If
    result = result & "|enabled=" & CStr(form.MultiMain.Enabled) & ":" & CStr(form.StripMain.Enabled)
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = result
    Unload form
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-885-failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
