Attribute VB_Name = "modUiState"
Option Explicit

Private Const DEFAULT_USERFORM_FONT As String = "ＭＳ ゴシック"

Public Function DefaultUserFormFontName() As String
    DefaultUserFormFontName = DEFAULT_USERFORM_FONT
End Function

Public Sub ApplyDefaultUserFormFont(ByVal targetForm As Object)
    ApplyUserFormFont targetForm, DEFAULT_USERFORM_FONT
End Sub

Public Sub ApplyDefaultControlFont(ByVal targetControl As Object)
    ApplyControlFont targetControl, DEFAULT_USERFORM_FONT
End Sub

Public Sub ApplyUserFormFont(ByVal targetForm As Object, ByVal fontName As String)
    On Error GoTo ErrHandler

    ApplyControlFont targetForm, fontName
    ApplyControlsFont targetForm.Controls, fontName
    Exit Sub

ErrHandler:
    modLogger.LogErrorSafe "UiState", "modUiState.ApplyUserFormFont", Err.description, "", "", TypeName(targetForm)
    Err.Raise Err.Number, Err.source, Err.description
End Sub

Public Function UserFormFontMatches(ByVal targetForm As Object, Optional ByVal fontName As String = DEFAULT_USERFORM_FONT) As Boolean
    If StrComp(targetForm.Font.Name, fontName, vbTextCompare) <> 0 Then
        Exit Function
    End If

    UserFormFontMatches = ControlsFontMatch(targetForm.Controls, fontName)
End Function

Private Sub ApplyControlsFont(ByVal targetControls As Object, ByVal fontName As String)
    Dim control As Object
    For Each control In targetControls
        ApplyControlFont control, fontName
        If StrComp(TypeName(control), "Frame", vbBinaryCompare) = 0 Then
            ApplyControlsFont control.Controls, fontName
        End If
    Next control
End Sub

Private Sub ApplyControlFont(ByVal targetControl As Object, ByVal fontName As String)
    targetControl.Font.Name = fontName
End Sub

Private Function ControlsFontMatch(ByVal targetControls As Object, ByVal fontName As String) As Boolean
    Dim control As Object
    For Each control In targetControls
        If StrComp(control.Font.Name, fontName, vbTextCompare) <> 0 Then
            ControlsFontMatch = False
            Exit Function
        End If
        If StrComp(TypeName(control), "Frame", vbBinaryCompare) = 0 Then
            If Not ControlsFontMatch(control.Controls, fontName) Then
                ControlsFontMatch = False
                Exit Function
            End If
        End If
    Next control

    ControlsFontMatch = True
End Function
