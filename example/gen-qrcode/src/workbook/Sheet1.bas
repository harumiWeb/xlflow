

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
        If Target Is Nothing Then
                Exit Sub
        End If

        If Target.CountLarge <> 1 Then
                Exit Sub
        End If

        If Intersect(Target, Me.Range("B2")) Is Nothing Then
                Exit Sub
        End If

        App.HandleInputChanged Me
End Sub
