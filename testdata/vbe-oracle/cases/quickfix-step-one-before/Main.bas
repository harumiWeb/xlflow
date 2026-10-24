Attribute VB_Name = "Main"
Option Explicit

Public Sub Probe()
    Dim i As Long
    For i = 1 To 3 Step 1 ' retain this comment
        Debug.Print i
    Next i
    For i = 1 To _
        3 Step 1
        Debug.Print i
    Next i
End Sub
