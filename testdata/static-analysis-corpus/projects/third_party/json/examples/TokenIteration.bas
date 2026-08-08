Attribute VB_Name = "TokenIteration"
Option Explicit

Public Sub Example_TokenIteration()
    Dim text As String
    text = "{""rows"":[{""name"":""Ana"",""score"":10,""active"":true},{""name"":""Bia"",""score"":20,""active"":false}]}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Dim rows As JSON
    Set rows = doc.Node("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenString(t, "name")
        Debug.Print rows.TokenNumber(t, "score")
        Debug.Print rows.TokenBool(t, "active")

        t = rows.NextToken(t)
    Loop
End Sub

Public Sub Example_RawNestedField()
    Dim text As String
    text = "{""rows"":[{""name"":""Ana"",""profile"":{""rank"":""S"",""level"":12}}]}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Dim rows As JSON
    Set rows = doc.Node("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    If t <> 0 Then
        Debug.Print rows.TokenRawField(t, "profile")
    End If
End Sub
