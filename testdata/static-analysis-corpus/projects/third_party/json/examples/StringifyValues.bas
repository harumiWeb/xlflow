Attribute VB_Name = "StringifyValues"
Option Explicit

Public Sub Example_StringifyDictionary()
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("name") = "JSON"
    data("version") = "1.0.1"
    data("fast") = True

    Debug.Print JSON.StringifyValue(data, True)
End Sub

Public Sub Example_StringifyCollection()
    Dim values As Collection
    Set values = New Collection

    values.Add "Excel"
    values.Add "Access"
    values.Add "Word"

    Debug.Print JSON.StringifyValue(values, True)
End Sub

Public Sub Example_StringifyArray()
    Dim values(0 To 2) As Variant
    values(0) = "VBA"
    values(1) = "JSON"
    values(2) = True

    Debug.Print JSON.StringifyValue(values)
End Sub

Public Sub Example_StringifyParsedNode()
    Dim doc As JSON
    Set doc = JSON.Parse("{""name"":""JSON"",""tags"":[""VBA"",""parser""]}")

    Debug.Print doc.Stringify()
    Debug.Print doc.Stringify(True)
    Debug.Print doc.StringifyWithIndent(True, vbTab)
End Sub
