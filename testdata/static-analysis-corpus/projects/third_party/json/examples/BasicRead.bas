Attribute VB_Name = "BasicRead"
Option Explicit

Public Sub Example_ReadObjectFields()
    Dim text As String
    text = "{""name"":""Ueslei"",""age"":18,""active"":true}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringValue("name")
    Debug.Print doc.NumberValue("age")
    Debug.Print doc.BoolValue("active")
End Sub

Public Sub Example_ReadNestedObject()
    Dim text As String
    text = "{""user"":{""name"":""Ueslei"",""role"":""developer""}}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Dim user As JSON
    Set user = doc.Node("user")

    If Not user Is Nothing Then
        Debug.Print user.StringValue("name")
        Debug.Print user.StringValue("role")
    End If
End Sub

Public Sub Example_ReadArray()
    Dim items As JSON
    Set items = JSON.Parse("[""Excel"",""Access"",""Word""]")

    Dim token As Long
    token = items.FirstChildToken()
    
    Do While token <> 0
        Debug.Print items.TokenStringValue(token)
        token = items.NextToken(token)
    Loop
End Sub

Public Sub Example_DefaultMemberChaining()
    Dim myJson As JSON
    Set myJson = JSON.Parse("{""names"":[""Ana"",""Bia"",""Caio""]}")

    Debug.Print myJson("names")(0)
    Debug.Print myJson("names")(1)
End Sub
