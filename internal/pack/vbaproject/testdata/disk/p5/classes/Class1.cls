VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strName As String


Private Sub Class_Initialize()
    strName = "Hello, world"
End Sub


Private Sub Class_Terminate()
    Debug.Print strName
End Sub

Public Property Get Name() As String
    Name = strName
End Property

Public Property Let Name(name_ As String)
    strName = name_
End Property
