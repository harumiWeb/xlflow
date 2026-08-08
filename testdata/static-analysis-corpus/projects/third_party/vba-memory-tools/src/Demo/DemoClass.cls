VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private m_id As Long

Public Function Factory(ByVal newID As Long) As DemoClass
    Dim c As New DemoClass
    Init c, newID
    Set Factory = c
End Function

Private Function Init(ByVal c As DemoClass, ByVal newID As Long) As Boolean
    RedirectInstance Init, VarPtr(Init), Me, c
    m_id = newID
End Function

Public Function Factory2(ByVal newID As Long) As DemoClass
    Dim c As New DemoClass
    '
    c.Init2 newID
    Set Factory2 = c
End Function
Public Function Init2(ByVal newID As Long) As DemoClass
    m_id = newID
End Function

Public Property Get ID() As Long
    ID = m_id
End Property
