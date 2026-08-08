VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.29.01  published: 29.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chTooltippPosition
    ttpAverage = 1
    ttpNearest = 2
End Enum


'### Members

Private m_BackgroundColor As String
Private m_BorderColor As String
Private m_BorderWidth As Integer
Private m_CaretSize As Integer
Private m_CaretPadding As Integer
Private m_DisplayColors As Boolean
Private m_Enabled As Boolean
Private m_Padding As Integer
Private m_Position As chTooltippPosition
Private m_UsePointStyle As Boolean


'### Properties

Public Property Let BackgroundColor(ByVal newValue As String)
    m_BackgroundColor = newValue
End Property
Public Property Get BackgroundColor() As String
    BackgroundColor = m_BackgroundColor
End Property

Public Property Let BorderColor(ByVal newValue As String)
    m_BorderColor = newValue
End Property
Public Property Get BorderColor() As String
    BorderColor = m_BorderColor
End Property

Public Property Let BorderWidth(ByVal newValue As Integer)
    m_BorderWidth = newValue
End Property
Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let CaretSize(ByVal newValue As Integer)
    m_CaretSize = newValue
End Property
Public Property Get CaretSize() As Integer
    CaretSize = m_CaretSize
End Property

Public Property Let CaretPadding(ByVal newValue As Integer)
    m_CaretPadding = newValue
End Property
Public Property Get CaretPadding() As Integer
    CaretPadding = m_CaretPadding
End Property

Public Property Let DisplayColors(ByVal newValue As Boolean)
    m_DisplayColors = newValue
End Property
Public Property Get DisplayColors() As Boolean
    DisplayColors = m_DisplayColors
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    m_Enabled = newValue
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Padding(ByVal newValue As Integer)
    m_Padding = newValue
End Property
Public Property Get Padding() As Integer
    Padding = m_Padding
End Property

Public Property Let Position(ByVal newValue As chTooltippPosition)
    m_Position = newValue
End Property
Public Property Get Position() As chTooltippPosition
    Position = m_Position
End Property

Public Property Let UsePointStyle(ByVal newValue As Boolean)
    m_UsePointStyle = newValue
End Property
Public Property Get UsePointStyle() As Boolean
    UsePointStyle = m_UsePointStyle
End Property


'### Instancing

Private Sub Class_Initialize()

    m_Enabled = True
    m_Position = ttpAverage
    m_Padding = 6
    m_CaretSize = 5
    m_CaretPadding = 2
    m_DisplayColors = True
    m_UsePointStyle = False
    
End Sub


'### Generator Function

Friend Function getTootip() As String
    
    With New BAC_StringBuilder
        If m_Enabled = False Then
            .AddLine "                tooltip: {"
            .AddLine "                    enabled: false"
            .AddLine "                },"
        Else
            .AddLine "                tooltip: {"
            .AddLine "                    enabled: true,"
            .AddLine "                    position: '" & getPositionAsText & "',"
            If Len(Me.BackgroundColor) > 0 Then
                .AddLine "                    backgroundColor: '" & Me.BackgroundColor & "',"
            End If
            .AddLine "                    padding: " & Me.Padding & ","
            .AddLine "                    caretSize: " & Me.CaretSize & ","
            .AddLine "                    caretPadding: " & Me.CaretPadding & ","
            If Me.DisplayColors = True Then
                .AddLine "                    displayColors: true,"
            Else
                .AddLine "                    displayColors: false,"
            End If
            If Me.UsePointStyle = True Then
                .AddLine "                    usePointStyle: true,"
            Else
                .AddLine "                    usePointStyle: false,"
            End If
            .AddLine "                    borderWidth: " & Me.BorderWidth & ","
            If Len(Me.BorderColor) > 0 Then
                .AddLine "                    borderColor: '" & Me.BorderColor & "',"
            End If
            .CutTrailingComma
            .AddLine "                },"
        End If
        getTootip = .ToString
    End With
    
End Function

Private Function getPositionAsText() As String

    'Declarations
    Dim strType As String

    Select Case Me.Position
        Case chTooltippPosition.ttpAverage
            strType = "average"
        Case chTooltippPosition.ttpNearest
            strType = "nearest"
        Case Else
    End Select

    getPositionAsText = strType

End Function

