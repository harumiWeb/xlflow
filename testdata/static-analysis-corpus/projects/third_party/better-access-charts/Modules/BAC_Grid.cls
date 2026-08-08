VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Grid"
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


'### Members

Private m_Show As Boolean
Private m_Color As String
Private m_LineWidth As Integer
Private m_Circular As Boolean

'### Properties

Public Property Let Show(ByVal newValue As Boolean)
    m_Show = newValue
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Color(ByVal newValue As String)
    m_Color = newValue
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let LineWidth(ByVal newValue As Integer)
    m_LineWidth = newValue
End Property
Public Property Get LineWidth() As Integer
    LineWidth = m_LineWidth
End Property

Public Property Let Circular(ByVal newValue As Boolean)
    m_Circular = newValue
End Property
Public Property Get Circular() As Boolean
    Circular = m_Circular
End Property


'### Instancing

Private Sub Class_Initialize()

    m_Show = True
    m_LineWidth = 1
    m_Circular = False
    
End Sub


'### Generator Function

Friend Function getGrid() As String
    
    With New BAC_StringBuilder
        If m_Show = False Then
            .AddLine "                    grid: {"
            .AddLine "                        display: false"
            .AddLine "                    },"
        Else
            .AddLine "                    grid: {"
            .AddLine "                        display: true,"
            If Len(m_Color) > 0 Then
                  .AddLine "                        color: '" & m_Color & "',"
            End If
            If m_LineWidth > 1 Then
                  .AddLine "                        lineWidth: " & m_LineWidth & ","
            End If
            If m_Circular = True Then
                  .AddLine "                        circular: true,"
            End If
            .CutTrailingComma
            .AddLine "                    },"
        End If
        getGrid = .ToString
    End With
    
End Function

