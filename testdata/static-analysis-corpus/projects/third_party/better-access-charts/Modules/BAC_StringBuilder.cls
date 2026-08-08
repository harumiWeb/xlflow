VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_StringBuilder"
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

Private m_String As String


'### Public Methods

Public Sub Add(ByVal TheString As String)
    m_String = m_String & TheString
End Sub

Public Sub AddLine(ByVal TheString As String)
    If m_String = vbNullString Then
        m_String = TheString & vbCrLf
    ElseIf TheString = vbNullString Then
        m_String = m_String & vbCrLf
    Else
        m_String = m_String & TheString & vbCrLf
    End If
End Sub

Public Sub CutTrailingComma()
    m_String = BAC.Helper.CutTrailingComma(m_String)
End Sub

Public Function ToString() As String
    ToString = m_String
End Function
