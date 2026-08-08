VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Font"
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

Private m_Color As String
Private m_Family As String
Private m_Size As Integer
Private m_Weight As chFontWeight
Private m_Style As chFontStyle


' ### Properties

Public Property Let Color(ByVal This_Color As String)
    m_Color = This_Color
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Family(ByVal This_Family As String)
    m_Family = This_Family
End Property
Public Property Get Family() As String
    Family = m_Family
End Property

Public Property Let Size(ByVal This_Size As Integer)
    m_Size = This_Size
End Property
Public Property Get Size() As Integer
    Size = m_Size
End Property

Public Property Let Weight(ByVal This_Weight As chFontWeight)
    m_Weight = This_Weight
End Property
Public Property Get Weight() As chFontWeight
    Weight = m_Weight
End Property

Public Property Let Style(ByVal This_Style As chFontStyle)
    m_Style = This_Style
End Property
Public Property Get Style() As chFontStyle
    Style = m_Style
End Property


' ### Functions

Friend Property Get SectionIsNeeded(Optional Parent As String) As Boolean

    Select Case Parent
        Case "DataLabel"
            SectionIsNeeded = (Family <> vbNullString) Or (Size > 0) Or (Weight <> fwUndefined) Or (Style <> fsUndefined)
        Case "Legend"
            SectionIsNeeded = (Color <> vbNullString) Or (Family <> vbNullString) Or (Size > 0) Or (Weight <> fwUndefined) Or (Style <> fsUndefined)
        Case "LegendTitle"
            SectionIsNeeded = (Family <> vbNullString) Or (Size > 0) Or (Weight <> fwUndefined) Or (Style <> fsUndefined)
        Case "Title"
            SectionIsNeeded = (Family <> vbNullString) Or (Size > 0) Or (Weight <> fwUndefined) Or (Style <> fsUndefined)
        Case Else
            'Do nothing
    End Select
    
End Property

Friend Property Get StyleText() As String

    Select Case Me.Style
        Case fsUndefined
            StyleText = vbNullString
        Case fsNormal
            StyleText = "'normal'"
        Case fsItalic
            StyleText = "'italic'"
        Case fsOblique
            StyleText = "'oblique'"
        Case Else
            'Do nothing
    End Select

End Property
