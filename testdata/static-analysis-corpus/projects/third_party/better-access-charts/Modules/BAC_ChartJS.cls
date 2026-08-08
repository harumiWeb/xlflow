VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ChartJS"
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

Private m_Source As chScriptSource
Private m_CDN As String
Private m_PathFilename As String


' ### Properties

Public Property Let Source(ByVal This_Source As chScriptSource)
    m_Source = This_Source
End Property
Public Property Get Source() As chScriptSource
    Source = m_Source
End Property

Public Property Let CDN(ByVal This_CDN As String)
    m_CDN = This_CDN
End Property
Public Property Get CDN() As String
    CDN = m_CDN
End Property

Public Property Let PathFilename(ByVal This_PathFilename As String)
    m_PathFilename = This_PathFilename
End Property
Public Property Get PathFilename() As String
    PathFilename = m_PathFilename
End Property


' ### Instanzing

Private Sub Class_Initialize()

    Me.Source = chScriptSource.CDN
    Me.CDN = "https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"
    Me.PathFilename = CurrentProject.Path & "\chart.umd.min.js"
  
End Sub


' ### Public Functions

Public Function SaveToDisk() As Boolean

    SaveToDisk = BAC.Helper.SaveFileToDisk("chart.umd.min.js", CurrentProject.Path)
    
End Function


' ### Generator Function

Friend Function getLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.Source
        Case chScriptSource.CDN
            strText = "'" & Me.CDN & "'"
        Case chScriptSource.LocalFile
            strText = "'" & "https://msaccess/" & Replace(Me.PathFilename, "\", "/") & "'"
        Case Else
            'Do nothing
    End Select
            
    getLocation = strText

End Function

