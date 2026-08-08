VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Legend"
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

Private m_Font As BAC_Font
Private m_Title As BAC_Legend_Title
Private m_Show As Boolean
Private m_Position As chPosition
Private m_Align As chAlign


' ### Properties

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property

Public Property Get Title() As BAC_Legend_Title
    Set Title = m_Title
End Property

Public Property Let Show(ByVal newValue As Boolean)
    m_Show = newValue
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Position(ByVal newValue As chPosition)
    m_Position = newValue
End Property
Public Property Get Position() As chPosition
    Position = m_Position
End Property

Public Property Let Align(ByVal newValue As chAlign)
    m_Align = newValue
End Property
Public Property Get Align() As chAlign
    Align = m_Align
End Property


'### Instancing

Private Sub Class_Initialize()
    Set m_Font = New BAC_Font
    Set m_Title = New BAC_Legend_Title
End Sub


' ### Generator Function

Friend Function getLegend() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            .AddLine "                legend: {"
            .AddLine "                    display: false"
            .AddLine "                },"
        Else
            .AddLine "                legend: {"
            .AddLine "                    display: true,"
            .AddLine "                    position: " & BAC.Helper.getPositionString(Me.Position) & ","
            If Me.Align > 0 Then
                .AddLine "                        align: '" & getAlignString(Me.Align) & "',"
            End If
            If Me.Font.SectionIsNeeded("Legend") = True Then
                .AddLine "                    labels: {"
                If Me.Font.Color <> vbNullString Then
                    .AddLine "                        color: '" & Me.Font.Color & "',"
                End If
                .AddLine "                        font: {"
                If Me.Font.Family <> vbNullString Then
                    .AddLine "                            family: '" & Me.Font.Family & "',"
                End If
                If Me.Font.Size > 0 Then
                    .AddLine "                            size: " & Me.Font.Size & ","
                End If
                If Me.Font.Weight <> chFontWeight.fwUndefined Then
                    .AddLine "                            weight: " & Me.Font.Weight & ","
                End If
                If Me.Font.Style <> chFontStyle.fsUndefined Then
                    .AddLine "                            style: " & Me.Font.StyleText
                End If
                .CutTrailingComma
                .AddLine "                        },"
                .CutTrailingComma
                .AddLine "                    },"
            End If
            If Me.Title.Show = True Then
                .Add Me.Title.getTitle
            End If
            .CutTrailingComma
            .AddLine "                },"
        End If
        getLegend = .ToString
    End With
    
End Function

Private Function getAlignString(ByVal ThisAlign As chAlign) As String

    Select Case ThisAlign
        Case chAlign.alStart
            getAlignString = "start"
        Case chAlign.alCenter
            getAlignString = "center"
        Case chAlign.alEnd
            getAlignString = "end"
        Case Else
            getAlignString = "center"
    End Select
    
End Function

