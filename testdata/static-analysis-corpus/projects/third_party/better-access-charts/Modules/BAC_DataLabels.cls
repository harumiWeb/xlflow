VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_DataLabels"
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
Private m_Anchor As chDataLabelAnchor
Private m_Font As BAC_Font

Private m_SectionIsNeeded As Boolean


' ### Properties

Public Property Let Show(ByVal This_Show As Boolean)
    m_Show = This_Show
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Anchor(ByVal This_Anchor As chDataLabelAnchor)
    m_Anchor = This_Anchor
    If This_Anchor > 0 Then
        m_SectionIsNeeded = True
    End If
End Property
Public Property Get Anchor() As chDataLabelAnchor
    Anchor = m_Anchor
End Property

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property


' ### Instanzing

Private Sub Class_Initialize()

    Set m_Font = New BAC_Font
    
End Sub


' ### Generator Function

Friend Function getDataLabel(Options As BAC_Chart) As String

    If Not Me.Show Then Exit Function
    
    With New BAC_StringBuilder
        .AddLine "                datalabels: {"
        If Me.Anchor > 0 Then
            .AddLine "                    anchor: '" & getAnchorString(Me.Anchor) & "',"
        End If
        Select Case Options.ChartType
            Case chChartType.Bar, chChartType.Bubble, chChartType.Line, chChartType.Scatter
                .AddLine "                    color: (context) => {"
                If Options.Horizontal = True Then
                    .AddLine "                        const y = context.chart.scales.y;"
                    .AddLine "                        if (context.dataIndex >= y.min && context.dataIndex <= y.max) {"
                Else
                    .AddLine "                        const x = context.chart.scales.x;"
                    .AddLine "                        if (context.dataIndex >= x.min && context.dataIndex <= x.max) {"
                End If
                If Me.Font.Color <> vbNullString Then
                    .AddLine "                            return '" & Me.Font.Color & "';"
                Else
                    .AddLine "                            return ;"
                End If
                .AddLine "                        } else {"
                .AddLine "                            return 'transparent';"
                .AddLine "                        }"
                .AddLine "                    },"
            Case Else
                'Do Nothing
        End Select
        If Me.Font.SectionIsNeeded("DataLabel") Then
            .AddLine "                    font: {"
            If Me.Font.Family <> vbNullString Then
                .AddLine "                        family: '" & Me.Font.Family & "',"
            End If
            If Me.Font.Size > 0 Then
                .AddLine "                        size: " & Me.Font.Size & ","
            End If
            If Me.Font.Weight <> chFontWeight.fwUndefined Then
                .AddLine "                        weight: " & Me.Font.Weight & ","
            End If
            If Me.Font.Style <> chFontStyle.fsUndefined Then
                .AddLine "                        style: " & Me.Font.StyleText
            End If
            .CutTrailingComma
            .AddLine "                        },"
        End If
        .CutTrailingComma
        .AddLine "                },"
        
        getDataLabel = .ToString
    End With
    
End Function

Private Function getAnchorString(ByVal ThisAnchor As chDataLabelAnchor) As String

    Select Case ThisAnchor
        Case chDataLabelAnchor.anStart
            getAnchorString = "start"
        Case chDataLabelAnchor.anCenter
            getAnchorString = "center"
        Case chDataLabelAnchor.anEnd
            getAnchorString = "end"
        Case Else
            getAnchorString = "center"
    End Select
    
End Function

