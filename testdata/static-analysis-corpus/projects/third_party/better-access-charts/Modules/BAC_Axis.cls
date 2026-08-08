VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Axis"
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

Public Enum chCartesianAxis
    apXAxis = 1
    apYAxis = 2
End Enum

Public Enum chAxisType
    atNotDefined = 0
    atLinear = 1
    atLogarithmic = 2
    atCategory = 3
    atTime = 4
    atTimeSeries = 5
End Enum


'### Members

Private m_Axis As chCartesianAxis
Private m_Show As Boolean
Private m_LabelText As String
Private m_BeginAtZero As Boolean
Private m_TicksMin As String
Private m_TicksMax As String
Private m_StepSize As String
Private m_AxisType As chAxisType
Private m_Grid As BAC_Grid

Private m_TicksNeeded As Boolean


'### Properties

Public Property Let Show(ByVal This_ShowAxis As Boolean)
    m_Show = This_ShowAxis
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let LabelText(ByVal This_LabelText As String)
    m_LabelText = This_LabelText
End Property
Public Property Get LabelText() As String
    LabelText = m_LabelText
End Property

Public Property Let beginAtZero(ByVal This_BeginAtZero As Boolean)
    m_BeginAtZero = This_BeginAtZero
End Property
Public Property Get beginAtZero() As Boolean
    beginAtZero = m_BeginAtZero
End Property

Public Property Let TicksMin(ByVal This_TicksMin As String)
    m_TicksMin = This_TicksMin
End Property
Public Property Get TicksMin() As String
    TicksMin = m_TicksMin
End Property

Public Property Let TicksMax(ByVal This_TicksMax As String)
    m_TicksMax = This_TicksMax
End Property
Public Property Get TicksMax() As String
    TicksMax = m_TicksMax
End Property

Public Property Let StepSize(ByVal This_StepSize As String)
    m_StepSize = This_StepSize
    If Len(Nz(This_StepSize, vbNullString)) > 0 Then
        m_TicksNeeded = True
    End If
End Property
Public Property Get StepSize() As String
    StepSize = m_StepSize
End Property

Public Property Let AxisType(ByVal This_AxisType As chAxisType)
    m_AxisType = This_AxisType
End Property
Public Property Get AxisType() As chAxisType
    AxisType = m_AxisType
End Property

Public Property Get Grid() As BAC_Grid
    Set Grid = m_Grid
End Property


'### Initializer

Private Sub Class_Initialize()

    Set m_Grid = New BAC_Grid
    m_AxisType = chAxisType.atNotDefined
    
End Sub

Friend Sub Init(ByVal Axis As chCartesianAxis)

    m_Axis = Axis
    
End Sub


'### Generator Functions

Friend Function getCartesianAxis(ByVal StackedChart As Boolean) As String

    With New BAC_StringBuilder
        .AddLine "                #: {"
        .AddLine "                    id: 'first-#-Axis',"
    
        If StackedChart = True Then
            .AddLine "                    stacked: true,"
        End If
    
        If Me.Show = False Then
            .AddLine "                    display: false"
        Else
            .AddLine "                    display: true,"
            .Add Me.Grid.getGrid
            If Me.AxisType <> atNotDefined Then
                .AddLine "                    type: '" & getAxisTypeAsText & "',"
            End If
            If Len(Me.LabelText) > 0 Then
                .AddLine "                    title: {"
                .AddLine "                       display: true,"
                .AddLine "                       text: '" & Me.LabelText & "'"
                .AddLine "                    },"
            End If
            If Me.beginAtZero = True Then
                .AddLine "                    beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                    min: " & Str$(Me.TicksMin) & " ,"
                Else
                    .AddLine "                    min: '" & Me.TicksMin & "' ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                    max: " & Str$(Me.TicksMax) & " ,"
                Else
                    .AddLine "                    max: '" & Me.TicksMax & "' ,"
                End If
            End If
            If m_TicksNeeded = True Then
                .AddLine "                    ticks: {"
            End If
            If Len(Me.StepSize) > 0 Then
                .AddLine "                        stepSize: " & Str$(Me.StepSize) & " ,"
            End If
            If m_TicksNeeded = True Then
                .AddLine "                    },"
            End If
        End If
        .CutTrailingComma
        .AddLine "                },"
    
        If m_Axis = apXAxis Then
            getCartesianAxis = Replace(.ToString, "#", "x")
        ElseIf m_Axis = apYAxis Then
            getCartesianAxis = Replace(.ToString, "#", "y")
        End If
    
    End With
    
End Function

Friend Function getRadialAxis() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            .AddLine "                r: {"
            .AddLine "                    display: false"
            .AddLine "                }"
        Else
            .AddLine "                r: {"
            .AddLine "                    display: true,"
            .Add Me.Grid.getGrid
            If Me.beginAtZero = True Then
                .AddLine "                    beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                    min: " & Str$(Me.TicksMin) & " ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                    max: " & Str$(Me.TicksMax) & " ,"
                End If
            End If
            If m_TicksNeeded = True Then
                .AddLine "                    ticks: {"
            End If
            If Len(Me.StepSize) > 0 Then
                .AddLine "                        stepSize:" & Str$(Me.StepSize) & ","
            End If
            If m_TicksNeeded = True Then
                .AddLine "                    }"
            End If
            .CutTrailingComma
            .AddLine "                }"
        End If
        getRadialAxis = .ToString
    End With

End Function

Private Function getAxisTypeAsText() As String

    'Declarations
    Dim strType As String

    Select Case Me.AxisType
        Case chAxisType.atNotDefined
            strType = "Not defined"
        Case chAxisType.atLinear
            strType = "linear"
        Case chAxisType.atLogarithmic
            strType = "logarithmic"
        Case chAxisType.atCategory
            strType = "category"
        Case chAxisType.atTime
            strType = "time"
        Case chAxisType.atTimeSeries
            strType = "timeseries"
        Case Else
            strType = "Not defined"
    End Select

    getAxisTypeAsText = strType

End Function
