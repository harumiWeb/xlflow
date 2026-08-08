VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC__Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.29.01  published: 29.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Helper  As BAC_Helper
Private m_Scripts As BAC_Scripts


'### Instancing

Private Sub Class_Initialize()
    Set m_Helper = New BAC_Helper
    Set m_Scripts = New BAC_Scripts
End Sub

    
'### Initializer

Public Function Chart(Optional ByVal TheControl As Edge, _
                Optional ByVal ChartType As chChartType = chChartType.Bar, _
                Optional ByVal DataSourceType As chDataSourceType = chDataSourceType.dstDemo, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldNames As Variant, _
                Optional ByVal LabelFieldName As String) As BAC_Chart
  
    Set Chart = New BAC_Chart
    Chart.Init TheControl, ChartType, DataSourceType, DataSource, DataFieldNames, LabelFieldName
  
End Function


'### Public Methods

Friend Function Helper() As BAC_Helper
    Set Helper = m_Helper
End Function

Public Function Scripts() As BAC_Scripts
    Set Scripts = m_Scripts
End Function

Public Function Generator(ByVal ChartOptions As BAC_Chart, ByVal ChartData As BAC_Data) As BAC_ChartGenerator
    Set Generator = New BAC_ChartGenerator
    Generator.Init ChartOptions, ChartData
End Function


