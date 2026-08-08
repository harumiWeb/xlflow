VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Scripts"
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

Private m_ChartJS       As BAC_ChartJS
Private m_AutoColors    As BAC_ChartJS_Plugin_AutoColors
Private m_DataLabels    As BAC_ChartJS_Plugin_Datalabels


'### Object Properties

Public Property Get ChartJS() As BAC_ChartJS
    Set ChartJS = m_ChartJS
End Property

Public Property Get Plugin_AutoColors() As BAC_ChartJS_Plugin_AutoColors
    Set Plugin_AutoColors = m_AutoColors
End Property

Public Property Get Plugin_DataLabels() As BAC_ChartJS_Plugin_Datalabels
    Set Plugin_DataLabels = m_DataLabels
End Property


'### Instancing

Private Sub Class_Initialize()

    Set m_ChartJS = New BAC_ChartJS
    Set m_AutoColors = New BAC_ChartJS_Plugin_AutoColors
    Set m_DataLabels = New BAC_ChartJS_Plugin_Datalabels
    
End Sub
