VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Data"
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

Private m_LabelCount As Integer
Private m_Labels() As String        ' 1..LabelCount

Private m_DataCount As Integer
Private m_Data() As String          ' 1..DataCount
Private m_DatasetNames() As String  ' 1..DataCount


' ### Properties

Public Property Get DataCount() As Integer
    DataCount = m_DataCount
End Property
Public Property Let DataCount(ByVal newValue As Integer)
    m_DataCount = newValue
    ReDim Preserve m_Data(newValue - 1)
    ReDim Preserve m_DatasetNames(newValue - 1)
End Property

Public Property Get IsSingleDataset() As Boolean
    IsSingleDataset = (DataCount = 1)
End Property

Public Property Get Data(ByVal Index As Integer) As String
    Data = m_Data(Index)
End Property
Public Property Let Data(ByVal Index As Integer, ByVal newValue As String)
    m_Data(Index) = newValue
End Property

Public Property Get DatasetName(ByVal Index As Integer) As String
    DatasetName = m_DatasetNames(Index)
End Property
Public Property Let DatasetName(ByVal Index As Integer, ByVal newValue As String)
    m_DatasetNames(Index) = newValue
End Property

Public Property Get LabelCount() As Integer
    LabelCount = m_LabelCount
End Property
Public Property Let LabelCount(ByVal newValue As Integer)
    m_LabelCount = newValue
    ReDim Preserve m_Labels(newValue - 1)
End Property

Public Property Get Label(ByVal Index As Integer) As String
    Label = m_Labels(Index)
End Property
Public Property Let Label(ByVal Index As Integer, ByVal newValue As String)
    m_Labels(Index) = Replace(newValue, "'", "´")
End Property

Public Property Get Labels(Optional ByVal Quoted As Boolean = True) As String
    Dim arrLabels() As String
    Dim i As Integer
    
    If Quoted Then
        ReDim arrLabels(LabelCount)
        For i = 0 To LabelCount - 1
            arrLabels(i) = Quote(m_Labels(i))
        Next
    Else
        arrLabels = m_Labels
    End If
    Labels = Strings.Join(arrLabels, ",")
End Property


'### Initializer

Public Sub Init()

    m_LabelCount = 0
    m_DataCount = 0
    
End Sub


' ### Methods

Public Sub AddLabel(ByVal TheLabel As String)

    LabelCount = LabelCount + 1
    Label(LabelCount - 1) = TheLabel
    
End Sub

Public Sub AddLabels(ByVal TheLabels As Variant)

    If IsArray(TheLabels) Then
        m_Labels = TheLabels
    Else
        m_Labels = Strings.Split(TheLabels, ",")
    End If
    m_LabelCount = UBound(m_Labels) + 1
    
End Sub

Public Sub AddDatasetName(ByVal TheDatasetName As String)

    DataCount = DataCount + 1
    DatasetName(DataCount - 1) = TheDatasetName
    
End Sub

Public Sub AddData(ByVal TheDataset As Variant, ByVal TheData As String)

    'Declarations
    Dim idx As Integer
    
    Select Case VarType(TheDataset)
        Case vbString: idx = GetDatasetIndex(TheDataset)
        Case vbInteger, vbLong: idx = TheDataset
        Case Else: Err.Raise vbObjectError, , "Invalid Datatype for TheDataset": Exit Sub
    End Select
    If idx < 0 Then
        Err.Raise vbObjectError, , "Invalid Index for AddData": Exit Sub
    End If
    If Data(idx) = "" Then
        Data(idx) = TheData
    Else
        Data(idx) = Data(idx) & "," & TheData
    End If
    
End Sub


' ### Public Function

Public Function Result() As BAC_Data

    ' Cleanup the data and return me as Result
    Dim i  As Integer
    
    For i = 0 To DataCount - 1
        Data(i) = BAC.Helper.CutTrailingComma(Data(i))
    Next
    For i = 0 To LabelCount - 1
        Label(i) = BAC.Helper.CutTrailingComma(Label(i))
    Next
    Set Result = Me
    
End Function


' ### Private Functions

Private Function GetDatasetIndex(ByVal TheDatasetName As String)

    'Declarations
    Dim i As Integer
    
    For i = 0 To DataCount - 1
        If DatasetName(i) = TheDatasetName Then
            GetDatasetIndex = i
            Exit Function
        End If
    Next
    
    'Add dataset name if not already exists
    Call AddDatasetName(TheDatasetName)
    GetDatasetIndex = DataCount - 1
    
End Function

Private Function Quote(ByVal TheString As String) As String

    If Left(TheString, 1) <> "'" And Right(TheString, 1) <> "'" Then
        Quote = "'" & TheString & "'"
    Else
        Quote = TheString
    End If
    
End Function
