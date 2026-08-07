Attribute VB_Name = "modMathCumulative"
Option Compare Database
Option Explicit

Public Function CumulativeValue( _
ByVal pvarSortCol As Variant, _
ByVal pstrSumColNme As String, _
ByVal pvarSumColValue As Variant, _
Optional pbSortAscending As Boolean = True) As Double
On Error GoTo ErrTrap
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Cumulative values within queries
' Notes:                If any of the below conditions are not true the results are UNPREDICTABLE
'                       No more than 10 running sums are included in any one query
' Example:
' Ascending             CumulativeValue([EndDate],"SomeNumber",[SomeNumber], True)
' Descending            CumulativeValue([EndDate],"SomeNumber",[SomeNumber], False)
'
' Variables:
' varSortCol            The column MUST have unique value, ex. Date, Autonumber, or Counter and the sort has to take place in the query itself.
' strSumColNme          Pass in the name of the column so you can use this function for more than one cumulative value in the same query.
' dblSumColValue        Pass in the number to sum as second variable.  If the number is null it will be set to zero.
' bSortAscending        Default is ascending.
'
' Ver.  Date            Author              Details
' 1.00  03 DEC 2002     Roger Saunders      Initial version.
' 1.01  19 FEB 2009     Anthony Duguid      added null handling as well as multiple columns
'--------------------------------------------------------------------------------------------------------------------
Static varLastRecord As Variant
Static RunningTotal(1 To 11, 1) 'Alias for field (text preceding colon in alias for output of field)
Dim intCount As Integer
Dim bAddQty As Boolean 'Add the value total
Dim dblSumColValue As Double

    dblSumColValue = CDbl(Nz(pvarSumColValue, 0)) 'used for stupid null value damn it
    
    If IsNull(pvarSortCol) Then
       bAddQty = False     ' Start of a new Query so reset
    End If
    
    If pbSortAscending Then ' Sorted Ascending
        If varLastRecord > pvarSortCol Then
           bAddQty = False ' Start of a new Query so reset
         Else
           bAddQty = True
        End If
    Else                   ' Sorted Descending
        If varLastRecord < pvarSortCol Then
           bAddQty = False ' Start of a new Query so reset
         Else
           bAddQty = True
        End If
    End If

    If bAddQty = False Then
       For intCount = 1 To 11
          RunningTotal(intCount, 0) = ""
          RunningTotal(intCount, 1) = 0
       Next intCount
    End If
    
    For intCount = 1 To 11
       If intCount = 11 Then
          RunningTotal(intCount, 1) = 0
          GoTo ExitProcedure
       End If
       If RunningTotal(intCount, 0) = "" Then
          '** Field Name is new add it to first empty spot in array
          RunningTotal(intCount, 0) = pstrSumColNme
          Exit For  'Exit for loop with intCount set to match field
       ElseIf RunningTotal(intCount, 0) = pstrSumColNme Then
          Exit For  'Exit for loop with intCount set to match field
       End If
    Next intCount
   RunningTotal(intCount, 1) = RunningTotal(intCount, 1) + dblSumColValue
   
ExitProcedure:
    On Error Resume Next
    CumulativeValue = RunningTotal(intCount, 1)
    varLastRecord = pvarSortCol
    Exit Function
    
ErrTrap:
    Debug.Print "Error Line: " & Erl
    Debug.Print "Error #: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    CumulativeValue = -1
    Exit Function
    
End Function

Public Function DeCumulativeValue( _
ByVal pvarSortCol As Variant, _
ByVal pstrSumColNme As String, _
ByVal pvarSumColValue As Variant, _
Optional pbSortAscending As Boolean = True) As Variant
On Error GoTo ErrTrap
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Used to Uncumulative values within queries
' Notes:                If any of the below conditions are not true the results are UNPREDICTABLE
'                       No more than 10 running sums are included in any one query
' Example:
' Ascending             DeCumulativeValue([EndDate],"SomeNumber",[SomeNumber], True)
' Descending            DeCumulativeValue([EndDate],"SomeNumber",[SomeNumber], False)
'
' Variables:
' varSortCol            The column MUST have unique value, ex. Date, Autonumber, or Counter and the sort has to take place in the query itself.
' strSumColNme          Pass in the name of the column so you can use this function for more than one cumulative value in the same query.
' dblSumColValue        Pass in the number to sum as second variable.  If the number is null it will be set to zero.
' bSortAscending        Default is ascending.
'
' Ver.  Date            Author              Details
' 1.00  07 SEP 2009     Anthony Duguid      Initial version, created from CumulativeValue function
' 1.00  07 OCT 2009     Anthony Duguid      pass back null value if parameter is null
'--------------------------------------------------------------------------------------------------------------------
Static varLastRecord As Variant
Static RunningTotal(1 To 11, 1) 'Alias for field (text preceding colon in alias for output of field)
Static CumValue(1 To 11, 1)
Dim intCount As Integer
Dim bAddQty As Boolean 'Add the value total
Dim dblSumColValue As Double

    If IsNull(pvarSumColValue) Then  'must do this for the null and negative values.
        DeCumulativeValue = Null
        Exit Function 'exit and do not set any of the static variables
    End If
    
    dblSumColValue = CDbl(Nz(pvarSumColValue, 0)) 'this should never be null
    
    If IsNull(pvarSortCol) Then
       bAddQty = False     ' Start of a new Query so reset
    End If
    
    If pbSortAscending Then ' Sorted Ascending
        If varLastRecord > pvarSortCol Then
           bAddQty = False ' Start of a new Query so reset
         Else
           bAddQty = True
        End If
    Else                   ' Sorted Descending
        If varLastRecord < pvarSortCol Then
           bAddQty = False ' Start of a new Query so reset
         Else
           bAddQty = True
        End If
    End If

    If bAddQty = False Then
       For intCount = 1 To 11
          RunningTotal(intCount, 0) = ""
          RunningTotal(intCount, 1) = 0
          CumValue(intCount, 1) = 0
       Next intCount
    End If
    
    For intCount = 1 To 11
       If intCount = 11 Then
          RunningTotal(intCount, 1) = 0
          CumValue(intCount, 1) = 0
          GoTo ExitProcedure
       End If
       If RunningTotal(intCount, 0) = "" Then
          '** Field Name is new add it to first empty spot in array
          RunningTotal(intCount, 0) = pstrSumColNme
          Exit For  'Exit for loop with intCount set to match field
       ElseIf RunningTotal(intCount, 0) = pstrSumColNme Then
          Exit For  'Exit for loop with intCount set to match field
       End If
    Next intCount
    
    If dblSumColValue - CumValue(intCount, 1) < 0 Then
        RunningTotal(intCount, 1) = 0
        CumValue(intCount, 1) = dblSumColValue
    Else
        RunningTotal(intCount, 1) = dblSumColValue - CumValue(intCount, 1)
        CumValue(intCount, 1) = dblSumColValue
    End If
    

ExitProcedure:
    On Error Resume Next
    DeCumulativeValue = RunningTotal(intCount, 1)
    varLastRecord = pvarSortCol
    Exit Function
    
ErrTrap:
    Debug.Print "Error Line: " & Erl
    Debug.Print "Error #: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    DeCumulativeValue = -666
    Exit Function
    
End Function
