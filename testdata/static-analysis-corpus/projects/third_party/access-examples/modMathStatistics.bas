Attribute VB_Name = "modMathStatistics"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Statistics
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Function Combin(n As Integer, M As Integer) As Double
'
' Returns the number of combinations, without regard to order,
' of M items that can be made from a pool of N items.
'
' N must be greater of equal to M.
' M must be between 0 and N.
'
  Combin = Factorial(n) / Factorial(M) / Factorial(n - M)
End Function

Function Factorial(x As Integer) As Double
'
' Non-recursive factorial
'
Dim i As Integer, f As Double
  If x < 0 Or x > 170 Then
    Factorial = 0#
    Exit Function
  End If
  f = 1#
  For i = 2 To x
    f = f * i
  Next i
  Factorial = f
End Function

Function FactorialR(x As Integer) As Double
'
' Recursive factorial
'
  If x < 0 Or x > 170 Then
    FactorialR = 0#
  ElseIf x = 0 Then
    FactorialR = 1#
  Else
    FactorialR = x * FactorialR(x - 1)
  End If
End Function

Function Permut(n As Integer, M As Integer) As Double
'
' Returns the number of combinations, with regard to order,
' of M items that can be made from a pool of N items.
'
' Use this for lottory-style calculations.
'
' N must be greater of equal to M.
' M must be between 0 and N.
'
  Permut = Factorial(n) / Factorial(n - M)
End Function

Sub Regress(a() As Double, ByVal XCol As Integer, YCol As Integer, ByVal n As Integer, Slope As Double, YIntercept As Double)
'
' Performs Linear regression on array A()
' A() must be a 2-dimensional array - e.g. Redim A(50, 0 To 1) as Double
' XCol indicates which column represents the X Axis
' YCol indicates which column represents the Y Axis
' The slope of the best fit line is returned in Slope
' The Y intercept is returned in YIntercept
' The X intercept = -YIntercept/Slope - This will produce a run-time error if the line is horizontal
'
' The function will produce a run-time error if there are less than 2 points,
' the line is vertical, or XCol or YCol are invalid
'
' The function sums all rows from 1 to N
'
Dim i As Integer, SumX As Double, SumY As Double, SumXY As Double, SumXX As Double, x As Double, y As Double
  SumX = 0
  SumY = 0
  SumXY = 0
  SumXX = 0
  For i = 1 To n
    x = a(i, XCol)
    y = a(i, YCol)
    SumX = SumX + x
    SumY = SumY + y
    SumXY = SumXY + x * y
    SumXX = SumXX + x * x
  Next i
  Slope = (n * SumXY - SumX * SumY) / (n * SumXX - SumX * SumX)
  YIntercept = (SumY * SumXX - SumX * SumXY) / (n * SumXX - SumX * SumX)
End Sub

Sub TestRegress()
'
' Use this procedure to test the Regress procedure
'
Dim Slope As Double, YIntercept As Double, XIntercept As Double
ReDim a(5, 1) As Double
  a(1, 0) = 1: a(1, 1) = 1
  a(2, 0) = 2: a(2, 1) = 3
  a(3, 0) = 3: a(3, 1) = 5
  a(4, 0) = 4: a(4, 1) = 7
  a(5, 0) = 5: a(5, 1) = 9
  Regress a(), 0, 1, 5, Slope, YIntercept
  XIntercept = -YIntercept / Slope ' This will error if line is horizontal
  Debug.Print "Y = " & Slope & " * X + " & YIntercept
  Debug.Print "X = " & (1 / Slope) & " * Y + " & XIntercept
End Sub
