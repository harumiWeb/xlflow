Attribute VB_Name = "modMathTrig"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Trigonometry
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Function ArcCos(x As Double) As Double
' Inverse Cosine
  If x = 1 Then
    ArcCos = 0
  ElseIf x = -1 Then
    ArcCos = -pi()
  Else
    ArcCos = Atn(x / Sqr(-x * x + 1)) + pi() / 2
  End If
End Function

Function Arccosec(x As Double) As Double
' Inverse Cosecant
  Arccosec = Atn(x / Sqr(x * x - 1)) + (Sgn(x) - 1) * pi() / 2
End Function

Function Arccotan(x As Double) As Double
' Inverse Cotangent
  Arccotan = Atn(x) + pi() / 2
End Function

Function ArcSec(x As Double) As Double
' Inverse Secant
  ArcSec = Atn(x / Sqr(x * x - 1)) + Sgn(Sgn(x) - 1) * pi() / 2
End Function

Function ArcSin(x As Double) As Double
' Inverse Sine
  If x = 1 Then
    ArcSin = pi() / 2
  ElseIf x = -1 Then
    ArcSin = -pi() / 2
  Else
    ArcSin = Atn(x / Sqr(-x * x + 1))
  End If
End Function

Function ATan2(x As Double, y As Double) As Double
'
' Returns the ArcTangent based on X and Y coordinates
' If both X and Y are zero an error will occur.
'
' The positive X axis is assumed to be 0, going poistive in the
' counterclockwise direction, and negative in the clockwise direction.
'
  If x = 0 Then
    If y = 0 Then
      ATan2 = 1 / 0
    ElseIf y > 0 Then
      ATan2 = pi() / 2
    Else
      ATan2 = -pi() / 2
    End If
  ElseIf x > 0 Then
    If y = 0 Then
      ATan2 = 0
    Else
      ATan2 = Atn(y / x)
    End If
  Else
    If y = 0 Then
      ATan2 = pi()
    Else
      ATan2 = (pi() - Atn(Abs(y) / Abs(x))) * Sgn(y)
    End If
  End If
End Function

Function Cosec(x As Double) As Double
' Cosecant
  Cosec = 1 / Sin(x)
End Function

Function Cotan(x As Double) As Double
' Cotangent
  Cotan = 1 / Tan(x)
End Function

Function Deg2Rad(x As Double) As Double
' Degrees to radians
  Deg2Rad = x / 180 * pi()
End Function

Function HArccos(x As Double) As Double
' Inverse Hyperbolic Cosine
  HArccos = Log(x + Sqr(x * x - 1))
End Function

Function HArccosec(x As Double) As Double
' Inverse Hyperbolic Cosecant
  HArccosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function

Function HArccotan(x As Double) As Double
' Inverse Hyperbolic Tangent
  HArccotan = Log((x + 1) / (x - 1)) / 2
End Function

Function HArcsec(x As Double) As Double
' Inverse Hyperbolic Secant
  HArcsec = Log((Sqr(-x * x + 1) + 1) / x)
End Function

Function HArcsin(x As Double) As Double
' Inverse Hyperbolic Sine
  HArcsin = Log(x + Sqr(x * x + 1))
End Function

Function HArctan(x As Double) As Double
' Inverse Hyperbolic Tangent
  HArctan = Log((1 + x) / (1 - x)) / 2
End Function

Function HCos(x As Double) As Double
' Hyperbolic Cosine
  HCos = (Exp(x) + Exp(-x)) / 2
End Function

Function HCosec(x As Double) As Double
' Hyperbolic Cosecant = 1/HSin(X)
  HCosec = 2 / (Exp(x) - Exp(-x))
End Function

Function HCotan(x As Double) As Double
' Hyperbolic Cotangent = 1/HTan(X)
  HCotan = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function

Function HSec(x As Double) As Double
' Hyperbolic Secant = 1/HCos(X)
  HSec = 2 / (Exp(x) + Exp(-x))
End Function

Function HSin(x As Double) As Double
' Hyperbolic Sine
  HSin = (Exp(x) - Exp(-x)) / 2
End Function

Function HTan(x As Double) As Double
' Hyperbolic Tangent = HSin(X)/HCos(X)
  HTan = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function

Function pi() As Double
  pi = Atn(1) * 4
End Function

Function Rad2Deg(x As Double) As Double
' Radians to Degrees
  Rad2Deg = x / pi() * 180
End Function

Function Sec(x As Double) As Double
' Secant
' This will choke at PI/2 and 3PI/2 radians (90 & 270 degrees)
'
  Sec = 1# / Cos(x)
End Function
