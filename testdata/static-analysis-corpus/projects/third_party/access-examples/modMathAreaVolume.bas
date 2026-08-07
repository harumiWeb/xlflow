Attribute VB_Name = "modMathAreaVolume"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Area and Volume
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Function ACircle(Radius As Double) As Double
'
' Area of circle given radius
' Uses PI() from Trig module
'
  ACircle = Radius * Radius * pi()
End Function

Function ARect(L As Double, W As Double) As Double
'
' Area of rectangle given length and width
'
  ARect = L * W
End Function

Function ARing(InnerRadius As Double, OuterRadius As Double) As Double
'
' Area of a ring defined by 2 radii
'
  ARing = ACircle(OuterRadius) - ACircle(InnerRadius)
End Function

Function ASphere(r As Double) As Double
'
' Area of a sphere given the radius
'
  ASphere = 4 * pi() * r * r
End Function

Function ASquare(Side As Double) As Double
'
' Area of square give length of a side
'
  ASquare = Side * Side
End Function

Function ASquare2(Diag As Double) As Double
'
' Area of square given length of diagonal
'
  ASquare2 = Diag * Diag / 2
End Function

Function ATrap(h As Double, L1 As Double, L2 As Double)
'
' Area of Trapezoid or Trapezium given lengths of parallel sides
' and perpendicular height
'
  ATrap = h * (L1 + L2) / 2
End Function

Function ATriangle(L As Double, h As Double) As Double
'
' Area of triangle given length of a side and perpendicular height
'
  ATriangle = L * h / 2
End Function

Function ATriangle2(a As Double, b As Double, c As Double) As Double
'
' Area of a triangle given the lengths of the 3 sides
'
Dim CosC As Double
  CosC = (a * a + b * b - c * c) / (2 * a * b)
  ATriangle2 = a * b * Sqr(1 - CosC * CosC) / 2
End Function

Function RectDiag(W As Double, L As Double) As Double
'
' Length of diagonal of a rectangle given the 2 sides
'
  RectDiag = Sqr(W * W + L * L)
End Function

Function SquareDiag(L As Double) As Double
'
' Length of the diagonal of a square side length L
'
  SquareDiag = L * Sqr(2)
End Function

Function VCone(h As Double, r As Double) As Double
'
' Volume of a cone given radius of base and height
'
  VCone = h * r * r * pi() / 3
End Function

Function VCylinder(h As Double, r As Double) As Double
'
' Volume of Cylinder given height and radius
' Uses PI() from Trig module
'
  VCylinder = pi() * r * r * h
End Function

Function VPipe(h As Double, OuterRadius As Double, InnerRadius As Double) As Double
'
' Volume of a pipe by subtracting 2 cylinders
'
  VPipe = VCylinder(h, OuterRadius) - VCylinder(h, InnerRadius)
End Function

Function VPyramid(h As Double, BaseArea As Double) As Double
'
' Volume of pyramid or cone given area of the base and the height
'
  VPyramid = h * BaseArea / 3
End Function

Function VSphere(r As Double) As Double
'
' Volume of a sphere given the radius
'
  VSphere = pi() * r * r * r * 4 / 3
End Function

Function VTruncPyramid(h As Double, BaseArea1 As Double, BaseArea2 As Double) As Double
'
' Volume of truncated pyramid given height and area of base and top
'
  VTruncPyramid = h * (BaseArea1 + BaseArea2 + Sqr(BaseArea1) * Sqr(BaseArea2)) / 3
End Function
