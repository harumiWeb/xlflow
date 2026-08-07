Attribute VB_Name = "modMathXYZ"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Latitude/Longitude
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Sub DegToDMS(ByVal L As Double, d As Integer, M As Integer, s As Double)
'
' Converts a decimal degree to Degrees, Minutes, and Seconds
' Seconds may contain up to 3 decimal places.
' e.g. 15.5 -> 15,30,0
'
  d = Int(L)
  L = (L - d) * 60
  M = Int(L)
  s = Val(Format((L - M) * 60, "#.###"))
End Sub

Function DegToDMSStr(ByVal L As Double) As String
'
' Converts a decimal value to Degrees, Minutes, and Seconds
' Processes Seconds up to 3 decimal places.
' e.g. 15.5 -> 15 30' 0"
'
Dim d As Integer, M As Integer, s As Double
  d = Int(L)
  L = (L - d) * 60
  M = Int(L)
  s = Val(Format((L - M) * 60, "#.###"))
  DegToDMSStr = d & " " & M & "' " & s & """"
End Function

Function DMSStrToDeg(ByVal DMS As String) As Double
'
' Converts a string value in the format [d m' s"] to a decimal number.
' Not all elements need be present, and they may contain decimal digits.
' e.g. 15 30' 0" -> 15.5
'
Dim i As Integer, W As String, Temp As Double
  Temp = 0
  For i = 1 To 3
    'w = CutWord(DMS, DMS)
    Select Case Right(W, 1)
      Case "'"
        Temp = Temp + Val(W) / 60
      Case """"
        Temp = Temp + Val(W) / 3600
      Case Else
        Temp = Temp + Val(W)
    End Select
  Next i
  DMSStrToDeg = Temp
End Function

Function DMSToDeg(d As Double, M As Double, s As Double) As Double
'
' Converts separate Degree, Minute, Second values into decimal degrees.
' e.g. 15,30,0 -> 15.5
'
  DMSToDeg = d + M / 60 + s / 3600
End Function

Function GreatArcDistance(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double, Radius As Double) As Double
'
' Calculates the Great Arc (shortest) distance between 2 locations on the globe.
'
' Uses functions from Trigonometry
'
Dim X1 As Double, Y1 As Double, Z1 As Double, X2 As Double, Y2 As Double, Z2 As Double
Dim CosX As Double, ChordLen As Double
  LatLongToXYZ Lat1, Lon1, Radius, X1, Y1, Z1
  LatLongToXYZ Lat2, Lon2, Radius, X2, Y2, Z2
  ChordLen = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2) + (Z1 - Z2) * (Z1 - Z2))
  CosX = 1 - ChordLen * ChordLen / (2 * Radius * Radius)
  Debug.Print X1, Y1, Z1
  Debug.Print X2, Y2, Z2
  Debug.Print ChordLen, CosX
  If CosX = 1 Or CosX = -1 Then
    GreatArcDistance = 0
  Else
    GreatArcDistance = Sqr(1 - CosX * CosX) * Radius * pi() / 2
  End If
End Function

Sub LatLongToXYZ(Lat As Double, Lon As Double, Radius As Double, x As Double, y As Double, z As Double)
'
' Converts Latitude, Longitude, Radius to 3d-Cartesian coordinates
'
' Assumes:
'   X axis runs through 270 (-X) and 90 (+X) Latitude
'   Y axis runs North (+Y) to South (-Y)
'   Z axis runs through 0 (-Z) and 180 (+Z) Latitude
'
  y = Radius * Sin(Deg2Rad(Lat))
  x = Radius * Sin(Deg2Rad(Lon)) * Cos(Deg2Rad(Lat))
  z = -Radius * Cos(Deg2Rad(Lon)) * Cos(Deg2Rad(Lat))
End Sub

Sub testxyz()
'
' Procedure to test the LatLongToXYZ and XYZToLatLong functions
'
  Dim Lat As Double, Lon As Double, Radius As Double
  Dim x As Double, y As Double, z As Double
  
  Lat = 90
  Lon = 50
  Radius = 1000
  LatLongToXYZ Lat, Lon, Radius, x, y, z
  Debug.Print Lat; Lon; Radius; x; y; z
  XYZToLatLong x, y, z, Lat, Lon, Radius
  Debug.Print Lat; Lon; Radius

End Sub

Sub XYZToLatLong(x As Double, y As Double, z As Double, Lat As Double, Lon As Double, Radius As Double)
'
' Converts 3d-Cartesian coordinates to Latitude, Longitude, and Radius
'
' Assumes:
'   X axis runs through 270 (-X) and 90 (+X) Latitude
'   Y axis runs North (+Y) to South (-Y)
'   Z axis runs through 0 (-Z) and 180 (+Z) Latitude
'
' Uses functions from the Trigonometry module
'
  Radius = Sqr(x * x + y * y + z * z)
  If Abs(Radius) < 0.0000000001 Then Radius = 0  ' Accomodate round-off error
  If Radius = 0 Then                             ' Zero radius has no other coordinates
    Radius = 0
    Lat = 0
    Lon = 0
  Else
    Lat = Rad2Deg(ArcSin(y / Radius))
    If (Lat Mod 90) = 0 Then                     ' North/South pole has no longitude
      Lon = 0
    Else
      Lon = Rad2Deg(ATan2(-z / Cos(Deg2Rad(Lat)), x / Cos(Deg2Rad(Lat))))
    End If
  End If
End Sub
