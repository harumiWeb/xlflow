Attribute VB_Name = "CsvUtil"
Option Explicit

Public Function ParseCsvText(ByVal csvText As String) As Variant
  Dim normalized As String
  Dim lines() As String
  Dim lineIndex As Long
  Dim delimiter As String

  normalized = Replace(csvText, vbCrLf, vbLf)
  normalized = Replace(normalized, vbCr, vbLf)
  If Len(normalized) > 0 Then
    If Left$(normalized, 1) = ChrW$(&HFEFF) Then
      normalized = Mid$(normalized, 2)
    End If
  End If

  ValidateCsvInput normalized

  lines = Split(normalized, vbLf)
  delimiter = DetectCsvDelimiter(lines)

  If InStr(normalized, """") = 0 Then
    ParseCsvText = ParsePlainCsvLines(lines, delimiter)
    Exit Function
  End If

  ParseCsvText = ParseQuotedCsvLines(lines, delimiter)
End Function

Private Function ParsePlainCsvLines(ByRef lines() As String, ByVal delimiter As String) As Variant
  Dim rowCount As Long
  Dim columnCount As Long
  Dim lineIndex As Long
  Dim rowIndex As Long
  Dim fields() As String
  Dim output() As Variant
  Dim columnIndex As Long
  Dim currentLine As String

  rowCount = 0
  columnCount = 0

  For lineIndex = LBound(lines) To UBound(lines)
    currentLine = Trim$(lines(lineIndex))
    If Len(currentLine) > 0 Then
      fields = Split(lines(lineIndex), delimiter)
      rowCount = rowCount + 1
      If (UBound(fields) - LBound(fields) + 1) > columnCount Then
        columnCount = UBound(fields) - LBound(fields) + 1
      End If
    End If
  Next lineIndex

  If rowCount = 0 Then
    Err.Raise vbObjectError + 801, "CsvUtil.ParsePlainCsvLines", "CSV data is empty."
  End If

  ReDim output(1 To rowCount, 1 To columnCount)
  rowIndex = 0

  For lineIndex = LBound(lines) To UBound(lines)
    If Len(Trim$(lines(lineIndex))) > 0 Then
      fields = Split(lines(lineIndex), delimiter)
      rowIndex = rowIndex + 1
      For columnIndex = LBound(fields) To UBound(fields)
        output(rowIndex, columnIndex - LBound(fields) + 1) = fields(columnIndex)
      Next columnIndex
    End If
  Next lineIndex

  ParsePlainCsvLines = output
End Function

Private Function ParseQuotedCsvLines(ByRef lines() As String, ByVal delimiter As String) As Variant
  Dim parsedRows As Collection
  Dim lineIndex As Long
  Dim maxColumns As Long
  Dim rowCount As Long
  Dim rowFields As Variant
  Dim output() As Variant
  Dim columnIndex As Long

  Set parsedRows = New Collection

  For lineIndex = LBound(lines) To UBound(lines)
    If Len(Trim$(lines(lineIndex))) > 0 Then
      rowFields = ParseCsvLine(lines(lineIndex), delimiter)
      parsedRows.Add rowFields
      If (UBound(rowFields) + 1) > maxColumns Then
        maxColumns = UBound(rowFields) + 1
      End If
    End If
  Next lineIndex

  rowCount = parsedRows.Count
  If rowCount = 0 Then
    Err.Raise vbObjectError + 801, "CsvUtil.ParseQuotedCsvLines", "CSV data is empty."
  End If

  ReDim output(1 To rowCount, 1 To maxColumns)
  For lineIndex = 1 To rowCount
    rowFields = parsedRows(lineIndex)
    For columnIndex = 0 To UBound(rowFields)
      output(lineIndex, columnIndex + 1) = rowFields(columnIndex)
    Next columnIndex
  Next lineIndex

  ParseQuotedCsvLines = output
End Function

Public Function CsvHeaderIndex(ByVal csvTable As Variant, ByVal headerName As String) As Long
  Dim columnIndex As Long

  For columnIndex = LBound(csvTable, 2) To UBound(csvTable, 2)
    If StrComp(CStr(csvTable(1, columnIndex)), headerName, vbTextCompare) = 0 Then
      CsvHeaderIndex = columnIndex
      Exit Function
    End If
  Next columnIndex

  Err.Raise vbObjectError + 802, "CsvUtil.CsvHeaderIndex", "Header not found: " & headerName
End Function

Private Sub ValidateCsvInput(ByVal normalized As String)
  Dim trimmed As String
  Dim preview As String

  trimmed = LTrim$(normalized)
  If Len(trimmed) = 0 Then
    Exit Sub
  End If

  If Left$(trimmed, 1) = "{" Or Left$(trimmed, 1) = "[" Then
    preview = Left$(Replace(trimmed, vbLf, " "), 240)
    Err.Raise vbObjectError + 804, "CsvUtil.ValidateCsvInput", "Expected CSV but received non-CSV response: " & preview
  End If
End Sub

Private Function DetectCsvDelimiter(ByRef lines() As String) As String
  Dim lineIndex As Long
  Dim candidate As String

  For lineIndex = LBound(lines) To UBound(lines)
    candidate = Trim$(lines(lineIndex))
    If Len(candidate) > 0 Then
      If InStr(candidate, ";") > 0 Then
        DetectCsvDelimiter = ";"
      Else
        DetectCsvDelimiter = ","
      End If
      Exit Function
    End If
  Next lineIndex

  DetectCsvDelimiter = ","
End Function

Private Function ParseCsvLine(ByVal line As String, ByVal delimiter As String) As Variant
  Dim fields() As String
  Dim fieldCount As Long
  Dim currentValue As String
  Dim inQuotes As Boolean
  Dim charIndex As Long
  Dim currentChar As String

  ReDim fields(0 To 0)

  For charIndex = 1 To Len(line)
    currentChar = Mid$(line, charIndex, 1)
    If inQuotes Then
      If currentChar = """" Then
        If charIndex < Len(line) And Mid$(line, charIndex + 1, 1) = """" Then
          currentValue = currentValue & """"
          charIndex = charIndex + 1
        Else
          inQuotes = False
        End If
      Else
        currentValue = currentValue & currentChar
      End If
    Else
      Select Case currentChar
        Case delimiter
          fields(fieldCount) = currentValue
          fieldCount = fieldCount + 1
          ReDim Preserve fields(0 To fieldCount)
          currentValue = vbNullString
        Case """"
          inQuotes = True
        Case Else
          currentValue = currentValue & currentChar
      End Select
    End If
  Next charIndex

  fields(fieldCount) = currentValue
  ParseCsvLine = fields
End Function
