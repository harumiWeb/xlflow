Attribute VB_Name = "QrCode"
Option Explicit

Private Const MODE_BYTE As Long = 4
Private Const ERROR_CORRECTION_LEVEL_FORMAT As Long = 1

Private expTable(0 To 511) As Long
Private logTable(0 To 255) As Long
Private fieldInitialized As Boolean

Public Function BuildMatrix(ByVal text As String) As Variant
  Dim dataBytes() As Byte
  Dim version As Long
  Dim dataCodewords() As Long
  Dim eccCodewords() As Long
  Dim allCodewords() As Long
  Dim baseMatrix As Variant
  Dim reserved As Variant
  Dim bestMask As Long

  dataBytes = Utf8Bytes(text)
  version = ChooseVersion(CountBytes(dataBytes))
  dataCodewords = CreateDataCodewords(dataBytes, version)
  eccCodewords = CreateErrorCorrection(dataCodewords, ErrorCorrectionCodewordCount(version))
  allCodewords = JoinCodewords(dataCodewords, eccCodewords)

  PlaceBasePatterns version, baseMatrix, reserved
  PlaceDataBits baseMatrix, reserved, allCodewords
  bestMask = SelectBestMask(baseMatrix, reserved, version)
  BuildMatrix = ApplyMaskAndFormat(baseMatrix, reserved, version, bestMask)
End Function

Private Function Utf8Bytes(ByVal text As String) As Byte()
  Dim stream As Object
  Dim binaryData As Variant
  Dim bytes() As Byte
  Dim index As Long

  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText text
  stream.Position = 0
  stream.Type = 1
  stream.Position = 3
  binaryData = stream.Read
  stream.Close

  bytes = binaryData
  If CountBytes(bytes) = 0 Then
    ReDim bytes(0 To -1)
  Else
    For index = LBound(bytes) To UBound(bytes)
      bytes(index) = bytes(index) And &HFF
    Next index
  End If
  Utf8Bytes = bytes
End Function

Private Function CountBytes(ByRef bytes() As Byte) As Long
  On Error GoTo EmptyBytes
  CountBytes = UBound(bytes) - LBound(bytes) + 1
  Exit Function

EmptyBytes:
  CountBytes = 0
End Function

Private Function ChooseVersion(ByVal byteCount As Long) As Long
  If byteCount <= 17 Then
    ChooseVersion = 1
  ElseIf byteCount <= 32 Then
    ChooseVersion = 2
  ElseIf byteCount <= 53 Then
    ChooseVersion = 3
  ElseIf byteCount <= 78 Then
    ChooseVersion = 4
  Else
    Err.Raise vbObjectError + 1000, "QrCode.ChooseVersion", "Input is too long for this demo. Keep UTF-8 payloads within 78 bytes."
  End If
End Function

Private Function CreateDataCodewords(ByRef bytes() As Byte, ByVal version As Long) As Long()
  Dim dataCapacity As Long
  Dim bitStream As String
  Dim index As Long
  Dim bitLength As Long
  Dim codewords() As Long
  Dim padToggle As Boolean

  dataCapacity = DataCodewordCount(version)
  bitStream = BitsFromValue(MODE_BYTE, 4)
  bitStream = bitStream & BitsFromValue(CountBytes(bytes), 8)

  For index = LBound(bytes) To UBound(bytes)
    bitStream = bitStream & BitsFromValue(bytes(index), 8)
  Next index

  bitLength = dataCapacity * 8
  If Len(bitStream) > bitLength Then
    Err.Raise vbObjectError + 1001, "QrCode.CreateDataCodewords", "Encoded payload exceeded the selected QR version capacity."
  End If

  bitStream = bitStream & String$(MinLong(4, bitLength - Len(bitStream)), "0")
  Do While (Len(bitStream) Mod 8) <> 0
    bitStream = bitStream & "0"
  Loop

  ReDim codewords(0 To dataCapacity - 1)
  For index = 0 To (Len(bitStream) \ 8) - 1
    codewords(index) = ByteFromBits(Mid$(bitStream, index * 8 + 1, 8))
  Next index

  padToggle = True
  For index = (Len(bitStream) \ 8) To dataCapacity - 1
    If padToggle Then
      codewords(index) = &HEC
    Else
      codewords(index) = &H11
    End If
    padToggle = Not padToggle
  Next index

  CreateDataCodewords = codewords
End Function

Private Function DataCodewordCount(ByVal version As Long) As Long
  Select Case version
    Case 1
      DataCodewordCount = 19
    Case 2
      DataCodewordCount = 34
    Case 3
      DataCodewordCount = 55
    Case 4
      DataCodewordCount = 80
    Case Else
      Err.Raise vbObjectError + 1002, "QrCode.DataCodewordCount", "Unsupported version."
  End Select
End Function

Private Function ErrorCorrectionCodewordCount(ByVal version As Long) As Long
  Select Case version
    Case 1
      ErrorCorrectionCodewordCount = 7
    Case 2
      ErrorCorrectionCodewordCount = 10
    Case 3
      ErrorCorrectionCodewordCount = 15
    Case 4
      ErrorCorrectionCodewordCount = 20
    Case Else
      Err.Raise vbObjectError + 1003, "QrCode.ErrorCorrectionCodewordCount", "Unsupported version."
  End Select
End Function

Private Function CreateErrorCorrection(ByRef dataCodewords() As Long, ByVal ecLength As Long) As Long()
  Dim generator() As Long
  Dim remainder() As Long
  Dim index As Long
  Dim innerIndex As Long
  Dim factor As Long

  EnsureGaloisField
  generator = BuildGeneratorPolynomial(ecLength)
  ReDim remainder(0 To ecLength - 1)

  For index = LBound(dataCodewords) To UBound(dataCodewords)
    factor = dataCodewords(index) Xor remainder(0)
    For innerIndex = 0 To ecLength - 2
      remainder(innerIndex) = remainder(innerIndex + 1)
    Next innerIndex
    remainder(ecLength - 1) = 0
    If factor <> 0 Then
      For innerIndex = 0 To ecLength - 1
        remainder(innerIndex) = remainder(innerIndex) Xor GfMultiply(generator(innerIndex), factor)
      Next innerIndex
    End If
  Next index

  CreateErrorCorrection = remainder
End Function

Private Function BuildGeneratorPolynomial(ByVal ecLength As Long) As Long()
  Dim polynomial() As Long
  Dim term(0 To 1) As Long
  Dim index As Long

  ReDim polynomial(0 To 0)
  polynomial(0) = 1

  For index = 0 To ecLength - 1
    term(0) = 1
    term(1) = GfPower(index)
    polynomial = MultiplyPolynomials(polynomial, term)
  Next index

  BuildGeneratorPolynomial = SlicePolynomial(polynomial, 1)
End Function

Private Function MultiplyPolynomials(ByRef leftPoly() As Long, ByRef rightPoly() As Long) As Long()
  Dim result() As Long
  Dim leftIndex As Long
  Dim rightIndex As Long

  ReDim result(0 To UBound(leftPoly) + UBound(rightPoly))
  For leftIndex = LBound(leftPoly) To UBound(leftPoly)
    For rightIndex = LBound(rightPoly) To UBound(rightPoly)
      result(leftIndex + rightIndex) = result(leftIndex + rightIndex) Xor GfMultiply(leftPoly(leftIndex), rightPoly(rightIndex))
    Next rightIndex
  Next leftIndex

  MultiplyPolynomials = result
End Function

Private Function SlicePolynomial(ByRef values() As Long, ByVal startIndex As Long) As Long()
  Dim result() As Long
  Dim index As Long

  ReDim result(0 To UBound(values) - startIndex)
  For index = startIndex To UBound(values)
    result(index - startIndex) = values(index)
  Next index
  SlicePolynomial = result
End Function

Private Sub EnsureGaloisField()
  Dim index As Long
  Dim value As Long

  If fieldInitialized Then
    Exit Sub
  End If

  value = 1
  For index = 0 To 254
    expTable(index) = value
    logTable(value) = index
    value = value * 2
    If value >= 256 Then
      value = (value Xor &H11D) And &HFF
    End If
  Next index

  For index = 255 To 511
    expTable(index) = expTable(index - 255)
  Next index

  fieldInitialized = True
End Sub

Private Function GfPower(ByVal exponent As Long) As Long
  EnsureGaloisField
  GfPower = expTable(exponent Mod 255)
End Function

Private Function GfMultiply(ByVal leftValue As Long, ByVal rightValue As Long) As Long
  If leftValue = 0 Or rightValue = 0 Then
    GfMultiply = 0
  Else
    GfMultiply = expTable(logTable(leftValue) + logTable(rightValue))
  End If
End Function

Private Function JoinCodewords(ByRef leftCodewords() As Long, ByRef rightCodewords() As Long) As Long()
  Dim result() As Long
  Dim index As Long
  Dim offset As Long

  ReDim result(0 To UBound(leftCodewords) + UBound(rightCodewords) + 1)
  offset = 0
  For index = LBound(leftCodewords) To UBound(leftCodewords)
    result(offset) = leftCodewords(index)
    offset = offset + 1
  Next index
  For index = LBound(rightCodewords) To UBound(rightCodewords)
    result(offset) = rightCodewords(index)
    offset = offset + 1
  Next index
  JoinCodewords = result
End Function

Private Sub PlaceBasePatterns(ByVal version As Long, ByRef matrix As Variant, ByRef reserved As Variant)
  Dim size As Long
  Dim alignmentCenters As Variant
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim centerRow As Long
  Dim centerColumn As Long

  size = 17 + version * 4
  ReDim matrix(0 To size - 1, 0 To size - 1) As Long
  ReDim reserved(0 To size - 1, 0 To size - 1) As Boolean

  PlaceFinder matrix, reserved, 0, 0
  PlaceFinder matrix, reserved, 0, size - 7
  PlaceFinder matrix, reserved, size - 7, 0

  For columnIndex = 0 To size - 1
    If Not reserved(6, columnIndex) Then
      matrix(6, columnIndex) = (columnIndex + 1) Mod 2
      reserved(6, columnIndex) = True
    End If
  Next columnIndex

  For rowIndex = 0 To size - 1
    If Not reserved(rowIndex, 6) Then
      matrix(rowIndex, 6) = (rowIndex + 1) Mod 2
      reserved(rowIndex, 6) = True
    End If
  Next rowIndex

  alignmentCenters = AlignmentCentersForVersion(version)
  If Not IsEmpty(alignmentCenters) Then
    For rowIndex = LBound(alignmentCenters) To UBound(alignmentCenters)
      centerRow = alignmentCenters(rowIndex)
      For columnIndex = LBound(alignmentCenters) To UBound(alignmentCenters)
        centerColumn = alignmentCenters(columnIndex)
        If Not reserved(centerRow, centerColumn) Then
          PlaceAlignment matrix, reserved, centerRow - 2, centerColumn - 2
        End If
      Next columnIndex
    Next rowIndex
  End If

  ReserveFormatAreas reserved, size
  matrix(size - 8, 8) = 1
  reserved(size - 8, 8) = True
End Sub

Private Sub PlaceFinder(ByRef matrix As Variant, ByRef reserved As Variant, ByVal topRow As Long, ByVal leftColumn As Long)
  Dim rowOffset As Long
  Dim columnOffset As Long
  Dim actualRow As Long
  Dim actualColumn As Long

  For rowOffset = -1 To 7
    For columnOffset = -1 To 7
      actualRow = topRow + rowOffset
      actualColumn = leftColumn + columnOffset
      If actualRow >= 0 And actualColumn >= 0 And actualRow <= UBound(matrix, 1) And actualColumn <= UBound(matrix, 2) Then
        reserved(actualRow, actualColumn) = True
        If rowOffset = -1 Or rowOffset = 7 Or columnOffset = -1 Or columnOffset = 7 Then
          matrix(actualRow, actualColumn) = 0
        ElseIf rowOffset = 0 Or rowOffset = 6 Or columnOffset = 0 Or columnOffset = 6 Then
          matrix(actualRow, actualColumn) = 1
        ElseIf rowOffset = 1 Or rowOffset = 5 Or columnOffset = 1 Or columnOffset = 5 Then
          matrix(actualRow, actualColumn) = 0
        Else
          matrix(actualRow, actualColumn) = 1
        End If
      End If
    Next columnOffset
  Next rowOffset
End Sub

Private Sub PlaceAlignment(ByRef matrix As Variant, ByRef reserved As Variant, ByVal topRow As Long, ByVal leftColumn As Long)
  Dim rowOffset As Long
  Dim columnOffset As Long
  Dim actualRow As Long
  Dim actualColumn As Long

  For rowOffset = 0 To 4
    For columnOffset = 0 To 4
      actualRow = topRow + rowOffset
      actualColumn = leftColumn + columnOffset
      reserved(actualRow, actualColumn) = True
      If rowOffset = 0 Or rowOffset = 4 Or columnOffset = 0 Or columnOffset = 4 Then
        matrix(actualRow, actualColumn) = 1
      ElseIf rowOffset = 2 And columnOffset = 2 Then
        matrix(actualRow, actualColumn) = 1
      Else
        matrix(actualRow, actualColumn) = 0
      End If
    Next columnOffset
  Next rowOffset
End Sub

Private Function AlignmentCentersForVersion(ByVal version As Long) As Variant
  Dim centers() As Long

  Select Case version
    Case 1
      AlignmentCentersForVersion = Empty
    Case 2
      ReDim centers(0 To 1)
      centers(0) = 6
      centers(1) = 18
      AlignmentCentersForVersion = centers
    Case 3
      ReDim centers(0 To 1)
      centers(0) = 6
      centers(1) = 22
      AlignmentCentersForVersion = centers
    Case 4
      ReDim centers(0 To 1)
      centers(0) = 6
      centers(1) = 26
      AlignmentCentersForVersion = centers
    Case Else
      Err.Raise vbObjectError + 1004, "QrCode.AlignmentCentersForVersion", "Unsupported version."
  End Select
End Function

Private Sub ReserveFormatAreas(ByRef reserved As Variant, ByVal size As Long)
  Dim index As Long

  For index = 0 To 8
    If index <> 6 Then
      reserved(8, index) = True
      reserved(index, 8) = True
    End If
  Next index

  For index = 0 To 7
    reserved(8, size - 1 - index) = True
  Next index

  For index = 0 To 6
    reserved(size - 1 - index, 8) = True
  Next index
End Sub

Private Sub PlaceDataBits(ByRef matrix As Variant, ByRef reserved As Variant, ByRef codewords() As Long)
  Dim size As Long
  Dim rightColumn As Long
  Dim rowStep As Long
  Dim rowIndex As Long
  Dim currentRow As Long
  Dim currentColumn As Long
  Dim pairOffset As Long
  Dim bitIndex As Long
  Dim totalBits As Long

  size = UBound(matrix, 1) + 1
  totalBits = (UBound(codewords) - LBound(codewords) + 1) * 8
  bitIndex = 0

  For rightColumn = size - 1 To 1 Step -2
    If rightColumn = 6 Then
      rightColumn = 5
    End If

    If (((size - 1) - rightColumn) \ 2) Mod 2 = 0 Then
      rowStep = -1
      rowIndex = size - 1
    Else
      rowStep = 1
      rowIndex = 0
    End If

    Do While rowIndex >= 0 And rowIndex < size
      For pairOffset = 0 To 1
        currentColumn = rightColumn - pairOffset
        currentRow = rowIndex
        If Not reserved(currentRow, currentColumn) Then
          If bitIndex < totalBits Then
            matrix(currentRow, currentColumn) = BitAt(codewords, bitIndex)
            bitIndex = bitIndex + 1
          Else
            matrix(currentRow, currentColumn) = 0
          End If
        End If
      Next pairOffset
      rowIndex = rowIndex + rowStep
    Loop
  Next rightColumn
End Sub

Private Function BitAt(ByRef codewords() As Long, ByVal bitIndex As Long) As Long
  Dim byteIndex As Long
  Dim shiftCount As Long

  byteIndex = bitIndex \ 8
  shiftCount = 7 - (bitIndex Mod 8)
  BitAt = (codewords(byteIndex) \ (2 ^ shiftCount)) And 1
End Function

Private Function SelectBestMask(ByRef baseMatrix As Variant, ByRef reserved As Variant, ByVal version As Long) As Long
  Dim maskIndex As Long
  Dim bestPenalty As Long
  Dim currentPenalty As Long
  Dim firstMask As Boolean
  Dim trialMatrix As Variant

  firstMask = True
  For maskIndex = 0 To 7
    trialMatrix = ApplyMaskAndFormat(baseMatrix, reserved, version, maskIndex)
    currentPenalty = ScoreMatrix(trialMatrix)
    If firstMask Or currentPenalty < bestPenalty Then
      bestPenalty = currentPenalty
      SelectBestMask = maskIndex
      firstMask = False
    End If
  Next maskIndex
End Function

Private Function ApplyMaskAndFormat(ByRef baseMatrix As Variant, ByRef reserved As Variant, ByVal version As Long, ByVal maskIndex As Long) As Variant
  Dim matrix As Variant
  Dim rowIndex As Long
  Dim columnIndex As Long

  matrix = CopyMatrix(baseMatrix)
  For rowIndex = 0 To UBound(matrix, 1)
    For columnIndex = 0 To UBound(matrix, 2)
      If Not reserved(rowIndex, columnIndex) Then
        If MaskBit(maskIndex, rowIndex, columnIndex) Then
          matrix(rowIndex, columnIndex) = matrix(rowIndex, columnIndex) Xor 1
        End If
      End If
    Next columnIndex
  Next rowIndex

  PlaceFormatBits matrix, version, maskIndex
  ApplyMaskAndFormat = matrix
End Function

Private Function CopyMatrix(ByRef source As Variant) As Variant
  Dim Target As Variant
  Dim rowIndex As Long
  Dim columnIndex As Long

  ReDim Target(LBound(source, 1) To UBound(source, 1), LBound(source, 2) To UBound(source, 2)) As Long
  For rowIndex = LBound(source, 1) To UBound(source, 1)
    For columnIndex = LBound(source, 2) To UBound(source, 2)
      Target(rowIndex, columnIndex) = source(rowIndex, columnIndex)
    Next columnIndex
  Next rowIndex

  CopyMatrix = Target
End Function

Private Function MaskBit(ByVal maskIndex As Long, ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
  Select Case maskIndex
    Case 0
      MaskBit = ((rowIndex + columnIndex) Mod 2) = 0
    Case 1
      MaskBit = (rowIndex Mod 2) = 0
    Case 2
      MaskBit = (columnIndex Mod 3) = 0
    Case 3
      MaskBit = ((rowIndex + columnIndex) Mod 3) = 0
    Case 4
      MaskBit = (((rowIndex \ 2) + (columnIndex \ 3)) Mod 2) = 0
    Case 5
      MaskBit = (((rowIndex * columnIndex) Mod 2) + ((rowIndex * columnIndex) Mod 3)) = 0
    Case 6
      MaskBit = ((((rowIndex * columnIndex) Mod 2) + ((rowIndex * columnIndex) Mod 3)) Mod 2) = 0
    Case 7
      MaskBit = ((((rowIndex + columnIndex) Mod 2) + ((rowIndex * columnIndex) Mod 3)) Mod 2) = 0
    Case Else
      Err.Raise vbObjectError + 1005, "QrCode.MaskBit", "Unsupported mask index."
  End Select
End Function

Private Sub PlaceFormatBits(ByRef matrix As Variant, ByVal version As Long, ByVal maskIndex As Long)
  Dim bits As Long
  Dim size As Long
  Dim index As Long

  size = 17 + version * 4
  bits = FormatBits(maskIndex)

  For index = 0 To 5
    matrix(index, 8) = GetBit(bits, index)
  Next index
  matrix(7, 8) = GetBit(bits, 6)
  matrix(8, 8) = GetBit(bits, 7)
  matrix(8, 7) = GetBit(bits, 8)
  For index = 9 To 14
    matrix(8, 14 - index) = GetBit(bits, index)
  Next index

  For index = 0 To 7
    matrix(8, size - 1 - index) = GetBit(bits, index)
  Next index
  For index = 8 To 14
    matrix(size - 15 + index, 8) = GetBit(bits, index)
  Next index

  matrix(size - 8, 8) = 1
End Sub

Private Function FormatBits(ByVal maskIndex As Long) As Long
  Dim dataValue As Long
  Dim remainder As Long
  Dim bitIndex As Long

  dataValue = (ERROR_CORRECTION_LEVEL_FORMAT * 8) Or maskIndex
  remainder = dataValue
  For bitIndex = 0 To 9
    remainder = remainder * 2
    If (remainder And &H400) <> 0 Then
      remainder = remainder Xor &H537
    End If
  Next bitIndex

  FormatBits = ((dataValue * 1024) Or (remainder And &H3FF)) Xor &H5412
End Function

Private Function GetBit(ByVal value As Long, ByVal bitIndex As Long) As Long
  GetBit = (value \ (2 ^ bitIndex)) And 1
End Function

Private Function ScoreMatrix(ByRef matrix As Variant) As Long
  ScoreMatrix = ScoreRuns(matrix) + ScoreBlocks(matrix) + ScoreFinders(matrix) + ScoreDarkBalance(matrix)
End Function

Private Function ScoreRuns(ByRef matrix As Variant) As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim runColor As Long
  Dim runLength As Long

  For rowIndex = 0 To UBound(matrix, 1)
    runColor = matrix(rowIndex, 0)
    runLength = 1
    For columnIndex = 1 To UBound(matrix, 2)
      If matrix(rowIndex, columnIndex) = runColor Then
        runLength = runLength + 1
      Else
        ScoreRuns = ScoreRuns + PenaltyForRun(runLength)
        runColor = matrix(rowIndex, columnIndex)
        runLength = 1
      End If
    Next columnIndex
    ScoreRuns = ScoreRuns + PenaltyForRun(runLength)
  Next rowIndex

  For columnIndex = 0 To UBound(matrix, 2)
    runColor = matrix(0, columnIndex)
    runLength = 1
    For rowIndex = 1 To UBound(matrix, 1)
      If matrix(rowIndex, columnIndex) = runColor Then
        runLength = runLength + 1
      Else
        ScoreRuns = ScoreRuns + PenaltyForRun(runLength)
        runColor = matrix(rowIndex, columnIndex)
        runLength = 1
      End If
    Next rowIndex
    ScoreRuns = ScoreRuns + PenaltyForRun(runLength)
  Next columnIndex
End Function

Private Function PenaltyForRun(ByVal runLength As Long) As Long
  If runLength >= 5 Then
    PenaltyForRun = 3 + (runLength - 5)
  Else
    PenaltyForRun = 0
  End If
End Function

Private Function ScoreBlocks(ByRef matrix As Variant) As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim colorValue As Long

  For rowIndex = 0 To UBound(matrix, 1) - 1
    For columnIndex = 0 To UBound(matrix, 2) - 1
      colorValue = matrix(rowIndex, columnIndex)
      If colorValue = matrix(rowIndex + 1, columnIndex) _
        And colorValue = matrix(rowIndex, columnIndex + 1) _
        And colorValue = matrix(rowIndex + 1, columnIndex + 1) Then
        ScoreBlocks = ScoreBlocks + 3
      End If
    Next columnIndex
  Next rowIndex
End Function

Private Function ScoreFinders(ByRef matrix As Variant) As Long
  Dim rowIndex As Long
  Dim columnIndex As Long

  For rowIndex = 0 To UBound(matrix, 1)
    For columnIndex = 0 To UBound(matrix, 2) - 10
      If MatchesFinderPattern(matrix, rowIndex, columnIndex, True) Then
        ScoreFinders = ScoreFinders + 40
      End If
    Next columnIndex
  Next rowIndex

  For columnIndex = 0 To UBound(matrix, 2)
    For rowIndex = 0 To UBound(matrix, 1) - 10
      If MatchesFinderPattern(matrix, rowIndex, columnIndex, False) Then
        ScoreFinders = ScoreFinders + 40
      End If
    Next rowIndex
  Next columnIndex
End Function

Private Function MatchesFinderPattern(ByRef matrix As Variant, ByVal startRow As Long, ByVal startColumn As Long, ByVal horizontal As Boolean) As Boolean
  Dim values(0 To 10) As Long
  Dim index As Long

  For index = 0 To 10
    If horizontal Then
      values(index) = matrix(startRow, startColumn + index)
    Else
      values(index) = matrix(startRow + index, startColumn)
    End If
  Next index

  MatchesFinderPattern = SequenceMatches(values, "10111010000") Or SequenceMatches(values, "00001011101")
End Function

Private Function SequenceMatches(ByRef values() As Long, ByVal pattern As String) As Boolean
  Dim index As Long

  SequenceMatches = True
  For index = 0 To UBound(values)
    If values(index) <> CLng(Mid$(pattern, index + 1, 1)) Then
      SequenceMatches = False
      Exit Function
    End If
  Next index
End Function

Private Function ScoreDarkBalance(ByRef matrix As Variant) As Long
  Dim darkCount As Long
  Dim totalCount As Long
  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim percent As Long
  Dim deviation As Long

  totalCount = (UBound(matrix, 1) + 1) * (UBound(matrix, 2) + 1)
  For rowIndex = 0 To UBound(matrix, 1)
    For columnIndex = 0 To UBound(matrix, 2)
      If matrix(rowIndex, columnIndex) = 1 Then
        darkCount = darkCount + 1
      End If
    Next columnIndex
  Next rowIndex

  percent = (darkCount * 100) \ totalCount
  deviation = Abs(percent - 50) \ 5
  ScoreDarkBalance = deviation * 10
End Function

Private Function BitsFromValue(ByVal value As Long, ByVal bitCount As Long) As String
  Dim index As Long

  For index = bitCount - 1 To 0 Step -1
    If ((value \ (2 ^ index)) And 1) = 1 Then
      BitsFromValue = BitsFromValue & "1"
    Else
      BitsFromValue = BitsFromValue & "0"
    End If
  Next index
End Function

Private Function ByteFromBits(ByVal bits As String) As Long
  Dim index As Long

  For index = 1 To Len(bits)
    ByteFromBits = ByteFromBits * 2
    If Mid$(bits, index, 1) = "1" Then
      ByteFromBits = ByteFromBits + 1
    End If
  Next index
End Function

Private Function MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
  If leftValue < rightValue Then
    MinLong = leftValue
  Else
    MinLong = rightValue
  End If
End Function
