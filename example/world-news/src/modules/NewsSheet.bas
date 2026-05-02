Attribute VB_Name = "NewsSheet"
Option Explicit

Private Const NEWS_SHEET_NAME As String = "News"
Private Const HEADER_ROW As Long = 5
Private Const FIRST_DATA_ROW As Long = 6
Private Const MAX_ARTICLE_COUNT As Long = 15
Private Const ARTICLE_ROW_HEIGHT As Double = 108
Private Const IMAGE_SHAPE_PREFIX As String = "NewsImage_"
Private Const IMAGE_CACHE_FOLDER_NAME As String = "xlflow_world_news_images"
Private Const IMAGE_PADDING As Double = 4

Public Sub RenderArticles(ByVal wb As Workbook, ByVal newsRoot As Object, Optional ByVal includeImages As Boolean = True)
  Dim ws As Worksheet
  Dim articles As Collection
  Dim previousScreenUpdating As Boolean

  previousScreenUpdating = Application.ScreenUpdating
  On Error GoTo ErrHandler

  Application.ScreenUpdating = False
  Set ws = GetOrCreateNewsWorksheet(wb)
  Set articles = GetChildArray(newsRoot, "articles")

  PrepareCanvas ws
  WriteSummary ws, articles.Count
  WriteHeaders ws
  WriteArticleRows ws, articles, includeImages

CleanExit:
  Application.ScreenUpdating = previousScreenUpdating
  Exit Sub

ErrHandler:
  Application.ScreenUpdating = previousScreenUpdating
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub RenderFailure(ByVal wb As Workbook, ByVal sourceName As String, ByVal message As String)
  Dim ws As Worksheet
  Dim previousScreenUpdating As Boolean

  previousScreenUpdating = Application.ScreenUpdating
  On Error GoTo ErrHandler

  Application.ScreenUpdating = False
  Set ws = GetOrCreateNewsWorksheet(wb)

  PrepareCanvas ws
  With ws.Range("A1:F1")
    .Merge
    .Value = "World News"
    .Font.Bold = True
    .Font.Size = 18
    .HorizontalAlignment = xlLeft
  End With

  ws.Range("A2").Value = "Refresh failed: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
  ws.Range("A3").Value = "Source: " & sourceName
  ws.Range("A4").Value = message
  ws.Range("A6").Value = "Set NEWSAPI_KEY in the environment and reopen the workbook or run Main.Run."

  With ws.Range("A2:A6")
    .Font.Color = RGB(156, 0, 6)
    .WrapText = True
  End With

  ws.Range("A4:F4").Merge
  ws.Range("A4:F4").Interior.Color = RGB(255, 235, 235)
  ws.Range("A4:F4").Borders.LineStyle = xlContinuous
  ws.Range("A4:F4").Borders.Weight = xlThin

CleanExit:
  Application.ScreenUpdating = previousScreenUpdating
  Exit Sub

ErrHandler:
  Application.ScreenUpdating = previousScreenUpdating
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PrepareCanvas(ByVal ws As Worksheet)
  DeleteManagedImages ws
  ws.Cells.UnMerge
  ws.Cells.Clear
  ws.Hyperlinks.Delete
  ClearImageCacheFolder

  ws.Columns("A").ColumnWidth = 22
  ws.Columns("B").ColumnWidth = 20
  ws.Columns("C").ColumnWidth = 24
  ws.Columns("D").ColumnWidth = 42
  ws.Columns("E").ColumnWidth = 56
  ws.Columns("F").ColumnWidth = 16

  ws.Rows("1:3").RowHeight = 22
  ws.Rows(HEADER_ROW).RowHeight = 24

  With ws.Cells.Font
    .Name = "Calibri"
    .Size = 11
  End With
End Sub

Private Sub WriteSummary(ByVal ws As Worksheet, ByVal articleCount As Long)
  With ws.Range("A1:F1")
    .Merge
    .Value = "World News"
    .Font.Bold = True
    .Font.Size = 18
    .HorizontalAlignment = xlLeft
  End With

  ws.Range("A2").Value = "Latest stories from NewsAPI"
  ws.Range("A3").Value = "Updated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Query: world | Showing: " & CStr(MinLong(articleCount, MAX_ARTICLE_COUNT))
  ws.Range("A3:F3").Merge
  ws.Range("A3:F3").Interior.Color = RGB(242, 242, 242)
End Sub

Private Sub WriteHeaders(ByVal ws As Worksheet)
  With ws.Range("A" & CStr(HEADER_ROW) & ":F" & CStr(HEADER_ROW))
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(221, 235, 247)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
  End With

  ws.Cells(HEADER_ROW, 1).Value = "Published"
  ws.Cells(HEADER_ROW, 2).Value = "Source"
  ws.Cells(HEADER_ROW, 3).Value = "Image"
  ws.Cells(HEADER_ROW, 4).Value = "Title"
  ws.Cells(HEADER_ROW, 5).Value = "Description"
  ws.Cells(HEADER_ROW, 6).Value = "Link"
End Sub

Private Sub WriteArticleRows(ByVal ws As Worksheet, ByVal articles As Collection, ByVal includeImages As Boolean)
  Dim rowIndex As Long
  Dim articleIndex As Long
  Dim article As Object
  Dim urlValue As String
  Dim imageUrl As String

  If articles.Count = 0 Then
    ws.Cells(FIRST_DATA_ROW, 1).Value = "No articles returned."
    Exit Sub
  End If

  For articleIndex = 1 To MinLong(articles.Count, MAX_ARTICLE_COUNT)
    rowIndex = FIRST_DATA_ROW + articleIndex - 1
    Set article = articles.Item(articleIndex)
    urlValue = GetChildStringOrDefault(article, "url", "")
    imageUrl = GetChildStringOrDefault(article, "urlToImage", "")
    ws.Rows(rowIndex).RowHeight = ARTICLE_ROW_HEIGHT

    ws.Cells(rowIndex, 1).Value = NormalizeCellText(GetChildStringOrDefault(article, "publishedAt", ""))
    ws.Cells(rowIndex, 2).Value = SourceName(article)
    ws.Cells(rowIndex, 4).Value = NormalizeCellText(GetChildStringOrDefault(article, "title", "(untitled)"))
    ws.Cells(rowIndex, 5).Value = NormalizeCellText(GetChildStringOrDefault(article, "description", ""))
    WriteImageCell ws, ws.Cells(rowIndex, 3), imageUrl, includeImages

    If Len(urlValue) > 0 Then
      ws.Hyperlinks.Add Anchor:=ws.Cells(rowIndex, 6), Address:=urlValue, TextToDisplay:="Open article"
    Else
      ws.Cells(rowIndex, 6).Value = "-"
    End If

    With ws.Range("A" & CStr(rowIndex) & ":F" & CStr(rowIndex))
      .Borders.LineStyle = xlContinuous
      .Borders.Weight = xlThin
      .VerticalAlignment = xlTop
      .WrapText = True
    End With

    ws.Cells(rowIndex, 1).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 2).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 3).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 3).VerticalAlignment = xlCenter
    ws.Cells(rowIndex, 6).HorizontalAlignment = xlCenter
  Next articleIndex
End Sub

Private Sub WriteImageCell(ByVal ws As Worksheet, ByVal targetCell As Range, ByVal imageUrl As String, ByVal includeImages As Boolean)
  Dim imagePath As String
  Dim picture As Shape
  Dim maxWidth As Double
  Dim maxHeight As Double

  targetCell.ClearContents
  DeleteShapeIfExists ws, IMAGE_SHAPE_PREFIX & CStr(targetCell.Row)

  If Not includeImages Or Len(imageUrl) = 0 Then
    Exit Sub
  End If

  imagePath = DownloadTempImage(imageUrl, targetCell.Row)
  If Len(imagePath) = 0 Then
    Exit Sub
  End If

  On Error GoTo ImageLoadFailed
  Set picture = ws.Shapes.AddPicture(imagePath, 0, -1, targetCell.Left, targetCell.Top, -1, -1)
  picture.Name = IMAGE_SHAPE_PREFIX & CStr(targetCell.Row)
  picture.Placement = xlMoveAndSize
  picture.LockAspectRatio = -1

  maxWidth = targetCell.Width - (IMAGE_PADDING * 2)
  maxHeight = targetCell.Height - (IMAGE_PADDING * 2)
  If picture.Width <= 0 Or picture.Height <= 0 Or maxWidth <= 0 Or maxHeight <= 0 Then
    GoTo ImageLoadFailed
  End If

  picture.Width = maxWidth
  If picture.Height > maxHeight Then
    picture.Height = maxHeight
  End If

  picture.Left = targetCell.Left + ((targetCell.Width - picture.Width) / 2)
  picture.Top = targetCell.Top + ((targetCell.Height - picture.Height) / 2)
  Exit Sub

ImageLoadFailed:
  DeleteShapeIfExists ws, IMAGE_SHAPE_PREFIX & CStr(targetCell.Row)
End Sub

Private Function GetOrCreateNewsWorksheet(ByVal wb As Workbook) As Worksheet
  Dim ws As Worksheet

  For Each ws In wb.Worksheets
    If StrComp(ws.Name, NEWS_SHEET_NAME, vbTextCompare) = 0 Then
      Set GetOrCreateNewsWorksheet = ws
      Exit Function
    End If
  Next ws

  If GetOrCreateNewsWorksheet Is Nothing Then
    Set GetOrCreateNewsWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    GetOrCreateNewsWorksheet.Name = NEWS_SHEET_NAME
  End If
End Function

Private Function GetChildArray(ByVal parent As Object, ByVal key As String) As Collection
  If Not parent.Exists(key) Then
    Err.Raise vbObjectError + 2400, "NewsSheet.GetChildArray", "Missing JSON property '" & key & "'."
  End If

  If Not IsObject(parent(key)) Then
    Err.Raise vbObjectError + 2401, "NewsSheet.GetChildArray", "JSON property '" & key & "' is not an array."
  End If

  Set GetChildArray = parent(key)
End Function

Private Function GetChildObject(ByVal parent As Object, ByVal key As String) As Object
  If Not parent.Exists(key) Then
    Err.Raise vbObjectError + 2402, "NewsSheet.GetChildObject", "Missing JSON property '" & key & "'."
  End If

  If Not IsObject(parent(key)) Then
    Err.Raise vbObjectError + 2403, "NewsSheet.GetChildObject", "JSON property '" & key & "' is not an object."
  End If

  Set GetChildObject = parent(key)
End Function

Private Function GetChildStringOrDefault(ByVal parent As Object, ByVal key As String, ByVal defaultValue As String) As String
  If Not parent.Exists(key) Then
    GetChildStringOrDefault = defaultValue
  ElseIf IsNull(parent(key)) Then
    GetChildStringOrDefault = defaultValue
  Else
    GetChildStringOrDefault = CStr(parent(key))
  End If
End Function

Private Function SourceName(ByVal article As Object) As String
  Dim sourceNode As Object

  If Not article.Exists("source") Then
    SourceName = "-"
    Exit Function
  End If

  If Not IsObject(article("source")) Then
    SourceName = "-"
    Exit Function
  End If

  Set sourceNode = GetChildObject(article, "source")
  SourceName = GetChildStringOrDefault(sourceNode, "name", "-")
End Function

Private Function NormalizeCellText(ByVal value As String) As String
  NormalizeCellText = Replace(value, vbCrLf, vbLf)
  NormalizeCellText = Replace(NormalizeCellText, vbCr, vbLf)
End Function

Private Function DownloadTempImage(ByVal imageUrl As String, ByVal rowNumber As Long) As String
  Dim request As Object
  Dim stream As Object
  Dim streamOpened As Boolean
  Dim tempFolder As String
  Dim filePath As String
  Dim contentType As String

  On Error GoTo ErrHandler

  tempFolder = BuildImageCacheFolder()
  EnsureFolder tempFolder

  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
  request.SetTimeouts 5000, 5000, 15000, 15000
  request.Open "GET", imageUrl, False
  request.SetRequestHeader "User-Agent", "xlflow-world-news/1.0"
  request.Send

  If CLng(request.Status) <> 200 Then
    Exit Function
  End If

  contentType = LCase$(request.GetResponseHeader("Content-Type"))
  filePath = tempFolder & "\news_" & Format$(rowNumber, "00") & GetImageExtension(imageUrl, contentType)

  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 1
  stream.Open
  streamOpened = True
  stream.Write request.ResponseBody
  stream.SaveToFile filePath, 2

  DownloadTempImage = filePath

CleanExit:
  If streamOpened Then
    stream.Close
  End If
  Exit Function

ErrHandler:
  If streamOpened Then
    stream.Close
  End If
  If Len(filePath) > 0 And Len(Dir$(filePath)) > 0 Then
    Kill filePath
  End If
  DownloadTempImage = vbNullString
End Function

Private Function BuildImageCacheFolder() As String
  BuildImageCacheFolder = Environ$("TEMP") & "\" & IMAGE_CACHE_FOLDER_NAME
End Function

Private Sub ClearImageCacheFolder()
  Dim folderPath As String
  Dim fileName As String

  folderPath = BuildImageCacheFolder()
  If Len(Dir$(folderPath, vbDirectory)) = 0 Then
    Exit Sub
  End If

  fileName = Dir$(folderPath & "\*.*")
  Do While Len(fileName) > 0
    Kill folderPath & "\" & fileName
    fileName = Dir$
  Loop
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
  If Len(Dir$(folderPath, vbDirectory)) = 0 Then
    MkDir folderPath
  End If
End Sub

Private Function GetImageExtension(ByVal imageUrl As String, ByVal contentType As String) As String
  Dim separatorIndex As Long

  separatorIndex = InStr(contentType, ";")
  If separatorIndex > 0 Then
    contentType = Left$(contentType, separatorIndex - 1)
  End If

  Select Case contentType
    Case "image/jpeg", "image/jpg"
      GetImageExtension = ".jpg"
    Case "image/png"
      GetImageExtension = ".png"
    Case "image/gif"
      GetImageExtension = ".gif"
    Case "image/bmp"
      GetImageExtension = ".bmp"
    Case "image/webp"
      GetImageExtension = ".webp"
    Case Else
      GetImageExtension = GetFileExtensionFromUrl(imageUrl)
  End Select

  If Len(GetImageExtension) = 0 Then
    GetImageExtension = ".jpg"
  End If
End Function

Private Function GetFileExtensionFromUrl(ByVal imageUrl As String) As String
  Dim queryIndex As Long
  Dim dotIndex As Long
  Dim slashIndex As Long

  queryIndex = InStr(imageUrl, "?")
  If queryIndex > 0 Then
    imageUrl = Left$(imageUrl, queryIndex - 1)
  End If

  slashIndex = InStrRev(imageUrl, "/")
  dotIndex = InStrRev(imageUrl, ".")

  If dotIndex > slashIndex Then
    GetFileExtensionFromUrl = Mid$(imageUrl, dotIndex)
  End If
End Function

Private Sub DeleteManagedImages(ByVal ws As Worksheet)
  Dim shapeIndex As Long

  For shapeIndex = ws.Shapes.Count To 1 Step -1
    If InStr(1, ws.Shapes(shapeIndex).Name, IMAGE_SHAPE_PREFIX, vbTextCompare) = 1 Then
      ws.Shapes(shapeIndex).Delete
    End If
  Next shapeIndex
End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
  Dim shapeIndex As Long

  For shapeIndex = ws.Shapes.Count To 1 Step -1
    If ws.Shapes(shapeIndex).Name = shapeName Then
      ws.Shapes(shapeIndex).Delete
      Exit Sub
    End If
  Next shapeIndex
End Sub

Private Function MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
  If leftValue < rightValue Then
    MinLong = leftValue
  Else
    MinLong = rightValue
  End If
End Function

