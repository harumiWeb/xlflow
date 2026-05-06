Attribute VB_Name = "NewsSheetTests"
Option Explicit

Public Sub Test_RenderArticles_WritesRows()
  Dim root As Object
  Dim articles As Collection
  Dim article As Object
  Dim sourceNode As Object
  Dim ws As Worksheet

  Set root = CreateObject("Scripting.Dictionary")
  Set articles = New Collection
  Set article = CreateObject("Scripting.Dictionary")
  Set sourceNode = CreateObject("Scripting.Dictionary")

  sourceNode.Add "name", "Example Source"
  article.Add "source", sourceNode
  article.Add "publishedAt", "2026-05-02T12:00:00Z"
  article.Add "urlToImage", "https://example.com/story.jpg"
  article.Add "title", "Example headline"
  article.Add "description", "Example story summary"
  article.Add "url", "https://example.com/story"
  articles.Add article

  root.Add "status", "ok"
  root.Add "articles", articles

  NewsSheet.RenderArticles ThisWorkbook, root, False
  Set ws = ThisWorkbook.Worksheets("News")

  AssertEquals "World News", ws.Range("A1").Value2, "sheet title"
  AssertEquals "Published", ws.Cells(5, 1).Value2, "header label"
  AssertEquals "Image", ws.Cells(5, 3).Value2, "image header"
  AssertEquals "Example Source", ws.Cells(6, 2).Value2, "source value"
  AssertEquals "Example headline", ws.Cells(6, 4).Value2, "title value"
  AssertEquals "https://example.com/story", ws.Hyperlinks(1).Address, "hyperlink address"
End Sub
