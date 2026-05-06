Attribute VB_Name = "App"
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
  Dim newsRoot As Object

  On Error GoTo ErrHandler

  Set newsRoot = NewsApi.FetchWorldNews()
  NewsSheet.RenderArticles wb, newsRoot
  Exit Sub

ErrHandler:
  NewsSheet.RenderFailure wb, Err.Source, Err.Description
End Sub
