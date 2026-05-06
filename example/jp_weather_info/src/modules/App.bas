Attribute VB_Name = "App"
Option Explicit

Public Sub RunCore(ByVal wb As Workbook)
  Dim forecastRoot As Object
  Set forecastRoot = WeatherApi.FetchForecast(WeatherApi.DefaultCityCode())
  WeatherSheet.RenderForecast wb, forecastRoot
End Sub
