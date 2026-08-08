VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ChartGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.29.01  published: 29.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Options  As BAC_Chart
Private m_Data     As BAC_Data


' ### Properties

Private Property Get Options() As BAC_Chart
    Set Options = m_Options
End Property
Private Property Get Data() As BAC_Data
    Set Data = m_Data
End Property


'### Initializer

Public Sub Init(ByVal ChartOptions As BAC_Chart, ByVal ChartData As BAC_Data)

    Set m_Options = ChartOptions
    Set m_Data = ChartData
    
End Sub


'### Public Methods

Public Function GenerateHTML() As String
    
    Dim strHtml As New BAC_StringBuilder
    With strHtml
        .Add getHeader
        .Add getChartType
        .Add getChartData
        .Add getChartOptions
        .Add getFooter
    End With
    GenerateHTML = strHtml.ToString()
    
End Function


'### Private Methods / Functions

Private Function getHeader() As String
    
    'Declarations
    Dim strText As New BAC_StringBuilder
    
    With strText
        .AddLine "<!DOCTYPE html>"
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "    <title>Chart</title>"
        .AddLine "    <meta charset='utf-8'>"
        .AddLine "    <script src=" & BAC.Scripts.ChartJS.getLocation & "></script>"
        .AddLine "    <script src=" & BAC.Scripts.Plugin_AutoColors.getDataLabelLocation & "></script>"
        If Options.DataLabels.Show = True Then
            .AddLine "    <script src=" & BAC.Scripts.Plugin_DataLabels.getDataLabelLocation & "></script>"
        End If
        .AddLine "    <style>"
        .AddLine "        body { margin: 0; padding: 0; }"
        .AddLine "        #container { width: 100%; }"
        If Len(Options.BackgroundColor) > 0 Then
            .AddLine "        #myChart { background-color: " & Options.BackgroundColor & " }"
        End If
        .AddLine "    </style>"
        .AddLine "</head>"
        
        .AddLine "<body>"
        .AddLine "<div id='container'>"
        .AddLine "    <canvas id='myChart'></canvas>"
        .AddLine "</div>"
        
        If Options.ShowDownloadLink = True Then
            .AddLine "<a href='#' id='download' onclick='daunload()'>Download</a>"
        End If
        
        .AddLine "<script>"
        
        .AddLine "    Chart.register(window['chartjs-plugin-autocolors']);"
        
        If Options.DataLabels.Show = True Then
            .AddLine "    Chart.register(ChartDataLabels);"
        End If
        
        .Add getDefaultFont
        .Add Options.Animation.getAnimation
        .AddLine "    var ctx = document.getElementById('myChart').getContext('2d');"
        .AddLine "    var myChart = new Chart(ctx, {"
        
        getHeader = .ToString
    End With
    
End Function

Private Function getDefaultFont() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    Dim Font As BAC_Font
    
    With strText
        Set Font = Options.Default.Font
        If Len(Font.Color) > 0 Then
            .AddLine "    Chart.defaults.color = '" & Font.Color & "';"
        End If
        If Len(Font.Family) > 0 Then
            .AddLine "    Chart.defaults.font.family = '" & Font.Family & "';"
        End If
        If Font.Size > 0 Then
            .AddLine "    Chart.defaults.font.size = " & Font.Size & ";"
        End If
        If Font.Weight <> chFontWeight.fwUndefined Then
            .AddLine "    Chart.defaults.font.weight = " & Font.Weight & ";"
        End If
        If Font.Style <> chFontStyle.fsUndefined Then
            .AddLine "    Chart.defaults.font.style = " & Font.StyleText & ";"
        End If
        getDefaultFont = .ToString
    End With
    
End Function

Private Function getChartType() As String

    'Declarations
    Dim strType As String
    
    Select Case Options.ChartType
        Case chChartType.Line
            strType = "line"
        Case chChartType.Bar
            strType = "bar"
        Case chChartType.Radar
            strType = "radar"
        Case chChartType.Doughnut
            strType = "doughnut"
        Case chChartType.Pie
            strType = "pie"
        Case chChartType.PolarArea
            strType = "polarArea"
        Case chChartType.Bubble
            strType = "bubble"
        Case chChartType.Scatter
            strType = "scatter"
        Case Else
            'Do Nothing
    End Select
        
    getChartType = "        type: '" & strType & "'," & vbCrLf

End Function

Private Function getChartData() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    Dim i As Integer

    With strText
        .AddLine "        data: {"
        If Not IsBubbleOrScatterChart Then
            .AddLine "            labels: [" & Data.Labels & "],"
        End If
        
        .AddLine "            datasets: ["
        For i = 0 To Data.DataCount - 1
        
            If IsBubbleOrScatterChart Then
                .AddLine "               {label: '" & Data.Label(i) & "' ,"
            Else
                .AddLine "               {label: '" & Data.DatasetName(i) & "' ,"
            End If
            
            Select Case Options.ChartType
                Case chChartType.Line, chChartType.Radar
                    If Options.AreaFilled = True Then
                        .AddLine "                fill: true ,"
                    Else
                        .AddLine "                fill: false ,"
                    End If
                Case Else
                    'Do Nothing
            End Select
                    
            .AddLine "                data: [" & Data.Data(i) & "],"
            
'' ### Reminder for new coloring solution
''            If IsSingleBarChart Then
''                .AddLine "                backgroundColor: Chart['colorschemes']." & Options.ColorSchema.Name & ","
''            End If
            
            .AddLine "                borderWidth: 1},"
            
        Next i
        .CutTrailingComma
        .AddLine "             ]"
        .AddLine "        },"
        
        getChartData = .ToString
    End With
    
End Function
    
Private Function getChartOptions() As String
    
    'Declarations
    Dim strText As New BAC_StringBuilder
    
    With strText
        .AddLine "        options: {"
        If Options.Horizontal = True Then
            If Options.ChartType = chChartType.Bar Or Options.ChartType = chChartType.Line Then
            .AddLine "            indexAxis: 'y',"
            End If
        End If
        If Options.AspectRatio <> 2# Then
            .AddLine "            aspectRatio: " & Str(Options.AspectRatio) & ","
        End If
        .Add getScales
        .Add getPlugins
        .CutTrailingComma
        .AddLine "        }"
        
        getChartOptions = .ToString
    End With
    
End Function

Private Function getFooter() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    
    With strText
        .AddLine "    });"  ' closing: var myChart = new Chart(ctx, {"
        
        If Options.ShowDownloadLink = True Then
            .AddLine "    function daunload() {"
            .AddLine "        var myChart = document.getElementById('myChart');"
            .AddLine "        var image = myChart.toDataURL('image/png').replace('image/png', 'image/octet-stream');"
            .AddLine "        window.location.href=image;"
            .AddLine "    }"
        End If
        .AddLine "</script>"
        .AddLine "</body>"
        .AddLine "</html>"
        
        getFooter = strText.ToString
    End With
    
End Function

Private Function getScales() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    
    Select Case Options.ChartType
        Case chChartType.Bar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
            With strText
                .AddLine "            scales: {"
                .Add Options.YAxis.getCartesianAxis(Options.StackedChart)
                .Add Options.XAxis.getCartesianAxis(Options.StackedChart)
                .CutTrailingComma
                .AddLine "            },"
                getScales = .ToString
            End With
        Case chChartType.PolarArea, chChartType.Radar
            With strText
                .AddLine "            scales: {"
                .Add Options.RadialAxis.getRadialAxis
                .CutTrailingComma
                .AddLine "            },"
                getScales = .ToString
            End With
        Case chChartType.Pie, chChartType.Doughnut
            'No axes available
        Case Else
            'Do Nothing
    End Select
    
End Function

Private Function getPlugins() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    
    With strText
        .AddLine "            plugins: {"
        .Add Options.Title.getTitle
        .Add Options.SubTitle.getTitle
        .Add Options.Legend.getLegend
        .Add Options.Tooltip.getTootip
        .Add Options.DataLabels.getDataLabel(Options)
        .Add getAutoColors
        .CutTrailingComma
        .AddLine "            },"
        
        getPlugins = .ToString
    End With
    
End Function

Private Function IsSingleBarChart()

    Select Case Options.ChartType
        Case chChartType.Bar
            If Data.IsSingleDataset = True Then
                IsSingleBarChart = True
            Else
                IsSingleBarChart = False
            End If
        Case Else
            IsSingleBarChart = False
    End Select

End Function

Private Function IsBubbleOrScatterChart() As Boolean

    If Options.ChartType = Bubble Then
        IsBubbleOrScatterChart = True
    ElseIf Options.ChartType = Scatter Then
        IsBubbleOrScatterChart = True
    Else
        IsBubbleOrScatterChart = False
    End If
    
End Function

Private Function getAutoColors() As String

    If Options.ChartType = Pie Or Options.ChartType = Doughnut Or Options.ChartType = PolarArea Then
        With New BAC_StringBuilder
            .AddLine "                autocolors: {"
            .AddLine "                    mode: 'data'"
            .AddLine "                }"
            getAutoColors = .ToString
        End With
    Else
        'Return nothing
    End If

End Function

