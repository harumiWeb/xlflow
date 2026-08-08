VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.29.01  published: 29.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chChartType
    Line = 1
    Bar = 2
    Radar = 3
    Doughnut = 4
    Pie = 5
    PolarArea = 6
    Bubble = 7
    Scatter = 8
End Enum

Public Enum chDataSourceType
    dstDemo
    dstTableName
    dstQueryName
    dstSQLStament
    dstRecordset
    dstData
    dstEmpty
End Enum

Public Enum chPosition
    posTop = 1
    posLeft = 2
    posBottom = 3
    posRight = 4
End Enum

Public Enum chScriptSource
    CDN = 1
    LocalFile = 2
End Enum

Public Enum chAlign
    alStart = 1
    alCenter = 2
    alEnd = 3
End Enum

Public Enum chDataLabelAnchor
    anStart = 1
    anCenter = 2
    anEnd = 3
End Enum

Public Enum chDisplayIn
    chEdgeBrowserControl = 1
    chSystemBrowser = 4
End Enum

Public Enum chEasing
    linear = 0
    easeInQuad = 1
    easeOutQuad = 2
    easeInOutQuad = 3
    easeInCubic = 4
    easeOutCubic = 5
    easeInOutCubic = 6
    easeInQuart = 7
    easeOutQuart = 8
    easeInOutQuart = 9
    easeInQuint = 10
    easeOutQuint = 11
    easeInOutQuint = 12
    easeInSine = 13
    easeOutSine = 14
    easeInOutSine = 15
    easeInExpo = 16
    easeOutExpo = 17
    easeInOutExpo = 18
    easeInCirc = 19
    easeOutCirc = 20
    easeInOutCirc = 21
    easeInElastic = 22
    easeOutElastic = 23
    easeInOutElastic = 24
    easeInBack = 25
    easeOutBack = 26
    easeInOutBack = 27
    easeInBounce = 28
    easeOutBounce = 29
    easeInOutBounce = 30
End Enum

Public Enum chFontWeight
    fwUndefined = 0
    fwThin = 100
    fwExtraLight = 200
    fwLight = 300
    fwNormal = 400
    fwMedium = 500
    fwSemiBold = 600
    fwBold = 700
    fwExtraBold = 800
    fwBlack = 900
End Enum

Public Enum chFontStyle
    fsUndefined = 0
    fsNormal = 1
    fsItalic = 2
    fsOblique = 3
End Enum


'### Members

Private m_EdgeBrowserControl As Access.Edge
Private m_ChartType As chChartType
Private m_Horizontal As Boolean
Private m_StackedChart As Boolean
Private m_AreaFilled As Boolean
Private m_BackgroundColor As String
Private m_ShowDownloadLink As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_AspectRatio As Double
Private m_DisplayIn As chDisplayIn
    
Private m_Animation   As BAC_Animation
Private m_DataSource  As BAC_DataSource
Private m_DataLabels  As BAC_DataLabels
Private m_Defaults    As BAC_Defaults
Private m_XAxis       As BAC_Axis
Private m_YAxis       As BAC_Axis
Private m_RadialAxis  As BAC_Axis
Private m_Legend      As BAC_Legend
Private m_Title       As BAC_Title
Private m_SubTitle    As BAC_SubTitle
Private m_Tooltip     As BAC_Tooltip

    
'### Initializer

Public Sub Init(Optional ByRef TheControl As Object, _
                Optional ByVal ChartType As chChartType = chChartType.Bar, _
                Optional ByVal DataSourceType As chDataSourceType, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldNames As Variant, _
                Optional ByVal LabelFieldName As String)
  
    Set Me.EdgeBrowserControl = TheControl
    Me.ChartType = ChartType
    Me.DataSource.Init DataSourceType, DataSource, DataFieldNames, LabelFieldName
    
End Sub

    
'### Properties

Public Property Set EdgeBrowserControl(ByRef New_Control As Access.Edge)
    Set m_EdgeBrowserControl = New_Control
    m_DisplayIn = chEdgeBrowserControl
End Property
Public Property Get EdgeBrowserControl() As Access.Edge
    Set EdgeBrowserControl = m_EdgeBrowserControl
End Property

Public Property Let ChartType(ByVal This_ChartType As chChartType)
    m_ChartType = This_ChartType
    Select Case This_ChartType
        Case chChartType.Bar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
            YAxis.Show = True
            XAxis.Show = True
            RadialAxis.Show = False
        Case chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea, chChartType.Radar
            YAxis.Show = False
            XAxis.Show = False
            RadialAxis.Show = True
        Case Else
            'Do Nothing
    End Select
End Property
Public Property Get ChartType() As chChartType
    ChartType = m_ChartType
End Property

Public Property Let Horizontal(ByVal This_Horizontal As Boolean)
    m_Horizontal = This_Horizontal
End Property
Public Property Get Horizontal() As Boolean
    Horizontal = m_Horizontal
End Property

Public Property Let StackedChart(ByVal This_StackedChart As Boolean)
    m_StackedChart = This_StackedChart
End Property
Public Property Get StackedChart() As Boolean
    Select Case ChartType
        Case chChartType.Line, chChartType.Bar
            StackedChart = m_StackedChart
        Case Else
            StackedChart = False
    End Select
End Property

Public Property Let AreaFilled(ByVal This_AreaFilled As Boolean)
    m_AreaFilled = This_AreaFilled
End Property
Public Property Get AreaFilled() As Boolean
    AreaFilled = m_AreaFilled
End Property

' Proportion of Width / Height
Public Property Let AspectRatio(ByVal This_Ratio As Double)
    m_AspectRatio = This_Ratio
End Property
Public Property Get AspectRatio() As Double
    AspectRatio = m_AspectRatio
End Property

Public Property Let BackgroundColor(ByVal This_BackgroundColor As String)
    m_BackgroundColor = This_BackgroundColor
End Property
Public Property Get BackgroundColor() As String
    BackgroundColor = m_BackgroundColor
End Property

Public Property Let ShowDownloadLink(ByVal This_ShowDownloadLink As Boolean)
    m_ShowDownloadLink = This_ShowDownloadLink
End Property
Public Property Get ShowDownloadLink() As Boolean
    ShowDownloadLink = m_ShowDownloadLink
End Property

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_SaveHtmlFileForDebugging = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = m_SaveHtmlFileForDebugging
End Property

Public Property Let DisplayIn(ByVal newValue As chDisplayIn)
    m_DisplayIn = newValue
End Property
Public Property Get DisplayIn() As chDisplayIn
    DisplayIn = m_DisplayIn
End Property



'### Object Properties

Public Property Get Animation() As BAC_Animation
    Set Animation = m_Animation
End Property

Public Property Get DataSource() As BAC_DataSource
    Set DataSource = m_DataSource
End Property

Public Property Get Default() As BAC_Defaults
  Set Default = m_Defaults
End Property

Public Property Get Legend() As BAC_Legend
    Set Legend = m_Legend
End Property

Public Property Get Title() As BAC_Title
    Set Title = m_Title
End Property

Public Property Get SubTitle() As BAC_SubTitle
    Set SubTitle = m_SubTitle
End Property

Public Property Get XAxis() As BAC_Axis
    Set XAxis = m_XAxis
End Property

Public Property Get YAxis() As BAC_Axis
    Set YAxis = m_YAxis
End Property

Public Property Get RadialAxis() As BAC_Axis
    Set RadialAxis = m_RadialAxis
End Property

Public Property Get DataLabels() As BAC_DataLabels
    Set DataLabels = m_DataLabels
End Property

Public Property Get Data() As BAC_Data
    Set Data = m_DataSource.Data
End Property

Public Property Get Tooltip() As BAC_Tooltip
    Set Tooltip = m_Tooltip
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'Create Sub Classes
    Set m_Animation = New BAC_Animation
    Set m_DataSource = New BAC_DataSource
    Set m_DataLabels = New BAC_DataLabels
    Set m_Defaults = New BAC_Defaults
    Set m_Legend = New BAC_Legend
    Set m_Title = New BAC_Title
    Set m_SubTitle = New BAC_SubTitle
    Set m_XAxis = New BAC_Axis
    m_XAxis.Init apXAxis
    Set m_YAxis = New BAC_Axis
    m_YAxis.Init apYAxis
    Set m_RadialAxis = New BAC_Axis
    Set m_Tooltip = New BAC_Tooltip
    
    'Set default values
    ChartType = chChartType.Bar
    StackedChart = False
    AspectRatio = 2
    DisplayIn = chEdgeBrowserControl
    'Title
    Title.Show = False
    Title.Position = posTop
    Title.Text = "My first Better Access Chart"
    'SubTitle
    SubTitle.Show = False
    'Legend
    Legend.Show = False
    Legend.Position = posRight
    'Axis
    YAxis.Show = True
    XAxis.Show = True

End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub CreateChart()

    'Declarations
    Dim strHtml As String
    Dim js As String
    
    Select Case Me.DisplayIn
        Case chDisplayIn.chEdgeBrowserControl
            If Me.EdgeBrowserControl.LocationURL = vbNullString Then
                CreateEmptyHtmlFile
                Me.EdgeBrowserControl.Navigate "https://msaccess/" & getHTMLFileName(, True)
                Do While Me.EdgeBrowserControl.ReadyState <> 4
                    DoEvents
                Loop
            End If
            strHtml = BAC.Generator(Me, Me.DataSource.getData(Me.ChartType)).GenerateHTML
            If Me.SaveHtmlFileForDebugging = True Then
                Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
            End If
            strHtml = Replace(strHtml, vbCrLf, "")
            js = js & "document.open();"
            js = js & "document.write(""" & strHtml & """);"
            js = js & "document.close();"
            Me.EdgeBrowserControl.ExecuteJavascript js
        Case chDisplayIn.chSystemBrowser
            CreateChartFile
            BAC.Helper.ShellExecute getHTMLFileName
        Case Else
            'Do nothing
    End Select
    
End Sub

    
'### Private Methods

Private Sub CreateEmptyHtmlFile()

    'Decalarations
    Dim strHtml As New BAC_StringBuilder
    
    With strHtml
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "</body>"
        .AddLine "</html>"
    End With
    
    Call saveHtmlFile(vbNullString, strHtml.ToString, True)
    
End Sub

Private Sub CreateChartFile()

    'Declarations
    Dim strHtml As String
    
    With BAC.Generator(Me, Me.DataSource.getData(Me.ChartType))
        strHtml = .GenerateHTML
    End With
 
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
    End If
    
    Call saveHtmlFile(BAC.Helper.getTempFolder, strHtml)
    
End Sub


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal Folder As String, ByVal strHtml As String, Optional ByVal fEmpty As Boolean = False)

    'Declarations
    Dim objFS As Object
    Dim FileName As String
    Dim strHtmlWithoutPrefix
       
    FileName = getHTMLFileName(Folder, fEmpty)
    strHtmlWithoutPrefix = Replace(strHtml, "https://msaccess/", vbNullString)
    
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtmlWithoutPrefix
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
    
End Sub

Private Function getHTMLFileName(Optional Path As String = vbNullString, Optional fEmpty As Boolean = False) As String
    
    'Declarations
    Dim strFilename As String
    Const HtmlFileName As String = "BetterAccessChart_{}.html"

    If Path = vbNullString Then
        Path = BAC.Helper.getTempFolder
    End If
    If fEmpty = True Then
        strFilename = "Empty.html"
    Else
        strFilename = Replace(HtmlFileName, "{}", ObjPtr(Me))
    End If

    getHTMLFileName = Path & strFilename
    
End Function

