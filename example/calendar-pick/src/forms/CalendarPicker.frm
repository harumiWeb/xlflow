VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarPicker 
   Caption         =   "日付を選択"
   ClientHeight    =   4680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "CalendarPicker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalendarPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DAYS_IN_WEEK As Long = 7
Private Const WEEKS_IN_VIEW As Long = 6
Private Const HEADER_TOP As Single = 12
Private Const HEADER_LEFT As Single = 6
Private Const HEADER_WIDTH As Single = 32
Private Const HEADER_HEIGHT As Single = 12
Private Const HEADER_GAP As Single = 2
Private Const BUTTON_TOP As Single = 27
Private Const BUTTON_LEFT As Single = 6
Private Const BUTTON_WIDTH As Single = 32
Private Const BUTTON_HEIGHT As Single = 18
Private Const BUTTON_GAP As Single = 2
Private Const YEAR_MIN As Long = 1900
Private Const YEAR_MAX As Long = 2100
Private Const COLOR_SUNDAY As Long = 192
Private Const COLOR_SATURDAY As Long = 13395456
Private Const COLOR_SUNDAY_MUTED As Long = 10079487
Private Const COLOR_SATURDAY_MUTED As Long = 14181852
Private Const COLOR_WEEKDAY As Long = 0
Private Const COLOR_WEEKDAY_MUTED As Long = 9211020

Private mDisplayMonth As Date
Private mSelectedDate As Date
Private mHasSelection As Boolean
Private mIsConfirmed As Boolean
Private mDayHandlers As Collection
Private mIsInitializing As Boolean

Public Property Get IsConfirmed() As Boolean
   IsConfirmed = mIsConfirmed
End Property

Public Property Get SelectedDate() As Date
   SelectedDate = mSelectedDate
End Property

Public Function PickDate(ByRef wasConfirmed As Boolean) As Date
   Me.Show
   wasConfirmed = mIsConfirmed
   If mIsConfirmed Then
      PickDate = mSelectedDate
   End If
End Function

Private Sub BtnCancel_Click()
   mIsConfirmed = False
   Me.Hide
End Sub

Private Sub BtnNextMonth_Click()
   mDisplayMonth = DateAdd("m", 1, mDisplayMonth)
   RenderCalendar
End Sub

Private Sub BtnPrevMonth_Click()
   mDisplayMonth = DateAdd("m", -1, mDisplayMonth)
   RenderCalendar
End Sub

Private Sub BtnToday_Click()
   HandleDayButtonClick Date
End Sub

Private Sub CmbYear_Change()
   If mIsInitializing Then Exit Sub
   If Len(CmbYear.Value) = 0 Then Exit Sub

   mDisplayMonth = DateSerial(CLng(CmbYear.Value), Month(mDisplayMonth), 1)
   RenderCalendar
End Sub

Public Sub HandleDayButtonClick(ByVal chosenDate As Date)
   mSelectedDate = chosenDate
   mHasSelection = True
   mIsConfirmed = True
   LblSelectedDate.Caption = "選択日: " & Format$(chosenDate, "yyyy/mm/dd")
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
   mIsInitializing = True
   Set mDayHandlers = New Collection
   mDisplayMonth = DateSerial(Year(Date), Month(Date), 1)
   mHasSelection = False
   mIsConfirmed = False

   PopulateYearOptions Year(mDisplayMonth)
   BuildWeekdayHeaders
   RenderCalendar
   mIsInitializing = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      mIsConfirmed = False
   End If
End Sub

Private Sub BuildWeekdayHeaders()
   Dim dayNames As Variant
   Dim dayIndex As Long
   Dim headerLabel As MSForms.Label

   dayNames = Array("日", "月", "火", "水", "木", "金", "土")

   For dayIndex = 0 To UBound(dayNames)
      Set headerLabel = Me.FrameCalendar.Controls.Add("Forms.Label.1", "WeekdayLabel" & CStr(dayIndex), True)
      headerLabel.Caption = CStr(dayNames(dayIndex))
      headerLabel.Left = HEADER_LEFT + (HEADER_WIDTH + HEADER_GAP) * dayIndex
      headerLabel.Top = HEADER_TOP
      headerLabel.Width = HEADER_WIDTH
      headerLabel.Height = HEADER_HEIGHT
      headerLabel.TextAlign = fmTextAlignCenter
      headerLabel.ForeColor = GetColumnColor(dayIndex, True)
   Next dayIndex
End Sub

Private Sub PopulateYearOptions(ByVal selectedYear As Long)
   Dim yearValue As Long

   CmbYear.Clear
   CmbYear.Style = fmStyleDropDownList

   For yearValue = YEAR_MIN To YEAR_MAX
      CmbYear.AddItem CStr(yearValue)
   Next yearValue

   CmbYear.Value = CStr(selectedYear)
End Sub

Private Sub ClearDayButtons()
   Dim controlIndex As Long

   For controlIndex = Me.FrameCalendar.Controls.Count - 1 To 0 Step -1
      If TypeName(Me.FrameCalendar.Controls.Item(controlIndex)) = "CommandButton" Then
         Me.FrameCalendar.Controls.Remove Me.FrameCalendar.Controls.Item(controlIndex).Name
      End If
   Next controlIndex

   Set mDayHandlers = New Collection
End Sub

Private Sub RenderCalendar()
   Dim firstVisibleDate As Date
   Dim dayOffset As Long
   Dim currentDate As Date
   Dim dayButton As MSForms.CommandButton
   Dim dayHandler As CalendarDayButton
   Dim rowIndex As Long
   Dim columnIndex As Long

   ClearDayButtons
   mIsInitializing = True
   CmbYear.Value = CStr(Year(mDisplayMonth))
   mIsInitializing = False
   LblMonthYear.Caption = Format$(mDisplayMonth, "m月")

   firstVisibleDate = DateAdd("d", -Weekday(mDisplayMonth, vbSunday) + 1, mDisplayMonth)

   For dayOffset = 0 To (DAYS_IN_WEEK * WEEKS_IN_VIEW) - 1
      currentDate = DateAdd("d", dayOffset, firstVisibleDate)
      rowIndex = dayOffset \ DAYS_IN_WEEK
      columnIndex = dayOffset Mod DAYS_IN_WEEK

      Set dayButton = Me.FrameCalendar.Controls.Add("Forms.CommandButton.1", "DayButton" & Format$(dayOffset, "00"), True)
      dayButton.Caption = CStr(Day(currentDate))
      dayButton.Left = BUTTON_LEFT + (BUTTON_WIDTH + BUTTON_GAP) * columnIndex
      dayButton.Top = BUTTON_TOP + (BUTTON_HEIGHT + BUTTON_GAP) * rowIndex
      dayButton.Width = BUTTON_WIDTH
      dayButton.Height = BUTTON_HEIGHT
      dayButton.ForeColor = GetColumnColor(columnIndex, Month(currentDate) = Month(mDisplayMonth))

      If mHasSelection Then
         If CLng(currentDate) = CLng(mSelectedDate) Then
            dayButton.BackColor = RGB(211, 233, 255)
         End If
      ElseIf CLng(currentDate) = CLng(Date) Then
         dayButton.BackColor = RGB(240, 245, 255)
      End If

      Set dayHandler = New CalendarDayButton
      dayHandler.Bind Me, dayButton, currentDate
      mDayHandlers.Add dayHandler
   Next dayOffset

   If mHasSelection Then
      LblSelectedDate.Caption = "選択日: " & Format$(mSelectedDate, "yyyy/mm/dd")
   Else
      LblSelectedDate.Caption = "選択日: -"
   End If
End Sub

Private Function GetColumnColor(ByVal columnIndex As Long, ByVal isCurrentMonth As Boolean) As Long
   Select Case columnIndex
      Case 0
         If isCurrentMonth Then
            GetColumnColor = COLOR_SUNDAY
         Else
            GetColumnColor = COLOR_SUNDAY_MUTED
         End If
      Case 6
         If isCurrentMonth Then
            GetColumnColor = COLOR_SATURDAY
         Else
            GetColumnColor = COLOR_SATURDAY_MUTED
         End If
      Case Else
         If isCurrentMonth Then
            GetColumnColor = COLOR_WEEKDAY
         Else
            GetColumnColor = COLOR_WEEKDAY_MUTED
         End If
   End Select
End Function
