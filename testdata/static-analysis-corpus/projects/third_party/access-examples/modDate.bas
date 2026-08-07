Attribute VB_Name = "modDate"
Option Compare Database
'--------------------------------------------------------------------------------------------------------------------
' Purpose:              Date and time subroutines
'
' Ver.  Date            Author              Details
' 1.00  01 JAN 2001     Microsoft           Initial version.
'--------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Sub Timedelay(secondsdelay As Integer)
'Call from Code to delay for Passed in Number of seconds
'For Instance to Delay for 5 seconds insert the following line of code without the '
'call Timedelay(5)
Dim tStartTime
   
   tStartTime = Now
   Do Until Second(tStartTime - Now) >= secondsdelay
   Loop

End Sub

Public Function DaysInMonth(MyDate)
' Note pass date in this format       DaysInMonth (#mm/dd/yyyy#)
' example for entering Date           DaysInMonth (#2/2/1995#)
' example for Prompting               DaysInMonth ([Please Enter Date?])
' or if refering to field in query    DaysInMonth ([WELL_MNLY_ALOC_DTE])
      
      Dim NextMonth, EndOfMonth
      NextMonth = DateAdd("m", 1, MyDate)
      EndOfMonth = NextMonth - DatePart("d", NextMonth)
      DaysInMonth = DatePart("d", EndOfMonth)

End Function

Public Function Age(Bdate, DateToday) As Integer
On Error Resume Next
'
' Returns the Age in years between 2 dates
' Doesn't handle negative date ranges i.e. Bdate > DateToday
'
  If month(DateToday) < month(Bdate) Or (month(DateToday) = month(Bdate) And day(DateToday) < day(Bdate)) Then
    Age = year(DateToday) - year(Bdate) - 1
  Else
    Age = year(DateToday) - year(Bdate)
  End If
  
End Function

Function DaysInMonthMS(d As Variant) As Variant
On Error Resume Next
'
' Returns the number of days in a month.
' Requires a date argument, since February can change if it's a leap year
'
  If VarType(d) <> 7 Then
    DaysInMonthMS = Null
  Else
    Select Case month(d)
      Case 2
        If LeapYear(year(d)) Then
          DaysInMonthMS = 29
        Else
          DaysInMonthMS = 28
        End If
      Case 4, 6, 9, 11
        DaysInMonthMS = 30
      Case 1, 3, 5, 7, 8, 10, 12
        DaysInMonthMS = 31
    End Select
  End If
End Function

Function DaysInMonth2(d As Variant) As Variant
On Error Resume Next
'
' Returns the number of days in a month
' Requires a date argument, since February can change if it's a leap year
' Lets Access figure it out
'
  If VarType(d) <> 7 Then
    DaysInMonth2 = Null
  Else
    DaysInMonth2 = DateSerial(year(d), month(d) + 1, 1) - DateSerial(year(d), month(d), 1)
  End If
End Function

Function EndOfMonth(d As Variant) As Variant
On Error Resume Next
'
' Returns the date representing the last day of the current month.
'
' Arguments:
' D            = Date
  
  EndOfMonth = DateSerial(year(d), month(d) + 1, 0)
  
End Function

Function EndOfWeek(d As Variant, Optional FirstWeekday As Integer) As Variant
On Error Resume Next
'
' Returns the date representing the last day of the current week.
'
' Arguments:
' D            = Date
' FirstWeekday = (Optional argument) Integer that represents the first
' day of the week (e.g., 1=Sun..7=Sat).
'
If IsMissing(FirstWeekday) Then 'Sunday is the assumed first day of week.
  EndOfWeek = d - WeekDay(d) + 7
Else
  EndOfWeek = d - WeekDay(d, FirstWeekday) + 7
End If
End Function

Function FormatInterval(ByVal Interval As Variant, Fmt As String)
On Error Resume Next
'
' Formats the difference between two dates or sum of two times
' to show day as well as hours, minutes, and seconds.
'
' Supports the following formats:
'   D H                    5 Days 5 Hours
'   D H:MM                 5 Days 5:15
'   D HH:MM                5 Days 05:15
'   D H:MM:SS              5 Days 5:15:45
'   D HH:MM:SS             5 Days 05:15:45
'   H M                    125 Hours 15 Minutes
'   H:MM                   125:15
'   H:MM:SS                125:15:45
'   M S                    7515 Minutes 45 Seconds
'
Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long
'
' Check for Date or Double
'
  If VarType(Interval) <> 7 And VarType(Interval) <> 5 Then Exit Function
'
' Parse Days
'
  Days = Int(Interval)
  Interval = Interval - Days
  If Interval > #11:59:59 PM# Then
    Days = Days + 1
    Interval = 0#
  End If
'
' Parse Hours
'
  Interval = Interval * 24
  Hours = Int(Interval)
  Interval = Interval - Hours
  If Interval > 3599# / 3600# Then
    Hours = Hours + 1
    Interval = 0#
  End If
'
' Parse Minutes
'
  Interval = Interval * 60
  Minutes = Int(Interval)
  Interval = Interval - Minutes
  If Interval > 59# / 60# Then
    Minutes = Minutes + 1
    Interval = 0#
  End If
'
' Parse Seconds
'
  Seconds = Int(Interval * 60 + 0.5)
'
' Normalize
'
  If Seconds = 60 Then
    Minutes = Minutes + 1
    Seconds = 0
  End If
  If Minutes > 59 Then
    Hours = Hours + 1
    Minutes = Minutes - 60
  End If
  If Hours > 23 Then
    Days = Days + 1
    Hours = Hours - 24
  End If
'
' Create format
'
  Select Case Fmt
    Case "D H"
      FormatInterval = Days & IIf(Days <> 1, " Days ", " Day ") & Hours & IIf(Hours <> 1, " Hours", " Hour")
    Case "D H:MM"
      FormatInterval = Days & IIf(Days <> 1, " Days ", " Day ") & Hours & ":" & Format(Minutes, "00")
    Case "D HH:MM"
      FormatInterval = Days & IIf(Days <> 1, " Days ", " Day ") & Format(Hours, "00") & ":" & Format(Minutes, "00")
    Case "D H:MM:SS"
      FormatInterval = Days & IIf(Days <> 1, " Days ", " Day ") & Hours & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
    Case "D HH:MM:SS"
      FormatInterval = Days & IIf(Days <> 1, " Days ", " Day ") & Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
    Case "H M"
      Hours = Hours + Days * 24
      FormatInterval = Hours & IIf(Hours <> 1, " Hours ", " Hour ") & Minutes & IIf(Minutes <> 1, " Minutes", " Minute")
    Case "H:MM"
      Hours = Hours + Days * 24
      FormatInterval = Hours & ":" & Format(Minutes, "00")
    Case "H:MM:SS"
      Hours = Hours + Days * 24
      FormatInterval = Hours & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
    Case "M S"
      Minutes = Minutes + (Hours + Days * 24) * 60
      FormatInterval = Minutes & IIf(Minutes <> 1, " Minutes ", " Minute ") & Seconds & IIf(Seconds <> 1, " Seconds", " Second")
    Case Else
      FormatInterval = Null
  End Select
End Function

Function LastBusDay(d As Variant) As Variant
On Error Resume Next
'
' Returns the date of the last business day (Mon-Fri) in a month
'
Dim D2 As Variant
  If VarType(d) <> 7 Then
    LastBusDay = Null
  Else
    D2 = DateSerial(year(d), month(d) + 1, 0)
    Do While WeekDay(D2) = 1 Or WeekDay(D2) = 7
      D2 = D2 - 1
    Loop
    LastBusDay = D2
  End If
End Function

Function LeapYear(YYYY As Integer) As Integer
On Error Resume Next
'
' Leap Year from standard rules
' YYYY: 4-digit year
'
  LeapYear = YYYY Mod 4 = 0 And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
End Function

Function LeapYear2(YYYY As Integer) As Integer
On Error Resume Next
'
' Leap Year letting Access figure out the rules
' YYYY: 4-digit year
'
  LeapYear2 = month(DateSerial(YYYY, 2, 29)) = 2
End Function


Function NextDay(d As Variant, DayCode As Integer) As Variant
On Error Resume Next
'
' Returns the date of the next DayCode (1=Sun ... 7=Sat) after the
' date D.  e.g.  NextDay(#5/12/94#,6) returns the date of the next
' Friday after 5/12/94.
'
  NextDay = d - WeekDay(d) + DayCode + IIf(WeekDay(d) < DayCode, 0, 7)
End Function


Function NextDay1(d As Variant, DayCode As Integer) As Variant
On Error Resume Next
'
' Returns the date of the next DayCode (1=Sun ... 7=Sat) on or after the
' date D.
'
' e.g.  NextDay1(#5/12/94#,6) returns the date of the next Friday after 5/12/94,
' or, if 5/12/94 is a Friday, then returns that date (5/12/94).
'
  NextDay1 = d - WeekDay(d) + DayCode + IIf(WeekDay(d) <= DayCode, 0, 7)
End Function


Function Num2Date(ByVal n, Fmt As String)
On Error Resume Next
'
' Converts numbers to dates depending on Fmt
' See Help on Format for what the various strings mean
' Comments use 27 May, 1993 as example date to show input text
'
  If Not IsNumeric(n) Then
    Num2Date = Null
    Exit Function
  End If
  n = CLng(Int(n))

  Select Case Fmt
    Case "MMDDYY"                                                                '052793
      Num2Date = CVDate(n \ 10000 & "/" & n \ 100 Mod 100 & "/" & n Mod 100)
    Case "MMDDYYYY"                                                              '05271993
      Num2Date = CVDate(n \ 1000000 & "/" & n \ 10000 Mod 100 & "/" & n Mod 10000)
    Case "DDMMYY"                                                                '270593
      Num2Date = CVDate(n \ 100 Mod 100 & "/" & n \ 10000 & "/" & n Mod 100)
    Case "DDMMYYYY"                                                              '27051993
      Num2Date = CVDate(n \ 10000 Mod 100 & "/" & n \ 1000000 & "/" & n Mod 10000)
    Case "YYMMDD", "YYYYMMDD"                                                    '930527   19930527
      Num2Date = CVDate(n \ 100 Mod 100 & "/" & n Mod 100 & "/" & n \ 10000)
    Case Else
      Num2Date = Null
  End Select
End Function


Function PriorDay(d As Variant, DayCode As Integer) As Variant
On Error Resume Next
'
' Returns the date of the last DayCode (1=Sun ... 7=Sat) before the
' date D.  e.g.  PriorDay(#5/12/94#,6) returns the date of the
' Friday prior to 5/12/94.
'
  PriorDay = d - WeekDay(d) + DayCode - IIf(WeekDay(d) > DayCode, 0, 7)
End Function

Function PriorDay1(d As Variant, DayCode As Integer) As Variant
'
' Returns the date of the last DayCode (1=Sun ... 7=Sat) on or
' before the date D.
'
' e.g.  PriorDay1(#5/12/94#,6) returns the date of the Friday prior to 5/12/94,
' or if 5/12/94 is a Friday, then returns that date (5/12/94).
'
  PriorDay1 = d - WeekDay(d) + DayCode - IIf(WeekDay(d) >= DayCode, 0, 7)
End Function


Function StartOfMonth(d As Variant) As Variant
On Error Resume Next
'
' Returns the date representing the first day of the current month.
'
' Arguments:
' D            = Date
  
  StartOfMonth = DateSerial(year(d), month(d), 1)
  
End Function

Public Function StartOfWeek(d As Variant, FirstWeekday As Integer, Optional StartWeek As Integer = vbSunday) As Variant
'---------------------------------------------------------------------------
'Purpose:           Used in the calendar style report
'---------------------------------------------------------------------------
On Error Resume Next
Dim intTemp As Integer

    intTemp = FirstWeekday - DatePart("w", d)
    If intTemp = 0 Then
        StartOfWeek = d
    Else
        StartOfWeek = DateAdd("d", intTemp, d)
    End If

End Function

Function String2Date(s, Fmt As String)
On Error Resume Next
'
' Converts strings to dates depending on Fmt
' See Help on Format for what the various strings mean
' Comments use 27 May, 1993 as example date to show input text
'
  If VarType(s) <> 8 Then
    String2Date = Null
    Exit Function
  End If
  Select Case Fmt
    Case "MMDDYY", "MMDDYYYY"                                                    '052793   05271993
      String2Date = CVDate(Left(s, 2) & "/" & Mid(s, 3, 2) & "/" & Mid(s, 5))
    Case "DDMMYY", "DDMMYYYY"                                                    '270593   27051993
      String2Date = CVDate(Mid(s, 3, 2) & "/" & Left(s, 2) & "/" & Mid(s, 5))
    Case "YYMMDD"                                                                '930527
      String2Date = CVDate(Mid(s, 3, 2) & "/" & Right(s, 2) & "/" & Left(s, 2))
    Case "YYYYMMDD"                                                              '19930527
      String2Date = CVDate(Mid(s, 5, 2) & "/" & Right(s, 2) & "/" & Left(s, 4))
    Case "MM/DD/YY", "MM/DD/YYYY", "M/D/Y", "M/D/YY", "M/D/YYYY", "DD-MMM-YY", "DD-MMM-YYYY"
      String2Date = CVDate(s)
    Case "DD/MM/YY", "DD/MM/YYYY"                                                '27/05/93   27/05/1993
      String2Date = CVDate(Mid(s, 4, 3) & Left(s, 3) & Mid(s, 7))
    Case "YY/MM/DD"                                                              '93/05/27
      String2Date = CVDate(Mid(s, 4, 3) & Right(s, 2) & "/" & Left(s, 2))
    Case "YYYY/MM/DD"                                                            '1993/05/27
      String2Date = CVDate(Mid(s, 6, 3) & Right(s, 2) & "/" & Left(s, 4))
    Case Else
      String2Date = Null
  End Select
End Function

Function WorkDays(BegDate As Variant, EndDate As Variant) As Integer
' Note that this function does not account for holidays.
' Example of Reference in a query
'   WorkDays([StartDate],[FinishDate])
Dim WholeWeeks As Variant
Dim DateCnt As Variant
Dim EndDays As Integer

    BegDate = DateValue(BegDate)

    EndDate = DateValue(EndDate)
        WholeWeeks = DateDiff("w", BegDate, EndDate)
        DateCnt = DateAdd("ww", WholeWeeks, BegDate)
        EndDays = 0
        Do While DateCnt < EndDate
            If Format(DateCnt, "ddd") <> "Sun" And _
                          Format(DateCnt, "ddd") <> "Sat" Then
                    EndDays = EndDays + 1
            End If
            DateCnt = DateAdd("d", 1, DateCnt)
        Loop
        WorkDays = WholeWeeks * 5 + EndDays
End Function
