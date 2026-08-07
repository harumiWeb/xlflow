VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMonthCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 VBA
' There is a seperate A2K or higher version required because
' A97 does not support AddressOf.
' That is the only difference between the A97 and the A2K versions.
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.
'           Pedro Gil
'           Please feel free to use this code within your own
'           projects whether they are private or commercial applications
'           without obligation.
'           This code may not be resold by itself or as part of a collection.
'
'Name:      clsMonthCal
'
'Version:   2.05
'
'Purpose:
'           To allow for the selection of a Date or Dates using the standard
'           Windows Month Calendar control without having to use the ActiveX control.
' 
'Authors:   Stephen Lebans
'           Pedro Gil

'Email:     Stephen@lebans.com
'           pmpg98@hotmail.com
'
'Web Site:  www.lebans.com
'           http://www.geocities.com/pmpg98_pt/
'
'Date:      December 09, 2004, 1:11:11 PM
'
'Credits:   Pedro Gil for doing the background
'           research necessary to code a working model.
'           Stephen Lebans for the API non ActiveX solution,
'           with autosizing, user selectable display.
'           Michael Kaplan for the idea to create the
'           Month Calendar directly from the Common Control DLL.
'           jwolzvb@yahoo.de for a VB example of implementing
'           a 1 month ActiveX MonthCalendar control.
'
'BUGS:      Please report any bugs to Stephen@Lebans.com
'
'What's Missing:
'           Proper error handling.
'
'How it Works:
'           The Month Calendar is created directly with the
'           API's contained in the Common Controls DLL. In this manner we bypass
'           the DatePicker ActiveX control, which is simply a wrapper for these
'           calls anyway. This removes any problems from distribution and
'           especially version issues of using the ActiveX control.

'           DAYSTATE:
'           Have a look at the Initialize sub for this Class to see how to
'           set the DayStates for the months. You can also set the bits
'           in the array yourself if you are converting data that you
'           are storing in your own tables.
'
' This is the ?th release. We have exposed most of the properties of the
' Month Calendar. Simply add your own for any we have left out.

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

'Window Style Bits
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000

' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1

'Windows Style Bits Specific to DTP
Private Const DTS_UPDOWN          As Long = &H1   '0x0001 // use UPDOWN instead of MONTHCAL
Private Const DTS_SHOWNONE        As Long = &H2   '0x0002 // allow a NONE selection
Private Const DTS_SHORTDATEFORMAT As Long = &H0   '0x0000 // use the short date format (app must forward WM_WININICHANGE messages)
Private Const DTS_LONGDATEFORMAT  As Long = &H4   '0x0004 // use the long date format (app must forward WM_WININICHANGE messages)
Private Const DTS_TIMEFORMAT      As Long = &H9   '0x0009 // use the time format (app must forward WM_WININICHANGE messages)
Private Const DTS_APPCANPARSE     As Long = &H10  '0x0010 // allow user entered strings (app MUST respond to DTN_USERSTRING)
Private Const DTS_RIGHTALIGN      As Long = &H20  '0x0020 // right-align popup instead of left-align it

' Used to change Font for Month Calendar common control
Private Const WM_SETFONT = &H30

'// bit-packed array of "bold" info for a month
'// if a bit is on, that day is drawn bold
Private Type MONTHDAYSTATE
  lpMONTHDAYSTATE As Long
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type


Private Type NMHDR
hwndFrom As Long
    idfrom As Long
    Code As Long 'Integer
    End Type
    
Private Type NMDAYSTATE
 nmhd As NMHDR ' // this must be first, so we don't break WM_NOTIFY
 stStart As SYSTEMTIME
cDayState As Long
' FOr ease of use always specify 12 months of data
prgDayState(11) As MONTHDAYSTATE '; // points to cDayState MONTHDAYSTATEs
End Type

'Get the value of the Control
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 
' Private Type SYSTEMTIME
'   wYear As Integer
'   wMonth As Integer
'   wDayOfWeek As Integer
'   wDay As Integer
'   wHour As Integer
'   wMinute As Integer
'   wSecond As Integer
'   wMilliseconds As Integer
' End Type
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 


Private Declare Function apiCreateWindowEx Lib "user32.dll" _
  Alias "CreateWindowExA" _
  (ByVal dwExStyle As Long, _
  ByVal lpClassName As String, _
  ByVal lpWindowName As String, _
  ByVal dwStyle As Long, _
  ByVal x As Long, _
  ByVal y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hWndParent As Long, _
  ByVal hMenu As Long, _
  ByVal hInstance As Long, _
  lpParam As Any) As Long
  
Private Declare Function apiDestroyWindow Lib "user32" _
  Alias "DestroyWindow" _
  (ByVal hwnd As Long) As Long

Private Declare Function apiCreateFont Lib "gdi32" Alias "CreateFontA" _
(ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, _
ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, _
ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, _
ByVal PAF As Long, ByVal f As String) As Long

Private Declare Function apiMulDiv Lib "kernel32" Alias "MulDiv" (ByVal nNumber As Long, _
ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'To register the class
Private Type INITCOMMONCONTROLSEXSTRUCT
  dwsize As Long
  dwICC As Long
End Type

Private Declare Function apiInitCommonControlsEx Lib "Comctl32" _
  Alias "InitCommonControlsEx" _
  (lpInitCtrls As INITCOMMONCONTROLSEXSTRUCT) As Long
 
Private Const ICC_DATE_CLASSES = &H100

'Get Instance of Access by KEN GETZ
'required for CreateWindowEx hInstance
'Put this in the declarations area of a standard module.
Private Declare Function apiGetWindowLong Lib "user32" _
  Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Const GWL_HINSTANCE = (-6)

'To check the space needed by the control
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPSPERINCH = 1440

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function AdjustWindowRect Lib "user32" _
(lpRectl As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long

Private Declare Function AdjustWindowRectEx Lib "user32" _
(lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function apiSendMessage Lib "user32" _
  Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lparam As Any) As Long

Private Declare Function apiSetWindowPos Lib "user32" _
  Alias "SetWindowPos" _
  (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal x As Long, _
  ByVal y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal wFlags As Long) As Long

Private Declare Function apiCreateIC Lib "gdi32" _
  Alias "CreateICA" _
  (ByVal lpDriverName As String, _
  ByVal lpDeviceName As String, _
  ByVal lpOutput As String, _
  lpInitData As Any) As Long

Private Declare Function apiGetDeviceCaps Lib "gdi32" _
  Alias "GetDeviceCaps" _
  (ByVal hdc As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function apiDeleteDC Lib "gdi32" _
  Alias "DeleteDC" _
  (ByVal hdc As Long) As Long
 
 
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
 (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

 Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function Beep Lib "kernel32" _
(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


' Color Types
Private Const CTLCOLOR_MSGBOX = 0
Private Const CTLCOLOR_EDIT = 1
Private Const CTLCOLOR_LISTBOX = 2
Private Const CTLCOLOR_BTN = 3
Private Const CTLCOLOR_DLG = 4
Private Const CTLCOLOR_SCROLLBAR = 5
Private Const CTLCOLOR_STATIC = 6
Private Const CTLCOLOR_MAX = 8   '  three bits max

Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNHIGHLIGHT = 20

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 
'SysMonthCal32 messages
Private Const MCM_FIRST = &H1000&
Private Const MCM_GETCURSEL = &H1001& '(MCM_FIRST + 1)
Private Const MCM_SETCURSEL = (MCM_FIRST + 2)
Private Const MCM_GETMINREQRECT = (MCM_FIRST + 9)
Private Const MCM_SETMAXSELCOUNT = (MCM_FIRST + 4)

Private Const MCM_GETRANGE = (MCM_FIRST + 17)
Private Const MCM_SETRANGE = (MCM_FIRST + 18)
Private Const MCM_GETSELRANGE = (MCM_FIRST + 5)
Private Const MCM_GETMONTHRANGE = (MCM_FIRST + 7)


Private Const MCM_SETCOLOR = (MCM_FIRST + 10)
Private Const MCM_GETCOLOR = (MCM_FIRST + 11)

Private Const MCM_SETTODAY = (MCM_FIRST + 12)
Private Const MCM_GETTODAY = (MCM_FIRST + 13)
Private Const MCM_SETFIRSTDAYOFWEEK = (MCM_FIRST + 15)
Private Const MCM_GETFIRSTDAYOFWEEK = (MCM_FIRST + 16)
Private Const MCM_SETDAYSTATE = (MCM_FIRST + 8)
Private Const MCM_SETSELRANGE = (MCM_FIRST + 6)


Private Const MCM_GETMAXTODAYWIDTH = (MCM_FIRST + 21)

'Color part's of the Calendar
Private Const MCSC_BACKGROUND = 0    '// the background color (between months)
Private Const MCSC_TEXT = 1          '// the dates
Private Const MCSC_TITLEBK = 2       '// background of the title
Private Const MCSC_TITLETEXT = 3
Private Const MCSC_MONTHBK = 4       '// background within the month cal
Private Const MCSC_TRAILINGTEXT = 5  '// the text color of header & trailing days

Private Const MCS_DAYSTATE = &H1
Private Const MCS_MULTISELECT = &H2
Private Const MCS_WEEKNUMBERS = &H4
Private Const MCS_NOTODAYCIRCLE = &H8
Private Const MCS_NOTODAY = &H10
Private Const GMR_VISIBLE = 0       '// visible portion of display
Private Const GMR_DAYSTATE = 1      '    // above plus the grayed out parts of
                                    '// partially displayed months

' Window field offsets for GetWindowLong() and GetWindowWord()
Private Const GWL_WNDPROC = (-4)
'Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_USERDATA = (-21)
Private Const GWL_ID = (-12)

Private Const POINTSPERINCH As Long = 72

Private Const cerrPropertyNotFound As Integer = 3270

' Generally use Hex values but
' Decimal is clearer for educational purposes.
Const Bit10 = 128
Const Bit9 = 128
Const Bit8 = 128
Const Bit7 = 64
Const Bit6 = 32
Const Bit5 = 16
Const Bit4 = 8
Const Bit3 = 4
Const Bit2 = 2
Const Bit1 = 1

' Not used anymore
Private Const NotAColor = &HEF000000

' Leftover DAO props so that Reference to DAO is no longer required.
Private Const dbLongDAO = 4
Private Const dbTextDAO = 10


' Module Level Variables

' Junk vars
Private lngRet As Long
Private blRet As Boolean
Private varTemp As Variant

' Current Y axis Resolution in Dots per Inch
Private lngYdpi As Long

' General purpose RECT structure
Private udtRECT As RECT

' There is a bug in Access when using the SetWindowPos API
' on a Form from an external Class module.
' Specifically if the Form contains a Header or a
' Command Button they are resized along with the Form.
' To get around this we need to store the original
' height of the CommandButton.
Private originalCmdButHeight As Integer

' hWnd for the Month Calendar control
Private m_hWndDTP As Long
' hWnd for the Form used to display the Calendar
'Private m_Form As Access.Form

'Properties used to create the Month Calendar
Private m_MonthColumns  As Integer
Private m_MonthRows  As Integer
Private m_SelectedDate As Date

' To init SetSelectedDateRange
Private m_StartDate As Date
Private m_EndDate As Date

Private m_FontName As String
Private m_FontSize As Integer
Private m_ShowWeekNumbers As Long

' Allow the user to select a Range of Dates
Private m_MultiSelect As Long

' Max number of days in a user selectable range
Private m_MaxSelectRangeofDays As Integer

' Offset from the Top of our Form where we draw the
' Month Calendar. Allows any controls we have placed at the
' top of the Form to be visible.
Private m_CalendarYOffset As Integer

' Month Calendar Color props
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_MonthBackColor As Long
Private m_TitleBackColor As Long
Private m_TitleForeColor As Long
Private m_TrailingForeColor As Long
' Original values of above color props
' that exist at time of Calendar creation.
' These color values are as per the Users Window setting
Private m_BackColorOrig As Long
Private m_ForeColorOrig As Long
Private m_MonthBackColorOrig As Long
Private m_TitleBackColorOrig As Long
Private m_TitleForeColorOrig As Long
Private m_TrailingForeColorOrig As Long

Private m_NoToday As Long
Private m_NoTodayCircle As Long

' Store current Cursor location for this session
Private m_cursorX As Long
Private m_cursorY As Long
Private m_PositionAtCursor As Boolean
' Cursor coords when Calendar is first loaded
Private m_cursorXinitpos As Long
Private m_cursorYinitpos As Long
' Positioning of Calendar Window
Private m_WindowLocation As Long

' Requried to calculate Calendar Height
Dim m_TodayHeight As Long

' Temp vars to hold the size of our Calendar
Private lngTempRight As Long
Private lngTempBottom As Long
Private hFontPrevious As Long

Public m_Hwnd As Long
Private m_hWndForm As Long

' Click once to close the Calendar?
Private m_oneClick As Boolean

' Max 12 Months displayed at once
' plus one month before and after = 14 months
' Allow for years 1000 to 3000
Private BoldDayStates(1000 To 3000, 1 To 14) As MONTHDAYSTATE


Public Sub CreateDTPControl()
' Create the Month Calendar

' Only create MonthCalendar if there is a host window
If m_Hwnd = 0 Then Exit Sub

' Temp vars
Dim udtICC As INITCOMMONCONTROLSEXSTRUCT
Dim lngStyle As Long
Dim hInstance As Long

'Temp handle to Information context
Dim lngIC As Long

If m_hWndDTP <> 0 Then
    Call apiDestroyWindow(m_hWndDTP)
End If

' Courtesy of Ken Getz
hInstance = apiGetWindowLong(Application.hWndAccessApp, GWL_HINSTANCE)
udtICC.dwsize = Len(udtICC)
udtICC.dwICC = ICC_DATE_CLASSES
Call apiInitCommonControlsEx(udtICC)

' The Calendar MultiSelect property determines whether the
' user can select a range of dates ' WS_BORDER OrOr WS_OVERLAPPED
lngStyle = WS_BORDER Or WS_CHILD Or m_MultiSelect _
Or m_ShowWeekNumbers Or m_NoToday Or m_NoTodayCircle Or m_NoToday Or WS_VISIBLE Or MCS_DAYSTATE ' DayState not ready yet! Or MCS_DAYSTATE


' Create our Month Calendar
m_hWndDTP = apiCreateWindowEx(0&, _
                            "SysMonthCal32", _
                            "DateTimePicker", _
                            lngStyle, _
                               0, _
                            0, _
                             250, _
                            250, _
                             m_Hwnd, _
                            0&, _
                            hInstance, _
                            ByVal 0&)
                
' Modified to allow for different screen resolutions
' and printer output. Needed to Calculate Font size
lngIC = apiCreateIC("DISPLAY", vbNullString, vbNullString, vbNullString)
If lngIC <> 0 Then
  lngYdpi = apiGetDeviceCaps(lngIC, LOGPIXELSY)
  apiDeleteDC (lngIC)
Else
  lngYdpi = 120 'Default average value
End If


' Call our Sub to redraw the Form holding our Calendar
' and position the Calendar appropriately.
ReDraw


End Sub


Public Sub ReDraw()
' Redraw and Position the Parent Form and the Calendar

' Only create MonthCalendar if there is a host window
If m_Hwnd = 0 Then Exit Sub

Dim lngXtraPixels As Long
' Make sure the Calendar has time to be initially rendered.
DoEvents
  


' ****************************
' NOT READY YET!!!!
' TEST FOR DAYSTATE
'Me.DayState = 1


' Set max range of days user can select
Call apiSendMessage(m_hWndDTP, MCM_SETMAXSELCOUNT, m_MaxSelectRangeofDays, ByVal 0&)
  
' If EndDate is Set and MultiSelect then call SetSelectedDateRange
If m_EndDate <> 0 And m_EndDate <> 0 And m_MultiSelect = MCS_MULTISELECT Then
    SetSelectedDateRange m_StartDate, m_EndDate
Else

    Dim udtST As SYSTEMTIME
    With udtST
    .wYear = year(m_StartDate)
    .wMonth = month(m_StartDate)
    .wDay = day(m_StartDate)
    End With
    Call apiSendMessage(m_hWndDTP, MCM_SETCURSEL, ByVal 0&, udtST)
End If

' Get the size required to display 1 complete month of the calendar.
Call apiSendMessage(m_hWndDTP, MCM_GETMINREQRECT, ByVal 0&, udtRECT)

' Check is TODAY is showing at bottom of Calendar
If m_NoToday = 0 Then
    lngXtraPixels = apiSendMessage(m_hWndDTP, MCM_GETMAXTODAYWIDTH, ByVal 0&, ByVal 0&)
Else
    lngXtraPixels = udtRECT.Right ' = lngXtraPixels

End If

' Set color props
If m_BackColor <> NotAColor Then
Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_BACKGROUND, ByVal m_BackColor)
End If

If m_ForeColor <> NotAColor Then
 Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TEXT, ByVal m_ForeColor)
End If

If m_MonthBackColor <> NotAColor Then
Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_MONTHBK, ByVal m_MonthBackColor)
End If

If m_TitleForeColor <> NotAColor Then
Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLETEXT, ByVal m_TitleForeColor)
End If

If m_TitleBackColor <> NotAColor Then
    Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TITLEBK, ByVal m_TitleBackColor)
End If

If m_TrailingForeColor <> NotAColor Then
    Call apiSendMessage(m_hWndDTP, MCM_SETCOLOR, MCSC_TRAILINGTEXT, ByVal m_TrailingForeColor)
End If



' THis is the function to tell us what size the window would be
' to hold the entire CLient Rectangle area. In other words it adds
' on the height for the Title Bar, Bottom Border etc.
Dim rc1 As RECT
Dim rc2 As RECT
Dim rc3 As RECT
Dim BarBordersY As Integer
Dim BarBordersX As Integer

' Get rectangle for our Form
'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
lngRet = GetWindowRect(m_hWndDTP, rc1)
lngRet = GetClientRect(m_hWndDTP, rc2)
' Calculate Borders and Title Bar area.
BarBordersY = ((rc1.Bottom - rc1.Top) - (rc2.Bottom - rc2.Top)) + 10
BarBordersX = ((rc1.Right - rc1.Left) - (rc2.Right - rc2.Left)) + 10

With udtRECT
  ' Handle possible extra Today string Width
  udtRECT.Right = lngXtraPixels
  
  ' An OS bug/feature requires us to round up
    If m_NoToday <> 0 Then
   lngTempRight = (.Right * m_MonthColumns) + (m_MonthColumns - 1) * 6 'BarBordersX 'm_MonthColumns * 2 '5
  lngTempBottom = (.Bottom * m_MonthRows) + (m_MonthRows * (m_TodayHeight * 1.5)) ' (m_TodayHeight * 4)  '(m_MonthRows - 1) * 24 '12 'BarBordersY 'm_MonthRows * m_TodayHeight ')) ' + 10 ''m_MonthRows * 15
' ADD CODE TO HELP WITH NUMBER OF ROWS.
' If only 1 run then only need *2 or *3

Else
lngTempRight = (.Right * m_MonthColumns) + (m_MonthColumns - 1) * 6 'BarBordersX '+ m_MonthColumns * 2 '5
  lngTempBottom = (.Bottom * m_MonthRows) ' + (m_MonthRows - 1) * 1 '12 'BarBordersY '+ m_MonthRows * 10
End If

End With
' Resize the Month Calendar to display the user selected
' number of months. The CalendarYOffset is used to allow
' any controls we have placed at the Top of our Form
' to be visible.
' ***DEBUG - BUG FIX ******
' Try to fix visual display bug
' when only 1 month is selected.
' The left most column dissappears when
' when selecting a range of dates
' Add 4 pixels to the COntrol's Width

If m_MonthRows = 1 And m_MonthColumns = 1 Then
Call apiSetWindowPos(m_hWndDTP, 0&, 0&, _
0&, lngTempRight + 4, lngTempBottom, 0&)
Else
Call apiSetWindowPos(m_hWndDTP, 0&, 0&, _
0&, lngTempRight, lngTempBottom, 0&)
End If

' Get rectangle for our Form
'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
'lngRet = GetWindowRect(m_Hwnd, rc1)
'lngRet = GetClientRect(m_Hwnd, rc2)
' Get rectangle for our Calendar
lngRet = GetWindowRect(m_hWndDTP, rc3)


lngRet = GetWindowLong(m_Hwnd, GWL_STYLE)
lngRet = AdjustWindowRect(rc3, lngRet, -1)


 
' Resize our Form to display the entire Calendar
' Position Window at current cursor location.
' We will use the Cursor location when the Window was first
' created, not the current location. User could be
' selecting a Menu when this Redraw Sub is called.
' We also need to allow the user to manually move
' and place this window and respect those
' coordinates for this session.
' So if the current XY WIndow Coords are not the same
' as the stored Mouse Coords then let the current
' Calendars Parent Window coords stand.
' We do this directly in the Windows Procedure
' by resetting the CursorX and CursorY props
' on the Window Move message.

If m_PositionAtCursor Then
    ' Position Window at Cursor Location when
    ' Calendar was created.
    '
 
    ' Also add an option to position cursor itself in middle of calendar window
    ' May 05 - update m_cursorXinitpos and m_cursorYinitpos to reflect current cursor location when this calendar is opened
    ' This is done in the modCalendar module when the parent window for the Calendar is created.
    
    'June 7, 2003 fixing position at cursor logic
   ' PositionAtCursor = True
    
    ' First check to ensure entire calendar window will fit.
    ' If not MAKE IT!
    ' COMING IN NEXT RELEASE!
    
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With

Else
    ' Position Window according to users Menu selections
    ' a) 0 -Pop at cursor location when user activates Calendar
    ' b) 1 -Where they manually move/leave it at
    ' c) 2 -Centered in Access App Window
    ' d) 3 -Centered on entire screen
    ' d) 4 -Top Left Corner
'Debug.Print "m_WindowLocation:" & m_WindowLocation

    Select Case m_WindowLocation


    Case 1
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With

    Case 2
    ' Centered within main Access App Window
    ' Get rectangle for our Calendar
    lngRet = GetWindowRect(Application.hWndAccessApp, rc2)
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, ((rc2.Right - rc2.Left) - (rc3.Right - rc3.Left)) / 2, _
        ((rc2.Bottom - rc2.Top) - (rc3.Bottom - rc3.Top)) / 2, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With
    
    Case 3
    ' Centered within Screen
    ' Get rectangle for our Calendar
    Dim hWndDeskTop As Long
    hWndDeskTop = GetDesktopWindow()
    lngRet = GetWindowRect(hWndDeskTop, rc2)
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, ((rc2.Right - rc2.Left) - (rc3.Right - rc3.Left)) / 2, _
        ((rc2.Bottom - rc2.Top) - (rc3.Bottom - rc3.Top)) / 2, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With

     Case 4
    ' Upper Left Hand Corner of Screen
    ' Get rectangle for our Calendar
    lngRet = GetWindowRect(hWndDeskTop, rc2)
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, 2, _
        2, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With

    Case Else
    With udtRECT
        Call apiSetWindowPos(m_Hwnd, 0&, m_cursorX, m_cursorY, rc3.Right - rc3.Left, _
        rc3.Bottom - rc3.Top, 0&)
    End With
    
    End Select


End If

End Sub

Public Property Let MultiSelect(ms As Boolean)
' Allow user to select a Range of dates instead of just one.
  If ms Then
    m_MultiSelect = MCS_MULTISELECT
  Else
    m_MultiSelect = 0
  End If
' Update custom property
 SetProperty "MultiSelect", dbLongDAO, m_MultiSelect

End Property

Public Property Get MultiSelect() As Boolean
  If m_MultiSelect = MCS_MULTISELECT Then
    MultiSelect = True
  Else
    MultiSelect = False
  End If
End Property

Public Property Let MaxSelectRangeofDays(maxdays As Integer)
' Maximum days user can select in a range of dates.
  If maxdays > 365 Or maxdays < 0 Then
    m_MaxSelectRangeofDays = 7
  Else
    m_MaxSelectRangeofDays = maxdays
  End If
  CreateDTPControl
End Property

Public Property Get MaxSelectRangeofDays() As Integer
   MaxSelectRangeofDays = m_MaxSelectRangeofDays
End Property

Public Property Let CalendarYOffset(YOffset As Integer)
' Offset from the Top of our Form where we will
' draw our Calendat. This will leave us room at the
' the top of our Form for our controls.
' Some boundary checking. We'll arbitrarily use 3 inches for now.
' This prop is expressed in TWIPS!
  If YOffset > 3 * TWIPSPERINCH Or YOffset < 0 Then
    m_CalendarYOffset = 0
  Else
    m_CalendarYOffset = YOffset
  End If
  'CreateDTPControl
End Property

Public Property Get CalendarYOffset() As Integer
CalendarYOffset = m_CalendarYOffset
End Property
   
Public Property Let MonthRows(ByVal newMonthRows As Byte)
' The Month Calendar allows for a maximum of 12 months
' to be displayed at once.
  If newMonthRows > 12 Or newMonthRows < 0 Then
  newMonthRows = 1
  End If
  If newMonthRows * m_MonthColumns > 12 Then
  m_MonthRows = 12 / m_MonthColumns
  Else
    m_MonthRows = newMonthRows
   End If
   
' Update custom property
Call SetProperty("MonthRows", dbLongDAO, m_MonthRows)

End Property

Public Property Get MonthRows() As Byte
  MonthRows = m_MonthRows
End Property

Public Property Let MonthColumns(ByVal newMonthColumns As Byte)
' The Month Calendar allows for a maximum of 12 months
' to be displayed at once.
  If newMonthColumns > 12 Or newMonthColumns < 0 Then
    newMonthColumns = 1
  End If
  
  If newMonthColumns * m_MonthRows > 12 Then
    m_MonthColumns = 12 / m_MonthRows
  Else
    m_MonthColumns = newMonthColumns
  End If
' Update custom property
Call SetProperty("MonthColumns", dbLongDAO, m_MonthColumns)
End Property

Public Property Get MonthColumns() As Byte
  MonthColumns = m_MonthColumns
End Property

Public Property Let WindowLocation(loc As Long)
' Allow user to select position of Calendar Window
  'Debug.Print "Prop Let:" & loc
  If loc < 0 Or loc > 4 Then
    m_WindowLocation = 0
  Else
    m_WindowLocation = loc
  End If
  
' Update custom property
 SetProperty "WindowLocation", dbLongDAO, m_WindowLocation
 'Debug.Print "Saved Prop :" & m_WindowLocation
End Property

Public Property Get WindowLocation() As Long
WindowLocation = m_WindowLocation
End Property


Public Property Let PositionAtCursor(bl As Boolean)
' Does user want to popup Calendar at Current Cursor location
' When they invoke the Calendar?
 Dim pt As POINTAPI
m_PositionAtCursor = bl
If bl Then
'    lngRet = GetCursorPos(pt)
    m_cursorX = m_cursorXinitpos
    m_cursorY = m_cursorYinitpos
End If
End Property


Public Property Get PositionAtCursor() As Boolean
  PositionAtCursor = m_PositionAtCursor
End Property


Public Property Let CursorXinit(x As Long)
' Store Cursor position
    m_cursorXinitpos = x
    If Me.PositionAtCursor = True Then
        m_cursorX = x
    End If
' Update custom property
'Call SetProperty("CursorX", dbLongDAO, m_cursorX)
End Property

Public Property Let CursorYinit(y As Long)
' Store Cursor position
    m_cursorYinitpos = y
    If Me.PositionAtCursor = True Then
        m_cursorY = y
    End If
    ' Redraw the control
CreateDTPControl
' Update custom property
'Call SetProperty("CursorY", dbLongDAO, m_cursorY)
End Property

Public Property Let CursorX(x As Long)
' X pos.
    m_cursorX = x
' Update custom property
Call SetProperty("CursorX", dbLongDAO, m_cursorX)
End Property

Public Property Let CursorY(y As Long)
' Y pos.
    m_cursorY = y
' Update custom property
Call SetProperty("CursorY", dbLongDAO, m_cursorY)
End Property


Public Property Get CursorX() As Long
' X pos.
    CursorX = m_cursorX

End Property

Public Property Get CursorY() As Long
' Y Pos.
    CursorY = m_cursorY

End Property


Public Property Let FontSize(FS As Integer)
' Font Size to render the Calendar with.
If FS = 0 Or FS < 0 Then FS = 8
If FS > 24 Then FS = 8
    m_FontSize = FS
' Update custom property
SetProperty "FontSize", dbLongDAO, m_FontSize
End Property

Public Property Get FontSize() As Integer
  FontSize = m_FontSize
End Property

Public Property Let FontName(fn As String)
' Font Name to render the Calendar with.
    m_FontName = fn
' Update custom property
Call SetProperty("FontName", dbTextDAO, m_FontName)
End Property

Public Property Get FontName() As String
  FontName = m_FontName
End Property

Public Property Let ShowWeekNumbers(bl As Boolean)
' Show Week Numbers
    If bl Then
        m_ShowWeekNumbers = MCS_WEEKNUMBERS
    Else
        m_ShowWeekNumbers = 0
    End If
' Update custom property
Call SetProperty("MCS_WEEKNUMBERS", dbLongDAO, m_ShowWeekNumbers)

' Redraw the control
CreateDTPControl
End Property

Public Property Get ShowWeekNumbers() As Boolean
  If m_ShowWeekNumbers = MCS_WEEKNUMBERS Then
    ShowWeekNumbers = True
  Else
    ShowWeekNumbers = False
  End If
  End Property
  
Public Property Get NoTodayCircle() As Boolean
If m_NoTodayCircle = 0 Then
    NoTodayCircle = False
Else
    NoTodayCircle = True
End If
End Property


  Public Property Let NoTodayCircle(bl As Boolean)
' Show Circle around Today's Date
    If bl Then
        m_NoTodayCircle = MCS_NOTODAYCIRCLE
    Else
        m_NoTodayCircle = 0
    End If

' Update custom property
Call SetProperty("NoTodayCircle", dbLongDAO, m_NoTodayCircle)


' Redraw the control
CreateDTPControl
End Property

Public Property Get NoToday() As Boolean
If m_NoToday = 0 Then
    NoToday = False
Else
    NoToday = True
End If
End Property
  
  Public Property Let NoToday(bl As Boolean)
' Show Today's date at bottom of Calendar
    If bl Then
        m_NoToday = MCS_NOTODAY
    Else
        m_NoToday = 0
    End If
' Update custom property
Call SetProperty("NoToday", dbLongDAO, m_NoToday)


' Redraw the control
CreateDTPControl
End Property


Public Property Let hwnd(h As Long)
' Host Window
   m_Hwnd = h

' ReCreate the control
CreateDTPControl

End Property

Public Property Get hwnd() As Long
' MonthCalendar Parent Window
   hwnd = m_Hwnd
End Property

Public Property Let hWndForm(h As Long)
' Calling Form's Window Handle
   m_hWndForm = h
End Property

Public Property Get hWndForm() As Long
' Calling Form's Window Handle
   hWndForm = m_hWndForm
End Property

  
Public Property Get hWndCal() As Long
' MonthCalendar  Window
   hWndCal = m_hWndDTP
End Property
 
Public Property Let OneClick(bl As Boolean)
' Click Once to Close our Calendar.
    m_oneClick = bl
' Update custom property
Call SetProperty("OneClick", dbLongDAO, m_oneClick)
End Property

Public Property Get OneClick() As Boolean
 OneClick = m_oneClick
End Property

Public Property Let SelectedDate(ByVal newSelectedDate As Date)
' Date pre-selected in the Calendar.
' Default is Today's date.
Dim udtST As SYSTEMTIME
  m_SelectedDate = newSelectedDate
  With udtST
    .wYear = year(newSelectedDate)
    .wMonth = month(newSelectedDate)
    .wDay = day(newSelectedDate)
  End With
  
  m_StartDate = newSelectedDate
  If m_MultiSelect <> MCS_MULTISELECT Then
    Call apiSendMessage(m_hWndDTP, MCM_SETCURSEL, ByVal 0&, udtST)
  Else
    m_StartDate = newSelectedDate
    m_EndDate = m_StartDate
    Call SetSelectedDateRange(m_StartDate, m_EndDate)
' Sets the
  End If

End Property


Public Sub SetSelectedDateRange(ByVal STARTDATE As Date, ByVal EndDate As Date)
' Sets the Range of Dates visible in the Calendar
Dim udtST(1) As SYSTEMTIME

If m_MultiSelect <> MCS_MULTISELECT Then Exit Sub

' Fill in the start and ending SYSTEMTIME structures
With udtST(0)
.wDay = day(STARTDATE)
.wMonth = month(STARTDATE)
.wYear = year(STARTDATE)
End With

With udtST(1)
.wDay = day(EndDate)
.wMonth = month(EndDate)
.wYear = year(EndDate)
End With
lngRet = apiSendMessage(m_hWndDTP, MCM_SETSELRANGE, ByVal 0&, udtST(0))
  With udtST(0)
    m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
  End With
  'With udtST(1)
   ' m_endSelectedDate = DateSerial(.wYear, .wMonth, .wDay)
  'End With
  
  
End Sub


Public Property Get StartSelectedDate() As Date
' Either the current single date selected by the user
' or the First date in a user selected range of dates.
Dim udtST As SYSTEMTIME

If m_MultiSelect = MCS_MULTISELECT Then
StartSelectedDate = GetSelectedDates(True)

Else
  Call apiSendMessage(m_hWndDTP, MCM_GETCURSEL, ByVal 0&, udtST)
  With udtST
    m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
  End With
  StartSelectedDate = m_SelectedDate
End If

End Property

Public Property Get EndSelectedDate() As Date
' The End date in a user selected range of dates.
Dim udtST As SYSTEMTIME

If m_MultiSelect = MCS_MULTISELECT Then
EndSelectedDate = GetSelectedDates(False)

Else
  Call apiSendMessage(m_hWndDTP, MCM_GETCURSEL, ByVal 0&, udtST)
  With udtST
    m_SelectedDate = DateSerial(.wYear, .wMonth, .wDay)
  End With
  EndSelectedDate = m_SelectedDate
End If

End Property

Private Function GetSelectedDates(StartEnd As Boolean) As Date
' Return Start or END date of user selected Range
' If StartEnd = TRUE then return Start Date
' If StartEnd = FALSE then return END Date
Dim arrayTime(0 To 1) As SYSTEMTIME
Dim DateStart As Date
Dim DateEnd As Date

' Get the currently selected range
lngRet = apiSendMessage(m_hWndDTP, MCM_GETSELRANGE&, ByVal 0&, arrayTime(0))

' Convert to our Date format
With arrayTime(0)
    DateStart = DateSerial(.wYear, .wMonth, .wDay)
End With
With arrayTime(1)
    DateEnd = DateSerial(.wYear, .wMonth, .wDay)
End With

' Return specified Start or End date
If StartEnd Then
    GetSelectedDates = DateStart
Else
    GetSelectedDates = DateEnd
End If
End Function

Public Sub SetViewableMonths(numMonths As Integer)
' How many months are viewable in the Calendar.
' Options are 1,2,3,4,6,9,12
   
Select Case numMonths
    Case 1
    MonthColumns = 1
    MonthRows = 1

    Case 2
    MonthColumns = 2
    MonthRows = 1

    Case 3
    MonthColumns = 3
    MonthRows = 1
    
    Case 4
    MonthColumns = 2
    MonthRows = 2

    Case 6
    MonthColumns = 3
    MonthRows = 2
    
    Case 8
    MonthColumns = 4
    MonthRows = 2
    
    Case 9
    MonthColumns = 3
    MonthRows = 3
    
    Case 12
    MonthColumns = 4
    MonthRows = 3
    
Case Else
    MonthColumns = 1
    MonthRows = 1
        
End Select
   
' Create the new Month Calender
If m_hWndDTP = 0 Then
CreateDTPControl
Else
ReDraw
End If

End Sub


Public Function GetDAYSTATE(ByVal year As Long, ByVal month As Long) As Long
   GetDAYSTATE = BoldDayStates(year, month).lpMONTHDAYSTATE
End Function


' Turn on Bolding for specified days of the month
' Use named arguments when calling this function to
' cut down on your typing.
' ie. SetBoldDayState numYear 2004, numMonth:=1, day = 19
Public Function SetBoldDayState(numYear As Long, numMonth As Long, day As Long, Optional ResetMonth As Boolean = False)   '_
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 
' Optional One As Boolean, Optional Two As Boolean, _
' Optional Three As Boolean, Optional Four As Boolean, _
' Optional Five As Boolean, Optional Six As Boolean, _
' Optional Seven As Boolean, Optional Eight As Boolean, _
' Optional Nine As Boolean, Optional Ten As Boolean, _
' Optional Eleven As Boolean, Optional Twelve As Boolean, _
' Optional Thirteen As Boolean, Optional Fourteen As Boolean, _
' Optional Fifthteen As Boolean, Optional Sixteen As Boolean, _
' Optional Seventeen As Boolean, Optional Eighteen As Boolean, _
' Optional NineTeen As Boolean, Optional Twenty As Boolean, _
' Optional TwentyOne As Boolean, Optional TwentyTwo As Boolean, _
' Optional TwentyThree As Boolean, Optional TwentyFour As Boolean, _
' Optional TwentyFive As Boolean, Optional TwentySix As Boolean, _
' Optional TwentySeven As Boolean, Optional TwentyEight As Boolean, _
' Optional TwentyNine As Boolean, Optional Thirty As Boolean, _
' Optional ThirtyOne As Boolean, Optional ThirtyTwo As Boolean)
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 

'
' Init the month we are working with.
' Only clear this var if ResetMonth = TrUE
If ResetMonth Then
BoldDayStates(numYear, numMonth).lpMONTHDAYSTATE = 0
End If

If day > 0 And day < 31 Then day = day - 1
'BoldDayStates(numMonth) = BoldDays
BoldDayStates(numYear, numMonth).lpMONTHDAYSTATE = BoldDayStates(numYear, numMonth).lpMONTHDAYSTATE Or 2 ^ day
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 
' If One Then BoldDayStates(numMonth).lpMONTHDAYSTATE = 1 'BoldDayStates(numMonth).lpMONTHDAYSTATE Or 1
' If Two Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 1
' If Three Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 2
' If Four Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 3
' If Five Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 4
' If Six Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 5
' If Seven Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 6
' If Eight Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 7
' If Nine Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 8
' If Ten Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 9
' If Eleven Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 10
' If Twelve Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 11
' If Thirteen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 12
' If Fourteen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 13
' If Fifthteen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 14
' If Sixteen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 15
' If Seventeen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 16
' If Eighteen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 17
' If NineTeen Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 18
' 
' If Twenty Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 19
' If TwentyOne Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 20
' If TwentyTwo Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 21
' If TwentyThree Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 22
' If TwentyFour Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 23
' If TwentyFive Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 24
' If TwentySix Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 25
' If TwentySeven Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 26
' If TwentyEight Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 27
' If TwentyNine Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 28
' 
' If Thirty Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 29
' If ThirtyOne Then BoldDayStates(numMonth).lpMONTHDAYSTATE = BoldDayStates(numMonth).lpMONTHDAYSTATE Or 2 ^ 30
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 


End Function


Private Sub Class_Initialize()
' Number of pixels to leave at the Top of the form before we
' begin to draw our Calendar. This leaves room for any
' controls we have at the top of the form.
Me.CalendarYOffset = 20

' Default Number of Columns and Rows
' 6 month display
'Me.MonthColumns = 3
'Me.MonthRows = 2
' Check and see if Custom Prop exists
blRet = GetProperty("MonthColumns", varTemp)
If blRet Then
m_MonthColumns = varTemp
Else
m_MonthColumns = 3
End If

blRet = GetProperty("MonthRows", varTemp)
If blRet Then
m_MonthRows = varTemp
Else
m_MonthRows = 2
End If


' Default Font stuff
' Check and see if Custom Prop exists
blRet = GetProperty("FontName", varTemp)
If blRet Then
Me.FontName = varTemp
Else
Me.FontName = "Arial"
End If

blRet = GetProperty("FontSize", varTemp)
If blRet Then
Me.FontSize = varTemp
Else
Me.FontSize = 8
End If

hFontPrevious = 0

blRet = GetProperty("MCS_MULTISELECT", varTemp)
If blRet Then
MultiSelect = varTemp
Else
' MultiSelect to start
MultiSelect = MCS_MULTISELECT '0 to Turn Off
End If

blRet = GetProperty("MCS_WEEKNUMBERS", varTemp)
If blRet Then
ShowWeekNumbers = varTemp
Else
' MultiSelect to start
ShowWeekNumbers = 0 'MCS_WEEKNUMBERS '0 to Turn Off
End If


blRet = GetProperty("NoToday", varTemp)
If blRet Then
m_NoToday = varTemp
Else
' Show Today is turned off
 m_NoToday = 0 'MCS_NOTODAY '0 to Turn Off
End If

blRet = GetProperty("NoTodayCircle", varTemp)
If blRet Then
m_NoTodayCircle = varTemp
Else
' Show Today is turned off
 m_NoToday = 0 'MCS_NOTODAYCIRCLE '0 to Turn Off
End If


 blRet = GetProperty("WindowLocation", varTemp)
    If blRet Then
    m_WindowLocation = varTemp
    Else
    ' Window Location
     m_WindowLocation = 0
    End If
    
    ' June 7, 2003
    ' Fixing PositionAtCursor logic
    If varTemp = 0 Then
    ' Init
        Me.PositionAtCursor = True
    Else
        Me.PositionAtCursor = False
    End If

' Only load if PositionAtCursor is False
    If Not Me.PositionAtCursor Then
    blRet = GetProperty("Cursorx", varTemp)
    If blRet Then
    m_cursorX = varTemp
    Else
    ' Leave Alone!
     m_cursorX = 0
    End If
End If

 'Only load if PositionAtCursor is False
If Not Me.PositionAtCursor Then
    blRet = GetProperty("CursorY", varTemp)
    If blRet Then
        m_cursorY = varTemp
    Else
        ' Get current Cursor Position!
        'Dim pt As POINTAPI
        'lngRet = GetCursorPos(pt)
        ' m_cursorXinit = pt.x
        ' m_cursorYinit = pt.y
        ' m_cursorX = pt.x
         'm_cursorY = pt.y
         m_cursorY = 0
    End If
End If


' Check our Color props.
' Set var to NotAColor constant if prop doesn't exist
blRet = GetProperty("BackColor", varTemp)
If blRet Then
m_BackColor = varTemp
Else
' Set var to nonexistent color
m_BackColor = NotAColor
End If

blRet = GetProperty("MonthBackColor", varTemp)
If blRet Then
m_MonthBackColor = varTemp
Else
' Set var to nonexistent color
m_MonthBackColor = NotAColor
End If


blRet = GetProperty("ForeColor", varTemp)
If blRet Then
m_ForeColor = varTemp
Else
' Set var to nonexistent color
 m_ForeColor = NotAColor
End If

blRet = GetProperty("TitleBackColor", varTemp)
If blRet Then
m_TitleBackColor = varTemp
Else
' Set var to nonexistent color
m_TitleBackColor = NotAColor
End If

blRet = GetProperty("TitleForeColor", varTemp)
If blRet Then
m_TitleForeColor = varTemp
Else
' Set var to nonexistent color
m_TitleForeColor = NotAColor
End If

blRet = GetProperty("TrailingForeColor", varTemp)
If blRet Then
m_TrailingForeColor = varTemp
Else
' Set var to nonexistent color
m_TrailingForeColor = NotAColor
End If

blRet = GetProperty("OneClick", varTemp)
If blRet Then
m_oneClick = varTemp
Else
' Set var to nonexistent color
m_oneClick = True
End If


' Max number of days user can select in a continuous range
m_MaxSelectRangeofDays = 31
' To select a Range of Dates add code like below
'm_EndDate = #2/27/00#
'm_StartDate = #2/21/00#

' Set DayState
' For this example we will set one day per month
' corresponding to the Month number
Erase BoldDayStates
' Here is an example of how to Set the BoldDayState prop for Jan 19,2004.
' ie. SetBoldDayState numYear 2004, numMonth:=1, day = 19

End Sub

Public Sub Class_Terminate()
' Cleanup after ourselves.
On Error Resume Next


'Debug.Print "start ClsMonthCalendat Terminated:" & Now
' March 05 Debug
' Not getting a clean exit when user Closes
' Calling Form with Calendar still Open
If m_hWndDTP <> 0 Then
lngRet = apiDestroyWindow(m_hWndDTP)
'Debug.Print "Destroyed m_hWndDTP:" & lngRet
End If

' Wait until Window is definately gone!
lngRet = apiDestroyWindow(Me.m_Hwnd)
'Debug.Print "Destroyed Me.m_Hwnd:" & lngRet
DoEvents

 Const CLASSNAME = "MonthCalendar"
 Const TITLE = "Month Calendar"
Dim hWndTemp As Long

' If this window already exists then exit!
hWndTemp = 1
Do While hWndTemp <> 0
hWndTemp = FindWindow(CLASSNAME, TITLE)
'Debug.Print "MonthCal Hwnd -hWndTemp:" & hWndTemp

Loop

' See if we have a new Font handle hanging around.
If hFontPrevious <> 0 Then
'Debug.Print "deleting FOnt"
DeleteObject hFontPrevious
End If

'Debug.Print "End ClsMonthCalendar Terminated:" & Now
End Sub

' *** CODE END ***


Private Function GetProperty(ByVal strPropName As String, ByRef strPropValue As Variant) As Boolean
' Changed strPropValue as String to as Variant

Const cProcedureName As String = "GetProperty"
On Error GoTo Err_Handler

Dim sDefault As String

Select Case strPropName
    
    Case "BackColor"
    sDefault = vbWhite
    
    Case "FontName"
    sDefault = "Arial"
    
    Case "FontSize"
    sDefault = 8
    
    Case "ForeColor"
    sDefault = GetSysColor(COLOR_WINDOWTEXT)
    
    Case "MonthBackColor"
    sDefault = vbWhite
    
    Case "MonthColumns"
    sDefault = 1
    
    Case MultiSelect
    sDefault = "0" 'false
    
    Case "NoToday"
    sDefault = MCS_NOTODAY
    'Debug.Print "NoToday: " & sDefault
    
    Case "NoTodayCircle"
    sDefault = MCS_NOTODAYCIRCLE
    'Debug.Print "NoTodayCircle: " & sDefault
    
    Case "OneClick"
    sDefault = "0" ' false 'was = 1 changed 4/15/2005 ALD
    'Debug.Print "OneClick: " & sDefault
    
    Case "MCS_WEEKNUMBERS"
    sDefault = MCS_WEEKNUMBERS
    'Debug.Print "WEEK: " & sDefault
    
    Case "MonthRows"
    sDefault = 1
    
    Case "TitleBackColor"
    sDefault = GetSysColor(COLOR_ACTIVECAPTION)
    
    Case "TitleForeColor"
    sDefault = vbWhite 'GetSysColor()  'vbTitleBarText
    
    Case "TrailingForeColor"
    sDefault = GetSysColor(COLOR_GRAYTEXT) '  apiSendMessage(m_hWndDTP, MCM_GETCOLOR, MCSC_TRAILINGTEXT, 0&)
    
    Case "WindowLocation"
    ' Position Window according to users Menu selections
    ' a) 0 -Pop at cursor location when user activates Calendar
    ' b) 1 -Where they manually move/leave it at
    ' c) 2 -Centered in Access App Window
    ' d) 3 -Centered on entire screen
    ' d) 4 -Top Left Corner
    sDefault = 2
    
    
    
    Case Else
    ' not possible...I hope!
    sDefault = ""
    GetProperty = False
    Exit Function
    
End Select

strPropValue = GetSetting("MonthCalendarSL", "CalendarProperties", strPropName, sDefault)
GetProperty = True

Exit_Sub:
  On Error GoTo 0
  Exit Function
Err_Handler:
  Resume Exit_Sub
End Function


Sub SetProperty(ByVal strPropName As String, ByVal varPropType As Variant, _
                ByVal varPropValue As Variant)
Const cProcedureName As String = "SetProperty"
On Error GoTo Err_Handler

SaveSetting "MonthCalendarSL", "CalendarProperties", strPropName, varPropValue

Exit_Sub:
  On Error GoTo 0
  Exit Sub
Err_Handler:
  Resume Exit_Sub

End Sub



'Private Function GetProperty(ByVal strPropName As String, ByRef strPropValue As Variant) As Boolean
'' Changed strPropValue as String to as Variant
'
'Const cProcedureName As String = "GetProperty"
'On Error GoTo Err_Handler
'Dim db As DAO.Database
'
'    Set db = CurrentDb
'    strPropValue = db.Properties(strPropName)
'    GetProperty = True
'
'Exit_Sub:
'  On Error GoTo 0
'  Set db = Nothing
'  Exit Function
'Err_Handler:
'  GetProperty = False
'  Select Case Err
'  Case cerrPropertyNotFound
'  Case Else
'   ' Call LogError(Err.Number, Err.Description, cModuleName & cProcedureName)
'  End Select
'  Resume Exit_Sub
'End Function





'Sub SetProperty(ByVal strPropName As String, ByVal varPropType As Variant, _
'                ByVal varPropValue As Variant)
'Const cProcedureName As String = "SetProperty"
'On Error GoTo Err_Handler
'Dim db As DAO.Database
'Dim prp As DAO.Property
'
'    Set db = DBEngine(0)(0)
'    db.Properties(strPropName).Value = varPropValue
'
'Exit_Sub:
'  On Error GoTo 0
'  Set prp = Nothing
'  Set db = Nothing
'  Exit Sub
'Err_Handler:
'  Select Case Err
'  Case cerrPropertyNotFound
'    Set prp = db.CreateProperty(strPropName, varPropType, varPropValue)
'    db.Properties.Append prp
'  Case Else
'    'Call LogError(Err.Number, Err.Description, cModuleName & cProcedureName)
'  End Select
'  Resume Exit_Sub
'
'End Sub




'typedef DWORD MONTHDAYSTATE, FAR * LPMONTHDAYSTATE;
'// BOOL MonthCal_SetDayState(HWND hmc, int cbds, DAYSTATE *rgds)
'//   cbds is the count of DAYSTATE items in rgds and it must be equal
'//   to the value returned from MonthCal_GetMonthRange(hmc, GMR_DAYSTATE, NULL)
'//   This sets the DAYSTATE bits for each month (grayed and non-grayed
'//   days) displayed in the calendar. The first bit in a month's DAYSTATE
'//   corresponts to bolding day 1, the second bit affects day 2, etc.
'#define MCM_SETDAYSTATE     (MCM_FIRST + 8)
'#define MonthCal_SetDayState(hmc, cbds, rgds)
'SNDMSG(hmc, MCM_SETDAYSTATE, (WPARAM)(cbds), (LPARAM)(rgds))



'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 
'         lngStyle = GetWindowLong(m_hWndDTP, GWL_STYLE)
'         'We can Set/Reset these styles with recreating the Control:
'         lngStyle = lngStyle And Not (MCS_WEEKNUMBERS Or MCS_NOTODAY Or MCS_NOTODAYCIRCLE)
'         If m_WeekNumbers Then lngStyle = lngStyle Or MCS_WEEKNUMBERS
'         If Not m_ShowToday Then lngStyle = lngStyle Or MCS_NOTODAY
'         If Not m_TodayCircle Then lngStyle = lngStyle Or MCS_NOTODAYCIRCLE
'         SetWindowLong m_hWndDTP, GWL_STYLE, lngStyle
'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 

' ************************
' DEBUG GPF
' ************************
'Public Property Get xx() As Long
'xx = 1
'End Property

Public Function IsCalendar() As Boolean
Const CLASSNAME = "MonthCalendar"
 Const TITLE = "Month Calendar"
Dim hWndTemp As Long

' If this window already exists then exit!
hWndTemp = 0
hWndTemp = FindWindow(CLASSNAME, TITLE)
If hWndTemp <> 0 Then
IsCalendar = True
 Call Beep(12000, 1000)

Else
IsCalendar = False
End If
End Function
