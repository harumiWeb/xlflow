Attribute VB_Name = "modCalendar"

'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 and Access2K
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.
'           Please feel free to use this code within your own
'           projects whether they are private or commercial applications
'           without obligation.
'           This code may not be resold by itself or as part of a collection.
'
'Name:      modCalendar
'
'Version:   2.05
'
'Purpose:
'           Create a Window with a Window Procedure to house the
'           Month Calendar control. Further provide a Menu interface
'           to allow the user to modify the Calendar's properties.
'           The Window procedure must reside in a standard Code module.
' 
'Author:    Stephen Lebans
'
'Email:     Stephen@lebans.com
'
'Web Site:  www.lebans.com
'
'Date:      Dec 09, 2004, 11:11:11 PM
'
'Credits:   Based on code by:
'           Ray Mercer - Window Creation & Messaging in VB
'           Ken Getz & Michael Kaplan - AddrOf
'           Charles Petzold - Window Creation and Message loops
'           Dev Ashish - AddrOf implementation - Access Version checking
'           Pedro Gil - Initial framework and props
'           MSDN KB
'
'BUGS:      Fixed the bug that appears as a result of  Access
'           is caching the WinProc.
'           Added a call to UnregisterClass to resolve the issue.
'
'What's Missing:
'           You tell me!
'
'           Proper Error handling.
'
'How it Works:
'           The Month Calendar is created directly with the
'           API's contained in the Common Controls DLL. In this manner we bypass
'           the DatePicker ActiveX control, which is simply a wrapper for these
'           calls anyway. This removes any problems from distribution and
'           especially version issues of using the ActiveX control.
'
' This is the 10th major release.

' To exit from the Window Procedure,
' thereby closing the MonthCalendar Control,
' you can either:
'1) Press the Escape Key
'2) Click on the Window's Close Button(x)
'3) Double Click or Single Click the Left Mouse Button on a Date
'   It depends on your settings for the Calendar Properties

' ****************************************************
'
'               WARNING
'
' If you place a Breakpoint within the Window Procedure
' you will cause a GPF!
'
' ****************************************************


Option Explicit
Option Compare Database

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type


Type POINTAPI
        x As Long
        y As Long
End Type

Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lparam As Long
    time As Long
    pt As POINTAPI
End Type

Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

'// bit-packed array of "bold" info for a month
'// if a bit is on, that day is drawn bold
Private Type MONTHDAYSTATE
  lpMONTHDAYSTATE As Long
  ' SHould really be array of 4 bytes because
  ' of VB's Signed datatypes
End Type

' Control Message Header
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long 'Integer
    End Type

'The actual Date/Time values are stored this way
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

' MonthCalendar SelectChange
Private Type NMSELCHANGE
nm As NMHDR
stSelStart As SYSTEMTIME
stSelEnd As SYSTEMTIME
End Type

' DayState Header
Private Type NMDAYSTATE
    nmhd As NMHDR ' // this must be first, so we don't break WM_NOTIFY
    stStart As SYSTEMTIME
    cDayState As Long ' F0r ease of use always specify 12 months of data
    prgDayState As Long 'MONTHDAYSTATE '; // points to cDayState MONTHDAYSTATEs
End Type


Private Type MCHITTESTINFO
        cbSize As Long
        pt As POINTAPI
        uHit As Long
        st As SYSTEMTIME
End Type


' ********************************
' VB6 RUNTIMES must be present to resolve this call
' Returns address of the address of the associated SafeArray descriptor
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ptr() As Any) As Long
'*********************************************************************

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function GetDoubleClickTime Lib "user32" () As Long

Private Declare Function GetMessageTime Lib "user32" () As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
(ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
(ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal length As Long)
                                                  
Private Declare Function setWindowtext Lib "user32" Alias "SetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
(ByVal hwnd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

 Private Declare Function apiSendMessage Lib "user32" _
  Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lparam As Any) As Long

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
ByVal dwExStyle As Long, _
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
                                                    
Private Declare Function ClientToScreen Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long
                    
Private Declare Function PostMessageString Lib "user32" Alias "PostMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lparam As String) As Long
                                                      
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Declare Function CreatePopupMenu Lib "user32" () As Long

Private Declare Function CreateMenu Lib "user32" () As Long

Private Declare Function CheckMenuItem Lib "user32" _
(ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
                                                                                                        
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" _
(ByVal hInstance As Long, ByVal lpIconName As String) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
(ByVal hInstance As Long, ByVal lpCursorName As String) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" _
(pcWndClassEx As WNDCLASSEX) As Integer

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" _
(lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long

Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" _
(lpMsg As Msg) As Long

Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Private Declare Function BeginPaint Lib "user32" _
(ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" _
(ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
(ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
lpRect As RECT, ByVal wFormat As Long) As Long
                                                                    
Private Declare Function apiGetWindowLong Lib "user32" _
  Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias _
"FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function UnregisterClass Lib "user32" _
Alias "UnregisterClassA" (ByVal lpClassName As String, _
ByVal hInstance As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function MessageBeep Lib "user32" _
Alias "BeepA" (ByVal wType As Long) As Long

Private Declare Function Beep Lib "kernel32" _
(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' Enable/Disable Main Access Window
Private Declare Function EnableWindow Lib "user32" _
(ByVal hwnd As Long, ByVal fEnable As Long) As Long

Private Declare Function IsWindowEnabled Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32" _
(ByVal hWndLock As Long) As Long

Private Declare Function SetCapture Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

' Button Control Styles
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_DEFPUSHBUTTON = &H1&
Private Const BS_CHECKBOX = &H2&
Private Const BS_AUTOCHECKBOX = &H3&
Private Const BS_RADIOBUTTON = &H4&
Private Const BS_3STATE = &H5&
Private Const BS_AUTO3STATE = &H6&
Private Const BS_GROUPBOX = &H7&
Private Const BS_USERBUTTON = &H8&
Private Const BS_AUTORADIOBUTTON = &H9&
Private Const BS_OWNERDRAW = &HB&
Private Const BS_LEFTTEXT = &H20&

' User Button Notification Codes
Private Const BN_CLICKED = 0
Private Const BN_PAINT = 1
Private Const BN_HILITE = 2
Private Const BN_UNHILITE = 3
Private Const BN_DISABLE = 4
Private Const BN_DOUBLECLICKED = 5

' Button Control Messages
Private Const BM_GETCHECK = &HF0
Private Const BM_SETCHECK = &HF1
Private Const BM_GETSTATE = &HF2
Private Const BM_SETSTATE = &HF3
Private Const BM_SETSTYLE = &HF4

' CONSTANTS
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_DEADCHAR = &H103
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_SYSCHAR = &H106

' GetWindowLong  / SetWindowLong
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_STYLE = (-16)

Private Const MF_ENABLED = &H0&
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZE As Long = &H1000000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_BORDER As Long = &H800000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_ACCEPTFILES As Long = &H10&
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_GROUP As Long = &H20000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_ICONIC As Long = WS_MINIMIZE
Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_POPUP As Long = &H80000000
Private Const WS_POPUPWINDOW As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_SIZEBOX As Long = WS_THICKFRAME
Private Const WS_TILED As Long = WS_OVERLAPPED
Private Const WS_TILEDWINDOW As Long = WS_OVERLAPPEDWINDOW
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const CS_HREDRAW As Long = &H2
Private Const CS_VREDRAW As Long = &H1
Private Const IDI_APPLICATION As Long = 32512&
Private Const IDC_ARROW As Long = 32512&
Private Const WHITE_BRUSH As Integer = 0
Private Const BLACK_BRUSH As Integer = 4

Private Const WM_CLOSE As Long = &H10
Private Const WM_DESTROY As Long = &H2
Private Const WM_PAINT As Long = &HF
Private Const WM_NOTIFY = &H4E
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_SETTEXT = &HC
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_ENTERIDLE = &H121



' ShowWindow() Commands
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

' Window Message
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5


Private Const WM_ENABLE = &HA
Private Const WM_SETREDRAW = &HB

' Virtual Keys, Standard Set
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_CANCEL = &H3
Private Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Private Const VK_BACK = &H8
Private Const VK_TAB = &H9

Private Const VK_CLEAR = &HC
Private Const VK_RETURN = &HD

Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12
Private Const VK_PAUSE = &H13
Private Const VK_CAPITAL = &H14

Private Const VK_ESCAPE = &H1B
Private Const VK_SPACE = &H20
Private Const VK_PRIOR = &H21
Private Const VK_NEXT = &H22
Private Const VK_END = &H23
Private Const VK_HOME = &H24
Private Const VK_LEFT = &H25
Private Const VK_UP = &H26
Private Const VK_RIGHT = &H27
Private Const VK_DOWN = &H28

Private Const MB_ICONHAND = &H10&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONINFORMATION = MB_ICONASTERISK
Private Const MB_ICONSTOP = MB_ICONHAND

Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_USECHECKBITMAPS = &H200&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_SEPARATOR = &H800&

Private Const NM_FIRST = 0   '  // generic to all controls
Private Const NM_LAST = -99
Private Const NM_RELEASEDCAPTURE = (NM_FIRST - 16)

Private Const MF_BYPOSITION = &H400&
Private Const MF_POPUP = &H10&
Private Const MF_STRING = &H0&
Private Const GWL_WNDPROC = (-4)
Private Const WM_COMMAND = &H111

Private Const DTN_FIRST  As Long = -760
Private Const DTN_LAST  As Long = -799

Private Const MCN_FIRST  As Long = -750
Private Const MCN_LAST As Long = -799
Private Const MCN_GETDAYSTATE As Long = (MCN_FIRST + 3)

Private Const MCN_SELECT As Long = (MCN_FIRST + 4)
Private Const MCN_SELCHANGE As Long = (MCN_FIRST + 1)

Private Const NM_KEYDOWN = (NM_FIRST - 15)
Private Const NM_DBLCLK = (NM_FIRST - 3)


'Color part's of the Calendar
Private Const MCSC_BACKGROUND = 0    '// the background color (between months)
Private Const MCSC_TEXT = 1          '// the dates
Private Const MCSC_TITLEBK = 2       '// background of the title
Private Const MCSC_TITLETEXT = 3
Private Const MCSC_MONTHBK = 4       '// background within the month cal
Private Const MCSC_TRAILINGTEXT = 5  '/

Private Const MCM_FIRST = &H1000&
Private Const MCM_HITTEST = MCM_FIRST + 14
        
Private Const MCHT_TITLE = &H10000
Private Const MCHT_CALENDAR = &H20000
Private Const MCHT_TODAYLINK = &H30000

Private Const MCHT_NEXT = &H1000000
'// these indicate that hitting
Private Const MCHT_PREV = &H2000000
'// here will go to the next/prev month

Private Const MCHT_NOWHERE = &H0

Private Const MCHT_TITLEBK = (MCHT_TITLE)
Private Const MCHT_TITLEMONTH = (MCHT_TITLE Or &H1)
Private Const MCHT_TITLEYEAR = (MCHT_TITLE Or &H2)
Private Const MCHT_TITLEBTNNEXT = (MCHT_TITLE Or MCHT_NEXT Or &H3)
Private Const MCHT_TITLEBTNPREV = (MCHT_TITLE Or MCHT_PREV Or &H3)

Private Const MCHT_CALENDARBK = (MCHT_CALENDAR)
Private Const MCHT_CALENDARDATE = (MCHT_CALENDAR Or &H1)
Private Const MCHT_CALENDARDATENEXT = (MCHT_CALENDARDATE Or MCHT_NEXT)
Private Const MCHT_CALENDARDATEPREV = (MCHT_CALENDARDATE Or MCHT_PREV)
Private Const MCHT_CALENDARDAY = (MCHT_CALENDAR Or &H2)
Private Const MCHT_CALENDARWEEKNUM = (MCHT_CALENDAR Or &H3)


' We'll translate above color indexes into Menu ID's
' by adding 1000 to the values

' MISC Properties

'Use Single Or Double Click to Select a Date
Private Const SingleOrDouble = 720
Private Const SingleClick = 721
Private Const DoubleClick = 722


'Show Week Numbers
Private Const ShowWeekNum = 700
Private Const ShowWeekNumYES = 701
Private Const ShowWeekNumNO = 702

'Show Today TodayNumbers
Private Const ShowToday = 705
Private Const ShowTodayYES = 706
Private Const ShowTodayNO = 707

'Show CircleToday Numbers
Private Const ShowcircleToday = 708
Private Const ShowCircleTodayYES = 709
Private Const ShowCircleTodayNO = 710

' Font Dialog Menu
Private Const FontDialog = 820

' Font Size Menu
Private Const Fontx5 = 805

' Weeks Menu
Private Const Monthx1 = 901
Private Const Monthx2 = 902
Private Const Monthx3 = 903
Private Const Monthx4 = 904
Private Const Monthx6 = 906
Private Const Monthx8 = 908
Private Const Monthx9 = 909
Private Const Monthx12 = 912

' WindowPosition menu
Private Const Positionx0 = 920
Private Const Positionx1 = 921
Private Const Positionx2 = 922
Private Const Positionx3 = 923
Private Const Positionx4 = 924
Private Const Positionx5 = 925
Private Const Positionx6 = 926
Private Const Positionx7 = 927
Private Const Positionx8 = 928

Private Const CLASSNAME = "MonthCalendar"
Private Const TITLE = "Calendar"


' Variables to store our dynamic menu's item IDs
Private Menu1 As Long
Private Menu2 As Long
Private Menu3 As Long
Private Menu4 As Long
Private Menu5 As Long
Private Menu6 As Long
Private Menu7 As Long

' Junk Vars
Private lngRet As Long
Private lngTemp As Long

' Module level var to hold handle to our Calendars hWnd
Private hWndCalendar As Long

' Module level var to hold reference to our Calendar object
' We need this to access the Class from the WindowProc function
Private mc As clsMonthCal

' Module level variable to hold the currently selected date
Private SelectedDate As Date

' Module level variables to hold local copy of
' the currently selected Starting and Ending date Ranges
Private localStartSelectedDate As Date
Private localEndSelectedDate As Date

' Module Var to track whether a Font or Color
' Dialog window is currently Open.
Private blDialogOpen As Boolean


' Required to be Module level in order for WindowProc to have access to Menu handles
Dim hMenu As Long
Dim hMenuPop As Long
Dim hMenuPopMisc As Long
Dim hMenuPopTime As Long
Dim hMenuPopMiscShowWeekNumbers As Long
Dim hMenuPopTimeValueBegin As Long
Dim hMenuPopTimeValueEnd As Long
Dim hMenuPopMiscFont  As Long
Dim hMenuPopMiscColor As Long
Dim hMenuPopMiscToday  As Long
Dim hMenuPopMiscCircleToday  As Long
Dim hMenuPopMiscWindowPosition  As Long
Dim hMenuPopMiscOneClick  As Long

' To allow for Keyboard selection of Date(s)
Public SelChangeDateStart As Date
Dim SelChangeDateEnd As Date

' for the filldate subroutine
Public gdteFillCtrl As String
' for the time part of the date
Public gdteTimeValue As String

Public Function ShowMonthCalendar(ByRef clsMC As clsMonthCal, _
ByRef StartSelectedDate As Date, _
Optional ByRef EndSelectedDate As Date = 0) As Boolean

' ************************************************************
' March 22, 2004
' Major modification to the function logic including calling Parameter order.
' Changed function to return Boolean FALSE and "StartSelectedDate =0"
' if user did not select a date from the MonthCalendar.
' The hWndForm param is no longer optional.
' ************************************************************

' ********* WARNING *************
' In order for this function to return Focus to the calling Form properly
' you must set the MonthCalendar class's hWndForm property BEFORE
' Calling this function!!!!!!!!!!
' *******************************

' This function will always return the date selected by the
' user in the MonthCalendar as the return value for this function.
' If this function is called with the optional Date Range variables
' then it will also return the starting and ending dates of the
' range of dates selected by the user.
' Finally if StartSelectedDate and EndSelectedDate <> 0 Then their
' values will be used initialize the Calendar.
'

Const mcCLASSNAME = "MonthCalendar"
Const mcTITLE = "Calendar"
Dim hwnd As Long
Dim wc As WNDCLASSEX
' Class Atom
Dim lngClassAtom As Long
Dim message As Msg
Dim hInstance As Long
Dim lTemp As Long
Dim ctr As Long
Dim pt As POINTAPI
Dim s As String
Dim blFormIsPopup As Boolean
Dim blAppWindowIsModal As Boolean

On Error Resume Next

' Make sure the instance of MonthCalendar class is valid!
If clsMC Is Nothing Then
    s = " The MonthCalendar class instance you passed to this function is INVALID!" & vbCrLf
    s = s & " YOu must instantiate the MonthCalendar Class object before you call this function" & vbCrLf
    s = s & " The code behind the sample Form shows you how to do this in the Form's Load event" & vbCrLf & vbCrLf
    s = s & "' This must appear here!" & vbCrLf
    s = s & "' Create an instance of our Class" & vbCrLf
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "Set mc = New clsMonthCal" & vbCrLf
    s = s & "' You must set the class hWndForm prop!!!" & vbCrLf
    s = s & "mc.hWndForm = Me.hWnd"
    MsgBox s, vbOKOnly, "Invalid MonthCalendar object!"
    ' Return nothing!
    ShowMonthCalendar = 0
    Exit Function
End If


' If this window already exists then exit!
lngRet = FindWindow(mcCLASSNAME, mcTITLE)
If lngRet <> 0 Then
    's = "The Calendar Window Already Exists!" & vbCrLf
    's = s & "Please Close and then Restart Access!"
    'MsgBox s, vbCritical, "Critical Error. The MonthCalendar Window already exists"
    ' Return nothing!
    'ShowMonthCalendar = 0
    ' We can just Return. The user has tried to open another instance of the Calendar.
    ' Up to this point, Version 98b, we only support one open instance at a time
    ShowMonthCalendar = 0
    Exit Function
End If

' Create a local copy of the MonthCalendar class
Set mc = clsMC

' Update our init cursor props.
lngRet = GetCursorPos(pt)
mc.CursorXinit = pt.x
mc.CursorYinit = pt.y

' Ensure our SelChange vars are reset
SelChangeDateStart = 0
SelChangeDateEnd = 0


' MENU creation time!
hMenu = CreateMenu
hMenuPop = CreatePopupMenu
hMenuPopMisc = CreatePopupMenu
hMenuPopTime = CreatePopupMenu
hMenuPopMiscShowWeekNumbers = CreatePopupMenu
hMenuPopMiscFont = CreatePopupMenu
hMenuPopMiscColor = CreatePopupMenu
hMenuPopMiscToday = CreatePopupMenu
hMenuPopMiscCircleToday = CreatePopupMenu
hMenuPopMiscWindowPosition = CreatePopupMenu
hMenuPopMiscOneClick = CreatePopupMenu

' Viewable Months Menu
lngRet = InsertMenu(hMenuPopMisc, 1&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPop, "&Viewable Months")
' Viewable Months SubMenus
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx1, "1 Month")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx2, "2 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx3, "3 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx4, "4 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx6, "6 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx8, "8 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx9, "9 Months")
lngRet = InsertMenu(hMenuPop, 0&, MF_STRING Or MF_BYPOSITION, Monthx12, "12 Months")
' Erase all check marks

For ctr = 0 To 7
lngRet = CheckMenuItem(hMenuPop, 0, MF_UNCHECKED Or MF_BYPOSITION)
Next ctr
' Now set the Menu Check the current number of months displayed
lTemp = (mc.MonthColumns * mc.MonthRows)
Select Case lTemp
    Case 1
    ctr = 7

    Case 2
    ctr = 6
    
    Case 3
    ctr = 5
    
    Case 4
    ctr = 4
    
    Case 6
    ctr = 3
    
    Case 8
    ctr = 2
    
    Case 9
    ctr = 1
    
    Case 12
    ctr = 0

End Select

' Now set the Menu Check
lngRet = CheckMenuItem(hMenuPop, ctr, MF_CHECKED Or MF_BYPOSITION)


' Misc Properties Menu
lngRet = InsertMenu(hMenu, 4&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMisc, "&Options")

' Let's add Top level Menu Item that does not contain any submen items.
' We will use it like a CommandButton to allow the users to Close the Calendar Window.
'lngRet = InsertMenu(hMenu, 1&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopTime, " &Time ") 'used to select the time
'lngRet = InsertMenu(hMenu, 2&, MF_BYPOSITION Or MF_ENABLED, 999, "   &Fill   ") 'used to call function to fill dates in form
lngRet = InsertMenu(hMenu, 3&, MF_BYPOSITION Or MF_ENABLED, 998, " &Close ")

' Show Time Value SubMenu
lngRet = InsertMenu(hMenuPopTime, 1&, MF_STRING Or MF_ENABLED, 100, "&Start Time (12:00 AM)")
lngRet = InsertMenu(hMenuPopTime, 1&, MF_STRING Or MF_ENABLED, 200, "&End Time   (11:59 PM)")
lngRet = InsertMenu(hMenuPopTime, 1&, MF_STRING Or MF_ENABLED, 300, "&Manual Entry")

' Show WeekNumbers SubMenu
'lngRet = InsertMenu(hMenuPopMisc, 1&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscShowWeekNumbers, "ShowWeek#'s")
'lngRet = InsertMenu(hMenuPopMiscShowWeekNumbers, 0&, MF_STRING Or MF_BYPOSITION, ShowWeekNumYES, "YES")
'lngRet = InsertMenu(hMenuPopMiscShowWeekNumbers, 0&, MF_STRING Or MF_BYPOSITION, ShowWeekNumNO, "NO")
'If mc.ShowWeekNumbers = False Then
mc.ShowWeekNumbers = False
'    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_CHECKED Or MF_BYPOSITION)
'    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_UNCHECKED Or MF_BYPOSITION)
'Else
'    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_CHECKED Or MF_BYPOSITION)
'    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_UNCHECKED Or MF_BYPOSITION)
'End If

' Show Today's Date
'lngRet = InsertMenu(hMenuPopMisc, 4&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscToday, "Show Today")
'lngRet = InsertMenu(hMenuPopMiscToday, 0&, MF_STRING Or MF_BYPOSITION, ShowTodayYES, "YES")
'lngRet = InsertMenu(hMenuPopMiscToday, 0&, MF_STRING Or MF_BYPOSITION, ShowTodayNO, "NO")
'
'If mc.NoToday = True Then
mc.NoToday = False
'    lngRet = CheckMenuItem(hMenuPopMiscToday, 0, MF_CHECKED Or MF_BYPOSITION)
'    lngRet = CheckMenuItem(hMenuPopMiscToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
'Else
'    lngRet = CheckMenuItem(hMenuPopMiscToday, 1, MF_CHECKED Or MF_BYPOSITION)
'    lngRet = CheckMenuItem(hMenuPopMiscToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
'End If
 
  
' Circle Today's Date
'lngRet = InsertMenu(hMenuPopMisc, 5&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscCircleToday, "Circle Today")
'lngRet = InsertMenu(hMenuPopMiscCircleToday, 0&, MF_STRING Or MF_BYPOSITION, ShowCircleTodayYES, "YES")
'lngRet = InsertMenu(hMenuPopMiscCircleToday, 0&, MF_STRING Or MF_BYPOSITION, ShowCircleTodayNO, "NO")
'If mc.NoTodayCircle = True Then
mc.NoTodayCircle = False
'lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_CHECKED Or MF_BYPOSITION)
'lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
'Else
'lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_CHECKED Or MF_BYPOSITION)
'lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
'End If


' Window Position
lngRet = InsertMenu(hMenuPopMisc, 6&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscWindowPosition, "&Calendar Location")
lngRet = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx0, "Cursor Location when Calendar Opened")
lngRet = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx1, "Where User Last Dragged")
lngRet = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx2, "Center of Access App Window")
lngRet = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx3, "Center of Screen")
lngRet = InsertMenu(hMenuPopMiscWindowPosition, 0&, MF_STRING Or MF_BYPOSITION, Positionx4, "Top Left Corner")

For ctr = 0 To 4
lngRet = CheckMenuItem(hMenuPopMiscWindowPosition, ctr, MF_UNCHECKED Or MF_BYPOSITION)
Next ctr
' Now set the Menu Check the current number of months displayed
lTemp = (mc.WindowLocation)
' Now set the Menu Check
lngRet = CheckMenuItem(hMenuPopMiscWindowPosition, 4 - lTemp, MF_CHECKED Or MF_BYPOSITION)


' Single or Double Click to select Date
'lngRet = InsertMenu(hMenuPopMisc, 7&, MF_POPUP Or MF_BYPOSITION Or MF_ENABLED, hMenuPopMiscOneClick, "Single Or Double Click")
'lngRet = InsertMenu(hMenuPopMiscOneClick, 0&, MF_STRING Or MF_BYPOSITION, DoubleClick, "Double Click to Select Date")
'lngRet = InsertMenu(hMenuPopMiscOneClick, 0&, MF_STRING Or MF_BYPOSITION, SingleClick, "Single Click to Select Date")

'If mc.OneClick = True Then
'lngRet = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_CHECKED Or MF_BYPOSITION)
'lngRet = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_UNCHECKED Or MF_BYPOSITION)
'Else
'lngRet = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_CHECKED Or MF_BYPOSITION)
'lngRet = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_UNCHECKED Or MF_BYPOSITION)
'End If

' Get instance of this App
hInstance = apiGetWindowLong(Application.hWndAccessApp, GWL_HINSTANCE)
 
' From code by Ray Mercer
    ' Set up and register window class
    wc.cbSize = Len(wc)
    wc.style = CS_HREDRAW Or CS_VREDRAW
    
    wc.lpfnWndProc = GetFuncPtr(AddressOf WindowProc) 'ALD
    
    wc.cbClsExtra = 0&
    wc.cbWndExtra = 0&
    wc.hInstance = hInstance
    wc.hIcon = LoadIcon(hInstance, IDI_APPLICATION)
    wc.hCursor = LoadCursor(hInstance, IDC_ARROW)
    wc.hbrBackground = GetStockObject(WHITE_BRUSH)
    wc.lpszMenuName = 0&
    wc.lpszClassName = CLASSNAME
    wc.hIconSm = LoadIcon(hInstance, IDI_APPLICATION)
    
    ' Register this Class
   lngClassAtom = RegisterClassEx(wc)
 
 
 Dim lngEXStyle As Long
 ' Force window to always stay on top
 lngEXStyle = WS_EX_DLGMODALFRAME ' April 6 trying to fix WIn98 Form in Popup view Or WS_EX_TOPMOST
 
    ' Create a window Set to be NOT VISIBLE TO START Or WS_VISIBLE
 hwnd = CreateWindowEx(lngEXStyle, _
                        CLASSNAME, _
                        TITLE, _
                        WS_POPUPWINDOW Or WS_CAPTION, _
                        CW_USEDEFAULT, _
                        CW_USEDEFAULT, _
                        CW_USEDEFAULT, _
                        CW_USEDEFAULT, _
                           mc.hWndForm, _
                        hMenu, _
                        hInstance, _
                         0&)
                        
    ' We have to allow for the following:
    ' 1) The calling Form's Modal prop is turned on
    lngRet = GetWindowLong(Application.hWndAccessApp, GWL_STYLE)
    blAppWindowIsModal = lngRet And WS_DISABLED
    
    ' 2) The calling Form's Popup prop is turned on
    lngRet = GetWindowLong(mc.hWndForm, GWL_STYLE)
    blFormIsPopup = lngRet And WS_POPUP
     
     
    ' We will actually create our MonthCal window by setting the
    ' Class hWnd property.
    ' Set the Control's Parent Window property
    mc.hwnd = hwnd
       
    ' Init the Calendar to the date(s) supplied by the
    ' user in the calling function
    If StartSelectedDate <> 0 And EndSelectedDate <> 0 Then
        mc.SetSelectedDateRange StartSelectedDate, EndSelectedDate
        ' Update our local copies of these vars
        ' Need to redo the logic to get rid of these local vars
        ' See the date select code in the WindProc
        localStartSelectedDate = StartSelectedDate
        localEndSelectedDate = EndSelectedDate
        ' Clear our Return Date local Var.
        SelectedDate = 0
    Else
        If StartSelectedDate <> 0 Then
            mc.SelectedDate = StartSelectedDate
            ' Clear our Return Date local Var.
            SelectedDate = 0
        Else
            SelectedDate = 0
            localStartSelectedDate = 0
            localEndSelectedDate = 0
        End If
    End If
         
         
    ' The following logic is required to ensure our MonthCalendar window
    ' is MODAL(the user can only click in this window)
    ' If parent form's Popup prop is turned on then
    ' we have to Disable this Form ourselves
    If blFormIsPopup Then lngRet = EnableWindow(mc.hWndForm, 0)
       
    ' We only want to Disable the main app window if
    ' the Form's Modal prop is not true.
    ' Check and see if the main Access app window
    ' is disabled already - if not then disable it
    If Not blAppWindowIsModal Then
        lngRet = EnableWindow(Application.hWndAccessApp, 0)
    End If
         
    ' Show the Calendar's Parent window first then the MonthCal window
    ShowWindow hwnd, SW_SHOWNORMAL
    ShowWindow mc.hWndCal, SW_SHOWNORMAL
         
    ' Enter message loop
    '(all window messages are handled in WindowProc())
    Do While 0 <> GetMessage(message, 0&, 0&, 0&)
        TranslateMessage message
        DispatchMessage message
    Loop
    
    ' User has closed the MonthCalendar window
    ' Return the Selected Date
    ' If the user has called this function with the optional
    ' date range vars then fill them in.
    If SelectedDate <> 0 Then
        ' The Calendar Window is closed so we cannot
        ' use our Class methods that use SendMessage
        ' to get their current values.
        'if
        StartSelectedDate = SelectedDate
        EndSelectedDate = localEndSelectedDate
        ShowMonthCalendar = True
    Else
        ' User did not SELECT a Date
        StartSelectedDate = 0
        EndSelectedDate = 0
        ShowMonthCalendar = False
    End If
    
        
    ' Unregister our Custom Window Class
    ' If you don't then you will GPF on the next init of the class
    lngRet = UnregisterClass(CLASSNAME, hInstance)
        
    ' If Form was Popup then Enable this window first
    If blFormIsPopup Then
        lngRet = EnableWindow(mc.hWndForm, 1)
    End If
    
    ' In order to prevent screen flashing upon closing
    ' our MonthCalendar window we have to enable the
    ' main Access application window in the MonthCalendar's
    ' WinProc's WM_CLOSE message handler. From here now though,
    ' we  have to Disable the main Access application window
    ' if the calling form's Modal prop was turned on.
    If blAppWindowIsModal Then
        'Disable Access App window
        lngRet = EnableWindow(Application.hWndAccessApp, 0)
    End If
     
     ' Release Class reference required to be visible to our WindProc
    Set mc = Nothing
        
    ' Ensure focus returns to calling form.
    SetFocus clsMC.hWndForm
    
End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
'Main message handler for the MonthCalendar window
' *** WARNING ***
' DO NOT PLACE DEBUG BREAKPOINTS IN THIS FUNCTION
' *** WARNING ***

Dim ps          As PAINTSTRUCT
Dim rc          As RECT
Dim hdc         As Long
Dim strTemp     As String
Dim strTemp1    As String
Dim lngRet      As Long
Dim blRet       As Boolean
Dim intTemp     As Integer
Dim arrayTime(0 To 1) As SYSTEMTIME
' Mouse selection of Date(s)
Dim DateStart As Date
Dim DateEnd As Date


Dim nmsc As NMSELCHANGE
Dim hdr As NMHDR
Dim nmds As NMDAYSTATE
' THere is a bug or I am having alignment problems
' so we pass the second element of this array
' and leave the first zero'd out.
Dim aMDS(-1 To 13) As MONTHDAYSTATE

' To hold local copy of the current Message
Dim CurMessage As Msg
Dim lngCurMessagetime  As Long
Dim lngWMmessage As Long
Dim lngDoubleClickTime As Long

Static lngLastMouseDown As Long
' Flag to make sure we have a MouseUp bewtween our
' MouseDown messages to signify a Double Click
' not just the Mouse Button held down
Static blMouseUp As Boolean

' Temp Window Handle for Dialogs
Dim hWndTemp As Long

' You cannot have unhandled errors in a WinProc so
' we will jsut ingnore them all!!<grin>
' Really though, this is very heavily debugged code!
On Error Resume Next


Select Case message
          
    Case WM_MOVE
    ' Update the MonthCalendar's current
    Call UpdateCursor(lparam, hwnd)
    
    Case WM_PAINT
        ' Must leave this in to ensure Window is Redrawn!!!
        hdc = BeginPaint(hwnd, ps)
        Call EndPaint(hwnd, ps)
        Exit Function

    Case WM_KEYDOWN, WM_KEYUP
                    
            ' Select case on the Virtual Key Code
            Select Case wParam
        
            Case VK_ESCAPE
            Call PostMessage(hwnd, WM_CLOSE, 0, 0)
            Exit Function
        
        
            Case VK_SHIFT, VK_LEFT, VK_RIGHT, VK_DOWN, VK_UP, VK_HOME, VK_END, vbKeyPageDown, vbKeyPageUp
            KeysToMonthCal hwnd, message, wParam, lparam
            Exit Function
        
            Case VK_RETURN
            ' If the SelChangeDateStart var != 0 then send our MCN_SELECT Message
            If SelChangeDateStart = 0 Then Exit Function
            
            If SelChangeDateEnd = SelChangeDateStart Then
            mc.SelectedDate = SelChangeDateStart
            Else
            mc.SetSelectedDateRange SelChangeDateStart, SelChangeDateEnd
            End If
            
            ' Update our local var
            SetSelectedDate SelChangeDateStart
            ' Update our Class starting and ending date range vars
            UpdateRangeVars SelChangeDateStart, SelChangeDateEnd
            
            
           ' Let's CLose the Calendar
             Call PostMessage(hwnd, WM_CLOSE, 0, 0)
            'Debug
            'Debug.Print "Used Enter key to select date!"
            Exit Function
        
            Case Else
            WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
           
            Exit Function
            End Select
        
     Case WM_CLOSE
    ' April 12, 2004
    ' FINALLY resolved issue of screen flickering with Win2K or higher!!
    ' We have to temporarily Enable the main Access application window
    lngRet = EnableWindow(Application.hWndAccessApp, 1)
     lngRet = ShowWindow(Application.hWndAccessApp, SW_SHOW)
     
     WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
    Exit Function
        
    Case WM_DESTROY
        'Exit Function
        ' Enable Main Access Window now
        ' that the MonthCalendar is closed!
        'lngRet = ShowWindow(Application.hWndAccessApp, SW_SHOW)
        PostQuitMessage 0&
        Exit Function
    
 
    Case WM_PARENTNOTIFY
        ' Grab the lower WORD
        lngWMmessage = (wParam And &HFFFF)
        ' Switch on Window Message
        Select Case lngWMmessage

            Case WM_LBUTTONDOWN

           ' Mod Nov 24 -2002
           ' Removed MouseButton logic to determine when to close
           ' calendar. Now we simply check it from the SELECT notification
           ' and close the window if CHeckOneClick property is TRUE.
           ' We do not use the DoubleCLick logic either.
           ' Get the current Double Click interval
            lngCurMessagetime = GetMessageTime
            lngDoubleClickTime = GetDoubleClickTime

            ' Make sure the Cursor is double clicking
            ' on an actual Date not on a Calendar control
            blRet = LocationCursorOnCalendar(lparam)
            If Not blRet Then
                ' Call the default WIndow proc
                WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
                Exit Function
            End If

            ' Debug. A2K closing date range on one click!
            If Abs((lngCurMessagetime - lngLastMouseDown)) < lngDoubleClickTime Then ' Or CheckOneClick = True Then
                ' Double CLicked-or CheckOneClick-Let's CLose the Calendar
                Call PostMessage(hwnd, WM_CLOSE, 0, 0)
                lngLastMouseDown = 0
                blMouseUp = False
                Exit Function
            End If

            ' Always update our last left mouse button pressed var
            lngLastMouseDown = lngCurMessagetime


            Case Else
            ' Call the default Window proc
            WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
            Exit Function

            ' All Done!
            End Select
 
 
    Case WM_NOTIFY
        ' Update our class startdate, and range date props.
        ' Copy the NMRH structure to our local copy
        CopyMem hdr, ByVal lparam, Len(hdr)
        
        Select Case hdr.Code
            
            ' Modified Nov 24 -2002
            ' SELECT is when the user explicitly clicks to select a date.
            ' SELCHANGE is when the user scrolls through the calendar automatically
            ' updating the selected date.
            ' Thanks to Blake Sell for catching this!
            
            Case MCN_SELECT
            ' *** this needs to be fixed up to have seperate routines
            ' for single vs range date selections.
            ' Drop local vars and use the MonthCalendar Class only
            ' Grab the struct info
            CopyMem nmsc, ByVal lparam, Len(nmsc)
        
            ' Convert to our Date format
            With nmsc.stSelStart '(0)
                DateStart = DateSerial(.wYear, .wMonth, .wDay)
            End With
            With nmsc.stSelEnd '(1)
                DateEnd = DateSerial(.wYear, .wMonth, .wDay)
            End With
        
            ' Update our local var
            SetSelectedDate DateStart
            ' Update our Class starting and ending date range vars
            UpdateRangeVars DateStart, DateEnd
            
            
           ' Mod Nov 24 -2002
           ' Removed MouseButton logic to determine when to close
           ' calendar. Now we simply check it from the SELECT notification
           ' and close the window if CHeckOneClick property is TRUE.
            If mc.OneClick = True Then
                ' Double CLicked-or CheckOneClick-Let's CLose the Calendar
                Call PostMessage(hwnd, WM_CLOSE, 0, 0)
                lngLastMouseDown = 0
                blMouseUp = False
                'Exit Function
            End If
            
            Exit Function
   
   
   
   ' June 2 - 2004 - adding support for ENTER key to select currently highlighted date.
   Case MCN_SELCHANGE
   
    ' Grab the struct info
            CopyMem nmsc, ByVal lparam, Len(nmsc)
        
            ' Convert to our Date format
            With nmsc.stSelStart '(0)
                SelChangeDateStart = DateSerial(.wYear, .wMonth, .wDay)
            End With
            With nmsc.stSelEnd '(1)
                SelChangeDateEnd = DateSerial(.wYear, .wMonth, .wDay)
            End With
           ' Debug.Print "DateStart:" & DateStart

   
            Case MCN_GETDAYSTATE
            Dim s As SYSTEMTIME
            Dim lngTemp As Long
            Dim ptrArray As Long
            
            Dim x As Integer
            Dim intStartMonth As Integer
            Dim intCurrentMonth As Integer
            Dim intCurrentYear As Integer
            
            For x = -1 To UBound(aMDS)
            aMDS(x).lpMONTHDAYSTATE = 0
            Next
            
             CopyMem nmds, ByVal lparam, Len(nmds)
            intTemp = nmds.cDayState
            'Debug.Print "Months requested:" & intTemp
            'Debug.Print time
            ' Have to allow for the fact that the month before and
            ' the month after are always requested. THis means the starting year
            ' can be one year before the year of the first fully displayed month.
            intStartMonth = nmds.stStart.wMonth
            intCurrentYear = nmds.stStart.wYear
            
            intCurrentMonth = intStartMonth '+ x
            For x = 0 To intTemp - 1
            
            If intCurrentMonth > 12 Then
            intCurrentMonth = intCurrentMonth - 12 '1
            intCurrentYear = intCurrentYear + 1
            End If
            aMDS(x).lpMONTHDAYSTATE = mc.GetDAYSTATE(intCurrentYear, intCurrentMonth)
            intCurrentMonth = intCurrentMonth + 1
            Next x
            ' set the address of our array
            lngTemp = VarPtr(aMDS(0))
            CopyMem ByVal lparam + (Len(nmds) - 4), lngTemp, 4
        
            ' Signal we want this message to be processed
            WindowProc = 0
            Exit Function
            
   
            Case Else
            WindowProc = DefWindowProc(hwnd, message, wParam, lparam)


        End Select

    Case WM_COMMAND:
               ' WM_COMMAND is sent to the window
               ' whenever someone clicks a menu.
               ' The menu's item ID is stored in wParam.
               'Debug.Print wparam
            Select Case wParam
                    Case Monthx1 To Monthx12
                    'Call MsgBox("You clicked Dynamic Sub Menu 1!", vbExclamation)
                    SetMonths (CInt(wParam) - 900)
                    Exit Function
                        
                    Case ShowWeekNumYES
                    ShowWeekNums True
                    Exit Function
                    
                    Case ShowWeekNumNO
                    ShowWeekNums False
                    Exit Function

                    ' Show Todays Date at bottom of Calendar
                    Case ShowTodayYES
                    sShowToday False
                    Exit Function
                    
                    Case ShowTodayNO
                    sShowToday True
                    Exit Function
                    
                    ' Circle Today's Date
                    Case ShowCircleTodayYES
                    sShowcircleToday False
                    Exit Function
                    
                    Case ShowCircleTodayNO
                    sShowcircleToday True
                    Exit Function
                    
                    ' WindowPosition menu
                    Case Positionx0 To Positionx8
                    sWindowPosition wParam, hwnd
                    
                    Case SingleClick
                    sClick True
                    Exit Function
                    
                    Case DoubleClick
                        sClick False
                        Exit Function
                    
                    Case 999 'Fill All Dates Function
                        If SelChangeDateStart = 0 Then
                            'Call FillDate   'Function ALD 04/13/2005
                            Call PostMessage(hwnd, WM_CLOSE, 0, 0)
                            Exit Function
                        End If
                        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
                        'Call FillDate   'Function ALD 04/13/2005
                         lngLastMouseDown = 0
                        blMouseUp = False
                        'Debug.Print SelChangeDateStart
                        Exit Function
                        
                    Case 998 'Close Form
                        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
                        lngLastMouseDown = 0
                        blMouseUp = False
                        Exit Function
                    Case 100
                        gdteTimeValue = "00:00:00" ' MsgBox "set begin time global variable"
                        Exit Function
                        'set begin time global variable
                    Case 200
                        gdteTimeValue = "23:59:59" ' MsgBox "set end time global variable"
                        Exit Function
                    Case 300
                        Dim varTempTime As Variant
                            varTempTime = InputBox("Enter a time to be used for the date selected." & vbCrLf & "Any format is acceptable.", "Time Entry.", gdteTimeValue)
                        If StrPtr(varTempTime) = 0 Then
                            Exit Function 'user pressed cancel
                        End If
                        If IsDate(varTempTime) Then
                            gdteTimeValue = timeValue(varTempTime)
                        Else
                            MsgBox "The entry was not a valid time.", vbInformation, "No action taken."
                        End If
                        'Debug.Print gdteTimeValue
                    Case Else
                    
                   ' Call the Default Window Procedure for all other WM_COMMAND'
                   WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
                   Exit Function
                   End Select

Case Else
    'pass all other messages to default window procedure
WindowProc = DefWindowProc(hwnd, message, wParam, lparam)
        
End Select
 
End Function
 
Function GetFuncPtr(ByVal lngFnPtr As Long) As Long
    'wrapper function to allow AddressOf to be used within VB
    GetFuncPtr = lngFnPtr
End Function

Private Function SetSelectedDate(ByVal dt As Date)
SelectedDate = dt
End Function

Private Function SetMonths(ByVal mth As Integer)
mc.SetViewableMonths mth
'Exit Function

Dim ctr As Long
Dim lTemp As Long
Dim lRet As Long

' 7 Possible/Total Menu Items to uncheck
For ctr = 0 To 7
lRet = CheckMenuItem(hMenuPop, ctr, MF_UNCHECKED Or MF_BYPOSITION)
Next ctr
' Now set the Menu Check the current number of months displayed
lTemp = (mc.MonthColumns * mc.MonthRows)
Select Case lTemp
    Case 1
    ctr = 7

    Case 2
    ctr = 6
    
    Case 3
    ctr = 5
    
    Case 4
    ctr = 4
    
    Case 6
    ctr = 3
    
    Case 8
    ctr = 2
    
    Case 9
    ctr = 1
    
    Case 12
    ctr = 0

End Select

' Now set the Menu Check
lRet = CheckMenuItem(hMenuPop, ctr, MF_CHECKED Or MF_BYPOSITION)

End Function

Private Sub sClick(bl As Boolean)
' Sets the Class's OneClick property and the
' appropriate Menu Check Marks
If bl Then
    mc.OneClick = True
    lngRet = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_UNCHECKED Or MF_BYPOSITION)
Else
    mc.OneClick = False
    lngRet = CheckMenuItem(hMenuPopMiscOneClick, 1, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscOneClick, 0, MF_UNCHECKED Or MF_BYPOSITION)
End If
End Sub

Private Function ShowWeekNums(ByVal yn As Boolean)
If yn = True Then
    mc.ShowWeekNumbers = True
    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_UNCHECKED Or MF_BYPOSITION)
Else
    mc.ShowWeekNumbers = False
    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 0, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscShowWeekNumbers, 1, MF_UNCHECKED Or MF_BYPOSITION)
End If
End Function

Private Sub sShowToday(bl As Boolean)
If bl Then
mc.NoToday = True
lngRet = CheckMenuItem(hMenuPopMiscToday, 0, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
Else
mc.NoToday = False
lngRet = CheckMenuItem(hMenuPopMiscToday, 1, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
End If
End Sub

Private Sub sShowcircleToday(bl As Boolean)
If bl = True Then
    mc.NoTodayCircle = True
    lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_UNCHECKED Or MF_BYPOSITION)
Else
    lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 1, MF_CHECKED Or MF_BYPOSITION)
    lngRet = CheckMenuItem(hMenuPopMiscCircleToday, 0, MF_UNCHECKED Or MF_BYPOSITION)
    mc.NoTodayCircle = False
End If
End Sub

Private Sub sWindowPosition(wParam As Long, hwnd As Long)
' Position Window according to users Menu selections
    ' a) 0 -Pop at cursor location when user activates Calendar
    ' b) 1 -Where they manually move/leave it at
    ' c) 2 -Centered in Access App Window
    ' d) 3 -Centered on entire screen
    ' d) 4 -Top Left Corner

Dim rc1 As RECT
Dim pt As POINTAPI
Dim lngRet As Long
Dim ctr As Long
Dim lTemp As Long


Select Case wParam

Case Positionx0
' Pop at Cursor
mc.PositionAtCursor = True


Case Positionx1
    mc.PositionAtCursor = False
    ' Use current position of Calendar Window
    ' Get rectangle for our Form
    'Debug.Print "GetWindowRect- Me.hWnd:" & m_Form.hWnd
    lngRet = GetWindowRect(hwnd, rc1)

    mc.CursorX = rc1.Left 'pt.x 'rc1.Left
    mc.CursorY = rc1.Top 'pt.y

Case Positionx2 To Positionx8
    mc.PositionAtCursor = False

Case Else

End Select

' Update Window Position property
mc.WindowLocation = wParam - 920
'Debug.Print "modCalendar - mc.Windowlocation:" & wparam ' mc.WindowLocation
For ctr = 0 To 4
lngRet = CheckMenuItem(hMenuPopMiscWindowPosition, ctr, MF_UNCHECKED Or MF_BYPOSITION)
Next ctr
' Now set the Menu Check the current number of months displayed
lTemp = (mc.WindowLocation)
' Now set the Menu Check
lngRet = CheckMenuItem(hMenuPopMiscWindowPosition, 4 - lTemp, MF_CHECKED Or MF_BYPOSITION)

mc.ReDraw
End Sub

Private Sub KeysToMonthCal(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long)
Call PostMessage(ByVal mc.hWndCal, ByVal Msg, ByVal wParam, ByVal lparam)
End Sub

Private Sub UpdateRangeVars(ByVal DateStart, ByVal DateEnd)
localStartSelectedDate = DateStart
localEndSelectedDate = DateEnd
End Sub

Private Sub UpdateCursor(ByVal lparam, ByVal hwnd As Long)
'xPos = (int)(short) LOWORD(lParam);   // horizontal position
'yPos = (int)(short) HIWORD(lParam);   // vertical position

Dim pt As POINTAPI
Dim rc As RECT
Dim lRet As Long

' Should not happen
If mc.hwnd = 0 Then Exit Sub
' Only update if the window is visible
lRet = GetWindowLong(hwnd, GWL_STYLE)
If Not (lRet And WS_VISIBLE) Then Exit Sub

' If PositionAtCursor is True then
' DO NOT UPDATE
If mc.PositionAtCursor Then Exit Sub

lngRet = GetWindowRect(hwnd, rc)
pt.x = rc.Left
pt.y = rc.Top

'Debug.Print time & "  UpdateCursor -X:" & rc.Left & "  Y:" & rc.Top
mc.CursorX = pt.x
mc.CursorY = pt.y

'UpdateCursor -X:" & mc.CursorX & "  Y:" & mc.CursorY

End Sub

Private Function LocationCursorOnCalendar(ByVal lparam As Long) As Boolean
Dim ht As MCHITTESTINFO
' The x-coordinate of the cursor is the low-order word,
' and the y-coordinate of the cursor is the high-order word.

ht.pt.x = LoWord(lparam)
ht.pt.y = HiWord(lparam)

' Set structure size
ht.cbSize = Len(ht)
lngRet = apiSendMessage(ByVal mc.hWndCal, ByVal MCM_HITTEST, ByVal 0&, ht)
If ht.uHit <> MCHT_CALENDARDATE Then
LocationCursorOnCalendar = False
Else
LocationCursorOnCalendar = True
End If
End Function

Private Function ReleaseClass()
Set mc = Nothing
End Function

Private Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then ' &H8000& = &H00008000
       LoWord = DWord Or &HFFFF0000
    Else
       LoWord = DWord And &HFFFF&
    End If
End Function

Private Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

