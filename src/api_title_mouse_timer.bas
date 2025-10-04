Attribute VB_Name = "api_title_mouse_timer"

'********************************************************************
'** 32/64 bit api's
'********************************************************************
#If Win64 Then

    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As theCursor) As Long
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    
    Dim TimerID As LongPtr
    Dim hdc As LongPtr
    Dim mhWndForm As LongPtr
    
#Else
    
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (p As theCursor) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    
    Dim TimerID As Long
    Dim hdc As Long
    Dim mhWndForm As Long
    
#End If

'used for conversion of mouse to form positoning
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

'mouse position
Public Type theCursor
    Left As Long
    Top As Long
End Type

'timer options
Dim TimerSeconds As Single, tim As Boolean, Counter As Long

'used to get the screen height info
Private Const SM_CXMAXIMIZED = 61
Private Const SM_CYMAXIMIZED = 62

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public g_screenHeight As Long
Public g_screenWidth As Long

'converts X position for mouse
Private Function pointsPerPixelX() As Double
    hdc = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
End Function

'converts Y position for mouse
Private Function pointsPerPixelY() As Double
    hdc = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
End Function

Public Sub getScreenHeight()
    'get the screen height
    g_screenHeight = GetSystemMetrics32(SM_CYMAXIMIZED)
    g_screenHeight = g_screenHeight * pointsPerPixelY
    
    'get the screen width
    g_screenWidth = GetSystemMetrics32(SM_CXMAXIMIZED)
    g_screenWidth = g_screenWidth * pointsPerPixelX
    
    'bump for some comfort space
    g_screenHeight = g_screenHeight - 20
    g_screenWidth = g_screenWidth - 20
    
    'check to be sure we have a generally positive number
    If g_screenHeight < 300 Then g_screenHeight = 300
    If g_screenWidth < 300 Then g_screenWidth = 300
End Sub

'positions the passed in form's top/left to the current mouse position
Sub MoveFormToMouse(theForm)

    Dim mousePos As theCursor
    
    'get the position of the mouse
    GetCursorPos mousePos
    
    'set the form postion to the mouse position
    theForm.Top = pointsPerPixelX * mousePos.Top
    theForm.Left = pointsPerPixelY * mousePos.Left
    
'    'make sure we don't run off the bottom
'    If g_screenHeight > 0 Then
'        If (theForm.Top + theForm.Height) > g_screenHeight Then theForm.Top = g_screenHeight - theForm.Height
'    End If
'
'    'make sure we don't run off the right
'    If g_screenWidth > 0 Then
'        If (theForm.Left + theForm.Width) > g_screenWidth Then theForm.Left = g_screenWidth - theForm.Width
'    End If
    
End Sub


'********************************************************
'TIMER
'********************************************************

'starts the timer
Sub StartTimer()
    'set the timer for 1 second
    TimerSeconds = 1
    TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
End Sub

'stops the timer
Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
End Sub

'timer called on the second, update the calendar clock
Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
    If IsUserFormLoaded("datepickerform") Then
        datepickerform.timeButton.Caption = VBA.Time
    Else
        EndTimer
    End If
End Sub

'checks to see if a particular form is loaded
Private Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    For Each UForm In VBA.UserForms
        IsUserFormLoaded = UForm.Name = UFName
        If IsUserFormLoaded Then
            Exit For
        End If
    Next
End Function

'********************************************************
'REMOVE TITLE BAR
'********************************************************

'hides the caption for a userform
Sub removeCaption(objForm As Object)
    Dim lStyle As Long
    Dim hMenu As Long
     
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", objForm.Caption) 'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    End If
    
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub
