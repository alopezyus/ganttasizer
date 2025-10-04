Attribute VB_Name = "api_scroll_wheel"
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const cSCROLLCHANGE As Long = 10
 
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
 
Private mFormHwnd As Long
Private mbHook As Boolean

Dim mForm As Object

'********************************************************************
'** 32/64 bit api's
'********************************************************************
#If Win64 Then

    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long

    Private mLngMouseHook As LongPtr

    'mouse message incoming to form
    Private Function MouseProc(ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
        On Error GoTo errH
        If (ncode = HC_ACTION) Then
            If GetActiveWindow = mFormHwnd Then
                'see if its a mouse wheel action
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
                    If lParam.hwnd > 0 Then
                        'scroll up, to go previous month
                        datepickerform.prevMonthButton_Click
                    Else
                        'scroll down, go to next month
                        datepickerform.nextMonthButton_Click
                    End If
                    Exit Function
                End If
            End If
        End If
        
        'pass message along
        MouseProc = CallNextHookEx(mLngMouseHook, ncode, wParam, ByVal lParam)
        Exit Function
errH:
        'if there is an error then stop everything
        UnhookFormScroll
    End Function

    'listen to all the mouse events on the form passed in
    Sub DoHookFormScroll(oForm As Object)
        'too unstable on 64 bit
    End Sub
    Sub UnhookFormScroll()
        'too unstable on 64 bit
    End Sub

#Else

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    
    Private mLngMouseHook              As Long
    
    'mouse message incoming to form
    Private Function MouseProc(ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
        On Error GoTo errH
        If (ncode = HC_ACTION) Then
            If GetActiveWindow = mFormHwnd Then
                'see if its a mouse wheel action
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
                    If lParam.hwnd > 0 Then
                        'scroll up, to go previous month
                        datepickerform.prevMonthButton_Click
                    Else
                        'scroll down, go to next month
                        datepickerform.nextMonthButton_Click
                    End If
                    Exit Function
                End If
            End If
        End If
        
        'pass message along
        MouseProc = CallNextHookEx(mLngMouseHook, ncode, wParam, ByVal lParam)
        Exit Function
errH:
        'if there is an error then stop everything
        UnhookFormScroll
    End Function

    'listen to all the mouse events on the form passed in
    Sub DoHookFormScroll(oForm As Object)
        Dim lngAppInst                  As Long
        Dim hwndUnderCursor             As Long
         
        Set mForm = oForm
        hwndUnderCursor = FindWindow("ThunderDFrame", oForm.Caption)
        If mFormHwnd <> hwndUnderCursor Then
            UnhookFormScroll
            mFormHwnd = hwndUnderCursor
            lngAppInst = GetWindowLong(mFormHwnd, GWL_HINSTANCE)
            If Not mbHook Then
                mLngMouseHook = SetWindowsHookEx( _
                WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                mbHook = mLngMouseHook <> 0
            End If
        End If
    End Sub

    'stops listening for all the mouse messages to the form
    Sub UnhookFormScroll()
        If mbHook Then
            UnhookWindowsHookEx mLngMouseHook
            mLngMouseHook = 0
            mFormHwnd = 0
            mbHook = False
        End If
    End Sub

#End If

