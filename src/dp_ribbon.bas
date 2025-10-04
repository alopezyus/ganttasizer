Attribute VB_Name = "dp_ribbon"
'adds the date picker to the right click menu

'ribbon init
Sub DatePickerRibbonInit(Ribbon As IRibbonUI)
    'get a copy of the ribbon object
    Set theRibbon = Ribbon
End Sub

Sub AddDatePickerRightClick()
'    Dim ContextMenu As CommandBar
'
'    'add to context menus
'    If DatePickerOnMenu(Application.CommandBars("Cell")) = False Then AddDatePickerToMenu Application.CommandBars("Cell")
'    If DatePickerOnMenu(Application.CommandBars("List Range Popup")) = False Then AddDatePickerToMenu Application.CommandBars("List Range Popup")
End Sub

Sub AddDatePickerToMenu(cb As CommandBar)
    ' Add one custom button to the Cell context menu.
    With cb.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "DatePicker_Click"
        .FaceId = 1992
        .Caption = "Date Picker"
        .Tag = "samrad_dp"
        .BeginGroup = True
    End With
End Sub

'removes date picker menu item from context menu, cannot be called from the menu item
Sub RemoveDatePickerRightClick()
    'get the cell context menu
    RemoveDatePickerFromMenu Application.CommandBars("Cell")
    RemoveDatePickerFromMenu Application.CommandBars("List Range Popup")
End Sub

'removes the date picker item from the menu
Sub RemoveDatePickerFromMenu(cb As CommandBar)
    Dim ctrl As CommandBarControl
    
    'loop through and find the date picker
    For Each ctrl In cb.Controls
        If ctrl.Tag = "samrad_dp" Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl
End Sub

'checks to see if the date picker menu item already exists
Private Function DatePickerOnMenu(cb As CommandBar) As Boolean
    Dim ctrl As CommandBarControl
    DatePickerOnMenu = False
    
    'loop through and find the date picker
    For Each ctrl In cb.Controls
        If ctrl.Tag = "samrad_dp" Then
            DatePickerOnMenu = True
            Exit For
        End If
    Next ctrl
End Function

'updates the right click menu based on the options
Sub updateRightClickMenu()
    'turn on or off for right click
    If fShowDPRightClick Then
        'add date picker to context menus
        AddDatePickerRightClick
    Else
        'remove date picker from context menus
        RemoveDatePickerRightClick
    End If
    
End Sub

'Callback for btnInsertTodaysDate onAction
Sub InsertTodaysDate_Click(control As IRibbonControl)
    dayPicked 68
End Sub

'Callback for btnInsertTodaysDateTime onAction
Sub InsertTodaysDateTime_Click(control As IRibbonControl)
    dayPicked 68, True
End Sub

'Callback for ShowDPMenu onAction
Sub ShowDPMenu_Click(control As IRibbonControl, pressed As Boolean)
    
    'update the flag
    fShowDPRightClick = pressed
    
    'save the setting
    VBA.SaveSetting "samradapps_datepicker", "ribbon", "fShowDPRightClick", fShowDPRightClick
    
    'refresh the option
    updateRightClickMenu
    
End Sub

'Callback for ShowDPGrid onAction
Sub ShowDPGrid_Click(control As IRibbonControl, pressed As Boolean)
    
    'update the flag
    fShowDPInGrid = pressed
    
    'save the setting
    VBA.SaveSetting "samradapps_datepicker", "ribbon", "fShowDPInGrid", fShowDPInGrid
    
    'refresh the option
    updateRightClickMenu
    
End Sub

'Callback for btnInsertTodaysDate getLabel
Sub InsertTodaysDate_Label(control As IRibbonControl, ByRef returnedVal)
    returnedVal = VBA.Date
End Sub

'Callback for btnInsertTodaysDateTime getLabel
Sub InsertTodaysDateTime_Label(control As IRibbonControl, ByRef returnedVal)
    returnedVal = VBA.Date & " " & VBA.Time
    
    'delay the reset
    Application.OnTime Now + TimeValue("00:00:05"), "resetRibbonControls"
End Sub

'Callback for ShowDPMenu getPressed
Sub ShowDPMenu_State(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fShowDPRightClick
End Sub

'Callback for ShowDPGrid getPressed
Sub ShowDPGrid_State(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fShowDPInGrid
End Sub

'ribbon menu click
Sub DatePicker_Click(Optional control As IRibbonControl)
    showDatePicker
End Sub

'called to make sure we updates the labels on insert today in the ribbon
Private Function resetRibbonControls()
    On Error Resume Next
    
    theRibbon.InvalidateControl "btnInsertTodaysDate"
    theRibbon.InvalidateControl "btnInsertTodaysDateTime"
    
    On Error GoTo 0
End Function


