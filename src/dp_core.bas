Attribute VB_Name = "dp_core"
Global g_oDP As DatePickerManager 'controls the date picker
Global cMonth As Integer 'curent month for the date picker
Global cYear As Integer 'current year for the date picker
Global datePicked As String 'date that get picked and inserted
Global theRibbon As IRibbonUI 'ref to the ribbon
Global fShowDPRightClick As Boolean 'flag if we show the date picker on the right click menu
Global fShowDPInGrid As Boolean 'flag if we show the date picker in the grid
Global highlightDate As Date 'check for cell to highlight in calendar
Global highlightPicker As Integer 'special value to highlight in the "my" picker
Global g_gridDP As shape 'dp entry point for grid
Global pickerMode As Integer '0=normal, 1=months, 2=years
Global pickerYearsOffset As Integer 'tracks how far off of cYear to show in the years picker

'gets the month and year and populates the date picker
Sub populateDatePickerDays()
    
    Dim the_startOfMonth As Date 'date object that is the start of the current month
    Dim trackingDate As Date 'the top left position of the calendar control, then incremented as we fill out the calendar
    Dim iStartOfMonthDay As Integer 'day of the week the month starts on
    
    Dim cDayLabel1 As control 'the label control for the day we're editing
    Dim cDayLabelBG1 As control 'the BG label control for the day we're editing
    
    'get the start of the month
    the_startOfMonth = VBA.DateSerial(cYear, cMonth, 1)
    
    'get the day of the week for the start
    iStartOfMonthDay = VBA.Weekday(the_startOfMonth, vbSunday)
    
    'get the top left date for the calendar
    trackingDate = VBA.DateAdd("d", -iStartOfMonthDay + 1, the_startOfMonth)
    
    'set the month label
    datepickerform.monthTitle.Caption = VBA.MonthName(cMonth, True)
    
    'set the year label
    datepickerform.yearTitle.Caption = cYear

    'loop through each week
    For i = 1 To 6
        
        'loop though each day of the week
        For j = 1 To 7
            'get the controls for the day
            Set cDayLabel1 = datepickerform.Controls("day" & i & j)
            Set cDayLabelBG1 = datepickerform.Controls("dayBG" & i & j)
            
            'set the caption for the day
            cDayLabel1.Caption = VBA.day(trackingDate)
            
            'set the tag info for inserting
            cDayLabel1.Tag = trackingDate
            cDayLabelBG1.Tag = trackingDate
            
            'check to see if we should gray the label
            If VBA.Month(trackingDate) <> cMonth Then
                'not current month, fade some
                cDayLabel1.ForeColor = 8421504
            Else
                'current month, make normal
                cDayLabel1.ForeColor = -2147483630
            End If
            
            'check to see if we should highlight the day
            specialHighlight cDayLabelBG1
            
            'move to the next day
            trackingDate = VBA.DateAdd("d", 1, trackingDate)
        Next j
    Next i
    
End Sub

'highlights days in the calendar
Sub specialHighlight(theBGControl As control, Optional picker As Boolean = False)
    
    'check to see if its already flagged and remove it
    '(bug with changing months)
    If theBGControl.BackColor = 12632319 Then
        theBGControl.BackColor = 16777215
    End If
    
    'see if we should highlight it
    If picker Then
        'use highlight picker for the picker control
        If highlightPicker = theBGControl.Tag Then
            theBGControl.BackColor = 12632319
        End If
    Else
        'use highlight date for calendar
        If highlightDate = theBGControl.Tag Then
            theBGControl.BackColor = 12632319
        End If
    End If
    
End Sub

'fills out Sun - Sat captions
Sub populateWeekdayNames()

    For i = 1 To 7
         datepickerform.Controls("dayofweek" & i).Caption = VBA.WeekdayName(i, True, vbSunday)
    Next i
    
End Sub

'shows the date picker
Sub showDatePicker()
    'make sure the manager is loaded
    ensureDPManager

    'flag that picker is visible
    g_oDP.PickerVisible = True

    'setup the special highlight date
    If fDPstartexc Then
        If frmCalExc.txtStart = "" Then
            highlightDate = VBA.Date
        Else: highlightDate = CDate(frmCalExc.txtStart)
        End If
    ElseIf fDPfinishexc Then
        If frmCalExc.txtFinish = "" Then
            highlightDate = VBA.Date
        Else: highlightDate = CDate(frmCalExc.txtFinish)
        End If
    ElseIf fDPribbon Then
        If xl_cutoff = "" Then
            highlightDate = VBA.Date
        Else: highlightDate = xl_cutoff
        End If
    ElseIf VBA.IsDate(ActiveCell) Then
        'use the current cells value
            highlightDate = VBA.DateValue(ActiveCell)
    Else
        'use today
        highlightDate = VBA.Date
    End If
    
    'set the global month and year
    cMonth = VBA.Month(highlightDate)
    cYear = VBA.Year(highlightDate)
    
    'show the calendar form
    datepickerform.Show
    
    'be sure the form moves to position
    MoveFormToMouse datepickerform
    
    'hook up mousewheel
    DoHookFormScroll datepickerform
End Sub

'closes the date picker
Sub closeDatePicker()
    'ensure DP object
    ensureDPManager
    
    'flags = False
    fDPribbon = False
    fDPstartexc = False
    fDPfinishexc = False

    'check if its visible
    If g_oDP.PickerVisible Then
        'unhook mouse wheel scrolling
        UnhookFormScroll
        
        'get rid of the date picker form
        Unload datepickerform
        
        'update the manager to false
        g_oDP.PickerVisible = False
    End If
End Sub

'user clicked to insert today, do some processing before handing off to day picked
Sub nowPicked(id As Integer, Button As Integer)
    If Button > 1 Then
        'right click or something else, include time
        dayPicked 68, True
    Else
        'left click, just the date
        dayPicked 68
    End If
End Sub

'called by the buttons on the date picked when clicked
Sub dayPicked(id As Integer, Optional inculdeTime As Boolean = False)
    If id <= 67 Then
        'get the tag info for the control
        datePicked = datepickerform.Controls("day" & id).Tag
    ElseIf id = 68 Then
        'get the date and time for now
        datePicked = VBA.Date
        
        'see if we add time
        If inculdeTime Then datePicked = datePicked & " " & VBA.Time
    End If
    
    'insert the date to the cell(s)
    If fDPstartexc Then
        frmCalExc.txtStart = datePicked
        fDPstartexc = False
        StartExcUpdate
    ElseIf fDPfinishexc Then
        frmCalExc.txtFinish = datePicked
        fDPfinishexc = False
        FinishExcUpdate
    ElseIf fDPribbon Then
        xl_cutoff = format(datePicked, "dd/mmm/yyyy")
        ribbonUI.InvalidateControl "CutoffDateEdit"
        fDPribbon = False
    Else
        DoAction Date_Picker, True
    End If
    
    'close the date picker
    closeDatePicker
End Sub

'called by the button on the "my" picker when clicked
Sub pickerClicked(id As Integer)

    'check mode
    If pickerMode = 1 Then
        'update month value
        cMonth = datepickerform.Controls("my" & id).Tag
    ElseIf pickerMode = 2 Then
        'update year value
        cYear = datepickerform.Controls("my" & id).Tag
    End If
    
    'reload
    populateDatePickerDays
    
    'close picker
    setPickerMode 0
End Sub

'destroys the grid entry point
Sub killGridDP()
    On Error Resume Next
    
    'delete it
    g_gridDP.Delete
    
    'Delete all grid shapes that are still on the chart
    Dim d As shape
    For Each d In ActiveSheet.Shapes
        If d.Name = "VB_GRID" Then d.Delete
    Next

    'make sure its set to nothing
    Set g_gridDP = Nothing
    
    On Error GoTo 0
End Sub

'makes the grid entry point and places it
Sub createGridDP()
    On Error GoTo create_err
    
    'be sure the current one is deleted
    If Not (g_gridDP Is Nothing) Then
        killGridDP
    End If
    
    'create a new one
    Set g_gridDP = ActiveSheet.Shapes.AddShape(msoShapeRectangle, ActiveCell.Left + ActiveCell.Width + 6, ActiveCell.Top, 12, 12)
    
    'add image
    With g_gridDP.Fill
        .Visible = msoTrue
        .UserPicture pathToIcon
        .TextureTile = msoFalse
        .RotateWithObject = msoTrue
    End With
    
    g_gridDP.Name = "VB_GRID"
    
    'format the border away
    g_gridDP.Line.Visible = msoFalse
    
    'add click event
    g_gridDP.OnAction = "gridDP_Click"
    
create_err:
End Sub

'user clicked on the gridDP entry
Public Sub gridDP_Click() ' #visible
    If g_oDP.PickerVisible Then
        closeDatePicker
    Else
        fDPribbon = False
        fDPstartexc = False
        fDPfinishexc = False
        showDatePicker
    End If
End Sub

'toggles the month picker
Sub toggleMonthPicker()
    If pickerMode = 1 Then
        setPickerMode 0
    Else
        setPickerMode 1
        showPicker
    End If
End Sub

'toggles the year picker
Sub toggleYearPicker()
    If pickerMode = 2 Then
        setPickerMode 0
    Else
        setPickerMode 2
        showPicker
    End If
End Sub

'shows the my picker
Sub showPicker()
    datepickerform.Controls("myFrame").Top = 30
End Sub

'hides the my picker
Sub hidePicker()
    datepickerform.Controls("myFrame").Top = datepickerform.Height + 20
End Sub

'set the mode for the picker, and updates the UI
Sub setPickerMode(mode As Integer)
    If mode = 0 Then
        'hide the picker
        pickerMode = 0
        hidePicker
        
        'show buttons
        showPrevNextMonthButtons
    ElseIf mode = 1 Then
        'populate months
        populatePickerMonths
        
        'hide buttons
        hidePrevNextMonthButtons
        
        'set the flag to months
        pickerMode = 1
    ElseIf mode = 2 Then
        'reset offset for year scrolling
        pickerYearsOffset = -6
        
        'populate years
        populatePickerYears
        
        'show buttons
        showPrevNextMonthButtons
        
        'set the flag
        pickerMode = 2
    End If
End Sub

'updates the "my" picker strings to months
Sub populatePickerMonths()
    'ref to the control to update
    Dim myControl As control
    Dim myBGControl As control
    
    'loop through and populate the months
    For i = 1 To 12
        'get a ref to the control to update
        Set myControl = datepickerform.Controls("my" & i)
        Set myBGControl = datepickerform.Controls("mybg" & i)
        
        'set the string
        myControl.Caption = VBA.MonthName(i, True)
        
        'set the tag to the value we'll act on later
        myControl.Tag = i
        myBGControl.Tag = i
        
        'clear any highlight
        myBGControl.BackColor = 16777215
        
        'see if we should highlight
        specialHighlight myBGControl, True
    Next i
End Sub

Sub populatePickerYears()
    'ref to the control to update
    Dim myControl As control
    Dim myBGControl As control
    Dim loopStart As Integer
    
    loopStart = cYear + pickerYearsOffset
    
    'loop through and populate the months
    For i = 1 To 12
        'get a ref to the control to update
        Set myControl = datepickerform.Controls("my" & i)
        Set myBGControl = datepickerform.Controls("mybg" & i)
        
        'set the string
        myControl.Caption = loopStart
        
        'set the tag to the value we'll act on later
        myControl.Tag = loopStart
        myBGControl.Tag = loopStart
        
        'clear any highlight
        myBGControl.BackColor = 16777215
        
        'see if we should highlight
        specialHighlight myBGControl, True
        
        'inc the year
        loopStart = loopStart + 1
    Next i
End Sub

'hides/shows the buttons when they don't apply (for month picker)
Sub hidePrevNextMonthButtons()
    datepickerform.Controls("prevMonthButton").Visible = False
    datepickerform.Controls("nextMonthButton").Visible = False
End Sub
Sub showPrevNextMonthButtons()
    datepickerform.Controls("prevMonthButton").Visible = True
    datepickerform.Controls("nextMonthButton").Visible = True
End Sub

