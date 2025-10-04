VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} datepickerform 
   Caption         =   "datepicker"
   ClientHeight    =   10220
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6450
   OleObjectBlob   =   "datepickerform.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "datepickerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim cHighlight As Integer       'current day that has the highlight
Dim c1Highlight_Picker As Integer 'current picker highlight

'click on the month picker
Private Sub monthTitle_Click()

    'set the special highlight
    highlightPicker = cMonth
    
    'reset the current highlight
    c1Highlight_Picker = 0
    
    'show the picker
    toggleMonthPicker
    
End Sub

'click on the year picker
Private Sub yearTitle_Click()

    'set the special highlight
    highlightPicker = cYear
    
    'reset the current highlight
    c1Highlight_Picker = 0
    
    'show the picker
    toggleYearPicker
    
End Sub

'next month button click
Public Sub nextMonthButton_Click()
    
    If pickerMode = 0 Then
        'move the month by one
        cMonth = cMonth + 1
        
        'make sure its a valid month
        If cMonth > 12 Then
            cMonth = 1
            cYear = cYear + 1
        End If
        
        'populate the calendar
        populateDatePickerDays
    
    ElseIf pickerMode = 2 Then
        'move to the next set of picker years
        pickerYearsOffset = pickerYearsOffset + 3
        
        'update the picker
        populatePickerYears
    End If
    
End Sub

'previous month button click
Public Sub prevMonthButton_Click()

    If pickerMode = 0 Then
        'move the month by one
        cMonth = cMonth - 1
        
        'make sure its a valid month
        If cMonth < 1 Then
            cMonth = 12
            cYear = cYear - 1
        End If
        
        'populate the calendar
        populateDatePickerDays
    
    ElseIf pickerMode = 2 Then
        'move to previous set of picker years
        pickerYearsOffset = pickerYearsOffset - 3
        
        'update the picker
        populatePickerYears
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    'get rid of the title bar
    removeCaption Me
    
    'once the borders are gone, we need to resize it to fit the content
    Me.Width = redBG.Width
    Me.Height = barHeight.Height
        
    'init the weekday names
    populateWeekdayNames
    
    'populate the calendar
    populateDatePickerDays
    
    'set the current date and time
    todayButton.Caption = VBA.WeekdayName(VBA.Weekday(VBA.Date, vbSunday)) & ", " & VBA.MonthName(VBA.Month(VBA.Date)) & " " & VBA.day(VBA.Date) & ", " & VBA.Year(VBA.Date)
    timeButton.Caption = VBA.Time
    
    'default the mode
    setPickerMode 0
    
    'start clock timer
    StartTimer
    
End Sub

'processes the mouse move highlight
Sub process2MM(day)

    'see if there is an active highlight
    If cHighlight <> 0 Then
        'see if we need to highlight something else
        If cHighlight <> day Then
            'remove the current highlight
            If cHighlight <= 67 Then
                'remove current highlight
                Me.Controls("daybg" & cHighlight).BackColor = "16777215"
                
                'check to see if its a special date
                specialHighlight Me.Controls("daybg" & cHighlight)
            ElseIf cHighlight = 68 Then
                'insert today button
                datetimebg.BackColor = "16777215"
            ElseIf cHighlight = 69 Then
                'month title caption
                monthTitle.Font.Underline = False
            ElseIf cHighlight = 70 Then
                'year title caption
                yearTitle.Font.Underline = False
            End If
        End If
    End If
    
    'see if we need to apply a new highlight
    If day > 0 Then
        'need to apply new highlight
        If day <= 67 Then
            Me.Controls("daybg" & day).BackColor = "14737632"
        ElseIf day = 68 Then
            'insert today button
            datetimebg.BackColor = "14737632"
        ElseIf day = 69 Then
            'month title caption
            monthTitle.Font.Underline = True
        ElseIf day = 70 Then
            'year title caption
            yearTitle.Font.Underline = True
        End If
        
        'store that its highlighted
        cHighlight = day
    End If
    
End Sub

'process the mouse move with the picker UI
Private Sub process1MMpicker(itemIndex As Integer)
    
    'see if there is an active highlight
    If c1Highlight_Picker <> 0 Then
        'see if we need to highlight something else
        If c1Highlight_Picker <> itemIndex Then
            'remove current highlight
            Me.Controls("mybg" & c1Highlight_Picker).BackColor = "16777215"
            
            'check to see if we should highlight
            specialHighlight Me.Controls("mybg" & c1Highlight_Picker), True
        End If
    End If
    
    'see if we need to apply a new highlight
    If itemIndex > 0 Then
        Me.Controls("mybg" & itemIndex).BackColor = "14737632"
        c1Highlight_Picker = itemIndex
    End If
    
End Sub

'on userform mouse move clear the highlights
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    process2MM 0
End Sub

'on header mouse move clear the highlights
Private Sub redBG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    process2MM 0
End Sub

'*****************************************
'bunch of mouse click handling for the days
'*****************************************
Private Sub day11_Click(): dayPicked 11: End Sub
Private Sub day12_Click(): dayPicked 12: End Sub
Private Sub day13_Click(): dayPicked 13: End Sub
Private Sub day14_Click(): dayPicked 14: End Sub
Private Sub day15_Click(): dayPicked 15: End Sub
Private Sub day16_Click(): dayPicked 16: End Sub
Private Sub day17_Click(): dayPicked 17: End Sub
Private Sub daybg11_Click(): dayPicked 11: End Sub
Private Sub daybg12_Click(): dayPicked 12: End Sub
Private Sub daybg13_Click(): dayPicked 13: End Sub
Private Sub daybg14_Click(): dayPicked 14: End Sub
Private Sub daybg15_Click(): dayPicked 15: End Sub
Private Sub daybg16_Click(): dayPicked 16: End Sub
Private Sub daybg17_Click(): dayPicked 17: End Sub
Private Sub day21_Click(): dayPicked 21: End Sub
Private Sub day22_Click(): dayPicked 22: End Sub
Private Sub day23_Click(): dayPicked 23: End Sub
Private Sub day24_Click(): dayPicked 24: End Sub
Private Sub day25_Click(): dayPicked 25: End Sub
Private Sub day26_Click(): dayPicked 26: End Sub
Private Sub day27_Click(): dayPicked 27: End Sub
Private Sub daybg21_Click(): dayPicked 21: End Sub
Private Sub daybg22_Click(): dayPicked 22: End Sub
Private Sub daybg23_Click(): dayPicked 23: End Sub
Private Sub daybg24_Click(): dayPicked 24: End Sub
Private Sub daybg25_Click(): dayPicked 25: End Sub
Private Sub daybg26_Click(): dayPicked 26: End Sub
Private Sub daybg27_Click(): dayPicked 27: End Sub
Private Sub day31_Click(): dayPicked 31: End Sub
Private Sub day32_Click(): dayPicked 32: End Sub
Private Sub day33_Click(): dayPicked 33: End Sub
Private Sub day34_Click(): dayPicked 34: End Sub
Private Sub day35_Click(): dayPicked 35: End Sub
Private Sub day36_Click(): dayPicked 36: End Sub
Private Sub day37_Click(): dayPicked 37: End Sub
Private Sub daybg31_Click(): dayPicked 31: End Sub
Private Sub daybg32_Click(): dayPicked 32: End Sub
Private Sub daybg33_Click(): dayPicked 33: End Sub
Private Sub daybg34_Click(): dayPicked 34: End Sub
Private Sub daybg35_Click(): dayPicked 35: End Sub
Private Sub daybg36_Click(): dayPicked 36: End Sub
Private Sub daybg37_Click(): dayPicked 37: End Sub
Private Sub day41_Click(): dayPicked 41: End Sub
Private Sub day42_Click(): dayPicked 42: End Sub
Private Sub day43_Click(): dayPicked 43: End Sub
Private Sub day44_Click(): dayPicked 44: End Sub
Private Sub day45_Click(): dayPicked 45: End Sub
Private Sub day46_Click(): dayPicked 46: End Sub
Private Sub day47_Click(): dayPicked 47: End Sub
Private Sub daybg41_Click(): dayPicked 41: End Sub
Private Sub daybg42_Click(): dayPicked 42: End Sub
Private Sub daybg43_Click(): dayPicked 43: End Sub
Private Sub daybg44_Click(): dayPicked 44: End Sub
Private Sub daybg45_Click(): dayPicked 45: End Sub
Private Sub daybg46_Click(): dayPicked 46: End Sub
Private Sub daybg47_Click(): dayPicked 47: End Sub
Private Sub day51_Click(): dayPicked 51: End Sub
Private Sub day52_Click(): dayPicked 52: End Sub
Private Sub day53_Click(): dayPicked 53: End Sub
Private Sub day54_Click(): dayPicked 54: End Sub
Private Sub day55_Click(): dayPicked 55: End Sub
Private Sub day56_Click(): dayPicked 56: End Sub
Private Sub day57_Click(): dayPicked 57: End Sub
Private Sub daybg51_Click(): dayPicked 51: End Sub
Private Sub daybg52_Click(): dayPicked 52: End Sub
Private Sub daybg53_Click(): dayPicked 53: End Sub
Private Sub daybg54_Click(): dayPicked 54: End Sub
Private Sub daybg55_Click(): dayPicked 55: End Sub
Private Sub daybg56_Click(): dayPicked 56: End Sub
Private Sub daybg57_Click(): dayPicked 57: End Sub
Private Sub day61_Click(): dayPicked 61: End Sub
Private Sub day62_Click(): dayPicked 62: End Sub
Private Sub day63_Click(): dayPicked 63: End Sub
Private Sub day64_Click(): dayPicked 64: End Sub
Private Sub day65_Click(): dayPicked 65: End Sub
Private Sub day66_Click(): dayPicked 66: End Sub
Private Sub day67_Click(): dayPicked 67: End Sub
Private Sub daybg61_Click(): dayPicked 61: End Sub
Private Sub daybg62_Click(): dayPicked 62: End Sub
Private Sub daybg63_Click(): dayPicked 63: End Sub
Private Sub daybg64_Click(): dayPicked 64: End Sub
Private Sub daybg65_Click(): dayPicked 65: End Sub
Private Sub daybg66_Click(): dayPicked 66: End Sub
Private Sub daybg67_Click(): dayPicked 67: End Sub
Private Sub timeButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): nowPicked 68, Button: End Sub
Private Sub todayButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): nowPicked 68, Button: End Sub
Private Sub datetimebg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): nowPicked 68, Button: End Sub

'*****************************************
'bunch of mouse moves for the highlights
'*****************************************
Private Sub day11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 11: End Sub
Private Sub day12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 12: End Sub
Private Sub day13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 13: End Sub
Private Sub day14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 14: End Sub
Private Sub day15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 15: End Sub
Private Sub day16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 16: End Sub
Private Sub day17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 17: End Sub
Private Sub daybg11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 11: End Sub
Private Sub daybg12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 12: End Sub
Private Sub daybg13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 13: End Sub
Private Sub daybg14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 14: End Sub
Private Sub daybg15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 15: End Sub
Private Sub daybg16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 16: End Sub
Private Sub daybg17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 17: End Sub
Private Sub day21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 21: End Sub
Private Sub day22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 22: End Sub
Private Sub day23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 23: End Sub
Private Sub day24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 24: End Sub
Private Sub day25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 25: End Sub
Private Sub day26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 26: End Sub
Private Sub day27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 27: End Sub
Private Sub daybg21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 21: End Sub
Private Sub daybg22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 22: End Sub
Private Sub daybg23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 23: End Sub
Private Sub daybg24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 24: End Sub
Private Sub daybg25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 25: End Sub
Private Sub daybg26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 26: End Sub
Private Sub daybg27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 27: End Sub
Private Sub day31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 31: End Sub
Private Sub day32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 32: End Sub
Private Sub day33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 33: End Sub
Private Sub day34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 34: End Sub
Private Sub day35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 35: End Sub
Private Sub day36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 36: End Sub
Private Sub day37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 37: End Sub
Private Sub daybg31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 31: End Sub
Private Sub daybg32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 32: End Sub
Private Sub daybg33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 33: End Sub
Private Sub daybg34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 34: End Sub
Private Sub daybg35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 35: End Sub
Private Sub daybg36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 36: End Sub
Private Sub daybg37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 37: End Sub
Private Sub day41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 41: End Sub
Private Sub day42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 42: End Sub
Private Sub day43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 43: End Sub
Private Sub day44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 44: End Sub
Private Sub day45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 45: End Sub
Private Sub day46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 46: End Sub
Private Sub day47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 47: End Sub
Private Sub daybg41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 41: End Sub
Private Sub daybg42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 42: End Sub
Private Sub daybg43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 43: End Sub
Private Sub daybg44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 44: End Sub
Private Sub daybg45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 45: End Sub
Private Sub daybg46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 46: End Sub
Private Sub daybg47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 47: End Sub
Private Sub day51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 51: End Sub
Private Sub day52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 52: End Sub
Private Sub day53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 53: End Sub
Private Sub day54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 54: End Sub
Private Sub day55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 55: End Sub
Private Sub day56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 56: End Sub
Private Sub day57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 57: End Sub
Private Sub daybg51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 51: End Sub
Private Sub daybg52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 52: End Sub
Private Sub daybg53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 53: End Sub
Private Sub daybg54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 54: End Sub
Private Sub daybg55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 55: End Sub
Private Sub daybg56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 56: End Sub
Private Sub daybg57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 57: End Sub
Private Sub day61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 61: End Sub
Private Sub day62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 62: End Sub
Private Sub day63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 63: End Sub
Private Sub day64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 64: End Sub
Private Sub day65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 65: End Sub
Private Sub day66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 66: End Sub
Private Sub day67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 67: End Sub
Private Sub daybg61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 61: End Sub
Private Sub daybg62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 62: End Sub
Private Sub daybg63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 63: End Sub
Private Sub daybg64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 64: End Sub
Private Sub daybg65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 65: End Sub
Private Sub daybg66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 66: End Sub
Private Sub daybg67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 67: End Sub
Private Sub datetimebg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 68: End Sub
Private Sub monthTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 69: End Sub
Private Sub yearTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process2MM 70: End Sub

'*****************************************
'bunch of mouse moves for the PICKER highlights
'*****************************************
Private Sub my1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 1: End Sub
Private Sub my2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 2: End Sub
Private Sub my3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 3: End Sub
Private Sub my4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 4: End Sub
Private Sub my5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 5: End Sub
Private Sub my6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 6: End Sub
Private Sub my7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 7: End Sub
Private Sub my8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 8: End Sub
Private Sub my9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 9: End Sub
Private Sub my10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 10: End Sub
Private Sub my11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 11: End Sub
Private Sub my12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 12: End Sub
Private Sub mybg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 1: End Sub
Private Sub mybg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 2: End Sub
Private Sub mybg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 3: End Sub
Private Sub mybg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 4: End Sub
Private Sub mybg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 5: End Sub
Private Sub mybg6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 6: End Sub
Private Sub mybg7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 7: End Sub
Private Sub mybg8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 8: End Sub
Private Sub mybg9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 9: End Sub
Private Sub mybg10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 10: End Sub
Private Sub mybg11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 11: End Sub
Private Sub mybg12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 12: End Sub
Private Sub myFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): process1MMpicker 0: End Sub


'*****************************************
'bunch of mouse click handling for the picker
'*****************************************
Private Sub my1_Click(): pickerClicked 1: End Sub
Private Sub my2_Click(): pickerClicked 2: End Sub
Private Sub my3_Click(): pickerClicked 3: End Sub
Private Sub my4_Click(): pickerClicked 4: End Sub
Private Sub my5_Click(): pickerClicked 5: End Sub
Private Sub my6_Click(): pickerClicked 6: End Sub
Private Sub my7_Click(): pickerClicked 7: End Sub
Private Sub my8_Click(): pickerClicked 8: End Sub
Private Sub my9_Click(): pickerClicked 9: End Sub
Private Sub my10_Click(): pickerClicked 10: End Sub
Private Sub my11_Click(): pickerClicked 11: End Sub
Private Sub my12_Click(): pickerClicked 12: End Sub
Private Sub mybg1_Click(): pickerClicked 1: End Sub
Private Sub mybg2_Click(): pickerClicked 2: End Sub
Private Sub mybg3_Click(): pickerClicked 3: End Sub
Private Sub mybg4_Click(): pickerClicked 4: End Sub
Private Sub mybg5_Click(): pickerClicked 5: End Sub
Private Sub mybg6_Click(): pickerClicked 6: End Sub
Private Sub mybg7_Click(): pickerClicked 7: End Sub
Private Sub mybg8_Click(): pickerClicked 8: End Sub
Private Sub mybg9_Click(): pickerClicked 9: End Sub
Private Sub mybg10_Click(): pickerClicked 10: End Sub
Private Sub mybg11_Click(): pickerClicked 11: End Sub
Private Sub mybg12_Click(): pickerClicked 12: End Sub
