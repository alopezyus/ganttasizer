VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'hooks into Excel objects
Private WithEvents c1App As Application
Attribute c1App.VB_VarHelpID = -1

'local vars
Private m_PickerVisible As Boolean

'**************************************************************
'CLASS EVENTS
'**************************************************************

'INIT
Private Sub Class_Initialize()
    On Error Resume Next

    'get the application object
    Set c1App = Application
    
    'set the defaults
    m_PickerVisible = False
    
    On Error GoTo 0
End Sub

'TERMINATE
Private Sub Class_Terminate()
    clearAllUI
End Sub

'**************************************************************
'PROPERTIES
'**************************************************************

'Date Picker Visible Tracking
Public Property Get PickerVisible() As Variant
    PickerVisible = m_PickerVisible
End Property
Public Property Let PickerVisible(ByVal vNewValue As Variant)
    m_PickerVisible = vNewValue
End Property

'**************************************************************
'SHARED UPDATES
'**************************************************************

'sheet changed
Private Sub DoSheetChange()
    processSelectionChange
End Sub

'workbook changed
Private Sub DoWorkbookChange()
    DoSheetChange
End Sub

'clear all UI
Private Sub clearAllUI()
    closeDatePicker
    killGridDP
End Sub

'process the cell changing
Private Sub processSelectionChange()
    'clear the ui
    clearAllUI
    
    'see if we should show the DP icon
    If checkForGridDP Then
        'show the date picker
        createGridDP
    End If
End Sub

'called to see if we should show the dp entry point in the grid
Private Function checkForGridDP() As Boolean
    On Error GoTo check_err
    
    'ref to cell above
    Dim cellAbove As Range
    
    'default to false
    checkForGridDP = False
    
    'be sure the flag is enabled
    'If (fShowDPInGrid) And TypeName(ActiveCell) = "Range" Then
    If TypeName(ActiveCell) = "Range" Then

        Set wsSch = ActiveSheet
        'Comprobar si la hoja tiene encabezados
        varReturn = returnName
        If Not varReturn(0) Then Exit Function
        
        SetHeaderRef (False)
        intLastRow = ActLastRow
        If Not (ActiveCell.row <= intLastRow And ActiveCell.row > rngRef.row) Then Exit Function
        
        If ActiveCell.Column = rngStart.Column Or ActiveCell.Column = rngFinish.Column Or _
            ActiveCell.Column = rngStartBL.Column Or ActiveCell.Column = rngFinishBL.Column Or _
            ActiveCell.Column = rngStartAct.Column Or ActiveCell.Column = rngFinishAct.Column Or _
            ActiveCell.Column = rngResume.Column Or ActiveCell.Column = rngConstraint.Column Then checkForGridDP = True
    End If
    
check_err:
End Function

'**************************************************************
'APP EVENTS
'**************************************************************

Private Sub c1App_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    DoSheetChange
End Sub

'App New workbook created
Private Sub c1App_NewWorkbook(ByVal Wb As Workbook)
    DoWorkbookChange
End Sub

'App before print
Private Sub c1App_WorkbookBeforePrint(ByVal Wb As Workbook, Cancel As Boolean)
    clearAllUI
End Sub

'App before save
Private Sub c1App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    clearAllUI
End Sub

'App window resize
Private Sub c1App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    closeDatePicker
End Sub

'App sheet activated
Private Sub c1App_SheetActivate(ByVal sh As Object)
    DoSheetChange
End Sub

'App workbook activated
Private Sub c1App_WorkbookActivate(ByVal Wb As Workbook)
    DoWorkbookChange
End Sub

