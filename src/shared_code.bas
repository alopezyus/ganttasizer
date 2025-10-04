Attribute VB_Name = "shared_code"
'Action types
Enum ActionOnCell
    Date_Picker
End Enum

'called by the click events to begin the action
Sub DoAction(iType As ActionOnCell, Optional noTableGrow As Boolean = False)
    
    'loop through the cells
    LoopCellsController iType, noTableGrow

End Sub

'detects tables and multiple ranges before calling the sub to change the cell
Sub LoopCellsController(iType As ActionOnCell, Optional noTableGrow As Boolean = False)
    Dim Target As Range     'target range object to loop through
    Dim block As Range      'each smaller area of the target to loop through
    
    'see if we have a valid selection
    If TypeName(Selection) = "Range" Then
        'set the target to the current selection
        Set Target = Selection
        
        'see if its just a single cell
        If Target.Cells.Count = 1 Then
            'see if we're in a table and grow
            If IsCellInTable(Target) And noTableGrow = False Then
            'make sure the table has rows
                If Target.ListObject.ListRows.Count > 1 Then
                    'get the column range and fill it
                    Set Target = Target.ListObject.ListColumns(Target.Column - Target.ListObject.DataBodyRange.Column + 1).DataBodyRange
                End If
            End If
        End If
        
        'loop through all the selection ranges
        For Each block In Target.Areas
            'populate the cells
            LoopCellsPopulator block, iType
        Next block
    
    Else
        'not a valid range
    End If
    
End Sub

'performs the action on the range passed in
Sub LoopCellsPopulator(oRange As Range, iType As ActionOnCell)
    Dim cell As Range
    
    'loop through the cells
    For Each cell In oRange
        Select Case iType
            'date picker
            Case Is = ActionOnCell.Date_Picker
                cell = GetDatePicked
        End Select
    Next cell

End Sub

'called for each cell the date picker is putting a value in
Private Function GetDatePicked() As Date
    GetDatePicked = datePicked
End Function

'returns true if active cell is in a table and false if it isn't.
Function IsCellInTable(Target As Range) As Boolean

    On Error Resume Next
    
    'check to see if there is a table name
    IsCellInTable = (Target.ListObject.Name <> "")
    
    On Error GoTo 0
    
End Function

'returns the number of days in a month
'Function daysInMonth(theDate As Date)
'    daysInMonth = day(DateSerial(Year(theDate), Month(theDate) + 1, 1) - 1)
'End Function

'makes sure the date picker manager is loaded
Sub ensureDPManager()
    If g_oDP Is Nothing Then Set g_oDP = New DatePickerManager
End Sub

'loads the global settings
Sub LoadGlobalSettings()
    
    'show on right click menu
    fShowDPRightClick = VBA.GetSetting("samradapps_datepicker", "ribbon", "fShowDPRightClick", True)
    
    'show in cell
    fShowDPInGrid = VBA.GetSetting("samradapps_datepicker", "ribbon", "fShowDPInGrid", True)
    
    'update right click menus
    updateRightClickMenu
    
    'get the screen height
    getScreenHeight
    
    'make sure the icon exists
    CreateDPIcon
End Sub
