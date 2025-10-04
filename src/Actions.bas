Attribute VB_Name = "Actions"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public Sub iniACT()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    StopEvents
    'stopBarEvents
    SetPrjVar
    booPrjStartSet = True
End Sub

Public Sub finACT()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    StartEvents
    'If xl_UpdChart Then startbarEvents
End Sub

Public Sub ACT_CreateTemplate()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    SetEdition
    StopEvents
    'No se puede llamar a iniACT SetPrjVar porque SetPrjVar intenta comprobar la validez de los encabezados
    
    NewSheet
    'CreateHeaders
    strUpdate = "Template"
    ShowStatusBar
    
    finACT
    
'    CustomMsgBox "Start Creating your Gantt chart:" & vbNewLine & _
'                    "Add your first activity by defining the Activity ID or Description and its Start and Finish Dates, then click on Draw Chart."

End Sub

Public Sub StatusBar_CreateTemplate()
    UpdateProgressBar 0.3
    CreateHeaders
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition
Public Sub ACT_CopyWs()
    If intEdition = 1 Then Exit Sub
    iniACT
    If Not booHeaders Then GoTo finSub
    
    strUpdate = "CopyWs"
    ShowStatusBar
    
finSub:
    finACT
End Sub

Public Sub StatusBar_CopyWs()
    CopySheet
    UpdateProgressBar 1
End Sub

Public Sub ACT_ClearCalendar()
    iniACT
    If Not booHeaders Then GoTo finSub
    
    strUpdate = "ClearCalendar"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_ClearCalendar()
    UpdateProgressBar 0.1
    ClearCalendar
    UpdateProgressBar 1
End Sub

Public Sub ACT_CreateCalendar()
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    wsSch.Select
    strUpdate = "CreateCalendar"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_CreateCalendar()
    UpdateProgressBar 0.1
    ClearCalendar
    UpdateProgressBar 0.4
    CreateCalendar
    UpdateProgressBar 1
End Sub

Public Sub ACT_ClearChart()
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If

    strUpdate = "ClearChart"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_ClearChart()
    UpdateProgressBar 0.1
    ClearChart
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition and Pro Edition
Public Sub ACT_FilterShapes()
    If intEdition > 0 Then Exit Sub
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If

    strUpdate = "FilterShapes"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_FilterShapes()
    UpdateProgressBar 0.1
    FilterShapes
    UpdateProgressBar 1
End Sub

Public Sub ACT_CreateChart_shortcut()
ACT_CreateChart
End Sub

Public Sub ACT_CreateChart(Optional ByVal intRowUpdate As Variant = Empty)
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If

    booFloatBar = False
    
    If IsEmpty(intRowUpdate) Then
        strUpdate = "CreateChart"
        ShowStatusBar
    Else 'sin Status Bar
        ContentsWBS intRowUpdate
        CreateChart intRowUpdate
        If xl_UpdUnits = True Then DistributeUnits intRowUpdate
    End If

finSub:
    finACT
End Sub

Public Sub StatusBar_CreateChart()
    ClearFilter
    
    If xl_TimeScl Then
        ClearCalendar
        UpdateProgressBar 0.1
        CreateCalendar
        intLastCol = CalLastColumn
        intActLastRow = ActLastRow
    End If
    UpdateProgressBar 0.3
    ContentsWBS
    UpdateProgressBar 0.5
    CreateChart
    UpdateProgressBar 0.9
    If xl_UpdUnits Then DistributeUnits
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition
Public Sub ACT_FormatWBS()
    If intEdition = 1 Then Exit Sub
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    strUpdate = "FormatWBS"
    ShowStatusBar
    
finSub:
    finACT
End Sub

Public Sub StatusBar_FormatWBS()
    'Actualización de Progress Bar desde el procedimiento llamado
    ClearFilter
    
    FormatWBS
    UpdateProgressBar 0.3
    ContentsWBS
    UpdateProgressBar 0.5
    CreateChart
    UpdateProgressBar 0.9
    If xl_UpdUnits = True Then DistributeUnits
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition
Public Sub ACT_ClearConnectors()
    If intEdition = 1 Then Exit Sub
    iniACT
    If Not booHeaders Then Exit Sub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    strUpdate = "ClearConnectors"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_ClearConnectors()
    UpdateProgressBar 0.1
    ClearConnectors
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition
Public Sub ACT_CreateConnectors()
    If intEdition = 1 Then Exit Sub
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    strUpdate = "CreateConnectors"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_CreateConnectors()
    ClearFilter
    
    UpdateProgressBar 0.1
    ClearConnectors
    UpdateProgressBar 0.4
    CreateConnectors
    UpdateProgressBar 1
End Sub

'Disabled in Free Edition and Pro Edition
Public Sub ACT_CalculateNetwork_shortcut()
If intEdition > 0 Then Exit Sub
ACT_CalculateNetwork
End Sub

'Disabled in Free Edition and Pro Edition
Public Sub ACT_CalculateNetwork(Optional ByVal booAutoSchedule As Boolean = False)
    If intEdition > 0 Then Exit Sub
    If xl_cutoff = "" Then
        datCutoff = Date
        xl_cutoff = Date
        ribbonUI.InvalidateControl "CutoffDateEdit"
    End If
    
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    WeekCalendar
    If intWorkingDays = 0 Then
        CustomMsgBox "Please, revise calendar." + vbCrLf + "At least one working day must be selected.", vbOKOnly + vbExclamation
        GoTo finSub
    End If
    
    booFloatBar = True
    
    If booAutoSchedule Then
        CalculateSchedule True
        If booLoopStatusPh1 Then
            CustomMsgBox "Loops have been found in the Network.", vbOKOnly + vbExclamation
        ElseIf Not IsEmpty(intRowUpd) Then
            ContentsWBS intRowUpd
            CreateChart intRowUpd
            If xl_UpdUnits = True Then DistributeUnits intRowUpd
        End If
    Else
        strUpdate = "CalculateSchedule"
        ShowStatusBar
        If booLoopStatusPh1 Then
            CustomMsgBox "Loops have been found in the Network.", vbOKOnly + vbExclamation
        Else
            SetPrjVar
            strUpdate = "CreateChart"
            ShowStatusBar
        End If
    End If
    
finSub:
    finACT
End Sub

'Disabled in Free Edition and Pro Edition
Public Sub ACT_DistributeUnits()
    If intEdition > 0 Then Exit Sub
    iniACT
    If Not booHeaders Then GoTo finSub
    If intActLastRow = rngRef.row Then
        CreateActMsg
        GoTo finSub
    End If
    
    strUpdate = "DistributeUnits"
    ShowStatusBar

finSub:
    finACT
End Sub

Public Sub StatusBar_DistributeUnits()
    UpdateProgressBar 0.1
    DistributeUnits
    UpdateProgressBar 1
End Sub


Public Sub ClearFilter()
  If wsSch.FilterMode = True Then
    wsSch.ShowAllData
  End If
End Sub

Private Sub CreateActMsg()
    CustomMsgBox "You need to add an activity first." & vbNewLine & vbNewLine & _
                    "Add an activity by defining the Activity ID or Description and its Start and Finish Dates, then click on Draw Chart."
End Sub

