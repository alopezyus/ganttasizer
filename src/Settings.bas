Attribute VB_Name = "Settings"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition
Option Explicit

Public Sub SaveSettings()
    If intEdition = 1 Then Exit Sub
    
    Dim ws As Worksheet
    Dim c As Integer
    Dim strWsName As String
    Dim rng As Range
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    'Contar número de hojas de configuración existentes
    c = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = "ganttasizerSettings" Or ws.Name Like "ganttasizerSettings (*)" Then c = c + 1
    Next
    strWsName = "ganttasizerSettings" & IIf(c > 0, " (" & c + 1 & ")", "")
    
    'Crear hoja configuracion
    Worksheets.Add.Name = strWsName
    Set ws = ActiveWorkbook.Worksheets(strWsName)
    
    Set rng = ws.Cells(2, 1)
    rng = "GANTTASIZER SETTINGS"
        
    i = 2
    rng.Offset(i, 0) = "Group WBS"
    rng.Offset(i, 1) = xl_wbsOutline
    i = i + 1
    rng.Offset(i, 0) = "Calendar Period"
    rng.Offset(i, 1) = xl_period
    i = i + 1
    rng.Offset(i, 0) = "Week Start Day"
    rng.Offset(i, 1) = xl_weekStart
    i = i + 1
    rng.Offset(i, 0) = "Period Width"
    rng.Offset(i, 1) = xl_periodWidth
    i = i + 1
    rng.Offset(i, 0) = "Start Extra Periods"
    rng.Offset(i, 1) = xl_startExtra
    i = i + 1
    rng.Offset(i, 0) = "Finsih Extra Periods"
    rng.Offset(i, 1) = xl_finishExtra
    i = i + 1
    rng.Offset(i, 0) = "Cutoff Date"
    rng.Offset(i, 1) = xl_cutoff
    i = i + 1
    rng.Offset(i, 0) = "Bar Style"
    rng.Offset(i, 1) = xl_barStyle
    i = i + 1
    rng.Offset(i, 0) = "Milestone Style"
    rng.Offset(i, 1) = xl_milStyle
    i = i + 1
    rng.Offset(i, 0) = "Shape Height"
    rng.Offset(i, 1) = xl_shpHgt
    i = i + 1
    rng.Offset(i, 0) = "Label: Description"
    rng.Offset(i, 1) = xl_lblDesc
    i = i + 1
    rng.Offset(i, 0) = "Label: Finish"
    rng.Offset(i, 1) = xl_lblFinish
    i = i + 1
    rng.Offset(i, 0) = "Label: Duration"
    rng.Offset(i, 1) = xl_lblDur
    i = i + 1
    rng.Offset(i, 0) = "Label: Start"
    rng.Offset(i, 1) = xl_lblStart
    i = i + 1
    rng.Offset(i, 0) = "Label: Show on Actuals"
    rng.Offset(i, 1) = xl_lblActuals
    i = i + 1
    rng.Offset(i, 0) = "Remaining Bar Color"
    rng.Offset(i, 1) = xl_rmgBarColor
    i = i + 1
    rng.Offset(i, 0) = "Actual Bar Color"
    rng.Offset(i, 1) = xl_actBarColor
    i = i + 1
    rng.Offset(i, 0) = "BL Bar Color"
    rng.Offset(i, 1) = xl_blBarColor
    i = i + 1
    rng.Offset(i, 0) = "Progress Bar Color"
    rng.Offset(i, 1) = xl_prgBarColor
    i = i + 1
    rng.Offset(i, 0) = "Float Bar Color"
    rng.Offset(i, 1) = xl_FltBarColor
    i = i + 1
    rng.Offset(i, 0) = "Milestone Color"
    rng.Offset(i, 1) = xl_mileColor
    i = i + 1
    rng.Offset(i, 0) = "Cutoff Line Color"
    rng.Offset(i, 1) = xl_cutoffColor
    i = i + 1
    rng.Offset(i, 0) = "Relationship Type"
    rng.Offset(i, 1) = xl_relType
    i = i + 1
    rng.Offset(i, 0) = "Relationship Lag"
    rng.Offset(i, 1) = xl_relLag
    i = i + 1
    rng.Offset(i, 0) = "Connector Style"
    rng.Offset(i, 1) = xl_conStyle
    i = i + 1
    rng.Offset(i, 0) = "Connector Thickness"
    rng.Offset(i, 1) = xl_conThick
    i = i + 1
    rng.Offset(i, 0) = "Sunday"
    rng.Offset(i, 1) = xl_sunday
    i = i + 1
    rng.Offset(i, 0) = "Monday"
    rng.Offset(i, 1) = xl_monday
    i = i + 1
    rng.Offset(i, 0) = "Tuesday"
    rng.Offset(i, 1) = xl_tuesday
    i = i + 1
    rng.Offset(i, 0) = "Wednesday"
    rng.Offset(i, 1) = xl_wednesday
    i = i + 1
    rng.Offset(i, 0) = "Thursday"
    rng.Offset(i, 1) = xl_thursday
    i = i + 1
    rng.Offset(i, 0) = "Friday"
    rng.Offset(i, 1) = xl_friday
    i = i + 1
    rng.Offset(i, 0) = "Saturday"
    rng.Offset(i, 1) = xl_saturday
    i = i + 1
    rng.Offset(i, 0) = "Units Distribution Curve"
    rng.Offset(i, 1) = xl_unitsCurve
    i = i + 1
    rng.Offset(i, 0) = "Auto Update Chart"
    rng.Offset(i, 1) = xl_UpdChart
    i = i + 1
    rng.Offset(i, 0) = "Auto Distribute Units"
    rng.Offset(i, 1) = xl_UpdUnits
    i = i + 1
    rng.Offset(i, 0) = "Auto Update Schedule"
    rng.Offset(i, 1) = xl_UpdSch
    i = i + 1
    rng.Offset(i, 0) = "Auto Update Row Height"
    rng.Offset(i, 1) = xl_UpdRow
    i = i + 1
    rng.Offset(i, 0) = "Update Time Scale with Chart"
    rng.Offset(i, 1) = xl_TimeScl
    i = i + 1
    rng.Offset(i, 0) = "Allow Set Actuals Color"
    rng.Offset(i, 1) = xl_SetActColor
    i = i + 1
    rng.Offset(i, 0) = "Show Base Line"
    rng.Offset(i, 1) = xl_BlBar
    i = i + 1
    rng.Offset(i, 0) = "Show Progress Bar"
    rng.Offset(i, 1) = xl_PrgBar
    i = i + 1
    rng.Offset(i, 0) = "Show Float Bar"
    rng.Offset(i, 1) = xl_FltBar
    i = i + 1
    rng.Offset(i, 0) = "Calendar Exceptions"
    rng.Offset(i, 1) = ActiveWorkbook.CustomDocumentProperties("cdpCalExc").value
    
    rng.EntireColumn.AutoFit
    rng.Offset(0, 1).EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Public Sub LoadSettings()
    If intEdition = 1 Then Exit Sub
    
    Dim rng As Range
    Dim i As Integer
    
    Set rng = ActiveSheet.Cells(2, 1)
    'Comprobar titulos
    
    If Not rng = "GANTTASIZER SETTINGS" Then GoTo IsNotSettings
    i = 2
    If Not rng.Offset(i, 0) = "Group WBS" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Calendar Period" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Week Start Day" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Period Width" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Start Extra Periods" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Finsih Extra Periods" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Cutoff Date" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Bar Style" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Milestone Style" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Shape Height" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Label: Description" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Label: Finish" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Label: Duration" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Label: Start" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Label: Show on Actuals" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Remaining Bar Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Actual Bar Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "BL Bar Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Progress Bar Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Float Bar Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Milestone Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Cutoff Line Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Relationship Type" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Relationship Lag" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Connector Style" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Connector Thickness" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Sunday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Monday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Tuesday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Wednesday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Thursday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Friday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Saturday" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Units Distribution Curve" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Auto Update Chart" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Auto Distribute Units" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Auto Update Schedule" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Auto Update Row Height" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Update Time Scale with Chart" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Allow Set Actuals Color" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Show Base Line" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Show Progress Bar" Then GoTo IsNotSettings
    i = i + 1
    If Not rng.Offset(i, 0) = "Show Float Bar" Then GoTo IsNotSettings

    'Comprobar contenido
    i = 2
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 5) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 6) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 5) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 5) Then GoTo IsNotSettings
    i = i + 1
    If Not (rng.Offset(i, 1) = "" Or IsDate(CDate(rng.Offset(i, 1)))) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 6) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 9) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 3) Then GoTo IsNotSettings
    i = i + 1
    If Not (rng.Offset(i, 1) = 0 Or IsNumeric(rng.Offset(i, 1))) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 3) Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 10) Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not (VarType(rng.Offset(i, 1)) = vbDouble And rng.Offset(i, 1) >= 0 And rng.Offset(i, 1) <= 3) Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
    i = i + 1
    If Not VarType(rng.Offset(i, 1)) = vbBoolean Then GoTo IsNotSettings
'    i = i + 1
'    If Not VarType(rng.Offset(i, 1)) = vbString Then GoTo IsNotSettings

    'Asignar valores
    i = 2
    xl_wbsOutline = rng.Offset(i, 1)
    i = i + 1
    xl_period = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_weekStart = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_periodWidth = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_startExtra = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_finishExtra = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_cutoff = format(rng.Offset(i, 1), "dd/mmm/yyyy")
    i = i + 1
    xl_barStyle = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_milStyle = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_shpHgt = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_lblDesc = rng.Offset(i, 1)
    i = i + 1
    xl_lblFinish = rng.Offset(i, 1)
    i = i + 1
    xl_lblDur = rng.Offset(i, 1)
    i = i + 1
    xl_lblStart = rng.Offset(i, 1)
    i = i + 1
    xl_lblActuals = rng.Offset(i, 1)
    i = i + 1
    xl_rmgBarColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_actBarColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_blBarColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_prgBarColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_FltBarColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_mileColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_cutoffColor = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_relType = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_relLag = rng.Offset(i, 1)
    i = i + 1
    xl_conStyle = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_conThick = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_sunday = rng.Offset(i, 1)
    i = i + 1
    xl_monday = rng.Offset(i, 1)
    i = i + 1
    xl_tuesday = rng.Offset(i, 1)
    i = i + 1
    xl_wednesday = rng.Offset(i, 1)
    i = i + 1
    xl_thursday = rng.Offset(i, 1)
    i = i + 1
    xl_friday = rng.Offset(i, 1)
    i = i + 1
    xl_saturday = rng.Offset(i, 1)
    i = i + 1
    xl_unitsCurve = CInt(rng.Offset(i, 1))
    i = i + 1
    xl_UpdChart = rng.Offset(i, 1)
    i = i + 1
    xl_UpdUnits = rng.Offset(i, 1)
    i = i + 1
    xl_UpdSch = rng.Offset(i, 1)
    i = i + 1
    xl_UpdRow = rng.Offset(i, 1)
    i = i + 1
    xl_TimeScl = rng.Offset(i, 1)
    i = i + 1
    xl_SetActColor = rng.Offset(i, 1)
    i = i + 1
    xl_BlBar = rng.Offset(i, 1)
    i = i + 1
    xl_PrgBar = rng.Offset(i, 1)
    i = i + 1
    xl_FltBar = rng.Offset(i, 1)
    i = i + 1
    updateCustomDocumentProperty "cdpCalExc", rng.Offset(i, 1), msoPropertyTypeString
    
    ribbonUI.Invalidate
    Exit Sub
    
IsNotSettings:
    CustomMsgBox "Ganttasizer Settings are not correctly defined in the current worksheet."
End Sub
