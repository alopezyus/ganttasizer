Attribute VB_Name = "UI_Actions"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Public ribbonUI As IRibbonUI

Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (destination As Any, source As Any, _
    ByVal length As Long)
    
'Callback for customUI.onLoad
Sub onLoad_Ribbon(Ribbon As IRibbonUI)
    'PURPOSE: Run code when Ribbon loads the UI to store Ribbon Object's Pointer ID code
    
    'Handle variable declaration if 32-bit or 64-bit Excel
    #If VBA7 Then
    Dim StoreRibbonPointer As LongPtr
    #Else
    Dim StoreRibbonPointer As Long
    #End If
    
    'Store Ribbon Object to Public variable
    Set ribbonUI = Ribbon
    
    'Store pointer to IRibbonUI in a Named Range within add-in file
    StoreRibbonPointer = ObjPtr(Ribbon)
    ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:=StoreRibbonPointer
  
    'Set ribbonUI = Ribbon
    
    StartEvents
    If (Application.ActiveProtectedViewWindow Is Nothing) Then
        startbarEvents
    End If

    initializeRibbonVar
    
End Sub

Sub initializeRibbonVar()
    xl_wbsOutline = updateProjectVarProperty("cdpWBSoutline", True)
    xl_period = updateProjectVarProperty("cdpPeriod", 2)
    xl_weekStart = updateProjectVarProperty("cdpWeekStart", 1)
    xl_periodWidth = updateProjectVarProperty("cdpPeriodWidth", 3)
    xl_startExtra = updateProjectVarProperty("cdpStartExtra", 1)
    xl_finishExtra = updateProjectVarProperty("cdpFinishExtra", 1)
    xl_cutoff = updateProjectVarProperty("cdpCutoff", "")
    xl_barStyle = updateProjectVarProperty("cdpBarStyle", 1)
    xl_milStyle = updateProjectVarProperty("cdpMilStyle", 0)
    xl_shpHgt = updateProjectVarProperty("cdpShpHgt", 3)
    xl_lblDesc = updateProjectVarProperty("cdpLblDesc", True)
    xl_lblFinish = updateProjectVarProperty("cdpLblFinish", False)
    xl_lblDur = updateProjectVarProperty("cdpLblDur", False)
    xl_lblStart = updateProjectVarProperty("cdpLblStart", False)
    xl_lblActuals = updateProjectVarProperty("cdpLblActuals", False)
    xl_rmgBarColor = updateProjectVarProperty("cdpRmgBarColor", 3)
    xl_actBarColor = updateProjectVarProperty("cdpActBarColor", 5)
    xl_blBarColor = updateProjectVarProperty("cdpBlBarColor", 1)
    xl_prgBarColor = updateProjectVarProperty("cdpPrgBarColor", 2)
    xl_FltBarColor = updateProjectVarProperty("cdpFltBarColor", 9)
    xl_mileColor = updateProjectVarProperty("cdpMileColor", 7)
    xl_cutoffColor = updateProjectVarProperty("cdpCutoffColor", 6)
    xl_windowColor = updateProjectVarProperty("cdpWindowColor", 8)
    xl_relType = updateProjectVarProperty("cdpRelType", 0)
    xl_relLag = updateProjectVarProperty("cdpRelLag", 0)
    xl_conStyle = updateProjectVarProperty("cdpConStyle", 0)
    xl_conThick = updateProjectVarProperty("cdpConThick", 2)
    xl_UpdChart = updateProjectVarProperty("cdpUpdChart", False)
    xl_UpdUnits = updateProjectVarProperty("cdpUpdUnits", False)
    xl_UpdRow = updateProjectVarProperty("cdpUpdRow", False)
    xl_TimeScl = updateProjectVarProperty("cdpTimeScl", True)
    xl_SetActColor = updateProjectVarProperty("cdpSetActColor", False)
    xl_UpdSch = updateProjectVarProperty("cdpUpdSch", False)
    xl_BlBar = updateProjectVarProperty("cdpBlBar", True)
    xl_PrgBar = updateProjectVarProperty("cdpPrgBar", False)
    xl_FltBar = updateProjectVarProperty("cdpFltBar", False)
    xl_sunday = updateProjectVarProperty("cdpSunday", False)
    xl_monday = updateProjectVarProperty("cdpMonday", True)
    xl_tuesday = updateProjectVarProperty("cdpTuesday", True)
    xl_wednesday = updateProjectVarProperty("cdpWednesday", True)
    xl_thursday = updateProjectVarProperty("cdpThursday", True)
    xl_friday = updateProjectVarProperty("cdpFriday", True)
    xl_saturday = updateProjectVarProperty("cdpSaturday", False)
    xl_unitsCurve = updateProjectVarProperty("cdpUnitsCurve", 1)
    
End Sub

Sub RefreshRibbon()
'PURPOSE: Refresh Ribbon UI

Dim myRibbon As Object

SetEdition

booPrjStartSet = False

On Error GoTo RestartExcel
  If ribbonUI Is Nothing Then
    Set ribbonUI = GetRibbon(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", ""))
  End If
  
  'Redo Ribbon Load
    ribbonUI.Invalidate
    
    StartEvents
    startbarEvents
    
    initializeRibbonVar
On Error GoTo 0

Exit Sub

'ERROR MESSAGES:
RestartExcel:
'MsgBox "Please restart Excel for Ribbon UI changes to take affect", , "Ribbon UI Refresh Failed"

End Sub
    
#If VBA7 Then
  Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
  Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If

  Dim objRibbon As Object
  
  CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = objRibbon
  Set objRibbon = Nothing
  
End Function

'ACTION BUTTONS-------------------------------------------------------------------------------
'Callback for templateBtn onAction
Sub codeTemplate(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CreateTemplate
End Sub

'Callback for copyWsBtn onAction
Sub codeCopyWs(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CopyWs
End Sub

'Callback for colBtn onAction
Sub codeColumns(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
openColSelForm
End Sub

'Callback for loadSetBtn onAction
Sub codeLoadSet(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
LoadSettings
End Sub

'Callback for saveSetBtnM onAction
Sub codeSaveSet(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
SaveSettings
End Sub

'Callback for addCalBtn onAction
Sub codeAddCal(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CreateCalendar
End Sub

'Callback for delCalBtn onAction
Sub codeDelCal(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_ClearCalendar
End Sub

'Callback for addChrBtn onAction
Sub codeAddChr(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CreateChart
End Sub

'Callback for delChrBtn onAction
Sub codeDelChr(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_ClearChart
End Sub

'Callback for chkFilterBtnM onAction
Sub codeChkFilter(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_FilterShapes
End Sub

'Callback for formatWBSBtnM onAction
Sub codeFormatWBS(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_FormatWBS
End Sub

'Callback for indentBtn onAction
Sub codeIndentWBS(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
WBS_indent
End Sub

'Callback for outdentBtn onAction
Sub codeOutdentWBS(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
WBS_indent False
End Sub

'Callback for addConBtn onAction
Sub codeAddCon(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CreateConnectors
End Sub

'Callback for delConBtn onAction
Sub codeDelCon(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_ClearConnectors
End Sub

'Callback for netBtn onAction
Sub codeCalc(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_CalculateNetwork
End Sub

'Callback for crvBtnM onAction
Sub codeDist(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
ACT_DistributeUnits
End Sub

'Callback for qrelBtn onAction
Sub codeAddRelations(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
MngRel
End Sub

'Callback for delRelBtn onAction
Sub codeDelRelations(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
MngRel True
End Sub

'Callback for infoBtn onAction
Sub codeLicense(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
LicenseFile
End Sub

'Callback for helpBtn onAction
Sub codeHelp(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
HelpFile
End Sub

'Callback for webBtn onAction
Sub codeApp(control As IRibbonControl)
    'ActiveWorkbook.FollowHyperlink "https://ganttasizer.com"
    Call RefreshRibbon
    checkDatePicker
    InfoGanttasizer
End Sub

'Callback for excBtn onAction
Sub calendarExc(control As IRibbonControl)
Call RefreshRibbon
checkDatePicker
openCalendarExceptions
End Sub

'WBS-----------------------------------------------------------------------------------------------------
'Callback for wbsOutline onAction
Sub wbsOutline_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_wbsOutline = pressed
    ribbonUI.InvalidateControl "wbsOutline"
End Sub

'Callback for wbsOutline getPressed
Sub wbsOutline_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_wbsOutline
    updateCustomDocumentProperty "cdpWBSoutline", xl_wbsOutline, msoPropertyTypeBoolean
End Sub

'CALENDAR------------------------------------------------------------------------------------------------
'Callback for periodDropDown onAction
Sub period_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_period = index
    ribbonUI.InvalidateControl "periodDropDown"
End Sub

'Callback for periodDropDown getSelectedItemIndex
Sub period_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_period
    updateCustomDocumentProperty "cdpPeriod", xl_period, msoPropertyTypeNumber
End Sub

'Callback for weekStartDropDown onAction
Sub wStart_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_weekStart = index
    ribbonUI.InvalidateControl "weekStartDropDown"
End Sub

'Callback for weekStartDropDown getSelectedItemIndex
Sub weekStart_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_weekStart
    updateCustomDocumentProperty "cdpWeekStart", xl_weekStart, msoPropertyTypeNumber
End Sub

'Callback for periodWidthDropDown onAction
Sub pWidth_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_periodWidth = index
    ribbonUI.InvalidateControl "periodWidthDropDown"
End Sub

'Callback for periodWidthDropDown getSelectedItemIndex
Sub periodWidth_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_periodWidth
    updateCustomDocumentProperty "cdpPeriodWidth", xl_periodWidth, msoPropertyTypeNumber
End Sub

'Callback for startCalExtraDDown onAction
Sub startCalExtra_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_startExtra = index
    ribbonUI.InvalidateControl "startCalExtraDDown"
End Sub

'Callback for startCalExtraDDown getSelectedItemIndex
Sub startExtra_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_startExtra
    updateCustomDocumentProperty "cdpStartExtra", xl_startExtra, msoPropertyTypeNumber
End Sub

'Callback for finishCalExtraDDown onAction
Sub finishCalExtra_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_finishExtra = index
    ribbonUI.InvalidateControl "finishCalExtraDDown"
End Sub

'Callback for finishCalExtraDDown getSelectedItemIndex
Sub finishExtra_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_finishExtra
    updateCustomDocumentProperty "cdpFinishExtra", xl_finishExtra, msoPropertyTypeNumber
End Sub

'CUTOFF DATE---------------------------------------------------------------------------------------------
'Callback for CutoffDateEdit onChange
Sub cutoff_onChange(control As IRibbonControl, text As String)
    Call RefreshRibbon
    checkDatePicker
    If IsDate(text) Then
        xl_cutoff = format(text, "dd/mmm/yyyy")
    Else
        xl_cutoff = ""
    End If
    ribbonUI.InvalidateControl "CutoffDateEdit"
End Sub

'Callback for CutoffDateEdit getText
Sub cutoff_getDate(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_cutoff
    updateCustomDocumentProperty "cdpCutoff", xl_cutoff, msoPropertyTypeString
End Sub

'Callback for eraseBtn onAction
Sub eraseBtn_onAction(control As IRibbonControl)
    Call RefreshRibbon
    checkDatePicker
    xl_cutoff = ""
    ribbonUI.InvalidateControl "CutoffDateEdit"
End Sub

'Callback for pickerBtn onAction
Sub pickerBtn_onAction(control As IRibbonControl)
    Call RefreshRibbon
    ensureDPManager
    If g_oDP.PickerVisible Then
        closeDatePicker
    Else
        fDPribbon = True
        showDatePicker
    End If
End Sub

'CHART---------------------------------------------------------------------------------------------------
'Callback for barStyleDropDown onAction
Sub barStyle_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_barStyle = index
    ribbonUI.InvalidateControl "barStyleDropDown"
End Sub

'Callback for barStyleDropDown getSelectedItemIndex
Sub barStyle_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_barStyle
    updateCustomDocumentProperty "cdpBarStyle", xl_barStyle, msoPropertyTypeNumber
End Sub

'Callback for milStyleDropDown onAction
Sub milStyle_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_milStyle = index
    ribbonUI.InvalidateControl "milStyleDropDown"
End Sub

'Callback for milStyleDropDown getSelectedItemIndex
Sub milStyle_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_milStyle
    updateCustomDocumentProperty "cdpMilStyle", xl_milStyle, msoPropertyTypeNumber
End Sub

'Callback for heightEdit onAction
Sub barHgt_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_shpHgt = index
    ribbonUI.InvalidateControl "heightEdit"
End Sub

'Callback for heightEdit getSelectedItemIndex
Sub shpHgt_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_shpHgt
    updateCustomDocumentProperty "cdpShpHgt", xl_shpHgt, msoPropertyTypeNumber
End Sub

'Callback for lblDescChk onAction
Sub lblDesc_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_lblDesc = pressed
    ribbonUI.InvalidateControl "lblDescChk"
End Sub

'Callback for lblDescChk getPressed
Sub lblDesc_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_lblDesc
    updateCustomDocumentProperty "cdpLblDesc", xl_lblDesc, msoPropertyTypeBoolean
End Sub

'Callback for lblFinishChk onAction
Sub lblFinish_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_lblFinish = pressed
    ribbonUI.InvalidateControl "lblFinishChk"
End Sub

'Callback for lblFinishChk getPressed
Sub lblFinish_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_lblFinish
    updateCustomDocumentProperty "cdpLblFinish", xl_lblFinish, msoPropertyTypeBoolean
End Sub

'Callback for lblDurChk onAction
Sub lblDur_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_lblDur = pressed
    ribbonUI.InvalidateControl "lblDurChk"
End Sub

'Callback for lblDurChk getPressed
Sub lblDur_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_lblDur
    updateCustomDocumentProperty "cdpLblDur", xl_lblDur, msoPropertyTypeBoolean
End Sub

'Callback for lblStartChk onAction
Sub lblStart_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_lblStart = pressed
    ribbonUI.InvalidateControl "lblStartChk"
End Sub

'Callback for lblStartChk getPressed
Sub lblStart_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_lblStart
    updateCustomDocumentProperty "cdpLblStart", xl_lblStart, msoPropertyTypeBoolean
End Sub

'Callback for lblActualsChk onAction
Sub lblActuals_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_lblActuals = pressed
    ribbonUI.InvalidateControl "lblActualsChk"
End Sub

'Callback for lblActualsChk getPressed
Sub lblActuals_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_lblActuals
    updateCustomDocumentProperty "cdpLblActuals", xl_lblActuals, msoPropertyTypeBoolean
End Sub

'RELATIONSHIPS--------------------------------------------------------------------------------------------
'Callback for conStyleDropDown onAction
Sub conStyle_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_conStyle = index
    ribbonUI.InvalidateControl "conStyleDropDown"
End Sub

'Callback for relTypeDropDown getSelectedItemIndex
Sub relType_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_relType
    updateCustomDocumentProperty "cdpRelType", xl_relType, msoPropertyTypeNumber
End Sub

'Callback for relTypeDropDown onAction
Sub relType_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_relType = index
    ribbonUI.InvalidateControl "relTypeDropDown"
End Sub

'Callback for LagEdit onChange
Sub lag_onChange(control As IRibbonControl, text As String)
    Call RefreshRibbon
    checkDatePicker
    If IsNumeric(text) Then
        xl_relLag = CInt(text)
    Else
        xl_relLag = 0
    End If
    ribbonUI.InvalidateControl "LagEdit"
End Sub

'Callback for LagEdit getText
Sub lag_getInt(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_relLag
    updateCustomDocumentProperty "cdpRelLag", xl_relLag, msoPropertyTypeNumber
End Sub

'Callback for conStyleDropDown getSelectedItemIndex
Sub conStyle_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_conStyle
    updateCustomDocumentProperty "cdpConStyle", xl_conStyle, msoPropertyTypeNumber
End Sub

'Callback for conThikEdit onAction
Sub ConThik_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_conThick = index
    ribbonUI.InvalidateControl "conThikEdit"
End Sub

'Callback for conThikEdit getSelectedItemIndex
Sub conThik_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_conThick
    updateCustomDocumentProperty "cdpConThick", xl_conThick, msoPropertyTypeNumber
End Sub

''Callback for conTransEdit onAction
'Sub Transp_onAction(control As IRibbonControl, id As String, index As Integer)
'    checkDatePicker
'    xl_conTransp = index
'    ribbonUI.InvalidateControl "conTransEdit"
'End Sub
'
''Callback for conTransEdit getSelectedItemIndex
'Sub conTrans_getIndex(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = xl_conTransp
'    updateCustomDocumentProperty "cdpConTransp", xl_conTransp, msoPropertyTypeNumber
'End Sub

'CALCULATIONS---------------------------------------------------------------------------------------------
'Callback for autoUpdateChk onAction
Sub autoChart_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_UpdChart = pressed
    If xl_UpdChart Then
        startbarEvents
    Else
        startbarEvents
        'stopBarEvents
    End If
    ribbonUI.InvalidateControl "autoUpdateChk"
End Sub

'Callback for autoUpdateChk getPressed
Sub autoUpd_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_UpdChart
    updateCustomDocumentProperty "cdpUpdChart", xl_UpdChart, msoPropertyTypeBoolean
End Sub

'Callback for autoUpdateUnits onAction
Sub autoUnits_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_UpdUnits = pressed
    ribbonUI.InvalidateControl "autoUpdateUnits"
End Sub

'Callback for autoUpdateUnits getPressed
Sub autoUnits_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_UpdUnits
    updateCustomDocumentProperty "cdpUpdUnits", xl_UpdUnits, msoPropertyTypeBoolean
End Sub

'Callback for autoRowChk onAction
Sub autoRow_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_UpdRow = pressed
    ribbonUI.InvalidateControl "autoRowChk"
End Sub

'Callback for autoRowChk getPressed
Sub autoRow_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_UpdRow
    updateCustomDocumentProperty "cdpUpdRow", xl_UpdRow, msoPropertyTypeBoolean
End Sub

'Callback for autoTimeScl onAction
Sub autoTimeScl_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_TimeScl = pressed
    ribbonUI.InvalidateControl "autoTimeScl"
End Sub

'Callback for autoTimeScl getPressed
Sub autoTimeScl_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_TimeScl
    updateCustomDocumentProperty "cdpTimeScl", xl_TimeScl, msoPropertyTypeBoolean
End Sub

'Callback for autoSetActColor onAction
Sub autoSetActColor_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_SetActColor = pressed
    ribbonUI.InvalidateControl "autoSetActColor"
End Sub

'Callback for autoSetActColor getPressed
Sub autoSetActColor_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_SetActColor
    updateCustomDocumentProperty "cdpSetActColor", xl_SetActColor, msoPropertyTypeBoolean
End Sub

'Callback for autoUpdSch onAction
Sub autoUpdSch_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_UpdSch = pressed
    ribbonUI.InvalidateControl "autoUpdSch"
End Sub

'Callback for autoUpdSch getPressed
Sub autoUpdSch_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_UpdSch
    updateCustomDocumentProperty "cdpUpdSch", xl_UpdSch, msoPropertyTypeBoolean
End Sub

'Callback for BlBarChk onAction
Sub BlBar_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_BlBar = pressed
    ribbonUI.InvalidateControl "BlBarChk"
End Sub

'Callback for BlBarChk getPressed
Sub BlBar_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_BlBar
    updateCustomDocumentProperty "cdpBlBar", xl_BlBar, msoPropertyTypeBoolean
End Sub

'Callback for PrgBarChk onAction
Sub PrgBar_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_PrgBar = pressed
    ribbonUI.InvalidateControl "PrgBarChk"
End Sub

'Callback for PrgBarChk getPressed
Sub PrgBar_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_PrgBar
    updateCustomDocumentProperty "cdpPrgBar", xl_PrgBar, msoPropertyTypeBoolean
End Sub

'Callback for FltBarChk onAction
Sub FltBar_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_FltBar = pressed
    ribbonUI.InvalidateControl "FltBarChk"
End Sub

'Callback for FltBarChk getPressed
Sub FltBar_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_FltBar
    updateCustomDocumentProperty "cdpFltBar", xl_FltBar, msoPropertyTypeBoolean
End Sub

'Callback for SundayChk onAction
Sub sunday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_sunday = pressed
    ribbonUI.InvalidateControl "SundayChk"
End Sub

'Callback for SundayChk getPressed
Sub sunday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_sunday
    updateCustomDocumentProperty "cdpSunday", xl_sunday, msoPropertyTypeBoolean
End Sub

'Callback for MondayChk onAction
Sub monday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_monday = pressed
    ribbonUI.InvalidateControl "MondayChk"
End Sub

'Callback for MondayChk getPressed
Sub monday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_monday
    updateCustomDocumentProperty "cdpMonday", xl_monday, msoPropertyTypeBoolean
End Sub

'Callback for TuesdayChk onAction
Sub tuesday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_tuesday = pressed
    ribbonUI.InvalidateControl "TuesdayChk"
End Sub

'Callback for TuesdayChk getPressed
Sub tuesday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_tuesday
    updateCustomDocumentProperty "cdpTuesday", xl_tuesday, msoPropertyTypeBoolean
End Sub

'Callback for WednesdayChk onAction
Sub wednesday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_wednesday = pressed
    ribbonUI.InvalidateControl "WednesdayChk"
End Sub

'Callback for WednesdayChk getPressed
Sub wednesday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_wednesday
    updateCustomDocumentProperty "cdpWednesday", xl_wednesday, msoPropertyTypeBoolean
End Sub

'Callback for ThursdayChk onAction
Sub thursday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_thursday = pressed
    ribbonUI.InvalidateControl "ThursdayChk"
End Sub

'Callback for ThursdayChk getPressed
Sub thursday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_thursday
    updateCustomDocumentProperty "cdpThursday", xl_thursday, msoPropertyTypeBoolean
End Sub

'Callback for FridayChk onAction
Sub friday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_friday = pressed
    ribbonUI.InvalidateControl "FridayChk"
End Sub

'Callback for FridayChk getPressed
Sub friday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_friday
    updateCustomDocumentProperty "cdpFriday", xl_friday, msoPropertyTypeBoolean
End Sub

'Callback for SaturdayChk onAction
Sub saturday_onAction(control As IRibbonControl, pressed As Boolean)
    Call RefreshRibbon
    checkDatePicker
    xl_saturday = pressed
    ribbonUI.InvalidateControl "SaturdayChk"
End Sub

'Callback for SaturdayChk getPressed
Sub saturday_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_saturday
    updateCustomDocumentProperty "cdpSaturday", xl_saturday, msoPropertyTypeBoolean
End Sub

'Callback for resCurveDropDown onAction
Sub resCurve_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    checkDatePicker
    xl_unitsCurve = index
    ribbonUI.InvalidateControl "resCurveDropDown"
End Sub

'Callback for resCurveDropDown getSelectedItemIndex
Sub resCurve_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_unitsCurve
    
    updateCustomDocumentProperty "cdpUnitsCurve", xl_unitsCurve, msoPropertyTypeNumber
End Sub

'COLORS---------------------------------------------------------------------------------------------------
'Callback for rmgBarColor getText
Sub rmgBarColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_rmgBarColor
    updateCustomDocumentProperty "cdpRmgBarColor", xl_rmgBarColor, msoPropertyTypeNumber
End Sub

'Callback for rmngBarColorGal onAction
Sub rmgBarColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_rmgBarColor = index
    ribbonUI.InvalidateControl "rmgBarColor"
    ribbonUI.Invalidate
End Sub

'Callback for rmngBarColorGal getVisible
Sub rmgBarColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal0 getVisible
Sub rmgBarColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal1 getVisible
Sub rmgBarColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal2 getVisible
Sub rmgBarColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal3 getVisible
Sub rmgBarColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal4 getVisible
Sub rmgBarColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal5 getVisible
Sub rmgBarColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal6 getVisible
Sub rmgBarColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal7 getVisible
Sub rmgBarColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal8 getVisible
Sub rmgBarColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal9 getVisible
Sub rmgBarColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_rmgBarColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


'Callback for actBarColor getText
Sub actBarColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_actBarColor
    updateCustomDocumentProperty "cdpActBarColor", xl_actBarColor, msoPropertyTypeNumber
End Sub

'Callback for rmngBarColorGal onAction
Sub actBarColor_onAction(control As IRibbonControl, id As String, index As Integer)
    xl_actBarColor = index
    ribbonUI.InvalidateControl "actBarColor"
    ribbonUI.Invalidate
End Sub

'Callback for rmngBarColorGal getVisible
Sub actBarColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal0 getVisible
Sub actBarColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal1 getVisible
Sub actBarColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal2 getVisible
Sub actBarColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal3 getVisible
Sub actBarColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal4 getVisible
Sub actBarColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal5 getVisible
Sub actBarColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal6 getVisible
Sub actBarColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal7 getVisible
Sub actBarColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal8 getVisible
Sub actBarColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal9 getVisible
Sub actBarColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_actBarColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


'Callback for blBarColor getText
Sub blBarColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_blBarColor
    updateCustomDocumentProperty "cdpBlBarColor", xl_blBarColor, msoPropertyTypeNumber
End Sub

'Callback for blBarColorGal onAction
Sub blBarColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_blBarColor = index
    ribbonUI.InvalidateControl "blBarColor"
    ribbonUI.Invalidate
End Sub

'Callback for blBarColorGal getVisible
Sub blBarColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal0 getVisible
Sub blBarColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal1 getVisible
Sub blBarColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal2 getVisible
Sub blBarColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal3 getVisible
Sub blBarColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal4 getVisible
Sub blBarColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal5 getVisible
Sub blBarColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal6 getVisible
Sub blBarColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal7 getVisible
Sub blBarColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal8 getVisible
Sub blBarColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for blBarColorGal9 getVisible
Sub blBarColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_blBarColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColor getText
Sub prgBarColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_prgBarColor
    updateCustomDocumentProperty "cdpPrgBarColor", xl_prgBarColor, msoPropertyTypeNumber
End Sub

'Callback for prgBarColorGal onAction
Sub prgBarColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_prgBarColor = index
    ribbonUI.InvalidateControl "prgBarColor"
    ribbonUI.Invalidate
End Sub

'Callback for prgBarColorGal getVisible
Sub prgBarColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal0 getVisible
Sub prgBarColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal1 getVisible
Sub prgBarColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal2 getVisible
Sub prgBarColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal3 getVisible
Sub prgBarColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal4 getVisible
Sub prgBarColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal5 getVisible
Sub prgBarColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal6 getVisible
Sub prgBarColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal7 getVisible
Sub prgBarColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal8 getVisible
Sub prgBarColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for prgBarColorGal9 getVisible
Sub prgBarColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_prgBarColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


'Callback for fltBarColor getText
Sub fltBarColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_FltBarColor
    updateCustomDocumentProperty "cdpFltBarColor", xl_FltBarColor, msoPropertyTypeNumber
End Sub

'Callback for fltBarColorGal onAction
Sub fltBarColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_FltBarColor = index
    ribbonUI.InvalidateControl "fltBarColor"
    ribbonUI.Invalidate
End Sub

'Callback for fltBarColorGal getVisible
Sub fltBarColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal0 getVisible
Sub fltBarColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal1 getVisible
Sub fltBarColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal2 getVisible
Sub fltBarColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal3 getVisible
Sub fltBarColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal4 getVisible
Sub fltBarColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal5 getVisible
Sub fltBarColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal6 getVisible
Sub fltBarColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal7 getVisible
Sub fltBarColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal8 getVisible
Sub fltBarColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for fltBarColorGal9 getVisible
Sub fltBarColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_FltBarColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


'Callback for mileColor getText
Sub mileColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_mileColor
    updateCustomDocumentProperty "cdpMileColor", xl_mileColor, msoPropertyTypeNumber
End Sub

'Callback for rmngBarColorGal onAction
Sub mileColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_mileColor = index
    ribbonUI.InvalidateControl "mileColor"
    ribbonUI.Invalidate
End Sub

'Callback for rmngBarColorGal getVisible
Sub mileColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal0 getVisible
Sub mileColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal1 getVisible
Sub mileColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal2 getVisible
Sub mileColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal3 getVisible
Sub mileColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal4 getVisible
Sub mileColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal5 getVisible
Sub mileColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal6 getVisible
Sub mileColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal7 getVisible
Sub mileColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal8 getVisible
Sub mileColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for rmngBarColorGal9 getVisible
Sub mileColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_mileColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


'Callback for cutoffColor getText
Sub cutoffColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_cutoffColor
    updateCustomDocumentProperty "cdpCutoffColor", xl_cutoffColor, msoPropertyTypeNumber
End Sub

'Callback for cutoffColorGal onAction
Sub cutoffColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_cutoffColor = index
    ribbonUI.InvalidateControl "cutoffColor"
    ribbonUI.Invalidate
End Sub

'Callback for cutoffColorGal getVisible
Sub cutoffColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal0 getVisible
Sub cutoffColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal1 getVisible
Sub cutoffColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal2 getVisible
Sub cutoffColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal3 getVisible
Sub cutoffColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal4 getVisible
Sub cutoffColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal5 getVisible
Sub cutoffColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal6 getVisible
Sub cutoffColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal7 getVisible
Sub cutoffColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal8 getVisible
Sub cutoffColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for cutoffColorGal9 getVisible
Sub cutoffColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_cutoffColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColor getText
Sub windowColor_getIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = xl_windowColor
    updateCustomDocumentProperty "cdpwindowColor", xl_windowColor, msoPropertyTypeNumber
End Sub

'Callback for windowColorGal onAction
Sub windowColor_onAction(control As IRibbonControl, id As String, index As Integer)
    Call RefreshRibbon
    xl_windowColor = index
    ribbonUI.InvalidateControl "windowColor"
    ribbonUI.Invalidate
End Sub

'Callback for windowColorGal getVisible
Sub windowColorGal_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal0 getVisible
Sub windowColorGal0_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "0" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal1 getVisible
Sub windowColorGal1_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "1" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal2 getVisible
Sub windowColorGal2_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "2" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal3 getVisible
Sub windowColorGal3_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "3" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal4 getVisible
Sub windowColorGal4_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "4" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal5 getVisible
Sub windowColorGal5_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "5" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal6 getVisible
Sub windowColorGal6_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "6" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal7 getVisible
Sub windowColorGal7_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "7" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal8 getVisible
Sub windowColorGal8_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "8" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for windowColorGal9 getVisible
Sub windowColorGal9_getVisibile(control As IRibbonControl, ByRef returnedVal)
    If xl_windowColor = "9" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub


Sub checkDatePicker()
    ensureDPManager
    If g_oDP.PickerVisible Then closeDatePicker
End Sub

'DISABLE CONTROLS IN FREE AND PRO EDITIONS-----------------------------------------------

'Only disable in Free Edition
Sub free_getEnabled(control As IRibbonControl, ByRef returnedVal)
    SetEdition
    If intEdition = 1 Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

'Disable in Free and Pro Edition
Sub pro_getEnabled(control As IRibbonControl, ByRef returnedVal)
    SetEdition
    If intEdition >= 1 Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

'Only enable in Free Edition
Sub freeOnly_getEnabled(control As IRibbonControl, ByRef returnedVal)
    SetEdition
    If intEdition = 1 Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

