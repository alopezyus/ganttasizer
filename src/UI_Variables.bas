Attribute VB_Name = "UI_Variables"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

'Private myRibUI As IRibbonUI
Private intPeriod As Integer
Private intWeekStart As Integer
Private intPeriodWidth As Integer
Private intStartExtra As Integer
Private intFinishExtra As Integer
Private strCutoff As String
Private intBarStyle As Integer
Private intMilStyle As Integer
Private intShpHgt As Integer
Private booWBSoutline As Boolean
Private booLblDesc As Boolean
Private booLblFinish As Boolean
Private booLblDur As Boolean
Private booLblStart As Boolean
Private booLblActuals As Boolean
Private intRmgBarColor As String
Private intActBarColor As String
Private intBlBarColor As String
Private intMileColor As String
Private intCutoffColor As String
Private intWindowColor As String
Private intPrgBarColor As String
Private intFltBarColor As String
Private intRelType As Integer
Private intRelLag As Integer
Private intConStyle As Integer
Private intThick As Integer
'Private intTransp As Integer
Private booAutoChart As Boolean
Private booAutoUnits As Boolean
Private booAutoRow As Boolean
Private booTimeScl As Boolean
Private booSetActColor As Boolean
Private booAutoSch As Boolean
Private booBlBar As Boolean
Private booPrgBar As Boolean
Private booFltBar As Boolean
Private booSunday As Boolean
Private booMonday As Boolean
Private booTuesday As Boolean
Private booWednesday As Boolean
Private booThursday As Boolean
Private booFriday As Boolean
Private booSaturday As Boolean
Private intUnitsCurve As Integer

'PROPIEDADES-------------------------------------------------------------------------------------------
'Public Property Let ribbonUI(irib As IRibbonUI)
'    Set myRibUI = irib
'End Property
'Public Property Get ribbonUI() As IRibbonUI
'    Set ribbonUI = myRibUI
'End Property

Public Property Let xl_wbsOutline(strVal As Boolean)
    booWBSoutline = strVal
End Property
Public Property Get xl_wbsOutline() As Boolean
    xl_wbsOutline = booWBSoutline
End Property

Public Property Let xl_period(strVal As String)
    intPeriod = strVal
End Property
Public Property Get xl_period() As String
    xl_period = intPeriod
End Property

Public Property Let xl_weekStart(strVal As String)
    intWeekStart = strVal
End Property
Public Property Get xl_weekStart() As String
    xl_weekStart = intWeekStart
End Property

Public Property Let xl_periodWidth(strVal As String)
    intPeriodWidth = strVal
End Property
Public Property Get xl_periodWidth() As String
    xl_periodWidth = intPeriodWidth
End Property

Public Property Let xl_startExtra(strVal As String)
    intStartExtra = strVal
End Property
Public Property Get xl_startExtra() As String
    xl_startExtra = intStartExtra
End Property

Public Property Let xl_finishExtra(strVal As String)
    intFinishExtra = strVal
End Property
Public Property Get xl_finishExtra() As String
    xl_finishExtra = intFinishExtra
End Property

Public Property Let xl_cutoff(strVal As String)
    strCutoff = strVal
End Property
Public Property Get xl_cutoff() As String
    xl_cutoff = strCutoff
End Property

Public Property Let xl_barStyle(strVal As String)
    intBarStyle = strVal
End Property
Public Property Get xl_barStyle() As String
    xl_barStyle = intBarStyle
End Property

Public Property Let xl_milStyle(strVal As String)
    intMilStyle = strVal
End Property
Public Property Get xl_milStyle() As String
    xl_milStyle = intMilStyle
End Property

Public Property Let xl_shpHgt(strVal As String)
    intShpHgt = strVal
End Property
Public Property Get xl_shpHgt() As String
    xl_shpHgt = intShpHgt
End Property

Public Property Let xl_lblDesc(strVal As Boolean)
    booLblDesc = strVal
End Property
Public Property Get xl_lblDesc() As Boolean
    xl_lblDesc = booLblDesc
End Property

Public Property Let xl_lblFinish(strVal As Boolean)
    booLblFinish = strVal
End Property
Public Property Get xl_lblFinish() As Boolean
    xl_lblFinish = booLblFinish
End Property

Public Property Let xl_lblDur(strVal As Boolean)
    booLblDur = strVal
End Property
Public Property Get xl_lblDur() As Boolean
    xl_lblDur = booLblDur
End Property

Public Property Let xl_lblStart(strVal As Boolean)
    booLblStart = strVal
End Property
Public Property Get xl_lblStart() As Boolean
    xl_lblStart = booLblStart
End Property

Public Property Let xl_lblActuals(strVal As Boolean)
    booLblActuals = strVal
End Property

Public Property Get xl_lblActuals() As Boolean
    xl_lblActuals = booLblActuals
End Property

Public Property Let xl_relType(strVal As String)
    intRelType = strVal
End Property
Public Property Get xl_relType() As String
    xl_relType = intRelType
End Property

Public Property Let xl_relLag(strVal As String)
    intRelLag = strVal
End Property
Public Property Get xl_relLag() As String
    xl_relLag = intRelLag
End Property

Public Property Let xl_conStyle(strVal As String)
    intConStyle = strVal
End Property
Public Property Get xl_conStyle() As String
    xl_conStyle = intConStyle
End Property

Public Property Let xl_conThick(strVal As String)
    intThick = strVal
End Property
Public Property Get xl_conThick() As String
    xl_conThick = intThick
End Property

'Public Property Let xl_conTransp(strVal As String)
'    intTransp = strVal
'End Property
'Public Property Get xl_conTransp() As String
'    xl_conTransp = intTransp
'End Property

Public Property Let xl_UpdChart(strVal As Boolean)
    booAutoChart = strVal
End Property
Public Property Get xl_UpdChart() As Boolean
    xl_UpdChart = IIf(intEdition = 1, False, booAutoChart)  'Free Edition: xl_UpdSch = False
End Property

Public Property Let xl_UpdUnits(strVal As Boolean)
    booAutoUnits = strVal
End Property
Public Property Get xl_UpdUnits() As Boolean
    xl_UpdUnits = IIf(intEdition > 0, False, booAutoUnits) 'Free Edition and Pro Edition: xl_UpdSch = False
End Property

Public Property Let xl_UpdRow(strVal As Boolean)
    booAutoRow = strVal
End Property
Public Property Get xl_UpdRow() As Boolean
    xl_UpdRow = booAutoRow
End Property

Public Property Let xl_TimeScl(strVal As Boolean)
    booTimeScl = strVal
End Property
Public Property Get xl_TimeScl() As Boolean
    xl_TimeScl = IIf(intEdition = 1, False, booTimeScl) 'Free Edition: xl_BlBar = False
End Property

Public Property Let xl_SetActColor(strVal As Boolean)
    booSetActColor = strVal
End Property

Public Property Get xl_SetActColor() As Boolean
    xl_SetActColor = booSetActColor
End Property

Public Property Let xl_UpdSch(strVal As Boolean)
    booAutoSch = strVal
End Property

Public Property Get xl_UpdSch() As Boolean
    xl_UpdSch = IIf(intEdition > 0, False, booAutoSch) 'Free Edition and Pro Edition: xl_UpdSch = False
End Property

Public Property Let xl_BlBar(strVal As Boolean)
    booBlBar = strVal
End Property
Public Property Get xl_BlBar() As Boolean
    xl_BlBar = IIf(intEdition = 1, False, booBlBar) 'Free Edition: xl_BlBar = False
End Property

Public Property Let xl_PrgBar(strVal As Boolean)
    booPrgBar = strVal
End Property
Public Property Get xl_PrgBar() As Boolean
    xl_PrgBar = IIf(intEdition = 1, False, booPrgBar) 'Free Edition: xl_PrgBar = False
End Property

Public Property Let xl_FltBar(strVal As Boolean)
    booFltBar = strVal
End Property
Public Property Get xl_FltBar() As Boolean
    xl_FltBar = IIf(intEdition > 0, False, booFltBar) 'Free Edition and Pro Edition: xl_PrgBar = False
End Property

Public Property Let xl_sunday(strVal As Boolean)
    booSunday = strVal
End Property
Public Property Get xl_sunday() As Boolean
    xl_sunday = booSunday
End Property

Public Property Let xl_monday(strVal As Boolean)
    booMonday = strVal
End Property
Public Property Get xl_monday() As Boolean
    xl_monday = booMonday
End Property

Public Property Let xl_tuesday(strVal As Boolean)
    booTuesday = strVal
End Property
Public Property Get xl_tuesday() As Boolean
    xl_tuesday = booTuesday
End Property

Public Property Let xl_wednesday(strVal As Boolean)
    booWednesday = strVal
End Property
Public Property Get xl_wednesday() As Boolean
    xl_wednesday = booWednesday
End Property

Public Property Let xl_thursday(strVal As Boolean)
    booThursday = strVal
End Property
Public Property Get xl_thursday() As Boolean
    xl_thursday = booThursday
End Property
Public Property Let xl_friday(strVal As Boolean)
    booFriday = strVal
End Property
Public Property Get xl_friday() As Boolean
    xl_friday = booFriday
End Property

Public Property Let xl_saturday(strVal As Boolean)
    booSaturday = strVal
End Property
Public Property Get xl_saturday() As Boolean
    xl_saturday = booSaturday
End Property

Public Property Let xl_unitsCurve(strVal As String)
    intUnitsCurve = strVal
End Property
Public Property Get xl_unitsCurve() As String
    xl_unitsCurve = intUnitsCurve
End Property

Public Property Let xl_rmgBarColor(strVal As String)
    intRmgBarColor = strVal
End Property
Public Property Get xl_rmgBarColor() As String
    xl_rmgBarColor = intRmgBarColor
End Property

Public Property Let xl_actBarColor(strVal As String)
    intActBarColor = strVal
End Property
Public Property Get xl_actBarColor() As String
    xl_actBarColor = intActBarColor
End Property

Public Property Let xl_blBarColor(strVal As String)
    intBlBarColor = strVal
End Property
Public Property Get xl_blBarColor() As String
    xl_blBarColor = intBlBarColor
End Property

Public Property Let xl_prgBarColor(strVal As String)
    intPrgBarColor = strVal
End Property
Public Property Get xl_prgBarColor() As String
    xl_prgBarColor = intPrgBarColor
End Property

Public Property Let xl_FltBarColor(strVal As String)
    intFltBarColor = strVal
End Property
Public Property Get xl_FltBarColor() As String
    xl_FltBarColor = intFltBarColor
End Property

Public Property Let xl_mileColor(strVal As String)
    intMileColor = strVal
End Property
Public Property Get xl_mileColor() As String
    xl_mileColor = intMileColor
End Property

Public Property Let xl_cutoffColor(strVal As String)
    intCutoffColor = strVal
End Property
Public Property Get xl_cutoffColor() As String
    xl_cutoffColor = intCutoffColor
End Property

Public Property Let xl_windowColor(strVal As String)
    intWindowColor = strVal
End Property
Public Property Get xl_windowColor() As String
    xl_windowColor = intWindowColor
End Property



Public Function updateProjectVarProperty(strPropertyName As String, defaultValue As Variant) As Variant
On Error GoTo DefaultValues

updateProjectVarProperty = ActiveWorkbook.CustomDocumentProperties(strPropertyName).value
Exit Function

DefaultValues:
updateProjectVarProperty = defaultValue
End Function

Public Sub updateCustomDocumentProperty(strPropertyName As String, _
    varValue As Variant, docType As Office.MsoDocProperties)
     'VB TYPES:
        'vbEmpty                0       Empty (uninitialized)
        'vbNull                 1       Null (no valid data)
        'vbInteger              2       Integer
        'vbLong                 3       Long integer
        'vbSingle               4       Single-precision floating-point number
        'vbDouble               5       Double-precision floating-point number
        'vbCurrency             6       Currency value
        'vbDate                 7       Date value
        'vbString               8       String
        'vbObject               9       Object
        'vbError                10      Error value
        'vbBoolean              11      Boolean value
        'vbVariant              12      Variant (used only with arrays of variants)
        'vbDataObject           13      A data access object
        'vbDecimal              14      Decimal value
        'vbByte                 17      Byte value
        'vbUserDefinedType      36      Variants that contain user-defined types
        'vbArray                8192    Array

    'OFFICE.MSODOCPROPERTIES.TYPES
        'msoPropertyTypeNumber  1       Integer value.
        'msoPropertyTypeBoolean 2       Boolean value.
        'msoPropertyTypeDate    3       Date value.
        'msoPropertyTypeString  4       String value.
        'msoPropertyTypeFloat   5       Floating point value.

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties(strPropertyName).value = varValue
    If Err.Number > 0 Then
        ActiveWorkbook.CustomDocumentProperties.Add _
            Name:=strPropertyName, _
            LinkToContent:=False, _
            Type:=docType, _
            value:=varValue
    End If
End Sub

Public Sub deleteCustomDocumentProperty(strPropertyName As String)
    ActiveWorkbook.CustomDocumentProperties(strPropertyName).Delete
End Sub

