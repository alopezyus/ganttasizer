Attribute VB_Name = "Variables"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public wsSch As Worksheet
Public rngRef, rngActStyle, rngShpHgt, rngConStyle, rngLabPos, rngTmlMod, rngTmlCod, rngSchMod, rngDistCrv, rngActID, rngWBS, rngDesc, rngTotDur, rngRmgDur, _
        rngStart, rngFinish, rngStartBL, rngFinishBL, rngStartAct, rngFinishAct, rngResume, rngConstraint, rngPred, rngFloat, rngPeriod, rngProgress, rngBdgUnt, rngRmgUnt As Range
Public lngRemBarColor, lngActBarColor, lngBLBarColor, lngPrgBarColor, lngMilColor, lngCutoffColor, lngWindowColor, lngFltBarColor As Long
Public datCutoff, datPrjFirst, datPrjLast, datPrjStart, datPrjFinish, datChartStart As Date
Public intPeriod, intExtraPeriodS, intExtraPeriodsF As Integer
Public dblWidth, dblHeightStd As Double
Public dblBarHgt, dblBarPos, dblConTrn, dblConThk As Double
Public intBarStyle, intMilStyle, intActLastRow, intSumHgt, intConStyle, intDistCrv, intWeekStart, intCalFirstRowOffset, intWorkingDays, intFirstCol, intLastCol As Integer
Public strHeadArr() As Variant
Public strRelType As String
Public intRelLag As Integer
Public intEdition As Integer
Public booSun, booMon, booTue, booWed, booThu, booFri, booSat As Boolean
Public booLabDesc, booLabFinish, booLabDur, booLabStart, booLabActuals As Boolean
Public booGroupWBS As Boolean
Public fDPribbon, fDPstartexc, fDPfinishexc As Boolean
Public booHeaders As Boolean 'Si el procedimiento SetHeaderRef devuelve error está variable se pone False y permite salir de las llamadas aguas arriba
Public booPrjStartSet As Boolean 'Variables creadas para acelerar la actualizacion de arrastrar y soltar en barras
Public arrDatesExc As Variant

Public Sub SetEdition()
    'Ganttasizer edition
        '0 --> Master Edition
        '1 --> Free Edition
        '2 --> Pro Edition
    'Debug.Print intEdition
    intEdition = 0
End Sub


'Variables de uso para todo el proyecto
Public Sub SetPrjVar()
    SetEdition
    
    'Hojas de trabajo
    Set wsSch = ActiveSheet
        
    'Set default colors
    lngRemBarColor = rgbColor(xl_rmgBarColor)
    lngActBarColor = rgbColor(xl_actBarColor)
    lngBLBarColor = rgbColor(xl_blBarColor)
    lngPrgBarColor = rgbColor(xl_prgBarColor)
    lngFltBarColor = rgbColor(xl_FltBarColor)
    lngMilColor = rgbColor(xl_mileColor)
    lngCutoffColor = rgbColor(xl_cutoffColor)
    lngWindowColor = rgbColor(xl_windowColor)
    
    'Set time scale variables
    intPeriod = xl_period + 1
    'Enable in Free Edition only
    If intEdition = 1 Then
        Select Case intPeriod
        Case 1 'semanas
            intPeriod = 3
        Case 2 'meses
            intPeriod = 5
        Case 3 'años
            intPeriod = 7
        End Select
    End If
    intExtraPeriodS = xl_startExtra
    intExtraPeriodsF = xl_finishExtra
    intWeekStart = xl_weekStart + 1
    dblWidth = periodWidth(xl_periodWidth + 1)
    dblHeightStd = 15
    
    'Set bars and milestones shapes
    intBarStyle = xl_barStyle + 1
    intMilStyle = xl_milStyle + 1 + 10
    'Set bar height
    dblBarHgt = shapeHeight(xl_shpHgt + 1)
    intSumHgt = 1

    'Set Relationships options
    strRelType = RelType(xl_relType + 1)
    intRelLag = xl_relLag
    
    'Set connectors options
    dblConThk = conThick(xl_conThick + 1)
    dblConTrn = 0 'conTransp(xl_conTransp + 1)
    intConStyle = xl_conStyle + 1
    
    'Set distribution curve
    intDistCrv = IIf(intEdition > 0, Empty, xl_unitsCurve + 1) 'Disabled in Free Edition and Pro Edition
    
    'Set Cutoff date
    datCutoff = IIf(intEdition = 1, Empty, SetCutoff) 'Free Edition: datCutoff = Empty

    'Set label options
    booLabDesc = xl_lblDesc
    booLabFinish = IIf(intEdition = 1, False, xl_lblFinish) 'Free Edition: booLabFinish = False
    booLabDur = IIf(intEdition = 1, False, xl_lblDur) 'Free Edition: booLabDur = False
    booLabStart = IIf(intEdition = 1, False, xl_lblStart) 'Free Edition: booLabStart = False
    booLabActuals = IIf(intEdition = 1, False, xl_lblActuals) 'Free Edition: booLabActuals = False
    
    booGroupWBS = xl_wbsOutline
    
    'Set headers references. The rest of variables require headers to be calculated
    SetHeaderRef
    If Not booHeaders Then Exit Sub
    
    'Ultima fila de la lista de actividades y utlima columna del calendario
    intActLastRow = ActLastRow
    intLastCol = CalLastColumn
    intFirstCol = TblFirstColumn
    
    SetPrjDates 'Calculo datPrjStart, datPrjFinish
    datChartStart = datPrjStart
    
    arrDatesExc = IIf(intEdition > 0, Empty, GetCalendarExceptions) 'Free Edition and Pro Edition: arrDateExc = Empty
    
    'Cálculo del posicionamiento de la barra: Calculado para cada actividad
'    dblBarPos = (1 - dblBarHgt * IIf(WorksheetFunction.Count(wsSch.Range(rngStartBL.Offset(1, 0), rngStartBL.Offset(intActLastRow - rngRef.Row, 0))) + _
'                WorksheetFunction.Count(wsSch.Range(rngFinishBL.Offset(1, 0), rngFinishBL.Offset(intActLastRow - rngRef.Row, 0))) > 0, 1.5, 1)) / 2

    'Establecer primera fila del calendario
    If rngRef.row >= 3 Then
        If rngRef.Offset(0, 0).value <= 31 Or Len(rngRef.Offset(0, 0).value) <= 2 Then
            intCalFirstRowOffset = 2
        ElseIf Len(rngRef.Offset(0, 0).value) = 3 Or rngRef.Offset(0, 0).value Like "Q*" Then
            intCalFirstRowOffset = 1
        Else: intCalFirstRowOffset = 0
        End If
    ElseIf rngRef.row = 2 Then
        If Len(rngRef.Offset(0, 0).value) = 3 Or rngRef.Offset(0, 0).value Like "Q*" Then
            intCalFirstRowOffset = 1
        Else: intCalFirstRowOffset = 0
        End If
    Else: intCalFirstRowOffset = 0
    End If
End Sub

Public Sub GetHeaderArray()
    strHeadArr = Array("act / mil style", "shape height", "connect style", "label pos", "timeline mode", "timeline code", "schedule mode", "units distrib. curve", "ACTIVITY ID", "WBS", _
                        "DESCRIPTION", "TOTAL DURATION", "REMAINING DURATION", "START DATE", "FINISH DATE", "BL START DATE", "BL FINISH DATE", "ACTUAL START DATE", "ACTUAL FINISH DATE", _
                        "RESUME DATE", "CONSTRAINT DATE", "PREDECESSORS", "FLOAT", "PROGRESS %", "BUDGET UNITS", "REMAINING UNITS", "Period")
End Sub

Public Sub SetHeaderRef(Optional booWarning As Boolean = True)
On Error GoTo errHandler
    Dim i As Integer
    Dim varReturn As Variant
    
    booHeaders = True
    
    varReturn = returnName
    
    If Not varReturn(0) Then
        GoTo errHandler
    End If
    
    Set rngActStyle = Range("VB_" & varReturn(1) & "_00")
    Set rngShpHgt = Range("VB_" & varReturn(1) & "_01")
    Set rngConStyle = Range("VB_" & varReturn(1) & "_02")
    Set rngLabPos = Range("VB_" & varReturn(1) & "_03")
    Set rngTmlMod = Range("VB_" & varReturn(1) & "_04")
    Set rngTmlCod = Range("VB_" & varReturn(1) & "_05")
    Set rngSchMod = Range("VB_" & varReturn(1) & "_06")
    Set rngDistCrv = Range("VB_" & varReturn(1) & "_07")
    Set rngActID = Range("VB_" & varReturn(1) & "_08")
    Set rngWBS = Range("VB_" & varReturn(1) & "_09")
    Set rngDesc = Range("VB_" & varReturn(1) & "_10")
    Set rngTotDur = Range("VB_" & varReturn(1) & "_11")
    Set rngRmgDur = Range("VB_" & varReturn(1) & "_12")
    Set rngStart = Range("VB_" & varReturn(1) & "_13")
    Set rngFinish = Range("VB_" & varReturn(1) & "_14")
    Set rngStartBL = Range("VB_" & varReturn(1) & "_15")
    Set rngFinishBL = Range("VB_" & varReturn(1) & "_16")
    Set rngStartAct = Range("VB_" & varReturn(1) & "_17")
    Set rngFinishAct = Range("VB_" & varReturn(1) & "_18")
    Set rngResume = Range("VB_" & varReturn(1) & "_19")
    Set rngConstraint = Range("VB_" & varReturn(1) & "_20")
    Set rngPred = Range("VB_" & varReturn(1) & "_21")
    Set rngFloat = Range("VB_" & varReturn(1) & "_22")
    Set rngProgress = Range("VB_" & varReturn(1) & "_23")
    Set rngBdgUnt = Range("VB_" & varReturn(1) & "_24")
    Set rngRmgUnt = Range("VB_" & varReturn(1) & "_25")
    Set rngPeriod = Range("VB_" & varReturn(1) & "_26")
    'Celda de referencia para calendario
    Set rngRef = rngPeriod.Offset(0, 1)
    Exit Sub
    
errHandler:
        If booWarning Then CustomMsgBox "Headers are not correctly defined in this worksheet.", error:=True
        booHeaders = False
End Sub

Public Sub SetPrjDates()
    Dim rngAllDates As Range
    Dim i, j As Integer
    Dim arrVal, arrData, arrAux As Variant
    Dim arrAct As Variant
    Dim intDim, intFields As Integer
    
    'Se establece la dimensión del vector: número de filas
    intDim = intActLastRow - rngRef.row
    If intDim = 0 Then Exit Sub
    'Se establece el número de campos
    intFields = 4
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngStart.Offset(1), rngStart.Offset(intDim)).value, _
                    Range(rngFinish.Offset(1), rngFinish.Offset(intDim)).value, _
                    Range(rngStartBL.Offset(1), rngStartBL.Offset(intDim)).value, _
                    Range(rngFinishBL.Offset(1), rngFinishBL.Offset(intDim)).value)
    'Se dimensiona el vector que contendrá la información final al número de filas y un vector auxiliar al número de campos
    ReDim arrData(intDim - 1)
    ReDim arrAux(intFields)
    
    'Para cada fila
    For i = 0 To intDim - 1
        'Para cada campo
        For j = 0 To intFields
            'Se actualiza el vector auxiliar de campos con los valores en el vector de inicio
            If intDim = 1 Then
                arrAux(j) = arrVal(j)
            Else
                arrAux(j) = arrVal(j)(i + 1, 1)
            End If
        Next
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux
    Next

    datPrjFirst = Empty
    datPrjLast = Empty
    
    i = 0
    For Each arrAct In arrData
        If Not arrData(i)(0) Like "WBS-*" Then
            If IsError(arrAct(1)) Then arrAct(1) = Empty
            If Not IsEmpty(arrAct(1)) And Not arrAct(1) = "" Then
                arrAct(1) = CDate(arrAct(1))
                If datPrjFirst = Empty Or datPrjFirst > arrAct(1) Then datPrjFirst = arrAct(1)
                If datPrjLast = Empty Or datPrjLast < arrAct(1) Then datPrjLast = arrAct(1)
            End If
            If IsError(arrAct(2)) Then arrAct(2) = Empty
            If Not IsEmpty(arrAct(2)) And Not arrAct(2) = "" Then
                arrAct(2) = CDate(arrAct(2))
                If datPrjFirst = Empty Or datPrjFirst > arrAct(2) Then datPrjFirst = arrAct(2)
                If datPrjLast = Empty Or datPrjLast < arrAct(2) Then datPrjLast = arrAct(2)
            End If
            If xl_BlBar Then
                If IsError(arrAct(3)) Then arrAct(3) = Empty
                If Not IsEmpty(arrAct(3)) And Not arrAct(3) = "" Then
                    arrAct(3) = CDate(arrAct(3))
                    If datPrjFirst = Empty Or datPrjFirst > arrAct(3) Then datPrjFirst = arrAct(3)
                    If datPrjLast = Empty Or datPrjLast < arrAct(3) Then datPrjLast = arrAct(3)
                End If
                If IsError(arrAct(4)) Then arrAct(4) = Empty
                If Not IsEmpty(arrAct(4)) And Not arrAct(4) = "" Then
                    arrAct(4) = CDate(arrAct(4))
                    If datPrjFirst = Empty Or datPrjFirst > arrAct(4) Then datPrjFirst = arrAct(4)
                    If datPrjLast = Empty Or datPrjLast < arrAct(4) Then datPrjLast = arrAct(4)
                End If
            End If
        End If
        i = i + 1
    Next
    If IsDate(datCutoff) And datCutoff > 0 And datCutoff < datPrjFirst Then datPrjFirst = datCutoff
    
    datPrjStart = IIf(intPeriod <= 2, datPrjFirst - intExtraPeriodS * 1, _
                IIf(intPeriod = 3, datPrjFirst - intExtraPeriodS * 7 - (Weekday(datPrjFirst, intWeekStart) - 1), _
                IIf(intPeriod = 4, datPrjFirst - intExtraPeriodS * 14 - (Weekday(datPrjFirst, intWeekStart) - 1), _
                IIf(intPeriod = 5, DateSerial(Year(datPrjFirst), Month(datPrjFirst) - intExtraPeriodS, 1), _
                IIf(intPeriod = 6, DateSerial(Year(datPrjFirst), 3 * ((Month(datPrjFirst) - 1) \ 3) + 1 - 3 * intExtraPeriodS, 1), _
                    DateSerial(Year(datPrjFirst) - intExtraPeriodS, 1, 1))))))
    datPrjFinish = IIf(intPeriod <= 2, datPrjLast + intExtraPeriodsF * 1, _
                IIf(intPeriod = 3, datPrjLast + intExtraPeriodsF * 7 + (7 - (Weekday(datPrjLast, intWeekStart) - 1) - 1), _
                IIf(intPeriod = 4, datPrjLast + intExtraPeriodsF * 14 + (7 - (Weekday(datPrjLast, intWeekStart) - 1) - 1), _
                IIf(intPeriod = 5, DateSerial(Year(datPrjLast), Month(datPrjLast) + intExtraPeriodsF + 1, 1) - 1, _
                IIf(intPeriod = 6, DateSerial(Year(datPrjLast), 3 * ((Month(datPrjLast) - 1) \ 3) + 4 + 3 * intExtraPeriodsF, 1) - 1, _
                    DateSerial(Year(datPrjLast) + intExtraPeriodsF, 12, 31))))))

End Sub

Public Function SetCutoff()
    'Set Cutoff date
    If xl_cutoff = "" Then
        SetCutoff = 0
    Else
        SetCutoff = CDate(xl_cutoff)
    End If
End Function

Private Sub SetDateFormat(rng As Range)
    If IsDate(rng) = False Then
        rng = format(rng, "dd/mm/yy")
        rng.HorizontalAlignment = xlLeft
    End If
End Sub

'Free Edition: All booDay = True
Public Sub WeekCalendar()
    booSun = IIf(intEdition = 1, True, xl_sunday)
    booMon = IIf(intEdition = 1, True, xl_monday)
    booTue = IIf(intEdition = 1, True, xl_tuesday)
    booWed = IIf(intEdition = 1, True, xl_wednesday)
    booThu = IIf(intEdition = 1, True, xl_thursday)
    booFri = IIf(intEdition = 1, True, xl_friday)
    booSat = IIf(intEdition = 1, True, xl_saturday)
    
    intWorkingDays = IIf(booSun, 1, 0) + IIf(booMon, 1, 0) + IIf(booTue, 1, 0) + IIf(booWed, 1, 0) + IIf(booThu, 1, 0) + IIf(booFri, 1, 0) + IIf(booSat, 1, 0)
End Sub

Public Function ActLastRow(Optional ByRef intStartRow As Integer = 0) As Integer
    Dim booLoop As Boolean
    Dim rngCurrAct, rngCurrDesc As Range
    Dim strType As String

    booLoop = True
    If Not intStartRow = 0 Then intStartRow = intStartRow - rngRef.row
    Set rngCurrAct = rngActID.Offset(intStartRow)
    Set rngCurrDesc = rngDesc.Offset(intStartRow)
        
    Do While booLoop
        If IsError(rngCurrAct) Or IsError(rngCurrDesc) Or IsError(rngCurrAct.Offset(1, 0)) Or IsError(rngCurrDesc.Offset(1, 0)) Then
            Set rngCurrAct = rngCurrAct.Offset(1, 0)
            Set rngCurrDesc = rngCurrDesc.Offset(1, 0)
        ElseIf Not rngCurrAct = Empty Or Not rngCurrDesc = Empty _
             Or Not rngCurrAct.Offset(1, 0) = Empty Or Not rngCurrDesc.Offset(1, 0) = Empty Then
            Set rngCurrAct = rngCurrAct.Offset(1, 0)
            Set rngCurrDesc = rngCurrDesc.Offset(1, 0)
        Else:
            Set rngCurrAct = rngCurrAct.Offset(-1, 0)
            Set rngCurrDesc = rngCurrDesc.Offset(-1, 0)
            booLoop = False
        End If
        
        If Not (IsError(rngCurrAct) Or IsError(rngCurrDesc)) Then
            If (rngCurrAct.value = rngActID.value Or rngCurrDesc.value = rngDesc.value) And rngCurrAct.row > rngRef.row And rngCurrAct.row > 1 Then
                Set rngCurrAct = rngCurrAct.Offset(-1, 0)
                Set rngCurrDesc = rngCurrDesc.Offset(-1, 0)
                booLoop = False
            End If
        End If
    Loop
    
    ActLastRow = IIf(rngCurrAct.row > rngCurrDesc.row, rngCurrAct.row, rngCurrDesc.row)
    'Enable in Free Edition only
    If intEdition = 1 Then
        If ActLastRow - rngRef.row > 30 Then ActLastRow = 30 + rngRef.row
    End If

End Function

Public Function TblFirstColumn() As Integer

    TblFirstColumn = WorksheetFunction.Min(rngRef.Column, rngActStyle.Column, rngShpHgt.Column, rngConStyle.Column, rngLabPos.Column, rngTmlMod.Column, rngTmlCod.Column, _
                                            rngSchMod.Column, rngDistCrv.Column, rngActID.Column, rngWBS.Column, rngDesc.Column, rngTotDur.Column, rngRmgDur.Column, _
                                            rngStart.Column, rngFinish.Column, rngStartBL.Column, rngFinishBL.Column, rngStartAct.Column, rngFinishAct.Column, rngResume.Column, _
                                            rngPred.Column, rngFloat.Column, rngPeriod.Column, rngProgress.Column, rngBdgUnt.Column, rngRmgUnt.Column)

End Function

Public Function CalLastColumn() As Integer
    Dim booLoop As Boolean
    Dim rngCurr As Range
    Dim strType As String
    
    booLoop = True
    Set rngCurr = rngRef
    
    strType = CellType(rngCurr)
    
    Do While booLoop
        If Not rngCurr = Empty And CellType(rngCurr) = strType Then
            Set rngCurr = rngCurr.Offset(0, 1)
        Else:
            Set rngCurr = rngCurr.Offset(0, -1)
            booLoop = False
        End If
    Loop

    CalLastColumn = rngCurr.Column

End Function

'Disabled in Free Edition and Pro Edition
Public Function GetCalendarExceptions() As Variant
    If intEdition > 0 Then Exit Function
    
    Dim strDatesExc, datDatesExc As Variant
    Dim varException As Variant
    Dim i As Integer
    
    On Error GoTo errHandler
    
    strDatesExc = Split(ActiveWorkbook.CustomDocumentProperties("cdpCalExc").value, ",")
    ReDim datDatesExc(0)
    
    For Each varException In strDatesExc
        varException = Trim(varException)
        If Len(varException) = 8 Then
            varException = CDate(Right(varException, 2) & "-" & Mid(varException, 5, 2) & "-" & Left(varException, 4))
            If WorkingWeekDay(varException) Then
                If Not IsEmpty(datDatesExc(0)) Then ReDim Preserve datDatesExc(UBound(datDatesExc) + 1)
                datDatesExc(UBound(datDatesExc)) = varException
            End If
        End If
    Next
    GetCalendarExceptions = datDatesExc
    Exit Function
    
errHandler:
    updateCustomDocumentProperty "cdpCalExc", "", msoPropertyTypeString
    ReDim datDatesExc(0)
End Function


Public Function periodWidth(level As Integer) As Double
    Select Case level
    Case 1 '12 pixels
        periodWidth = 1
    Case 2 '24 pixels
        periodWidth = 2.71
    Case 3 '36 pixels
        periodWidth = 4.43
    Case 4 '48 pixels
        periodWidth = 6.14
    Case 5 '60 pixels
        periodWidth = 7.86
    Case 6 '72 pixels
        periodWidth = 9.57
    Case 7 '84 pixels
        periodWidth = 11.29
    Case 8 '96 pixels
        periodWidth = 13
    Case 9 '108 pixels
        periodWidth = 14.71
    Case 10 '120 pixels
        periodWidth = 16.43
    End Select
End Function

Public Function shapeHeight(level As Integer) As Double
    Select Case level
    Case 1
        shapeHeight = 0.2
    Case 2
        shapeHeight = 0.3
    Case 3
        shapeHeight = 0.4
    Case 4
        shapeHeight = 0.5
    Case 5
        shapeHeight = 0.6
    Case 6
        shapeHeight = 0.65
    Case 7
        shapeHeight = 0.7
    Case 8
        shapeHeight = 0.8
    Case 9
        shapeHeight = 0.9
    Case 10
        shapeHeight = 1
    End Select
End Function

Public Function RelType(level As Integer) As String
    Select Case level
    Case 1
        RelType = "FS"
    Case 2
        RelType = "FF"
    Case 3
        RelType = "SS"
    Case 4
        RelType = "SF"
    End Select
End Function

Public Function conThick(level As Integer) As Double
    Select Case level
    Case 1
        conThick = 0.5
    Case 2
        conThick = 0.75
    Case 3
        conThick = 1
    Case 4
        conThick = 1.25
    Case 5
        conThick = 1.5
    Case 6
        conThick = 1.75
    Case 7
        conThick = 2
    Case 8
        conThick = 2.25
    Case 9
        conThick = 2.5
    Case 10
        conThick = 2.75
    Case 11
        conThick = 3
    End Select
End Function

Public Function conTransp(level As Integer) As Double
    Select Case level
    Case 1
        conTransp = 0
    Case 2
        conTransp = 0.1
    Case 3
        conTransp = 0.2
    Case 4
        conTransp = 0.3
    Case 5
        conTransp = 0.4
    Case 6
        conTransp = 0.5
    Case 7
        conTransp = 0.6
    Case 8
        conTransp = 0.7
    Case 9
        conTransp = 0.8
    Case 10
        conTransp = 0.9
    Case 11
        conTransp = 1
    End Select
End Function

Public Function rgbColor(code As Integer) As Long
    Select Case code
    Case 0 'Rojo
        rgbColor = RGB(255, 0, 0)
    Case 1 'Naranja
        rgbColor = RGB(255, 192, 0)
    Case 2 'Amarillo
        rgbColor = RGB(255, 255, 0)
    Case 3 'Verde Claro
        rgbColor = RGB(146, 208, 80)
    Case 4 'Verde Oscuro
        rgbColor = RGB(0, 176, 80)
    Case 5 'Azul Claro
        rgbColor = RGB(0, 176, 240)
    Case 6 'Azul Oscuro
        rgbColor = RGB(0, 112, 192)
    Case 7 'Morado
        rgbColor = RGB(112, 48, 160)
    Case 8 'Gris
        rgbColor = RGB(208, 206, 206)
    Case 9 'Negro
        rgbColor = RGB(0, 0, 0)
    End Select
End Function

Function CellType(rng)
    Application.Volatile
    Set rng = rng.Range("A1")
    Select Case True
        Case IsEmpty(rng)
            CellType = "Blank"
        Case WorksheetFunction.IsText(rng)
            CellType = "Text"
        Case WorksheetFunction.IsLogical(rng)
            CellType = "Logical"
        Case WorksheetFunction.IsErr(rng)
            CellType = "Error"
        Case IsDate(rng)
            CellType = "Date"
        Case InStr(1, rng.text, ":") <> 0
            CellType = "Time"
        Case IsNumeric(rng)
            CellType = "Value"
    End Select
End Function

