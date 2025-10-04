Attribute VB_Name = "Resources"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition and Pro Edition
Option Explicit

Dim arrData() As Variant
Dim arrAct As Variant

'Construcción del vector de actividades
Private Sub SetActArray(ByVal intRowUpdate As Variant)
    If intEdition > 0 Then Exit Sub
    
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux() As Variant
    Dim i, j As Integer
    Dim row As Variant
    
    'Se establece la dimensión del vector: número de filas
    intDim = ActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 9
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value, _
                    Range(rngTmlCod.Offset(1), rngTmlCod.Offset(intDim)).value, _
                    Range(rngDistCrv.Offset(1), rngDistCrv.Offset(intDim)).value, _
                    Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngWBS.Offset(1), rngWBS.Offset(intDim)).value, _
                    Range(rngStart.Offset(1), rngStart.Offset(intDim)).value, _
                    Range(rngFinish.Offset(1), rngFinish.Offset(intDim)).value, _
                    Range(rngResume.Offset(1), rngResume.Offset(intDim)).value, _
                    Range(rngRmgUnt.Offset(1), rngRmgUnt.Offset(intDim)).value, _
                    Range(rngRmgDur.Offset(1), rngRmgDur.Offset(intDim)).value)

    'Se dimensiona el vector que contendrá la información final al número de filas y un vector auxiliar al número de campos
    ReDim arrData(intDim - 1)
    ReDim arrAux(intFields + 1)
    
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
        'Inicialización del campo de Actividad Actualizable
        arrAux(10) = IIf(IsEmpty(intRowUpdate), True, False)
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux
    Next
    
    If Not IsEmpty(intRowUpdate) Then
        For Each row In intRowUpdate
            arrData(row - rngRef.row - 1)(10) = True
        Next
    End If
    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Timeline Mode
    '1 --> Timeline Code
    '2 --> Distribution Curve
    '3 --> Activity ID
    '4 --> WBS
    '5 --> Start
    '6 --> Finish
    '7 --> Resume Date
    '8 --> Remaining Unit
    '9 --> Remaining Duration
    'Dimensiones extra
    '10 --> Actividad actualizable segun vector intRowUpd
End Sub

Public Sub DistributeUnits(Optional ByVal intRowUpdate As Variant = Empty)
    If intEdition > 0 Then Exit Sub
    
    Dim intR, intRsum, k, l, i As Integer
    Dim intCurve As Integer
    Dim datActStart, datActFinish, datActResume, datPeriodStart, datPeriodFinish, datRangeStart, datRangeFinish, datCalFinish As Date
    Dim dblLinearK, dblMeanS, dblMeanF, dblMeanB, dblVarS, dblVarF, dblVarB, dblErrorS, dblErrorF, dblErrorB As Double
    Dim dblDistPeriod As Double
    Dim dblUnits As Variant
    Dim intPeriodIni, intPeriodFin, intTmlRow, intWBSRow, intCalLastCol As Integer
    Dim rngInsertUnits As Range
    Dim strTmlMod, strTmlCod, strActID, strWBS As String
    Dim booOneDayDur, booMilestone As Boolean
    Dim intRowUpdVal As Variant

    intCalLastCol = CalLastColumn
    
    'Recorrer lista de actividades
    SetActArray intRowUpdate
    intR = rngRef.row + 1
    i = 0
    
    'Primer recorrido para calcular distribución de actividades
    'También recorre timelines y sumarios para combinar celdas donde sea necesario
    For Each arrAct In arrData
        booMilestone = False
        'Comprobar si la actividad actual está en la lista de actividades a actualizar.
        If Not arrAct(10) Then GoTo NextIteration1

        'Eliminar contenido en la fila y descombinar celdas
        With Range(rngRef.Offset(intR - rngRef.row, 0), rngRef.Offset(intR - rngRef.row, intCalLastCol))
            .ClearContents
            .UnMerge
        End With
        'Si la fecha de terminación es anterior al corte no pueden quedar unidades remanentes
        If arrAct(6) <= datCutoff And arrAct(8) <> 0 Then
            rngRmgUnt.Offset(intR - rngRef.row, 0) = 0
            arrAct(8) = 0
            arrData(i)(8) = 0
        End If
        'Si no hay horas remanentes para distribuir y no es una WBS ni un Timeline se salta a la siguiente iteración
        If arrAct(8) = 0 And Not (arrAct(3) Like "WBS-*" Or Not IsEmpty(arrAct(0))) Then GoTo NextIteration1

        dblUnits = arrAct(8)
        If IsNumeric(dblUnits) And dblUnits > 0 Then
            'Establecemos variables de entrada de la actividad: fechas y periodo
            datActStart = arrAct(5)
            datActFinish = arrAct(6)
            datActResume = arrAct(7)
            datActStart = IIf(IsEmpty(datActStart), datActStart, IIf(datActStart >= datCutoff, datActStart, _
                        IIf(IsDate(datActResume) And datActResume > datCutoff And datActResume < datActFinish, datActResume, datCutoff + 1)))
            dblDistPeriod = 1 + DateDiffCal(datActStart, datActFinish) 'rngTotDur.Offset(intR- rngRef.Row, 0)
            'Si la actividad solo tiene un día de duracion no se puede calcular la curva normal
            If arrAct(9) = 0 Then
                'GoTo NextIteration1
                booMilestone = True
                GoTo insertUnits
            ElseIf arrAct(9) = 1 Then
                booOneDayDur = True
                GoTo insertUnits
            End If
            'Tipo de Curva de distribución
            intCurve = IIf(IsNumeric(Left(arrAct(2), 1)), Left(arrAct(2), 1), intDistCrv)
            'Parámetros para distribución Lineal
            dblLinearK = dblUnits / dblDistPeriod
            'Parámetros para distribución S-Curve
            dblMeanS = Round(dblDistPeriod / 2, 0)
            dblVarS = dblMeanS / 3
            dblErrorS = (1 - WorksheetFunction.Norm_Dist(dblDistPeriod, dblMeanS, dblVarS, True)) + WorksheetFunction.Norm_Dist(0, dblMeanS, dblVarS, True)
            dblErrorS = dblErrorS / dblDistPeriod
            'Parámetros para distribución Front Loaded
            dblMeanF = 0
            dblVarF = Round(dblDistPeriod / 2, 0) * 2 / 3
            dblErrorF = 2 * ((1 - WorksheetFunction.Norm_Dist(dblDistPeriod, dblMeanF, dblVarF, True)) + (WorksheetFunction.Norm_Dist(0, dblMeanF, dblVarF, True) - 0.5))
            dblErrorF = dblErrorF / dblDistPeriod
            'Parámetros para distribución Back Loaded
            dblMeanB = dblDistPeriod
            dblVarB = Round(dblDistPeriod / 2, 0) * 2 / 3
            dblErrorB = 2 * ((0.5 - WorksheetFunction.Norm_Dist(dblDistPeriod, dblMeanB, dblVarB, True)) + WorksheetFunction.Norm_Dist(0, dblMeanB, dblVarB, True))
            dblErrorB = dblErrorB / dblDistPeriod
            
insertUnits:
            'Posicionar primer y último día del periodo y el final del calendario
            Select Case intPeriod
                Case 1, 2 'Diario
                    datPeriodStart = datPrjStart
                    datPeriodFinish = datPeriodStart
                    datCalFinish = datPrjFinish
                Case 3 'Semanal
                    datPeriodStart = datPrjStart
                    datPeriodFinish = datPeriodStart + 7 - 1
                    datCalFinish = datPrjFinish - (Weekday(datPrjFinish, intWeekStart) + 1) + 7
                Case 4 'Bi-semanal
                    datPeriodStart = datPrjStart
                    datPeriodFinish = datPeriodStart + 14 - 1
                    datCalFinish = datPrjFinish - (Weekday(datPrjFinish, intWeekStart) + 1) + 14
                Case 5 'Mensual
                    datPeriodStart = DateSerial(Year(datPrjStart), Month(datPrjStart), 1)
                    datPeriodFinish = DateSerial(Year(datPeriodStart), Month(datPeriodStart) + 1, 1) - 1
                    datCalFinish = DateSerial(Year(datPrjFinish), Month(datPrjFinish) + 1, 1) - 1
                Case 6 'Trimestre
                    datPeriodStart = DateSerial(Year(datPrjStart), 3 * ((Month(datPrjStart) - 1) \ 3) + 1, 1)
                    datPeriodFinish = DateSerial(Year(datPeriodStart), 3 * ((Month(datPeriodStart) - 1) \ 3) + 4, 1) - 1
                    datCalFinish = DateSerial(Year(datPrjFinish), 3 * ((Month(datPrjFinish) - 1) \ 3) + 4, 1) - 1
                Case 7 'Anual
                    datPeriodStart = DateSerial(Year(datPrjStart), 1, 1)
                    datPeriodFinish = DateSerial(Year(datPeriodStart) + 1, 1, 1)
                    datCalFinish = DateSerial(Year(datPrjFinish), 12, 31)
            End Select
            
            k = 0
            intPeriodIni = 0
            intPeriodFin = 0
            Do While datPeriodStart <= datCalFinish
                Set rngInsertUnits = rngRef.Offset(intR - rngRef.row, k)
                If datPeriodStart <= datActFinish And datPeriodFinish >= datActStart Then
                    'Rango de fechas para que se va a calcular la distribución
                    datRangeStart = IIf(datPeriodStart > datActStart, datPeriodStart, datActStart)
                    datRangeFinish = IIf(datPeriodFinish < datActFinish, datPeriodFinish, datActFinish)
'                        datRangeFinish = IIf(IIf(datPeriodFinish + 1 = datActStart, datActStart, datPeriodFinish) < datActFinish, _
'                                            IIf(datPeriodFinish + 1 = datActStart, datActStart, datPeriodFinish), datActFinish)
                    'Periodo inicial y final de la distribución acumulada
                    intPeriodIni = intPeriodFin 'DateDiffCal(datActStart, datRangeStart)
                    intPeriodFin = DateDiffCal(datActStart, datRangeFinish) + IIf(intPeriod <= 2, DateDiffCal(datRangeStart - 1, datRangeFinish), 1)
                    intPeriodFin = IIf(intPeriodFin > intPeriodIni, intPeriodFin, intPeriodIni)
                    'Actualizar celda unidades
                    If booMilestone Then
                        If datPeriodStart <= IIf(IsDate(datActStart), datActStart, datActFinish) And _
                            datPeriodFinish >= IIf(IsDate(datActStart), datActStart, datActFinish) Then _
                            rngInsertUnits = rngRmgUnt.Offset(intR - rngRef.row)
                    ElseIf booOneDayDur Then
                        rngInsertUnits = rngRmgUnt.Offset(intR - rngRef.row)
                    ElseIf Not (arrAct(3) Like "WBS-*" Or Not IsEmpty(arrAct(0))) Then
                        Select Case intCurve
                        Case 1 'Linear
                            rngInsertUnits = (intPeriodFin - intPeriodIni) * dblLinearK
                        Case 2 'S-Curve
                            rngInsertUnits = ((WorksheetFunction.Norm_Dist(intPeriodFin, dblMeanS, dblVarS, True) - WorksheetFunction.Norm_Dist(intPeriodIni, dblMeanS, dblVarS, True)) + _
                                                (intPeriodFin - intPeriodIni) * dblErrorS) * dblUnits
                        Case 3 'Front Loaded
                            rngInsertUnits = (2 * (WorksheetFunction.Norm_Dist(intPeriodFin, dblMeanF, dblVarF, True) - WorksheetFunction.Norm_Dist(intPeriodIni, dblMeanF, dblVarF, True)) + _
                                                (intPeriodFin - intPeriodIni) * dblErrorF) * dblUnits
                        Case 4 'Back Loaded
                            rngInsertUnits = (2 * (WorksheetFunction.Norm_Dist(intPeriodFin, dblMeanB, dblVarB, True) - WorksheetFunction.Norm_Dist(intPeriodIni, dblMeanB, dblVarB, True)) + _
                                                (intPeriodFin - intPeriodIni) * dblErrorB) * dblUnits
                        Case 5 'Step 0-100
                            rngInsertUnits = IIf(datActFinish <= datRangeFinish And datActFinish >= datRangeStart, dblUnits, 0)
                        End Select
                    End If
                End If
                
                If (intPeriod = 3 Or intPeriod = 4) And Not Month(datPeriodStart) = Month(datPeriodFinish) And k > 0 Then
                    If rngInsertUnits > 0 Then Range(rngInsertUnits, rngInsertUnits.Offset(0, 1)).Merge
                    k = k + 2
                Else: k = k + 1
                End If

                Select Case intPeriod
                    Case 1, 2 'Diario
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = datPeriodStart
                    Case 3 'Semanal
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = datPeriodStart + 7 - 1
                    Case 4 'Bi-semanal
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = datPeriodStart + 14 - 1
                    Case 5 'Mensual
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = DateSerial(Year(datPeriodStart), Month(datPeriodStart) + 1, 1) - 1
                    Case 6 'Trimestre
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = DateSerial(Year(datPeriodStart), 3 * ((Month(datPeriodStart) - 1) \ 3) + 4, 1) - 1
                    Case 7 'Anual
                        datPeriodStart = datPeriodFinish + 1
                        datPeriodFinish = DateSerial(Year(datPeriodStart) + 1, 1, 1) - 1
                End Select
            Loop
        End If
NextIteration1:
        intR = intR + 1
        i = i + 1
    Next


'------------------------------------------------------------------------------------------------------------------------------------------
    'Segundo recorrido para calcular distribución de sumarios
    intR = rngRef.row + 1
    
    For Each arrAct In arrData
        'Comprobar si la actividad actual está en la lista de actividades a actualizar.
        If Not arrAct(10) Then GoTo NextIteration2

        strTmlMod = arrAct(0)
        strActID = arrAct(3)
        strWBS = arrAct(4)
        'Primero se comprueba si es un timeline
        If strTmlMod = "SUM" Or strTmlMod = "MIL" Or strTmlMod = "ACT" Then
            strTmlCod = arrAct(1)
            'Variable intRsum ajustada a las posiciones del vector
            intRsum = intR - rngRef.row - 1
            Do While strTmlCod = arrData(intRsum + 1)(1)
                intRsum = intRsum + 1
            Loop
            'Variable intRsum ajustada a las filas de la lista de actividades
            intRsum = intRsum + rngRef.row + 1
            'Si la posición final del Timeline es mayor que a inicial
            If intRsum > intR Then
                'Total Remaining Units del timeline
                dblUnits = 0
                For k = intR - rngRef.row To intRsum - rngRef.row - 1
                    dblUnits = dblUnits + arrData(k)(8)
                Next
                rngRmgUnt.Offset(intR - rngRef.row, 0) = dblUnits
                'Remaining Units de cada periodo
                If dblUnits > 0 Then
                    For l = rngRef.Column To intCalLastCol
                        dblUnits = WorksheetFunction.Sum(Range(rngRef.Offset(intR - rngRef.row + 1, l - rngRef.Column), rngRef.Offset(intRsum - rngRef.row, l - rngRef.Column)))
                        If dblUnits > 0 Then rngRef.Offset(intR - rngRef.row, l - rngRef.Column) = dblUnits
                    Next
                End If
            End If
            
            
        'Si no, se comprueba si es un sumario de WBS
        ElseIf Not strWBS = rngWBS.Offset(intR - rngRef.row - 1) And strActID Like "WBS-*" Then
            'Total Remaining Units del sumario
            dblUnits = 0
            For k = intR - rngRef.row To intActLastRow - rngRef.row - 1
                If Not arrData(k)(3) Like "WBS-*" And (arrData(k)(4) = strWBS Or arrData(k)(4) Like strWBS & ".*") And IsEmpty(arrData(k)(0)) Then
                    dblUnits = dblUnits + arrData(k)(8)
                End If
            Next
            rngRmgUnt.Offset(intR - rngRef.row, 0) = dblUnits
            
            'Remaining Units de cada periodo
            If dblUnits > 0 Then
                For l = rngRef.Column To intCalLastCol
                    dblUnits = WorksheetFunction.SumIfs(Range(rngRef.Offset(intR - rngRef.row + 1, l - rngRef.Column), rngRef.Offset(intActLastRow - rngRef.row, l - rngRef.Column)), _
                                        Range(rngWBS.Offset(intR - rngRef.row + 1, 0), rngWBS.Offset(intActLastRow - rngRef.row, 0)), strWBS, _
                                        Range(rngActID.Offset(intR - rngRef.row + 1, 0), rngActID.Offset(intActLastRow - rngRef.row, 0)), "<>WBS-*") + _
                                WorksheetFunction.SumIfs(Range(rngRef.Offset(intR - rngRef.row + 1, l - rngRef.Column), rngRef.Offset(intActLastRow - rngRef.row, l - rngRef.Column)), _
                                        Range(rngWBS.Offset(intR - rngRef.row + 1, 0), rngWBS.Offset(intActLastRow - rngRef.row, 0)), strWBS & ".*", _
                                        Range(rngActID.Offset(intR - rngRef.row + 1, 0), rngActID.Offset(intActLastRow - rngRef.row, 0)), "<>WBS-*")
                                                                                           
                    If dblUnits > 0 Then rngRef.Offset(intR - rngRef.row, l - rngRef.Column) = dblUnits
                Next
            End If
        End If
        
NextIteration2:
        intR = intR + 1
    Next
End Sub

