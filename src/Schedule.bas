Attribute VB_Name = "Schedule"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition and Pro Edition
Option Explicit

Public booLoopStatusPh1 As Boolean
Public intRowUpd As Variant 'vector para almacenar actividades que han cambiado para actualización automática


Sub CalculateSchedule(Optional ByVal booAutoSchedule As Boolean = False)
    If intEdition > 0 Then Exit Sub
    
    Dim ActList() As Variant
    Dim objAct As actObj
    Dim arrActID(), arrRmgDur(), arrStart(), arrFinish(), arrActStart(), arrActFinish(), arrResume(), arrConstraint(), arrFloat(), arrSchMod(), arrPred(), arrTmlMod As Variant
    Dim itmAct, itmPred, itmSucc As Variant
    Dim i, j, k, l, m As Integer 'contadores
    Dim intDim As Integer
    Dim strActID As String
    'Variables para tabla de Loops
    Dim wsLoop As Worksheet
    Dim tblLoop As ListObject
    Dim rowNew As ListRow
    
    
    '1) CREACION DE LA ESTRUCTURA DE DATOS -----------------------------------------------------------------------------------------------------------------
    
    intDim = intActLastRow - rngRef.row
    
    If intDim = 1 Then
        If rngActID.Offset(1).value = "" Or Not IsNumeric(rngRmgDur.Offset(1).value) Then Exit Sub
        'Lectura de datos de la hoja de trabajo
        ReDim arrActID(1, 1)
        ReDim arrRmgDur(1, 1)
        ReDim arrStart(1, 1)
        ReDim arrFinish(1, 1)
        ReDim arrActStart(1, 1)
        ReDim arrActFinish(1, 1)
        ReDim arrResume(1, 1)
        ReDim arrConstraint(1, 1)
        ReDim arrFloat(1, 1)
        ReDim arrSchMod(1, 1)
        ReDim arrPred(1, 1)
        ReDim arrTmlMod(1, 1)
        arrActID(1, 1) = rngActID.Offset(1).value
        arrRmgDur(1, 1) = rngRmgDur.Offset(1).value
        arrStart(1, 1) = rngStart.Offset(1).value
        arrFinish(1, 1) = rngFinish.Offset(1).value
        arrActStart(1, 1) = rngStartAct.Offset(1).value
        arrActFinish(1, 1) = rngFinishAct.Offset(1).value
        arrResume(1, 1) = rngResume.Offset(1).value
        arrConstraint(1, 1) = rngConstraint.Offset(1).value
        arrFloat(1, 1) = rngFloat.Offset(1).value
        arrSchMod(1, 1) = rngSchMod.Offset(1).value
        arrPred(1, 1) = rngPred.Offset(1).value
        arrTmlMod(1, 1) = rngTmlMod.Offset(1).value

    Else
        'Lectura de datos de la hoja de trabajo
        arrActID = Range(rngActID.Offset(1), rngActID.Offset(intDim)).value
        arrRmgDur = Range(rngRmgDur.Offset(1), rngRmgDur.Offset(intDim)).value
        arrStart = Range(rngStart.Offset(1), rngStart.Offset(intDim)).value
        arrFinish = Range(rngFinish.Offset(1), rngFinish.Offset(intDim)).value
        arrActStart = Range(rngStartAct.Offset(1), rngStartAct.Offset(intDim)).value
        arrActFinish = Range(rngFinishAct.Offset(1), rngFinishAct.Offset(intDim)).value
        arrResume = Range(rngResume.Offset(1), rngResume.Offset(intDim)).value
        arrConstraint = Range(rngConstraint.Offset(1), rngConstraint.Offset(intDim)).value
        arrFloat = Range(rngFloat.Offset(1), rngFloat.Offset(intDim)).value
        arrSchMod = Range(rngSchMod.Offset(1), rngSchMod.Offset(intDim)).value
        arrPred = Range(rngPred.Offset(1), rngPred.Offset(intDim)).value
        arrTmlMod = Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value
    End If
    
        k = 0
        'Recorrer lista de actividades e ir ampliando el vector ActList con las nuevas filas
        For i = 1 To intDim
            If Not (arrActID(i, 1) = Empty Or arrActID(i, 1) Like "WBS-*") Then
                ReDim Preserve ActList(k)
                Set ActList(k) = New actObj
                'Este método pasa los datos de la lista de actividades a las propiedades del objeto creado en la nueva posición del vector
                ActList(k).WriteProperties rngRef.row + i, arrActID(i, 1), arrRmgDur(i, 1), arrStart(i, 1), arrFinish(i, 1), arrActStart(i, 1), arrActFinish(i, 1), _
                        arrResume(i, 1), arrConstraint(i, 1), arrFloat(i, 1), arrSchMod(i, 1), arrPred(i, 1), arrTmlMod(i, 1)
                k = k + 1
            End If
        Next
        
        'Para cada actividad en el vector ActList
        i = 0
        For Each itmAct In ActList
            'Si el vector de predecesoras no está vacío
            If Not itmAct.fPredEmpty Then
                'Se recorre el vector de predecesoras
                j = 0
                For Each itmPred In itmAct.PredList
                    k = 0
                    'buscando la posición en la que se encuentra la predecesora en el vector ActList
                    strActID = itmPred.ActID
                    Do Until ActList(k).ActID = strActID
                        k = k + 1
                        If k > UBound(ActList) Then GoTo missedPred
                    Loop
                    'y se añade la referencia d ela posición de la propiedad ArrID de la predecesora
                    itmPred.ArrID = k
                    
                    '*****Podría encapsular esta parte para permitir un cálculo sin fechas Late y sin Float
                    'Además, se rellena el vector de sucesoras
                    'Primero se determina la dimensión del vector de destino
                    If ActList(k).fSuccEmpty Then
                        l = 0
                        ActList(k).fSuccEmpty = False
                    Else
                        l = UBound(ActList(k).SuccList) + 1
                    End If
                    'Este método redimensiona el vector SuccList dentro del elemento actual de ActList
                    ActList(k).RedimSuccList l
                    'Y asignamos los valores al nuevo elemento de SuccList
                    ActList(k).SuccList()(l).ActID = itmAct.ActID
                    ActList(k).SuccList()(l).ArrID = i
                    ActList(k).SuccList()(l).RelType = itmAct.PredList()(j).RelType
                    ActList(k).SuccList()(l).Lag = itmAct.PredList()(j).Lag
                    '*****
                    j = j + 1
                Next
            End If
            i = i + 1
        Next

    If Not booAutoSchedule Then UpdateProgressBar 0.1
    
    
    '2.1) COMPROBACION DE LOOPS FASE 1 -----------------------------------------------------------------------------------------------------------------
    Dim cntLoopChk, cntLoopChkPre, cntPredLoopChk, cntSuccLoopChk As Integer
    'Inicialización de variables para no bloquear la entrada inicial al Loop
    cntLoopChk = 0
    cntLoopChkPre = -1
    
    'Mientras el recuento de Loops chequeados en la iteración anterior sea menor que el recuento de la última iteración se sigue iterando
    Do While cntLoopChkPre < cntLoopChk
        'Inicialización de las variables
        cntLoopChkPre = cntLoopChk
        cntLoopChk = 0
        'Se recorre la lista de actividades
        For Each itmAct In ActList
            'Para cada actividad
            With itmAct
                'Si no está chequeado todavía que no forma parte de un Loop
                If Not .fLoopChkPh1 Then
                    'Si la actividad no tiene predecesoras o no va a formar parte del cálculo queda chequeada
                    If .fPredEmpty Or .fSuccEmpty Or .fSchNo Then
                        .fLoopChkPh1 = True
                    Else 'En el resto de casos se debe comprobar el estado de todas sus predecesoras o sucesoras
                        cntPredLoopChk = 0
                        'Para cada predecesora se comprueba si está chequeada
                        For Each itmPred In .PredList
                            'Si lo está se suma 1 al contador de chequeo de predecesoras
                            If ActList(itmPred.ArrID).fLoopChkPh1 Then cntPredLoopChk = cntPredLoopChk + 1
                        Next
                        'Si el nº de predecesoras que efectivamente no forman parte de un loop es igual a la dimensión del vector de predecesoras
                        'la actividad queda chequeada
                        If cntPredLoopChk = UBound(.PredList) + 1 Then
                            .fLoopChkPh1 = True
                        Else 'Si la actividad no queda chequeada por sus predecesoras quizá sí que quede chequeada por sus sucesoras
                            cntSuccLoopChk = 0
                            'Para cada sucesora se comprueba si está chequeada
                            For Each itmSucc In .SuccList
                                'Si lo está se suma 1 al contador de chequeo de predecesoras
                                If ActList(itmSucc.ArrID).fLoopChkPh1 Then cntSuccLoopChk = cntSuccLoopChk + 1
                            Next
                            'Si el nº de sucesoras que efectivamente no forman parte de un loop es igual a la dimensión del vector de sucesoras
                            'la actividad queda chequeada
                            If cntSuccLoopChk = UBound(.SuccList) + 1 Then .fLoopChkPh1 = True
                        End If
                    End If
                End If
                'Si la actividad está chequeada se suma 1 al contador de actividades chqueadas en esta iteración
                If .fLoopChkPh1 Then cntLoopChk = cntLoopChk + 1
            End With
        Next
    Loop
    
    If Not booAutoSchedule Then UpdateProgressBar 0.2
    
    
    '2.2) COMPROBACION DE LOOPS FASE 2 -----------------------------------------------------------------------------------------------------------------
    '***** Podría encapsular esta parte para permitir un cálculo sin detección de elementos en Loop
    'Entramos sólo si en la fase 1 hay menos actividades chequeadas que actividades totales
    If cntLoopChk < UBound(ActList) + 1 Then
        booLoopStatusPh1 = True
        'Variables para controlar el número de Loop y la posición de una actividad dentro del Loop
        Dim intLoopNo, intLoopPos As Integer
        intLoopNo = 1
        intLoopPos = 1
        'Recorremos la lista de actividad una sola vez. Cada vez que encuentre el inicio del Loop va a recorrer el Loop completo y luego itera a la siguiente posición del vector.
        For k = 0 To UBound(ActList)
            'Variable adicional para controlar la posición del vector ActList en la que estamos trabajando
            i = k
'Inicio de la evaluación de la posicón del vector ActList
NextPredInLoop:
            'Para la posición actual
            With ActList(i)
                'Si la posición no ha sido chequeada en la Fase 1 y no ha sido identificada en un Loop en la Fase 2
                If Not .fLoopChkPh1 And Not .LoopStatusPh2 = 2 Then
                    'Recorremos la lista de predecesoras de la posición actual
                    For j = 0 To UBound(.PredList)
                        'Si la predecesora actual no ha sido chequeada en la Fase 1 y no ha sido identificada en un Loop en la Fase 2 la actividad actual puede formar parte de un Loop
                        'Si la predecesora está en Status 2 y la actividad actual está en Status 1 significa que es la última actividad de un bucle a punto de ser cerrado
                        If (Not ActList(.PredList()(j).ArrID).fLoopChkPh1 And Not ActList(.PredList()(j).ArrID).LoopStatusPh2 = 2) Or _
                            (ActList(.PredList()(j).ArrID).LoopStatusPh2 = 2 And .LoopStatusPh2 = 1) Then
                            'Si el Status de la actividad actual es 0 (nunca revisada en esta iteración) se pasa a 1
                            If .LoopStatusPh2 = 0 Then
                                .LoopStatusPh2 = 1
                            'Si el Status de la actividad actual es 0 (revisada una vez en esta iteración) se pasa a 2
                            'Esto significa que se ha cerrado un Loop y se comienza la identificación de sus parámetros
                            ElseIf .LoopStatusPh2 = 1 Then
                                .LoopNo = intLoopNo
                                .LoopPos = intLoopPos
                                intLoopPos = intLoopPos + 1
                                .LoopStatusPh2 = 2
                                'Si la actividad actual era la última del Loop se itera el valor de intLoopNo y se resetean los valores de LoopStatusPh2 = 1 a 0
                                If ActList(.PredList()(j).ArrID).LoopStatusPh2 = 2 Then
                                    intLoopNo = intLoopNo + 1
                                    'Recorremos las sucesoras de la primera actividad en el Loop buscando actividades con LoopStatusPh2=1 e iteramos hasta que no encontremos más
PrevStatus1:
                                    m = .PredList()(j).ArrID
                                    For l = 0 To UBound(ActList(m).SuccList)
                                        If ActList(ActList(m).SuccList()(l).ArrID).LoopStatusPh2 = 1 Then
                                            ActList(ActList(m).SuccList()(l).ArrID).LoopStatusPh2 = 0
                                            GoTo PrevStatus1
                                        End If
                                    Next
                                End If
                            End If
                            'Si el Status previo de la actividad era 0 ó 1, se debe seguir recorriendo el Loop por la predecesora actual que se convertirá en la actividad analizada
                            i = .PredList()(j).ArrID
                            GoTo NextPredInLoop
                        End If
                    Next
                    'Si todas las predecesoras han sido chequeadas en la Fase 1 o han sido ya identificadas en un Loop en la Fase 2
                    'la actividad actual queda chequeada
                    .fLoopChkPh1 = True
                End If
            End With
            'Antes de pasar a la siguiente iteración se reseta la variable LoopPos a 1 y se analiza la siguiente posición del vecto ActList
            intLoopPos = 1
        Next
        
        'Mostrar loops en una hoja
        'Crear nueva hoja si no existe y tabla de actividades
        For Each wsLoop In ActiveWorkbook.Worksheets
            If wsLoop.Name = "loops_summary" Then
                Set wsLoop = ActiveWorkbook.Worksheets("loops_summary")
                Set tblLoop = wsLoop.ListObjects(1)
                GoTo LoopsWsExists
            End If
        Next
        Worksheets.Add.Name = "loops_summary"
        Set wsLoop = ActiveWorkbook.Worksheets("loops_summary")
        wsLoop.Cells(1, 1) = "Loop No"
        wsLoop.Cells(1, 2) = "Loop Step"
        wsLoop.Cells(1, 3) = "Activity ID"
        Set tblLoop = wsLoop.ListObjects.Add(xlSrcRange, wsLoop.Range(Cells(1, 1), Cells(1, 3)), , xlYes)
LoopsWsExists:
        With tblLoop
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete
            End If
        End With
        'Llenar tabla
        For i = 0 To UBound(ActList)
            If ActList(i).LoopStatusPh2 = 2 Then
                Set rowNew = tblLoop.ListRows.Add
                With rowNew
                    .Range(1) = ActList(i).LoopNo
                    .Range(2) = ActList(i).LoopPos
                    .Range(3) = "'" & ActList(i).ActID
                 End With
            End If
        Next
        'Ordenar tabla
        With wsLoop.ListObjects(tblLoop.Name).Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range(tblLoop.Name & "[Loop No]"), SortOn _
                :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range(tblLoop.Name & "[Loop Step]"), SortOn _
                :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Apply
        End With
        
        UpdateProgressBar 1
        wsLoop.Activate
        
        'Salir del procedimiento
        Exit Sub
    Else 'Si no se han encontrado Loops eliminar tabla de loops de ejecuciones anteriores
        booLoopStatusPh1 = False
        For Each wsLoop In ActiveWorkbook.Worksheets
        If wsLoop.Name = "loops_summary" Then
            Application.DisplayAlerts = False
            wsLoop.Delete
            Application.DisplayAlerts = True
        End If
        Next
    End If
    
    'Salir del cálculo si no hay fecha de corte
    If datCutoff = 0 Then Exit Sub
    
    '3) CALCULO DE FECHAS EARLY -----------------------------------------------------------------------------------------------------------------
    Dim intEarlyDone As Integer
    Dim datEFinish, datEFinishMax As Date
    Dim strRelType As String
    Dim dblRmgDur, dblLag As Double
    Dim datRefStart, datRefFinish As Date
    Dim datPrjFin As Date
    'Contador de actividades con fecha Early calculada que se inicializa a cero
    intEarlyDone = 0
    datPrjFin = datCutoff
    'El bucle va a iterarse hasta que todos los elementos de la lista hayan sido calculados
    Do While intEarlyDone < UBound(ActList) + 1
        'Inicializar contador en cada iteración
        intEarlyDone = 0
        'Se recorre el vector ActList completo
        For Each itmAct In ActList
            With itmAct
            'Si la actividad todavía no ha sido calculada y el SchMode no es "NO"
                If Not .fEarlyDone And Not .fSchNo Then
                    'Cálculo corregido de la duración remanente
                    dblRmgDur = IIf(.ActType = "?ML" Or .RmgDuration = 0, 0, .RmgDuration - 1)
                    If .CnstType = "MANU" Then
                        'Código para cálculo de fechas Early en actividades restringidas
                        .EarlyStart = .ConstraintStart
                        .EarlyFinish = .ConstraintFinish
                    Else
                        If .fPredEmpty Then
                            'Código para cálculo de fechas Early en actividades sin predecesoras
                            datEFinishMax = DateAddCal(dblRmgDur + 1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff)
                        Else 'La actividad tiene predecesoras de modo que se recorren todas
                            For j = 0 To UBound(.PredList)
                                'Si una de las predecesoras todavía no está calculada
                                'se salta a la siguiente iteración y la actividad se analizará en una iteración posterior
                                If Not ActList(.PredList()(j).ArrID).fEarlyDone And Not ActList(.PredList()(j).ArrID).fSchNo Then GoTo NextActivity_EarlyCalc
                            Next
                            'Si se sale del bucle normalmente significa que la actividad es calculable.
                            'Se inicializa el cálculo con las fechas Early basadas en Cutoff date
                            datEFinishMax = DateAddCal(dblRmgDur + 1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff)
                            'Se vuelve a recorrer el vector de predecesoras calculando la fecha Early Finish fijada por cada relación y comparando con la fijada por la anterior predecesora
                            For j = 0 To UBound(.PredList)
                                'Las actividades con el flag fSchNo no intervienen en el cálculo y se deben saltar
                                If Not ActList(.PredList()(j).ArrID).fSchNo Then
                                    'Código para cálculo de fechas Early en actividades con todas las predecesoras calculadas
                                    dblLag = .PredList()(j).Lag
                                    datRefStart = IIf(ActList(.PredList()(j).ArrID).ActualStart > 0, ActList(.PredList()(j).ArrID).ActualStart, ActList(.PredList()(j).ArrID).EarlyStart)
                                    datRefFinish = IIf(ActList(.PredList()(j).ArrID).ActualFinish > 0, ActList(.PredList()(j).ArrID).ActualFinish, ActList(.PredList()(j).ArrID).EarlyFinish)
                                    Select Case .PredList()(j).RelType
                                    Case "FS"
                                        datEFinish = DateAddCal(dblLag + dblRmgDur + 1 - IIf(WorkingDay(datRefFinish), 0, 1), datRefFinish)
                                    Case "FF"
                                        datEFinish = DateAddCal(dblLag, datRefFinish)
                                    Case "SS"
                                        datEFinish = DateAddCal(dblLag + dblRmgDur, datRefStart)
                                    Case "SF"
                                        datEFinish = DateAddCal(dblLag, datRefStart)
                                    End Select
                                    'Comparación de la fecha Early Finish calculada con la máxima registrada y actualización de valores
                                    datEFinishMax = IIf(datEFinish > datEFinishMax, datEFinish, datEFinishMax)
                                End If
                            Next
                        End If
                        'Cálculo de fechas en función de constraints y guardado en las propiedades
                        If Not IsEmpty(.CnstType) And .ConstraintDate > 0 Then
                            'Si hay una constraint definida y hay fecha
                            Select Case .CnstType
                            Case "FMN" 'Mandatory Finish
                                'Si la fecha de la Constraint es anterior al corte se calcula como una actividad sin predecesora
                                If .ConstraintDate <= datCutoff Then
                                    .EarlyStart = DateAddCal(1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff)
                                    .EarlyFinish = DateAddCal(dblRmgDur, .EarlyStart)
                                Else 'Si no, se fija EarlyFinish a la ConstraintDate
                                    .EarlyFinish = .ConstraintDate
                                    'Si la EarlyStart calculada desde la ConstraintDate es posterior al Cutoff,
                                    If DateAddCal(-dblRmgDur, .EarlyFinish) > datCutoff Then
                                        'Se asigna el cálculo
                                        .EarlyStart = DateAddCal(-dblRmgDur, .EarlyFinish)
                                    Else: 'Si no, se calcula como el día siguiente al cutoff
                                        .EarlyStart = DateAddCal(1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff)
                                    End If
                                End If
                            Case "SMN" 'Mandatory Start
                                'Si la fecha de la Constraint es anterior al corte se calcula como una actividad sin predecesora
                                If .ConstraintDate <= datCutoff Then
                                    .EarlyStart = DateAddCal(1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff)
                                    .EarlyFinish = DateAddCal(dblRmgDur, .EarlyStart)
                                Else 'Si no, se fija EarlyStart a la ConstraintDate
                                    .EarlyStart = .ConstraintDate
                                    'Si la EarlyFinish calculada desde la ConstraintDate es posterior a la fecha fijada por sus predecesoras,
                                    If DateAddCal(dblRmgDur, .EarlyStart) > datEFinishMax Then
                                        'Se asigna el cálculo
                                        .EarlyFinish = DateAddCal(-dblRmgDur, .EarlyStart)
                                    Else: 'Si no, se mantiene la fecha fijada por sus predecesoras
                                        .EarlyFinish = datEFinishMax
                                    End If
                                End If
                            Case "FON", "FOB", "FOL" 'Finish On / or Later / or Before
                                datEFinish = .ConstraintDate
                                'Si la fecha de la Constraint es anterior al corte o es anterior a la fecha datFinishMax , se mantiene datFinishMax y se calcula el inicio
                                If datEFinish <= datCutoff Or datEFinish <= datEFinishMax Then
                                    .EarlyFinish = datEFinishMax
                                    .EarlyStart = DateAddCal(-dblRmgDur, datEFinishMax)
                                Else 'La fecha de la Constraint es posterior a datEFinishMax
                                    If .CnstType = "FOB" Then
                                        .EarlyFinish = datEFinishMax
                                    Else
                                        .EarlyFinish = datEFinish
                                    End If
                                    .EarlyStart = DateAddCal(-dblRmgDur, .EarlyFinish)
                                End If
                            Case "SON", "SOB", "SOL" 'Start On / or Later / or Before
                                datEFinish = DateAddCal(dblRmgDur, .ConstraintDate)
                                'Si la fecha de la Constraint es anterior al corte o es anterior a la fecha datFinishMax , se mantiene datFinishMax y se calcula el inicio
                                If datEFinish <= datCutoff Or datEFinish <= datEFinishMax Then
                                    .EarlyFinish = datEFinishMax
                                    .EarlyStart = DateAddCal(-dblRmgDur, datEFinishMax)
                                Else 'La fecha de la Constraint es posterior a datEFinishMax
                                    If .CnstType = "SOB" Then
                                        .EarlyFinish = datEFinishMax
                                    Else
                                        .EarlyFinish = datEFinish
                                    End If
                                    .EarlyStart = DateAddCal(-dblRmgDur, .EarlyFinish)
                                End If
                            Case Else 'Para cualquier otro tipo de Constraint el cálculo es como sin Constraint
                                .EarlyFinish = datEFinishMax
                                .EarlyStart = DateAddCal(-dblRmgDur, datEFinishMax)
                            End Select
                        Else 'Si no hay constraint definida o no hay ConstraintDate
                            .EarlyFinish = datEFinishMax
                            .EarlyStart = DateAddCal(-dblRmgDur, datEFinishMax)
                        End If
                    End If
                    'Comparación de la fecha Early Finish calculada con la máxima registrada para el final del proyecto
                    datPrjFin = IIf(datPrjFin > .EarlyFinish, datPrjFin, .EarlyFinish)
                End If
NextActivity_EarlyCalc:
            'Actualizamos contador en función de si la actividad actual está calculada
            If .fEarlyDone Or .fSchNo Then intEarlyDone = intEarlyDone + 1
            End With
        Next
    Loop
    
    If Not booAutoSchedule Then UpdateProgressBar 0.5
    
    '4) CALCULO DE FECHAS LATE -----------------------------------------------------------------------------------------------------------------
    Dim intLateDone As Integer
    Dim datLStart, datLStartMin As Date
    Dim datEFinishMin As Date

    'Contador de actividades con fecha Early calculada que se inicializa a cero
    intLateDone = 0
    'El bucle va a iterarse hasta que todos los elementos de la lista hayan sido calculados
    Do While intLateDone < UBound(ActList) + 1
        'Inicializar contador en cada iteración
        intLateDone = 0
        'Se recorre el vector ActList completo
        For Each itmAct In ActList
            With itmAct
            'Si la actividad todavía no ha sido calculada y el SchMode no es "NO"
                If Not .fLateDone And Not .fSchNo Then
                    'Cálculo corregido de la duración remanente
                    dblRmgDur = IIf(.ActType = "?ML" Or .RmgDuration = 0, 0, .RmgDuration - 1)
                    'Si la actividad tiene una Constraint Mandatory las fechas Late deben ser iguales a las fechas Early
                    If (.CnstType = "SMN" Or .CnstType = "FMN") And Not IsEmpty(.ConstraintDate) Then
                        .LateFinish = .EarlyFinish
                        .LateStart = .EarlyStart
                    Else
                        If .fSuccEmpty Then
                            'Código para cálculo de fechas Early en actividades sin predecesoras
                            datLStartMin = DateAddCal(-dblRmgDur, datPrjFin)
                        Else 'La actividad tiene sucesoras de modo que se recorren todas
                            For j = 0 To UBound(.SuccList)
                                'Si una de las predecesoras todavía no está calculada
                                'se salta a la siguiente iteración y la actividad se analizará en una iteración posterior
                                If Not ActList(.SuccList()(j).ArrID).fLateDone And Not ActList(.SuccList()(j).ArrID).fSchNo Then GoTo NextActivity_LateCalc
                            Next
                            'Si se sale del bucle normalmente significa que la actividad es calculable.
                            'Se inicializa el cálculo con las fechas Early basadas en Cutoff date
                            datLStartMin = DateAddCal(-dblRmgDur, datPrjFin)
                            'Se vuelve a recorrer el vector de predecesoras calculando la fecha Early Finish fijada por cada relación y comparando con la fijada por la anterior predecesora
                            For j = 0 To UBound(.SuccList)
                                'Las actividades con el flag fSchNo no intervienen en el cálculo y se deben saltar
                                If Not ActList(.SuccList()(j).ArrID).fSchNo Then
                                    'Código para cálculo de fechas Early en actividades con todas las predecesoras calculadas
                                    dblLag = .SuccList()(j).Lag
                                    datRefStart = ActList(.SuccList()(j).ArrID).LateStart
                                    datRefFinish = ActList(.SuccList()(j).ArrID).LateFinish
                                    Select Case .SuccList()(j).RelType
                                    Case "FS"
                                        datLStart = DateAddCal(-(dblLag + dblRmgDur + 1), datRefStart)
                                    Case "FF"
                                        datLStart = DateAddCal(-(dblLag + dblRmgDur), datRefFinish)
                                    Case "SS"
                                        datLStart = DateAddCal(-dblLag, datRefStart)
                                    Case "SF"
                                        datLStart = DateAddCal(-dblLag, datRefFinish)
                                    End Select
                                    'Comparación de la fecha Early Finish calculada con la máxima registrada y actualización de valores
                                    datLStartMin = IIf(datLStart < datLStartMin, datLStart, datLStartMin)
                                    datLStartMin = IIf(datLStartMin < datCutoff, DateAddCal(1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff), datLStartMin)
                                End If
                            Next
                        End If
                        'Cálculo de fechas en función de constraints y guardado en las propiedades
                        If Not IsEmpty(.CnstType) And .ConstraintDate > 0 Then
                            'Si hay una constraint definida y hay fecha
                            Select Case .CnstType
                            Case "FON"
                                .LateFinish = .ConstraintDate
                                .LateStart = DateAddCal(-dblRmgDur, .ConstraintDate)
                            Case "FOB"
                                .LateFinish = IIf(DateAddCal(dblRmgDur, datLStartMin) < .ConstraintDate, DateAddCal(dblRmgDur, datLStartMin), .ConstraintDate)
                                .LateStart = DateAddCal(-dblRmgDur, .LateFinish)
                            Case "FOL"
                                .LateFinish = IIf(DateAddCal(dblRmgDur, datLStartMin) > .ConstraintDate, DateAddCal(dblRmgDur, datLStartMin), .ConstraintDate)
                                .LateStart = DateAddCal(-dblRmgDur, .LateFinish)
                            Case "SON"
                                .LateFinish = DateAddCal(dblRmgDur, .ConstraintDate)
                                .LateStart = .ConstraintDate
                            Case "SOB"
                                .LateStart = IIf(datLStartMin < .ConstraintDate, datLStartMin, .ConstraintDate)
                                .LateFinish = DateAddCal(dblRmgDur, .LateStart)
                            Case "SOL"
                                .LateStart = IIf(datLStartMin > .ConstraintDate, datLStartMin, .ConstraintDate)
                                .LateFinish = DateAddCal(dblRmgDur, .LateStart)
                            Case Else 'Para cualquier otro tipo de Constraint el cálculo es como sin Constraints
                                .LateFinish = DateAddCal(dblRmgDur, datLStartMin)
                                .LateStart = datLStartMin
                            End Select
                        Else 'Si no hay constraint definida o no hay fecha
                            .LateFinish = DateAddCal(dblRmgDur, datLStartMin)
                            .LateStart = datLStartMin
                        End If
                    End If
                    'Si tenemos una restricción ALAP la fecha early es las más temprana fijada por sus sucesoras.
                    If .CnstType = "ALAP" Then
                        If .fSuccEmpty Then
                            .EarlyFinish = .LateFinish
                            .EarlyStart = .LateStart
                        Else
                            datEFinishMin = .LateFinish
                            'Se vuelve a recorrer el vector de sucesoras calculando la fecha Early Finish fijada por cada relación y comparando con la fijada por la anterior predecesora
                            For j = 0 To UBound(.SuccList)
                                dblLag = .SuccList()(j).Lag
                                datRefStart = ActList(.SuccList()(j).ArrID).EarlyStart
                                datRefFinish = ActList(.SuccList()(j).ArrID).EarlyFinish
                                Select Case .SuccList()(j).RelType
                                Case "FS"
                                    datEFinish = DateAddCal(-(dblLag + 1), datRefStart)
                                Case "FF"
                                    datEFinish = DateAddCal(-dblLag, datRefFinish)
                                Case "SS"
                                    datEFinish = DateAddCal(-dblLag + dblRmgDur, datRefStart)
                                Case "SF"
                                    datEFinish = DateAddCal(-dblLag + dblRmgDur, datRefFinish)
                                End Select
                                'Comparación de la fecha Early Finish calculada con la mínima registrada y actualización de valores
                                datEFinishMin = IIf(datEFinish < datEFinishMin, datEFinish, datEFinishMin)
                            Next
                            .EarlyStart = DateAddCal(-dblRmgDur, datEFinishMin)
                            .EarlyFinish = datEFinishMin
                        End If
                    End If
                End If
NextActivity_LateCalc:
            'Actualizamos contador en función de si la actividad actual está calculada
            If .fLateDone Or .fSchNo Then intLateDone = intLateDone + 1
            End With
        Next
    Loop
    
    If Not booAutoSchedule Then UpdateProgressBar 0.8

    '5) ACTUALIZACION DE VALORES EN HOJA DE DATOS -----------------------------------------------------------------------------------------------------------------
    Dim booRowUpdated As Boolean
    
    For Each itmAct In ActList
        booRowUpdated = False
        With itmAct
            If (.Start <> .CalcStart And Not (.Start = 0 And .CalcStart = "")) Then
                rngStart.Offset(.row - rngRef.row) = .CalcStart
                booRowUpdated = True
            End If
            If (.Finish <> .CalcFinish And Not (.Finish = 0 And .CalcFinish = "")) Then
                rngFinish.Offset(.row - rngRef.row) = .CalcFinish
                booRowUpdated = True
            End If
            If (.ResumeDate <> .CalcResumeDate And Not (.ResumeDate = 0 And .CalcResumeDate = "")) Then
                rngResume.Offset(.row - rngRef.row) = .CalcResumeDate
                booRowUpdated = True
            End If
            If .Float <> .CalcFloat Or (IsEmpty(.Float) And .CalcFloat = 0) Then
                rngFloat.Offset(.row - rngRef.row) = .CalcFloat
                booRowUpdated = True
            End If
            If booAutoSchedule And booRowUpdated Then
                If IsEmpty(intRowUpd) Then
                    ReDim intRowUpd(0)
                Else
                    ReDim Preserve intRowUpd(UBound(intRowUpd) + 1)
                End If
                intRowUpd(UBound(intRowUpd)) = .row
            End If
        End With
    Next
    
    If Not booAutoSchedule Then
        UpdateProgressBar 1
    Else
        intRowUpd = GetEditRows_wSum(intRowUpd)
    End If
    
    Exit Sub
    
missedPred:
    UpdateProgressBar 1
    CustomMsgBox "Missing predecessor at row: " & ActList(i).row & vbCrLf & _
    "Activity ID: " & ActList(i).ActID & vbCrLf & _
    "Predecessor: " & strActID, vbCritical + vbOKOnly, error:=True
End Sub

