Attribute VB_Name = "Functions"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public Function WorkingDay(ByVal DateChk As Date) As Boolean
    Dim booWorkingDay As Boolean
    Dim CalException, FilterException As Variant
    
    WeekCalendar
    
    Select Case Weekday(DateChk)
        Case 1
            booWorkingDay = booSun
        Case 2
            booWorkingDay = booMon
        Case 3
            booWorkingDay = booTue
        Case 4
            booWorkingDay = booWed
        Case 5
            booWorkingDay = booThu
        Case 6
            booWorkingDay = booFri
        Case 7
            booWorkingDay = booSat
    End Select
    If booWorkingDay Then
        If InStr(ActiveWorkbook.CustomDocumentProperties("cdpCalExc").value, _
                format(DateChk, "yyyymmdd")) > 0 Then
            booWorkingDay = False
        End If
    End If
    WorkingDay = booWorkingDay
End Function

Public Function WorkingWeekDay(ByVal Date1 As Date) As Boolean
    WeekCalendar
        Select Case Weekday(Date1)
        Case 1
            WorkingWeekDay = booSun
        Case 2
            WorkingWeekDay = booMon
        Case 3
            WorkingWeekDay = booTue
        Case 4
            WorkingWeekDay = booWed
        Case 5
            WorkingWeekDay = booThu
        Case 6
            WorkingWeekDay = booFri
        Case 7
            WorkingWeekDay = booSat
    End Select
End Function

Public Function DateAddCal(ByVal Duration As Double, ByVal Date1 As Date) As Date
    Dim booWorkingDay As Boolean
    Dim datDateAdd As Date
    Dim intFullWeeks, intDuration, intSign As Integer
    
    WeekCalendar
    
    intDuration = CInt(Duration)
    datDateAdd = Date1
    intSign = IIf(Duration >= 0, 1, -1)

    'Si la fecha de entrada esta en un día no laborable, mover fecha hasta el siguiente día laborable
    Do While Not booWorkingDay
        booWorkingDay = WorkingWeekDay(datDateAdd)
        If Not booWorkingDay Then datDateAdd = DateAdd("d", intSign, datDateAdd)
    Loop
    
    'Añadir número de días correspondientes a semanas enteras en base a días laborables por semana
    intFullWeeks = intDuration \ intWorkingDays
    datDateAdd = DateAdd("d", intFullWeeks * 7, datDateAdd)

    'Si la fecha de salida esta en un día no laborable, mover fecha hasta el anterior día laborable
    booWorkingDay = False
    Do While Not booWorkingDay
        booWorkingDay = WorkingWeekDay(datDateAdd)
        If Not booWorkingDay Then datDateAdd = DateAdd("d", -intSign, datDateAdd)
    Loop
    
    'Sumar remanente de días teniendo en cuenta sólo días laborables
    intDuration = intDuration - intFullWeeks * intWorkingDays
    Do While Abs(intDuration) > 0
        If WorkingWeekDay(DateAdd("d", intSign, datDateAdd)) Then intDuration = intDuration - intSign
        datDateAdd = DateAdd("d", intSign, datDateAdd)
    Loop
    
    Dim datExc As Variant
    Dim i As Integer
    'arrDatesExc es una variable publica que contiene la lista de excepciones
    If Not VarType(arrDatesExc) = 0 Then
        If Not IsEmpty(arrDatesExc(0)) Then
            If intSign > 0 Then
                i = 0
                datExc = arrDatesExc(i)
                Do While datDateAdd >= datExc And i <= UBound(arrDatesExc)
                    If Date1 <= datExc Then
                        datDateAdd = DateAdd("d", intSign, datDateAdd)
                        Do While Not WorkingWeekDay(datDateAdd)
                            datDateAdd = DateAdd("d", intSign, datDateAdd)
                        Loop
                    End If
                i = i + 1
                If i <= UBound(arrDatesExc) Then datExc = arrDatesExc(i)
                Loop
            Else
                i = UBound(arrDatesExc)
                datExc = arrDatesExc(i)
                Do While datDateAdd <= datExc And i >= 0
                    If Date1 >= datExc Then
                        datDateAdd = DateAdd("d", intSign, datDateAdd)
                        Do While Not WorkingWeekDay(datDateAdd)
                            datDateAdd = DateAdd("d", intSign, datDateAdd)
                        Loop
                    End If
                i = i - 1
                If i >= 0 Then datExc = arrDatesExc(i)
                Loop
            
            End If
        End If
    End If

    'Devolver valor
    DateAddCal = datDateAdd

End Function

Public Function DateDiffCal(ByVal Date1, ByVal Date2 As Date) As Variant
    Dim datDateIni, datDateFin, datDateFin_wTail As Date
    Dim intDaysDiff, intDaysTail, intDaysExc As Integer

On Error GoTo errHandler
    
    WeekCalendar
    
    datDateIni = IIf(Date1 < Date2, Date1, Date2)
    datDateFin = IIf(Date1 > Date2, Date1, Date2)
    datDateFin_wTail = datDateFin
    
    intDaysTail = 0
    Do While Weekday(datDateIni) <> Weekday(datDateFin)
        If WorkingWeekDay(datDateFin) Then intDaysTail = intDaysTail + 1
        datDateFin = DateAdd("d", -1, datDateFin)
    Loop
    
    Dim datExc As Variant
    'arrDatesExc es una variable publica que contiene la lista de excepciones
    If Not VarType(arrDatesExc) = 0 Then
        If Not IsEmpty(arrDatesExc(0)) Then
            For Each datExc In arrDatesExc
                If datExc >= datDateIni And datExc <= datDateFin_wTail Then
                    intDaysExc = intDaysExc + 1
                End If
            Next
        End If
    End If
    
    intDaysDiff = intDaysTail + intWorkingDays * DateDiff("d", datDateIni, datDateFin) \ 7 - intDaysExc
    
    DateDiffCal = intDaysDiff * IIf(Date1 > Date2, -1, 1)
    Exit Function

errHandler:
    DateDiffCal = CVErr(xlErrValue)
    
End Function

'Ordenar alfabéticamente un vector
Public Function SortArrayAtoZ(myArray As Variant)
    Dim i, j As Long
    Dim Temp
    
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
            If UCase(myArray(i)) > UCase(myArray(j)) Then
                Temp = myArray(j)
                myArray(j) = myArray(i)
                myArray(i) = Temp
            End If
        Next j
    Next i

SortArrayAtoZ = myArray
        
End Function

'Ordenar vector numérico ascendente
Public Function SortArrayAscendent(myArray As Variant)
    Dim i, j As Long
    Dim Temp
    
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
            If myArray(i) > myArray(j) Then
                Temp = myArray(j)
                myArray(j) = myArray(i)
                myArray(i) = Temp
            End If
        Next j
    Next i

SortArrayAscendent = myArray
        
End Function

'Ordenar vector multidimensional por valor (0,i) numérico ascendente
Public Function SortArrayAscendentMulti(myArray As Variant)
    Dim i, j, k As Long
    Dim Temp
    
    For i = LBound(myArray) To UBound(myArray) - 1
        For j = i + 1 To UBound(myArray)
            If myArray(i, 0) > myArray(j, 0) Then
                For k = 0 To UBound(myArray, 2)
                    Temp = myArray(j, k)
                    myArray(j, k) = myArray(i, k)
                    myArray(i, k) = Temp
                Next
            End If
        Next j
    Next i

SortArrayAscendentMulti = myArray
        
End Function


'Invertir vector
Public Function ReverseArray(myArray As Variant)
    Dim i, j As Long
    Dim tempArray() As Variant
    
    ReDim tempArray(UBound(myArray))
    
    For i = LBound(myArray) To UBound(myArray)
        tempArray(UBound(myArray) - i) = myArray(i)
    Next i

ReverseArray = tempArray
        
End Function


Public Function GetEditRows_wSum(ByVal arrRows As Variant) As Variant
    Dim arrSummary() As Variant
    Dim row As Variant
    Dim i, j, k As Integer
    Dim arrActID(), arrWBS(), arrTmlMod(), arrTmlCod(), arrAux As Variant
    Dim strActID, strWBS, strTmlMod, strTmlCod As String
    
    'Si solo hay una actividad en la lista de actividades se sale de la función porque no va a haber ningún sumario que añadir
    If intActLastRow - rngRef.row = 1 Then GoTo ExitFunction
    'Se construyen 4 vectores con todos los valors de ActID, WBS, Timeline Mode y Timeline Code
    arrActID = Array(Range(rngActID.Offset(1), rngActID.Offset(intActLastRow - rngRef.row)).value)
    arrWBS = Array(Range(rngWBS.Offset(1), rngWBS.Offset(intActLastRow - rngRef.row)).value)
    arrTmlMod = Array(Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intActLastRow - rngRef.row)).value)
    arrTmlCod = Array(Range(rngTmlCod.Offset(1), rngTmlCod.Offset(intActLastRow - rngRef.row)).value)
    
    'Se inicializa el vector para las filas de los sumarios
    i = 0
    ReDim arrSummary(i)
    'Si el vector de entrada está vacío se sale de la función
    If IsEmpty(arrRows) Then Exit Function
    
    'Para cada fila en el vector de filas
    For Each row In arrRows
        'Guardamos el valor de ActID, WBS, TimelineMode y TimelineCode de la fila analizada en esta iteración
        strActID = IIf(IsError(arrActID(0)(row - rngRef.row, 1)), "", arrActID(0)(row - rngRef.row, 1))
        strWBS = arrWBS(0)(row - rngRef.row, 1)
        strTmlMod = arrTmlMod(0)(row - rngRef.row, 1)
        strTmlCod = arrTmlCod(0)(row - rngRef.row, 1)
        
        'Si no estamos en una fila de agrupación WBS entramos
        If Not strActID Like "WBS-*" Then
        'Se recorren todas las posiciones del vector ActID
            For j = 1 To UBound(arrActID(0))
                'Si el WBS es igual al correspondiente a la fila y el ActID contiene "WBS-*"
                If (arrWBS(0)(j, 1) = strWBS Or strWBS Like arrWBS(0)(j, 1) & ".*") And arrActID(0)(j, 1) Like "WBS-*" Then
                    'Se comprueba si la fila actual está ya incluida en el vector de sumarios y se sale si lo está
                    For k = 0 To UBound(arrSummary)
                        If arrSummary(k) = j + rngRef.row Then GoTo WBSinRowArray
                    Next
                    'Si se sale del bucle enterior por aquí se añade la fila al vector de sumarios
                    If i > 0 Then ReDim Preserve arrSummary(i)
                    arrSummary(i) = j + rngRef.row
                    i = i + 1
                End If
WBSinRowArray:
            Next
        End If
        
        'Si estamos en una fila agrupada en un Timeline entramos
        If Not IsEmpty(strTmlCod) And IsEmpty(strTmlMod) Then
            'Se recorren todas las posiciones del vector ActID
            For j = 1 To UBound(arrActID(0))
                'Si la fila tiene TimelineMode y el TimelineCode es igual al analizado en esta iteración
                If Not IsEmpty(arrTmlMod(0)(j, 1)) And arrTmlCod(0)(j, 1) = strTmlCod Then
                    'Se comprueba si la fila actual está ya incluida en el vector de sumarios y se sale si lo está
                    For k = 0 To UBound(arrSummary)
                        If arrSummary(k) = j + rngRef.row Then GoTo TMLinRowArray
                    Next
                    'Si se sale del bucle enterior por aquí se añade la fila al vector de sumarios
                    If i > 0 Then ReDim Preserve arrSummary(i)
                    arrSummary(i) = j + rngRef.row
                    i = i + 1
                End If
TMLinRowArray:
            Next
        End If
        
        'Si estamos en una fila de un Timeline entramos para añadir todas las actividades incluidas
        If Not IsEmpty(strTmlCod) And Not IsEmpty(strTmlMod) Then
            'Se recorren todas las posiciones del vector ActID
            For j = 1 To UBound(arrActID(0))
                'Si la fila no tiene TimelineMode y el TimelineCode es igual al analizado en esta iteración
                If IsEmpty(arrTmlMod(0)(j, 1)) And arrTmlCod(0)(j, 1) = strTmlCod Then
                    'Se comprueba si la fila actual está ya incluida en el vector de sumarios y se sale si lo está
                    For k = 0 To UBound(arrSummary)
                        If arrSummary(k) = j + rngRef.row Then GoTo TMLACTinRowArray
                    Next
                    'Si se sale del bucle enterior por aquí se añade la fila al vector de sumarios
                    If i > 0 Then ReDim Preserve arrSummary(i)
                    arrSummary(i) = j + rngRef.row
                    i = i + 1
                End If
TMLACTinRowArray:
            Next
        End If
    Next

    'Se incorporan los valores del vector sumarios al vector de filas original si el vector sumario tiene valores
    If Not IsEmpty(arrSummary(0)) Then
        For Each row In arrSummary
            ReDim Preserve arrRows(UBound(arrRows) + 1)
            arrRows(UBound(arrRows)) = row
        Next
    End If
ExitFunction:
    'Valor de salida ordenado
    GetEditRows_wSum = SortArrayAscendent(arrRows)
End Function

Public Function shapeExist(shName As String) As Boolean
    Dim shape As shape
    For Each shape In ActiveSheet.Shapes
        If shape.Name = shName Then
            shapeExist = True
            Exit Function
        End If
    Next
    shapeExist = False
End Function
