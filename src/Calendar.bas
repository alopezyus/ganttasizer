Attribute VB_Name = "Calendar"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public Sub CreateCalendar()
    Dim datCurr, datCalFinish, datBOmonth, datWeekStart As Date
    Dim intC, intY, intM, intYearOffset As Integer
    Dim strYear, strMonth, strQuarter As String
    Dim intDays, intRowsInserted As Integer
    
    'Comprobar si queda espacio para insertar calendario y si no insertar filas
        'y establecer offset para insertar valores de calendario
    intYearOffset = 0
    intRowsInserted = 0
    If intPeriod <= 6 Then
        intYearOffset = intYearOffset + 1
        If rngRef.row = 1 Then
            wsSch.Rows(1).Insert Shift:=xlDown
            intRowsInserted = intRowsInserted + 1
        End If
    End If
    If intPeriod <= 4 Then
        intYearOffset = intYearOffset + 1
        If rngRef.row = 2 Then
            wsSch.Rows(1).Insert Shift:=xlDown
            intRowsInserted = intRowsInserted + 1
        End If
    End If
    If intRowsInserted > 0 Then RenameShapes rngRef.row, intRowsInserted
    
    'Iinicialización de contadores y valor fecha
    intC = 0 'Counter
    intY = 0 'Year
    intM = 0 'Month
    
    'Posicionar primer y último día de calendario
    Select Case intPeriod
        Case 1, 2 'Diario
            datCurr = datPrjStart
            datCalFinish = datPrjFinish
        Case 3 'Semanal
            datCurr = datPrjStart - Weekday(datPrjStart, intWeekStart) + 1
            datCalFinish = datPrjFinish - (Weekday(datPrjFinish, intWeekStart) + 1) + 7
        Case 4 'Bi-semanal
            datCurr = datPrjStart - Weekday(datPrjStart, intWeekStart) + 1
            datCalFinish = datPrjFinish - (Weekday(datPrjFinish, intWeekStart) + 1) + _
                        IIf(Int((datPrjFinish + 1 - datPrjStart) / 14) = (datPrjFinish + 1 - datPrjStart) / 14, 7, 14)
        Case 5 'Mensual
            datCurr = DateSerial(Year(datPrjStart), Month(datPrjStart), 1)
            datCalFinish = DateSerial(Year(datPrjFinish), Month(datPrjFinish), 1)
        Case 6 'Trimestral
            datCurr = DateSerial(Year(datPrjStart), Month(datPrjStart), 1)
            datCalFinish = DateSerial(Year(datPrjFinish), Month(datPrjFinish), 1)
        Case 7 'Anual
            datCurr = DateSerial(Year(datPrjStart), 1, 1)
            datCalFinish = DateSerial(Year(datPrjFinish), 12, 31)
    End Select


    'UpdateProgressBar 0.1
    intDays = DateDiff("d", datCurr, datCalFinish)
    'Títulos para cada nivel de calendario
    
    'Bucle para recorrer todas las fechas distribuidas según período desde la fecha de inico hasta la de finalización
    Do While datCurr <= datCalFinish
        'Valor de año si es diferente al anterior y combinación de celdas
        If strYear <> format(datCurr, "yyyy") Or intC = 0 Then
            rngRef.Offset(-intYearOffset, intC) = format(datCurr, "yyyy")
            strYear = format(datCurr, "yyyy")
            If rngRef.Offset(-intYearOffset, intC) <> "" And intC > 0 Then
                With Range(rngRef.Offset(-intYearOffset, intY), rngRef.Offset(-intYearOffset, intC - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .NumberFormat = "General"
                End With
            End If
            intY = intC
        End If
        'Valor de trimestre si es diferente al anterior y combiación de celdas
        If (strQuarter <> "Q" & (Month(datCurr) - 1) \ 3 + 1 Or intC = 0) And intPeriod = 6 Then
            rngRef.Offset(-(intYearOffset - 1), intC) = "Q" & (Month(datCurr) - 1) \ 3 + 1
            strQuarter = "Q" & (Month(datCurr) - 1) \ 3 + 1
            If rngRef.Offset(-(intYearOffset - 1), intC) <> "" And intC > 0 Then
                With Range(rngRef.Offset(-(intYearOffset - 1), intM), rngRef.Offset(-(intYearOffset - 1), intC - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                End With
            End If
            intM = intC
        End If
        'Valor de mes si es diferente al anterior y combiación de celdas
        If (strMonth <> format(datCurr, "MMM") Or intC = 0) And intPeriod <= 5 And intPeriod <> 2 Then
            rngRef.Offset(-(intYearOffset - 1), intC) = format(datCurr, "MMM")
            strMonth = format(datCurr, "MMM")
            If rngRef.Offset(-(intYearOffset - 1), intC) <> "" And intC > 0 Then
                With Range(rngRef.Offset(-(intYearOffset - 1), intM), rngRef.Offset(-(intYearOffset - 1), intC - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                End With
            End If
            intM = intC
        End If
        
        'Valor de semana si es diferente al anterior y combiación de celdas
        'Se reutilizan las variable usadas para la actualización de las etiquetas de mes
        If (datWeekStart <> DateAdd("d", -(Weekday(datCurr, intWeekStart) - 1), datCurr) Or intC = 0) And intPeriod = 2 Then
            datWeekStart = DateAdd("d", -(Weekday(datCurr, intWeekStart) - 1), datCurr)
            rngRef.Offset(-(intYearOffset - 1), intC) = format(datWeekStart, "DD/MMM") & " - " & format(datWeekStart + 6, "DD/MMM")
            If rngRef.Offset(-(intYearOffset - 1), intC) <> "" And intC > 0 Then
                With Range(rngRef.Offset(-(intYearOffset - 1), intM), rngRef.Offset(-(intYearOffset - 1), intC - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                End With
            End If
            intM = intC
        End If
        
        'Valor correspondiente al período
        Select Case intPeriod
            Case 1 'Diario Calendario
                rngRef.Offset(0, intC) = format(datCurr, "dd")
                rngRef.Offset(0, intC).ColumnWidth = dblWidth
                rngRef.Offset(0, intC).HorizontalAlignment = xlCenter
                datCurr = datCurr + 1
            Case 2 'Diario Semana
                rngRef.Offset(0, intC) = Left(WeekdayName(Weekday(datCurr, intWeekStart), True, intWeekStart), 2)
                rngRef.Offset(0, intC).ColumnWidth = dblWidth
                rngRef.Offset(0, intC).HorizontalAlignment = xlCenter
                datCurr = datCurr + 1
            Case 3 'Semanal
                datBOmonth = DateSerial(Year(datCurr), Month(datCurr), 1)
                If datBOmonth > datCurr - 7 And datBOmonth < datCurr And intC > 0 Then
                    rngRef.Offset(0, intC + 1) = format(datCurr, "dd")
                    rngRef.Offset(0, intC + 1).ColumnWidth = dblWidth
                    rngRef.Offset(0, intC - 1).ColumnWidth = dblWidth * DateDiff("d", datCurr - 7, datBOmonth, intWeekStart) / 7
                    rngRef.Offset(0, intC).ColumnWidth = dblWidth * (1 - DateDiff("d", datCurr - 7, datBOmonth, intWeekStart) / 7)
                    'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                    If intC = 1 Then
                        Range(rngRef.Offset(-(intYearOffset - 1), 1), rngRef.Offset(-(intYearOffset - 1), 2)).Merge
                        Range(rngRef.Offset(-(intYearOffset), 1), rngRef.Offset(-(intYearOffset), 2)).Merge
                    End If
                    'Celdas combinadas semana
                    Range(rngRef.Offset(0, intC - 1), rngRef.Offset(0, intC)).Merge
                    'Corrección de intC para semanas a caballo entre dos meses en la primera fecha del calendario
                    intC = intC + IIf(intC = 1, 0, 1)
                Else
                    rngRef.Offset(0, intC) = format(datCurr, "dd")
                    rngRef.Offset(0, intC).ColumnWidth = dblWidth
                End If
                rngRef.Offset(0, intC).HorizontalAlignment = xlLeft
                datCurr = datCurr + 7
            Case 4 'Bi-semanal
                datBOmonth = DateSerial(Year(datCurr), Month(datCurr), 1)
                If datBOmonth > datCurr - 14 And datBOmonth < datCurr And intC > 0 Then
                    rngRef.Offset(0, intC + 1) = format(datCurr, "dd")
                    rngRef.Offset(0, intC + 1).ColumnWidth = dblWidth
                    rngRef.Offset(0, intC - 1).ColumnWidth = dblWidth * DateDiff("d", datCurr - 14, datBOmonth, intWeekStart) / 14
                    rngRef.Offset(0, intC).ColumnWidth = dblWidth * (1 - DateDiff("d", datCurr - 14, datBOmonth, intWeekStart) / 14)
                    'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                    If intC = 1 Then
                        Range(rngRef.Offset(-(intYearOffset - 1), 1), rngRef.Offset(-(intYearOffset - 1), 2)).Merge
                        Range(rngRef.Offset(-(intYearOffset), 1), rngRef.Offset(-(intYearOffset), 2)).Merge
                    End If
                    'Celdas combinadas semana
                    Range(rngRef.Offset(0, intC - 1), rngRef.Offset(0, intC)).Merge
                    'Corrección para semanas a caballo entre dos meses en la primera fecha del calendario
                    intC = intC + IIf(intC = 1, 0, 1)
                Else
                    rngRef.Offset(0, intC) = format(datCurr, "dd")
                    rngRef.Offset(0, intC).ColumnWidth = dblWidth
                End If
                rngRef.Offset(0, intC).HorizontalAlignment = xlLeft
                datCurr = datCurr + 14
            Case 5 'Mensual
                rngRef.Offset(0, intC).ColumnWidth = dblWidth
                rngRef.Offset(0, intC).HorizontalAlignment = xlCenter
                datCurr = DateSerial(Year(datCurr), Month(datCurr) + 1, day(datCurr))
            Case 6 'Trimestre
                rngRef.Offset(0, intC).ColumnWidth = dblWidth
                rngRef.Offset(0, intC).HorizontalAlignment = xlCenter
                datCurr = DateSerial(Year(datCurr), Month(datCurr) + 3, day(datCurr))
            Case 7 'Anual
                rngRef.Offset(0, intC).ColumnWidth = dblWidth
                rngRef.Offset(0, intC).HorizontalAlignment = xlCenter
                datCurr = DateSerial(Year(datCurr) + 1, Month(datCurr), day(datCurr))
        End Select
        intC = intC + 1
    'UpdateProgressBar IIf(intDays = 0, 0, (DateDiff("d", datCurr, datCalFinish) / intDays) * 0.8 + 0.1)
    Loop
    
    
    'Última combinación de celdas de año y mes
    If intPeriod = 5 Or intPeriod = 6 Then
        With Range(rngRef.Offset(-intYearOffset, intY), rngRef.Offset(-intYearOffset, intC - 1))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Merge
        End With
    Else
        With Range(rngRef.Offset(-intYearOffset, intY), rngRef.Offset(-intYearOffset, intC - 1))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        With Range(rngRef.Offset(-(intYearOffset - 1), intM), rngRef.Offset(-(intYearOffset - 1), intC - 1))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    End If
    
    'Correccón de la úlitma semana cuando se encuentra entre dos meses diferentes.
    If intPeriod = 3 Or intPeriod = 4 Then 'Semanal, Bi-semanal
        datBOmonth = DateSerial(Year(datCurr), Month(datCurr), 1)
        If datBOmonth > datCurr - IIf(intPeriod = 3, 7, 14) And datBOmonth < datCurr And intC > 0 Then
            rngRef.Offset(-1, intC) = format(datCurr, "mmm")
            rngRef.Offset(0, intC - 1).ColumnWidth = dblWidth * DateDiff("d", datCurr - IIf(intPeriod = 3, 7, 14), datBOmonth, intWeekStart) / IIf(intPeriod = 3, 7, 14)
            rngRef.Offset(0, intC).ColumnWidth = dblWidth * (1 - DateDiff("d", datCurr - IIf(intPeriod = 3, 7, 14), datBOmonth, intWeekStart) / IIf(intPeriod = 3, 7, 14))
            'Celdas combinadas semana y año
            Range(rngRef.Offset(-intYearOffset, intC - 1), rngRef.Offset(-intYearOffset, intC)).Merge
            Range(rngRef.Offset(0, intC - 1), rngRef.Offset(0, intC)).Merge
            'Corrección de intC para semanas a caballo entre dos meses
            intC = intC + 1
            rngRef.Offset(0, intC).HorizontalAlignment = xlLeft
        End If
    End If

    'Dibujar bordes del calendario
    CalendarBorders Range(rngRef.Offset(IIf(intPeriod = 7, 0, IIf(intPeriod >= 5, -1, -2)), 0), rngRef.Offset(0, intC - 1))
    
    GreyNonWorking
    
    'UpdateProgressBar 1
End Sub

Public Sub CalendarBottom()
    Dim i, intCalLastRow, intCalLastCol, intShpRow  As Integer
    Dim d As shape
    
    With Range(Cells(rngRef.row, intFirstCol), Cells(rngRef.row, intLastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    intCalLastRow = intActLastRow + 1
    intCalLastCol = CalLastColumn

    For i = 0 To intCalFirstRowOffset
        wsSch.Range(Cells(rngRef.row - i, 1), Cells(rngRef.row - i, intCalLastCol)).Copy
        wsSch.Cells(intCalLastRow + i, 1).Select
        ActiveSheet.Paste
        With wsSch.Range(Cells(intCalLastRow + i, rngPeriod.Column + 1), Cells(intCalLastRow + i, intCalLastCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Application.CutCopyMode = False
    Next
    
    For Each d In wsSch.Shapes
        If UBound(Split(d.Name, "_")) >= 2 Then
            intShpRow = CInt(Split(d.Name, "_")(2))
            On Error GoTo ErrorDeleteShape
            If intShpRow < intActLastRow And d.Top + d.Height > rngRef.Offset(intActLastRow + 1 - rngRef.row).Top Then d.Delete
        End If
NextShape:
    Next
    
    Exit Sub
    
ErrorDeleteShape:
Call RefreshRibbon
On Error GoTo 0
Resume NextShape
End Sub

Public Sub CalendarBorders(rngBorders As Range)
    rngBorders.Borders(xlDiagonalDown).LineStyle = xlNone
    rngBorders.Borders(xlDiagonalUp).LineStyle = xlNone
    With rngBorders.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rngBorders.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rngBorders.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rngBorders.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rngBorders.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rngBorders.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub ClearCalendar()
    
    If intLastCol >= rngRef.Column Then
        With Range(rngRef.Offset(-intCalFirstRowOffset, 0), rngRef.Offset(intActLastRow - rngRef.row + (intCalFirstRowOffset + 1), intLastCol - rngRef.Column))
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .UnMerge
            .ClearContents
            
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            .Interior.Color = xlNone
        End With
    End If
    'UpdateProgressBar 0.5
    
    'Elminar calendario parte inferior
    With wsSch.Cells(intActLastRow + 1, 1).EntireRow
        .ClearContents
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    'UpdateProgressBar 1

End Sub

Public Sub GreyNonWorking()
    Dim i As Integer
    Dim datRef As Date
      
    If intPeriod > 2 Then Exit Sub
    
    WeekCalendar
    
    i = 0
    For datRef = datPrjStart To datPrjFinish
        If Not WorkingDay(datRef) Then
            wsSch.Range(rngRef.Offset(0, i), rngRef.Offset(intActLastRow - rngRef.row, i)).Interior.Color = RGB(242, 242, 242)
        End If
        
        i = i + 1
    Next
    
End Sub
