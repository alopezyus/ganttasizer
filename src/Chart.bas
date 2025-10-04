Attribute VB_Name = "Chart"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Dim datStart, datFinish, datBLStart, datBLFinish, datResume As Variant
Dim strRef, strAct, strLabel As String
Dim datStartSum, datFinishSum, datBLStartSum, datBLFinishSum, datResumeSum As Variant
Dim strRefSum As String
Dim intSum, intRow, intRowSum As Integer
Dim varMilStyleRow, varProgress, varBdgUnt, varRmgUnt, varTotDur, varRmgDur, varFloat, varProgressSum, varBdgUntSum, varRmgUntSum, varTotDurSum, varRmgDurSum As Variant
Dim sh As shape
Dim dblProgress, dblBdgUnits, dblRmgUnits, dblShpHgt As Double
Dim intRowsInserted As Integer

Dim arrData() As Variant
Dim arrAct As Variant

Public booFloatBar As Boolean

'Construcción del vector de actividades
Private Sub SetActArray(Optional ByVal intRowUpdate As Variant = Empty)
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux(), arrShp() As Variant
    Dim row As Variant
    Dim i, j As Integer
    SetPrjVar
    'Se establece la dimensión del vector: número de filas
    intDim = ActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 18
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngActStyle.Offset(1), rngActStyle.Offset(intDim)).value, _
                    Range(rngShpHgt.Offset(1), rngShpHgt.Offset(intDim)).value, _
                    Range(rngLabPos.Offset(1), rngLabPos.Offset(intDim)).value, _
                    Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value, _
                    Range(rngTmlCod.Offset(1), rngTmlCod.Offset(intDim)).value, _
                    Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngDesc.Offset(1), rngDesc.Offset(intDim)).value, _
                    Range(rngStart.Offset(1), rngStart.Offset(intDim)).value, _
                    Range(rngFinish.Offset(1), rngFinish.Offset(intDim)).value, _
                    Range(rngStartBL.Offset(1), rngStartBL.Offset(intDim)).value, _
                    Range(rngFinishBL.Offset(1), rngFinishBL.Offset(intDim)).value, _
                    Range(rngResume.Offset(1), rngResume.Offset(intDim)).value, _
                    Range(rngProgress.Offset(1), rngProgress.Offset(intDim)).value, _
                    Range(rngBdgUnt.Offset(1), rngBdgUnt.Offset(intDim)).value, _
                    Range(rngRmgUnt.Offset(1), rngRmgUnt.Offset(intDim)).value, _
                    Range(rngTotDur.Offset(1), rngTotDur.Offset(intDim)).value, _
                    Range(rngRmgDur.Offset(1), rngRmgDur.Offset(intDim)).value, _
                    Range(rngFloat.Offset(1), rngFloat.Offset(intDim)).value, _
                    Range(rngWBS.Offset(1), rngWBS.Offset(intDim)).value)

    'Se dimensiona el vector que contendrá la información final al número de filas y un vector auxiliar al número de campos
    ReDim arrData(intDim - 1)
    ReDim arrShp(21) 'Ver referencias en función ShapeTypeToArrayRef
    
    For i = 0 To UBound(arrShp)
        arrShp(i) = False
    Next
    
    'Para cada fila
    For i = 0 To intDim - 1
        ReDim arrAux(intFields + 3)
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
        arrAux(19) = IIf(IsEmpty(intRowUpdate), True, False)
        'Valor del campo fila del Timeline en actividades sumario
        If Not IsEmpty(arrAux(3)) Then
            arrAux(20) = rngRef.row + 1 + i
        ElseIf i > 0 Then
            If arrAux(4) = arrData(i - 1)(4) Then arrAux(20) = arrData(i - 1)(20)
        End If
        'Redimensionar el campo de formas modificadas como un vector
        arrAux(21) = arrShp
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux
    Next
    
    If Not IsEmpty(intRowUpdate) Then
        For Each row In intRowUpdate
            arrData(row - rngRef.row - 1)(19) = True
        Next
    End If
    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Activity/Milestone Style
    '1 --> Shape Height
    '2 --> Label Position
    '3 --> Timeline Mode
    '4 --> Timeline Code
    '5 --> Activity ID
    '6 --> Description
    '7 --> Start
    '8 --> Finish
    '9 --> Start BL
    '10 --> Finish BL
    '11 --> Resume
    '12 --> Progress
    '13 --> Budget Unit
    '14 --> Remaining Unit
    '15 --> Total Duration
    '16 --> Remaining Duration
    '17 --> Float
    '18 --> WBS
    'Dimensiones extra
    '19 --> Actividad actualizable segun vector intRowUpd
    '20 --> Actividad sumario para actividades dentro de un Timeline
    '21 --> Vector de formas modificadas en la ejecución de CreateChart
End Sub

Private Sub SetActVar()
    Dim intRefRow As Integer
        intRefRow = rngRef.row
        
        datStart = IIf(IsError(arrAct(7)), Empty, arrAct(7))
        datFinish = IIf(IsError(arrAct(8)), Empty, arrAct(8))
        datBLStart = IIf(IsError(arrAct(9)), Empty, arrAct(9))
        datBLFinish = IIf(IsError(arrAct(10)), Empty, arrAct(10))
        datResume = IIf(IsError(arrAct(11)), Empty, arrAct(11))
        varProgress = IIf(IsError(arrAct(12)), 0, arrAct(12))
        varBdgUnt = IIf(IsError(arrAct(13)), 0, arrAct(13))
        varRmgUnt = IIf(IsError(arrAct(14)), 0, arrAct(14))
        varTotDur = IIf(IsError(arrAct(15)), 0, arrAct(15))
        varRmgDur = IIf(IsError(arrAct(16)), 0, arrAct(16))
        varFloat = arrAct(17)
        
        strRef = arrAct(4)
        varMilStyleRow = IIf(arrAct(0) = "WINDOW", "WINDOW", Trim(Left(arrAct(0), 2)))
        
        If IsNumeric(arrAct(1)) And arrAct(1) > 0 Then
            dblShpHgt = arrAct(1)
        Else: dblShpHgt = dblBarHgt
        End If
        
        'Cálculo del posicionamiento de la barra
        dblBarPos = (1 - dblShpHgt * IIf(xl_BlBar, 1.5, 1)) / 2 'IIf(IsDate(datBLStart) Or IsDate(datBLFinish), 1.5, 1)) / 2
        
        strLabel = arrAct(2)
        strAct = arrAct(6)
        
        intRowSum = 1
        intSum = 0
        Select Case arrAct(3)
        Case "MIL"
            intSum = 1
        Case "ACT"
            intSum = 2
        Case "SUM"
            intSum = 3
        End Select
End Sub

Private Sub SetSumVar(intRowSum As Integer)
    Dim intRefRow As Integer
        intRefRow = intRowSum - intRowsInserted
        
        datStartSum = IIf(IsError(arrData(intRefRow)(7)), Empty, arrData(intRefRow)(7))
        datFinishSum = IIf(IsError(arrData(intRefRow)(8)), Empty, arrData(intRefRow)(8))
        datBLStartSum = IIf(IsError(arrData(intRefRow)(9)), Empty, arrData(intRefRow)(9))
        datBLFinishSum = IIf(IsError(arrData(intRefRow)(10)), Empty, arrData(intRefRow)(10))
        datResumeSum = IIf(IsError(arrData(intRefRow)(11)), Empty, arrData(intRefRow)(11))
        varProgressSum = IIf(IsError(arrData(intRefRow)(12)), 0, arrData(intRefRow)(12))
        varBdgUntSum = IIf(IsError(arrData(intRefRow)(13)), 0, arrData(intRefRow)(13))
        varRmgUntSum = IIf(IsError(arrData(intRefRow)(14)), 0, arrData(intRefRow)(14))
        varTotDurSum = IIf(IsError(arrData(intRefRow)(15)), 0, arrData(intRefRow)(15))
        varRmgDurSum = IIf(IsError(arrData(intRefRow)(16)), 0, arrData(intRefRow)(16))
        
        strRefSum = arrData(intRefRow)(4)
        varMilStyleRow = Trim(Left(arrData(intRefRow)(0), 2))
        
        If IsNumeric(arrData(intRefRow)(1)) And arrData(intRefRow)(1) > 0 Then
            dblShpHgt = arrData(intRefRow)(1)
        Else: dblShpHgt = dblBarHgt
        End If
        'Cálculo del posicionamiento de la barra
        dblBarPos = (1 - dblShpHgt * IIf(xl_BlBar, 1.5, 1)) / 2 'IIf(IsDate(datBLStart) Or IsDate(datBLFinish), 1.5, 1)) / 2
        
        strLabel = arrData(intRefRow)(2)
        strAct = arrData(intRefRow)(6)
End Sub

Private Sub SetSumDates()
    Dim intRefRow As Integer
        intRefRow = rngRef.row
        
        arrAct(7) = datStart
        arrAct(8) = datFinish
        arrAct(9) = datBLStart
        arrAct(10) = datBLFinish
        
        rngStart.Offset(intRow - intRefRow) = datStart
        rngFinish.Offset(intRow - intRefRow) = datFinish
        rngStartBL.Offset(intRow - intRefRow) = datBLStart
        rngFinishBL.Offset(intRow - intRefRow) = datBLFinish
End Sub

'Eliminar todas las formas
Public Sub ClearChart(Optional ByVal intRowInsert As Variant = Empty)
    Dim d As shape
    Dim strSum As String
    Dim intRowLastTml, intRowShape As Integer
    Dim intRowUpdVal As Variant
    Dim strShpDef As String
    
        For Each d In wsSch.Shapes
            If d.Name Like "VB_*" Then
                If IsEmpty(intRowInsert) Then
                    d.Delete
                ElseIf d.Name Like "VB_*_*" Then
                    If CInt(Split(d.Name, "_")(2)) >= intRowInsert Then d.Delete
                End If
            End If
        Next d
    
    'Elminar calendario parte inferior
    With wsSch.Range(Cells(intActLastRow + 1, 1), Cells(intActLastRow + 1 + intCalFirstRowOffset, 1)).EntireRow
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
    
End Sub

Public Sub ClearShapes()
    Dim shp As shape
    Dim arrShapeName As Variant
    Dim intArrRef, intShpRef As Integer
    
    
    For Each shp In wsSch.Shapes
        If shp.Name Like "VB_*_*" Then
            arrShapeName = Split(shp.Name, "_")
            intArrRef = CInt(arrShapeName(2)) - rngRef.row - 1
            intShpRef = ShapeTypeToArrayRef(arrShapeName(1))
            If intArrRef > UBound(arrData) Or intArrRef < 0 Then
                shp.Delete
            ElseIf Not arrData(intArrRef)(21)(intShpRef) And arrData(intArrRef)(19) Then
                shp.Delete
            End If
        End If
    Next
End Sub

Public Sub ClearShapeByPos(dblTop, dblBottom As Double)
    Dim shp As shape
    Dim dblShapeMid As Double
    For Each shp In wsSch.Shapes
        dblShapeMid = shp.Top + shp.Height / 2
        If dblShapeMid < dblBottom And dblShapeMid > dblTop And shp.Name Like "VB_*" Then
            shp.Delete
        End If
    Next
End Sub

'Public Sub ClearShapeByRow()
'    Dim shp As shape
'    Dim dblShapeMid As Double
'    Dim intRow As Integer
'
'    For Each shp In wsSch.Shapes
'        dblShapeMid = shp.Top + shp.Height / 2
'        If shp.Name Like "VB_*_*" Then intRow = CInt(Split(shp.Name, "_")(2))
'        If intRow = 27 Then
'            Debug.Print dblShapeMid
'        End If
'        If dblShapeMid <= Cells(intRow, 1).Top And dblShapeMid >= Cells(intRow + 1, 1).Top Then
'            shp.Delete
'        End If
'    Next
'End Sub

Public Sub RenameShapes(intRowInsert, intRowsInserted As Integer)
    Dim shp As shape
    Dim strName As Variant
    Dim strType, strRow As String
    
    For Each shp In wsSch.Shapes
        If shp.Name Like "VB_*_*" Then
            strName = Split(shp.Name, "_")
            If CInt(strName(2)) >= intRowInsert Then
                strType = strName(0) & "_" & strName(1) & "_"
                strRow = format(CInt(strName(2)) + intRowsInserted, "00000")
                shp.Name = strType & strRow
            End If
        End If
    Next
End Sub

Private Function ShapeTypeToArrayRef(ByVal strShapeType As String) As Integer
    Select Case strShapeType
    Case "ACT" 'Barra Actual
        ShapeTypeToArrayRef = 0
    Case "MLA" 'Milestone Actual
        ShapeTypeToArrayRef = 1
    Case "REM" 'Barra Remaining
        ShapeTypeToArrayRef = 2
    Case "MLR" 'Milestone Remaining
        ShapeTypeToArrayRef = 3
    Case "BL0" 'Barra Base Line
        ShapeTypeToArrayRef = 4
    Case "BM0" 'Milestone Base Line
        ShapeTypeToArrayRef = 5
    Case "PRA" 'Progreso Actual
        ShapeTypeToArrayRef = 6
    Case "PRR" 'Progreso Remaining
        ShapeTypeToArrayRef = 7
    Case "FLT" 'Float
        ShapeTypeToArrayRef = 8
    Case "SMA" 'Sumario Milestone Actual
        ShapeTypeToArrayRef = 9
    Case "SMR" 'Sumario Milestone Remaining
        ShapeTypeToArrayRef = 10
    Case "SM0" 'Sumario Milestone Base Line
        ShapeTypeToArrayRef = 11
    Case "SAA" 'Sumario Barra Actual
        ShapeTypeToArrayRef = 12
    Case "SAR" 'Sumario Barra Remaining
        ShapeTypeToArrayRef = 13
    Case "SA0" 'Sumario Base Line
        ShapeTypeToArrayRef = 14
    Case "WIN" 'Ventana
        ShapeTypeToArrayRef = 15
    Case "DESC" 'Etiqueta descripción
        ShapeTypeToArrayRef = 16
    Case "DLIN" 'Línea de unión a descripción
        ShapeTypeToArrayRef = 17
    Case "DUR" 'Etiqueta duración
        ShapeTypeToArrayRef = 18
    Case "WDES" 'Etiqueta descripción de Ventanas
        ShapeTypeToArrayRef = 19
    Case "WLIN" 'Línea de unión a descripción de Ventnas
        ShapeTypeToArrayRef = 20
    Case "SDES" 'Etiqueta descripción en Sumarios
        ShapeTypeToArrayRef = 21
    End Select
End Function

Private Sub UpdateShapeArray(ByVal strShapeType As String, ByVal intRow As Integer)
    arrData(intRow - rngRef.row - 1 - intRowsInserted)(21)(ShapeTypeToArrayRef(strShapeType)) = True
End Sub

'Crear gráfico de barras
Public Sub CreateChart(Optional ByVal intRowUpdate As Variant = Empty)
'On Error GoTo errHandler
    Dim dblLeft, dblTop, dblWidth, dblHeight As Double
    Dim booCutoff As Boolean
    Dim col As Integer
    Dim intRowUpdVal As Variant
    Dim booHideTml As Boolean
    Dim intOutlineLevel As Integer
    Dim arrWBS() As Variant
    Dim strWBS As Variant
    Dim intWBS As Integer
    Dim booWBSgroup As Boolean
    
    'Se quitan todas las agrupaciones. Por si no hubiera agrupaciones previas se agrega una al final para evitar el error.
'    If IsEmpty(intRowUpdate) Then
'        Cells(intActLastRow + 5, 1).Rows.group
'        wsSch.Rows.Ungroup
'        Cells.EntireRow.Hidden = False
'    End If
    
    'Inserción y eliminación de filas para Timelines
    InsertRows intRowUpdate
    intActLastRow = ActLastRow

    SetActArray intRowUpdate
    intRow = rngRef.row + 1
    intRowsInserted = 0
    intWBS = 0
    booHideTml = False
    booWBSgroup = False
    'Inicio bucle para recorrer todas las filas
    For Each arrAct In arrData
        If Not arrAct(19) Then GoTo NextIteration

        'Instrucción de seguridad si intRow>intActLastRow
        If intRow > intActLastRow Then GoTo NextIteration

        'Para actualizaciones parciales, desagrupar actividades que requieran actualización en timelines
        If Not booHideTml Then
            If Rows(intRow).OutlineLevel > 1 And Rows(intRow).Hidden = True And Len(arrAct(4)) > 0 Then
                'Rows(intRow).Hidden = False
                wsSch.Rows(intRow).ShowDetail = True
                booHideTml = True
            Else: booHideTml = False
            End If
        End If
        
        'Desagrupar actividades dentro de WBS agrupado solo si es un calculo completo
        If IsEmpty(intRowUpdate) Then
        If Rows(intRow + 1).OutlineLevel > 1 And wsSch.Rows(intRow).ShowDetail = False And arrAct(5) Like "WBS-*" Then
            wsSch.Rows(intRow).ShowDetail = True
            
            ReDim Preserve arrWBS(intWBS)
            arrWBS(intWBS) = arrAct(18)
            booWBSgroup = True
            intWBS = intWBS + 1
        End If
        End If

        If xl_UpdRow = True And Not wsSch.Cells(intRow, 1).RowHeight = dblHeightStd Then
            wsSch.Cells(intRow, 1).RowHeight = dblHeightStd
        End If
        booCutoff = False
        SetActVar
        
        'Disabled in Free Edition
        If Not intEdition = 1 Then
            'Si la fila define una ventana
            If varMilStyleRow = "WINDOW" Then
                If (IsDate(datStart) And datStart > 0) Or (IsDate(datFinish) And datFinish > 0) Then
                    Dim intWindowRows As Integer
                    Dim rngWindowSetup As Range
                    Set rngWindowSetup = wsSch.Cells(intRow, 1)
                    CalculateDuration
                    dblLeft = BarPosX(IIf(datStart = 0, False, True), False, True, False)
                    If IsEmpty(datFinish) Then
                        dblWidth = 0
                    Else
                        dblWidth = BarPosX(False, False, True, False) - dblLeft
                    End If
                    dblTop = rngWindowSetup.Top
                    intWindowRows = IIf(arrAct(1) > 0, arrAct(1), intActLastRow - intRow) + 1
                    dblHeight = rngWindowSetup.Offset(intWindowRows).Top - dblTop
                    InsertWindow msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_WIN_", lngWindowColor
                End If
                GoTo NextIteration
            End If
        End If
                
        Select Case intSum
        Case 1
            If IsEmpty(strRef) Then GoTo NextIteration
            'Altura de la fila
            wsSch.Cells(intRow - 1, 1).RowHeight = 15 * maxLabelPos
            CreateSumMil (IsEmpty(intRowUpdate))
            If Not varProgress = dblProgress Then rngProgress.Offset(intRow - rngRef.row, 0) = dblProgress
            If Not varBdgUnt = dblBdgUnits Then rngBdgUnt.Offset(intRow - rngRef.row, 0) = dblBdgUnits
            If Not varRmgUnt = dblRmgUnits Then rngRmgUnt.Offset(intRow - rngRef.row, 0) = dblRmgUnits
            SetActVar
            If UCase(varMilStyleRow) = "NO" Then
                CalculateDuration
                GoTo NextIteration
            End If
            'Barra de Progreso: Si hay fecha de inicio y de fin (es una actividad) y hay progreso y se dibuja la barra de progreso y no es un summary
            If IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 And varProgress > 0 And xl_PrgBar Then
                dblTop = BarPosY(True, False)
                dblHeight = BarPosY(False, False) - dblTop
                dblTop = dblTop + dblHeight / 4
                dblHeight = dblHeight / 2
                If datStart <= datCutoff And IsDate(datCutoff) Then
                    dblLeft = BarPosX(True, False, False, False, True)
                    dblWidth = BarPosX(False, False, False, False, True) - dblLeft
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRA_", lngPrgBarColor, True
                    BarStyle
                End If
                If (datFinish > datCutoff And IsDate(datCutoff)) Or Not IsDate(datCutoff) Then
                    dblLeft = BarPosX(True, False, True, False, True)
                    dblWidth = BarPosX(False, False, True, False, True) - dblLeft
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRR_", lngPrgBarColor, True
                    BarStyle
                End If
            End If
        Case 2
            If IsEmpty(strRef) Then GoTo NextIteration
            'Altura de la fila
            wsSch.Cells(intRow - 1, 1).RowHeight = 15 * maxLabelPos
            CreateSumAct (IsEmpty(intRowUpdate))
            If Not varProgress = dblProgress Then rngProgress.Offset(intRow - rngRef.row, 0) = dblProgress
            If Not varBdgUnt = dblBdgUnits Then rngBdgUnt.Offset(intRow - rngRef.row, 0) = dblBdgUnits
            If Not varRmgUnt = dblRmgUnits Then rngRmgUnt.Offset(intRow - rngRef.row, 0) = dblRmgUnits
            'Insertar descripción
            SetActVar
            If IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 And datStart <= datFinish And (datCutoff < datFinish Or xl_lblActuals) Then
                dblLeft = BarPosX(True, False, True, booCutoff)
                dblTop = BarPosY(True, False)
                dblWidth = BarPosX(False, False, True, booCutoff) - dblLeft
                dblHeight = BarPosY(False, False) - dblTop 'IIf(intSum = 1, intSumHgt, BarPosY(False, False) - dblTop)
                strAct = arrAct(6) 'Descripción de la actividad     wsSch.Cells(intRow, rngDesc.Column)
                InsertDesc dblLeft + dblWidth, dblTop, dblWidth, dblHeight, False
            End If
            'Se salta a la primera actividad sumarizada para evitar que se dibuje otra barra sumario.
            CalculateDuration
            GoTo NextIteration
        Case 3
            If IsEmpty(strRef) Then GoTo NextIteration
            If IsEmpty(intRowUpdate) Then
                If Rows(intRow).OutlineLevel > 1 Then wsSch.Cells(intRow, 1).Rows.Ungroup
            End If
            'Altura de la fila
            wsSch.Cells(intRow - 1, 1).RowHeight = 0
            CreateSum (IsEmpty(intRowUpdate))
            If Not varProgress = dblProgress Then rngProgress.Offset(intRow - rngRef.row, 0) = dblProgress
            If Not varBdgUnt = dblBdgUnits Then rngBdgUnt.Offset(intRow - rngRef.row, 0) = dblBdgUnits
            If Not varRmgUnt = dblRmgUnits Then rngRmgUnt.Offset(intRow - rngRef.row, 0) = dblRmgUnits
            SetActVar
            'Establecer estilo de barra Summary
            rngActStyle.Offset(intRow - rngRef.row) = 8
            varMilStyleRow = 8
        End Select
        
        CalculateDuration
        
        If varMilStyleRow = 8 And IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 Then
            dblLeft = BarPosX(True, False, True, booCutoff)
            dblTop = BarPosY(True, False)
            dblWidth = BarPosX(False, False, True, booCutoff) - dblLeft
            dblHeight = BarPosY(False, False) - dblTop 'IIf(intSum = 1, intSumHgt, BarPosY(False, False) - dblTop)
            InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, IIf(datCutoff >= datFinish, "VB_ACT_", "VB_REM_"), _
                        IIf(datCutoff >= datFinish, lngActBarColor, lngRemBarColor), IIf(intSum = 1, True, False) 'IIf(datCutoff >= datFinish, lngActBarColor, lngRemBarColor)
            BarStyle
            If datCutoff < datFinish Or xl_lblActuals Then InsertDesc dblLeft + dblWidth, dblTop, dblWidth, dblHeight, False
            If datStart >= datCutoff Then InsertDuration dblLeft, dblTop, dblWidth, dblHeight
            GoTo ChkBLandPrg
        End If
        
        'Actual: sólo si hay fecha de corte definida
        If IsDate(datCutoff) Then
            'Actividad: Si hay fecha de inicio y de fin y la fecha de inicio es anterior al corte
            If IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 And datStart <= datFinish And datCutoff >= datStart Then
                'Cálculo de posiciones
                dblLeft = BarPosX(True, False, False, booCutoff)
                dblTop = BarPosY(True, False)
                dblWidth = BarPosX(False, False, False, booCutoff) - dblLeft
                dblHeight = BarPosY(False, False) - dblTop 'IIf(intSum = 1, intSumHgt, BarPosY(False, False) - dblTop)
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_ACT_", lngActBarColor, IIf(intSum = 1, True, False)
                BarStyle
                If datFinish <= datCutoff And booLabActuals Then InsertDesc dblLeft + IIf(strLabel = "0M", 0, dblWidth), dblTop, dblWidth, dblHeight, False
                booCutoff = True
                
            'Hito: Si hay fecha de inicio (o de terminación) y es anterior a la fecha de corte
            ElseIf (IsDate(datStart) And datStart > 0 And datCutoff > datStart) Or (IsDate(datFinish) And datFinish > 0 And datCutoff >= datFinish) Or _
                    (IsDate(datStart) And datStart > 0 And Not IsDate(datFinish) And datCutoff = datStart) Then
                dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblLeft = BarPosX(IIf(IsDate(datFinish) And datFinish > 0, False, True), False, True, booCutoff) - dblWidth / 2
                dblTop = BarPosY(True, False)
                InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_MLA_", lngActBarColor, IIf(intSum = 1, True, False)
                MilStyle
                If booLabActuals Then InsertDesc dblLeft + dblWidth, dblTop, dblWidth, dblHeight, False
                'Sh.Placement = IIf(Len(strRef) > 0, xlMoveAndSize, xlMove)
            End If
        End If
        
        'Remanente
        'Actividad: Si hay fecha de inicio y de fin y la fecha de fin es posterior al corte
        If IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 And datStart <= datFinish And datCutoff < datFinish Then
            dblLeft = BarPosX(True, False, True, booCutoff)
            dblTop = BarPosY(True, False)
            dblWidth = BarPosX(False, False, True, booCutoff) - dblLeft
            dblHeight = BarPosY(False, False) - dblTop 'IIf(intSum = 1, intSumHgt, BarPosY(False, False) - dblTop)
            InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_REM_", lngRemBarColor, IIf(intSum = 1, True, False)
            BarStyle
            InsertDesc dblLeft + IIf(strLabel = "0M", 0, dblWidth), dblTop, dblWidth, dblHeight, False
            If datStart >= datCutoff Then InsertDuration dblLeft, dblTop, dblWidth, dblHeight

        'Hito: Si hay fecha de inicio (o de terminación) y es posterior a la fecha de corte
        ElseIf (IsDate(datStart) And datStart > 0 And datCutoff < datStart) _
                Or (IsDate(datFinish) And datFinish > 0 And datCutoff < datFinish) Then
            dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
            dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
            dblLeft = BarPosX(IIf(IsDate(datFinish) And datFinish > 0, False, True), False, True, booCutoff) - dblWidth / 2
            dblTop = BarPosY(True, False)
            InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_MLR_", lngMilColor, IIf(intSum = 1, True, False)
            MilStyle
            'Sh.Placement = IIf(Len(strRef) > 0, xlMoveAndSize, xlMove)
            InsertDesc dblLeft + dblWidth, dblTop, dblWidth, dblHeight, False
        End If
        
ChkBLandPrg:
        'Linea Base
        'Actividad: Si hay fecha de inicio y de fin
        If xl_BlBar Then
            If IsDate(datBLStart) And datBLStart > 0 And IsDate(datBLFinish) And datBLFinish > 0 And datBLStart <= datBLFinish Then
                dblLeft = BarPosX(True, True, True, booCutoff)
                dblTop = BarPosY(True, True)
                dblWidth = BarPosX(False, True, True, booCutoff) - dblLeft
                dblHeight = BarPosY(False, True) - dblTop 'IIf(intSum = 1, intSumHgt, BarPosY(False, True) - dblTop)
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_BL0_", lngBLBarColor, IIf(intSum = 1 Or (varMilStyleRow >= 8 Or varMilStyleRow <= 10), True, False)
                
            'Hito: Si hay fecha de inicio (o de terminación)
            ElseIf (IsDate(datBLStart) And datBLStart > 0) Or (IsDate(datBLFinish) And datBLFinish > 0) Then
                dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                dblLeft = BarPosX(IIf(IsDate(datBLFinish) And datBLFinish > 0, False, True), True, True, booCutoff) - dblWidth / 2
                dblTop = BarPosY(True, True) '- dblHeight / 3
                InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_BM0_", lngBLBarColor, IIf(intSum = 1, True, False)
                MilStyle
                'Sh.Placement = IIf(Len(strRef) > 0, xlMoveAndSize, xlMove)
            End If
        End If

        'Barra de Progreso: Si hay fecha de inicio y de fin (es una actividad) y hay progreso y se dibuja la barra de progreso y no es un summary
        If IsDate(datStart) And IsDate(datFinish) And datStart > 0 And datFinish > 0 And IsNumeric(varProgress) And varProgress > 0 And xl_PrgBar And Not intSum = 1 Then
            dblTop = BarPosY(True, False)
            dblHeight = BarPosY(False, False) - dblTop
            dblTop = dblTop + dblHeight / 4
            dblHeight = dblHeight / 2
            If datStart <= datCutoff And IsDate(datCutoff) Then
                dblLeft = BarPosX(True, False, False, False, True)
                dblWidth = BarPosX(False, False, False, False, True) - dblLeft
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRA_", lngPrgBarColor
                BarStyle
            End If
            If (datFinish > datCutoff And IsDate(datCutoff)) Or Not IsDate(datCutoff) Then
                dblLeft = BarPosX(True, False, True, False, True)
                dblWidth = BarPosX(False, False, True, False, True)
                dblWidth = IIf(dblLeft < dblWidth, dblWidth - dblLeft, 0)
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRR_", lngPrgBarColor
                BarStyle
            End If
        End If
        
        'Barra de holgura: si hay holgura calculada y se ha seleccionado dibujar la barra de progreso y la holgura es mayor que cero
        If IsNumeric(varFloat) And booFloatBar And xl_FltBar Then
            If varFloat > 0 Then
                dblLeft = BarPosX(True, False, False, False, False, True)
                dblWidth = BarPosX(False, False, False, False, False, True) - dblLeft
                dblTop = BarPosY(True, False)
                dblHeight = (BarPosY(False, False) - dblTop) / 8
                dblTop = dblTop + 7 * dblHeight
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_FLT_", lngFltBarColor
            End If
        End If
            
        'Para actualizaciones parciales, volver a agrupar actividades que requieran actualización
        If booHideTml Then
            If intRow = intActLastRow Or arrData(intRow - rngRef.row - 2)(4) = arrAct(4) Then
                'Rows(intRow).Hidden = True
                wsSch.Rows(intRow).ShowDetail = False
                booHideTml = False
            End If
        End If
        
NextIteration:
        intRow = intRow + 1
    Next
    
    'Formato de las agrupaciones
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    
    If IsEmpty(intRowUpdate) Then UpdateProgressBar 0.8
    ClearShapes
    
    If Not IsEmpty(intRowUpdate) Then
        On Error Resume Next
        wsSch.Shapes("VB_CUTOFF").ZOrder msoBringToFront
        GoTo exSub
    End If
        
        
    On Error Resume Next
    wsSch.Shapes("VB_CUTOFF").Delete
    On Error GoTo 0
    If datCutoff >= datPrjStart And datCutoff <= datPrjFinish Then
        InsertCutoff
    End If
    
    'Ocultar todas las agrupaciones de timelines
    Dim i As Integer
    For i = rngRef.row + 1 To intActLastRow
        If Rows(i + 1).OutlineLevel > Rows(i).OutlineLevel And Rows(i + 1).Hidden = False And Len(Cells(i, rngTmlMod.Column)) > 0 And Len(Cells(i, rngTmlCod.Column)) > 0 Then
            wsSch.Rows(i).ShowDetail = False
        End If
    Next
    
    'Agrupar WBS agrupados al inicio si es un calculo completo
    If IsEmpty(intRowUpdate) And booWBSgroup Then
        arrData = ReverseArray(arrData)
        intRow = intActLastRow
        For Each arrAct In arrData
            If arrAct(5) Like "WBS-*" Then
                For Each strWBS In arrWBS
                    If arrAct(18) = strWBS And wsSch.Rows(intRow).ShowDetail = True Then
                        wsSch.Rows(intRow).ShowDetail = False
                        GoTo nextWBS
                    End If
                Next
            End If
nextWBS:
            intRow = intRow - 1
        Next
    End If

    ChartLines
    CalendarBottom
    rngRef.Select
    
exSub:
    Exit Sub
    
errHandler:
    CustomMsgBox "An error has occurred during execution." + vbCrLf + "Please revise.", vbCritical + vbOKOnly, error:=True
    Resume exSub
End Sub


'Cálculo de la duración
Private Sub CalculateDuration()
    Dim intTotDur, intRmgDur As Integer
    'Cálculo de Duración Total
    If Not IsDate(datStart) Or Not IsDate(datFinish) Then
        intTotDur = 0
    Else:
        intTotDur = DateDiffCal(datStart, datFinish) + IIf(WorkingDay(datStart), 1, 0)
    End If
    If intTotDur < 0 Then intTotDur = 0
    If Not varTotDur = intTotDur Then wsSch.Cells(intRow, rngTotDur.Column) = intTotDur
    
    'Cálculo de duración remanente
    'Si es un hito (duracion total 0) o la fecha fin es anterior al cutoff --> Duración remanente = 0
    If intTotDur = 0 Or (IsDate(datFinish) And datFinish <= datCutoff) Then
        intRmgDur = 0
    'Si la fecha de inicio es mayor que el cutoff --> Duración remanente = Duración total
    ElseIf datStart > datCutoff Then
        intRmgDur = intTotDur
    'Si la fecha de reanudación está entre el corte y la fecha de terminacion --> Duración remanente = Fecha fin - Fecha reanudación
    ElseIf IsDate(datResume) And datResume > datCutoff And datResume <= datFinish Then
        intRmgDur = DateDiffCal(datResume, datFinish) + IIf(WorkingDay(datResume), 1, 0)
    'Duración remanente = Fecha fin - Cutoff
    Else:
        intRmgDur = DateDiffCal(datCutoff, datFinish)
    End If
    If intRmgDur < 0 Then intRmgDur = 0
    If Not varRmgDur = intRmgDur Then wsSch.Cells(intRow, rngRmgDur.Column) = intRmgDur
End Sub
'Crear sumario en modo SUM
Private Sub CreateSum(Optional booGroup As Boolean = True)
    Dim dblLeft, dblTop, dblWidth, dblHeight As Double
    Dim booCutoff As Boolean
    Dim datStartRef As Date 'Fecha de inicio de referencia para la sumarización
    Dim datStartArr(), datFinishArr(), datBLStartArr(), datBLFinishArr() As Variant
    Dim i As Integer
    
    intRowSum = intRow + 1
    SetSumVar (intRowSum - rngRef.row - 1)
    i = 0
    dblProgress = 0
    dblBdgUnits = 0
    dblRmgUnits = 0
    
    datStartRef = IIf(IsDate(datStartSum), datStartSum, datFinishSum)
        
    Do While strRefSum = strRef
        'Si encuentra otra línea identificada como summary y la misma referencia, quita summary
        If Not IsEmpty(arrData(intRowSum - rngRef.row - 1)(3)) Then wsSch.Cells(intRowSum, rngTmlMod.Column) = ""
        
        'Establecer fecha de inicio de la actividad sumario y su BL
        datStart = IIf(IsDate(datStartSum), datStartSum, datFinishSum)
        datBLStart = IIf(IsDate(datBLStartSum), datBLStartSum, datBLFinishSum)
        'Establecer fecha de finalización de la actividad sumario y su BL
        datFinish = datFinishSum
        datBLFinish = datBLFinishSum
        'Establecer fecha de reinicio de la actividad sumario
        datResume = datResumeSum
        
        'Llenar vectores de fechas
        ReDim Preserve datStartArr(i)
        datStartArr(i) = datStartSum
        ReDim Preserve datFinishArr(i)
        datFinishArr(i) = datFinishSum
        ReDim Preserve datBLStartArr(i)
        datBLStartArr(i) = IIf(datBLStartSum = 0, Empty, datBLStartSum)
        ReDim Preserve datBLFinishArr(i)
        datBLFinishArr(i) = datBLFinishSum
                
        intRowSum = intRowSum + 1
        If intRowSum > intActLastRow Then GoTo exitLoop
        'Suma de unidades
        If IsNumeric(varProgressSum) And IsNumeric(varBdgUntSum) Then dblProgress = dblProgress + varBdgUntSum * varProgressSum
        If IsNumeric(varBdgUntSum) Then dblBdgUnits = dblBdgUnits + varBdgUntSum
        If IsNumeric(varRmgUntSum) Then dblRmgUnits = dblRmgUnits + varRmgUntSum
        SetSumVar (intRowSum - rngRef.row - 1)
        i = i + 1
     Loop
exitLoop:
     If dblBdgUnits = 0 Then
        dblProgress = 0
     Else: dblProgress = dblProgress / dblBdgUnits
     End If
     
On Error GoTo errHandler
    'Extraer fechas de los vectores
    datStart = datStartArr(0)
    datFinish = datFinishArr(0)
    datBLStart = datBLStartArr(0)
    datBLFinish = datBLFinishArr(0)
    For i = 0 To UBound(datStartArr)
           datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or IsEmpty(datStart), datStartArr(i), datStart)
           datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), datFinishArr(i) + 1, datStart)
           datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or Not IsDate(datStart), datStartArr(i), datStart)
           datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), IIf(IsEmpty(datFinishArr(i)), Empty, datFinishArr(i) + 1), datStart)
           datFinish = IIf(datFinish < datFinishArr(i) And datFinishArr(i) > 0, datFinishArr(i), datFinish)
           
           datBLStart = IIf((datBLStart > datBLStartArr(i) And datBLStartArr(i) > 0) Or Not IsDate(datBLStart), datBLStartArr(i), datBLStart)
           datBLStart = IIf((datBLStart > datBLFinishArr(i) And datBLFinishArr(i) > 0) Or Not IsDate(datBLStart), IIf(IsEmpty(datBLFinishArr(i)), Empty, datBLFinishArr(i) + 1), datBLStart)
           datBLFinish = IIf(datBLFinish < datBLFinishArr(i) And datBLFinishArr(i) > 0, datBLFinishArr(i), datBLFinish)
    Next

    'Guardar fechas para la actividad sumario en la hoja excel
    SetSumDates
    'Agrupar
    Dim intRowGroup As Integer
    'If booGroup Then Range(Cells(intRow + 1, 1), Cells(intRowSum - 1, 1)).Rows.group
    If booGroup Then
        For intRowGroup = intRow + 1 To intRowSum - 1
            GroupWBS arrData(intRowGroup - rngRef.row - 1)(5), arrData(intRowGroup - rngRef.row - 1)(18), intRowGroup, True
        Next
    End If

    Exit Sub
errHandler:
    If Err.Number <> 9 Then CustomMsgBox Err.Number, vbExclamation + vbOKOnly
End Sub

'Crear hitos para sumarios en modo MIL
Private Sub CreateSumMil(Optional booGroup As Boolean = True)
    Dim dblLeft, dblTop, dblWidth, dblHeight As Double
    Dim datStartArr(), datFinishArr(), datBLStartArr(), datBLFinishArr() As Variant
    Dim i As Integer
    
    intRowSum = intRow + 1
    SetSumVar (intRowSum - rngRef.row - 1)
    i = 0
    dblProgress = 0
    dblBdgUnits = 0
    dblRmgUnits = 0
            
    Do While strRefSum = strRef
        'Si encuentra otra línea identificada como summary y la misma referencia, quita summary
        If Not IsEmpty(arrData(intRowSum - rngRef.row - 1)(3)) Then wsSch.Cells(intRowSum, rngTmlMod.Column) = ""
        
        'Establecer fecha de inicio de la actividad sumario y su BL
        datStart = IIf(IsDate(datStartSum), datStartSum, datFinishSum)
        datBLStart = IIf(IsDate(datBLStartSum), datBLStartSum, datBLFinishSum)
        'Establecer fecha de finalización de la actividad sumario y su BL
        datFinish = datFinishSum
        datBLFinish = datBLFinishSum
        'Establecer fecha de reinicio de la actividad sumario
        datResume = datResumeSum
        
        'Llenar vectores de fechas
        ReDim Preserve datStartArr(i)
        datStartArr(i) = datStartSum
        ReDim Preserve datFinishArr(i)
        datFinishArr(i) = datFinishSum
        ReDim Preserve datBLStartArr(i)
        datBLStartArr(i) = datBLStartSum
        ReDim Preserve datBLFinishArr(i)
        datBLFinishArr(i) = datBLFinishSum
        
        If Not UCase(varMilStyleRow) = "NO" Then
            'Actual
            'Hito: Si hay fecha de inicio (o de terminación) y es anterior a la fecha de corte
            If ((IsDate(datFinishSum) And datFinishSum > 0 And datCutoff >= datFinishSum) Or (IsDate(datStartSum) And datStartSum > 0 And Not IsDate(datFinishSum) And datCutoff >= datStartSum)) Then
                dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblLeft = BarPosX(IIf(IsDate(datFinishSum) And datFinishSum > 0, False, True), False, True, False) - dblWidth / 2
                dblTop = BarPosY(True, False)
                InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SMA_", lngActBarColor
                MilStyle
                If booLabActuals Then InsertDesc dblLeft + dblWidth / 2, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
            End If
    
            'Remanente
            'Hito: Si hay fecha de inicio (o de terminación) y es posterior a la fecha de corte
            If ((IsDate(datFinishSum) And datFinishSum > 0 And datCutoff < datFinishSum) Or (IsDate(datStartSum) And datStartSum > 0 And Not IsDate(datFinishSum) And datCutoff < datStartSum)) Then
                dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblLeft = BarPosX(IIf(IsDate(datFinishSum) And datFinishSum > 0, False, True), False, True, False) - dblWidth / 2
                dblTop = BarPosY(True, False)
                InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SMR_", lngMilColor
                MilStyle
                'InsertDesc dblLeft, wsSch.Cells(intRow - 1, 1).Top, Len(strAct) * 15, wsSch.Cells(intRow - 1, 1).Height, , intRowSum
                InsertDesc dblLeft + dblWidth / 2, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
            End If
            
            'Linea Base
            'Hito: Si hay fecha de inicio (o de terminación)
            If xl_BlBar Then
                If ((IsDate(datBLStartSum) And datBLStartSum > 0) Or (IsDate(datBLFinishSum) And datBLFinishSum > 0)) Then
                    dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                    dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                    dblLeft = BarPosX(IIf(IsDate(datBLFinishSum) And datBLFinishSum > 0, False, True), True, True, False) - dblWidth / 2
                    dblTop = BarPosY(True, True) '- dblHeight / 3
                    InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SM0_", lngBLBarColor
                    MilStyle
                End If
            End If
        End If
        intRowSum = intRowSum + 1
        If intRowSum > intActLastRow Then GoTo exitLoop
        'Suma de unidades
        If IsNumeric(varProgressSum) And IsNumeric(varBdgUntSum) Then dblProgress = dblProgress + varBdgUntSum * varProgressSum
        If IsNumeric(varBdgUntSum) Then dblBdgUnits = dblBdgUnits + varBdgUntSum
        If IsNumeric(varRmgUntSum) Then dblRmgUnits = dblRmgUnits + varRmgUntSum
        SetSumVar (intRowSum - rngRef.row - 1)
        i = i + 1
    Loop
exitLoop:
     If dblBdgUnits = 0 Then
        dblProgress = 0
     Else: dblProgress = dblProgress / dblBdgUnits
     End If
     
On Error GoTo errHandler
    'Extraer fechas de los vectores
    datStart = datStartArr(0)
    datFinish = datFinishArr(0)
    datBLStart = datBLStartArr(0)
    datBLFinish = datBLFinishArr(0)
    For i = 0 To UBound(datStartArr)
           datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or IsEmpty(datStart), datStartArr(i), datStart)
           datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), datFinishArr(i) + 1, datStart)
           datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or Not IsDate(datStart), datStartArr(i), datStart)
           datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), IIf(IsEmpty(datFinishArr(i)), Empty, datFinishArr(i) + 1), datStart)
           datFinish = IIf(datFinish < datFinishArr(i) And datFinishArr(i) > 0, datFinishArr(i), datFinish)
           
           datBLStart = IIf((datBLStart > datBLStartArr(i) And datBLStartArr(i) > 0) Or Not IsDate(datBLStart), datBLStartArr(i), datBLStart)
           datBLStart = IIf((datBLStart > datBLFinishArr(i) And datBLFinishArr(i) > 0) Or Not IsDate(datBLStart), IIf(IsEmpty(datBLFinishArr(i)), Empty, datBLFinishArr(i) + 1), datBLStart)
           datBLFinish = IIf(datBLFinish < datBLFinishArr(i) And datBLFinishArr(i) > 0, datBLFinishArr(i), datBLFinish)
    Next
    'Guardar fechas para la actividad sumario en la hoja excel
    SetSumDates
    'Agrupar
    'If booGroup Then Range(Cells(intRow + 1, 1), Cells(intRowSum - 1, 1)).Rows.group
    Dim intRowGroup As Integer
    If booGroup Then
        For intRowGroup = intRow + 1 To intRowSum - 1
            GroupWBS arrData(intRowGroup - rngRef.row - 1)(5), arrData(intRowGroup - rngRef.row - 1)(18), intRowGroup, True
        Next
    End If
    'Guardar estilo de la barra sumario
    varMilStyleRow = Trim(Left(arrAct(0), 2))

    Exit Sub
errHandler:
    If Err.Number <> 9 Then CustomMsgBox Err.Number, vbExclamation + vbOKOnly
End Sub

'Crear actividades para sumarios en modo ACT
Private Sub CreateSumAct(Optional booGroup As Boolean = True)
    Dim dblLeft, dblTop, dblWidth, dblHeight As Double
    Dim booCutoff As Boolean
    Dim datStartRef As Date 'Fecha de inicio de referencia para la sumarización
    Dim datStartArr(), datFinishArr(), datBLStartArr(), datBLFinishArr() As Variant
    Dim i As Integer
    
    intRowSum = intRow + 1
    SetSumVar (intRowSum - rngRef.row - 1)
    i = 0
    dblProgress = 0
    dblBdgUnits = 0
    dblRmgUnits = 0
    
    datStartRef = IIf(IsDate(datStartSum), datStartSum, datFinishSum)
    
    Do While strRefSum = strRef
        'Si encuentra otra línea identificada como summary y la misma referencia, quita summary
        If Not IsEmpty(arrData(intRowSum - rngRef.row - 1)(3)) Then wsSch.Cells(intRowSum, rngTmlMod.Column) = ""
        
        booCutoff = False
        
        'Establecer fecha de inicio de la actividad sumario y su BL
        datStart = IIf(IsDate(datStartSum), datStartSum, datFinishSum)
        datBLStart = IIf(IsDate(datBLStartSum), datBLStartSum, datBLFinishSum)
        'Establecer fecha de finalización de la actividad sumario y su BL
        datFinish = datFinishSum
        datBLFinish = datBLFinishSum
        'Establecer fecha de reinicio de la actividad sumario
        datResume = datResumeSum
        'Establecer progreso de la actividad sumario
        varProgress = varProgressSum
        
        'Llenar vectores de fechas
        ReDim Preserve datStartArr(i)
        datStartArr(i) = datStartSum
        ReDim Preserve datFinishArr(i)
        datFinishArr(i) = datFinishSum
        ReDim Preserve datBLStartArr(i)
        datBLStartArr(i) = IIf(datBLStartSum = 0, Empty, datBLStartSum)
        ReDim Preserve datBLFinishArr(i)
        datBLFinishArr(i) = datBLFinishSum
        
        If Not UCase(varMilStyleRow) = "NO" Then
            'Actual: sólo si hay fecha de corte definida
            If IsDate(datCutoff) Then
                'Actividad: Si hay fecha de inicio y de fin y la fecha de inicio es anterior al corte
                If IsDate(datStartSum) And datStartSum > 0 And IsDate(datFinishSum) And datFinishSum > 0 And datStartSum <= datFinishSum And datCutoff >= datStartSum Then
                    'Cálculo de posiciones
                    dblLeft = BarPosX(True, False, False, booCutoff)
                    dblTop = BarPosY(True, False)
                    dblWidth = BarPosX(False, False, False, booCutoff) - dblLeft
                    dblHeight = BarPosY(False, False) - dblTop
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_SAA_", lngActBarColor
                    BarStyle
                    If datFinish <= datCutoff And booLabActuals Then InsertDesc dblLeft + dblWidth, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
                    booCutoff = True
                    
                'Hito: Si hay fecha de inicio (o de terminación) y es anterior a la fecha de corte
                ElseIf ((IsDate(datStartSum) And datStartSum > 0 And datCutoff > datStartSum) Or (IsDate(datFinishSum) And datFinishSum > 0 And datCutoff >= datFinishSum)) Or _
                        (IsDate(datStartSum) And datStartSum > 0 And Not IsDate(datFinishSum) And datCutoff = datStartSum) Then
                    dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
                    dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
                    dblLeft = BarPosX(IIf(IsDate(datFinishSum) And datFinishSum > 0, False, True), False, True, booCutoff) - dblWidth / 2
                    dblTop = BarPosY(True, False)
                    InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SMA_", lngActBarColor
                    MilStyle
                    If booLabActuals Then InsertDesc dblLeft + dblWidth, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
                End If
            End If
        
            'Remanente
            'Actividad: Si hay fecha de inicio y de fin y la fecha de fin es posterior al corte
            If IsDate(datStartSum) And datStartSum > 0 And IsDate(datFinishSum) And datFinishSum > 0 And datStartSum <= datFinishSum And datCutoff < datFinishSum Then
                dblLeft = BarPosX(True, False, True, booCutoff)
                dblTop = BarPosY(True, False)
                dblWidth = BarPosX(False, False, True, booCutoff) - dblLeft
                dblHeight = BarPosY(False, False) - dblTop
                InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_SAR_", lngRemBarColor
                BarStyle
                InsertDesc dblLeft + dblWidth, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
            
            'Hito: Si hay fecha de inicio (o de terminación) y es posterior a la fecha de corte
            ElseIf ((IsDate(datStartSum) And datStartSum > 0 And datCutoff < datStartSum) Or (IsDate(datFinishSum) And datFinishSum > 0 And datCutoff < datFinishSum)) Then
                dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt
                dblLeft = BarPosX(IIf(IsDate(datFinishSum) And datFinishSum > 0, False, True), False, True, booCutoff) - dblWidth / 2
                dblTop = BarPosY(True, False)
                InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SMR_", lngMilColor
                MilStyle
                InsertDesc dblLeft + dblWidth, wsSch.Cells(intRow, 1).Top, dblWidth, dblHeight, , intRowSum
            End If
                    
            'Linea Base
            'Actividad: Si hay fecha de inicio y de fin
            If xl_BlBar Then
                If IsDate(datBLStartSum) And datBLStartSum > 0 And IsDate(datBLFinishSum) And datBLFinishSum > 0 And datBLStartSum <= datBLFinishSum Then
                    dblLeft = BarPosX(True, True, True, booCutoff)
                    dblTop = BarPosY(True, True)
                    dblWidth = BarPosX(False, True, True, booCutoff) - dblLeft
                    dblHeight = BarPosY(False, True) - dblTop
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_SA0_", lngBLBarColor
                    
                'Hito: Si hay fecha de inicio (o de terminación)
                ElseIf ((IsDate(datBLStartSum) And datBLStartSum > 0) Or (IsDate(datBLFinishSum) And datBLFinishSum > 0)) Then
                    dblWidth = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                    dblHeight = wsSch.Cells(intRow, 1).Height * dblShpHgt * 0.5
                    dblLeft = BarPosX(IIf(IsDate(datBLFinishSum) And datBLFinishSum > 0, False, True), True, True, booCutoff) - dblWidth / 2
                    dblTop = BarPosY(True, True) '- dblHeight / 3
                    InsertShape msoShapeDiamond, dblLeft, dblTop, dblWidth, dblHeight, "VB_SM0_", lngBLBarColor
                    MilStyle
                End If
            End If
            
            'Barra de Progreso: Si hay fecha de inicio y de fin (es una actividad) y hay progreso y se dibuja la barra de progreso y no es un summary
            If IsDate(datStartSum) And IsDate(datFinishSum) And datStartSum > 0 And datFinishSum > 0 And varProgress > 0 And xl_PrgBar And Not intSum = 1 Then
                dblTop = BarPosY(True, False)
                dblHeight = BarPosY(False, False) - dblTop
                dblTop = dblTop + dblHeight / 4
                dblHeight = dblHeight / 2
                If datStartSum <= datCutoff And IsDate(datCutoff) Then
                    dblLeft = BarPosX(True, False, False, False, True)
                    dblWidth = BarPosX(False, False, False, False, True) - dblLeft
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRA_", lngPrgBarColor
                    BarStyle
                End If
                If (datFinishSum > datCutoff And IsDate(datCutoff)) Or Not IsDate(datCutoff) Then
                    dblLeft = BarPosX(True, False, True, False, True)
                    dblWidth = BarPosX(False, False, True, False, True) - dblLeft
                    InsertShape msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight, "VB_PRR_", lngPrgBarColor
                    BarStyle
                End If
            End If

        End If
        intRowSum = intRowSum + 1
        If intRowSum > intActLastRow Then GoTo exitLoop
        'Suma de unidades
        If IsNumeric(varProgressSum) And IsNumeric(varBdgUntSum) Then dblProgress = dblProgress + varBdgUntSum * varProgressSum
        If IsNumeric(varBdgUntSum) Then dblBdgUnits = dblBdgUnits + varBdgUntSum
        If IsNumeric(varRmgUntSum) Then dblRmgUnits = dblRmgUnits + varRmgUntSum
        SetSumVar (intRowSum - rngRef.row - 1)
        i = i + 1
     Loop
exitLoop:
     If dblBdgUnits = 0 Then
        dblProgress = 0
     Else: dblProgress = dblProgress / dblBdgUnits
     End If
     
On Error GoTo errHandler
        datStart = datStartArr(0)
        datFinish = datFinishArr(0)
        datBLStart = datBLStartArr(0)
        datBLFinish = datBLFinishArr(0)
        For i = 0 To UBound(datStartArr)
               datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or IsEmpty(datStart), datStartArr(i), datStart)
               datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), datFinishArr(i) + 1, datStart)
               datStart = IIf((datStart > datStartArr(i) And datStartArr(i) > 0) Or Not IsDate(datStart), datStartArr(i), datStart)
               datStart = IIf((datStart > datFinishArr(i) And datFinishArr(i) > 0) Or Not IsDate(datStart), IIf(IsEmpty(datFinishArr(i)), Empty, datFinishArr(i) + 1), datStart)
               datFinish = IIf(datFinish < datFinishArr(i) And datFinishArr(i) > 0, datFinishArr(i), datFinish)
               
               datBLStart = IIf((datBLStart > datBLStartArr(i) And datBLStartArr(i) > 0) Or Not IsDate(datBLStart), datBLStartArr(i), datBLStart)
               datBLStart = IIf((datBLStart > datBLFinishArr(i) And datBLFinishArr(i) > 0) Or Not IsDate(datBLStart), IIf(IsEmpty(datBLFinishArr(i)), Empty, datBLFinishArr(i) + 1), datBLStart)
               datBLFinish = IIf(datBLFinish < datBLFinishArr(i) And datBLFinishArr(i) > 0, datBLFinishArr(i), datBLFinish)
        Next
    
        'Guardar fechas para la actividad sumario en la hoja excel
        SetSumDates
        'Agrupar
        'If booGroup Then Range(Cells(intRow + 1, 1), Cells(intRowSum - 1, 1)).Rows.group
        Dim intRowGroup As Integer
        If booGroup Then
            For intRowGroup = intRow + 1 To intRowSum - 1
                GroupWBS arrData(intRowGroup - rngRef.row - 1)(5), arrData(intRowGroup - rngRef.row - 1)(18), intRowGroup, True
            Next
        End If

    Exit Sub
errHandler:
    If Err.Number <> 9 Then CustomMsgBox Err.Number, vbExclamation + vbOKOnly
End Sub

'Establecer posición horizontal
Private Function BarPosX(booStart As Boolean, booBL As Boolean, booRem As Boolean, booCutoff As Boolean, Optional booPrgBar As Boolean = False, Optional booFloat As Boolean = False) As Double
    Dim datRef, datD As Date
    Dim dblDif, dblPosX As Double
    Dim intC As Integer
    Dim booResume As Boolean
    Dim intDaysActBar, intDaysRemBar, intDaysPrgBar As Integer

    booResume = False
    If IsDate(datResume) And IsDate(datCutoff) And datResume > datCutoff And datResume <= datFinish Then booResume = True
    
    If booPrgBar Then
        'Calcular días de la barra Actual
        If datStart <= datCutoff And IsDate(datCutoff) Then
            If datFinish <= datCutoff Then
                intDaysActBar = DateDiff("d", datStart, datFinish) + 1
            Else: intDaysActBar = DateDiff("d", datStart, datCutoff) + 1
            End If
        Else: intDaysActBar = 0
        End If
        'Calcular día de la barra Remaining
        If (datFinish > datCutoff And IsDate(datCutoff)) Or Not IsDate(datCutoff) Then
            If booResume Then
                intDaysRemBar = DateDiff("d", datResume, datFinish) + 1
            ElseIf datStart < datCutoff And IsDate(datCutoff) Then
                intDaysRemBar = DateDiff("d", datCutoff, datFinish) + 1
            Else: intDaysRemBar = DateDiff("d", datStart, datFinish) + 1
            End If
        ElseIf datFinish <= datCutoff And IsDate(datCutoff) Then
            intDaysRemBar = 0
        End If
        'Calcular días progresados y fecha a la que se añaden los días
        intDaysPrgBar = Int((intDaysActBar + intDaysRemBar) * varProgress)
        If Not booRem Then
            datRef = datStart - IIf(Not booStart, 1, 0)
            If booStart Then
                intDaysPrgBar = 0
            ElseIf intDaysPrgBar > intDaysActBar Then
                intDaysPrgBar = intDaysActBar
            End If
        Else
            If booResume Then
                datRef = datResume - IIf(Not booStart, 1, 0) 'And intDaysPrgBar <= intDaysActBar
            Else: datRef = IIf(intDaysActBar > 0, datCutoff + IIf(booStart, 1, 0), datStart) - IIf(Not booStart, 1, 0)
            End If
            If booStart Then
                intDaysPrgBar = 0
            Else: intDaysPrgBar = intDaysPrgBar - intDaysActBar
                If intDaysPrgBar < 0 Then intDaysPrgBar = 0

            End If
        End If
        'Devolver fecha de referencia
        datRef = IIf(intDaysPrgBar = 0, datRef, DateAdd("d", intDaysPrgBar, datRef))
    ElseIf booFloat Then
        If booStart Then
            datRef = IIf(datFinish + 1 > datStart, datFinish + 1, datStart)
        Else
            datRef = DateAddCal(varFloat - 1, IIf(datFinish + 1 > datStart, datFinish + 1, datStart))
        End If
    ElseIf booBL Then
        If booStart Then
            datRef = datBLStart
        Else: datRef = datBLFinish
        End If
    ElseIf booRem Then
        If booStart Then
            datRef = IIf(booCutoff, IIf(booResume, datResume, datCutoff + 1), datStart)
        Else: datRef = datFinish
        End If
    Else
        If booStart Then
            datRef = datStart
        Else: datRef = IIf(datCutoff < datFinish, datCutoff, datFinish)
        End If
    End If
        
    dblDif = DateDiff("d", datChartStart, datRef, intWeekStart)
    datD = datChartStart
    Select Case intPeriod
        Case 1, 2 'Diario
            dblDif = dblDif + IIf(booStart, 0, 1)
        Case 3 'Semanal
            dblDif = 0
            datRef = datRef + 1 + IIf(booStart, -1, 0)
            Do While datRef > datD
                If Month(datD) = Month(datD + 7) Or day(datD + 7) = 1 Then
                    If datRef >= datD + 7 Then
                        dblDif = dblDif + 1
                    Else: dblDif = dblDif + DateDiff("d", datD, datRef, intWeekStart) / 7
                    End If
                Else
                    If datRef >= datD + 7 Then
                        'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                        dblDif = dblDif + 2
                    ElseIf datRef >= DateSerial(Year(datD), Month(datD) + 1, 1) Then
                        dblDif = dblDif + 1 + (datRef - DateSerial(Year(datD), Month(datD) + 1, 1)) / (datD + 7 - DateSerial(Year(datD), Month(datD) + 1, 1))
                    Else: dblDif = dblDif + (datRef - datD) / (DateSerial(Year(datD), Month(datD) + 1, 1) - datD)
                    End If
                End If
                datD = datD + 7
            Loop
        Case 4 'Bi-semanal
            dblDif = 0
            datRef = datRef + 1 + IIf(booStart, -1, 0)
            Do While datD < datRef
                If Month(datD) = Month(datD + 14) Or day(datD + 14) = 1 Then
                    If datRef >= datD + 14 Then
                        dblDif = dblDif + 1
                    Else: dblDif = dblDif + DateDiff("d", datD, datRef, intWeekStart) / 14
                    End If
                Else
                    If datRef >= datD + 14 Then
                        'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                        dblDif = dblDif + IIf(datD = datChartStart And booStart, 1, 2)
                    ElseIf datRef >= DateSerial(Year(datD), Month(datD) + 1, 1) Then
                        dblDif = dblDif + 1 + (datRef - DateSerial(Year(datD), Month(datD) + 1, 1)) / (datD + 14 - DateSerial(Year(datD), Month(datD) + 1, 1))
                    Else: dblDif = dblDif + (datRef - datD) / (DateSerial(Year(datD), Month(datD) + 1, 1) - datD)
                    End If
                End If
                datD = datD + 14
            Loop
        Case 5 'Mensual
            dblDif = Round(DateDiff("d", datChartStart, DateSerial(Year(datRef), Month(datRef), 1), intWeekStart) / 30.4, 0) + _
                        (DateDiff("d", DateSerial(Year(datRef), Month(datRef), 1), datRef, intWeekStart) + IIf(booStart, 0, 1)) / GetMonthDays(datRef)
        Case 6 'Trimestre
            dblDif = Round(DateDiff("d", datChartStart, DateSerial(Year(datRef), 3 * ((Month(datRef) - 1) \ 3) + 1, 1), intWeekStart) / 91.25, 0) + _
                        (DateDiff("d", DateSerial(Year(datRef), 3 * ((Month(datRef) - 1) \ 3) + 1, 1), datRef, intWeekStart) + IIf(booStart, 0, 1)) / GetQuarterDays(datRef)
        Case 7 'Anual
             dblDif = Round(DateDiff("d", datChartStart, DateSerial(Year(datRef), 1, 1)) / 365, 0) + _
                        (DateDiff("d", DateSerial(Year(datRef), 1, 1), datRef) + IIf(booStart, 0, 1)) / 365
    End Select
    
    'Offset programado en dos pasos para evitar bug en el offset cuando el inicio o el final de la actividad cae en una semana a caballo entre dos meses
    BarPosX = rngRef.Offset(1).Offset(, Int(dblDif)).Left + (dblDif - Int(dblDif)) * rngRef.Offset(1).Offset(, Int(dblDif)).Width
End Function

'Obtener fecha para selección de forma
Public Function GetDate(dblPosX As Double, Optional booStart As Boolean = False) As Date
    Dim intDays, intW, i As Integer
    Dim datCurr, datPre As Date
    Dim booFirstWeek2months As Boolean

    datCurr = datChartStart
    i = 0
    
    If dblPosX < rngRef.Left Then
        GetDate = datCurr
        Exit Function
    End If
    
    Select Case intPeriod
        Case 1, 2 'Diario
             Do While rngRef.Offset(0, i).Left <= dblPosX
                datPre = datCurr
                datCurr = datCurr + 1
                i = i + 1
             Loop
        Case 3, 4 'Semanal / Bi-semanal
            intW = IIf(intPeriod = 3, 7, 14)
            booFirstWeek2months = False
            Do While rngRef.Offset(0, i).Left < dblPosX
                If Month(datCurr) = Month(datCurr + intW) Or day(datCurr + intW) = 1 Then
                    datPre = datCurr
                    datCurr = datCurr + intW
                Else
                    If rngRef.Offset(0, i + 1).Left <= dblPosX Then
                        datPre = DateSerial(Year(datCurr), Month(datCurr) + 1, 1)
                        datCurr = datCurr + intW
                        i = i + 1 + IIf(i = 0, -1, 0)
                        booFirstWeek2months = IIf(i = 0, True, False)
                    Else
                        datPre = datCurr
                        datCurr = DateSerial(Year(datCurr), Month(datCurr) + 1, 1)
                    End If
                End If
                i = i + 1
            Loop
            If rngRef.Offset(0, i).Left = dblPosX And i = 1 And booFirstWeek2months Then
                i = i + 1
                datPre = datCurr
                datCurr = datCurr + intW
            End If
        Case 5 'Mensual
             Do While rngRef.Offset(0, i).Left <= dblPosX
                datPre = datCurr
                datCurr = DateSerial(Year(datCurr), Month(datCurr) + 1, 1)
                i = i + 1
             Loop
        Case 6 'Trimestral
             Do While rngRef.Offset(0, i).Left <= dblPosX
                datPre = datCurr
                datCurr = DateSerial(Year(datCurr), Month(datCurr) + 3, 1)
                i = i + 1
             Loop
        Case 7 'Anual
             Do While rngRef.Offset(0, i).Left <= dblPosX
                datPre = datCurr
                datCurr = DateSerial(Year(datCurr) + 1, 1, 1)
                i = i + 1
             Loop
    End Select
    
    GetDate = datPre + Round((datCurr - datPre) * (dblPosX - rngRef.Offset(0, i - 1).Left) / rngRef.Offset(0, i - 1).Width, 0) - IIf(booStart, 0, 1)
End Function

'Obtener fecha de barra de progreso seleccionada
Public Function GetPrgBarDate(ByVal intRow As Integer, booStart, booRem As Boolean)
    Dim datRef, datStart, datFinish, datResume As Date
    Dim varProgress As Variant
    Dim booResume As Boolean
    Dim intDaysActBar, intDaysRemBar, intDaysPrgBar As Integer

    datStart = rngStart.Offset(intRow - rngRef.row)
    datFinish = rngFinish.Offset(intRow - rngRef.row)
    datResume = rngResume.Offset(intRow - rngRef.row)
    varProgress = rngProgress.Offset(intRow - rngRef.row)
    
    booResume = False
    If IsDate(datResume) And IsDate(datCutoff) And datResume > datCutoff And datResume <= datFinish Then booResume = True
    
    'Calcular días de la barra Actual
    If datStart <= datCutoff And IsDate(datCutoff) Then
        If datFinish <= datCutoff Then
            intDaysActBar = DateDiff("d", datStart, datFinish)
        Else: intDaysActBar = DateDiff("d", datStart, datCutoff)
        End If
    Else: intDaysActBar = 0
    End If
    'Calcular día de la barra Remaining
    If (datFinish > datCutoff And IsDate(datCutoff)) Or Not IsDate(datCutoff) Then
        If booResume Then
            intDaysRemBar = DateDiff("d", datResume, datFinish)
        ElseIf datStart < datCutoff And IsDate(datCutoff) Then
            intDaysRemBar = DateDiff("d", datCutoff, datFinish)
        Else: intDaysRemBar = DateDiff("d", datStart, datFinish)
        End If
    ElseIf datFinish <= datCutoff And IsDate(datCutoff) Then
        intDaysRemBar = 0
    End If
    'Calcular días progresados y fecha a la que se añaden los días
    intDaysPrgBar = Round((intDaysActBar + intDaysRemBar) * varProgress, 0)
    If Not booRem Then
        datRef = datStart
        If booStart Then
            intDaysPrgBar = 0
        ElseIf intDaysPrgBar > intDaysActBar Then
            intDaysPrgBar = intDaysActBar
        End If
    Else
        datRef = IIf(booResume, datResume - IIf(Not booStart And intDaysPrgBar <= intDaysActBar, 1, 0), IIf(intDaysActBar > 0, datCutoff + IIf(booStart, 1, 0), datStart))
        If booStart Then
            intDaysPrgBar = 0
        Else: intDaysPrgBar = intDaysPrgBar - intDaysActBar
            If intDaysPrgBar < 0 Then intDaysPrgBar = 0

        End If
    End If
    'Devolver fecha de referencia
    GetPrgBarDate = IIf(intDaysPrgBar = 0, datRef, DateAdd("d", intDaysPrgBar, datRef))
End Function

'Establecer posición vertical
Private Function BarPosY(booTop As Boolean, booBL As Boolean) As Double
    
    BarPosY = wsSch.Cells(intRow, 1).Top + _
        wsSch.Cells(intRow, 1).RowHeight * (dblBarPos + IIf(booBL, dblShpHgt * IIf(booTop, 1, 1.5), IIf(booTop, 0, dblShpHgt)))
    
End Function

'Obtener posición Y teórica para selección de forma
Public Function GetPosY(intR As Integer, ByVal booTop As Boolean, ByVal booBL As Boolean) As Double
    Dim booArrow As Boolean
    intRow = intR
    
    'La forma es una flecha?
    If Not booBL Then
        Select Case Trim(Left(rngActStyle.Offset(intR - rngRef.row), 2))
        Case "8", "9", "10"
            booArrow = True
        Case Else
            booArrow = False
        End Select
    End If
    
    GetPosY = rngRef.Offset(intR - rngRef.row, 0).Top + _
        rngRef.Offset(intR - rngRef.row, 0).RowHeight * ((dblBarPos + IIf(booBL, dblShpHgt * IIf(booTop, 1, 1.5), IIf(booTop, 0, dblShpHgt))) + _
                                                            dblShpHgt / 2 * IIf(Not booArrow, 0, IIf(booTop, -1, 1)))
End Function

'Crear forma
Private Sub InsertShape(varShape As Variant, dblLeft, dblTop, dblWidth, dblHeight As Double, strName As String, ByVal lngColor As Long, Optional booSendBack As Boolean = False)
    Dim intR As Integer
    Dim strNameFull As String
    
    intR = IIf(strName Like "VB_S??_", intRowSum, intRow)
    strNameFull = strName & format(intR, "00000")
    
    On Error GoTo CreateShape
    Set sh = wsSch.Shapes(strNameFull)
    With sh
        If Not Round(.Left, 2) = Round(dblLeft, 2) Then .Left = dblLeft
        If Not Round(.Top, 2) = Round(dblTop, 2) Then .Top = dblTop
        If Not Round(.Width, 2) = Round(Abs(dblWidth), 2) Then .Width = Abs(dblWidth)
        If Not Round(.Height, 2) = Round(dblHeight, 2) Then .Height = dblHeight
    End With
    GoTo ShapeExists

CreateShape:
    'Crear forma
    Set sh = wsSch.Shapes.AddShape(varShape, dblLeft, dblTop, Abs(dblWidth), dblHeight)
    'Nombrar
    sh.Name = strNameFull
    
ShapeExists:
    On Error GoTo 0
    UpdateShapeArray Split(strName, "_")(1), intR
    With sh
        'Editar color de la forma
        If ((strName = "VB_REM_" Or strName = "VB_MLR_") Or (xl_SetActColor And (strName = "VB_ACT_" Or strName = "VB_MLA_"))) And _
            wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color <> 16777215 Then
            .Fill.ForeColor.RGB = wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        ElseIf ((strName = "VB_SMR_" Or strName = "VB_SAR_") Or (xl_SetActColor And (strName = "VB_SMA_" Or strName = "VB_SAA_"))) And _
            wsSch.Cells(intRowSum, rngActStyle.Column).DisplayFormat.Interior.Color <> 16777215 Then
            .Fill.ForeColor.RGB = wsSch.Cells(intRowSum, rngActStyle.Column).DisplayFormat.Interior.Color
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        Else:
            .Fill.ForeColor.RGB = lngColor
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        End If
        .Visible = True
        'Enviar al fondo
        If booSendBack Then .ZOrder msoSendToBack
    End With
End Sub

'Crear Etiqueta Descripcion
Private Sub InsertDesc(dblLeft, dblTop, dblWidth, dblHeight As Double, Optional booSum As Boolean = True, Optional ByVal intR As Integer = Empty)
    Dim strLabInput As String
     Dim strNameFull As String
     
    If UCase(strLabel) = "NO" Then Exit Sub
    
    strAct = IIf(IsError(strAct), "", strAct)
    strLabInput = IIf(booLabDesc, strAct, " ") & IIf(booLabDesc And booLabFinish And Len(strAct) > 0, ", ", " ") & _
                    IIf(booLabFinish, format(IIf(datFinish > 0, datFinish, datStart), IIf(datFinish > 0, rngFinish, rngStart).Offset(intRow - rngRef.row).NumberFormat), " ")
    strLabInput = Trim(strLabInput)
    If Len(strLabInput) = 0 Then Exit Sub
   
    strNameFull = IIf(intR = Empty, "VB_DESC_", "VB_SDES_") & format(IIf(intR = Empty, intRow, intRowSum), "00000")
    
    On Error GoTo CreateDescription
    Set sh = wsSch.Shapes(strNameFull)
    With sh
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        If Not Round(.Left, 2) = Round(dblLeft, 2) Then .Left = dblLeft
        If Not Round(.Top, 2) = Round(dblTop, 2) Then .Top = dblTop
    End With
    GoTo DescriptionExists
    
CreateDescription:
    'Crear forma
    Set sh = wsSch.Shapes.AddShape(msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight)
    'Nombrar
    sh.Name = strNameFull
    
DescriptionExists:
    On Error GoTo 0
    UpdateShapeArray IIf(intR = Empty, "DESC", "SDES"), IIf(intR = Empty, intRow, intRowSum)
    
    With sh
        'Texto
        .TextFrame2.TextRange.Characters.text = strLabInput
        .Placement = xlMoveAndSize
        .TextFrame.AutoSize = True
        .Visible = True
        'Enviar al fondo
      .ZOrder msoBringToFront
        'Editar color de la forma
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        If strLabel = "0M" Then
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .Width = dblWidth
        End If
    
    'Posicionar etiqueta
        If booSum And Len(Trim(strLabInput)) > 0 Then
            .TextFrame2.TextRange.Font.UnderlineStyle = msoUnderlineSingleLine
            .Placement = xlMoveAndSize
            'wsSch.Cells(intRow - 1, 1).RowHeight = IIf(wsSch.Cells(intRow, 1).Top - .Top > wsSch.Cells(intRow - 1, 1).Height, wsSch.Cells(intRow, 1).Top - .Top, wsSch.Cells(intR - 1, 1).Height)
            If strLabel = Empty Then strLabel = "1R"
            
            Select Case Right(strLabel, 1)
                Case "L"
                    .Left = .Left - .Width + IIf(Left(strLabel, 1) = "0", 0, .TextFrame.MarginRight * 2)
                    .Placement = xlMoveAndSize
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight

                Case "M"
                    .Left = .Left - .Width / 2
                    .Placement = xlMoveAndSize
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Case "R"
                    .Left = .Left - IIf(Left(strLabel, 1) = "0", 0, .TextFrame.MarginLeft)
            End Select
            Select Case Left(strLabel, 1)
                Case "0"
                    .Top = .Top + wsSch.Cells(intR, 1).Height * dblBarPos - .Height / 2 + dblHeight / 2
                Case "1", "2"
                    .Top = .Top - IIf(Left(strLabel, 1) = "1", 0.8, 1.5) * .Height
                    InsertLine dblLeft, .Top + .Height * 3 / 4, dblLeft, dblTop + wsSch.Cells(intR, 1).Height * (dblBarPos + 0.1), "VB_DLIN_", intRowSum
            End Select

        Else
            .Top = .Top - .Height * 1 / 2 + dblHeight * 1 / 2
        End If
    End With
End Sub

Private Sub InsertLine(begX, begY, endX, endY As Double, strName As String, Optional ByVal intR As Integer = Empty)
        Dim strNameFull As String
        
        strNameFull = strName & format(intR, "00000")
    
        On Error Resume Next
        wsSch.Shapes(strNameFull).Delete
        On Error GoTo 0
        'Crear forma
        Set sh = wsSch.Shapes.AddLine(begX, begY, endX, endY)
        'Nombrar
        sh.Name = strNameFull
        UpdateShapeArray Split(strName, "_")(1), intR
        
        With sh
            .Visible = True
            .Line.DashStyle = msoLineSolid
            .Line.Weight = 1
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        End With


End Sub

'Crear etiqueta duración
Private Sub InsertDuration(dblLeft, dblTop, dblWidth, dblHeight As Double, Optional ByVal intR As Integer = Empty)
    Dim strLabInput, strLabDur As String
    Dim strNameFull As String
    If UCase(strLabel) = "NO" Then Exit Sub
    
    strLabDur = varRmgDur & IIf(varRmgDur = 1, " day", " days")
    strLabInput = IIf(booLabStart, format(IIf(Not datFinish > 0, "", datStart), rngStart.Offset(intRow - rngRef.row).NumberFormat), "") & _
                    IIf(booLabDur And booLabStart, ", ", " ") & IIf(booLabDur, strLabDur, "")
    strLabInput = Trim(strLabInput)
    If Len(strLabInput) = 0 Then Exit Sub
    strNameFull = "VB_DUR_" & format(IIf(intR = Empty, intRow, intRowSum), "00000")
    
    On Error GoTo CreateDescription
    Set sh = wsSch.Shapes(strNameFull)
    With sh
        If Not .Left = dblLeft Then .Left = dblLeft
        If Not .Top = dblTop Then .Top = dblTop
        If Not .Width = Abs(dblWidth) Then .Width = Abs(dblWidth)
        If Not .Height = dblHeight Then .Height = dblHeight
    End With
    GoTo DescriptionExists
    
CreateDescription:
    'Crear forma
    Set sh = wsSch.Shapes.AddShape(msoShapeRectangle, dblLeft, dblTop, dblWidth, dblHeight)
    'Nombrar
    sh.Name = strNameFull
    
DescriptionExists:
    On Error GoTo 0
    UpdateShapeArray "DUR", IIf(intR = Empty, intRow, intRowSum)
    
    With sh
        'Texto
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .TextFrame2.TextRange.Characters.text = strLabInput
        .Placement = xlMoveAndSize
        .TextFrame.AutoSize = True
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        .Visible = True
        'Enviar al fondo
        .ZOrder msoBringToFront
        'Editar color de la forma
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        .Top = .Top - .Height * 1 / 2 + dblHeight * 1 / 2
        .Left = .Left - .Width + .TextFrame.MarginRight / 2
    End With
End Sub


'Crear ventana
Private Sub InsertWindow(varShape As Variant, dblLeft, dblTop, dblWidth, dblHeight As Double, strName As String, ByVal lngColor As Long)
    Dim intR As Integer
    Dim strNameFull As String
    
    intR = IIf(strName Like "VB_S??_", intRowSum, intRow)
    strNameFull = strName & format(intR, "00000")
    
    On Error GoTo CreateShape
    Set sh = wsSch.Shapes(strNameFull)
    With sh
        If Not Round(.Left, 2) = Round(dblLeft, 2) Then .Left = dblLeft
        If Not Round(.Top, 2) = Round(dblTop, 2) Then .Top = dblTop
        If Not Round(.Width, 2) = Round(Abs(dblWidth), 2) Then .Width = Abs(dblWidth)
        If Not Round(.Height, 2) = Round(dblHeight, 2) Then .Height = dblHeight
    End With
    GoTo ShapeExists

CreateShape:
    'Crear forma
    Set sh = wsSch.Shapes.AddShape(varShape, dblLeft, dblTop, Abs(dblWidth), dblHeight)
    'Nombrar
    sh.Name = strNameFull
    
ShapeExists:
    On Error GoTo 0
    UpdateShapeArray "WIN", intRow
    With sh
        'Editar color de la forma
        If wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color <> 16777215 Then
            .Fill.ForeColor.RGB = wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color
            .Line.ForeColor.RGB = wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color
        Else:
            .Fill.ForeColor.RGB = lngColor
            .Line.ForeColor.RGB = lngColor
        End If
        .Fill.Transparency = 0.75
        .Line.Weight = 2
        .Visible = True
        .ZOrder msoSendToBack
    End With
    '-------------------------------------------------------------------
    'Insertar descripción
    InsertWindowDescription sh, lngColor
    
End Sub

Private Sub InsertWindowDescription(sh As shape, lngColor As Long)
    Dim strDesc As String
    Dim dblTopDesc, dblLeftDesc, dblHeightDesc, dblWidthDesc As Double
    Dim shDesc As shape
    Dim strNameFull As String
    
    strLabel = UCase(strLabel)
    strLabel = IIf(strLabel = "", "TM", strLabel) 'TM: Top / Middle
    If strLabel = "NO" Then Exit Sub
    'Descripción en variable
    strDesc = arrAct(6)
    'Valores para posición por defecto
    dblHeightDesc = wsSch.Cells(intRow, rngLabPos.Column).Height
    dblTopDesc = sh.Top - dblHeightDesc
    dblLeftDesc = sh.Left
    dblWidthDesc = sh.Width
    
    strNameFull = "VB_WDES_" & format(intRow, "00000")
    
    On Error GoTo CreateDescription
    Set shDesc = wsSch.Shapes(strNameFull)
    GoTo DescriptionExists

CreateDescription:
    'Crear forma
    Set shDesc = wsSch.Shapes.AddShape(msoShapeRectangle, dblLeftDesc, dblTopDesc, dblWidthDesc, dblHeightDesc)
    'Nombrar
    shDesc.Name = strNameFull
    
DescriptionExists:
    On Error GoTo 0
    UpdateShapeArray "WDES", intRow

    With shDesc
        'Texto
        .TextFrame2.TextRange.Characters.text = strDesc
        .TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame.AutoSize = True
        .Visible = True
        'Enviar al frente
        .ZOrder msoBringToFront
        .Placement = xlMoveAndSize
        'Editar color de la forma
        If wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color <> 16777215 Then
            .Line.ForeColor.RGB = wsSch.Cells(intRow, rngActStyle.Column).DisplayFormat.Interior.Color
        Else:
            .Line.ForeColor.RGB = lngColor
        End If
        .Fill.ForeColor.RGB = 16777215
        
        .Fill.Transparency = 0.2
        .Line.Weight = 2
        
        'Reposicionar
        dblHeightDesc = shDesc.Height
        dblWidthDesc = shDesc.Width
        Select Case Right(strLabel, 1)
        Case "R"
            dblLeftDesc = sh.Left + sh.Width - dblWidthDesc / 2
        Case "M"
            dblLeftDesc = sh.Left + sh.Width / 2 - dblWidthDesc / 2
        Case "L"
            dblLeftDesc = sh.Left - dblWidthDesc / 2
        End Select
        Select Case Left(strLabel, 1)
        Case "0"
            dblTopDesc = sh.Top - dblHeightDesc
        Case "1"
            dblTopDesc = sh.Top - 2 * dblHeightDesc
        Case "2"
            dblTopDesc = sh.Top - 3 * dblHeightDesc
        Case "T"
            dblTopDesc = rngRef.Offset(1).Top - dblHeightDesc
        End Select
        .Left = dblLeftDesc
        .Top = dblTopDesc
    End With
    
    'Añadir conector para el cartel
    If Not Left(strLabel, 1) = "0" Then
        InsertLine shDesc.Left + shDesc.Width / 2, shDesc.Top + shDesc.Height, sh.Left + sh.Width / 2, sh.Top, "VB_WLIN_", intRow
    End If

End Sub


'Disabled in Free Edition
'Insertar línea de corte del programa actualizado
Public Sub InsertCutoff()
    If intEdition = 1 Then Exit Sub
    
    Dim dblPosLeft, dblPosTop, dblPosBottom As Double
    Dim intRow As Integer
    Dim datD As Date

    'Posicionar línea
    dblPosLeft = DateDiff("d", datChartStart, datCutoff, intWeekStart)
    datD = datChartStart
    Select Case intPeriod
        Case 1, 2 'Diario
            dblPosLeft = dblPosLeft + 1
        Case 3 'Semanal
            datCutoff = datCutoff + 1
            dblPosLeft = 0
            Do While datD < datCutoff
                If Month(datD) = Month(datD + 7) Or day(datD + 7) = 1 Then
                    If datCutoff >= datD + 7 Then
                        dblPosLeft = dblPosLeft + 1
                    Else: dblPosLeft = dblPosLeft + DateDiff("d", datD, datCutoff, intWeekStart) / 7
                    End If
                Else
                    If datCutoff >= datD + 7 Then
                        'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                        dblPosLeft = dblPosLeft + IIf(datD = datChartStart, 1, 2)
                    ElseIf datCutoff >= DateSerial(Year(datD), Month(datD) + 1, 1) Then
                        dblPosLeft = dblPosLeft + 1 + (datCutoff - DateSerial(Year(datD), Month(datD) + 1, 1)) / (datD + 7 - DateSerial(Year(datD), Month(datD) + 1, 1))
                    Else: dblPosLeft = dblPosLeft + (datCutoff - datD) / (DateSerial(Year(datD), Month(datD) + 1, 1) - datD)
                    End If
                End If
                datD = datD + 7
            Loop
        Case 4 'Bi-semanal
            datCutoff = datCutoff + 1
            dblPosLeft = 0
            Do While datD < datCutoff
                If Month(datD) = Month(datD + 14) Or day(datD + 14) = 1 Then
                    If datCutoff >= datD + 14 Then
                        dblPosLeft = dblPosLeft + 1
                    Else: dblPosLeft = dblPosLeft + DateDiff("d", datD, datCutoff, intWeekStart) / 14
                    End If
                Else
                    If datCutoff >= datD + 14 Then
                        'Corrección de celdas combinadas para semanas a caballo entre dos meses en la primera fecha del calendario
                        dblPosLeft = dblPosLeft + IIf(datD = datChartStart, 1, 2)
                    ElseIf datCutoff >= DateSerial(Year(datD), Month(datD) + 1, 1) Then
                        dblPosLeft = dblPosLeft + 1 + (datCutoff - DateSerial(Year(datD), Month(datD) + 1, 1)) / (datD + 14 - DateSerial(Year(datD), Month(datD) + 1, 1))
                    Else: dblPosLeft = dblPosLeft + (datCutoff - datD) / (DateSerial(Year(datD), Month(datD) + 1, 1) - datD)
                    End If
                End If
                datD = datD + 14
            Loop
        Case 5 'Mensual
            dblPosLeft = Round(DateDiff("d", datChartStart, DateSerial(Year(datCutoff), Month(datCutoff), 1), intWeekStart) / 30.4, 0) + _
                        (DateDiff("d", DateSerial(Year(datCutoff), Month(datCutoff), 1), datCutoff, intWeekStart) + 1) / GetMonthDays(datCutoff)
        Case 6 'Trimestre
            dblPosLeft = Round(DateDiff("d", datChartStart, DateSerial(Year(datCutoff), 3 * ((Month(datCutoff) - 1) \ 3) + 1, 1), intWeekStart) / 91.25, 0) + _
                        (DateDiff("d", DateSerial(Year(datCutoff), 3 * ((Month(datCutoff) - 1) \ 3) + 1, 1), datCutoff, intWeekStart) + 1) / GetQuarterDays(datCutoff)
        Case 7 'Anual
            dblPosLeft = Round(DateDiff("d", datChartStart, DateSerial(Year(datCutoff), 1, 1)) / 365, 0) + _
                        (DateDiff("d", DateSerial(Year(datCutoff), 1, 1), datCutoff) + 1) / 365
    End Select
    dblPosLeft = rngRef.Offset(0, Int(dblPosLeft)).Left + (dblPosLeft - Int(dblPosLeft)) * rngRef.Offset(0, Int(dblPosLeft)).Width
    
    intRow = rngRef.Offset(1, 0).row
    dblPosTop = rngRef.Offset(intRow - rngRef.row, 0).Top
    
    intRow = intActLastRow + 1
    dblPosBottom = rngRef.Offset(intRow - rngRef.row, 0).Top
    
    'Crear forma
    Set sh = wsSch.Shapes.AddLine(dblPosLeft, dblPosTop, dblPosLeft, dblPosBottom)
    With sh
        'Nombrar
        .Name = "VB_CUTOFF"
        'Editar color de la forma
        .Line.ForeColor.RGB = lngCutoffColor
        .Line.Weight = 2
    End With
End Sub

Private Sub BarStyle()
    Dim intBar As Integer
    Dim strShp As String
    intBar = IIf(varMilStyleRow > 0 And varMilStyleRow <= 10, varMilStyleRow, intBarStyle)
    With sh
        Select Case intBar
            Case 1
                .AutoShapeType = msoShapeRectangle
               .Fill.OneColorGradient msoGradientHorizontal, 4, 0
            Case 2
                .AutoShapeType = msoShapeRectangle
               .Fill.Solid
            Case 3
                .AutoShapeType = msoShapeRectangle
               .Fill.Patterned msoPatternWideUpwardDiagonal
            Case 4
                .AutoShapeType = msoShapeRectangle
               .Fill.Patterned msoPatternDarkVertical
            Case 5
                .AutoShapeType = msoShapeRectangle
               .Fill.Patterned msoPatternNarrowHorizontal
            Case 6
                .AutoShapeType = msoShapeRectangle
               .Fill.Patterned msoPatternLargeCheckerBoard
            Case 7
                .AutoShapeType = msoShapeRectangle
               .Fill.Patterned msoPatternLargeConfetti
            Case 8 'Left & Right Arrow
                Select Case Mid(.Name, 4, 3)
                Case "PRA"
                    If datFinish > datCutoff Then
                        strShp = msoShapeLeftArrow
                    Else: strShp = msoShapeLeftRightArrow
                    End If
                Case "PRR"
                    If datStart <= datCutoff Then
                        strShp = msoShapeRightArrow
                    Else: strShp = msoShapeLeftRightArrow
                    End If
                Case Else
                    strShp = msoShapeLeftRightArrow
                End Select
                .AutoShapeType = strShp
                .Fill.Solid
                .Height = .Height * 2
                .Top = .Top - .Height / 4
            Case 9 'Left Arrow
                Select Case Mid(.Name, 4, 3)
                Case "REM", "SAR", "PRR"
                    If datStart <= datCutoff Then
                        strShp = msoShapeRectangle
                    Else: strShp = msoShapeLeftArrow
                    End If
                Case Else
                    strShp = msoShapeLeftArrow
                End Select
                If strShp = msoShapeLeftArrow Then
                    .AutoShapeType = strShp
                    .Fill.Solid
                    .Height = .Height * 2
                    .Top = .Top - .Height / 4
                Else
'                    .Height = .Height * 0.75
'                    .Top = .Top + .Height * 0.5
                End If
            Case 10 'Right Arrow
                Select Case Mid(.Name, 4, 3)
                Case "ACT", "SAA", "PRA"
                    If datFinish > datCutoff Then
                        strShp = msoShapeRectangle
                    Else: strShp = msoShapeRightArrow
                    End If
                Case Else
                    strShp = msoShapeRightArrow
                End Select
                If strShp = msoShapeRightArrow Then
                    .AutoShapeType = strShp
                    .Fill.Solid
                    .Height = .Height * 2
                    .Top = .Top - .Height / 4
                Else
'                    .Height = .Height * 0.75
'                    .Top = .Top + .Height * 0.5
                End If
        End Select
    End With
End Sub

Private Sub MilStyle()
    Dim intMil As Integer
    intMil = IIf(varMilStyleRow > 10 And varMilStyleRow <= 17, varMilStyleRow, intMilStyle)
    With sh
        Select Case intMil
           Case 11
              .AutoShapeType = msoShapeDiamond
           Case 12
              .AutoShapeType = msoShapeFlowchartMerge
           Case 13
              .AutoShapeType = msoShapeFlowchartExtract
           Case 14
              .AutoShapeType = msoShape5pointStar
           Case 15
              .AutoShapeType = msoShape4pointStar
           Case 16
              .AutoShapeType = msoShapeHeart
           Case 17
              .AutoShapeType = msoShapeOval
        End Select
    End With
End Sub

Private Sub ChartLines()

    With wsSch.Range(Cells(rngRef.row + 1, rngRef.Column), Cells(intActLastRow, intLastCol))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    With wsSch.Range(Cells(rngRef.row, rngRef.Column), Cells(rngRef.row, intLastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

Private Function maxLabelPos()
    Dim intR As Integer
    Dim intPos As Integer
    Dim strLabPos As Variant
    
    intR = intRow + 2
    intPos = Empty
    Do While rngTmlCod.Offset(intR - rngRef.row) = strRef
        strLabPos = rngLabPos.Offset(intR - rngRef.row)
        If Not UCase(strLabPos) = "NO" Then
            If strLabPos = Empty Then strLabPos = "1R"
            strLabPos = Left(strLabPos, 1)
            If intPos = Empty Or intPos < CInt(strLabPos) Then intPos = CInt(strLabPos)
        End If
        intR = intR + 1
    Loop
    maxLabelPos = IIf(intPos = Empty, 0, intPos)
End Function

'Construcción del vector de actividades
Private Sub SetActArray_forInsertRows(Optional ByVal intRowUpdate As Variant = Empty)
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux(), arrShp() As Variant
    Dim row As Variant
    Dim i, j As Integer
    
    'Se establece la dimensión del vector: número de filas
    intDim = intActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 2
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value, _
                    Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngDesc.Offset(1), rngDesc.Offset(intDim)).value)
                    
    'Se dimensiona el vector que contendrá la información final al número de filas y un vector auxiliar al número de campos
    ReDim arrData(intDim - 1)
    
    'Para cada fila
    For i = 0 To intDim - 1
        ReDim arrAux(intFields + 2)
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
        arrAux(3) = IIf(IsEmpty(intRowUpdate), True, False)
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux

    Next
    
    If Not IsEmpty(intRowUpdate) Then
        For Each row In intRowUpdate
            arrData(row - rngRef.row - 1)(3) = True
            arrData(row - rngRef.row - 1)(4) = row
        Next
    End If
    
    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Timeline Mode
    '1 --> Activity ID
    '2 --> Description
    'Dimensiones extra
    '3 --> Actividad actualizable segun vector intRowUpd
    '4 --> Referencia a la fila actualizada intRowUpdate
End Sub

Private Sub InsertRows(Optional ByRef intRowUpdate As Variant = Empty)
    Dim col, posData, intInsertedRows, posRowUpd As Integer

    SetActArray_forInsertRows intRowUpdate

    intRow = rngRef.row + 1
    posData = 0
    intInsertedRows = 0
    For Each arrAct In arrData
        If Not arrAct(3) Then GoTo NextIteration
        
        'Si la fila anterior no tiene nombre de actividad y la fila actual no es un timeline se elimina
        If Not intRow = rngRef.row + 1 Then
            If IsEmpty(arrData(posData - 1)(1)) And IsEmpty(arrData(posData - 1)(2)) And IsEmpty(arrAct(0)) Then
                wsSch.Cells(intRow - 1, 1).EntireRow.Delete Shift:=xlUp
                intActLastRow = intActLastRow - 1
                intRow = intRow - 1
                intInsertedRows = intInsertedRows - 1
                arrData(posData)(4) = arrAct(4) + intInsertedRows
                GoTo NextIteration
            End If
        End If
        
        
        'Insertar fila si estamos en un timeline
        If Not IsEmpty(arrAct(0)) Then
            'En la primera fila
            If intRow = rngRef.row + 1 Then
                GoTo InsertRow
            'o la fila anterior no es una fila en blanco
            ElseIf Not (IsEmpty(arrData(posData - 1)(1)) And IsEmpty(arrData(posData - 1)(2))) Then
                GoTo InsertRow
            Else: GoTo NextIteration
            End If
        Else: GoTo NextIteration
        End If
InsertRow:
        'Insertar fila copiada/pegada de la fila siguiente y eliminar contenidos
        wsSch.Cells(intRow, 1).EntireRow.Insert Shift:=xlDown
        wsSch.Cells(intRow + 1, 1).EntireRow.Copy
        wsSch.Cells(intRow, 1).EntireRow.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        For col = rngPeriod.Column To 1 Step -1
            With wsSch.Cells(intRow, col)
                .ClearContents
                .Borders(xlEdgeBottom).LineStyle = xlNone
            End With
        Next
        wsSch.Rows(intRow).OutlineLevel = wsSch.Rows(intRow + 1).OutlineLevel
        'Se modifica la referencia de la última fila y la fila sumario
        intActLastRow = intActLastRow + 1
        intRow = intRow + 1
        intInsertedRows = intInsertedRows + 1
        arrData(posData)(4) = arrAct(4) + intInsertedRows
NextIteration:
        intRow = intRow + 1
        posData = posData + 1
    Next
    
    'Corrección del vector de posiciones a actualizar
    If Not IsEmpty(intRowUpdate) Then
        posRowUpd = 0
        For Each arrAct In arrData
            If arrAct(3) Then
                intRowUpdate(posRowUpd) = arrAct(4)
                posRowUpd = posRowUpd + 1
            End If
        Next
    End If
End Sub

Private Function GetMonthDays(ByVal datRef As Date) As Integer
    Select Case Month(datRef)
    Case 1
        GetMonthDays = 31
    Case 2
        GetMonthDays = 28
    Case 3
        GetMonthDays = 31
    Case 4
        GetMonthDays = 30
    Case 5
        GetMonthDays = 31
    Case 6
        GetMonthDays = 30
    Case 7
        GetMonthDays = 31
    Case 8
        GetMonthDays = 31
    Case 9
        GetMonthDays = 30
    Case 10
        GetMonthDays = 31
    Case 11
        GetMonthDays = 30
    Case 12
        GetMonthDays = 31
    End Select
End Function

Private Function GetQuarterDays(ByVal datRef As Date) As Integer
    Select Case Month(datRef)
    Case 1, 2, 3
        GetQuarterDays = 90
    Case 4, 5, 6
        GetQuarterDays = 91
    Case 7, 8, 9
        GetQuarterDays = 92
    Case 10, 11, 12
        GetQuarterDays = 92
    End Select
End Function
Public Sub FilterShapes()
    Dim shp As shape
    Dim arrShapeName As Variant
    Dim strShpCod As String
    Dim intShpRow As Integer
    Dim i As Integer
    
    For Each shp In wsSch.Shapes
        If shp.Name Like "VB_*_*" Then
            arrShapeName = Split(shp.Name, "_")
            strShpCod = arrShapeName(1)
            intShpRow = CInt(arrShapeName(2))
            
            If strShpCod Like "S*" Or strShpCod Like "DLIN" Then
                i = 0
                Do While Len(rngTmlMod.Offset(intShpRow - rngRef.row - i)) = 0
                    i = i + 1
                Loop
                If wsSch.Rows(intShpRow - i).Hidden = True Then
                    shp.Visible = False
                Else: shp.Visible = True
                End If
            ElseIf wsSch.Rows(intShpRow).Hidden = True And Not strShpCod Like "W*" Then
                shp.Visible = False
            Else
                shp.Visible = True
            End If
        End If
    Next
    
    'Ocultar todas las agrupaciones
    For i = rngRef.row + 1 To intActLastRow
        If Rows(i + 1).OutlineLevel > Rows(i).OutlineLevel And Rows(i + 1).Hidden = False And Len(Cells(i, rngTmlMod.Column)) > 0 And Len(Cells(i, rngTmlCod.Column)) > 0 Then
            wsSch.Rows(i).ShowDetail = False
        End If
    Next
End Sub
