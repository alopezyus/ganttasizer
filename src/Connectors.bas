Attribute VB_Name = "Connectors"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition
Option Explicit

Dim dblMaxX, dblMinX, dblY As Double
Dim intRowSucc, intRowPred As Integer
Dim sh As shape
Dim strTmlModLast As String

Dim arrData() As Variant
Dim arrAct As Variant
 
 'Construcción del vector de actividades
Private Sub SetActArray()
    If intEdition = 1 Then Exit Sub
    
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux() As Variant
    Dim i, j As Integer
    
    'Se establece la dimensión del vector: número de filas
    intDim = ActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 4
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngConStyle.Offset(1), rngConStyle.Offset(intDim)).value, _
                    Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value, _
                    Range(rngTmlCod.Offset(1), rngTmlCod.Offset(intDim)).value, _
                    Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngPred.Offset(1), rngPred.Offset(intDim)).value)

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
    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Connector Style
    '1 --> Timeline Mode
    '2 --> Timeline Code
    '3 --> Activity ID
    '4 --> Predecessors
End Sub

'Eliminar todos los conectores
Public Sub ClearConnectors()
    If intEdition = 1 Then Exit Sub
    Dim d As shape
    Dim intShapesCount, intShapesTot As Integer
    
    intShapesCount = 0
    intShapesTot = wsSch.Shapes.Count
    For Each d In wsSch.Shapes
        If d.Name Like "VB_CON*" Then
           d.Delete
        End If
        intShapesCount = intShapesCount + 1
        'UpdateProgressBar IIf(intShapesTot = 0, 0, intShapesCount / intShapesTot)
    Next d
    
    'UpdateProgressBar 1
End Sub

Public Sub CreateConnectors()
    If intEdition = 1 Then Exit Sub
    
On Error GoTo errHandler
    Dim dblBegX, dblBegY, dblEndX, dblEndY As Double
    Dim dblMaxXp, dblMinXp, dblYp, dblMaxXs, dblMinXs, dblYs As Double
    Dim strPredList, strShp, strActID, strRel, strRelPred, strRelSucc As String
    Dim strArray() As String
    Dim i, intCntTmlCod As Integer
    Dim intRowRefIni As String
    
'    wsSch.Select
    
    intRowRefIni = rngRef.row
    strTmlModLast = ""
    
    SetActArray
    intRowSucc = rngRef.row + 1
    'Inicio bucle para recorrer todas las filas
    For Each arrAct In arrData
    'Comienza bucle para recorrer todas las filas de la lista de actividades
        If Not IsEmpty(arrAct(1)) Then GoTo nextSuccessor
        
        strPredList = arrAct(4)
        
        'Vector con todas las predecesoras de esta fila
        strArray() = Split(strPredList, ",")
        'Quitar espacios por delante y por detrás
        For i = 0 To UBound(strArray)
            strArray(i) = Trim(strArray(i))
        Next
        
        'Si el vector de predecesoras no es nulo se continua con el procedimiento
        If UBound(strArray()) >= 0 Then

            'Guardamos Timeline Code para en procedimiento GetPostition poder dar tratamiento especial a los timelines
            If Not IsEmpty(arrAct(2)) Then
                intCntTmlCod = 0
                Do While Len(arrAct(2)) = Len(arrData(intRowSucc - (intCntTmlCod + 1) - rngRef.row - 1)(2))
                    intCntTmlCod = intCntTmlCod + 1
                Loop
                strTmlModLast = arrData(intRowSucc - intCntTmlCod - rngRef.row - 1)(1)
            Else: strTmlModLast = ""
            End If
       
            'Se extraen las posiciones extremas para las formas asociadas a la actividad
            GetPosition (intRowSucc)
            dblMaxXs = dblMaxX
            dblMinXs = dblMinX
            dblYs = dblY
            If IsEmpty(dblMaxXs) Then GoTo nextSuccessor
    
            'Comienza bucle para recorrer todas las relaciones definidas para la actividad de la fila analizada
            For i = 0 To UBound(strArray)
                'Separar las partes de la relación de precedencia definida
                strActID = Split(strArray(i), " ")(0)
                
                If InStr(1, strArray(i), " ") = 0 Then
                    strRel = "FS"
                Else
                    strRel = Left(Split(strArray(i), " ")(1), 2)
                End If
                strRelPred = Left(strRel, 1)
                strRelSucc = Right(strRel, 1)
                
                'Fijar posición de la sucesora
                dblEndX = IIf(strRelSucc = "S", dblMinXs, dblMaxXs)
                dblEndY = dblYs
                
                'Buscar predecesora por ActID y guardar código de fila en formato texto
'                intRowPred = rngActID.EntireColumn.Find(strActID, LookAt:=xlWhole).Row
                For intRowPred = 0 To UBound(arrData)
                    If arrData(intRowPred)(3) = strActID Then GoTo PredFound
                Next
PredFound:
                intRowPred = intRowPred + (rngRef.row + 1)
                'Guardamos Timeline Code para en procedimiento GetPostition poder dar tratamiento especial a los timelines
                If Not IsEmpty(arrData(intRowPred - (rngRef.row + 1))(2)) Then
                    intCntTmlCod = 0
                    Do While Len(arrData(intRowPred - (rngRef.row + 1))(2)) = Len(arrData(intRowPred - (intCntTmlCod + 1) - (rngRef.row + 1))(2))
                        intCntTmlCod = intCntTmlCod + 1
                    Loop
                    strTmlModLast = arrData(intRowPred - intCntTmlCod - (rngRef.row + 1))(1)
                Else: strTmlModLast = ""
                End If
                
                'Se extraen las posiciones extremas para las formas asociadas a la actividad
                GetPosition (intRowPred)
                dblMaxXp = dblMaxX
                dblMinXp = dblMinX
                dblYp = dblY
                If IsEmpty(dblMaxXp) Then GoTo nextPredecessor
                
                'Fijar posición de la predecesora
                dblBegX = IIf(strRelPred = "S", dblMinXp, dblMaxXp)
                dblBegY = dblYp
                
                'Correccion de dblEndX cuando su posición debería ser igual a dblBegX pero no hay ajuste exacto
                'Se asume que esto es, cuando la diferencia entre las posiciones en X es menor que la mitad del ancho de un día de calendario
                If Abs(dblEndX - dblBegX) < rngRef.Offset(, 1).Width / _
                    (2 * IIf(intPeriod = 7, 360, IIf(intPeriod = 6, 91, IIf(intPeriod = 5, 30, IIf(intPeriod = 4, 14, IIf(intPeriod = 3, 7, 1)))))) Then
                    dblEndX = dblBegX
                End If
                
                'Insertar conector
                Call InsertConnector(msoConnectorElbow, dblBegX, dblBegY, dblEndX, dblEndY, strRel)
nextPredecessor:
                intRowPred = intRowPred + 1
            Next
        End If
        'UpdateProgressBar IIf(intActLastRow = 0, 0, (rngRef.Row + 1) / intActLastRow)
nextSuccessor:
        intRowSucc = intRowSucc + 1
    Next
    
    'UpdateProgressBar 1
    Exit Sub
errHandler:
    'UpdateProgressBar 1
    If intRowRefIni < rngRef.row Then wsSch.Cells(1, 1).EntireRow.Delete Shift:=xlUp
    CustomMsgBox "An error has occurred during execution at row: " & intRowSucc, vbCritical + vbOKOnly, error:=True
End Sub

Private Sub GetPosition(intR As Integer)
    If intEdition = 1 Then Exit Sub
    Dim d As shape
    Dim strShp As String
    Dim booTml As Boolean
    
'    strShp = "VB_" & IIf(Not wsSch.Cells(intR, rngTmlCod.Column) = "" And (strTmlModLast Like "MIL" Or strTmlModLast Like "ACT"), "S*", "*") & format(intR, "00000")
    strShp = "VB_" & "*_" & format(intR, "00000")
    booTml = Not IsEmpty(arrData(intR - (rngRef.row + 1))(2)) And (strTmlModLast Like "MIL" Or strTmlModLast Like "ACT")
    
    dblMaxX = Empty
    dblMinX = Empty
    dblY = Empty
    
    For Each d In wsSch.Shapes
        If d.Name Like strShp And IIf(booTml, d.Name Like "VB_S*", Not d.Name Like "VB_S*") And Not (d.Name Like "*0_*") And _
            Not (d.Name Like "*DES*") And Not (d.Name Like "*DUR*") And Not (d.Name Like "*LIN*") And Not (d.Name Like "*FLT*") Then
            If dblMaxX = Empty Or dblMinX = Empty Then
                dblMaxX = d.Left + d.Width
                dblMinX = d.Left
            Else
                dblMaxX = IIf((d.Left + d.Width) > dblMaxX, (d.Left + d.Width), dblMaxX)
                dblMinX = IIf(d.Left < dblMinX, d.Left, dblMinX)
            End If
            If d.Name Like "*MLA*" Or d.Name Like "*MLR*" Or d.Name Like "*BM0*" Or d.Name Like "*SMA*" Or d.Name Like "*SMR*" Or d.Name Like "*SM0*" Then
                dblMaxX = dblMaxX - (dblMaxX - dblMinX) / 2
                dblMinX = dblMaxX
            End If
            dblY = d.Top + d.Height / 2
        End If
    Next d

End Sub

'Insertar Conector
Private Sub InsertConnector(varShape As Variant, dblBegX, dblBegY, dblEndX, dblEndY As Double, ByVal strRel As String) ', strName As String, ByVal rngColor As Range)
    If intEdition = 1 Then Exit Sub
    
    Dim dblElbow As Double
    Dim Lx, Ly, Cx, Cy As Double
    Dim dblBegXp, dblBegYp, dblEndXp, dblEndYp As Double
    
       
    'Insertar Conector
    Select Case strRel
        Case "FF", "SS"
            Set sh = wsSch.Shapes.AddConnector(varShape, dblBegX, dblBegY, dblEndX, dblEndY)
        Case "FS"
            If Round(dblBegX, 2) <= Round(dblEndX, 2) Then
                Set sh = wsSch.Shapes.AddConnector(varShape, dblBegX, dblBegY, dblEndX, dblEndY)
            Else 'Relación con start anterior al finish: recalcular dimensiones para luego girar el conector
                'Lado x
                Lx = dblEndX - dblBegX
                'Lado y
                Ly = dblEndY - dblBegY
                'Calculo coordenadas Punto Central
                Cx = dblBegX + Lx / 2
                Cy = dblBegY + Ly / 2
                'Calculo de nuevas coordenadas para posicionamiento de conector
                Lx = Abs(Lx)
                Ly = Abs(Ly)
                dblBegXp = Cx - Ly / 2
                dblBegYp = Cy - Lx / 2
                dblEndXp = Cx + Ly / 2
                dblEndYp = Cy + Lx / 2
                
                If dblBegYp < 0 Then
                    wsSch.Cells(1, 1).EntireRow.Insert , xlFormatFromRightOrBelow
                    wsSch.Cells(1, 1).RowHeight = Abs(dblBegYp)
                End If
                
                Set sh = wsSch.Shapes.AddConnector(varShape, dblBegXp, dblBegYp, dblEndXp, dblEndYp)
                sh.IncrementRotation 90
                If dblBegY > dblEndY Then
                    sh.Flip msoFlipVertical
                End If
                If dblBegYp < 0 Then wsSch.Cells(1, 1).EntireRow.Delete Shift:=xlUp
            
            End If
        Case "SF"
            If Round(dblBegX, 2) >= Round(dblEndX, 2) Then
                Set sh = wsSch.Shapes.AddConnector(varShape, dblBegX, dblBegY, dblEndX, dblEndY)
            Else 'Relación con lag positivo: recalcular dimensiones para luego girar el conector
                'Lado x
                Lx = dblEndX - dblBegX
                'Lado y
                Ly = dblEndY - dblBegY
                'Calculo coordenadas Punto Central
                Cx = dblBegX + Lx / 2
                Cy = dblBegY + Ly / 2
                'Calculo de nuevas coordenadas para posicionamiento de conector
                Lx = Abs(Lx)
                Ly = Abs(Ly)
                dblBegXp = Cx - Ly / 2
                dblBegYp = Cy - Lx / 2
                dblEndXp = Cx + Ly / 2
                dblEndYp = Cy + Lx / 2
                
                If dblBegYp < 0 Then
                    wsSch.Cells(1, 1).EntireRow.Insert , xlFormatFromRightOrBelow
                    wsSch.Cells(1, 1).RowHeight = Abs(dblBegYp)
                End If
                
                Set sh = wsSch.Shapes.AddConnector(varShape, dblBegXp, dblBegYp, dblEndXp, dblEndYp)
                sh.IncrementRotation -90
                If dblBegY < dblEndY Then
                    sh.Flip msoFlipVertical
                End If
                If dblBegYp < 0 Then wsSch.Cells(1, 1).EntireRow.Delete Shift:=xlUp
            End If
    End Select
    
    'Mover Codo del conector
    If dblEndX <> dblBegX Then
        'Buscamos codo de ancho fijo 7 puntos
        dblElbow = 7 / (dblEndX - dblBegX)
        Select Case strRel
            Case "FF"
                dblElbow = IIf(dblBegX <= dblEndX, 1 + dblElbow, dblElbow)
                sh.Adjustments.item(1) = dblElbow
            Case "SS"
                dblElbow = IIf(dblBegX <= dblEndX, -dblElbow, 1 - dblElbow)
                sh.Adjustments.item(1) = dblElbow
'            Case "FS"
'                dblElbow = IIf(dblBegX <= dblEndX, dblElbow, -dblElbow * 2)
'            Case "SF"
'                dblElbow = IIf(dblBegX >= dblEndX, -dblElbow, dblElbow * 2)
        End Select
        'Sh.Adjustments.item(1) = dblElbow
    End If
        
    'Dar formato al conector
    With sh.Line
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadTriangle
        .Visible = msoTrue
        If wsSch.Cells(intRowSucc, rngConStyle.Column).DisplayFormat.Interior.Color <> 16777215 Then
            .ForeColor.RGB = wsSch.Cells(intRowSucc, rngConStyle.Column).DisplayFormat.Interior.Color
        Else
            .ForeColor.RGB = RGB(0, 0, 0)
        End If
        .Transparency = dblConTrn
        .Weight = dblConThk
    End With
    sh.Name = "VB_CON"
        
    ConnectorStyle
    
End Sub

Private Sub ConnectorStyle()
    If intEdition = 1 Then Exit Sub
    
    Dim intCon As Integer
    Dim varConSucc As Variant
    
    varConSucc = Trim(Left(wsSch.Cells(intRowSucc, rngConStyle.Column), 2))
    
    If UCase(varConSucc) = "NO" Then
        sh.Delete
        Exit Sub
    End If
    intCon = IIf(varConSucc > 0 And varConSucc <= 4, varConSucc, intConStyle)
    
    With sh.Line
        Select Case intCon
           Case 1
              .EndArrowheadStyle = msoArrowheadTriangle
           Case 2
              .DashStyle = msoLineDash
              .EndArrowheadStyle = msoArrowheadTriangle
           Case 3
              .DashStyle = msoLineRoundDot
              .EndArrowheadStyle = msoArrowheadTriangle
           Case 4
              .EndArrowheadStyle = msoArrowheadNone
           Case 5
              .DashStyle = msoLineDash
              .EndArrowheadStyle = msoArrowheadNone
           Case 6
              .DashStyle = msoLineRoundDot
              .EndArrowheadStyle = msoArrowheadNone
        End Select
    End With
End Sub
