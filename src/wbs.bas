Attribute VB_Name = "wbs"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition
Option Explicit

Dim arrData() As Variant
Dim arrAct As Variant

'Construcción del vector de actividades
Private Sub SetActArray_format()
    If intEdition = 1 Then Exit Sub
    
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux() As Variant
    Dim i, j As Integer
    
    'Se establece la dimensión del vector: número de filas
    intDim = ActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 2
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngWBS.Offset(1), rngWBS.Offset(intDim)).value, _
                    Range(rngDesc.Offset(1), rngDesc.Offset(intDim)).value)

    'Se dimensiona el vector que contendrá la información final al número de filas y un vector auxiliar al número de campos
    ReDim arrData(intDim - 1)
    ReDim arrAux(intFields + 5)
    
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
        'Activar en Pro Edition
        If intEdition = 2 Then
            Dim strWBS() As Variant
            If InStr(arrAux(1), ".") > 0 Then
                arrAux(1) = Split(arrAux(1), ".")(0) & "." & Split(arrAux(1), ".")(1)
            End If
        End If
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux
    Next
    
    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Activity ID
    '1 --> WBS
    '2 --> Description
    'Dimensiones extra
    '3 --> Delete (boolean)
    '4 --> Add (boolean)
    '5 --> Color RGB (double)
    '6 --> WBS level (integer)
End Sub

Public Sub FormatWBS()
    If intEdition = 1 Then Exit Sub
    
    Dim arrAdd, arrNewWBS, arrCurrWBS, arrColorLevels As Variant
    Dim booAdd As Variant
    Dim intRow, intLevel, intLevelmax, intL, posData, i, j, k As Integer
    Dim strConcatWBS As String
    Dim strPartWBS, dblColor  As Variant
    Dim n As Name
    Dim varReturn As Variant
    Dim rngHead, rngWBSline As Range
    Dim dblTimer As Double
    
    varReturn = returnName
    
    'Actualización del elemento Delete en arrData
    SetActArray_format
    posData = 0
    intLevelmax = -1
    'Para cada actividad de la lista
    For Each arrAct In arrData
        'Si no tiene ni ActID ni descripción
        If IsEmpty(arrAct(0)) And IsEmpty(arrAct(2)) Then
            'Se debe borrar
            arrData(posData)(3) = True
            GoTo NextIteration1
        'Si no tiene ActID (tampoco puede ser una agrupación de WBS)
        ElseIf IsEmpty(arrAct(0)) Then
            'No se borra
            arrData(posData)(3) = False
            GoTo NextIteration1
        'Si no es una agrupación de WBS
        ElseIf Not arrAct(0) Like "WBS-*" Then
            'No se borra
            arrData(posData)(3) = False
            GoTo NextIteration1
        'Si no es la primera actividad (en todo caso, ya sí que estamos en una agrupación WBS)
        ElseIf posData > 0 Then
            'Si el código WBS es igual al de la actividad anterior significa que la agrupación debe ir en una posición superior
            If arrAct(1) = arrData(posData - 1)(1) Or arrData(posData - 1)(1) Like arrAct(1) & ".*" Then
                'Se borra
                arrData(posData)(3) = True
                GoTo NextIteration1
            End If
        End If
        'Si se ha llegado hasta aquí estamos analizando una agrupación WBS que está en la primera posición o tiene un código diferente a la actividad superior
        'Analizamos las actividades posteriores a la actual
        For i = posData + 1 To UBound(arrData)
            'Si la actividad analizada no es una agrupación WBS y no es una línea vacía
            If Not arrData(i)(0) Like "WBS-*" And Not (IsEmpty(arrData(i)(0)) And IsEmpty(arrData(i)(2))) Then
                'Si el código WBS de la actividad actual es igual al de la actividad analizada o está contenido en ella
                If arrAct(1) = arrData(i)(1) Or arrData(i)(1) Like arrAct(1) & ".*" Then
                    'No se borra el nivel
                    arrData(posData)(3) = False
                'Si no está contenido, se borra
                Else: arrData(posData)(3) = True
                End If
                GoTo NextIteration1
            End If
        Next
NextIteration1:
        'Guardamos color
        arrData(posData)(5) = rngWBS.Offset(posData + 1).Interior.Color
        'Guardamos nivel de las actividades con WBS que no se van a eliminar
        If IsEmpty(arrAct(1)) Then
            arrData(posData)(6) = -1
        Else: arrData(posData)(6) = Len(arrAct(1)) - Len(Replace(arrAct(1), ".", ""))
        End If
        'Tenemos en cuenta el nivel de la actividad actual para determinar el nivel máximo sólo si la actividad no va a ser borrada.
        If Not arrData(posData)(3) Then intLevelmax = IIf(arrData(posData)(6) > intLevelmax, arrData(posData)(6), intLevelmax)
        posData = posData + 1
    Next
    
    'Actualización el elemento Add en arrData
    ReDim arrAdd(0)
    ReDim arrNewWBS(0)
    posData = 0
    'Para cada actividad de la lista
    For Each arrAct In arrData
        'Todos los elementos que ya se encuentran en el vector arrData deben tene la propiedad Add en falso
        arrData(posData)(4) = False
        
        'Hay que analizar qué se debe añadir y guardarlo en dos vectores auxiliares
        'Si el WBS está vacío o se tiene que eliminar la fila
        If IsEmpty(arrAct(1)) Or arrAct(3) Then
            'En esta posición no se añade ninguna agrupación
            GoTo NextIteration2
        'Si es una agrupación WBS
        ElseIf arrAct(0) Like "WBS-*" Then
            'Si estamos en la primera posición
            If posData = 0 Then
                'Y es el nivel 0 no hay que añadir nada
                If arrAct(6) = 0 Then GoTo NextIteration2
            Else 'Si no estamos en la primera posición
                'El WBS anterior debe ser distinto y de igual nivel o contenido en el actual y de un nivel inferior
                If (Not arrData(posData - 1)(1) = arrAct(1) And arrData(posData - 1)(6) = arrAct(6)) Or _
                    (arrAct(1) Like arrData(posData - 1)(1) & ".*" And arrData(posData - 1)(6) + 1 = arrAct(6)) Then GoTo NextIteration2
            End If
        'Si no es la primera posición (en todo caso estamos en una actividad con WBS definido que no es una agrupación WBS)
        ElseIf posData > 0 Then
            'Si el WBS de la actividad actual es igual al WBS de la actividad anterior estamos dentro de la misma agrupación
            If arrAct(1) = arrData(posData - 1)(1) Then
                'En esta posición no se añade ninguna agrupación
                GoTo NextIteration2
            End If
        End If
        'Si se ha llegado hasta aquí, estamos analizando una actividad con WBS y cuyo WBS es distinto al de la actividad anterior (o es la primera actividad)
        'O estamos analizando una agrupación WBS que no tiene la agrupación de nivel inferior o una actividad de nivel equivalente justo encima
        'Tenemos que analizar qué agrupaciones debe tener esta actividad por encima y analizar si ya están añadidas o hay que añadirlas
        arrCurrWBS = Split(arrAct(1), ".")
        For i = 0 To UBound(arrCurrWBS) - IIf(arrAct(0) Like "WBS-*", 1, 0)
            strPartWBS = arrCurrWBS(i)
            If i = 0 Then
                strConcatWBS = strPartWBS
            Else: strConcatWBS = strConcatWBS & "." & strPartWBS
            End If
            'Para cada WBS, comprobar si la agrupación está en la lista de actividades por encima de la actividad actual y no eliminada y si no añadirlo
            For j = 0 To posData - 1
                'If arrData(j)(1) = strConcatWBS And arrData(j)(0) Like "WBS-*" And Not arrData(j)(3) Then GoTo WBSexists
                If (arrData(j)(1) = strConcatWBS Or arrData(j)(1) Like strConcatWBS & ".*") And Not arrData(j)(3) Then GoTo WBSexists
            Next
WBSadd:
            'Si se sale por aquí es que no existe la agrupación para el WBS analizado
            'Rellenar vector de referencias con el WBS analizado
            If IsEmpty(arrAdd(0)) Then
                arrAdd(0) = True
                arrNewWBS(0) = strConcatWBS
            Else
                ReDim Preserve arrAdd(UBound(arrAdd) + 1)
                ReDim Preserve arrNewWBS(UBound(arrNewWBS) + 1)
                arrAdd(UBound(arrAdd)) = True
                arrNewWBS(UBound(arrNewWBS)) = strConcatWBS
            End If
WBSexists:
        Next
NextIteration2:
'        If Not booAdd Then
            'Rellenar vector de referencias de añadido cuando no hay que añadir nada
            If IsEmpty(arrAdd(0)) Then
                arrAdd(0) = False
            Else
                ReDim Preserve arrAdd(UBound(arrAdd) + 1)
                ReDim Preserve arrNewWBS(UBound(arrNewWBS) + 1)
                arrAdd(UBound(arrAdd)) = False
            End If
'        End If
        posData = posData + 1
    Next
    
    'Introducimos las filas añadidas en vector arrData
    posData = 0
    For Each booAdd In arrAdd
        If booAdd Then
            'Redimensionar arrData y mover todos los valores una posición para dejar espacio a la línea añadida
            ReDim Preserve arrData(UBound(arrData) + 1)
            For j = UBound(arrData) To posData + 1 Step -1
                arrData(j) = arrData(j - 1)
            Next
            arrData(posData)(0) = "WBS-" & arrNewWBS(posData)
            arrData(posData)(1) = arrNewWBS(posData)
            arrData(posData)(2) = ""
            For j = 0 To UBound(arrData)
                If arrData(j)(0) = arrData(posData)(0) And Len(arrData(j)(0)) > 0 Then arrData(posData)(2) = arrData(j)(2)
            Next
            arrData(posData)(3) = False
            arrData(posData)(4) = True
            arrData(posData)(5) = RGB(255, 255, 255)
            arrData(posData)(6) = Len(arrNewWBS(posData)) - Len(Replace(arrNewWBS(posData), ".", ""))
            intLevelmax = IIf(arrData(posData)(6) > intLevelmax, arrData(posData)(6), intLevelmax)
        End If
        posData = posData + 1
    Next

    'Adición de actividades a la lista
    For i = 0 To UBound(arrData) Step 1
        intRow = i + 1
        If arrData(i)(4) Then
            rngActID.Offset(intRow).EntireRow.Insert Shift:=xlDown
            rngActID.Offset(intRow) = arrData(i)(0)
            rngWBS.Offset(intRow) = "'" & arrData(i)(1)
            rngDesc.Offset(intRow) = arrData(i)(2)
            'Si es la primera fila copiar formato de la siguiente. Si no, los formatos de fecha que se quedan no permiten dibujar las barras
            If i = 0 Then
                rngActID.Offset(intRow + 1).EntireRow.Copy
                rngActID.Offset(intRow).EntireRow.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            End If
            'Quitar el formato de color interior del nivel añadido para que pinte por defecto
            rngWBS.Offset(intRow).Interior.Color = RGB(255, 255, 255)
        End If
    Next
    
    'Eliminación de actividades de la lista
    For i = UBound(arrData) To 0 Step -1
        intRow = i + 1
        If arrData(i)(3) = True Then _
            rngActID.Offset(intRow).EntireRow.Delete Shift:=xlUp
    Next
    
    'Redimensionado del vector quitando actividades eliminadas
    k = 0
    For i = 0 To UBound(arrData)
CheckDelete:
        If arrData(i)(3) Then
            For j = i + 1 To UBound(arrData)
                arrData(j - 1) = arrData(j)
            Next
            k = k + 1
            GoTo CheckDelete
        End If
    Next
    ReDim Preserve arrData(UBound(arrData) - k)

    'Eliminar columnas agrupaciones WBS
    For Each n In ActiveWorkbook.Names
        If n.Name Like "VB_" & varReturn(1) & "_L*" Then
            If InStr(1, n.RefersTo, "#REF!") = 0 Then
                Set rngHead = Range(n.Name)
                rngHead.EntireColumn.Delete
            End If
            'Borrar sólo si es un nivel superior al máximo actual
            n.Delete
        End If
    Next
    
    'Si no se ha calculado nivel máximo no hay agrupaciones a las que dar formato
    If intLevelmax = -1 Then Exit Sub
    
    'Insertar columnas agrupaciones WBS
    For i = 0 To intLevelmax
        Range("A1").EntireColumn.Insert Shift:=xlToRight
        Set rngHead = wsSch.Cells(rngRef.row, 1)
        ActiveWorkbook.Names.Add Name:="VB_" & varReturn(1) & "_L" & format(intLevelmax - i, "00"), RefersToR1C1:=rngHead
        rngHead.ColumnWidth = 1
NextColumn:
    Next
    
    'Rango que contiene las celdas que se deben editar de la línea completa de la actividad
    Set rngWBSline = rngFormatWBS(varReturn(1))
    
    'Inicialización del vector de colores actuales
    ReDim arrColorLevels(intLevelmax)
    For i = 0 To intLevelmax
        arrColorLevels(i) = RGB(255, 255, 255)
    Next
    'Recorremos vector de actividades para dar formato
    i = 0
    For Each arrAct In arrData
        If Len(arrAct(1)) > 0 Then
            intRow = i + 1
            intLevel = arrAct(6)
            'Si es un nivel WBS se actualiza el vector de colores actuales
            If arrAct(0) Like "WBS-*" Then arrColorLevels(intLevel) = arrAct(5)
            'Formato agrupaciones WBS
            For intL = 0 To intLevelmax
                If intL < intLevel Then
                    borderWBS 1, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow)
                    colorWBS intL, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow), arrColorLevels(intL)
                Else:
                    If Not arrAct(0) Like "WBS-*" Then
                        If intL = intLevel Then
                            borderWBS 1, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow)
                            colorWBS intLevel, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow), arrColorLevels(intLevel)
                        Else:
                            colorWBS intLevel, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow), RGB(255, 255, 255), True
                        End If
                    Else
                        If intL = intLevel Then
                            borderWBS 2, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow)
                        Else: borderWBS 3, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow)
                        End If
                        colorWBS intLevel, Range("VB_" & varReturn(1) & "_L" & format(intL, "00")).Offset(intRow), arrColorLevels(intLevel)
                    End If
                End If
            Next
            'Formato para la fila en la tabla de actividades
            If Not arrAct(0) Like "WBS-*" Then
                colorWBS intLevel, rngWBSline.Offset(intRow), arrAct(5), True
                If Not i = UBound(arrData) Then borderWBS 0, rngWBSline.Offset(intRow)
                rngWBSline.Offset(intRow).Font.Bold = False
                rngActStyle.Offset(intRow) = IIf(rngActStyle.Offset(intRow) = "8", "", rngActStyle.Offset(intRow))
            Else
                colorWBS intLevel, rngWBSline.Offset(intRow), arrColorLevels(intLevel)
                borderWBS 3, rngWBSline.Offset(intRow)
                rngWBSline.Offset(intRow).Font.Bold = True
                rngActStyle.Offset(intRow) = IIf(rngActStyle.Offset(intRow) = "", "8", rngActStyle.Offset(intRow))
                rngActStyle.Offset(intRow).Interior.Color = IIf(rngActStyle.Offset(intRow).Interior.Color <> RGB(255, 255, 255), rngActStyle.Offset(intRow).Interior.Color, RGB(178, 178, 178))
            End If
            'Insertar indents en descripción para filas de WBS y filas de actividades
            intLevel = IIf(intLevel < 0, 0, intLevel)
            If intLevel - rngDesc.Offset(intRow).IndentLevel <> 0 Then rngDesc.Offset(intRow).InsertIndent intLevel - rngDesc.Offset(intRow).IndentLevel
            'Crear agrupaciones WBS
            GroupWBS arrAct(0), arrAct(1), rngRef.Offset(intRow).row
        End If
        i = i + 1
    Next
End Sub

'Construcción del vector de actividades
Private Sub SetActArray(ByVal intRowUpdate As Variant)
    If intEdition = 1 Then Exit Sub
    
    Dim intDim, intFields As Integer
    Dim arrVal(), arrAux() As Variant
    Dim i, j As Integer
    Dim row As Variant
    
    'Se establece la dimensión del vector: número de filas
    intDim = ActLastRow - rngRef.row
    'Se establece el número de campos
    intFields = 14
    
    'Se construye un vector con una posición por campo. Cada posición contiene un vector con los valores de ese campo para cada actividad
    arrVal = Array(Range(rngTmlCod.Offset(1), rngTmlCod.Offset(intDim)).value, _
                    Range(rngActID.Offset(1), rngActID.Offset(intDim)).value, _
                    Range(rngWBS.Offset(1), rngWBS.Offset(intDim)).value, _
                    Range(rngStart.Offset(1), rngStart.Offset(intDim)).value, _
                    Range(rngFinish.Offset(1), rngFinish.Offset(intDim)).value, _
                    Range(rngStartBL.Offset(1), rngStartBL.Offset(intDim)).value, _
                    Range(rngFinishBL.Offset(1), rngFinishBL.Offset(intDim)).value, _
                    Range(rngStartAct.Offset(1), rngStartAct.Offset(intDim)).value, _
                    Range(rngFinishAct.Offset(1), rngFinishAct.Offset(intDim)).value, _
                    Range(rngProgress.Offset(1), rngProgress.Offset(intDim)).value, _
                    Range(rngBdgUnt.Offset(1), rngBdgUnt.Offset(intDim)).value, _
                    Range(rngRmgUnt.Offset(1), rngRmgUnt.Offset(intDim)).value, _
                    Range(rngFloat.Offset(1), rngFloat.Offset(intDim)).value, _
                    Range(rngTmlMod.Offset(1), rngTmlMod.Offset(intDim)).value, _
                    Range(rngDesc.Offset(1), rngDesc.Offset(intDim)).value)

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
        arrAux(15) = IIf(IsEmpty(intRowUpdate), True, False)
        'Se asigna el vector auxiliar con los valores correspondientes a cada campo a la posición del vector correspondiente a la fila
        arrData(i) = arrAux
    Next

    If Not IsEmpty(intRowUpdate) Then
        For Each row In intRowUpdate
            arrData(row - rngRef.row - 1)(15) = True
        Next
    End If

    'Posiciones del subvector dentro de cada posición del vector arrData
    '0 --> Timeline Code
    '1 --> Activity ID
    '2 --> WBS
    '3 --> Start
    '4 --> Finish
    '5 --> Start BL
    '6 --> Finish BL
    '7 --> Start Actual
    '8 --> Finish Actual
    '9 --> Progress
    '10 --> Budget Unit
    '11 --> Remaining Unit
    '12 --> Float
    '13 --> Timeline Mode
    '14 --> Description
    'Dimensiones extra
    '15 --> Actividad actualizable segun vector intRowUpd
End Sub

Public Sub ContentsWBS(Optional ByVal intRowWBS As Variant = Empty)
    If intEdition = 1 Then Exit Sub
    
    Dim intR, intRchart, intArrPos As Integer
    Dim varStart, varFinish, varStartBL, varFinishBL, varStartAct, varFinishAct, varProgress, varBdgUnt, varRmgUnt, varFloat As Variant
    Dim intRowUpdWBS As Variant

    'Recorrer lista de actividades
    SetActArray intRowWBS
    intActLastRow = ActLastRow
    intR = rngRef.row + 1
    
    For Each arrAct In arrData
        'Comprobar si la actividad actual es una agrupación WBS
        If Not arrAct(1) Like "WBS-*" Then GoTo NextIteration
        'Comprobar si la actividad actual está en la lista de actividades a actualizar.
        If Not arrAct(15) Then GoTo NextIteration

        'Inicializar variables
        varStart = Empty
        varFinish = Empty
        varStartBL = Empty
        varFinishBL = Empty
        varStartAct = Empty
        varFinishAct = Empty
        varProgress = Empty
        varBdgUnt = Empty
        varRmgUnt = Empty
        varFloat = Empty

        intRchart = intR - rngRef.row
        For intArrPos = intR - rngRef.row To intActLastRow - rngRef.row - 1
            If (arrAct(2) = arrData(intArrPos)(2) Or arrData(intArrPos)(2) Like arrAct(2) & ".*") And Not arrData(intArrPos)(1) Like "WBS-*" And IsEmpty(arrData(intArrPos)(13)) Then
                If IsError(arrData(intArrPos)(3)) Then arrData(intArrPos)(3) = Empty
                If IsError(arrData(intArrPos)(4)) Then arrData(intArrPos)(4) = Empty
                If IsError(arrData(intArrPos)(5)) Then arrData(intArrPos)(5) = Empty
                If IsError(arrData(intArrPos)(6)) Then arrData(intArrPos)(6) = Empty
                If IsError(arrData(intArrPos)(7)) Then arrData(intArrPos)(7) = Empty
                If IsError(arrData(intArrPos)(8)) Then arrData(intArrPos)(8) = Empty
                'Selección del valor menor para Start en cada iteración
                If IsEmpty(varStart) And Not arrData(intArrPos)(3) = "" Then varStart = arrData(intArrPos)(3)
                If IsEmpty(varStart) And Not arrData(intArrPos)(4) = "" Then varStart = arrData(intArrPos)(4)
                If Not IsEmpty(arrData(intArrPos)(3)) And Not arrData(intArrPos)(3) = "" Then If varStart > arrData(intArrPos)(3) Then varStart = arrData(intArrPos)(3)
                If Not IsEmpty(arrData(intArrPos)(4)) And Not arrData(intArrPos)(4) = "" Then If varStart > arrData(intArrPos)(4) Then varStart = arrData(intArrPos)(4)
                'Selección del valor mayor para Finish en cada iteración
                If IsEmpty(varFinish) And Not arrData(intArrPos)(3) = "" Then varFinish = arrData(intArrPos)(3)
                If IsEmpty(varFinish) And Not arrData(intArrPos)(4) = "" Then varFinish = arrData(intArrPos)(4)
                If Not IsEmpty(arrData(intArrPos)(3)) And Not arrData(intArrPos)(3) = "" Then If varFinish < arrData(intArrPos)(3) Then varFinish = arrData(intArrPos)(3)
                If Not IsEmpty(arrData(intArrPos)(4)) And Not arrData(intArrPos)(4) = "" Then If varFinish < arrData(intArrPos)(4) Then varFinish = arrData(intArrPos)(4)
                'Selección del valor menor para Start BL en cada iteración
                If IsEmpty(varStartBL) And Not arrData(intArrPos)(5) = "" Then varStartBL = arrData(intArrPos)(5)
                If IsEmpty(varStartBL) And Not arrData(intArrPos)(6) = "" Then varStartBL = arrData(intArrPos)(6)
                If Not IsEmpty(arrData(intArrPos)(5)) And Not arrData(intArrPos)(5) = "" Then If varStartBL > arrData(intArrPos)(5) Then varStartBL = arrData(intArrPos)(5)
                If Not IsEmpty(arrData(intArrPos)(6)) And Not arrData(intArrPos)(6) = "" Then If varStartBL > arrData(intArrPos)(6) Then varStartBL = arrData(intArrPos)(6)
                'Selección del valor mayor para Finish BL en cada iteración
                If IsEmpty(varFinishBL) And Not arrData(intArrPos)(5) = "" Then varFinishBL = arrData(intArrPos)(5)
                If IsEmpty(varFinishBL) And Not arrData(intArrPos)(6) = "" Then varFinishBL = arrData(intArrPos)(6)
                If Not IsEmpty(arrData(intArrPos)(5)) And Not arrData(intArrPos)(5) = "" Then If varFinishBL < arrData(intArrPos)(5) Then varFinishBL = arrData(intArrPos)(5)
                If Not IsEmpty(arrData(intArrPos)(6)) And Not arrData(intArrPos)(6) = "" Then If varFinishBL < arrData(intArrPos)(6) Then varFinishBL = arrData(intArrPos)(6)
                'Selección del valor menor para Start Act en cada iteración
                If IsEmpty(varStartAct) And Not arrData(intArrPos)(7) = "" Then varStartAct = arrData(intArrPos)(7)
                If IsEmpty(varStartAct) And Not arrData(intArrPos)(8) = "" Then varStartAct = arrData(intArrPos)(8)
                If Not IsEmpty(arrData(intArrPos)(7)) And Not arrData(intArrPos)(7) = "" Then If varStartAct > arrData(intArrPos)(7) Then varStartAct = arrData(intArrPos)(7)
                If Not IsEmpty(arrData(intArrPos)(8)) And Not arrData(intArrPos)(8) = "" Then If varStartAct > arrData(intArrPos)(8) Then varStartAct = arrData(intArrPos)(8)
                'Selección del valor mayor para Finish Act en cada iteración
                If IsEmpty(varFinishAct) And Not arrData(intArrPos)(7) = "" Then varFinishAct = arrData(intArrPos)(7)
                If IsEmpty(varFinishAct) And Not arrData(intArrPos)(8) = "" Then varFinishAct = arrData(intArrPos)(8)
                If Not IsEmpty(arrData(intArrPos)(7)) And Not arrData(intArrPos)(7) = "" Then If varFinishAct < arrData(intArrPos)(7) Then varFinishAct = arrData(intArrPos)(7)
                If Not IsEmpty(arrData(intArrPos)(8)) And Not arrData(intArrPos)(8) = "" Then If varFinishAct < arrData(intArrPos)(8) Then varFinishAct = arrData(intArrPos)(8)
                'Suma de progreso
                varProgress = varProgress + arrData(intArrPos)(9) * arrData(intArrPos)(10)
                'Suma de Budget Units
                varBdgUnt = varBdgUnt + arrData(intArrPos)(10)
                'Suma de Remaining Units
                varRmgUnt = varRmgUnt + arrData(intArrPos)(11)
                'Selección del valor menor para Float BL en cada iteración
                If IsEmpty(varFloat) Then varFloat = arrData(intArrPos)(12)
                If Not IsEmpty(arrData(intArrPos)(12)) Then If varFloat > arrData(intArrPos)(12) Then varFloat = arrData(intArrPos)(12)
            End If
        Next
        If Not varBdgUnt = 0 Then varProgress = varProgress / varBdgUnt
        
        If arrAct(3) <> varStart Then rngStart.Offset(intRchart) = varStart
        If arrAct(4) <> varFinish Then rngFinish.Offset(intRchart) = varFinish
        If arrAct(5) <> varStartBL Then rngStartBL.Offset(intRchart) = varStartBL
        If arrAct(6) <> varFinishBL Then rngFinishBL.Offset(intRchart) = varFinishBL
        If arrAct(7) <> varStartAct Then rngStartAct.Offset(intRchart) = varStartAct
        If arrAct(8) <> varFinishAct Then rngFinishAct.Offset(intRchart) = varFinishAct
        If arrAct(9) <> varProgress Then rngProgress.Offset(intRchart) = varProgress
        If arrAct(10) <> varBdgUnt Then rngBdgUnt.Offset(intRchart) = varBdgUnt
        If arrAct(11) <> varRmgUnt Then rngRmgUnt.Offset(intRchart) = varRmgUnt
        If arrAct(12) <> varFloat Or (IsEmpty(arrAct(12)) And varFloat = 0) Then rngFloat.Offset(intRchart) = varFloat
        rngResume.Offset(intRchart) = ""
        
NextIteration:
        intR = intR + 1
    Next

End Sub

Sub colorWBS(ByVal level As Integer, rng As Range, ByVal dblCustomColor As Double, Optional booActLevel As Boolean = False)
    If intEdition = 1 Then Exit Sub
    
    If dblCustomColor <> RGB(255, 255, 255) Then
        rng.Interior.Color = dblCustomColor
    ElseIf booActLevel Then
        rng.Interior.Color = RGB(255, 255, 255)
    Else
        Select Case level Mod 10
        Case 0 'rojo
            rng.Interior.Color = RGB(255, 90, 51)
        Case 1 'azul
            rng.Interior.Color = RGB(102, 153, 255)
        Case 2 'verde
            rng.Interior.Color = RGB(0, 204, 102)
        Case 3 'amarillo
            rng.Interior.Color = RGB(255, 255, 153)
        Case 4 'marron
            rng.Interior.Color = RGB(204, 153, 0)
        Case 5 'naranja
            rng.Interior.Color = RGB(255, 153, 51)
        Case 6 'lila
            rng.Interior.Color = RGB(204, 153, 255)
        Case 7  'rosa
            rng.Interior.Color = RGB(255, 153, 204)
        Case 8 'azul claro
            rng.Interior.Color = RGB(153, 204, 255)
        Case 9 'verde claro
            rng.Interior.Color = RGB(153, 255, 153)
        End Select
    End If
End Sub

Function borderWBS(ByVal format As Integer, rng As Range)
    If intEdition = 1 Then Exit Function
    
    Select Case format
    Case 0
        With rng
            '.Borders(xlEdgeBottom).LineStyle = xlNone
            With .Borders(xlEdgeBottom)
                .LineStyle = xlDot
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With

    Case 1
        With rng
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With

    Case 2
        With rng
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    Case 3
        With rng
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    End Select
End Function

Function rngFormatWBS(ByVal strWsNumber As String) As Range
    If intEdition = 1 Then Exit Function
    
    Dim rng, rngSetup, rngFormat As Range
    Dim intFirstCol, intLastCol, intRow, c As Integer
    Dim n As Name
    
    Set rngSetup = Nothing
    Set rngFormat = Nothing
    intFirstCol = Empty
    intLastCol = Empty
    
    For Each n In ActiveWorkbook.Names
        If n.Name Like "VB_" & strWsNumber & "*" And Not n.Name Like "VB_" & strWsNumber & "_L*" Then
            Set rng = Range(n.Name)
            If CInt(Right(n.Name, 2)) < 8 Then
                If rngSetup Is Nothing Then
                    Set rngSetup = rng
                Else: Set rngSetup = Union(rngSetup, rng)
                End If
            End If
            
            If intFirstCol = Empty Then
                intFirstCol = rng.Column
            ElseIf intFirstCol > rng.Column Then
                intFirstCol = rng.Column
            End If
             If intLastCol = Empty Then
                intLastCol = rng.Column
            ElseIf intLastCol < rng.Column Then
                intLastCol = rng.Column
            End If
            intRow = rng.row
        End If
    Next
    'La última columna siempre será Period y vamos a editar hasta la anterior
    intLastCol = intLastCol - 1
    For c = intFirstCol To intLastCol
        Set rng = Cells(intRow, c)
        If Intersect(rngSetup, rng) Is Nothing Then
                If rngFormat Is Nothing Then
                    Set rngFormat = rng
                Else: Set rngFormat = Union(rngFormat, rng)
                End If
        End If
    Next
    Set rngFormatWBS = rngFormat
End Function


Sub WBS_indent(Optional booIndent As Boolean = True)
    If intEdition = 1 Then Exit Sub
    
    Dim rngArea, c As Range
    Dim intRowsArr, i, j  As Integer
    Dim strWBS, strWBSparent, strWBSfirstparent, strWBSupd As String
    Dim arrWBSfirstparent() As String
    Dim intWBSlevel, intWBSparentLevel As Integer
    
    SetPrjVar
    If Not booHeaders Then Exit Sub

    'Construir vector con filas del rango seleccionado
    ReDim intRowsArr(0)
    For Each rngArea In Application.Selection.Areas
        For Each c In rngArea
            If intRowsArr(0) = "" Then
                intRowsArr(0) = c.row
            Else
                For i = 1 To UBound(intRowsArr)
                    If intRowsArr(i) = c.row Then GoTo nextCell
                Next
                ReDim Preserve intRowsArr(UBound(intRowsArr) + 1)
                intRowsArr(UBound(intRowsArr)) = c.row
            End If
nextCell:
        Next
    Next
    
    'Recorrer filas
    For i = 0 To UBound(intRowsArr)
        'Si la fila no está dentro del rango de actividades se sale del procedimiento
        If intRowsArr(i) <= rngRef.row Or intRowsArr(i) > intActLastRow Then Exit Sub
        If intRowsArr(i) = rngRef.Offset(1).row Then
            If Len(rngWBS.Offset(1)) = 0 Then
                rngWBS.Offset(1) = "'1"
            Else: rngWBS.Offset(1) = "'" & Split(rngWBS.Offset(1), ".")(0)
            End If
        Else
            'Asignación de variable WBS padre y primer padre de la serie
            strWBSparent = rngWBS.Offset(intRowsArr(i) - 1 - rngRef.row)
            'Convertir WBS padre en cadena
            rngWBS.Offset(intRowsArr(i) - 1 - rngRef.row) = "'" & strWBSparent
            
            If i = 0 Then
                strWBSfirstparent = strWBSparent
                arrWBSfirstparent = Split(strWBSfirstparent, ".")
            End If
            'Si no hay WBS padre definido se sale del procedimiento
            If strWBSparent = "" Then
                CustomMsgBox "The WBS parent code is not defined."
                Exit Sub
            End If
            'Cálculo del nivel del WBS padre
            intWBSparentLevel = Len(strWBSparent) - Len(Replace(strWBSparent, ".", ""))
            
            'Cálculo del nivel del WBS actual
            strWBS = rngWBS.Offset(intRowsArr(i) - rngRef.row)
            If strWBS = "" Then
                intWBSlevel = -1
            Else: intWBSlevel = Len(strWBS) - Len(Replace(strWBS, ".", ""))
            End If
            
            'Acción según se vaya a crear un hijo o a recuperar el mismo nivel del padre
            If booIndent Then
                'Si el WBS actual tiene un nivel superior se iguala al padre y si no se convierte en su hijo
                If intWBSlevel < intWBSparentLevel Then
                    rngWBS.Offset(intRowsArr(i) - rngRef.row) = "'" & strWBSparent
                Else: rngWBS.Offset(intRowsArr(i) - rngRef.row) = "'" & strWBS & ".1"
                End If
            Else 'Si el WBS actual tiene un nivel inferior se iguala al padre y si no se convierte en un WBS nuevo
                If intWBSlevel > intWBSparentLevel Then
                    rngWBS.Offset(intRowsArr(i) - rngRef.row) = "'" & strWBSparent
                ElseIf intWBSlevel = intWBSparentLevel And InStr(strWBS, "-") = 0 Then
                    rngWBS.Offset(intRowsArr(i) - rngRef.row) = "'" & strWBSfirstparent & "-" & i + 1
                Else
                    strWBSupd = ""
                    For j = 0 To intWBSlevel - 1
                        strWBSupd = strWBSupd & arrWBSfirstparent(j) & IIf(j < intWBSlevel - 1, ".", "")
                    Next
                    If Not strWBSupd = "" Then rngWBS.Offset(intRowsArr(i) - rngRef.row) = "'" & strWBSupd & "-" & i + 1
                End If
            End If
        End If
    Next
End Sub

'Crear agrupaciones
Public Sub GroupWBS(ByVal strActID, ByVal strWBS As String, ByVal intRow As Integer, Optional booTml As Boolean = False)
    If intEdition = 1 Then Exit Sub
    
    Dim intOutlineLevel, intWBSlevel As Integer
    
    'Nivel de agrupación mínimo
    intOutlineLevel = 1
    
    'Nivel WBS
    If booGroupWBS Then
        intWBSlevel = 0
        If Not Len(strWBS) = 0 Then
            intWBSlevel = Len(strWBS) - Len(Replace(strWBS, ".", ""))
            If Not strActID Like "WBS-*" Then intWBSlevel = intWBSlevel + 1
        End If
        intOutlineLevel = intOutlineLevel + intWBSlevel
    End If
    
    'Nivel extra si la actividad está en un timeline
    If booTml Then intOutlineLevel = intOutlineLevel + 1
    
    'Actualizacion del nivel de agrupación
    If intOutlineLevel > 8 Then intOutlineLevel = 8
    If Not wsSch.Rows(intRow).OutlineLevel = intOutlineLevel Then
        wsSch.Rows(intRow).OutlineLevel = intOutlineLevel
    End If
    
End Sub
