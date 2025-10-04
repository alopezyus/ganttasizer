Attribute VB_Name = "Columns"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Dim strRefHidden(), strRefShown(), strNmsHidden(), strNmsShown()  As String
Dim booAllHidden, booAllShown As Boolean

Public Sub openColSelForm()
    Dim i As Integer
    Dim strColumnsArr As Variant
    Dim item As Variant
    
    SetPrjVar
    If Not booHeaders Then Exit Sub
    
    strColumnsArr = getColumnsArr
    
    With frmColSel
        With .lstColumns
            For i = 0 To UBound(strColumnsArr)
                .AddItem
                .List(i, 0) = strColumnsArr(i, 3) 'título visible al usuario
                .List(i, 1) = IIf(strColumnsArr(i, 4), "hidden", "SHOWN") 'Hidden/Shown
                .List(i, 2) = strColumnsArr(i, 0) 'Columna original
                .List(i, 3) = strColumnsArr(i, 0) 'Columna a modificar
                .List(i, 4) = strColumnsArr(i, 1) 'Nombre
            Next
            .MultiSelect = 2
        End With
        .cboLayout.List = Array("All Columns", "Draw Chart", "Draw Timeline", "Schedule Project", "Progress & Units", "WBS", "Custom")
        .Show
    End With
End Sub

Public Function getColumnsArr() As Variant
    Dim nm As Name
    Dim strRef As String
    Dim i, j As Integer
    Dim strColumnsArr As Variant
    
    SetPrjVar
    GetHeaderArray
    ReDim strColumnsArr(UBound(strHeadArr) - 1, 4)
    i = 0
    'Recorrer todos los nombres en el libro
    For Each nm In ActiveWorkbook.Names
        'Entrar solo si el nombre empieza por "VB*" y está en la hoja activa y no es Period
        If nm.Name Like "VB*" And InStr(nm, ActiveSheet.Name) > 0 And Not nm.Name Like "VB_*_L??" Then
            If Not nm.RefersToRange Like rngPeriod.value Then
                'Guardar el rango al que apunta el nombre
                strRef = Replace(nm, "=", "")
                'Guardar en vector y si el rango está visible
                strColumnsArr(i, 0) = Range(strRef).Column
                strColumnsArr(i, 1) = nm.Name
                strColumnsArr(i, 2) = strRef
                strColumnsArr(i, 3) = Range(strRef).Value2
                strColumnsArr(i, 4) = ActiveSheet.Range(strRef).EntireColumn.Hidden
                i = i + 1
            End If
        End If
    Next
    getColumnsArr = SortArrayAscendentMulti(strColumnsArr)
End Function

