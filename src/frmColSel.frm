VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColSel 
   Caption         =   "Columns Selection"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6980
   OleObjectBlob   =   "frmColSel.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmColSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Private Sub btnChngStatus_Click()
    Dim i As Integer
    
    For i = 0 To lstColumns.ListCount - 1
        If lstColumns.Selected(i) = True Then
            lstColumns.List(i, 1) = IIf(lstColumns.List(i, 1) = "SHOWN", "hidden", "SHOWN")
        End If
    Next
    Me.cboLayout = "Custom"
End Sub

Private Sub btnUp_Click()
    Dim i, j, k As Integer
    Dim strColUp As String
    Dim intSelection() As Integer
    
    j = 0
    For i = 0 To lstColumns.ListCount - 1
        If lstColumns.Selected(i) = True Then
            If i = 0 Then Exit Sub
            
            For k = 0 To 2
                strColUp = lstColumns.List(i, k)
                lstColumns.List(i, k) = lstColumns.List(i - 1, k)
                lstColumns.List(i - 1, k) = strColUp
            Next
            strColUp = lstColumns.List(i, 3) - 1
            lstColumns.List(i, k) = lstColumns.List(i - 1, k) + 1
            lstColumns.List(i - 1, k) = strColUp
            
            
            lstColumns.Selected(i) = False

            ReDim Preserve intSelection(j)
            intSelection(j) = i - 1
            j = j + 1
        End If
    Next
    For j = 0 To UBound(intSelection)
        lstColumns.Selected(intSelection(j)) = True
    Next
End Sub

Private Sub btnDwn_Click()
    Dim i, j, k As Integer
    Dim strColUp, strColDwn As String
    Dim intSelection() As Integer
    
    For i = lstColumns.ListCount - 1 To 0 Step -1
        If lstColumns.Selected(i) = True Then
            If i = lstColumns.ListCount - 1 Then Exit Sub
            
            For k = 0 To 2
                strColDwn = lstColumns.List(i, k)
                lstColumns.List(i, k) = lstColumns.List(i + 1, k)
                lstColumns.List(i + 1, k) = strColDwn
            Next
            strColDwn = lstColumns.List(i, 3) + 1
            lstColumns.List(i, 3) = lstColumns.List(i + 1, 3) - 1
            lstColumns.List(i + 1, 3) = strColDwn

            
            lstColumns.Selected(i) = False
            
            ReDim Preserve intSelection(j)
            intSelection(j) = i + 1
            j = j + 1
        End If
    Next
    
    For j = 0 To UBound(intSelection)
        lstColumns.Selected(intSelection(j)) = True
    Next
End Sub

Private Sub btnApply_Click()
    Dim i, j As Integer
    Dim strListNms(), strListRef() As String
    Dim strRef As String
    Dim nm As Name
    Dim booHidden As Boolean
    Dim rngInsert As Range
    Dim strColumnsArr As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    GetHeaderArray
    ReDim strColumnsArr(UBound(strHeadArr) - 1, 3)
    For i = 0 To UBound(strColumnsArr, 1)
        'Guardamos valores en la lista en un vector
        strColumnsArr(i, 0) = lstColumns.List(i, 0)
        strColumnsArr(i, 1) = IIf(lstColumns.List(i, 1) = "hidden", True, False)
        strColumnsArr(i, 2) = CInt(lstColumns.List(i, 2))
        strColumnsArr(i, 3) = CInt(lstColumns.List(i, 3))
        'Inicializamos la columna de origen a la columna de destino
        lstColumns.List(i, 2) = lstColumns.List(i, 3)
    Next
    
    '2: Columna Origen (ajustar)
    '3: Columna Destino
    'Trabajamos con el vector
    'Para cada columna
    For i = 0 To UBound(strColumnsArr, 1)
        'Si su posición de origen no es igual a su posición de destino se ha traído de una columna mayor
        If Not strColumnsArr(i, 2) = strColumnsArr(i, 3) Then
            'Cortar en origen e insertar columna en destino
            ActiveSheet.Cells(1, strColumnsArr(i, 2)).EntireColumn.Cut
            ActiveSheet.Cells(1, strColumnsArr(i, 3)).EntireColumn.Insert Shift:=xlToRight
            'Para cada columna con posición de destino superior hay que ajustar la posición de origen
            For j = i + 1 To UBound(strColumnsArr, 1)
                If strColumnsArr(j, 2) < strColumnsArr(i, 2) Then strColumnsArr(j, 2) = strColumnsArr(j, 2) + 1
            Next
        End If
        'Show/Hide
        ActiveSheet.Cells(1, strColumnsArr(i, 3)).EntireColumn.Hidden = strColumnsArr(i, 1)
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Private Sub btnAccept_Click()
    btnApply_Click
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lstColumns_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnChngStatus_Click
End Sub

Private Sub cboLayout_Change()
    Dim arrShow As Variant
    Dim strNameRef, strNameShw As Variant
    Dim i As Integer
    
    Select Case cboLayout.ListIndex
    Case 0 'All Columns
        arrShow = Array("00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", _
                        "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25")
    Case 1 'Draw Chart
        arrShow = Array("00", "01", "08", "10", "11", "12", "13", "14", "15", "16")
    Case 2 'Draw Timeline
        arrShow = Array("00", "01", "03", "04", "05", "08", "10", "13", "14")
    Case 3 'Schedule Project
        arrShow = Array("06", "08", "10", "11", "12", "13", "14", "17", "18", "20", "21", "22")
    Case 4 'Progress & Units
        arrShow = Array("07", "08", "10", "13", "14", "23", "24", "25")
    Case 5 'WBS
        arrShow = Array("08", "09", "10", "11", "12", "13", "14")
    Case Else 'Custom
        Exit Sub
    End Select

    For i = 0 To lstColumns.ListCount - 1
        strNameRef = Right(lstColumns.List(i, 4), 2)
        lstColumns.List(i, 1) = "hidden"
        For Each strNameShw In arrShow
            If strNameShw = strNameRef Then lstColumns.List(i, 1) = "SHOWN"
        Next
    Next
End Sub
