VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalExc 
   Caption         =   "Calendar Exceptions"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5300
   OleObjectBlob   =   "frmCalExc.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCalExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Private Sub btnAdd_Click()
    Dim datRef As Date
    Dim arrCalExc As Variant
    Dim datExc As Variant
    Dim i As Integer
    
    If g_oDP.PickerVisible Then closeDatePicker
    
    If Not (IsDate(Me.txtStart) And IsDate(Me.txtFinish)) Then
        Me.txtStart = ""
        Me.txtFinish = ""
    Else
        ReDim arrCalExc(0)
        'Volcado de fechas en lista a vector
        For i = 0 To lstCalExc.ListCount - 1
            If IsEmpty(arrCalExc(0)) Then
                arrCalExc(0) = CDate(lstCalExc.List(i))
            Else
                ReDim Preserve arrCalExc(i)
                arrCalExc(i) = CDate(lstCalExc.List(i))
            End If
        Next
        'Volcado de fechas nuevas a vector
        For datRef = Me.txtStart To Me.txtFinish Step 1
            'Si la fecha ya está en el vector no se añade
            For Each datExc In arrCalExc
                If Not IsEmpty(datExc) Then
                    If datExc = datRef Then GoTo NextIteration
                End If
            Next
            If IsEmpty(arrCalExc(0)) Then
                arrCalExc(0) = datRef
            Else
                ReDim Preserve arrCalExc(UBound(arrCalExc) + 1)
                arrCalExc(UBound(arrCalExc)) = datRef
            End If
NextIteration:
        Next
        'Se ordena el vector
        SortArrayAscendent arrCalExc
        'Se añade el vector a la lista
        With lstCalExc
            .Clear
            For i = 0 To UBound(arrCalExc)
                .AddItem format(arrCalExc(i), "dd-mmm-yyyy")
            Next
        End With
    End If
    
End Sub

Private Sub btnDel_Click()
    Dim i As Integer
    
    If g_oDP.PickerVisible Then closeDatePicker
    
    For i = lstCalExc.ListCount - 1 To 0 Step -1
        If lstCalExc.Selected(i) = True Then
            lstCalExc.RemoveItem (i)
        End If
    Next
End Sub

Private Sub btnAccept_Click()
    Dim strCalExc As String
    Dim arrCalExc As Variant
    Dim varDateExc As Variant
    Dim i As Integer
    
    If g_oDP.PickerVisible Then closeDatePicker
    
    
    ReDim arrCalExc(0)
    For i = 0 To lstCalExc.ListCount - 1
        If Not IsEmpty(arrCalExc(0)) Then
            ReDim Preserve arrCalExc(i)
        End If
        arrCalExc(i) = lstCalExc.List(i)
    Next
    
    If Not IsEmpty(arrCalExc(0)) Then
        For Each varDateExc In arrCalExc
            If Len(strCalExc) = 0 Then
                strCalExc = format(varDateExc, "yyyymmdd")
            Else
                strCalExc = strCalExc & ", " & format(varDateExc, "yyyymmdd")
            End If
        Next
    Else
        strCalExc = ""
    End If
    updateCustomDocumentProperty "cdpCalExc", strCalExc, msoPropertyTypeString
    Unload Me
End Sub

Private Sub btnCancel_Click()
    If g_oDP.PickerVisible Then closeDatePicker
    Unload Me
End Sub


Private Sub btnFinishPicker_Click()
    ensureDPManager
    If g_oDP.PickerVisible Then
        closeDatePicker
    Else
        fDPfinishexc = True
        showDatePicker
    End If
End Sub

Private Sub btnStartPicker_Click()
    ensureDPManager
    If g_oDP.PickerVisible Then
        closeDatePicker
    Else
        fDPstartexc = True
        showDatePicker
    End If
End Sub

Private Sub lstCalExc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnDel_Click
End Sub


Private Sub lstCalExc_Enter()
If g_oDP.PickerVisible Then closeDatePicker
End Sub

Private Sub txtFinish_AfterUpdate()
FinishExcUpdate
End Sub

Private Sub txtFinish_Enter()
    If g_oDP.PickerVisible Then closeDatePicker
End Sub

Private Sub txtStart_AfterUpdate()
StartExcUpdate
End Sub

Private Sub txtStart_Enter()
If g_oDP.PickerVisible Then closeDatePicker
End Sub

Private Sub UserForm_Click()
If g_oDP.PickerVisible Then closeDatePicker
End Sub
