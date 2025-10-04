Attribute VB_Name = "Exceptions"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition and Pro Edition
Option Explicit

Public Sub openCalendarExceptions()
    If intEdition > 0 Then Exit Sub
    
    Dim arrDatesExc As Variant
    Dim i As Integer
    
    On Error GoTo errHandler
    arrDatesExc = Split(ActiveWorkbook.CustomDocumentProperties("cdpCalExc").value, ",")
    
    With frmCalExc
        With .lstCalExc
            For i = 0 To UBound(arrDatesExc)
                arrDatesExc(i) = Trim(arrDatesExc(i))
                If Len(arrDatesExc(i)) = 8 Then
                    .AddItem format(CDate(Right(arrDatesExc(i), 2) & "-" & Mid(arrDatesExc(i), 5, 2) & "-" & Left(arrDatesExc(i), 4)), "dd-mmm-yyyy")
                End If
            Next
            .MultiSelect = 2
        End With
        .Show
    End With
    
    Exit Sub
errHandler:
    updateCustomDocumentProperty "cdpCalExc", "", msoPropertyTypeString
    ReDim datDatesExc(0)
    frmCalExc.Show
End Sub

Public Sub StartExcUpdate()
    If intEdition > 0 Then Exit Sub
    If Not IsDate(frmCalExc.txtStart) Then
        frmCalExc.txtStart = ""
    Else
        frmCalExc.txtStart = format(frmCalExc.txtStart, "dd/mmm/yyyy")
        frmCalExc.txtFinish = frmCalExc.txtStart
    End If
End Sub

Public Sub FinishExcUpdate()
    If intEdition > 0 Then Exit Sub
    If Not IsDate(frmCalExc.txtFinish) Or Not IsDate(frmCalExc.txtStart) Then
        frmCalExc.txtFinish = ""
    ElseIf CDate(frmCalExc.txtFinish) < CDate(frmCalExc.txtStart) Then
        frmCalExc.txtFinish = frmCalExc.txtStart
    Else
        frmCalExc.txtFinish = format(frmCalExc.txtFinish, "dd/mmm/yyyy")
    End If
End Sub
