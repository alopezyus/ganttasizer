Attribute VB_Name = "Template"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

'Devueleve un vector: en la posición 0 si la hoja ya tiene en cabezados y en la posición 1 se encuentra el código de la hoja (el propio código o el que le corresponde añadir)
Public Function returnName(Optional ByVal booCreateHeaders As Boolean = False) As Variant
    Dim n As Name
    Dim strName, strWsCode As String
    Dim intCode As Integer
    Dim varReturn(1) As Variant
    Dim wsCurr As Worksheet
        
    Set wsCurr = IIf(booCreateHeaders, ActiveSheet, wsSch)
    
    intCode = 0
    For Each n In ActiveWorkbook.Names
        strName = n.Name
        If Left(strName, 3) = "VB_" Then
            strWsCode = Mid(strName, 4, InStr(4, strName, "_") - 4)
    
             If InStr(1, Replace(Replace(n.RefersTo, "=", ""), "'", ""), wsCurr.Name & "!") = 1 Then
                varReturn(0) = True
                GoTo endFunct
            End If
            
            If InStr(1, n.RefersTo, "#REF!") > 0 Then
                n.Delete
            End If
            
            If intCode < CInt(strWsCode) Then intCode = CInt(strWsCode)
        End If
    Next
    
    strWsCode = format(intCode + 1, "000")
    varReturn(0) = False
    
endFunct:
    varReturn(1) = strWsCode
    returnName = varReturn
End Function

Public Sub NewSheet()
    'Limit to 3 Ganttasizer Sheets in Free Edition
    If intEdition = 1 Then
        Dim n As Name
        Dim cntVB As Integer
        cntVB = 0
        For Each n In ActiveWorkbook.Names
            If n.Name Like "VB_*_00" Then cntVB = cntVB + 1
        Next
        If cntVB >= 3 Then
            CustomMsgBox "To create more Gantt charts try the Pro Edition."
            End
        End If
    End If
    
    Worksheets.Add(After:=ActiveSheet).Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub CreateHeaders()
    Dim intC As Integer
    Dim wsCurr As Worksheet
    Dim rngHead As Range
    Dim varReturn As Variant
    
    varReturn = returnName(True)
    
    If varReturn(0) Then
        'UpdateProgressBar 1
        CustomMsgBox "There are already defined headers in this worksheet."
        Exit Sub
    End If
    
    Set wsCurr = ActiveSheet
    
    GetHeaderArray
    
    For intC = 0 To UBound(strHeadArr)
        Set rngHead = wsCurr.Cells(1, 1).Offset(0, intC)
        rngHead = strHeadArr(intC)
        ActiveWorkbook.Names.Add Name:="VB_" & varReturn(1) & "_" & format(intC, "00"), RefersToR1C1:=rngHead
        
        'Formatear celda
        With rngHead
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = IIf(intC = 26, 90, 0)
        End With
        With rngHead.Font
            .Name = "Calibri"
            .Italic = IIf(intC <= 7, True, False)
            .Size = IIf(intC <= 7 Or intC = 26, 10, 11)
        End With
    
        With rngHead.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngHead.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngHead.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With rngHead.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        rngHead.ColumnWidth = IIf(intC = 26, 3, IIf(intC <= 7, 7, 14))
        
        'Free Edition and Pro Edition: Hide columns not used
        If intEdition = 1 Then
            Select Case intC
            Case 2, 6, 7, 9, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25
                rngHead.Interior.Color = RGB(165, 165, 165)
                rngHead.EntireColumn.Hidden = True
            End Select
        End If
        If intEdition = 2 Then
            Select Case intC
            Case 6, 7, 17, 18, 19, 20, 22, 25
                rngHead.Interior.Color = RGB(165, 165, 165)
                rngHead.EntireColumn.Hidden = True
            End Select
        End If
        
        'UpdateProgressBar 0.9 * intC / UBound(strHeadArr)
    Next
    
    wsCurr.Range(Cells(1, 1), Cells(1, 8)).Columns.group
    rngHead.RowHeight = 36
    
    'UpdateProgressBar 1
End Sub

'Disabled in Free Edition
Public Sub CopySheet()
    If intEdition = 1 Then Exit Sub
    
    Dim wsCurr, wsNew As Worksheet
    Dim nm As Name
    Dim varReturn, varName As Variant
    Dim intShCode As Integer
    Dim strShCode As String
    Dim rngCurrWS As Range

    Set wsCurr = ActiveSheet
    ActiveSheet.Copy After:=ActiveSheet
    Set wsNew = ActiveSheet
    UpdateProgressBar 0.3
    
    intShCode = 0
    For Each nm In ActiveWorkbook.Names
        If InStr(nm.RefersTo, wsNew.Name) > 0 Then
            nm.Delete
        ElseIf Left(nm.Name, 3) = "VB_" Then
            If CInt(Split(nm.Name, "_")(1)) > intShCode Then
                intShCode = CInt(Split(nm.Name, "_")(1))
            End If
        End If
    Next
    strShCode = format(intShCode + 1, "000")
    UpdateProgressBar 0.6
    
    For Each nm In ActiveWorkbook.Names
        If InStr(nm.RefersTo, wsCurr.Name) > 0 And InStr(nm.Name, "VB_") > 0 Then
            Set rngCurrWS = Range(Replace(nm.RefersTo, "=", ""))
            ActiveWorkbook.Names.Add Name:="VB_" & strShCode & "_" & Split(nm.Name, "_")(2), RefersTo:=wsNew.Cells(rngCurrWS.row, rngCurrWS.Column)
        End If
    Next
End Sub

