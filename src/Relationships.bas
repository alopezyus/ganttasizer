Attribute VB_Name = "Relationships"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

'Disabled in Free Edition
Option Explicit

Sub MngRel(Optional booDelete As Boolean = False)
    If intEdition = 1 Then Exit Sub
    
    Dim rngArea, rngUnion, c As Range
    Dim intRowsArr, i  As Integer
    
    SetPrjVar
    If Not booHeaders Then Exit Sub
    
    Set rngUnion = Nothing
    
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
    
    For i = 1 To UBound(intRowsArr)
        If booDelete Then
            DelRel intRowsArr(i - 1), intRowsArr(i)
        Else: AddRelFS intRowsArr(i - 1), intRowsArr(i)
        End If
    Next
End Sub

Private Sub AddRelFS(ByVal intRowPred, ByVal intRowSucc As Integer)
    If intEdition = 1 Then Exit Sub
    
    Dim strPred, strSucc, strAdd, strInput As String
    
    strPred = rngActID.Offset(intRowPred - rngRef.row)
    strSucc = rngActID.Offset(intRowSucc - rngRef.row)
    
    If Len(strPred) = 0 Or Len(strSucc) = 0 Then Exit Sub
    If intRowPred = intRowSucc Then Exit Sub

    strAdd = strPred & " " & strRelType & IIf(intRelLag = 0, "", intRelLag)

    strInput = rngPred.Offset(intRowSucc - rngRef.row).value
    
    If InStr(1, strInput, strPred & " " & strRelType) > 0 Then Exit Sub
    
    If strInput = Empty Then
        strInput = strAdd
    Else
        strInput = strInput & ", " & strAdd
    End If
    
    rngPred.Offset(intRowSucc - rngRef.row, 0).value = strInput
    
End Sub


Private Sub DelRel(ByVal intRowPred, ByVal intRowSucc As Integer)
    If intEdition = 1 Then Exit Sub
    
    Dim strPred, strSucc, strDel, strInput, strRel As String
    Dim intRow, i As Integer
    Dim strArray() As String
    
    strPred = rngActID.Offset(intRowPred - rngRef.row)
    strSucc = rngActID.Offset(intRowSucc - rngRef.row)
    strRel = "FS"
    
    If Len(strPred) = 0 Then Exit Sub
    If Len(strSucc) = 0 Then
        rngPred.Offset(intRowSucc - rngRef.row, 0).value = ""
        Exit Sub
    End If

    
    strDel = strPred '& " " & strRel
    
    intRow = rngActID.EntireColumn.Find(strSucc, LookAt:=xlWhole).row
    strInput = rngPred.Offset(intRow - rngRef.row, 0).value
    
    'Vector con todas las predecesoras de esta fila
    strArray() = Split(strInput, ",")
    strInput = Empty
    'Quitar espacios por delante y por detrás
    For i = 0 To UBound(strArray)
        strArray(i) = Trim(strArray(i))
        If InStr(1, strArray(i), strDel) = 0 Then
            strInput = IIf(strInput = Empty, "", strInput & ", ") & strArray(i)
        End If
    Next

    rngPred.Offset(intRow - rngRef.row, 0).value = strInput

End Sub

