VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "actObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

'Variables privadas para uso interno de la clase
Private arrPred() As Variant
Private arrSucc() As Variant

'Variables publicas
Public ActID As String
Public row As Long
Public RmgDuration As Integer
Public Start As Date
Public Finish As Date
Public ActualStart As Date
Public ActualFinish As Date
Public ResumeDate As Date
Public ConstraintDate As Date
Public EarlyStart As Date
Public EarlyFinish As Date
Public LateStart As Date
Public LateFinish As Date
Public Float As Variant
Public ActType As String
Public CnstType As String
Public TmlMode As String
Public fPredEmpty As Boolean
Public fSuccEmpty As Boolean
Public fLoopChkPh1 As Boolean
Public LoopNo As Integer
Public LoopPos As Integer
Public LoopStatusPh2 As Integer

Private Sub Class_Initialize()
    ReDim PredList(0)
    ReDim SuccList(0)
    Me.fPredEmpty = True
    Me.fSuccEmpty = True
    Me.fLoopChkPh1 = False
End Sub



Property Get CalcStart() As Variant
    If ActType = "FML" Then
        CalcStart = ""
    ElseIf ActualStart > 0 Then
        CalcStart = ActualStart
    ElseIf ConstraintStart > 0 Then
        CalcStart = ConstraintStart
    Else
        CalcStart = EarlyStart
    End If
End Property

Property Get CalcFinish() As Variant
    If ActType = "SML" Then
        CalcFinish = ""
    ElseIf ActualFinish > 0 Then
        CalcFinish = ActualFinish
    ElseIf ConstraintFinish > 0 Then
        CalcFinish = ConstraintFinish
    Else
        CalcFinish = EarlyFinish
    End If
End Property

Property Get CalcResumeDate() As Variant
    If Me.ActType = "ACT" And ActualStart > 0 And ActualFinish = 0 And _
        Me.EarlyStart > DateAddCal(1 - IIf(WorkingDay(datCutoff), 0, 1), datCutoff) Then
        CalcResumeDate = Me.EarlyStart
    Else: CalcResumeDate = ""
    End If
End Property

Property Get CalcFloat() As Variant
    If Me.fSchNo Then
        CalcFloat = "NO"
    ElseIf ActType = "SML" And ActualStart > 0 Or _
        ActType = "FML" And ActualFinish > 0 Or _
        ActType = "ACT" And ActualStart > 0 And ActualFinish > 0 Or _
        (Not IsEmpty(TmlMode) And Not TmlMode = "") Then
        CalcFloat = ""
    Else: CalcFloat = DateDiffCal(EarlyFinish, LateFinish)
    End If
End Property

Property Get ConstraintStart() As Date
    If Me.CnstType = "NO" Or Me.CnstType = "MANU" Then
        ConstraintStart = Me.Start
    End If
End Property

Property Get ConstraintFinish() As Date
    If Me.CnstType = "NO" Or Me.CnstType = "MANU" Then
        ConstraintFinish = Me.Finish
    End If
End Property

Property Get fEarlyDone() As Boolean
    If (ActType = "ACT" And EarlyStart > 0 And EarlyFinish > 0) Or _
        (ActType = "SML" And EarlyStart > 0) Or _
        (ActType = "FML" And EarlyFinish > 0) Then
        fEarlyDone = True
    Else
        fEarlyDone = False
    End If
End Property

Property Get fLateDone() As Boolean
    If (ActType = "ACT" And LateStart > 0 And LateFinish > 0) Or _
        (ActType = "SML" And LateStart > 0) Or _
        (ActType = "FML" And LateFinish > 0) Then
        fLateDone = True
    Else
        fLateDone = False
    End If
End Property

Property Get fSchNo() As Boolean
    If Me.CnstType = "NO" Then
        fSchNo = True
    Else
        fSchNo = False
    End If
End Property

Property Get PredList() As Variant
    PredList = arrPred
End Property

Property Get SuccList() As Variant
    SuccList = arrSucc
End Property


'Definición de métodos
Public Sub WriteProperties(ByVal shRow As Long, strActID, dblRmgDur, datStart, datFinish, datActStart, datActFinish, datResumeDate, _
                            datConstraintDate, varFloat, strSchMod, strPredList, strTmlMod As Variant)
    Dim strAct, strRelType, strLag As String
    Dim arrPredList() As String
    Dim i As Integer
    Dim dblLag As Double
    
    If strActID = Empty Then Exit Sub
    
    Me.ActID = IIf(IsError(strActID), "", strActID)
    Me.row = shRow
    Me.RmgDuration = IIf(IsError(dblRmgDur), 0, CInt(dblRmgDur))
    Me.Start = IIf(IsError(datStart), Empty, datStart)
    Me.Finish = IIf(IsError(datFinish), Empty, datFinish)
    Me.ActualStart = IIf(IsError(datActStart), Empty, datActStart)
    Me.ActualFinish = IIf(IsError(datActFinish), Empty, datActFinish)
    Me.ResumeDate = IIf(IsError(datResumeDate), Empty, datResumeDate)
    Me.ConstraintDate = IIf(IsError(datConstraintDate), Empty, datConstraintDate)
    Me.Float = varFloat
    Me.CnstType = Trim(Left(strSchMod, 4))
    Me.TmlMode = strTmlMod
    
    If Me.RmgDuration > 0 Or (Me.ActualStart > 0 And Me.ActualFinish > 0) Then
        Me.ActType = "ACT"
    ElseIf Me.Start > 0 Or Me.ActualStart > 0 Then
        Me.ActType = "SML"
    Else
        Me.ActType = "FML"
    End If
    
    strPredList = IIf(IsError(strPredList), Empty, strPredList)
    If Len(strPredList) > 0 Then
        arrPredList = Split(strPredList, ",")
        For i = 0 To UBound(arrPredList)
            arrPredList(i) = Trim(arrPredList(i))
        
            strAct = Split(arrPredList(i), " ")(0)
            If InStr(1, arrPredList(i), " ") = 0 Then
                strRelType = "FS"
                dblLag = 0
            Else
                strRelType = Left(Split(arrPredList(i), " ")(1), 2)
                strLag = Mid(Split(arrPredList(i), " ")(1), 3, 20)
                dblLag = 0
                If Len(strLag) > 0 And IsNumeric(strLag) Then dblLag = CDbl(strLag)
            End If
            
            ReDim Preserve arrPred(i)
            Set arrPred(i) = New relObj
    
            arrPred(i).ActID = strAct
            arrPred(i).RelType = strRelType
            arrPred(i).Lag = dblLag
        Next
        Me.fPredEmpty = False
    End If

End Sub

Public Sub RedimSuccList(ByVal i As Integer)
    ReDim Preserve arrSucc(i)
    Set arrSucc(i) = New relObj
End Sub

