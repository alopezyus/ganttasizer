VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShapeEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private old_selection As Object
Private WithEvents bars As CommandBars
Attribute bars.VB_VarHelpID = -1

Public SheetShapesDict As Object

Private Sub Class_Initialize()
  Set bars = Application.CommandBars
End Sub

Private Sub bars_OnUpdate()
  Dim xSel As Object
  Set xSel = Selection

  If isShape(Selection) Then 'Or isShape(old_selection)
    'If selection is a shape then it could have changed or been selected,
    'otherwise if the old selection contained a shape and the new doesn't then
    'shape has been deselected
    If DetectShape(Selection) Then
      'Use the name to decide if it has been changed or selected
      If GetName(old_selection) = GetName(Selection) Or Not isShape(old_selection) Then
        'Raise Changed event - doesn't actually imply the shape changed...
        'if hash(shp) <> old_hash then ...
          'RaiseEvent Changed(Selection.ShapeRange(1))
          Changed Selection.ShapeRange(1)
      End If
    End If
  End If
  
  'Keep track of old selection
  Set old_selection = Selection
End Sub

Public Sub HookSheet(ByVal sht As Worksheet)
  If SheetShapesDict Is Nothing Then Set SheetShapesDict = CreateObject("Scripting.Dictionary")
  Set SheetShapesDict(sht.CodeName) = CreateObject("Scripting.Dictionary")
  SheetShapesDict(sht.CodeName)("=COUNT") = sht.Shapes.Count
  

  Dim shp As shape
  For Each shp In sht.Shapes
    Set SheetShapesDict(sht.CodeName)(shp.Name) = shp
  Next
End Sub
Public Sub UnhookSheet(ByVal sht As Worksheet)
  Call SheetShapesDict.Remove(sht.Name)
End Sub
Public Sub UnhookListener()
  Set bars = Nothing
End Sub

Private Function getShapeByName(sht As Worksheet, ByVal sName As String) As shape
  Dim shp As shape
  For Each shp In sht.Shapes
    If shp.Name = sName Then
      Set getShapeByName = shp
      Exit Function
    End If
  Next
  Set getShapeByName = Nothing
End Function

Private Function GetName(ByVal obj As Object) As String
  On Error Resume Next
  GetName = obj.Name
End Function

Private Function DetectShape(ByVal obj As Object) As Boolean
  On Error GoTo endDetect
    DetectShape = obj.ShapeRange.Count > 0
endDetect:
End Function

Private Function isShape(ByVal obj As Object) As Boolean
  Select Case TypeName(obj)
    Case "Rectangle", "Arc", "Drawing", "Picture", "Line", "Oval"
      isShape = True
    Case Else
      isShape = False
  End Select
End Function

Function ShapeData(shp As shape) As String
  Dim s As String

  With shp
    s = .Top & "," & .Left & "," & .Height & "," & .Width & "," & .AlternativeText & "," & .Name
    With .Fill
      s = s & "," & .BackColor.RGB & "," & .ForeColor.RGB
    End With
    With .Line
      s = s & "," & .BackColor.RGB & "," & .ForeColor.RGB
    End With
    s = s & "," & .Glow.Color.RGB
    With .TextFrame2.TextRange
      With .Font.Fill
        s = s & "," & .BackColor.RGB & "," & .ForeColor.RGB
      End With
      s = s & "," & .text
    End With
  End With

  ShapeData = s
End Function


Private Sub Changed(shape As shape)
    Dim intRow, intRowSum, intWindowRows As Integer
    Dim datStart, datFinish, datFirstChk As Variant
    Dim datStartTable, datFinishTable, datResumeTable, datStartChart, datFinishChart As Variant
    Dim booBL, booAct, booMil, booSum, booWin, booPrg, booFlt, booSameY, booSumBar, booResume As Boolean
    Dim dblShpLeft As Double
    Dim k As Integer
    Dim intRmgDur As Integer

    Application.EnableEvents = False

    'Set Cutoff date
    datCutoff = SetCutoff
    'Set Calendar Exceptions
    arrDatesExc = GetCalendarExceptions
    
    killGridDP
    
    'Nombre de la forma debe empezar por VB_
    If Not Left(shape.Name, 3) = "VB_" Then GoTo subTermination

    If Not booPrjStartSet Then
        SetPrjVar
        booPrjStartSet = True
    End If
    datFirstChk = datPrjFirst
    If Not booHeaders Then GoTo subTermination
    
    booPrg = False
    'Selección de casos posibles para el segundo código
    Select Case Split(shape.Name, "_")(1)
    Case "CUTOFF"
        If GetDate(shape.Left, False) <> datCutoff Then
            shape.Delete
            If IsDate(datCutoff) Then InsertCutoff
        End If
        GoTo subTermination
    Case "ACT"
        booBL = False
        booAct = True
        booMil = False
        booSum = False
    Case "MLA"
        booBL = False
        booAct = True
        booMil = True
        booSum = False
    Case "REM"
        booBL = False
        booAct = False
        booMil = False
        booSum = False
    Case "MLR"
        booBL = False
        booAct = False
        booMil = True
        booSum = False
    Case "BL0"
        booBL = True
        booAct = False
        booMil = False
        booSum = False
    Case "BM0"
        booBL = True
        booAct = False
        booMil = True
        booSum = False
    Case "SMA"
        booBL = False
        booAct = True
        booMil = True
        booSum = True
    Case "SMR"
        booBL = False
        booAct = False
        booMil = True
        booSum = True
    Case "SM0"
        booBL = True
        booAct = False
        booMil = True
        booSum = True
    Case "SAA"
        booBL = False
        booAct = True
        booMil = False
        booSum = True
    Case "SAR"
        booBL = False
        booAct = False
        booMil = False
        booSum = True
    Case "SA0"
        booBL = True
        booAct = False
        booMil = False
        booSum = True
    Case "WIN"
        booWin = True
'        booBL = False
'        booAct = True '<-----------------------
'        booMil = False
'        booSum = False
    Case "PRA"
        booPrg = True
        booAct = True
    Case "PRR"
        booPrg = True
        booAct = False
    Case "FLT"
        booFlt = True
    Case Else 'Si no coincide con ninguno de los anteriores se sale del procedimiento
        GoTo subTermination
    End Select

    'Número de fila contenido en la tercera sección del código
    intRow = Split(shape.Name, "_")(2)

    'Asignación de fechas según booBL, booMil y booAct
    If booPrg Then
        datStartTable = GetPrgBarDate(intRow, True, Not booAct)
        datFinishTable = GetPrgBarDate(intRow, False, Not booAct)
        datStartChart = GetDate(shape.Left, True)
        datFinishChart = GetDate(shape.Left + shape.Width, False)
        datStart = datStartChart
        datFinish = datFinishChart
    ElseIf booFlt Then
        datStartTable = rngStart.Offset(intRow - rngRef.row, 0)
        datFinishTable = rngFinish.Offset(intRow - rngRef.row, 0)
        datStartTable = IIf(datFinishTable + 1 > datStartTable, datFinishTable + 1, datStartTable)
        datFinishTable = DateAddCal(rngFloat.Offset(intRow - rngRef.row, 0) - 1, datStartTable)
        datStartChart = GetDate(shape.Left, True)
        datFinishChart = GetDate(shape.Left + shape.Width, False)
        datStart = datStartChart
        datFinish = datFinishChart
    ElseIf booWin Then
        datStartTable = rngStart.Offset(intRow - rngRef.row, 0)
        datFinishTable = rngFinish.Offset(intRow - rngRef.row, 0)
        datStartChart = GetDate(shape.Left, True)
        datFinishChart = GetDate(shape.Left + shape.Width, False)
        datStart = datStartChart
        datFinish = datFinishChart
    ElseIf booBL Then
        datStartTable = rngStartBL.Offset(intRow - rngRef.row, 0)
        datFinishTable = rngFinishBL.Offset(intRow - rngRef.row, 0)
        If booMil Then
            If IsDate(rngFinishBL.Offset(intRow - rngRef.row, 0)) Then
                datFinishChart = GetDate(shape.Left + shape.Width / 2, False)
                datStartChart = ""
            Else
                datFinishChart = ""
                datStartChart = GetDate(shape.Left + shape.Width / 2, True)
            End If
        Else
            datStartChart = GetDate(shape.Left, True)
            datFinishChart = GetDate(shape.Left + shape.Width, False)
        End If
        datStart = datStartChart
        datFinish = datFinishChart
    Else
        datStartTable = rngStart.Offset(intRow - rngRef.row, 0)
        datFinishTable = rngFinish.Offset(intRow - rngRef.row, 0)
        datResumeTable = rngResume.Offset(intRow - rngRef.row, 0)
        If booMil Then
            If IsDate(rngFinish.Offset(intRow - rngRef.row, 0)) Then
                datFinishChart = GetDate(shape.Left + shape.Width / 2, False)
                datStartChart = ""
            Else
                datFinishChart = ""
                datStartChart = GetDate(shape.Left + shape.Width / 2, True)
            End If
            datStart = datStartChart
            datFinish = datFinishChart
        Else
            datStartChart = GetDate(shape.Left, True)
            datFinishChart = GetDate(shape.Left + shape.Width, False)
            If booAct Then
                datStart = datStartChart
                If datFinishChart >= datCutoff And datFinishTable > datCutoff Then
                    datFinish = datFinishTable
                Else
                    datFinish = datFinishChart
                End If
            Else
                If IsDate(datResumeTable) And datResumeTable > datCutoff And datResumeTable <= datFinishTable Then
                    datStartTable = datResumeTable
                    booResume = True
                End If
                If datStartChart - 1 = datCutoff And datStartTable < datCutoff Then
                    datStart = datStartTable
                Else
                    datStart = datStartChart
                End If
                datFinish = datFinishChart
            End If
        End If
    End If

    'Modificación de condiciones particulares para booSum
    If booSum Then
        If datFinishChart = "" Then datFinishTable = ""
        If datStartChart = "" Then datStartTable = ""

        k = 0
        Do While rngTmlMod.Offset(intRow - k - rngRef.row, 0) = ""
            k = k + 1
        Loop
        intRowSum = rngTmlMod.Offset(intRow - k - rngRef.row, 0).row
    Else
        booSumBar = Len(rngTmlMod.Offset(intRow - rngRef.row, 0)) > 1
    End If

    'Booleano para la posición Y
    Dim lngTop, lngBottom, lngHeight As Long
    lngTop = GetPosY(IIf(booSum, intRowSum, intRow), True, booBL)
    lngBottom = GetPosY(IIf(booSum, intRowSum, intRow), False, booBL)
    lngHeight = lngBottom - lngTop
    If booPrg Then
        lngTop = lngTop + lngHeight / 4
        lngBottom = lngBottom - lngHeight / 4
    ElseIf booFlt Then
        lngTop = lngTop + lngHeight * 7 / 8
    End If
    
    booSameY = Abs(shape.Top - lngTop) < 1 And _
            Abs(shape.Top + shape.Height - lngBottom) < 1

    dblShpLeft = shape.Left + shape.Width

    'Para ventanas, calcular numero de filas
    If Not booSameY And booWin Then
        intWindowRows = 0
        Do While rngRef.Offset(CInt(intRow) - rngRef.row + intWindowRows + 1).Top < shape.Top + shape.Height
            intWindowRows = intWindowRows + 1
        Loop

        'Actualizar valor de numero de filas y booleano para recalculo de posicion Y
        Dim rngWindowRows As Range
        Set rngWindowRows = rngShpHgt.Offset(CInt(intRow) - rngRef.row)
        If (rngWindowRows.value = "" And intWindowRows < intActLastRow - rngRef.row) Or (rngWindowRows.value > 0 And intWindowRows <> rngWindowRows.value) And xl_UpdChart Then
            rngWindowRows = intWindowRows
            booSameY = False
        Else:
            If shape.Top <> rngRef.Offset(CInt(intRow) - rngRef.row).Top Then
                booSameY = False
            Else: booSameY = True
            End If
        End If
    End If
    
    'Ajustar manualmente fechas concretas cuando el detalle es pequeño (meses/años) es muy complicado
    'Además, genera modificaciones no desadas en las posiciones de las formas cuando estas se seleccionan
    'Por eso, cuando la fecha devuelta por la posición de la barra sea sólo de unos días de diferencia
    'Semantiene la fecha de la tabla de datos
    Select Case intPeriod
        Case 5
            If Not datStartTable = "" Then If Abs(datStartTable - datStart) <= 1 Then datStart = datStartTable
            If Not datFinishTable = "" Then If Abs(datFinishTable - datFinish) <= 1 Then datFinish = datFinishTable
        Case 6
            If Not datStartTable = "" Then If Abs(datStartTable - datStart) <= 3 Then datStart = datStartTable
            If Not datFinishTable = "" Then If Abs(datFinishTable - datFinish) <= 3 Then datFinish = datFinishTable
        Case 7
            If Not datStartTable = "" Then If Abs(datStartTable - datStart) <= 10 Then datStart = datStartTable
            If Not datFinishTable = "" Then If Abs(datFinishTable - datFinish) <= 10 Then datFinish = datFinishTable
    End Select
    
    'Comprobación de cambios y ejecución de la actualización
    If Not (datStartTable = datStart And datFinishTable = datFinish And booSameY) Then
        If xl_UpdChart Then
            If Not booBL And Not booSum And Not booSumBar And Not booPrg And Not booFlt Then
                If booResume Then
                    If Not rngResume.Offset(intRow - rngRef.row, 0) = datStart Then rngResume.Offset(intRow - rngRef.row, 0) = datStart
                Else
                    If Not rngStart.Offset(intRow - rngRef.row, 0) = datStart Then rngStart.Offset(intRow - rngRef.row, 0) = datStart
                End If
                If Not rngFinish.Offset(intRow - rngRef.row, 0) = datFinish Then rngFinish.Offset(intRow - rngRef.row, 0) = datFinish
            End If
        End If
        
        Dim arrRows As Variant
        arrRows = GetEditRows_wSum(Array(IIf(booSum, CInt(intRowSum), CInt(intRow))))
        If xl_UpdSch Then
            If rngActID.Offset(intRow - rngRef.row, 0) Like "WBS-*" Then
                ACT_CreateChart Array(IIf(booSum, CInt(intRowSum), CInt(intRow)))
            Else
                'Cálculo de duración remanente
                'Si es un hito (duracion total 0) o la fecha fin es anterior al cutoff --> Duración remanente = 0
                If (IsDate(datFinish) And datFinish <= datCutoff) Or (IsDate(datStart) And Not IsDate(datFinish)) Or (Not IsDate(datStart) And IsDate(datFinish)) Then
                    intRmgDur = 0
                'Si la fecha de inicio es mayor que el cutoff --> Duración remanente = Duración total
                ElseIf datStart > datCutoff Then
                    intRmgDur = DateDiffCal(datStart, datFinish) + IIf(WorkingDay(datStart), 1, 0)
                'Duración remanente = Fecha fin - Cutoff
                Else:
                    intRmgDur = DateDiffCal(datCutoff, datFinish)
                End If
                wsSch.Cells(intRow, rngRmgDur.Column) = intRmgDur
                
                CreateChart Array(IIf(booSum, CInt(intRowSum), CInt(intRow)))
                ACT_CalculateNetwork True
            End If
        Else
            shape.Delete
            ACT_CreateChart arrRows
        End If
    End If
subTermination:
    Application.EnableEvents = True
End Sub

