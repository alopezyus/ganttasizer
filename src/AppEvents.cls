VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myApp As Application
Attribute myApp.VB_VarHelpID = -1

Private varOldValue() As Variant
Private booOldValueExist As Boolean
Private intRows, intCols As Integer 'dimensiones del clipboard
Private intActLastRowOld, intVisibleRowsOld As Integer
Private intShapesCountOld As Integer
Private strActIDOld, strDescOld As String

Private Sub Class_Initialize()
    Set myApp = Application
    
    Call clipboardDims
    
    Set wsSch = ActiveSheet
    If (Application.ActiveProtectedViewWindow Is Nothing) Then
        Call dragndropCheck
        
        CreateShortcuts
    End If
End Sub

Private Sub Class_Terminate()
    Application.CellDragAndDrop = True
    DeleteShortcuts
End Sub

Private Sub myApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    initializeRibbonVar
End Sub

Private Sub myApp_WorkbookActivate(ByVal Wb As Workbook)
    Call clipboardDims
    
    Set wsSch = ActiveSheet
    Call dragndropCheck
End Sub

Private Sub myApp_SheetActivate(ByVal sh As Object)
    booPrjStartSet = False
    
    Call clipboardDims
    
    Set wsSch = sh
    Call dragndropCheck
End Sub

Private Sub myApp_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim varReturn As Variant
    Dim rng As Range
    Dim intArrPos, i, j As Integer
    Dim strPredSp() As String
    Dim strPred As String
    Dim datFinish, datStart, datCutoff As Date
    Dim intOffset As Integer
    
    Application.EnableCancelKey = xlDisabled
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsSch = sh
    'Comprobar si la hoja tiene encabezados
    varReturn = returnName
    If Not varReturn(0) Then GoTo subTermination
    'Comprobar si la hoja tiene todos los encabezados
    SetHeaderRef (False)
    If Not booHeaders Then GoTo UndoChanges
    'Calcular variables
    intActLastRow = ActLastRow
    datCutoff = SetCutoff
    arrDatesExc = GetCalendarExceptions
    
    'Mantenimiento de la referencia a la fila en el nombre de las formas después de insertar o eliminar filas
    If Not intActLastRowOld = intActLastRow And intActLastRowOld > 0 Then 'Si la última actividad de la lista no está en la misma fila que antes de la edición es que se han insertado o eliminado filas
        intOffset = intActLastRowOld - Target.Rows.Count - rngRef.row
        intOffset = IIf(intOffset < 0, 0, intOffset)
        'Tres posibilidadaes para lanzar el procedimiento de renombrado:
            '1: La fila insertada o eliminada está por encima del inicio de la lista de actividades y por tanto el procedimiento de detección de última fila funciona sin problemas
            '2: La última actividad está una fila por encima (una fila eliminada) o una o varias filas por debajo (fila insertada con información)
                'Si la última actividad está más de una fila por encima puede ser debido a varias actividades eliminadas o a varias líneas insertadas sin contenido
            '3: Si tanto el ActID como la Descripción de la última fila antes de la edición coinciden con los contenidos en la fila correspondiente a la última fila anterior menos
                'las filas seleccionadas en el Target de entrada, se han eliminado tantas filas como las seleccionadas en el Target de entrada
        If Target.row < rngRef.row Or _
            intActLastRow - intActLastRowOld >= -1 Or _
            (strActIDOld = rngActID.Offset(intOffset) And strDescOld = rngDesc.Offset(intOffset)) Then
            RenameShapes Target.row, intActLastRow - intActLastRowOld
            intActLastRowOld = intActLastRow
            'Si la inserción de actividades genera más de una fila en blanco, la última fila calculada es la justo anterior y las formas por debajo se eliminan
        Else: ClearChart Target.row
        End If
        'Si la inserción de una fila genera nuevas formas, se deben eliminar todas las formas que se hayan insertado ya que no podemos controlar el código de fila
        If intShapesCountOld < wsSch.Shapes.Count Then
            ClearShapeByPos Target.Top, Target.Top + Target.Height
            intShapesCountOld = wsSch.Shapes.Count
        End If
    End If

    'Si estamos trabajando fuera del rango de la lista actividades se sale del procedimiento
    If Target.row <= rngRef.row Or Target.row + Target.Rows.Count - 1 > intActLastRow + 1 Or Target.Column + Target.Columns.Count - 1 > rngPeriod.Column Then GoTo subTermination

    intArrPos = 0
    For Each rng In Target
        If IsError(rng.Value2) Then GoTo NextRange
        
        'COMPROBACION Y MANTENIMIENTO DE VALORES EN CAMPOS
        'Activity ID
        If rng.Column = rngActID.Column And rng.row > rngRef.row And rng.row <= intActLastRow Then
            'Forzar a que el Activity ID no contenga espacios
            If InStr(rng, " ") > 0 Then
                rng = Replace(rng, " ", "_")
            End If
    
            'Forzar a que el ActivityID sea único
            If Application.WorksheetFunction.CountIf(rngActID.Offset.Resize(intActLastRow - rngRef.row + 1), rng.Value2) > 1 Then
                i = 1
                Do While Application.WorksheetFunction.CountIf(rngActID.Offset(1).Resize(intActLastRow - rngRef.row + 1), rng.Value2 & "-" & i) > 0
                    i = i + 1
                Loop
                rng = "'" & rng.Value2 & "-" & i
            End If
        
            'Si se ha modificado el Activity ID y el indice no es mayor que el vector de valores antiguos
            If booOldValueExist And intArrPos <= UBound(varOldValue) Then
                If rng.Value2 <> varOldValue(intArrPos) Then
                    
                    'Forzar que Activity ID de WBS empiece por "WBS-" si se ha cambiado
                    If varOldValue(intArrPos) Like "WBS-*" Then
                        rng = "WBS-" & rngWBS.Offset(rng.row - rngRef.row)
                    End If
                    
                    'Actualizar Predecesoras
                    If Len(varOldValue(intArrPos)) > 0 Then
                        For i = rngRef.row + 1 To intActLastRow
                            If InStr(rngPred.Offset(i - rngRef.row), varOldValue(intArrPos)) Then
                                If Len(rng.Value2) = 0 Then
                                   strPredSp = Split(rngPred.Offset(i - rngRef.row), ",")
                                   strPred = ""
                                   For j = 0 To UBound(strPredSp)
                                        If InStr(strPredSp(j), varOldValue(intArrPos)) = 0 Then strPred = strPred & IIf(Len(strPred) = 0, "", ",") & strPredSp(j)
                                   Next
                                   rngPred.Offset(i - rngRef.row) = strPred
                                Else
                                    rngPred.Offset(i - rngRef.row) = Replace(rngPred.Offset(i - rngRef.row), (varOldValue(intArrPos)) & " ", rng.Value2 & " ")
                                End If
                            End If
                        Next
                    End If
                    
                End If
                'Actualizar OldValue
                varOldValue(intArrPos) = IIf(Len(rng.Value2) = 0, "", rng.Value2)
                intArrPos = intArrPos + 1
            End If
        End If
        
        'Forzar a que haya descripción cuando se haya definido una ventana
        If rngActStyle.Offset(rng.row - rngRef.row) = "WINDOW" And rngDesc.Offset(rng.row - rngRef.row) = "" Then
            rngDesc.Offset(rng.row - rngRef.row) = "Window"
        End If
        
        'Forzar a que el progreso tenga formato %
        If rng.Column = rngProgress.Column Then
            If IsNumeric(rng) And Not rng.NumberFormat Like "*%" Then
                rng.NumberFormat = "0.00%"
                rng = rng / 100
            End If
        End If
        
        'Forzar a que haya una ConstraintDate cuando se añade una constraint
        If rng.Column = rngSchMod.Column Then
            If Left(rng.Value2, 1) = "S" Then
                If Not IsDate(Cells(rng.row, rngConstraint.Column)) And Not Cells(rng.row, rngConstraint.Column) > 0 Then _
                    Cells(rng.row, rngConstraint.Column) = Cells(rng.row, rngStart.Column)
            ElseIf Left(rng.Value2, 1) = "F" Then
                If Not IsDate(Cells(rng.row, rngConstraint.Column)) And Not Cells(rng.row, rngConstraint.Column) > 0 Then _
                    Cells(rng.row, rngConstraint.Column) = Cells(rng.row, rngFinish.Column)
            End If
            rng.NumberFormat = "dd/mmm/yyyy"
        End If
        
NextRange:
    Next

    'Construccxión de vector con las filas editadas
    Dim arrRows, arrRowsUpd As Variant
    Dim row As Variant
    Dim booSchUpd, booUntUpd As Boolean
    
    i = 0
    j = 0
    ReDim arrRows(i)
    'Para cada celda en el rango editado
    For Each rng In Target
        'Si es error se salta al siguiente rango
'        If IsError(rng.Value2) Then GoTo NextRange2
        'Para cada fila en el vector de filas editadas
        For Each row In arrRows
            'Si la fila es igual a la fila de la celda de esta iteración la fila ya se ha añadido y se pasa a la siguiente fase
            If row = rng.row Then GoTo ColumnsCheck
        Next
        'Si se sale del bucle anterior añadimos la fila al vector
        If i > 0 Then ReDim Preserve arrRows(i)
        arrRows(i) = rng.row
        i = i + 1
        'Y comprobamos si no hay fechas y se eliminan las formas asociadas a la fila
        If (IsError(rngStart.Offset(rng.row - rngRef.row)) Or IsError(rngFinish.Offset(rng.row - rngRef.row))) Or _
            ((rngStart.Offset(rng.row - rngRef.row) = 0 Or Not IsDate(rngStart.Offset(rng.row - rngRef.row))) And _
            (rngFinish.Offset(rng.row - rngRef.row) = 0 Or Not IsDate(rngFinish.Offset(rng.row - rngRef.row)))) Then
            ClearShapeByPos rng.Top, rng.Top + rng.Height
        End If
ColumnsCheck:
        'Si se ha editado la columna de Remaining Duration, Predecessors o Fechas se activa el booleano para el cálculo de la red
        If rng.Column = rngRmgDur.Column Or rng.Column = rngPred.Column Or _
            rng.Column = rngStart.Column Or rng.Column = rngFinish.Column Or _
            rng.Column = rngStartAct.Column Or rng.Column = rngFinishAct.Column Or _
            rng.Column = rngConstraint.Column Or rng.Column = rngSchMod.Column Then
            booSchUpd = True
            'Si se ha editado la columna de Start o Finish se calcula la duración
            If xl_UpdSch And (rng.Column = rngStart.Column Or rng.Column = rngFinish.Column) Then
                Dim intRmgDur As Integer
                datFinish = wsSch.Cells(Target.row, rngFinish.Column)
                datStart = wsSch.Cells(Target.row, rngStart.Column)
                'Cálculo de duración remanente
                'Si es un hito (duracion total 0) o la fecha fin es anterior al cutoff --> Duración remanente = 0
                If (IsDate(datFinish) And datFinish <= datCutoff) Or (Not IsDate(datStart) And Not IsDate(datFinish)) Or _
                    (IsDate(datStart) And Not IsDate(datFinish)) Or (Not IsDate(datStart) And IsDate(datFinish)) Then
                    intRmgDur = 0
                'Si la fecha de inicio es mayor que el cutoff --> Duración remanente = Duración total
                ElseIf datStart > datCutoff Then
                    intRmgDur = DateDiffCal(datStart, datFinish) + IIf(WorkingDay(datStart), 1, 0)
                'Duración remanente = Fecha fin - Cutoff
                Else:
                    intRmgDur = DateDiffCal(datCutoff, datFinish)
                End If
                wsSch.Cells(Target.row, rngRmgDur.Column) = intRmgDur
            End If
        End If
        'Si se ha editado la columna de Remaining Unit se activa el booleano para la distribución de unidades
        If rng.Column = rngRmgUnt.Column Or rng.Column = rngDistCrv.Column Then booUntUpd = True
NextRange2:
    Next

    'Ejecución de actualización
    If xl_UpdSch And booSchUpd Then
        ContentsWBS arrRows
        ACT_CalculateNetwork True
    ElseIf Not UBound(arrRows) > 50 Then 'Si se están copiando más de 5 filas no se actualiza
        If Not IsEmpty(arrRows(0)) Then arrRowsUpd = GetEditRows_wSum(arrRows)
        If xl_UpdChart Then
            ACT_CreateChart arrRowsUpd
        ElseIf xl_UpdUnits And booUntUpd Then
            iniACT
            DistributeUnits arrRowsUpd
            finACT
        End If
    End If

    'Resetear OldValue
'    If TypeName(Selection) = "Range" Then
'        Selection.Offset(1).Select
'        Selection.Offset(-1).Select
'    End If

subTermination:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

UndoChanges:
    Application.Undo
    CustomMsgBox "Project headers cannot be deleted.", vbOKOnly + vbExclamation, error:=True
    GoTo subTermination
End Sub

'Añadir listas deplegables al seleccionar la celda correspondiente
Private Sub myApp_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    Dim DataObj As MSForms.DataObject
    Dim strClip As String
    Dim strFormula As String
    Dim rng, rngTarget As Range
    Dim intArrPos As Integer
    Dim varReturn As Variant
    Dim intRefRows, intRefCols As Integer
    Dim i, j As Integer
    
    Application.EnableCancelKey = xlDisabled
    
    'No deshabilitar eventos a la entrada de este evento.
    'Este evento redimensiona el rango seleccionado si hay un rango copiado, eso hace que vuelva a llamarse a sí mismo actualizando el vector de OldValues.
    'No hay ninguna edición de rangos ni de formas que pueda disparar otro evento.

    Call clipboardDims
        
    Set wsSch = sh
    'Comprobar si la hoja tiene encabezados
    varReturn = returnName
    If Not varReturn(0) Then
        If intRows + intCols = 0 And Application.CellDragAndDrop = False Then Application.CellDragAndDrop = True
        Exit Sub
    End If
    
    SetHeaderRef (False)
    If Not booHeaders Then Exit Sub
    intActLastRow = ActLastRow
    intActLastRowOld = intActLastRow
    intShapesCountOld = wsSch.Shapes.Count
    strActIDOld = rngActID.Offset(intActLastRow - rngRef.row)
    strDescOld = rngDesc.Offset(intActLastRow - rngRef.row)
    
'    'Comprobar si hay filas filtradas y en tal caso ocultar las formas de las filas no mostradas
'    If Not intVisibleRowsOld = 0 Then
'        Application.EnableEvents = False
'        If intVisibleRowsOld <> Range(rngActID, rngActID.Offset(intActLastRow - rngRef.row)).SpecialCells(xlCellTypeVisible).Count Then
'            FilterShapes
'        End If
'    End If
'    intVisibleRowsOld = Range(rngActID, rngActID.Offset(intActLastRow - rngRef.row)).SpecialCells(xlCellTypeVisible).Rows.Count
'    Application.EnableEvents = True
    
    'Si hemos seleccionado una fila por encima de la lista de actividades y sumando el número de filas copiadas entramos en la lista se impide el copiado
    If Target.row <= rngRef.row And Target.row + intRows - 1 > rngRef.row And Target.row + intRows - 1 <= intActLastRow Then Application.CutCopyMode = False
    
    'Si estamos trabajando fuera del rango de la lista actividades se sale del procedimiento
    If Target.row <= rngRef.row Or Target.row + Target.Rows.Count - 1 > intActLastRow + 1 Or Target.Column + Target.Columns.Count - 1 > rngPeriod.Column Then
        If Not rngActID.Column = Target.Column And Application.CellDragAndDrop = False And intRows = 0 Then Application.CellDragAndDrop = True
        Exit Sub
    End If
    
    'Guardar valores del rango seleccionado si son un ActivityID
    ReDim varOldValue(0)
    Set rngTarget = Cells(Target.row, Target.Column)
    intRefRows = IIf(intRows > Target.Rows.Count, intRows, Target.Rows.Count) - 1
    intRefCols = IIf(intCols > Target.Columns.Count, intCols, Target.Columns.Count) - 1
    For i = 0 To intRefRows
        For j = 0 To intRefCols
            Set rng = rngTarget.Offset(i, j)
            If rng.Column = rngActID.Column And rng.row > rngRef.row And rng.row <= intActLastRow Then
                If IsEmpty(varOldValue(0)) Then
                    varOldValue(0) = rng.Value2
                Else:
                    ReDim Preserve varOldValue(UBound(varOldValue) + 1)
                    varOldValue(UBound(varOldValue)) = rng.Value2
                End If
            End If
        Next
    Next
    booOldValueExist = True
    
    'Evitar que en la columna Activity ID se pueda hacer Drag&Drop
    If rngActID.Column = Target.Column And Application.CellDragAndDrop = True And intRows = 0 Then
        Application.CellDragAndDrop = False
    ElseIf Not rngActID.Column = Target.Column And Application.CellDragAndDrop = False And intRows = 0 Then
        Application.CellDragAndDrop = True
    End If
    
    'Configurar validación en columnas de setup
    ActiveSheet.Cells.Validation.Delete
    
    'Salir del procedimiento si se ha seleccionado más de una hoja
    If Application.ActiveWindow.SelectedSheets.Count > 1 Then Exit Sub
        
    Select Case Target.Column
    Case rngActStyle.Column
        'Limited in Free Edition
        If intEdition = 1 Then
            strFormula = "NO, 1 - 3D, 2 - Solid, 3 - Diagonal, 4 - Vertical, " & _
                        "11 - Diamond, 12 - Down Arrow, 13 - Up Arrow"
        
        Else
            strFormula = "NO, 1 - 3D, 2 - Solid, 3 - Diagonal, 4 - Vertical, 5 - Horizontal, 6 - Checkerboard, 7 - Confetti, 8 - Summary, 9 - Left Arrow, 10 - Right Arrow, " & _
                        "11 - Diamond, 12 - Down Arrow, 13 - Up Arrow, 14 - 5 Points Star, 15 - 4 Points Star, 16 - Heart, 17 - Oval, WINDOW"
        End If
    Case rngShpHgt.Column
        strFormula = "10%, 20%, 30%, 40%, 50%, 60%, 70%, 80%, 90%, 100%"
    Case rngConStyle.Column
        'Blank in Free Edition
        If intEdition = 1 Then
            strFormula = "-"
        Else
            strFormula = "NO, 1 - Solid Arrow, 2 - Dash Arrow, 3 - Dot Arrow, 4 - Solid Line, 5 - Dash Line, 6 - Dot Line"
        End If
    Case rngLabPos.Column
        strFormula = "NO, 0L, 0M, 0R, 1L, 1M, 1R, 2L, 2M, 2R"
    Case rngTmlMod.Column
        strFormula = "SUM, MIL, ACT"
    Case rngSchMod.Column
        'Blank in Free Edition and Pro Edition
        If intEdition > 0 Then
            strFormula = "-"
        Else
            strFormula = "NO, MANUAL, ALAP - As Late as Possible, " & _
                        "SON - Start On, SOB - Start On or Before, SOL - Start On or Later, SMN - Mandatory Start, " & _
                        "FON - Finish On, FOB - Finish On or Before, FOL - Finish On or Later, FMN - Mandatory Finish"
        End If
    Case rngDistCrv.Column
        'Blank in Free Edition and Pro Edition
                If intEdition > 0 Then
            strFormula = "-"
        Else
            strFormula = "1 - Linear, 2 - S-Curve, 3 - Front Loaded, 4 - Back Loaded, 5 - Step 0-100"
        End If
    Case rngStart.Column, rngFinish.Column, rngStartBL.Column, rngFinishBL.Column, rngStartAct.Column, rngFinishAct.Column, rngResume.Column
        createGridDP
        Exit Sub
    Case Else
        Exit Sub
    End Select
    
    If Target.Column = rngShpHgt.Column And rngActStyle.Offset(Target.row - rngRef.row) = "WINDOW" Then
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1", Formula2:="10000"
            .InCellDropdown = False
            .ShowInput = False
            .ShowError = True
        End With
    Else
        With Target.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=strFormula
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
        End With
    End If

End Sub

Private Sub clipboardDims()
    Dim DataObj As MSForms.DataObject
    Dim strClip As String
    'Obtener dimensiones del rango guardado (copiado) en el clipboard
    'https://stackoverflow.com/questions/9022245/get-text-from-clipboard-using-gettext-avoid-error-on-empty-clipboard
    Set DataObj = New MSForms.DataObject
    intRows = 0
    intCols = 0
    On Error GoTo exitSub

    DataObj.GetFromClipboard
    strClip = DataObj.GetText(1)

    intRows = Len(strClip) - Len(Replace(strClip, vbCr, ""))
    intCols = (Len(strClip) - Len(Replace(strClip, vbTab, ""))) / intRows + 1
exitSub:
End Sub

Private Sub dragndropCheck()
    Dim varReturn As Variant
    Call clipboardDims
    'Comprobar si la hoja tiene encabezados
    varReturn = returnName
    
    If intRows + intCols = 0 Then
        If Not varReturn(0) Then
            If Application.CellDragAndDrop = False Then Application.CellDragAndDrop = True
        ElseIf ActiveCell.Column <> Range("VB_" & varReturn(1) & "_08").Column Then
            If Application.CellDragAndDrop = False Then Application.CellDragAndDrop = True
        ElseIf Application.CellDragAndDrop = True Then Application.CellDragAndDrop = False
        End If
    End If
End Sub
