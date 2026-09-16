Attribute VB_Name = "Module_RestoreAnodNames"

' ================================================================
' module: Module_RestoreAnodNames
' purpose: restore named ranges on the "Anod" sheet
' scalar names are not deleted or recreated (except manual create).
' range names are updated without deletion (comments are kept).
' ================================================================
Option Explicit

Private Const SHEET_ANOD As String = "Anod"

' ================================================================
' main restore procedure (update without delete)
' ================================================================
Public Sub RestoreAnodNamedRanges()
    On Error GoTo CleanExit

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = "restoring Anod named ranges..."

    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets(SHEET_ANOD)

    If wsAnod Is Nothing Then
        ' лист '
        ' лandст '
        ' ' не найден!
        ' ' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_ANOD & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbCritical
        GoTo CleanExit
    End If

    ' --- check whether names need restoring ---
    If Not NeedRestoreAnodNames(wsAnod) Then
        ' все имена anod в порядке, восстановление не требуется
        ' все andмеon anod в порядке, восстановленandе не требуется
        Debug.Print Ru("0432 0441 0435 0020 0438 043C 0435 043D 0430 0020 0061 006E 006F 0064 0020 0432 0020 043F 043E 0440 044F 0434 043A 0435 002C 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432") & Ru("043B 0435 043D 0438 0435 0020 043D 0435 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F")
        Application.StatusBar = False
        GoTo CleanExit
    End If

    ' --- determine the column count ---
    Dim colBr As Long
    Dim pipeCountValue As Variant

    On Error Resume Next
    pipeCountValue = wsAnod.Range("D19").Value
    If Err.Number <> 0 Then
        pipeCountValue = 1
        Err.Clear
    End If
    On Error GoTo 0

    If IsNumeric(pipeCountValue) Then
        If pipeCountValue > 0 Then
            colBr = CLng(pipeCountValue)
        Else
            colBr = 1
        End If
    Else
        colBr = 1
    End If

    If colBr < 1 Then colBr = 1
    If colBr > 1024 Then colBr = 1024

    Application.StatusBar = "updating: " & colBr & " columns..."
    DoEvents

    ' --- update range names (without deletion) ---
    Call UpdateRangeOnlyAnodNames(wsAnod, colBr)

    Application.StatusBar = "restore finished: " & colBr & " columns"

    ' именованные диапазоны листа '
    ' named ranges лandста '
    ' ' восстановлены!
    ' ' restored!
    ' количество колонок: 
    ' column count: 
    ' всего диапазонных имен: 49
    ' total range names: 49
    ' восстановление завершено
    ' restore finished
    MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0435 0020 0434 0438 0430 043F 0430 0437 043E 043D 044B 0020 043B 0438 0441 0442 0430 0020 0027") & SHEET_ANOD & Ru("0027 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 044B 0021") & vbCrLf & _
           Ru("043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 043A 043E 043B 043E 043D 043E 043A 003A 0020") & colBr & vbCrLf & _
           Ru("0432 0441 0435 0433 043E 0020 0434 0438 0430 043F 0430 0437 043E 043D 043D 044B 0445 0020 0438 043C 0435 043D 003A 0020 0034 0039"), vbInformation, Ru("0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 0435 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E")

CleanExit:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc
    Application.StatusBar = False
End Sub

' ================================================================
' check whether names need restoring (all 49 range names)
' both rows and columns are checked now.
' ================================================================
Private Function NeedRestoreAnodNames(ByVal wsAnod As Worksheet) As Boolean
    On Error Resume Next

    ' get the expected column count from D19
    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = wsAnod.Range("D19").Value
    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If

    ' name lists by row count
    Const NAMES_1ROW As String = _
        "anodHeaders|avgProtectionCurrentCPOverLife|currentCP|currentEndLifeCP|depthToMidAG|" & _
        "diameterAG|dissolutionRateAG|drainWireCrossSection|factorCurrent|factorSoilHeterogeneity|" & _
        "factorUseMassAG|factorVoltageMarginCP|lengthElectrodeAG|lengthWireAGtoPipe|lengthWireCPtoPipe|" & _
        "lengthWorkPartDeepAG|massOneElectrodeAG|minDistancePipeToAG|nominalOutputCurrenCP|" & _
        "nominalOutputPowerCP|nominalOutputVoltageCP|pipeLengthCP|powerEndLifeCP|ratedCurrent|" & _
        "resistanceEndLifeAG|resistivity_i_layerDeepAG|resistivityMaterialAG|resistivitySoilAG|" & _
        "specificMaccOneMeterAG|specificRatedCurrent|typeAG|typeCP|typeDeliveryAG|typeInstallationAG|" & _
        "typeMaterial|typeMountingAG|voltageEndLifeCP|webLinkCP|wireResistivity|wiresResistanceCPpipeAG|" & _
        "cokeBreezeDiameterAG|cokeBreezelengthElectrodeAG|cokeBreezeResistivityAG"

    Const NAMES_10ROWS As String = _
        "correctResistanceAG|numElectrodesAG|oneElectrodeResistanceAG|serviceLifeAG|serviceLifeDeviation|weightWithoutFillingAG"

    ' check single-row names (43)
    Dim arr1 As Variant
    arr1 = Split(NAMES_1ROW, "|")
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        Dim nm1 As String
        nm1 = Trim(arr1(i))
        If Len(nm1) = 0 Then GoTo Next1
        
        Dim rng1 As Range
        Set rng1 = Nothing
        Set rng1 = wsAnod.names(nm1).RefersToRange
        If Err.Number <> 0 Or rng1 Is Nothing Then
            ' имя 
            ' name 
            '  не найдено, требуется восстановление
            '  not found, restore required
            Debug.Print Ru("0438 043C 044F 0020") & nm1 & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 002C 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D") & Ru("0438 0435")
            NeedRestoreAnodNames = True
            Exit Function
        End If
        If rng1.Columns.count <> colBr Or rng1.rows.count <> 1 Then
            ' имя 
            ' name 
            '  имеет размерность 
            '  has size 
            ' , ожидается 1x
            ' , expected 1x
            ' , требуется восстановление
            ' , restore required
            Debug.Print Ru("0438 043C 044F 0020") & nm1 & Ru("0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & rng1.rows.count & "x" & rng1.Columns.count & _
                        Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0031 0078") & colBr & Ru("002C 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 0435")
            NeedRestoreAnodNames = True
            Exit Function
        End If
        Err.Clear
Next1:
    Next i

    ' check ten-row names (6)
    Dim arr10 As Variant
    arr10 = Split(NAMES_10ROWS, "|")
    For i = LBound(arr10) To UBound(arr10)
        Dim nm10 As String
        nm10 = Trim(arr10(i))
        If Len(nm10) = 0 Then GoTo Next10
        
        Dim rng10 As Range
        Set rng10 = Nothing
        Set rng10 = wsAnod.names(nm10).RefersToRange
        If Err.Number <> 0 Or rng10 Is Nothing Then
            ' имя 
            ' name 
            '  не найдено, требуется восстановление
            '  not found, restore required
            Debug.Print Ru("0438 043C 044F 0020") & nm10 & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 002C 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D") & Ru("0438 0435")
            NeedRestoreAnodNames = True
            Exit Function
        End If
        If rng10.Columns.count <> colBr Or rng10.rows.count <> 10 Then
            ' имя 
            ' name 
            '  имеет размерность 
            '  has size 
            ' , ожидается 10x
            ' , expected 10x
            ' , требуется восстановление
            ' , restore required
            Debug.Print Ru("0438 043C 044F 0020") & nm10 & Ru("0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & rng10.rows.count & "x" & rng10.Columns.count & _
                        Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0031 0030 0078") & colBr & Ru("002C 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 0435")
            NeedRestoreAnodNames = True
            Exit Function
        End If
        Err.Clear
Next10:
    Next i

    NeedRestoreAnodNames = False
    On Error GoTo 0
End Function

' ================================================================
' delete non-scalar names only (kept for manual use)
' cp_List is no longer deleted. not used in the main procedure.
' ================================================================
Private Sub DeleteNonScalarAnodNames(ByVal wsAnod As Worksheet)
    On Error Resume Next
    Application.StatusBar = "removing old range names..."
    DoEvents

    Dim preservedNames As String
    preservedNames = Module_Constants.SCALAR_NAMES & "|CP_List"

    Dim scalarArray() As String
    scalarArray = Split(preservedNames, "|")

    Dim nm As name
    Dim deletedCount As Long
    deletedCount = 0

    For Each nm In wsAnod.names
        Dim pureName As String
        pureName = nm.name
        Dim pos As Long
        pos = InStr(pureName, "!")
        If pos > 0 Then
            pureName = Mid(pureName, pos + 1)
        End If

        Dim isPreserved As Boolean
        isPreserved = False
        Dim i As Long
        For i = LBound(scalarArray) To UBound(scalarArray)
            If StrComp(pureName, scalarArray(i), vbTextCompare) = 0 Then
                isPreserved = True
                Exit For
            End If
        Next i

        If Not isPreserved Then
            On Error Resume Next
            nm.Delete
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next nm

    ' удалено диапазонных имен с листа anod: 
    ' range names deleted from sheet anod: 
    Debug.Print Ru("0443 0434 0430 043B 0435 043D 043E 0020 0434 0438 0430 043F 0430 0437 043E 043D 043D 044B 0445 0020 0438 043C 0435 043D 0020 0441 0020 043B 0438 0441 0442 0430 0020 0041 006E") & Ru("006F 0064 003A 0020") & deletedCount
End Sub

' ================================================================
' update range names (without deletion)
' ================================================================
Private Sub UpdateRangeOnlyAnodNames(ByVal wsAnod As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit

    Application.StatusBar = "updating range names..."
    DoEvents

    Dim startCol As Long
    startCol = 4

    Dim endCol As Long
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, _
                     Application.WorksheetFunction.Min(1024, endCol))

    ' ================================================================
    ' 1. update/create range names (1 row x colBr)
    ' ================================================================
    Call SetRangeName1RowR1C1(wsAnod, "anodHeaders", 2, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "avgProtectionCurrentCPOverLife", 25, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "currentCP", 20, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "currentEndLifeCP", 21, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "depthToMidAG", 51, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "diameterAG", 39, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "dissolutionRateAG", 44, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "drainWireCrossSection", 11, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "factorCurrent", 9, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "factorSoilHeterogeneity", 54, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "factorUseMassAG", 55, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "factorVoltageMarginCP", 56, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "lengthElectrodeAG", 40, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "lengthWireAGtoPipe", 13, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "lengthWireCPtoPipe", 12, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "lengthWorkPartDeepAG", 59, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "massOneElectrodeAG", 43, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "minDistancePipeToAG", 50, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "nominalOutputCurrenCP", 29, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "nominalOutputPowerCP", 31, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "nominalOutputVoltageCP", 30, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "pipeLengthCP", 8, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "powerEndLifeCP", 24, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "ratedCurrent", 45, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "resistanceEndLifeAG", 58, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "resistivity_i_layerDeepAG", 53, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "resistivityMaterialAG", 46, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "resistivitySoilAG", 52, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "specificMaccOneMeterAG", 49, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "specificRatedCurrent", 48, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeAG", 38, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeCP", 26, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeDeliveryAG", 37, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeInstallationAG", 36, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeMaterial", 34, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "typeMountingAG", 35, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "voltageEndLifeCP", 22, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "webLinkCP", 27, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "wireResistivity", 10, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "wiresResistanceCPpipeAG", 23, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "cokeBreezeDiameterAG", 41, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "cokeBreezelengthElectrodeAG", 42, startCol, endCol)
    Call SetRangeName1RowR1C1(wsAnod, "cokeBreezeResistivityAG", 47, startCol, endCol)

    ' ================================================================
    ' 2. update/create range names (10 rows x colBr)
    ' ================================================================
    Call SetRangeName10RowsR1C1(wsAnod, "correctResistanceAG", 110, startCol, endCol)
    Call SetRangeName10RowsR1C1(wsAnod, "numElectrodesAG", 70, startCol, endCol)
    Call SetRangeName10RowsR1C1(wsAnod, "oneElectrodeResistanceAG", 60, startCol, endCol)
    Call SetRangeName10RowsR1C1(wsAnod, "serviceLifeAG", 90, startCol, endCol)
    Call SetRangeName10RowsR1C1(wsAnod, "serviceLifeDeviation", 100, startCol, endCol)
    Call SetRangeName10RowsR1C1(wsAnod, "weightWithoutFillingAG", 80, startCol, endCol)

    ' обновлены диапазонные имена на листе anod (
    ' range names updated on sheet anod (
    '  колонок, 49 шт.)
    '  columns, 49 pcs)
    Debug.Print Ru("043E 0431 043D 043E 0432 043B 0435 043D 044B 0020 0434 0438 0430 043F 0430 0437 043E 043D 043D 044B 0435 0020 0438 043C 0435 043D 0430 0020 043D 0430 0020 043B 0438 0441 0442") & Ru("0435 0020 0041 006E 006F 0064 0020 0028") & colBr & Ru("0020 043A 043E 043B 043E 043D 043E 043A 002C 0020 0034 0039 0020 0448 0442 002E 0029")

CleanExit:
End Sub

' ================================================================
' helper: update/create a range name (1 row)
' ================================================================
Private Sub SetRangeName1RowR1C1(ByVal ws As Worksheet, ByVal name As String, ByVal row As Long, ByVal startCol As Long, ByVal endCol As Long)
    On Error Resume Next

    Dim ref As String
    ref = "=" & ws.name & "!R" & row & "C" & startCol & ":R" & row & "C" & endCol

    Dim nm As name
    Set nm = ws.names(name)
    If nm Is Nothing Then
        ws.names.Add name:=name, RefersToR1C1:=ref
        '   создано: 
        '   created: 
        Debug.Print Ru("0020 0020 0441 043E 0437 0434 0430 043D 043E 003A 0020") & name & " -> " & ref
    Else
        nm.RefersToR1C1 = ref
        '   обновлено: 
        '   updated: 
        Debug.Print Ru("0020 0020 043E 0431 043D 043E 0432 043B 0435 043D 043E 003A 0020") & name & " -> " & ref
    End If

    If Err.Number <> 0 Then
        '   ошибка: 
        '   error: 
        Debug.Print Ru("0020 0020 043E 0448 0438 0431 043A 0430 003A 0020") & name & " (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ================================================================
' helper: update/create a range name (10 rows)
' ================================================================
Private Sub SetRangeName10RowsR1C1(ByVal ws As Worksheet, ByVal name As String, ByVal startRow As Long, ByVal startCol As Long, ByVal endCol As Long)
    On Error Resume Next

    Dim endRow As Long
    endRow = startRow + 9
    Dim ref As String
    ref = "=" & ws.name & "!R" & startRow & "C" & startCol & ":R" & endRow & "C" & endCol

    Dim nm As name
    Set nm = ws.names(name)
    If nm Is Nothing Then
        ws.names.Add name:=name, RefersToR1C1:=ref
        '   создано: 
        '   created: 
        Debug.Print Ru("0020 0020 0441 043E 0437 0434 0430 043D 043E 003A 0020") & name & " -> " & ref
    Else
        nm.RefersToR1C1 = ref
        '   обновлено: 
        '   updated: 
        Debug.Print Ru("0020 0020 043E 0431 043D 043E 0432 043B 0435 043D 043E 003A 0020") & name & " -> " & ref
    End If

    If Err.Number <> 0 Then
        '   ошибка: 
        '   error: 
        Debug.Print Ru("0020 0020 043E 0448 0438 0431 043A 0430 003A 0020") & name & " (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ================================================================
' restore scalar names only (if they were deleted)
' ================================================================
Public Sub CreateScalarNamesOnly()
    On Error Resume Next
    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets(SHEET_ANOD)
    If wsAnod Is Nothing Then Exit Sub

    Call SetScalarNameR1C1(wsAnod, "factorMutualInfluence", 6, 4)
    Call SetScalarNameR1C1(wsAnod, "lengthProtectiveZone", 18, 4)
    Call SetScalarNameR1C1(wsAnod, "maxProtectPotential", 4, 4)
    Call SetScalarNameR1C1(wsAnod, "minProtectPotential", 3, 4)
    Call SetScalarNameR1C1(wsAnod, "naturalPotential", 5, 4)
    Call SetScalarNameR1C1(wsAnod, "pipeCountCP", 19, 4)
    Call SetScalarNameR1C1(wsAnod, "pipeLength", 7, 4)
    Call SetScalarNameR1C1(wsAnod, "pipeShiftPotentialMin", 16, 4)
    Call SetScalarNameR1C1(wsAnod, "pipeShiftPotentialPoint", 17, 4)

    ' скалярные имена восстановлены
    ' scalar names restored
    Debug.Print Ru("0441 043A 0430 043B 044F 0440 043D 044B 0435 0020 0438 043C 0435 043D 0430 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 044B")
End Sub

Private Sub SetScalarNameR1C1(ByVal ws As Worksheet, ByVal name As String, ByVal row As Long, ByVal col As Long)
    On Error Resume Next

    Dim ref As String
    ref = "=" & ws.name & "!R" & row & "C" & col

    Dim nm As name
    Set nm = ws.names(name)
    If nm Is Nothing Then
        ws.names.Add name:=name, RefersToR1C1:=ref
    Else
        nm.RefersToR1C1 = ref
    End If

    If Err.Number = 0 Then
        ' скалярное имя: 
        ' scalar name: 
        Debug.Print Ru("0441 043A 0430 043B 044F 0440 043D 043E 0435 0020 0438 043C 044F 003A 0020") & name & " -> " & ref
    Else
        ' ошибка скалярного: 
        ' scalar error: 
        Debug.Print Ru("043E 0448 0438 0431 043A 0430 0020 0441 043A 0430 043B 044F 0440 043D 043E 0433 043E 003A 0020") & name & " (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ================================================================
' restore button (range names)
' ================================================================
Sub btnRestoreAnodNames()
    Call RestoreAnodNamedRanges
End Sub
