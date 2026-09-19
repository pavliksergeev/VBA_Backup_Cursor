Attribute VB_Name = "Module_PestorePipeNames"

' ================================================================
' module: Module_RestorePipeNames
' purpose: restore named ranges on the wsPipe sheet
' ================================================================
Option Explicit

' ================================================================
' main restore procedure
' ================================================================
Public Sub RestorePipeNamedRanges(Optional ByVal silent As Boolean = False)
    On Error GoTo CleanExit
    
    ' save settings
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
    Application.StatusBar = "restoring Pipe calculation named ranges..."
    
    Dim wsPipe As Worksheet
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    
    If wsPipe Is Nothing Then
        ' лист '
        ' лandст '
        ' ' не найден!
        ' ' not found!
        If Not silent Then
            MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbCritical
        End If
        GoTo CleanExit
    End If

    If Not NeedRestorePipeNames(wsPipe) Then
        Debug.Print Ru("0432 0441 0435 0020 0438 043C 0435 043D 0430 0020 0070 0069 0070 0065 0020 0432 0020 043F 043E 0440 044F 0434 043A 0435 002C 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432") & Ru("043B 0435 043D 0438 0435 0020 043D 0435 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F")
        Application.StatusBar = False
        GoTo CleanExit
    End If
    
    ' get the column count from D3 (pipeDifferentParametersNum)
    Dim colBr As Long
    Dim pipeCountValue As Variant
    
    On Error Resume Next
    pipeCountValue = wsPipe.Range("D3").Value
    If Err.Number <> 0 Then
        pipeCountValue = 1
        Err.Clear
    End If
    On Error GoTo 0
    
    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If
    
    If colBr < 1 Then colBr = 1
    If colBr > 1024 Then colBr = 1024
    
    Application.StatusBar = "restoring: " & colBr & " columns..."
    DoEvents
    
    Call ApplyPipeNames(wsPipe, colBr)
    
    Application.StatusBar = "restore finished: " & colBr & " columns"
    
    'MsgBox "Pipe calculation named ranges restored" & vbCrLf & _
    '       "columns: " & colBr, vbInformation, "restore finished"
    
CleanExit:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc
    Application.StatusBar = False
End Sub

' проверка динамических (1 x n) и скалярных (1x1) имён Pipe
' check dynamic (1 x n) and scalar (1x1) Pipe names
Private Function NeedRestorePipeNames(ByVal wsPipe As Worksheet) As Boolean
    On Error Resume Next

    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = wsPipe.Range("D3").Value
    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If

    Dim rangeNames As Variant
    rangeNames = Array( _
        "pipeSteelGrade", "pipeSteelResistivity", "pipeDiameter", _
        "pipeWallThickness", "pipeInsulationResistivityStartLife", "pipeLayingDepth", _
        "soilResistivityAvg", "serviceLifeDesigned", "pipeResistivityChangeFactor", _
        "pipeAlongResistance", "soilResistanceAroundPipe", "pipeTransientResistivity", _
        "soilResistivityAroundPipe", "pipeInsulationResistanceStartLife", _
        "pipeTransientResistance", "pipeTransientResistanceEndLife", _
        "factorPropagationCurrentAlongPipe", "factorPropagationCurrentAlongPipeEndLife", _
        "pipeImpedance", "pipeImpedanceEndLife" _
    )

    Dim nm As Variant
    Dim rng As Range
    For Each nm In rangeNames
        Set rng = Nothing
        Err.Clear
        Set rng = wsPipe.names(CStr(nm)).RefersToRange
        If Err.Number <> 0 Or rng Is Nothing Then
            NeedRestorePipeNames = True
            Exit Function
        End If
        If rng.rows.count <> 1 Or rng.Columns.count <> colBr Then
            NeedRestorePipeNames = True
            Exit Function
        End If
    Next nm

    Dim scalarNames As Variant
    scalarNames = Array("pipeDifferentParametersNum", "pipeInputResistance", "pipeInputResistanceEndLife")
    For Each nm In scalarNames
        Set rng = Nothing
        Err.Clear
        Set rng = wsPipe.names(CStr(nm)).RefersToRange
        If Err.Number <> 0 Or rng Is Nothing Then
            NeedRestorePipeNames = True
            Exit Function
        End If
        If rng.rows.count <> 1 Or rng.Columns.count <> 1 Then
            NeedRestorePipeNames = True
            Exit Function
        End If
    Next nm

    NeedRestorePipeNames = False
    On Error GoTo 0
End Function

' ================================================================
' update RefersTo in place (Add only if the name is missing)
' ================================================================
Private Sub ApplyPipeNames(ByVal wsPipe As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit
    
    Application.StatusBar = "updating Pipe names..."
    DoEvents
    
    Dim startCol As Long
    startCol = 4
    
    Dim endCol As Long
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, _
                     Application.WorksheetFunction.Min(1024, endCol))
    
    Call SetPipeName(wsPipe, "pipeDifferentParametersNum", wsPipe.Range("D3"))
    Call SetPipeName(wsPipe, "pipeInputResistance", wsPipe.Range("D28"))
    Call SetPipeName(wsPipe, "pipeInputResistanceEndLife", wsPipe.Range("D29"))
    
    Dim rangeNames As Variant
    Dim rangeRows As Variant
    
    rangeNames = Array( _
        "pipeSteelGrade", _
        "pipeSteelResistivity", _
        "pipeDiameter", _
        "pipeWallThickness", _
        "pipeInsulationResistivityStartLife", _
        "pipeLayingDepth", _
        "soilResistivityAvg", _
        "serviceLifeDesigned", _
        "pipeResistivityChangeFactor", _
        "pipeAlongResistance", _
        "soilResistanceAroundPipe", _
        "pipeTransientResistivity", _
        "soilResistivityAroundPipe", _
        "pipeInsulationResistanceStartLife", _
        "pipeTransientResistance", _
        "pipeTransientResistanceEndLife", _
        "factorPropagationCurrentAlongPipe", _
        "factorPropagationCurrentAlongPipeEndLife", _
        "pipeImpedance", _
        "pipeImpedanceEndLife" _
    )
    
    rangeRows = Array(5, 6, 7, 8, 9, 10, 11, 12, 13, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27)
    
    Dim nameIndex As Long
    Dim startRow As Long
    Dim newRng As Range
    
    For nameIndex = LBound(rangeNames) To UBound(rangeNames)
        startRow = rangeRows(nameIndex)
        Set newRng = wsPipe.Range(wsPipe.Cells(startRow, startCol), _
                                  wsPipe.Cells(startRow, endCol))
        If Not newRng Is Nothing Then
            Call SetPipeName(wsPipe, CStr(rangeNames(nameIndex)), newRng)
        End If
    Next nameIndex
    
CleanExit:
End Sub

Private Sub SetPipeName(ByVal wsPipe As Worksheet, ByVal rangeName As String, ByVal rng As Range)
    On Error Resume Next
    
    Dim nm As name
    Set nm = wsPipe.names(rangeName)
    If nm Is Nothing Then Set nm = thisWorkbook.names(rangeName)
    Err.Clear
    
    If nm Is Nothing Then
        wsPipe.names.Add name:=rangeName, RefersTo:=rng
        If Err.Number = 0 Then
            If DEBUG_MODE Then Debug.Print Ru("0441 043E 0437 0434 0430 043D 043E 0020 0438 043C 044F 003A 0020") & rangeName
        Else
            If DEBUG_MODE Then Debug.Print Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 0441 043E 0437 0434 0430 0442 044C 0020 0438 043C 044F 0020") & rangeName
            Err.Clear
        End If
    Else
        nm.RefersTo = rng
        If Err.Number = 0 Then
            If DEBUG_MODE Then Debug.Print Ru("043E 0431 043D 043E 0432 043B 0435 043D 043E 0020 0438 043C 044F 003A 0020") & rangeName
        Else
            If DEBUG_MODE Then Debug.Print Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043E 0431 043D 043E 0432 0438 0442 044C 0020 0438 043C 044F 0020") & rangeName
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
End Sub

' ================================================================
' button that runs restore
' ================================================================
Sub btnRestorePipeNames()
    Call RestorePipeNamedRanges
End Sub
