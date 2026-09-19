Attribute VB_Name = "Module_RunAllTests"

Option Explicit

Private testRunning As Boolean

' ================================================================
' Write a diagnostic message to a log file next to the workbook.
' Only writes when DEBUG_MODE is True.
' ================================================================
Public Sub LogTrace(ByVal msg As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    Dim logPath As String
    If Len(thisWorkbook.Path) = 0 Then
        logPath = Environ("TEMP") & Application.PathSeparator & "test_trace.log"
    Else
        logPath = thisWorkbook.Path & Application.PathSeparator & "test_trace.log"
    End If
    
    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Append As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #fnum
    
    On Error GoTo 0
End Sub

' ================================================================
' Clear the log file at the start of the test
' ================================================================
Private Sub InitLog()
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    
    Dim logPath As String
    If Len(thisWorkbook.Path) = 0 Then
        logPath = Environ("TEMP") & Application.PathSeparator & "test_trace.log"
    Else
        logPath = thisWorkbook.Path & Application.PathSeparator & "test_trace.log"
    End If
    
    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Output As #fnum
    Print #fnum, "=== test log started at " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Close #fnum
    
    On Error GoTo 0
End Sub

' Dump TestReport A:F to the Immediate window when DEBUG_MODE is True.
Private Sub DumpTestReportToImmediate(ByVal wsReport As Worksheet)
    If Not DEBUG_MODE Then Exit Sub
    If wsReport Is Nothing Then Exit Sub
    On Error Resume Next

    Dim lastRow As Long
    Dim lastF As Long
    lastRow = wsReport.Cells(wsReport.rows.count, 1).End(xlUp).row
    lastF = wsReport.Cells(wsReport.rows.count, 6).End(xlUp).row
    If lastF > lastRow Then lastRow = lastF
    If lastRow < 1 Then Exit Sub

    Debug.Print "=== TestReport ==="
    Dim r As Long
    Dim c As Long
    Dim line As String
    Dim v As Variant
    Dim hasVal As Boolean
    For r = 1 To lastRow
        line = ""
        hasVal = False
        For c = 1 To 6
            If c > 1 Then line = line & " | "
            v = wsReport.Cells(r, c).Value
            If IsError(v) Then
                line = line & "#ERR"
                hasVal = True
            ElseIf Not IsEmpty(v) And CStr(v) <> "" Then
                line = line & CStr(v)
                hasVal = True
            End If
        Next c
        If hasVal Then Debug.Print line
    Next r
    Debug.Print "=== TestReport end ==="
    On Error GoTo 0
End Sub

' ================================================================
' Clear the values in the listed named ranges on the Anod sheet.
' Waits (DoEvents) after each clear so that dependent formulas and
' validations have time to react before the test starts.
' ================================================================
Private Sub ClearNamedRangesForTest()
    Dim namesToClear As Variant
    Const NAMES_TO_CLEAR As String = _
        "typeMaterial|typeMountingAG|typeInstallationAG|typeDeliveryAG|typeAG|" & _
        "diameterAG|lengthElectrodeAG|cokeBreezeDiameterAG|cokeBreezelengthElectrodeAG|" & _
        "massOneElectrodeAG|dissolutionRateAG|ratedCurrent|resistivityMaterialAG|" & _
        "cokeBreezeResistivityAG|specificRatedCurrent|specificMaccOneMeterAG|" & _
        "resistanceEndLifeAG|lengthWorkPartDeepAG|oneElectrodeResistanceAG|" & _
        "oneElectrodeResistanceHorizAG|numElectrodesAG|weightWithoutFillingAG|" & _
        "serviceLifeAG|serviceLifeDeviation|correctResistanceAG"
    namesToClear = Split(NAMES_TO_CLEAR, "|")
    
    Call LogTrace("CLEAR: starting cleanup of named ranges on Anod")
    
    Dim nm As Variant
    For Each nm In namesToClear
        Call LogTrace("CLEAR: clearing '" & CStr(nm) & "'")
        On Error Resume Next
        Dim rng As Range
        Set rng = Nothing
        Set rng = wsAnod.Range(CStr(nm))
        If Err.Number <> 0 Then
            Call LogTrace("CLEAR: name '" & CStr(nm) & "' not found - skipped")
            Err.Clear
        ElseIf Not rng Is Nothing Then
            rng.ClearContents
            If Err.Number <> 0 Then
                Call LogTrace("CLEAR: failed to clear '" & CStr(nm) & "': " & Err.Description)
                Err.Clear
            End If
        End If
        On Error GoTo 0
        
    Next nm
    
    ' Final wait to make sure all dependencies are recalculated
    DoEvents
    Call LogTrace("CLEAR: cleanup of named ranges finished")
End Sub

' ================================================================
' Main test entry point
' ================================================================
Public Sub RunAllTests()
    ' --- guard against re-entry ---
    If testRunning Then
        ' тест уже выполняется. дождитесь завершения.
        ' test is already running. wait until it finishes.
        MsgBox Ru("0442 0435 0441 0442 0020 0443 0436 0435 0020 0432 044B 043F 043E 043B 043D 044F 0435 0442 0441 044F 002E 0020 0434 043E 0436 0434 0438 0442 0435 0441 044C 0020 0437 0430 0432") & Ru("0435 0440 0448 0435 043D 0438 044F 002E"), vbExclamation
        Exit Sub
    End If
    testRunning = True
    
    ' --- init log file (clears previous content) ---
    Call InitLog
    Call LogTrace("TEST: started")
    
    ' --- remember initial application state ---
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCursor As XlMousePointer
    Dim oldEnableCancelKey As XlEnableCancelKey
    Dim dumpedReport As Boolean
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCursor = Application.Cursor
    oldEnableCancelKey = Application.EnableCancelKey

    On Error GoTo CleanExit
    
    ' --- disable extras for the duration of the test ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' --- disable maxProtectPotential recalculation during the test ---
    wsAnod.isProgrammaticPipeCountChange = True
    Call LogTrace("TEST: programmatic mode enabled")

    ' --- timeout (10 minutes = 600 seconds) ---
    Dim startTime As Double
    startTime = Timer
    Const TIMEOUT_SECONDS As Long = 600

    ' ================================================================
    ' CLEANUP OF NAMED RANGES BEFORE THE TEST
    ' ================================================================
    Call LogTrace("TEST: clearing named ranges on Anod before start")
    Call ClearNamedRangesForTest
    Call LogTrace("TEST: named ranges cleaned, starting test")

    ' ================================================================
    ' CREATE TEST REPORT SHEET
    ' ================================================================
    Dim wsReport As Worksheet
    On Error Resume Next
    Set wsReport = thisWorkbook.Worksheets("TestReport")
    If wsReport Is Nothing Then
        Set wsReport = thisWorkbook.Worksheets.Add(After:=thisWorkbook.Worksheets(thisWorkbook.Worksheets.count))
        wsReport.name = "TestReport"
    End If
    On Error GoTo CleanExit
    
    ' Full reset of contents and formats without deleting the sheet
    wsReport.Cells.Clear
    wsReport.Cells.Interior.ColorIndex = xlNone
    wsReport.Cells.Font.ColorIndex = xlAutomatic
    wsReport.Cells.Font.Bold = False
    wsReport.rows.Hidden = False
    wsReport.Columns.Hidden = False
    
    ' отчёт о тестировании эхз (
    ' cp design test report (
    wsReport.Range("A1").Value = Ru("043E 0442 0447 0451 0442 0020 043E 0020 0442 0435 0441 0442 0438 0440 043E 0432 0430 043D 0438 0438 0020 044D 0445 0437 0020 0028") & TABLE_ALL_TEST & ")"
    ' дата: 
    ' date: 
    wsReport.Range("A2").Value = Ru("0434 0430 0442 0430 003A 0020") & Now
    ' сценарий
    ' scenario
    wsReport.Range("A4").Value = Ru("0441 0446 0435 043D 0430 0440 0438 0439")
    ' параметр
    ' parameter
    wsReport.Range("B4").Value = Ru("043F 0430 0440 0430 043C 0435 0442 0440")
    ' ожидаемое
    ' expected
    wsReport.Range("C4").Value = Ru("043E 0436 0438 0434 0430 0435 043C 043E 0435")
    ' фактическое
    ' actual
    wsReport.Range("D4").Value = Ru("0444 0430 043A 0442 0438 0447 0435 0441 043A 043E 0435")
    ' отклонение, %
    ' deviation, %
    wsReport.Range("E4").Value = Ru("043E 0442 043A 043B 043E 043D 0435 043D 0438 0435 002C 0020 0025")
    ' результат
    ' result
    wsReport.Range("F4").Value = Ru("0440 0435 0437 0443 043B 044C 0442 0430 0442")

    Dim row As Long
    row = 5

    ' ================================================================
    ' CHECK LOCAL NAMES ON SHEETS
    ' ================================================================
    Call LogTrace("CHECK: verifying local names")
    On Error Resume Next
    Dim testRng As Range
    Set testRng = wsAnod.Range("A1")
    If Err.Number <> 0 Then
        ' глобальный объект wsanod не найден! проверьте кодовое имя листа anod.
        ' global object wsanod not found! check the sheet code name anod.
        MsgBox Ru("0433 043B 043E 0431 0430 043B 044C 043D 044B 0439 0020 043E 0431 044A 0435 043A 0442 0020 0077 0073 0041 006E 006F 0064 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021") & Ru("0020 043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 043A 043E 0434 043E 0432 043E 0435 0020 0438 043C 044F 0020 043B 0438 0441 0442 0430 0020 0041 006E 006F 0064 002E"), vbCritical
        GoTo CleanExit
    End If
    Err.Clear
    Set testRng = wsPipe.Range("A1")
    If Err.Number <> 0 Then
        ' глобальный объект wspipe не найден! проверьте кодовое имя листа pipe.
        ' global object wspipe not found! check the sheet code name pipe.
        MsgBox Ru("0433 043B 043E 0431 0430 043B 044C 043D 044B 0439 0020 043E 0431 044A 0435 043A 0442 0020 0077 0073 0050 0069 0070 0065 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021") & Ru("0020 043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 043A 043E 0434 043E 0432 043E 0435 0020 0438 043C 044F 0020 043B 0438 0441 0442 0430 0020 0050 0069 0070 0065 002E"), vbCritical
        GoTo CleanExit
    End If
    On Error GoTo 0

    Dim pipeLocalNames As Variant
    pipeLocalNames = Array("pipeDifferentParametersNum", "pipeInputResistance", "pipeInputResistanceEndLife", _
                           "pipeImpedance", "pipeImpedanceEndLife")
    Dim anodLocalNames As Variant
    anodLocalNames = Array("pipeCountCP", "typeMaterial", "typeMountingAG", "typeInstallationAG", "typeAG")

    Dim missingPipe As String, missingAnod As String
    Dim nm As Variant
    For Each nm In pipeLocalNames
        On Error Resume Next
        Dim rng As Range
        Set rng = wsPipe.names(nm).RefersToRange
        If Err.Number <> 0 Then
            If missingPipe <> "" Then missingPipe = missingPipe & ", "
            missingPipe = missingPipe & nm
            Err.Clear
        End If
        On Error GoTo 0
    Next nm

    For Each nm In anodLocalNames
        On Error Resume Next
        Set rng = wsAnod.names(nm).RefersToRange
        If Err.Number <> 0 Then
            If missingAnod <> "" Then missingAnod = missingAnod & ", "
            missingAnod = missingAnod & nm
            Err.Clear
        End If
        On Error GoTo 0
    Next nm

    Dim missingAll As String
    If missingPipe <> "" Then missingAll = "pipe: " & missingPipe
    If missingAnod <> "" Then
        If missingAll <> "" Then missingAll = missingAll & "; "
        missingAll = missingAll & "anod: " & missingAnod
    End If

    If missingAll <> "" Then
        ' ошибка
        ' error
        wsReport.Cells(row, 1).Value = Ru("043E 0448 0438 0431 043A 0430")
        ' отсутствуют локальные имена
        ' missing local names
        wsReport.Cells(row, 2).Value = Ru("043E 0442 0441 0443 0442 0441 0442 0432 0443 044E 0442 0020 043B 043E 043A 0430 043B 044C 043D 044B 0435 0020 0438 043C 0435 043D 0430")
        wsReport.Cells(row, 3).Value = missingAll
        ' не пройден
        ' failed
        wsReport.Cells(row, 6).Value = Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D")
        row = row + 1
        ' отсутствуют локальные имена: 
        ' missing local names: 
        ' . восстановите их и запустите тест снова.
        ' . восстановandте andх and run тест снова.
        MsgBox Ru("043E 0442 0441 0443 0442 0441 0442 0432 0443 044E 0442 0020 043B 043E 043A 0430 043B 044C 043D 044B 0435 0020 0438 043C 0435 043D 0430 003A 0020") & missingAll & Ru("002E 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 0438 0442 0435 0020 0438 0445 0020 0438 0020 0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0442 0435 0441 0442 0020") & Ru("0441 043D 043E 0432 0430 002E"), vbCritical
        GoTo CleanExit
    Else
        ' успешно
        ' ok
        wsReport.Cells(row, 1).Value = Ru("0443 0441 043F 0435 0448 043D 043E")
        ' все локальные имена найдены
        ' все local names found
        wsReport.Cells(row, 2).Value = Ru("0432 0441 0435 0020 043B 043E 043A 0430 043B 044C 043D 044B 0435 0020 0438 043C 0435 043D 0430 0020 043D 0430 0439 0434 0435 043D 044B")
        ' пройден
        ' passed
        wsReport.Cells(row, 6).Value = Ru("043F 0440 043E 0439 0434 0435 043D")
        row = row + 1
    End If
    Call LogTrace("CHECK: local names verified")

    ' ================================================================
    ' проверка размерности расчётных имён: 1 x pipeCountCP на постоянных строках
    ' check result name dimensions: 1 x pipeCountCP on constant rows
    ' ================================================================
    Call LogTrace("CHECK: verifying named range dimensions")
    Dim dimCheckNames As Variant
    Dim dimCheckRows As Variant
    Call AnodResultNameSpecs(dimCheckNames, dimCheckRows)
    Dim dimIdx As Long
    Dim dimName As Variant
    Dim dimRng As Range
    Dim expRows As Long
    Dim expRow As Long
    expRows = 1
    Dim actRows As Long, actCols As Long
    Dim pipeCountCheck As Long
    pipeCountCheck = CLng(wsAnod.Range("pipeCountCP").Value)
    For dimIdx = LBound(dimCheckNames) To UBound(dimCheckNames)
        dimName = dimCheckNames(dimIdx)
        expRow = CLng(dimCheckRows(dimIdx))
        On Error Resume Next
        Set dimRng = wsAnod.names(dimName).RefersToRange
        If Err.Number <> 0 Then
            ' именованный диапазон '
            ' named range '
            ' ' не существует!
            ' ' does not exist!
            MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0439 0020 0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & dimName & Ru("0027 0020 043D 0435 0020 0441 0443 0449 0435 0441 0442 0432 0443 0435 0442 0021"), vbCritical
            GoTo CleanExit
        End If
        On Error GoTo 0
        actRows = dimRng.rows.count
        actCols = dimRng.Columns.count
        If actRows <> expRows Or actCols <> pipeCountCheck Or dimRng.row <> expRow Then
            ' именованный диапазон '
            ' named range '
            ' ' имеет размерность 
            ' ' has size 
            '  строка 
            '  row 
            ' , ожидается 
            ' , expected 
            '  строка 
            '  row 
            ' . тест прерван.
            ' . test aborted.
            MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0439 0020 0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & dimName & Ru("0027 0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & actRows & "x" & actCols & _
                   Ru("0020 0441 0442 0440 043E 043A 0430 0020") & dimRng.row & _
                   Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & expRows & "x" & pipeCountCheck & _
                   Ru("0020 0441 0442 0440 043E 043A 0430 0020") & expRow & Ru("002E 0020 0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 002E"), vbCritical
            GoTo CleanExit
        End If
    Next dimIdx
    Call LogTrace("CHECK: all named ranges have correct dimensions (" & expRows & "x" & pipeCountCheck & " on constant rows)")

    ' ================================================================
    ' PREPARE DATA ON PIPE SHEET
    ' ================================================================
    Call LogTrace("PIPE: preparing data")
    DoEvents
    wsPipe.Range("pipeDifferentParametersNum").Value = 2
    DoEvents

    ' начальные значения входных параметров трубы через именованные диапазоны
    ' initial pipe input values via named ranges
    Call FillPipeNamedRange("pipeSteelGrade", _
        Ru("0414 0430 043D 043D 044B 0435 0020 043E 0020 043C 0430 0440 043A 0435 0020 0441 0442 0430 043B 0438 0020 043E 0442 0441") & _
        Ru("0443 0442 0441 0442 0432 0443 044E 0442"))
    Call FillPipeNamedRange("pipeSteelResistivity", 0.000000245)
    Call FillPipeNamedRange("pipeDiameter", 1.22)
    Call FillPipeNamedRange("pipeWallThickness", 0.017)
    Call FillPipeNamedRange("pipeInsulationResistivityStartLife", 50000)
    Call FillPipeNamedRange("pipeLayingDepth", 1.6)
    Call FillPipeNamedRange("soilResistivityAvg", 50)
    Call FillPipeNamedRange("serviceLifeDesigned", 30)
    Call FillPipeNamedRange("pipeResistivityChangeFactor", 0.11)

    ' синхронизация колонок и валидация списков (сталь, диаметр, kиз, Rиз0)
    ' column sync and list validation (steel, diameter, kins, Rins0)
    Application.StatusBar = "pipe column sync and validation..."
    DoEvents
    Call LogTrace("PIPE: ForceSyncFromD3 (columns + validation lists)")
    On Error Resume Next
    Call wsPipe.ForceSyncFromD3
    If Err.Number <> 0 Then
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 043F 0440 0438 0020 0441 0438 043D 0445 0440 043E 043D 0438 0437 0430 0446 0438 0438 0020 043A 043E 043B 043E 043D 043E 043A 0020 043B 0438 0441") & _
               Ru("0442 0430 0020 0050 0069 0070 0065 003A 0020") & Err.Description, vbCritical
        Call LogTrace("PIPE: sync error " & Err.Description)
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo 0
    ' sync включает события и автопересчёт — вернуть режим теста
    ' sync turns events and auto-calc on — restore test mode
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Call LogTrace("PIPE: columns synced, validation lists created")

    ' --- pipe calculation ---
    Application.StatusBar = "pipe calculation..."
    DoEvents
    Call LogTrace("PIPE: running btnPipeCalculate")
    On Error Resume Next
    Call Module_calcPipe.btnPipeCalculate
    If Err.Number <> 0 Then
        ' ошибка при вызове btnpipecalculate: 
        ' error calling btnpipecalculate: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 043F 0440 0438 0020 0432 044B 0437 043E 0432 0435 0020 0062 0074 006E 0050 0069 0070 0065 0043 0061 006C 0063 0075 006C 0061 0074 0065 003A") & Ru("0020") & Err.Description, vbCritical
        Call LogTrace("PIPE: error " & Err.Description)
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo 0
    Call LogTrace("PIPE: calculation finished")

    Dim R_in As Double, R_in_end As Double
    Dim val1 As Variant, val2 As Variant
    val1 = wsPipe.Range("pipeInputResistance").Value
    val2 = wsPipe.Range("pipeInputResistanceEndLife").Value

    If IsError(val1) Or IsError(val2) Or Not IsNumeric(val1) Or Not IsNumeric(val2) Or val1 = 0 Or val2 = 0 Then
        ' ошибка: pipeinputresistance или pipeinputresistanceendlife содержат ошибку, не число или равны 0!
        ' error: pipeinputresistance or pipeinputresistanceendlife contain an error, are not a number, or equal 0!
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0070 0069 0070 0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0020 0438 043B 0438 0020 0070 0069 0070") & Ru("0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0045 006E 0064 004C 0069 0066 0065 0020 0441 043E 0434 0435 0440 0436 0430 0442 0020 043E 0448") & Ru("0438 0431 043A 0443 002C 0020 043D 0435 0020 0447 0438 0441 043B 043E 0020 0438 043B 0438 0020 0440 0430 0432 043D 044B 0020 0030 0021"), vbCritical
        GoTo CleanExit
    End If
    R_in = CDbl(val1)
    R_in_end = CDbl(val2)

    Call LogTrace("PIPE: R_in = " & R_in & ", R_in_end = " & R_in_end)

    Dim expR_in_end As Double
    expR_in_end = 0.0242
    Dim devR As Double
    If expR_in_end <> 0 Then
        devR = Abs((R_in_end - expR_in_end) / expR_in_end) * 100
    Else
        devR = 0
    End If

    ' трубопровод
    ' pipeline
    wsReport.Cells(row, 1).Value = Ru("0442 0440 0443 0431 043E 043F 0440 043E 0432 043E 0434")
    ' rвх(t)
    ' rin(t)
    wsReport.Cells(row, 2).Value = Ru("0072 0432 0445 0028 0074 0029")
    wsReport.Cells(row, 3).Value = expR_in_end
    wsReport.Cells(row, 4).Value = R_in_end
    wsReport.Cells(row, 5).Value = devR
    ' пройден
    ' passed
    ' не пройден
    ' failed
    wsReport.Cells(row, 6).Value = IIf(devR <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
    row = row + 1

    ' ================================================================
    ' PREPARE ANOD SHEET (GENERAL PARAMETERS)
    ' ================================================================
    Call LogTrace("ANOD: preparing general parameters")
    wsAnod.Range("minProtectPotential").Value = -0.85
    wsAnod.Range("maxProtectPotential").Value = -1.15
    wsAnod.Range("naturalPotential").Value = -0.55
    wsAnod.Range("factorMutualInfluence").Value = 0.5
    wsAnod.Range("pipeLength").Value = 350000

    ' --- protective zone calculation ---
    Application.StatusBar = "protective zone calculation..."
    DoEvents
    Call LogTrace("ANOD: running CalcProtectiveZone")
    On Error Resume Next
    Call Module_calcCP.CalcProtectiveZone(wsAnod)
    If Err.Number <> 0 Then
        ' ошибка при вызове calcprotectivezone: 
        ' error calling calcprotectivezone: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 043F 0440 0438 0020 0432 044B 0437 043E 0432 0435 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E") & Ru("0065 003A 0020") & Err.Description, vbCritical
        Call LogTrace("ANOD: error " & Err.Description)
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo 0
    Call LogTrace("ANOD: protective zone calculated")

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Call LogTrace("ANOD: test mode restored after CalcProtectiveZone")

    ' --- read currents ---
    Dim curCP As Double, curEndCP As Double
    Dim vCur As Variant, vCurEnd As Variant
    vCur = wsAnod.Range("currentCP").Cells(1, 1).Value
    vCurEnd = wsAnod.Range("currentEndLifeCP").Cells(1, 1).Value

    If IsError(vCur) Or Not IsNumeric(vCur) Then curCP = 0 Else curCP = CDbl(vCur)
    If IsError(vCurEnd) Or Not IsNumeric(vCurEnd) Then curEndCP = 0 Else curEndCP = CDbl(vCurEnd)

    Dim pipeCount As Variant
    pipeCount = wsAnod.Range("pipeCountCP").Value
    If IsError(pipeCount) Or Not IsNumeric(pipeCount) Then
        ' pipecountcp содержит ошибку или не число!
        ' pipecountcp contains an error or is not a number!
        MsgBox Ru("0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0441 043E 0434 0435 0440 0436 0438 0442 0020 043E 0448 0438 0431 043A 0443 0020 0438 043B 0438 0020 043D 0435 0020") & Ru("0447 0438 0441 043B 043E 0021"), vbCritical
        GoTo CleanExit
    End If
    Dim pipeCountLong As Long
    pipeCountLong = CLng(pipeCount)

    Dim expPipeCount As Long
    expPipeCount = 10
    Dim expCurCP As Double
    expCurCP = 1
    Dim expCurEndCP As Double
    expCurEndCP = 5

    Dim devN As Double
    If expPipeCount <> 0 Then devN = Abs(pipeCountLong - expPipeCount) / expPipeCount * 100 Else devN = 0
    ' защитная зона
    ' protective zone
    wsReport.Cells(row, 1).Value = Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430")
    ' nукз
    ' n_cp
    wsReport.Cells(row, 2).Value = Ru("006E 0443 043A 0437")
    wsReport.Cells(row, 3).Value = expPipeCount
    wsReport.Cells(row, 4).Value = pipeCountLong
    wsReport.Cells(row, 5).Value = devN
    ' пройден
    ' passed
    ' не пройден
    ' failed
    wsReport.Cells(row, 6).Value = IIf(devN <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
    row = row + 1

    Dim devI As Double
    If expCurCP <> 0 Then devI = Abs((curCP - expCurCP) / expCurCP) * 100 Else devI = 0
    ' защитная зона
    ' protective zone
    wsReport.Cells(row, 1).Value = Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430")
    ' iн
    ' in
    wsReport.Cells(row, 2).Value = Ru("0069 043D")
    wsReport.Cells(row, 3).Value = expCurCP
    wsReport.Cells(row, 4).Value = curCP
    wsReport.Cells(row, 5).Value = devI
    ' пройден
    ' passed
    ' не пройден
    ' failed
    wsReport.Cells(row, 6).Value = IIf(devI <= 10, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
    row = row + 1

    Dim devK As Double
    If expCurEndCP <> 0 Then devK = Abs((curEndCP - expCurEndCP) / expCurEndCP) * 100 Else devK = 0
    ' защитная зона
    ' protective zone
    wsReport.Cells(row, 1).Value = Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430")
    ' iк
    ' ik
    wsReport.Cells(row, 2).Value = Ru("0069 043A")
    wsReport.Cells(row, 3).Value = expCurEndCP
    wsReport.Cells(row, 4).Value = curEndCP
    wsReport.Cells(row, 5).Value = devK
    ' пройден
    ' passed
    ' не пройден
    ' failed
    wsReport.Cells(row, 6).Value = IIf(devK <= 10, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
    row = row + 1

    ' ================================================================
    ' SCENARIO DATA from ListAG TableAllTest (Russian IDs stay on the sheet)
    ' Table columns = scenarios (AG1..AG10). Data rows:
    '   1 material, 2 mounting, 3 installation, 4 delivery, 5 model
    ' Optional rows 6-9: expected R, N, T, G. Else ASCII defaults below.
    ' ================================================================
    Dim scenarios As Variant
    Dim nScenarios As Long
    scenarios = LoadScenariosFromTableAllTest()
    nScenarios = UBound(scenarios) - LBound(scenarios) + 1
    Call LogTrace("TEST: loaded " & nScenarios & " scenarios from " & TABLE_ALL_TEST)

    Dim col As Long
    Dim scIdx As Integer
    Dim startCol As Long
    startCol = 4

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ================================================================
    ' FILL ALL SCENARIO COLUMNS, THEN ONE Anod CALC
    ' ================================================================
    For scIdx = LBound(scenarios) To UBound(scenarios)
        If Timer - startTime > TIMEOUT_SECONDS Then
            Call LogTrace("SCENARIO: timeout exceeded at fill " & (scIdx + 1))
            MsgBox Ru("0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 0020 043F 043E 0020 0442 0430 0439 043C 002D 0430 0443 0442 0443 0020 0028 0431 043E 043B 0435 0435 0020") & TIMEOUT_SECONDS & Ru("0020 0441 0435 043A 0443 043D 0434 0029 002E") & vbCrLf & _
                   Ru("043F 0440 043E 0439 0434 0435 043D 043E 0020 0441 0446 0435 043D 0430 0440 0438 0435 0432 003A 0020") & scIdx & Ru("0020 0438 0437 0020") & nScenarios, vbCritical
            GoTo CleanExit
        End If

        col = startCol + scIdx
        Application.StatusBar = "filling scenario " & (scIdx + 1) & " of " & nScenarios & "... (column " & col & ")"
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": start (col=" & col & ")")

        wsAnod.Cells(34, col).Value = scenarios(scIdx)(1)
        wsAnod.Cells(35, col).Value = scenarios(scIdx)(2)
        wsAnod.Cells(36, col).Value = scenarios(scIdx)(3)
        wsAnod.Cells(37, col).Value = scenarios(scIdx)(4)
        wsAnod.Cells(38, col).Value = scenarios(scIdx)(5)
        wsAnod.Cells(51, col).Value = 1.6
        wsAnod.Cells(52, col).Value = 50
        wsAnod.Cells(53, col).Value = 50
        wsAnod.Cells(54, col).Value = 1
        wsAnod.Cells(55, col).Value = 0.77
        wsAnod.Cells(56, col).Value = 0.7
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": filters filled")

        Module_ValidationLogic.RefreshValidationForColumn wsAnod, col
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": validation refreshed")

        Module_AGData.InsertAGDataForColumn wsAnod, col
        UpdateCPData wsAnod, col
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": data inserted")

        If IsEmpty(wsAnod.Cells(39, col).Value) Then
            Call LogTrace("SCENARIO " & (scIdx + 1) & ": ERROR - row 39 is empty")
            MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0434 043B 044F 0020 043A 043E 043B 043E 043D 043A 0438 0020") & col & Ru("0020 043D 0435 0020 043F 043E 0434 0441 0442 0430 0432 043B 0435 043D 044B 0020 0434 0430 043D 043D 044B 0435 0020 0438 0437 0020 0054 0061 0062 006C 0065 0041 0047 0020 0028") & Ru("0441 0442 0440 043E 043A 0430 0020 0033 0039 0020 043F 0443 0441 0442 0430 0029 002E"), vbCritical
            GoTo CleanExit
        End If
    Next scIdx

    If Timer - startTime > TIMEOUT_SECONDS Then
        Call LogTrace("SCENARIO: timeout exceeded before anode calc")
        MsgBox Ru("0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 0020 043F 043E 0020 0442 0430 0439 043C 002D 0430 0443 0442 0443 0020 0028 0431 043E 043B 0435 0435 0020") & TIMEOUT_SECONDS & Ru("0020 0441 0435 043A 0443 043D 0434 0029 002E"), vbCritical
        GoTo CleanExit
    End If

    Application.StatusBar = "anode calculation (all scenarios)..."
    DoEvents
    Call Module_calcAnod.btnAnodFullCalc
    DoEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Call LogTrace("ANOD: one full calc for all scenario columns")

    For scIdx = LBound(scenarios) To UBound(scenarios)
        col = startCol + scIdx
        Application.StatusBar = "writing report " & (scIdx + 1) & " of " & nScenarios & "..."
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": reading results (col=" & col & ")")

        ' --- remember report start row for this scenario ---
        Dim reportStartRow As Long
        reportStartRow = row

        ' читаем результаты с постоянных строк (1 x pipeCountCP)
        ' read results from constant rows (1 x pipeCountCP)
        Dim R_p1 As Variant, N_el As Variant, T_p As Variant, G_total As Variant
        R_p1 = Empty: N_el = Empty: T_p = Empty: G_total = Empty

        Dim iTypeSc As Integer
        Dim isCombined As Boolean
        iTypeSc = Module_Visual.GetInstallationTypeIndex(CStr(scenarios(scIdx)(3)))
        ' комбинированный тип 3: Rp1 вертикальный (строка 60) и горизонтальный (строка 61) пишутся на лист
        ' combined type 3: vertical Rp1 (row 60) and horizontal Rp1 (row 61) are written to the sheet
        isCombined = (iTypeSc = 3)

        R_p1 = wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, col).Value

        N_el = wsAnod.Cells(ROW_NUM_ELECTRODES_AG, col).Value
        T_p = wsAnod.Cells(ROW_SERVICE_LIFE_AG, col).Value
        G_total = wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, col).Value

        ' модель не найдена
        ' model not found
        If IsEmpty(R_p1) Then R_p1 = Ru("043C 043E 0434 0435 043B 044C 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430")
        ' модель не найдена
        ' model not found
        If IsEmpty(N_el) Then N_el = Ru("043C 043E 0434 0435 043B 044C 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430")
        ' модель не найдена
        ' model not found
        If IsEmpty(T_p) Then T_p = Ru("043C 043E 0434 0435 043B 044C 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430")
        ' модель не найдена
        ' model not found
        If IsEmpty(G_total) Then G_total = Ru("043C 043E 0434 0435 043B 044C 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430")

        Dim rVal As Double, nVal As Double, tVal As Double, gVal As Double
        If IsNumeric(R_p1) Then rVal = CDbl(R_p1) Else rVal = 0
        If IsNumeric(N_el) Then nVal = CDbl(N_el) Else nVal = 0
        If IsNumeric(T_p) Then tVal = CDbl(T_p) Else tVal = 0
        If IsNumeric(G_total) Then gVal = CDbl(G_total) Else gVal = 0

        Dim expR As Double, expN As Double, expT As Double, expG As Double
        expR = scenarios(scIdx)(6)
        expN = scenarios(scIdx)(7)
        expT = scenarios(scIdx)(8)
        expG = scenarios(scIdx)(9)

        Dim scLabel As String
        ' колонка A: AG1 + модель (typeAG) + способ монтажа (typeInstallationAG)
        ' column A: AG1 + model (typeAG) + installation type (typeInstallationAG)
        scLabel = ScenarioReportLabel(scenarios(scIdx))

        ' --- write R_p1 ---
        wsReport.Cells(row, 1).Value = scLabel
        wsReport.Cells(row, 2).Value = "r_p1"
        wsReport.Cells(row, 3).Value = expR
        wsReport.Cells(row, 4).Value = R_p1
        If IsNumeric(R_p1) And expR <> 0 Then
            wsReport.Cells(row, 5).Value = Abs((rVal - expR) / expR) * 100
            ' пройден
            ' passed
            ' не пройден
            ' failed
            wsReport.Cells(row, 6).Value = IIf(wsReport.Cells(row, 5).Value <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
        Else
            wsReport.Cells(row, 5).Value = "-"
            ' не пройден
            ' failed
            ' не пройден (нет данных)
            ' failed (no data)
            wsReport.Cells(row, 6).Value = IIf(IsNumeric(R_p1), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029"))
        End If
        row = row + 1

        ' --- write N_el ---
        wsReport.Cells(row, 1).Value = scLabel
        ' n_э
        ' n_el
        wsReport.Cells(row, 2).Value = Ru("006E 005F 044D")
        wsReport.Cells(row, 3).Value = expN
        wsReport.Cells(row, 4).Value = N_el
        If IsNumeric(N_el) And expN <> 0 Then
            wsReport.Cells(row, 5).Value = Abs((nVal - expN) / expN) * 100
            ' пройден
            ' passed
            ' не пройден
            ' failed
            wsReport.Cells(row, 6).Value = IIf(wsReport.Cells(row, 5).Value <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
        Else
            wsReport.Cells(row, 5).Value = "-"
            ' не пройден
            ' failed
            ' не пройден (нет данных)
            ' failed (no data)
            wsReport.Cells(row, 6).Value = IIf(IsNumeric(N_el), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029"))
        End If
        row = row + 1

        ' --- write T_p ---
        wsReport.Cells(row, 1).Value = scLabel
        wsReport.Cells(row, 2).Value = "t_p"
        wsReport.Cells(row, 3).Value = expT
        wsReport.Cells(row, 4).Value = T_p
        If IsNumeric(T_p) And expT <> 0 Then
            wsReport.Cells(row, 5).Value = Abs((tVal - expT) / expT) * 100
            ' пройден
            ' passed
            ' не пройден
            ' failed
            wsReport.Cells(row, 6).Value = IIf(wsReport.Cells(row, 5).Value <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
        Else
            wsReport.Cells(row, 5).Value = "-"
            ' не пройден
            ' failed
            ' не пройден (нет данных)
            ' failed (no data)
            wsReport.Cells(row, 6).Value = IIf(IsNumeric(T_p), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029"))
        End If
        row = row + 1

        ' --- write G_total ---
        wsReport.Cells(row, 1).Value = scLabel
        ' g_общ
        ' g_total
        wsReport.Cells(row, 2).Value = Ru("0067 005F 043E 0431 0449")
        wsReport.Cells(row, 3).Value = expG
        wsReport.Cells(row, 4).Value = G_total
        If IsNumeric(G_total) And expG <> 0 Then
            wsReport.Cells(row, 5).Value = Abs((gVal - expG) / expG) * 100
            ' пройден
            ' passed
            ' не пройден
            ' failed
            wsReport.Cells(row, 6).Value = IIf(wsReport.Cells(row, 5).Value <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
        Else
            wsReport.Cells(row, 5).Value = "-"
            ' не пройден
            ' failed
            ' не пройден (нет данных)
            ' failed (no data)
            wsReport.Cells(row, 6).Value = IIf(IsNumeric(G_total), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029"))
        End If
        row = row + 1

        ' дополнительная проверка горизонтальной составляющей для комбинированного типа
        ' extra check of the horizontal component for the combined type
        If isCombined Then
            Dim R_horiz_calc As Variant
            R_horiz_calc = wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, col).Value
            If IsNumeric(R_horiz_calc) Then
                wsReport.Cells(row, 1).Value = ScenarioReportLabel(scenarios(scIdx), "horiz")
                wsReport.Cells(row, 2).Value = "r_p1"
                wsReport.Cells(row, 3).Value = 29.36
                wsReport.Cells(row, 4).Value = R_horiz_calc
                If CDbl(R_horiz_calc) > 0 Then
                    wsReport.Cells(row, 5).Value = Abs((CDbl(R_horiz_calc) - 29.36) / 29.36) * 100
                    ' пройден
                    ' passed
                    ' не пройден
                    ' failed
                    wsReport.Cells(row, 6).Value = IIf(wsReport.Cells(row, 5).Value <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
                Else
                    wsReport.Cells(row, 5).Value = "-"
                    ' не пройден (нет данных)
                    ' failed (no data)
                    wsReport.Cells(row, 6).Value = Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029")
                End If
                row = row + 1
            End If
        End If

        ' --- apply colors from wsAnod typeInstallationAG to report rows ---
        Call ApplyScenarioColorToReport(wsAnod, col, wsReport, reportStartRow, row - 1)
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": report rows written and colored")

    Next scIdx

    ' ================================================================
    ' FINAL VALIDATION REFRESH
    ' ================================================================
    Call LogTrace("FIN-01: before final validation loop")
    Dim pipeCountFinal As Long
    pipeCountFinal = CLng(wsAnod.Range("pipeCountCP").Value)
    Call Module_ValidationLogic.RefreshValidationForColumns(wsAnod, 4, 4 + pipeCountFinal - 1)
    Call LogTrace("FIN-03: after final validation loop")

    ' ================================================================
    ' FINAL REPORT FORMATTING
    ' ================================================================
    Call LogTrace("FIN-04: before AutoFit")
    wsReport.Columns("A").WrapText = False
    wsReport.Columns("A:F").AutoFit
    If row < 1 Then row = 1
    wsReport.Range("A1:F" & row).EntireRow.AutoFit
    Call LogTrace("FIN-05: after AutoFit")
    
    Call LogTrace("FIN-06: before Borders")
    wsReport.Range("A1:F" & row).Borders.LineStyle = xlContinuous
    Call LogTrace("FIN-07: after Borders")
    
    Call LogTrace("FIN-08: before Bold/Size")
    wsReport.Range("A4:F4").Font.Bold = True
    wsReport.Range("A1").Font.Size = 16
    Call LogTrace("FIN-09: after Bold/Size")

    ' ================================================================
    ' COUNT RESULTS
    ' ================================================================
    Call LogTrace("FIN-10: before CountIf")
    Dim passCount As Long, failCount As Long
    On Error Resume Next
    ' пройден
    ' passed
    passCount = Application.WorksheetFunction.CountIf(wsReport.Range("F5:F" & row), Ru("043F 0440 043E 0439 0434 0435 043D"))
    ' не пройден*
    ' failed*
    failCount = Application.WorksheetFunction.CountIf(wsReport.Range("F5:F" & row), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 002A"))
    On Error GoTo 0
    Call LogTrace("FIN-11: after CountIf, pass=" & passCount & ", fail=" & failCount)
    If DEBUG_MODE Then Debug.Print "RunAllTests: passed=" & passCount & ", failed=" & failCount
    Call DumpTestReportToImmediate(wsReport)
    dumpedReport = True

    Dim totalChecks As Long
    totalChecks = passCount + failCount
    If passCount < totalChecks Or failCount > 0 Then
        ' тест не пройден!
        ' test failed!
        ' пройдено: 
        ' passedо: 
        ' , не пройдено: 
        ' , failedо: 
        ' проверьте лист testreport для деталей.
        ' проверьте лandст testreport for details.
        MsgBox Ru("0442 0435 0441 0442 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0021") & vbCrLf & _
               Ru("043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & passCount & Ru("002C 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & failCount & vbCrLf & _
               Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 043B 0438 0441 0442 0020 0054 0065 0073 0074 0052 0065 0070 006F 0072 0074 0020 0434 043B 044F 0020 0434 0435 0442 0430 043B") & Ru("0435 0439 002E"), vbCritical
        GoTo CleanExit
    End If

    Call LogTrace("=== test report (" & nScenarios & " scenarios) ===")
    Call LogTrace("passed: " & passCount)
    Call LogTrace("failed: " & failCount)
    Call LogTrace("total checks: " & (passCount + failCount))

    Call LogTrace("FIN-12: before CheckNamedRanges")
    Call CheckNamedRanges
    Call LogTrace("FIN-13: after CheckNamedRanges")

    ' --- remember success status for the final message in CleanExit ---
    Dim testSucceeded As Boolean
    testSucceeded = True

    ' ================================================================
    ' fall through to CleanExit for state restoration
    ' ================================================================

CleanExit:
    Call LogTrace("CE-01: entered CleanExit")
    
    ' --- capture error info before any restore call clears it ---
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    On Error Resume Next
    
    ' --- restore programmatic mode for pipeCountCP ---
    Call LogTrace("CE-02: resetting isProgrammaticPipeCountChange")
    wsAnod.isProgrammaticPipeCountChange = False
    Call LogTrace("CE-03: isProgrammaticPipeCountChange = False")
    
    ' restore UI, but keep calc manual / events off until after MsgBox
    Call LogTrace("CE-04: restoring application state")
    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = oldStatusBar
    Application.Cursor = oldCursor
    Application.EnableCancelKey = oldEnableCancelKey
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Call LogTrace("CE-05: state restored (calc/events still test mode)")
    
    If Not wsReport Is Nothing Then
        Call LogTrace("CE-07: activating report sheet")
        wsReport.Activate
        Call LogTrace("CE-08: report sheet activated")
    End If
    
    If Not dumpedReport Then Call DumpTestReportToImmediate(wsReport)

    Call LogTrace("CE-12: before final MsgBox")
    
    If errNum <> 0 Then
        ' ошибка в тесте: 
        ' error in test: 
        ' код: 
        ' code: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0432 0020 0442 0435 0441 0442 0435 003A 0020") & errDesc & vbCrLf & Ru("043A 043E 0434 003A 0020") & errNum, vbCritical
    ElseIf testSucceeded Then
        ' тестирование завершено.
        ' testing finished.
        Dim doneMsg As String
        doneMsg = Ru("0442 0435 0441 0442 0438 0440 043E 0432 0430 043D 0438 0435 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E 002E") & vbCrLf & _
                  Ru("043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & passCount & Ru("002C 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & failCount & vbCrLf & _
                  Ru("043F 043E 0434 0440 043E 0431 043D 043E 0441 0442 0438 0020 043D 0430 0020 043B 0438 0441 0442 0435 0020 0054 0065 0073 0074 0052 0065 0070 006F 0072 0074 002E")
        If DEBUG_MODE Then
            doneMsg = doneMsg & vbCrLf & Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 0049 006D 006D 0065 0064 0069 0061 0074 0065 0020 0057 0069 006E 0064 006F 0077 0020 0434 043B 044F 0020 0434 0438 0430 0433") & Ru("043D 043E 0441 0442 0438 043A 0438 002E")
        End If
        MsgBox doneMsg, vbInformation
    End If
    
    Call LogTrace("CE-13: after final MsgBox")
    
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    
    testRunning = False
    Call LogTrace("CE-DONE: testRunning = False, exiting")
End Sub

' подпись сценария в колонке A: AG1, модель (typeAG), способ монтажа (typeInstallationAG)
' scenario label in column A: AG1, model (typeAG), installation type (typeInstallationAG)
Private Function ScenarioReportLabel(ByVal sc As Variant, Optional ByVal extra As String = "") As String
    Dim s As String
    s = CStr(sc(0)) & " | " & CStr(sc(5)) & " | " & CStr(sc(3))
    If Len(extra) > 0 Then s = s & " " & extra
    ScenarioReportLabel = s
End Function

' заполняем именованный диапазон трубы на все плечи (pipeDifferentParametersNum колонок)
' fill a pipe named range across all sections (pipeDifferentParametersNum columns)
Private Sub FillPipeNamedRange(ByVal rangeName As String, ByVal val As Variant)
    Dim rng As Range
    Dim c As Long
    Dim nCol As Long
    Set rng = wsPipe.Range(rangeName)
    nCol = CLng(wsPipe.Range("pipeDifferentParametersNum").Value)
    If nCol < 1 Then nCol = 1
    For c = 1 To nCol
        wsPipe.Cells(rng.row, START_COL + c - 1).Value = val
    Next c
End Sub

' ================================================================
' копируем заливку и шрифт typeInstallationAG (строка 36) на строки отчёта сценария
' apply background and font colors from wsAnod typeInstallationAG (row 36) to the report rows
' ================================================================
Private Sub ApplyScenarioColorToReport(ByVal wsAnod As Worksheet, ByVal col As Long, _
                                       ByVal wsReport As Worksheet, _
                                       ByVal reportStartRow As Long, ByVal reportEndRow As Long)
    On Error GoTo CleanExit
    
    Dim bgColor As Long
    Dim fontColor As Long
    
    ' читаем фактические цвета со строки 36 (typeInstallationAG)
    ' read actual colors from wsAnod row 36 (typeInstallationAG)
    bgColor = wsAnod.Cells(FILTER_START_ROW + 2, col).Interior.Color
    fontColor = wsAnod.Cells(FILTER_START_ROW + 2, col).Font.Color
    
    Dim rr As Long, cc As Long
    For rr = reportStartRow To reportEndRow
        For cc = 1 To 6
            wsReport.Cells(rr, cc).Interior.Color = bgColor
            wsReport.Cells(rr, cc).Font.Color = fontColor
        Next cc
    Next rr

CleanExit:
End Sub

' ================================================================
' расчётные имена Anod: 1 x pipeCountCP на постоянных строках
' Anod result names: 1 x pipeCountCP on constant rows
' ================================================================
Private Sub AnodResultNameSpecs(ByRef nmList As Variant, ByRef rowList As Variant)
    nmList = Array( _
        "resistanceEndLifeAG", _
        "lengthWorkPartDeepAG", _
        "oneElectrodeResistanceAG", _
        "oneElectrodeResistanceHorizAG", _
        "numElectrodesAG", _
        "weightWithoutFillingAG", _
        "serviceLifeAG", _
        "serviceLifeDeviation", _
        "correctResistanceAG")
    rowList = Array( _
        ROW_RESISTANCE_END_LIFE_AG, _
        ROW_LENGTH_WORK_PART_DEEP_AG, _
        ROW_ONE_ELECTRODE_RESISTANCE_AG, _
        ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, _
        ROW_NUM_ELECTRODES_AG, _
        ROW_WEIGHT_WITHOUT_FILLING_AG, _
        ROW_SERVICE_LIFE_AG, _
        ROW_SERVICE_LIFE_DEVIATION, _
        ROW_CORRECT_RESISTANCE_AG)
End Sub

' ================================================================
' проверка размерности именованных диапазонов
' check named ranges dimensions
' ================================================================
Private Sub CheckNamedRanges()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = wsAnod
    If ws Is Nothing Then
        Call LogTrace("CheckNamedRanges: Anod sheet not found")
        Exit Sub
    End If

    Dim pipeCount As Long
    Dim pipeVal As Variant
    pipeVal = ws.Range("pipeCountCP").Value
    If IsNumeric(pipeVal) And pipeVal > 0 Then
        pipeCount = CLng(pipeVal)
    Else
        pipeCount = 10
    End If

    Dim namesToCheck As Variant
    Dim rowsToCheck As Variant
    Call AnodResultNameSpecs(namesToCheck, rowsToCheck)
    Dim errorMsg As String
    Dim allOk As Boolean
    allOk = True

    Dim iNm As Long
    Dim nm As String
    Dim rng As Range
    Dim expectedRows As Long
    Dim expectedRow As Long
    expectedRows = 1

    For iNm = LBound(namesToCheck) To UBound(namesToCheck)
        nm = CStr(namesToCheck(iNm))
        expectedRow = CLng(rowsToCheck(iNm))
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.names(nm).RefersToRange
        If Err.Number <> 0 Then
            ' имя '
            ' name '
            ' ' не найдено!
            ' ' not found!
            errorMsg = errorMsg & vbCrLf & Ru("0438 043C 044F 0020 0027") & nm & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 0021")
            allOk = False
            Err.Clear
        Else
            If Not rng Is Nothing Then
                Dim actualRows As Long, actualCols As Long
                actualRows = rng.rows.count
                actualCols = rng.Columns.count
                If actualRows <> expectedRows Or actualCols <> pipeCount Or rng.row <> expectedRow Then
                    ' имя '
                    ' name '
                    ' ' имеет размерность 
                    ' ' has size 
                    '  строка 
                    '  row 
                    ' , ожидается 
                    ' , expected 
                    errorMsg = errorMsg & vbCrLf & Ru("0438 043C 044F 0020 0027") & nm & Ru("0027 0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & actualRows & "x" & actualCols & _
                               Ru("0020 0441 0442 0440 043E 043A 0430 0020") & rng.row & _
                               Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & expectedRows & "x" & pipeCount & _
                               Ru("0020 0441 0442 0440 043E 043A 0430 0020") & expectedRow
                    allOk = False
                End If
            End If
        End If
        On Error GoTo 0
    Next iNm

    If Not allOk Then
        Call LogTrace("CheckNamedRanges: errors found")
        Call LogTrace(errorMsg)
    Else
        Call LogTrace("CheckNamedRanges: all ranges have correct dimensions (" & expectedRows & "x" & pipeCount & " on constant rows)")
    End If
End Sub

' ================================================================
' Load scenarios from ListAG TableAllTest.
' Russian filter/model strings are read from the sheet (Windows-safe).
' Each item: name, material, mounting, installation, delivery, model, expR, expN, expT, expG
' ================================================================
Private Function LoadScenariosFromTableAllTest() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nCol As Long
    Dim nRow As Long
    Dim i As Long
    Dim body As Variant
    Dim scenarios() As Variant
    Dim expected As Variant
    Dim scName As String
    Dim expR As Double
    Dim expN As Double
    Dim expT As Double
    Dim expG As Double

    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 51, "LoadScenariosFromTableAllTest", _
                  "sheet '" & SHEET_LIST_AG & "' not found"
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_ALL_TEST)
    On Error GoTo 0
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 52, "LoadScenariosFromTableAllTest", _
                  "table '" & TABLE_ALL_TEST & "' not found on sheet " & SHEET_LIST_AG
    End If

    If tbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 53, "LoadScenariosFromTableAllTest", _
                  TABLE_ALL_TEST & " has no data rows"
    End If

    nCol = tbl.ListColumns.count
    nRow = tbl.ListRows.count
    If nCol < 1 Then
        Err.Raise vbObjectError + 54, "LoadScenariosFromTableAllTest", _
                  TABLE_ALL_TEST & " has no columns"
    End If
    If nRow < 5 Then
        Err.Raise vbObjectError + 55, "LoadScenariosFromTableAllTest", _
                  TABLE_ALL_TEST & " needs 5 data rows: material, mounting, installation, delivery, model"
    End If

    body = tbl.DataBodyRange.Value
    expected = DefaultExpectedResults()
    ReDim scenarios(0 To nCol - 1)

    For i = 1 To nCol
        scName = Trim$(CStr(tbl.HeaderRowRange.Cells(1, i).Value))
        If Len(scName) = 0 Then scName = "AG" & CStr(i)

        If nRow >= 9 Then
            expR = TableCellDbl(body, 6, i, 0)
            expN = TableCellDbl(body, 7, i, 0)
            expT = TableCellDbl(body, 8, i, 0)
            expG = TableCellDbl(body, 9, i, 0)
        ElseIf (i - 1) >= LBound(expected) And (i - 1) <= UBound(expected) Then
            expR = CDbl(expected(i - 1)(0))
            expN = CDbl(expected(i - 1)(1))
            expT = CDbl(expected(i - 1)(2))
            expG = CDbl(expected(i - 1)(3))
        Else
            expR = 0: expN = 0: expT = 0: expG = 0
        End If

        scenarios(i - 1) = Array( _
            scName, _
            TableCellText(body, 1, i), _
            TableCellText(body, 2, i), _
            TableCellText(body, 3, i), _
            TableCellText(body, 4, i), _
            TableCellText(body, 5, i), _
            expR, expN, expT, expG)
    Next i

    LoadScenariosFromTableAllTest = scenarios
End Function

Private Function DefaultExpectedResults() As Variant
    ' ASCII-only expected R, N, T, G for AG1..AG10
    DefaultExpectedResults = Array( _
        Array(26.2, 6, 92, 144), _
        Array(19.3, 5, 531, 27.5), _
        Array(37.65, 9, 464, 72), _
        Array(30.6, 7, 744, 38.5), _
        Array(16.8, 4, 206, 32), _
        Array(0.526, 0.12, 18, 33.6), _
        Array(0.527, 0.12, 162, 8.4), _
        Array(16.87, 4, 425, 22), _
        Array(30.6, 7, 744, 38.5), _
        Array(3.92, 0.86, 13, 24))
End Function

Private Function TableCellText(ByVal body As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    If IsArray(body) Then
        v = body(r, c)
    Else
        If r = 1 And c = 1 Then v = body Else v = Empty
    End If
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        TableCellText = ""
    Else
        TableCellText = Trim$(CStr(v))
    End If
End Function

Private Function TableCellDbl(ByVal body As Variant, ByVal r As Long, ByVal c As Long, ByVal defaultVal As Double) As Double
    Dim v As Variant
    If IsArray(body) Then
        v = body(r, c)
    Else
        If r = 1 And c = 1 Then v = body Else v = Empty
    End If
    If IsNumeric(v) Then
        TableCellDbl = CDbl(v)
    Else
        TableCellDbl = defaultVal
    End If
End Function
