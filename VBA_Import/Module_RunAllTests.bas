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

' ================================================================
' Clear the values in the listed named ranges on the Anod sheet.
' Waits (DoEvents) after each clear so that dependent formulas and
' validations have time to react before the test starts.
' ================================================================
Private Sub ClearNamedRangesForTest()
    Dim namesToClear As Variant
    namesToClear = Array( _
        "typeMaterial", _
        "typeMountingAG", _
        "typeInstallationAG", _
        "typeDeliveryAG", _
        "typeAG", _
        "diameterAG", _
        "lengthElectrodeAG", _
        "cokeBreezeDiameterAG", _
        "cokeBreezelengthElectrodeAG", _
        "massOneElectrodeAG", _
        "dissolutionRateAG", _
        "ratedCurrent", _
        "resistivityMaterialAG", _
        "cokeBreezeResistivityAG", _
        "specificRatedCurrent", _
        "specificMaccOneMeterAG", _
        "oneElectrodeResistanceAG", _
        "numElectrodesAG", _
        "weightWithoutFillingAG", _
        "serviceLifeAG", _
        "serviceLifeDeviation", _
        "correctResistanceAG" _
    )
    
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
        
        ' Give Excel time to process dependent formulas, validations, events
        DoEvents
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
    ' CHECK DIMENSIONS OF NAMED RANGES
    ' ================================================================
    Call LogTrace("CHECK: verifying named range dimensions")
    Dim dimCheckNames As Variant
    dimCheckNames = Array("oneElectrodeResistanceAG", "numElectrodesAG", "weightWithoutFillingAG", _
                          "serviceLifeAG", "serviceLifeDeviation", "correctResistanceAG")
    Dim dimName As Variant
    Dim dimRng As Range
    Dim expRows As Long
    expRows = 10
    Dim actRows As Long, actCols As Long
    Dim pipeCountCheck As Long
    pipeCountCheck = CLng(wsAnod.Range("pipeCountCP").Value)
    For Each dimName In dimCheckNames
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
        If actRows <> expRows Or actCols <> pipeCountCheck Then
            ' именованный диапазон '
            ' named range '
            ' ' имеет размерность 
            ' ' has size 
            ' , ожидается 
            ' , expected 
            ' . тест прерван.
            ' . test aborted.
            MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0439 0020 0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & dimName & Ru("0027 0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & actRows & "x" & actCols & _
                   Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & expRows & "x" & pipeCountCheck & Ru("002E 0020 0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 002E"), vbCritical
            GoTo CleanExit
        End If
    Next dimName
    Call LogTrace("CHECK: all named ranges have correct dimensions (" & expRows & "x" & pipeCountCheck & ")")

    ' ================================================================
    ' PREPARE DATA ON PIPE SHEET
    ' ================================================================
    Call LogTrace("PIPE: preparing data")
    DoEvents
    wsPipe.Range("pipeDifferentParametersNum").Value = 2
    DoEvents

    With wsPipe
        .Cells(6, 4).Value = 2.45E-07
        .Cells(7, 4).Value = 1.22
        .Cells(8, 4).Value = 0.017
        .Cells(9, 4).Value = 50000
        .Cells(10, 4).Value = 1.6
        .Cells(11, 4).Value = 50
        .Cells(12, 4).Value = 30
        .Cells(13, 4).Value = 0.11

        .Cells(6, 5).Value = 2.45E-07
        .Cells(7, 5).Value = 1.22
        .Cells(8, 5).Value = 0.017
        .Cells(9, 5).Value = 50000
        .Cells(10, 5).Value = 1.6
        .Cells(11, 5).Value = 50
        .Cells(12, 5).Value = 30
        .Cells(13, 5).Value = 0.11
    End With

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
        ' ================================================================
    ' LOOP OVER SCENARIOS
    ' ================================================================
    For scIdx = LBound(scenarios) To UBound(scenarios)
        ' --- timeout check (600 seconds) ---
        If Timer - startTime > TIMEOUT_SECONDS Then
            Call LogTrace("SCENARIO: timeout exceeded at scenario " & (scIdx + 1))
            ' тест прерван по тайм-ауту (более 
            ' test aborted on timeout (более 
            '  секунд).
            '  seconds).
            ' пройдено сценариев: 
            ' scenarios completed: 
            '  из 
            '  of 
            MsgBox Ru("0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 0020 043F 043E 0020 0442 0430 0439 043C 002D 0430 0443 0442 0443 0020 0028 0431 043E 043B 0435 0435 0020") & TIMEOUT_SECONDS & Ru("0020 0441 0435 043A 0443 043D 0434 0029 002E") & vbCrLf & _
                   Ru("043F 0440 043E 0439 0434 0435 043D 043E 0020 0441 0446 0435 043D 0430 0440 0438 0435 0432 003A 0020") & scIdx & Ru("0020 0438 0437 0020") & nScenarios, vbCritical
            GoTo CleanExit
        End If

        col = startCol + scIdx
        Application.StatusBar = "testing scenario " & (scIdx + 1) & " of " & nScenarios & "... (column " & col & ")"
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": start (col=" & col & ")")

        ' --- 1. fill filters ---
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

        ' --- 2. refresh validation ---
        Module_ValidationLogic.RefreshValidationForColumn wsAnod, col
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": validation refreshed")

        ' --- 3. insert data ---
        Module_AGData.InsertAGDataForColumn wsAnod, col
        UpdateCPData wsAnod, col
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": data inserted")

        ' --- 4. verify insertion ---
        If IsEmpty(wsAnod.Cells(39, col).Value) Then
            Call LogTrace("SCENARIO " & (scIdx + 1) & ": ERROR - row 39 is empty")
            ' ошибка: для колонки 
            ' error: for колонкand 
            '  не подставлены данные из tableag (строка 39 пуста).
            '  не data inserted from tableag (row 39 is empty).
            MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0434 043B 044F 0020 043A 043E 043B 043E 043D 043A 0438 0020") & col & Ru("0020 043D 0435 0020 043F 043E 0434 0441 0442 0430 0432 043B 0435 043D 044B 0020 0434 0430 043D 043D 044B 0435 0020 0438 0437 0020 0054 0061 0062 006C 0065 0041 0047 0020 0028") & Ru("0441 0442 0440 043E 043A 0430 0020 0033 0039 0020 043F 0443 0441 0442 0430 0029 002E"), vbCritical
            GoTo CleanExit
        End If

        ' --- 5. anode calculation ---
        Call Module_calcAnod.btnAnodFullCalc
        DoEvents
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": anode calculated")

        ' --- remember report start row for this scenario ---
        Dim reportStartRow As Long
        reportStartRow = row

        ' --- 6. read results ---
        Dim R_p1 As Variant, N_el As Variant, T_p As Variant, G_total As Variant
        Dim i As Long
        R_p1 = Empty: N_el = Empty: T_p = Empty: G_total = Empty

        Dim isCombined As Boolean
        isCombined = (StrComp(CStr(scenarios(scIdx)(0)), "AG4", vbTextCompare) = 0) Or (scIdx = 3)

        If Not isCombined Then
            For i = 1 To 10
                If Not IsEmpty(wsAnod.Cells(59 + i, col).Value) Then
                    R_p1 = wsAnod.Cells(59 + i, col).Value
                    Exit For
                End If
            Next i
        Else
            Dim l_el As Double, d_el As Double, h As Double, rho_soil As Double
            l_el = wsAnod.Cells(40, col).Value
            d_el = wsAnod.Cells(39, col).Value
            h = wsAnod.Cells(51, col).Value
            rho_soil = wsAnod.Cells(52, col).Value
            If IsNumeric(l_el) And IsNumeric(d_el) And IsNumeric(h) And IsNumeric(rho_soil) And l_el > 0 And d_el > 0 And (4 * h - l_el) > 0 Then
                Dim pi As Double
                pi = Application.pi()
                R_p1 = (rho_soil / (2 * pi * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
            Else
                ' нет данных
                ' no data
                R_p1 = Ru("043D 0435 0442 0020 0434 0430 043D 043D 044B 0445")
            End If
        End If

        For i = 1 To 10
            If Not IsEmpty(wsAnod.Cells(69 + i, col).Value) Then
                N_el = wsAnod.Cells(69 + i, col).Value
                Exit For
            End If
        Next i
        For i = 1 To 10
            If Not IsEmpty(wsAnod.Cells(89 + i, col).Value) Then
                T_p = wsAnod.Cells(89 + i, col).Value
                Exit For
            End If
        Next i
        For i = 1 To 10
            If Not IsEmpty(wsAnod.Cells(79 + i, col).Value) Then
                G_total = wsAnod.Cells(79 + i, col).Value
                Exit For
            End If
        Next i

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

        ' --- write R_p1 ---
        wsReport.Cells(row, 1).Value = scenarios(scIdx)(0)
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
        wsReport.Cells(row, 1).Value = scenarios(scIdx)(0)
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
        wsReport.Cells(row, 1).Value = scenarios(scIdx)(0)
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
        wsReport.Cells(row, 1).Value = scenarios(scIdx)(0)
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

        ' --- additional check for scenario 4 ---
        If scIdx = 3 Then
            Dim l_el_h As Double, d_el_h As Double, h_h As Double, rho_h As Double
            l_el_h = wsAnod.Cells(40, col).Value
            d_el_h = wsAnod.Cells(39, col).Value
            h_h = wsAnod.Cells(51, col).Value
            rho_h = wsAnod.Cells(52, col).Value
            If IsNumeric(l_el_h) And IsNumeric(d_el_h) And l_el_h > 0 And d_el_h > 0 Then
                Dim R_horiz_calc As Double
                Dim pi2 As Double
                pi2 = Application.pi()
                R_horiz_calc = (rho_h / (2 * pi2 * l_el_h)) * Log(2 * l_el_h / d_el_h)
                wsReport.Cells(row, 1).Value = CStr(scenarios(scIdx)(0)) & " horiz"
                wsReport.Cells(row, 2).Value = "r_p1"
                wsReport.Cells(row, 3).Value = 29.36
                wsReport.Cells(row, 4).Value = R_horiz_calc
                If R_horiz_calc > 0 Then
                    wsReport.Cells(row, 5).Value = Abs((R_horiz_calc - 29.36) / 29.36) * 100
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
    Dim colIdx As Long
    For colIdx = 4 To 4 + pipeCountFinal - 1
        Call LogTrace("FIN-02: RefreshValidationForColumn col=" & colIdx)
        Module_ValidationLogic.RefreshValidationForColumn wsAnod, colIdx
    Next colIdx
    Call LogTrace("FIN-03: after final validation loop")

    ' ================================================================
    ' FINAL REPORT FORMATTING
    ' ================================================================
    Call LogTrace("FIN-04: before AutoFit")
    wsReport.Columns("A:F").AutoFit
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
    
    ' --- restore application state (except EnableEvents) ---
    Call LogTrace("CE-04: restoring application state")
    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = oldStatusBar
    Application.Cursor = oldCursor
    Application.EnableCancelKey = oldEnableCancelKey
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation
    Call LogTrace("CE-05: state restored (except EnableEvents)")
    
    DoEvents
    Call LogTrace("CE-06: DoEvents passed")
    
    ' NOTE: Application.Calculate is intentionally removed - it may hang on complex sheets
    
    ' --- activate the report sheet BEFORE enabling events ---
    If Not wsReport Is Nothing Then
        Call LogTrace("CE-07: activating report sheet")
        wsReport.Activate
        wsReport.Range("A1").Select
        Call LogTrace("CE-08: report sheet activated")
    End If
    DoEvents
    
    ' --- now enable events (last step before the message) ---
    Call LogTrace("CE-09: enabling events")
    Application.EnableEvents = oldEnableEvents
    Call LogTrace("CE-10: events enabled")
    
    DoEvents
    Call LogTrace("CE-11: DoEvents before final MsgBox")
    
    ' --- ensure Excel window is in the foreground ---
    On Error Resume Next
    Application.WindowState = xlNormal
    On Error GoTo 0
    
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
        ' пройдено: 
        ' passedо: 
        ' , не пройдено: 
        ' , failedо: 
        ' подробности на листе testreport.
        ' details on sheet testreport.
        ' проверьте immediate window для диагностики.
        ' check the immediate window for diagnostics.
        MsgBox Ru("0442 0435 0441 0442 0438 0440 043E 0432 0430 043D 0438 0435 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E 002E") & vbCrLf & _
               Ru("043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & passCount & Ru("002C 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & failCount & vbCrLf & _
               Ru("043F 043E 0434 0440 043E 0431 043D 043E 0441 0442 0438 0020 043D 0430 0020 043B 0438 0441 0442 0435 0020 0054 0065 0073 0074 0052 0065 0070 006F 0072 0074 002E") & vbCrLf & _
               Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 0049 006D 006D 0065 0064 0069 0061 0074 0065 0020 0057 0069 006E 0064 006F 0077 0020 0434 043B 044F 0020 0434 0438 0430 0433") & Ru("043D 043E 0441 0442 0438 043A 0438 002E"), vbInformation
    End If
    
    Call LogTrace("CE-13: after final MsgBox")
    
    testRunning = False
    Call LogTrace("CE-DONE: testRunning = False, exiting")
End Sub

' ================================================================
' Apply background and font colors from the wsAnod typeInstallationAG
' (row 36) to the report rows of the current scenario
' ================================================================
Private Sub ApplyScenarioColorToReport(ByVal wsAnod As Worksheet, ByVal col As Long, _
                                       ByVal wsReport As Worksheet, _
                                       ByVal reportStartRow As Long, ByVal reportEndRow As Long)
    On Error GoTo CleanExit
    
    Dim bgColor As Long
    Dim fontColor As Long
    
    ' Read actual colors from wsAnod row 36 (typeInstallationAG)
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
' Check named ranges dimensions
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
    namesToCheck = Array("oneElectrodeResistanceAG", "numElectrodesAG", "weightWithoutFillingAG", _
                         "serviceLifeAG", "serviceLifeDeviation", "correctResistanceAG")
    Dim errorMsg As String
    Dim allOk As Boolean
    allOk = True

    Dim nm As Variant
    Dim rng As Range
    Dim expectedRows As Long
    expectedRows = 10

    For Each nm In namesToCheck
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
                If actualRows <> expectedRows Or actualCols <> pipeCount Then
                    ' имя '
                    ' name '
                    ' ' имеет размерность 
                    ' ' has size 
                    ' , ожидается 
                    ' , expected 
                    errorMsg = errorMsg & vbCrLf & Ru("0438 043C 044F 0020 0027") & nm & Ru("0027 0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & actualRows & "x" & actualCols & _
                               Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & expectedRows & "x" & pipeCount
                    allOk = False
                End If
            End If
        End If
        On Error GoTo 0
    Next nm

    If Not allOk Then
        Call LogTrace("CheckNamedRanges: errors found")
        Call LogTrace(errorMsg)
    Else
        Call LogTrace("CheckNamedRanges: all ranges have correct dimensions (" & expectedRows & "x" & pipeCount & ")")
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
