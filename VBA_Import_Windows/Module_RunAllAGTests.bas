Attribute VB_Name = "Module_RunAllAGTests"

Option Explicit

Private agTestRunning As Boolean

' ================================================================
' полный каталог: каждая модель typeAG из TableAG x все typeInstallationAG
' из TableMounting, допустимые для typeMountingAG этой модели.
' эталонов TableAllTest нет: пройден, если расчёт дал числовые R/N/T/G
' и строка 39 заполнена. лист отчёта: TestReportAG.
' catalog: each TableAG typeAG model x every TableMounting
' typeInstallationAG allowed for that model's typeMountingAG.
' no TableAllTest goldens: pass if calc yields numeric R/N/T/G
' and row 39 is filled. report sheet: TestReportAG.
' ================================================================
Public Sub RunAllAGTests()
    If agTestRunning Then
        ' тест уже выполняется. дождитесь завершения.
        ' test is already running. wait until it finishes.
        MsgBox Ru("0442 0435 0441 0442 0020 0443 0436 0435 0020 0432 044B 043F 043E 043B 043D 044F 0435 0442 0441 044F 002E 0020 0434 043E 0436 0434 0438 0442 0435 0441 044C 0020 0437 0430 0432") & Ru("0435 0440 0448 0435 043D 0438 044F 002E"), vbExclamation
        Exit Sub
    End If
    agTestRunning = True

    Call LogTrace("AGTEST: started")

    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCursor As XlMousePointer
    Dim oldEnableCancelKey As XlEnableCancelKey
    Dim dumpedReport As Boolean
    Dim testSucceeded As Boolean
    Dim passCount As Long
    Dim failCount As Long
    Dim wsReport As Worksheet
    Dim scenarios As Variant
    Dim nScenarios As Long

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCursor = Application.Cursor
    oldEnableCancelKey = Application.EnableCancelKey

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    wsAnod.isProgrammaticPipeCountChange = True
    Call LogTrace("AGTEST: programmatic mode enabled")

    Dim startTime As Double
    startTime = Timer
    ' 122 колонки: запас 30 минут
    ' 122 columns: 30 minute budget
    Const TIMEOUT_SECONDS As Long = 1800
    Const REPORT_SHEET As String = "TestReportAG"

    Call LogTrace("AGTEST: clearing named ranges on Anod before start")
    Call ClearNamedRangesForAGTest
    Call LogTrace("AGTEST: named ranges cleaned")

    On Error Resume Next
    Set wsReport = thisWorkbook.Worksheets(REPORT_SHEET)
    If wsReport Is Nothing Then
        Set wsReport = thisWorkbook.Worksheets.Add(After:=thisWorkbook.Worksheets(thisWorkbook.Worksheets.count))
        wsReport.name = REPORT_SHEET
    End If
    On Error GoTo CleanExit

    wsReport.Cells.Clear
    wsReport.Cells.Interior.ColorIndex = xlNone
    wsReport.Cells.Font.ColorIndex = xlAutomatic
    wsReport.Cells.Font.Bold = False
    wsReport.rows.Hidden = False
    wsReport.Columns.Hidden = False

    ' отчёт о тестировании эхз (
    ' cp design test report (
    wsReport.Range("A1").Value = Ru("043E 0442 0447 0451 0442 0020 043E 0020 0442 0435 0441 0442 0438 0440 043E 0432 0430 043D 0438 0438 0020 044D 0445 0437 0020 0028") & _
                                 TABLE_AG & " x " & TABLE_MOUNTING & ")"
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

    Call LogTrace("AGTEST: verifying local names")
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
        MsgBox Ru("043E 0442 0441 0443 0442 0441 0442 0432 0443 044E 0442 0020 043B 043E 043A 0430 043B 044C 043D 044B 0435 0020 0438 043C 0435 043D 0430 003A 0020") & missingAll & Ru("002E 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 0438 0442 0435 0020 0438 0445 0020 0438 0020 0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0442 0435 0441 0442 0020") & Ru("0441 043D 043E 0432 0430 002E"), vbCritical
        GoTo CleanExit
    Else
        ' успешно
        ' ok
        wsReport.Cells(row, 1).Value = Ru("0443 0441 043F 0435 0448 043D 043E")
        ' все локальные имена найдены
        ' all local names found
        wsReport.Cells(row, 2).Value = Ru("0432 0441 0435 0020 043B 043E 043A 0430 043B 044C 043D 044B 0435 0020 0438 043C 0435 043D 0430 0020 043D 0430 0439 0434 0435 043D 044B")
        ' пройден
        ' passed
        wsReport.Cells(row, 6).Value = Ru("043F 0440 043E 0439 0434 0435 043D")
        row = row + 1
    End If
    Call LogTrace("AGTEST: local names verified")

    ' ================================================================
    ' сценарии TableAG x TableMounting
    ' TableAG x TableMounting scenarios
    ' ================================================================
    scenarios = LoadScenariosFromTableAGAndMounting()
    nScenarios = UBound(scenarios) - LBound(scenarios) + 1
    Call LogTrace("AGTEST: loaded " & nScenarios & " scenarios from " & TABLE_AG & " x " & TABLE_MOUNTING)

    If nScenarios < 1 Then
        MsgBox Ru("043D 0435 0442 0020 0441 0446 0435 043D 0430 0440 0438 0435 0432 0020") & TABLE_AG & " x " & TABLE_MOUNTING & ".", vbCritical
        GoTo CleanExit
    End If
    If nScenarios > MAX_COL Then
        ' слишком много сочетаний (
        ' too many combinations (
        ' ). максимум колонок ukz:
        ' ). max ukz columns:
        MsgBox Ru("0441 043B 0438 0448 043A 043E 043C 0020 043C 043D 043E 0433 043E 0020 0441 043E 0447 0435 0442 0430 043D 0438 0439 0020 0028") & nScenarios & _
               Ru("0029 002E 0020 043C 0430 043A 0441 0438 043C 0443 043C 0020 043A 043E 043B 043E 043D 043E 043A 0020 0075 006B 007A 003A 0020") & MAX_COL & ".", vbCritical
        GoTo CleanExit
    End If

    ' каталог
    ' catalog
    wsReport.Cells(row, 1).Value = Ru("043A 0430 0442 0430 043B 043E 0433")
    wsReport.Cells(row, 2).Value = TABLE_AG & " x " & TABLE_MOUNTING
    wsReport.Cells(row, 3).Value = "-"
    wsReport.Cells(row, 4).Value = nScenarios
    wsReport.Cells(row, 5).Value = "-"
    wsReport.Cells(row, 6).Value = Ru("043F 0440 043E 0439 0434 0435 043D")
    row = row + 1

    ' ================================================================
    ' PREPARE DATA ON PIPE SHEET
    ' ================================================================
    Call LogTrace("PIPE: preparing data")
    Call SetStatusBar("pipe data...")
    DoEvents
    wsPipe.Range("pipeDifferentParametersNum").Value = 2
    DoEvents

    Call FillPipeNamedRangeAG("pipeSteelGrade", _
        Ru("0414 0430 043D 043D 044B 0435 0020 043E 0020 043C 0430 0440 043A 0435 0020 0441 0442 0430 043B 0438 0020 043E 0442 0441") & _
        Ru("0443 0442 0441 0442 0432 0443 044E 0442"))
    Call FillPipeNamedRangeAG("pipeSteelResistivity", 0.000000245)
    Call FillPipeNamedRangeAG("pipeDiameter", 1.22)
    Call FillPipeNamedRangeAG("pipeWallThickness", 0.017)
    Call FillPipeNamedRangeAG("pipeInsulationResistivityStartLife", 50000)
    Call FillPipeNamedRangeAG("pipeLayingDepth", 1.6)
    Call FillPipeNamedRangeAG("soilResistivityAvg", 50)
    Call FillPipeNamedRangeAG("serviceLifeDesigned", 30)
    Call FillPipeNamedRangeAG("pipeResistivityChangeFactor", 0.11)

    Call SetStatusBar("pipe column sync and validation...")
    DoEvents
    Call LogTrace("PIPE: ForceSyncFromD3")
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
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Call LogTrace("PIPE: columns synced")

    Call SetStatusBar("pipe calculation...")
    DoEvents
    Call LogTrace("PIPE: running btnPipeCalculate")
    On Error Resume Next
    Call Module_calcPipe.btnPipeCalculate
    If Err.Number <> 0 Then
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
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0070 0069 0070 0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074 0430 043D 0433 0435 0020 0438 043B 0438 0020 0070 0069 0070") & Ru("0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074 0430 043D 0433 0435 0045 006E 0064 004C 0069 0066 0065 0020 0441 043E 0434 0435 0440 0436 0430 0442 0020 043E 0448") & Ru("0438 0431 043A 0443 002C 0020 043D 0435 0020 0447 0438 0441 043B 043E 0020 0438 043B 0438 0020 0440 0430 0432 043D 044B 0020 0030 0021"), vbCritical
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
    wsReport.Cells(row, 6).Value = IIf(devR <= 5, Ru("043F 0440 043E 0439 0434 0435 043D"), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D"))
    row = row + 1

    ' ================================================================
    ' PREPARE ANOD SHEET
    ' ================================================================
    Call LogTrace("ANOD: preparing general parameters")
    wsAnod.Range("minProtectPotential").Value = -0.85
    wsAnod.Range("maxProtectPotential").Value = -1.15
    wsAnod.Range("naturalPotential").Value = -0.55
    wsAnod.Range("factorMutualInfluence").Value = 0.5
    wsAnod.Range("pipeLength").Value = 350000

    Call SetStatusBar("protective zone calculation...")
    DoEvents
    Call LogTrace("ANOD: running CalcProtectiveZone")
    On Error Resume Next
    Call Module_calcCP.CalcProtectiveZone(wsAnod)
    If Err.Number <> 0 Then
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

    Dim vCur As Variant, vCurEnd As Variant
    vCur = wsAnod.Range("currentCP").Cells(1, 1).Value
    vCurEnd = wsAnod.Range("currentEndLifeCP").Cells(1, 1).Value

    ' защитная зона (токи, без эталона TableAllTest)
    ' protective zone (currents, no TableAllTest goldens)
    Call WriteSanityRow(wsReport, row, Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430"), Ru("0069 043D"), vCur)
    row = row + 1
    Call WriteSanityRow(wsReport, row, Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430"), Ru("0069 043A"), vCurEnd)
    row = row + 1

    ' колонки Anod = число сценариев каталога
    ' Anod columns = catalog scenario count
    Call SetStatusBar("sync pipeCountCP = " & nScenarios & " ...")
    DoEvents
    Call LogTrace("ANOD: SetPipeCountCPProgrammatically " & nScenarios)
    On Error Resume Next
    Call wsAnod.SetPipeCountCPProgrammatically(nScenarios)
    If Err.Number <> 0 Then
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0443 0441 0442 0430 043D 043E 0432 043A 0438 0020 0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 003A 0020") & Err.Description, vbCritical
        Call LogTrace("ANOD: pipeCount set error " & Err.Description)
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo 0

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    wsAnod.isProgrammaticPipeCountChange = True

    Dim pipeCountLong As Long
    Dim pipeCount As Variant
    pipeCount = wsAnod.Range("pipeCountCP").Value
    If IsError(pipeCount) Or Not IsNumeric(pipeCount) Then
        MsgBox Ru("0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0441 043E 0434 0435 0440 0436 0438 0442 0020 043E 0448 0438 0431 043A 0443 0020 0438 043B 0438 0020 043D 0435 0020") & Ru("0447 0438 0441 043B 043E 0021"), vbCritical
        GoTo CleanExit
    End If
    pipeCountLong = CLng(pipeCount)
    If pipeCountLong <> nScenarios Then
        Call LogTrace("ANOD: pipeCountCP=" & pipeCountLong & " expected " & nScenarios)
        MsgBox "pipeCountCP=" & pipeCountLong & " / " & nScenarios, vbCritical
        GoTo CleanExit
    End If

    ' nукз
    ' n_cp
    wsReport.Cells(row, 1).Value = Ru("0437 0430 0449 0438 0442 043D 0430 044F 0020 0437 043E 043D 0430")
    wsReport.Cells(row, 2).Value = Ru("006E 0443 043A 0437")
    wsReport.Cells(row, 3).Value = nScenarios
    wsReport.Cells(row, 4).Value = pipeCountLong
    wsReport.Cells(row, 5).Value = 0
    wsReport.Cells(row, 6).Value = Ru("043F 0440 043E 0439 0434 0435 043D")
    row = row + 1

    If Not CheckAnodResultDims(pipeCountLong) Then GoTo CleanExit

    Dim col As Long
    Dim scIdx As Long
    Dim startCol As Long
    startCol = START_COL

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ================================================================
    ' FILL ALL CATALOG COLUMNS, THEN ONE Anod CALC
    ' ================================================================
    For scIdx = LBound(scenarios) To UBound(scenarios)
        If Timer - startTime > TIMEOUT_SECONDS Then
            Call LogTrace("SCENARIO: timeout exceeded at fill " & (scIdx + 1))
            MsgBox Ru("0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 0020 043F 043E 0020 0442 0430 0439 043C 002D 0430 0443 0442 0443 0020 0028 0431 043E 043B 0435 0435 0020") & TIMEOUT_SECONDS & Ru("0020 0441 0435 043A 0443 043D 0434 0029 002E") & vbCrLf & _
                   Ru("043F 0440 043E 0439 0434 0435 043D 043E 0020 0441 0446 0435 043D 0430 0440 0438 0435 0432 003A 0020") & scIdx & Ru("0020 0438 0437 0020") & nScenarios, vbCritical
            GoTo CleanExit
        End If

        col = startCol + scIdx
        Call PulseProgress("filling AG catalog", scIdx + 1, nScenarios)
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

        Module_ValidationLogic.RefreshValidationForColumn wsAnod, col
        Module_AGData.InsertAGDataForColumn wsAnod, col
        UpdateCPData wsAnod, col

        If IsEmpty(wsAnod.Cells(39, col).Value) Then
            Call LogTrace("SCENARIO " & (scIdx + 1) & ": row 39 empty after InsertAGData")
        End If
    Next scIdx

    If Timer - startTime > TIMEOUT_SECONDS Then
        Call LogTrace("SCENARIO: timeout exceeded before anode calc")
        MsgBox Ru("0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 0020 043F 043E 0020 0442 0430 0439 043C 002D 0430 0443 0442 0443 0020 0028 0431 043E 043B 0435 0435 0020") & TIMEOUT_SECONDS & Ru("0020 0441 0435 043A 0443 043D 0434 0029 002E"), vbCritical
        GoTo CleanExit
    End If

    Call SetStatusBar("anode calculation (all catalog columns)...")
    DoEvents
    Call Module_calcAnod.btnAnodFullCalc
    DoEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Call LogTrace("ANOD: one full calc for all catalog columns")

    For scIdx = LBound(scenarios) To UBound(scenarios)
        col = startCol + scIdx
        Call PulseProgress("writing AG catalog report", scIdx + 1, nScenarios)

        Dim reportStartRow As Long
        reportStartRow = row

        Dim R_p1 As Variant, N_el As Variant, T_p As Variant, G_total As Variant
        R_p1 = Empty: N_el = Empty: T_p = Empty: G_total = Empty

        Dim iTypeSc As Integer
        Dim isCombined As Boolean
        iTypeSc = Module_Visual.GetInstallationTypeIndex(CStr(scenarios(scIdx)(3)))
        isCombined = (iTypeSc = 3)

        R_p1 = wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, col).Value
        N_el = wsAnod.Cells(ROW_NUM_ELECTRODES_AG, col).Value
        T_p = wsAnod.Cells(ROW_SERVICE_LIFE_AG, col).Value
        G_total = wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, col).Value

        Dim scLabel As String
        scLabel = ScenarioReportLabelAG(scenarios(scIdx))

        Call WriteSanityRow(wsReport, row, scLabel, "r_p1", R_p1)
        row = row + 1
        Call WriteSanityRow(wsReport, row, scLabel, Ru("006E 005F 044D"), N_el)
        row = row + 1
        Call WriteSanityRow(wsReport, row, scLabel, "t_p", T_p)
        row = row + 1
        Call WriteSanityRow(wsReport, row, scLabel, Ru("0067 005F 043E 0431 0449"), G_total)
        row = row + 1

        If isCombined Then
            Dim R_horiz_calc As Variant
            R_horiz_calc = wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, col).Value
            Call WriteSanityRow(wsReport, row, ScenarioReportLabelAG(scenarios(scIdx), "horiz"), "r_p1", R_horiz_calc)
            row = row + 1
        End If

        If IsEmpty(wsAnod.Cells(39, col).Value) Then
            Call WriteSanityRow(wsReport, row, scLabel, "AG data r39", Empty)
            row = row + 1
        End If

        Call ApplyScenarioColorToReportAG(wsAnod, col, wsReport, reportStartRow, row - 1)
        Call LogTrace("SCENARIO " & (scIdx + 1) & ": report rows written")
    Next scIdx

    Call LogTrace("FIN-01: before final validation")
    Call Module_ValidationLogic.RefreshValidationForColumns(wsAnod, START_COL, START_COL + nScenarios - 1)
    Call LogTrace("FIN-03: after final validation")

    wsReport.Columns("A").WrapText = False
    wsReport.Columns("A:F").AutoFit
    wsReport.Range("A1:F" & row).Borders.LineStyle = xlContinuous
    wsReport.Range("A4:F4").Font.Bold = True
    wsReport.Range("A1").Font.Size = 16

    Call LogTrace("FIN-10: before CountIf")
    On Error Resume Next
    passCount = Application.WorksheetFunction.CountIf(wsReport.Range("F5:F" & row), Ru("043F 0440 043E 0439 0434 0435 043D"))
    failCount = Application.WorksheetFunction.CountIf(wsReport.Range("F5:F" & row), Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 002A"))
    On Error GoTo 0
    Call LogTrace("FIN-11: after CountIf, pass=" & passCount & ", fail=" & failCount)
    If DEBUG_MODE Then Debug.Print "RunAllAGTests: passed=" & passCount & ", failed=" & failCount
    Call DumpAGTestReportToImmediate(wsReport)
    dumpedReport = True

    Dim totalChecks As Long
    totalChecks = passCount + failCount
    If passCount < totalChecks Or failCount > 0 Then
        MsgBox Ru("0442 0435 0441 0442 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0021") & vbCrLf & _
               Ru("043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & passCount & Ru("002C 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & failCount & vbCrLf & _
               Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 043B 0438 0441 0442 0020") & REPORT_SHEET & Ru("0020 0434 043B 044F 0020 0434 0435 0442 0430 043B") & Ru("0435 0439 002E"), vbCritical
        GoTo CleanExit
    End If

    Call LogTrace("=== AG catalog report (" & nScenarios & " scenarios) ===")
    testSucceeded = True

CleanExit:
    Call LogTrace("AGTEST CE-01: entered CleanExit")

    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    On Error Resume Next

    Call LogTrace("AGTEST CE-02: resetting isProgrammaticPipeCountChange")
    wsAnod.isProgrammaticPipeCountChange = False

    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = oldStatusBar
    Application.Cursor = oldCursor
    Application.EnableCancelKey = oldEnableCancelKey
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    If Not wsReport Is Nothing Then
        wsReport.Activate
    End If

    If Not dumpedReport Then Call DumpAGTestReportToImmediate(wsReport)

    If errNum <> 0 Then
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0432 0020 0442 0435 0441 0442 0435 003A 0020") & errDesc & vbCrLf & Ru("043A 043E 0434 003A 0020") & errNum, vbCritical
    ElseIf testSucceeded Then
        Dim doneMsg As String
        doneMsg = Ru("0442 0435 0441 0442 0438 0440 043E 0432 0430 043D 0438 0435 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E 002E") & vbCrLf & _
                  Ru("043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & passCount & Ru("002C 0020 043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 043E 003A 0020") & failCount & vbCrLf & _
                  Ru("043F 043E 0434 0440 043E 0431 043D 043E 0441 0442 0438 0020 043D 0430 0020 043B 0438 0441 0442 0435 0020") & "TestReportAG."
        If DEBUG_MODE Then
            doneMsg = doneMsg & vbCrLf & Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 0049 006D 006D 0065 0064 0069 0430 0442 0435 0020 0057 0069 006E 0064 006F 0077 0020 0434 043B 044F 0020 0434 0438 0430 0433") & Ru("043D 043E 0441 0442 0438 043A 0438 002E")
        End If
        MsgBox doneMsg, vbInformation
    End If

    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation

    agTestRunning = False
    Call LogTrace("AGTEST CE-DONE")
End Sub

' ================================================================
' TableAG (AG_Material / AG_MountType / AG_Completion / AG_Model)
' x TableMounting col4 typeMountingAG -> col3 typeInstallationAG
' ================================================================
Private Function LoadScenariosFromTableAGAndMounting() As Variant
    Dim matVals As Variant, mountVals As Variant, shipVals As Variant, modelVals As Variant
    Dim byMount As Collection
    Dim scCol As Collection
    Dim instCol As Collection
    Dim nAg As Long
    Dim iAg As Long
    Dim iInst As Long
    Dim matRaw As String, mountRaw As String, shipRaw As String, modelRaw As String
    Dim mountKey As String
    Dim instRaw As String
    Dim scenarios() As Variant
    Dim iSc As Long

    matVals = LoadName2DAG("AG_Material")
    mountVals = LoadName2DAG("AG_MountType")
    shipVals = LoadName2DAG("AG_Completion")
    modelVals = LoadName2DAG("AG_Model")
    If IsEmpty(matVals) Or IsEmpty(mountVals) Or IsEmpty(shipVals) Or IsEmpty(modelVals) Then
        Err.Raise vbObjectError + 61, "LoadScenariosFromTableAGAndMounting", _
                  "named ranges AG_Material / AG_MountType / AG_Completion / AG_Model not found"
    End If

    Set byMount = LoadInstallsByMountType()
    If byMount Is Nothing Then
        Err.Raise vbObjectError + 62, "LoadScenariosFromTableAGAndMounting", _
                  "table '" & TABLE_MOUNTING & "' not found"
    End If

    nAg = UBound(matVals, 1)
    Set scCol = New Collection

    For iAg = 1 To nAg
        modelRaw = CellTextAG(modelVals, iAg)
        If Len(modelRaw) = 0 Then GoTo NextAgRow
        matRaw = CellTextAG(matVals, iAg)
        mountRaw = CellTextAG(mountVals, iAg)
        shipRaw = CellTextAG(shipVals, iAg)
        mountKey = Module_ValidationLists.NormalizeFilterText(mountRaw)

        Set instCol = Nothing
        On Error Resume Next
        Set instCol = byMount(mountKey)
        On Error GoTo 0
        If instCol Is Nothing Then
            Call LogTrace("AGTEST: no TableMounting installs for mount='" & mountRaw & "' model='" & modelRaw & "'")
            GoTo NextAgRow
        End If

        For iInst = 1 To instCol.Count
            instRaw = CStr(instCol(iInst))
            scCol.Add Array("AG" & CStr(scCol.Count + 1), matRaw, mountRaw, instRaw, shipRaw, modelRaw)
        Next iInst
NextAgRow:
    Next iAg

    If scCol.Count < 1 Then
        Err.Raise vbObjectError + 63, "LoadScenariosFromTableAGAndMounting", _
                  "no " & TABLE_AG & " x " & TABLE_MOUNTING & " combinations"
    End If

    ReDim scenarios(0 To scCol.Count - 1)
    For iSc = 1 To scCol.Count
        scenarios(iSc - 1) = scCol(iSc)
    Next iSc

    LoadScenariosFromTableAGAndMounting = scenarios
End Function

' TableMounting: ListColumns(4)=typeMountingAG, ListColumns(3)=typeInstallationAG
Private Function LoadInstallsByMountType() As Collection
    Dim tbl As ListObject
    Dim mVals As Variant, iVals As Variant
    Dim byMount As Collection
    Dim instCol As Collection
    Dim nRow As Long
    Dim i As Long
    Dim mountKey As String
    Dim mountRaw As String, instRaw As String

    Set tbl = FindTableMounting()
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    mVals = As2DAG(tbl.ListColumns(4).DataBodyRange.Value)
    iVals = As2DAG(tbl.ListColumns(3).DataBodyRange.Value)
    nRow = UBound(mVals, 1)

    Set byMount = New Collection
    For i = 1 To nRow
        mountRaw = CellTextAG(mVals, i)
        instRaw = CellTextAG(iVals, i)
        If Len(mountRaw) = 0 Or Len(instRaw) = 0 Then GoTo NextMountRow
        mountKey = Module_ValidationLists.NormalizeFilterText(mountRaw)
        If Len(mountKey) = 0 Then GoTo NextMountRow

        Set instCol = Nothing
        On Error Resume Next
        Set instCol = byMount(mountKey)
        On Error GoTo 0
        If instCol Is Nothing Then
            Set instCol = New Collection
            byMount.Add instCol, mountKey
        End If
        On Error Resume Next
        instCol.Add instRaw, instRaw
        Err.Clear
        On Error GoTo 0
NextMountRow:
    Next i

    Set LoadInstallsByMountType = byMount
End Function

Private Function FindTableMounting() As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(TABLE_MOUNTING)
    If Not tbl Is Nothing Then
        Set FindTableMounting = tbl
        On Error GoTo 0
        Exit Function
    End If

    For Each ws In thisWorkbook.Worksheets
        Set tbl = Nothing
        Set tbl = ws.ListObjects(TABLE_MOUNTING)
        If Not tbl Is Nothing Then
            Set FindTableMounting = tbl
            On Error GoTo 0
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Private Function LoadName2DAG(ByVal nm As String) As Variant
    Dim rng As Range
    On Error Resume Next
    Set rng = thisWorkbook.names(nm).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    LoadName2DAG = As2DAG(rng.Value)
End Function

Private Function As2DAG(ByVal vals As Variant) As Variant
    If IsArray(vals) Then
        As2DAG = vals
    Else
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = vals
        As2DAG = arr
    End If
End Function

Private Function CellTextAG(ByVal arr As Variant, ByVal iRow As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = arr(iRow, 1)
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then
        CellTextAG = ""
    Else
        CellTextAG = Trim$(CStr(v))
    End If
End Function

Private Function HasNumericValue(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    HasNumericValue = IsNumeric(v)
End Function

Private Sub WriteSanityRow(ByVal wsReport As Worksheet, ByVal r As Long, _
                           ByVal scLabel As String, ByVal paramName As String, _
                           ByVal actual As Variant)
    wsReport.Cells(r, 1).Value = scLabel
    wsReport.Cells(r, 2).Value = paramName
    wsReport.Cells(r, 3).Value = "-"
    wsReport.Cells(r, 4).Value = actual
    wsReport.Cells(r, 5).Value = "-"
    If HasNumericValue(actual) Then
        ' пройден
        ' passed
        wsReport.Cells(r, 6).Value = Ru("043F 0440 043E 0439 0434 0435 043D")
    Else
        ' не пройден (нет данных)
        ' failed (no data)
        wsReport.Cells(r, 6).Value = Ru("043D 0435 0020 043F 0440 043E 0439 0434 0435 043D 0020 0028 043D 0435 0442 0020 0434 0430 043D 043D 044B 0445 0029")
    End If
End Sub

Private Function ScenarioReportLabelAG(ByVal sc As Variant, Optional ByVal extra As String = "") As String
    Dim s As String
    s = CStr(sc(0)) & " | " & CStr(sc(5)) & " | " & CStr(sc(3))
    If Len(extra) > 0 Then s = s & " " & extra
    ScenarioReportLabelAG = s
End Function

Private Sub FillPipeNamedRangeAG(ByVal rangeName As String, ByVal val As Variant)
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

Private Sub ApplyScenarioColorToReportAG(ByVal wsSrc As Worksheet, ByVal col As Long, _
                                         ByVal wsRep As Worksheet, _
                                         ByVal reportStartRow As Long, ByVal reportEndRow As Long)
    On Error GoTo CleanExit

    Dim bgColor As Long
    Dim fontColor As Long
    bgColor = wsSrc.Cells(FILTER_START_ROW + 2, col).Interior.Color
    fontColor = wsSrc.Cells(FILTER_START_ROW + 2, col).Font.Color

    Dim rr As Long, cc As Long
    For rr = reportStartRow To reportEndRow
        For cc = 1 To 6
            wsRep.Cells(rr, cc).Interior.Color = bgColor
            wsRep.Cells(rr, cc).Font.Color = fontColor
        Next cc
    Next rr

CleanExit:
End Sub

Private Function CheckAnodResultDims(ByVal pipeCountCheck As Long) As Boolean
    Dim nmList As Variant
    Dim rowList As Variant
    Dim dimIdx As Long
    Dim dimName As String
    Dim dimRng As Range
    Dim actRows As Long, actCols As Long
    Dim expRow As Long

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

    For dimIdx = LBound(nmList) To UBound(nmList)
        dimName = CStr(nmList(dimIdx))
        expRow = CLng(rowList(dimIdx))
        On Error Resume Next
        Set dimRng = Nothing
        Set dimRng = wsAnod.names(dimName).RefersToRange
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0439 0020 0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & dimName & Ru("0027 0020 043D 0435 0020 0441 0443 0449 0435 0441 0442 0432 0443 0435 0442 0021"), vbCritical
            Exit Function
        End If
        On Error GoTo 0
        actRows = dimRng.rows.count
        actCols = dimRng.Columns.count
        If actRows <> 1 Or actCols <> pipeCountCheck Or dimRng.row <> expRow Then
            MsgBox Ru("0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0439 0020 0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & dimName & Ru("0027 0020 0438 043C 0435 0435 0442 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020") & actRows & "x" & actCols & _
                   Ru("0020 0441 0442 0440 043E 043A 0430 0020") & dimRng.row & _
                   Ru("002C 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & "1x" & pipeCountCheck & _
                   Ru("0020 0441 0442 0440 043E 043A 0430 0020") & expRow & Ru("002E 0020 0442 0435 0441 0442 0020 043F 0440 0435 0440 0432 0430 043D 002E"), vbCritical
            Exit Function
        End If
    Next dimIdx
    CheckAnodResultDims = True
End Function

Private Sub ClearNamedRangesForAGTest()
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

    Dim nm As Variant
    For Each nm In namesToClear
        On Error Resume Next
        Dim rng As Range
        Set rng = Nothing
        Set rng = wsAnod.Range(CStr(nm))
        If Not rng Is Nothing Then rng.ClearContents
        Err.Clear
        On Error GoTo 0
    Next nm
    DoEvents
End Sub

Private Sub DumpAGTestReportToImmediate(ByVal wsReport As Worksheet)
    If Not DEBUG_MODE Then Exit Sub
    If wsReport Is Nothing Then Exit Sub
    On Error Resume Next

    Dim lastRow As Long
    Dim lastF As Long
    lastRow = wsReport.Cells(wsReport.rows.count, 1).End(xlUp).row
    lastF = wsReport.Cells(wsReport.rows.count, 6).End(xlUp).row
    If lastF > lastRow Then lastRow = lastF
    If lastRow < 1 Then Exit Sub

    Debug.Print "=== TestReportAG ==="
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
    Debug.Print "=== TestReportAG end ==="
    On Error GoTo 0
End Sub
