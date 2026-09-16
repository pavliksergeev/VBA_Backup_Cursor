Attribute VB_Name = "Module_calcCP"

' ================================================================
' module: Module_calcCP
' purpose: cathodic-protection (CP) calculations
' ================================================================
Option Explicit

' ================================================================
' evenly distribute pipeline length among CP stations
' ================================================================
Public Sub DistributeLength(ByVal wsAnod As Worksheet)
    On Error GoTo CleanExit
    
    Application.StatusBar = "distributing pipe length between CPs"
    DoEvents
    
    Dim wsPipe As Worksheet
    Dim colCP As Long
    Dim arr As Variant
    Dim colCount As Long
    Dim rng As Range
    Dim originalCalc As XlCalculation
    Dim originalScreenUpdating As Boolean
    Dim originalEvents As Boolean
    
    If wsAnod Is Nothing Then
        ' ошибка: лист 
        ' error: лandст 
        '  не передан!
        '  not passed!
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 043B 0438 0441 0442 0020") & SHEET_ANOD & Ru("0020 043D 0435 0020 043F 0435 0440 0435 0434 0430 043D 0021"), vbCritical
        GoTo CleanExit
    End If
    
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    
    originalCalc = Application.Calculation
    originalScreenUpdating = Application.ScreenUpdating
    originalEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    arr = wsAnod.Range("pipeLengthCP").Value
    If Err.Number <> 0 Then
        ' диапазон 'pipelengthcp' не найден!
        ' range 'pipelengthcp' not found!
        MsgBox Ru("0434 0438 0430 043F 0430 0437 043E 043D 0020 0027 0070 0069 0070 0065 004C 0065 006E 0067 0074 0068 0043 0050 0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo CleanExit
    
    Set rng = wsAnod.Range("pipeLengthCP")
    colCount = rng.Columns.count
    
    If colCount < 1 Then
        ' диапазон 'pipelengthcp' не содержит колонок!
        ' range 'pipelengthcp' has no columns!
        MsgBox Ru("0434 0438 0430 043F 0430 0437 043E 043D 0020 0027 0070 0069 0070 0065 004C 0065 006E 0067 0074 0068 0043 0050 0027 0020 043D 0435 0020 0441 043E 0434 0435 0440 0436 0438 0442") & Ru("0020 043A 043E 043B 043E 043D 043E 043A 0021"), vbExclamation
        GoTo CleanExit
    End If
    
    Dim pipeCount As Variant
    On Error Resume Next
    pipeCount = wsAnod.Range("pipeCountCP").Value
    On Error GoTo CleanExit
    
    If IsEmpty(pipeCount) Or Not IsNumeric(pipeCount) Or pipeCount <= 0 Then
        ' 'pipecountcp' должен быть положительным числом!
        ' 'pipecountcp' must be a positive number!
        MsgBox Ru("0027 0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0027 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043F 043E 043B 043E 0436 0438 0442 0435 043B") & Ru("044C 043D 044B 043C 0020 0447 0438 0441 043B 043E 043C 0021"), vbExclamation
        GoTo CleanExit
    End If
    
    If colCount <> pipeCount Then
        ' количество колонок (
        ' column count (
        ' ) не соответствует pipecountcp (
        ' ) does not match pipecountcp (
        MsgBox Ru("043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 043A 043E 043B 043E 043D 043E 043A 0020 0028") & colCount & Ru("0029 0020 043D 0435 0020 0441 043E 043E 0442 0432 0435 0442 0441 0442 0432 0443 0435 0442 0020 0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0028") & pipeCount & ")!", vbExclamation
        GoTo CleanExit
    End If
    
    Dim pipeLength As Variant
    On Error Resume Next
    pipeLength = wsAnod.Range("pipeLength").Value
    On Error GoTo CleanExit
    
    If IsEmpty(pipeLength) Or Not IsNumeric(pipeLength) Or pipeLength <= 0 Then
        ' 'pipelength' должен быть положительным числом!
        ' 'pipelength' must be a positive number!
        MsgBox Ru("0027 0070 0069 0070 0065 004C 0065 006E 0067 0074 0068 0027 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043F 043E 043B 043E 0436 0438 0442 0435 043B 044C") & Ru("043D 044B 043C 0020 0447 0438 0441 043B 043E 043C 0021"), vbExclamation
        GoTo CleanExit
    End If
    
    ReDim arr(1 To 1, 1 To colCount)
    Dim lengthPerCP As Double
    lengthPerCP = pipeLength / pipeCount
    
    Application.StatusBar = "writing lengths for " & colCount & " CPs..."
    DoEvents
    
    For colCP = 1 To colCount
        arr(1, colCP) = lengthPerCP
    Next colCP
    
    wsAnod.Range("pipeLengthCP").Value = arr
    Application.Calculate
    
    If wsAnod.Range("C8").HasFormula Then
        wsAnod.Range("C8").Calculate
    End If
    
    Application.StatusBar = "length distributed: " & Format(lengthPerCP, "0.00") & " per CP"
    DoEvents
    
    ' укз равномерно распределены по длине трубопровода.
    ' укз evenly distributed along the pipeline.
    ' длина на одну укз 
    ' length per cp 
    '  м.
    ' перераспределите точки дренажа вручную при необходимости с учетом результатов изысканий
    ' redistribute drain points manually if needed using survey results
    ' распределение выполнено равномерно
    ' distribution completed evenly
    MsgBox Ru("0443 043A 0437 0020 0440 0430 0432 043D 043E 043C 0435 0440 043D 043E 0020 0440 0430 0441 043F 0440 0435 0434 0435 043B 0435 043D 044B 0020 043F 043E 0020 0434 043B 0438 043D") & Ru("0435 0020 0442 0440 0443 0431 043E 043F 0440 043E 0432 043E 0434 0430 002E") & vbCrLf & _
           Ru("0434 043B 0438 043D 0430 0020 043D 0430 0020 043E 0434 043D 0443 0020 0443 043A 0437 0020") & Format(lengthPerCP, "0.00") & Ru("0020 043C 002E") & vbCrLf & _
           Ru("043F 0435 0440 0435 0440 0430 0441 043F 0440 0435 0434 0435 043B 0438 0442 0435 0020 0442 043E 0447 043A 0438 0020 0434 0440 0435 043D 0430 0436 0430 0020 0432 0440 0443 0447") & Ru("043D 0443 044E 0020 043F 0440 0438 0020 043D 0435 043E 0431 0445 043E 0434 0438 043C 043E 0441 0442 0438 0020 0441 0020 0443 0447 0435 0442 043E 043C 0020 0440 0435 0437 0443") & Ru("043B 044C 0442 0430 0442 043E 0432 0020 0438 0437 044B 0441 043A 0430 043D 0438 0439"), _
           vbInformation, Ru("0440 0430 0441 043F 0440 0435 0434 0435 043B 0435 043D 0438 0435 0020 0432 044B 043F 043E 043B 043D 0435 043D 043E 0020 0440 0430 0432 043D 043E 043C 0435 0440 043D 043E")

CleanExit:
    Application.ScreenUpdating = originalScreenUpdating
    Application.EnableEvents = originalEvents
    Application.Calculation = originalCalc
    Application.StatusBar = False
End Sub

' ================================================================
' protective-zone calculation (main procedure)
' ================================================================
Public Sub CalcProtectiveZone(ByVal wsAnod As Worksheet)
    ' --- recursion guard ---
    Static inProgress As Boolean
    If inProgress Then
        ' >>> calcprotectivezone: уже выполняется, пропускаем
        ' >>> calcprotectivezone: already running, skipped
        Debug.Print Ru("003E 003E 003E 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E 0065 003A 0020 0443 0436 0435 0020 0432 044B 043F 043E 043B 043D 044F") & Ru("0435 0442 0441 044F 002C 0020 043F 0440 043E 043F 0443 0441 043A 0430 0435 043C")
        Exit Sub
    End If
    inProgress = True

    ' === save application state ===
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation

    ' === disable events and screen updating ===
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ' leave Calculation as is, or switch to manual if needed,
    ' but keep automatic here because the code triggers recalc itself.

    On Error GoTo CleanExit

    ' >>> calcprotectivezone: начало
    ' >>> calcprotectivezone: start
    Debug.Print Ru("003E 003E 003E 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E 0065 003A 0020 043D 0430 0447 0430 043B 043E")

    Application.StatusBar = "starting protective zone calculation..."
    DoEvents

    Dim wsPipe As Worksheet
    Dim colBr As Long
    Dim Utzm As Double, Utzo As Double

    If wsAnod Is Nothing Then
        ' ошибка: лист 
        ' error: лandст 
        '  не передан!
        '  not passed!
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 043B 0438 0441 0442 0020") & SHEET_ANOD & Ru("0020 043D 0435 0020 043F 0435 0440 0435 0434 0430 043D 0021"), vbCritical
        GoTo CleanExit
    End If

    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)

    If wsPipe Is Nothing Then
        ' лист '
        ' лandст '
        ' ' не найден!
        ' ' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        GoTo CleanExit
    End If

    ' >>> calcprotectivezone: листы найдены
    ' >>> calcprotectivezone: лandсты found
    Debug.Print Ru("003E 003E 003E 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E 0065 003A 0020 043B 0438 0441 0442 044B 0020 043D 0430 0439 0434 0435") & Ru("043D 044B")

    ' ================================================================
    ' 1. read only the required data from the pipe sheet
    ' ================================================================

    ' --- 1.1 read pipeInputResistanceEndLife (scalar) ---
    Dim pipeInputResistanceEndLife As Double
    On Error Resume Next
    pipeInputResistanceEndLife = wsPipe.Range("pipeInputResistanceEndLife").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать pipeinputresistanceendlife с листа pipe!
        ' failed to read pipeinputresistanceendlife с pipe sheet!
        ' убедитесь, что расчет трубопровода выполнен.
        ' make sure that pipe calculation is done.
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 0070 0069 0070 0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073") & Ru("0074 0061 006E 0063 0065 0045 006E 0064 004C 0069 0066 0065 0020 0441 0020 043B 0438 0441 0442 0430 0020 0070 0069 0070 0065 0021") & vbCrLf & _
               Ru("0443 0431 0435 0434 0438 0442 0435 0441 044C 002C 0020 0447 0442 043E 0020 0440 0430 0441 0447 0435 0442 0020 0442 0440 0443 0431 043E 043F 0440 043E 0432 043E 0434 0430 0020") & Ru("0432 044B 043F 043E 043B 043D 0435 043D 002E"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit  ' restore the error handler

    Debug.Print ">>> pipeInputResistanceEndLife = " & pipeInputResistanceEndLife

    ' --- 1.2 read factorPropagationCurrentAlongPipeEndLife (array) ---
    Dim alphaMax As Double
    Dim alphaVal As Variant
    Dim alphaRng As Range

    On Error Resume Next
    Set alphaRng = wsPipe.Range("factorPropagationCurrentAlongPipeEndLife")
    If Err.Number <> 0 Or alphaRng Is Nothing Then
        ' не удалось прочитать factorpropagationcurrentalongpipeendlife с листа pipe!
        ' failed to read factorpropagationcurrentalongpipeendlife с pipe sheet!
        ' убедитесь, что расчет трубопровода выполнен.
        ' make sure that pipe calculation is done.
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 0066 0061 0063 0074 006F 0072 0050 0072 006F 0070 0061 0067 0061 0074") & Ru("0069 006F 006E 0043 0075 0072 0072 0065 006E 0074 0041 006C 006F 006E 0067 0050 0069 0070 0065 0045 006E 0064 004C 0069 0066 0065 0020 0441 0020 043B 0438 0441 0442 0430 0020") & Ru("0070 0069 0070 0065 0021") & vbCrLf & _
               Ru("0443 0431 0435 0434 0438 0442 0435 0441 044C 002C 0020 0447 0442 043E 0020 0440 0430 0441 0447 0435 0442 0020 0442 0440 0443 0431 043E 043F 0440 043E 0432 043E 0434 0430 0020") & Ru("0432 044B 043F 043E 043B 043D 0435 043D 002E"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear

    alphaVal = alphaRng.Value
    If Err.Number <> 0 Then
        ' не удалось получить значение factorpropagationcurrentalongpipeendlife!
        ' failed to get value factorpropagationcurrentalongpipeendlife!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 043E 043B 0443 0447 0438 0442 044C 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020 0066 0061 0063 0074 006F 0072") & Ru("0050 0072 006F 0070 0061 0067 0061 0074 0069 006F 006E 0043 0075 0072 0072 0065 006E 0074 0041 006C 006F 006E 0067 0050 0069 0070 0065 0045 006E 0064 004C 0069 0066 0065 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit  ' restore the error handler

    If IsArray(alphaVal) Then
        alphaMax = -1E+308
        Dim hasValidAlpha As Boolean
        hasValidAlpha = False
        Dim r As Long, c As Long

        For r = LBound(alphaVal, 1) To UBound(alphaVal, 1)
            For c = LBound(alphaVal, 2) To UBound(alphaVal, 2)
                If IsNumeric(alphaVal(r, c)) Then
                    If alphaVal(r, c) > alphaMax Then
                        alphaMax = alphaVal(r, c)
                        hasValidAlpha = True
                    End If
                End If
            Next c
        Next r

        If Not hasValidAlpha Or alphaMax <= 0 Then
            ' нет корректных значений в factorpropagationcurrentalongpipeendlife!
            ' no valid values в factorpropagationcurrentalongpipeendlife!
            MsgBox Ru("043D 0435 0442 0020 043A 043E 0440 0440 0435 043A 0442 043D 044B 0445 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0432 0020 0066 0061 0063 0074 006F 0072 0050 0072 006F") & Ru("0070 0061 0067 0061 0074 0069 006F 006E 0043 0075 0072 0072 0065 006E 0074 0041 006C 006F 006E 0067 0050 0069 0070 0065 0045 006E 0064 004C 0069 0066 0065 0021"), vbExclamation
            GoTo CleanExit
        End If
    Else
        If IsNumeric(alphaVal) And alphaVal > 0 Then
            alphaMax = alphaVal
        Else
            ' нет корректных значений в factorpropagationcurrentalongpipeendlife!
            ' no valid values в factorpropagationcurrentalongpipeendlife!
            MsgBox Ru("043D 0435 0442 0020 043A 043E 0440 0440 0435 043A 0442 043D 044B 0445 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0432 0020 0066 0061 0063 0074 006F 0072 0050 0072 006F") & Ru("0070 0061 0067 0061 0074 0069 006F 006E 0043 0075 0072 0072 0065 006E 0074 0041 006C 006F 006E 0067 0050 0069 0070 0065 0045 006E 0064 004C 0069 0066 0065 0021"), vbExclamation
            GoTo CleanExit
        End If
    End If

    Debug.Print ">>> alphaMax = " & alphaMax

    ' ================================================================
    ' 2. read data from the anod sheet (scalars)
    ' ================================================================
    Dim minProtectPotential As Double
    Dim naturalPotential As Double
    Dim maxProtectPotential As Double
    Dim factorMutualInfluence As Double
    Dim pipeLength As Double

    On Error Resume Next
    minProtectPotential = wsAnod.Range("minProtectPotential").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать minprotectpotential!
        ' failed to read minprotectpotential!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 006D 0069 006E 0050 0072 006F 0074 0065 0063 0074 0050 006F 0074 0065") & Ru("006E 0074 0069 0061 006C 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit

    naturalPotential = wsAnod.Range("naturalPotential").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать naturalpotential!
        ' failed to read naturalpotential!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 006E 0061 0074 0075 0072 0061 006C 0050 006F 0074 0065 006E 0074 0069") & Ru("0061 006C 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit

    maxProtectPotential = wsAnod.Range("maxProtectPotential").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать maxprotectpotential!
        ' failed to read maxprotectpotential!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 006D 0061 0078 0050 0072 006F 0074 0065 0063 0074 0050 006F 0074 0065") & Ru("006E 0074 0069 0061 006C 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit

    factorMutualInfluence = wsAnod.Range("factorMutualInfluence").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать factormutualinfluence!
        ' failed to read factormutualinfluence!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 0066 0061 0063 0074 006F 0072 004D 0075 0074 0075 0061 006C 0049 006E") & Ru("0066 006C 0075 0065 006E 0063 0065 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit

    pipeLength = wsAnod.Range("pipeLength").Value
    If Err.Number <> 0 Then
        ' не удалось прочитать pipelength!
        ' failed to read pipelength!
        MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 043F 0440 043E 0447 0438 0442 0430 0442 044C 0020 0070 0069 0070 0065 004C 0065 006E 0067 0074 0068 0021"), vbExclamation
        On Error GoTo CleanExit
        GoTo CleanExit
    End If
    Err.Clear
    On Error GoTo CleanExit

    Debug.Print ">>> minProtectPotential = " & minProtectPotential
    Debug.Print ">>> naturalPotential = " & naturalPotential
    Debug.Print ">>> maxProtectPotential = " & maxProtectPotential
    Debug.Print ">>> factorMutualInfluence = " & factorMutualInfluence
    Debug.Print ">>> pipeLength = " & pipeLength

    ' --- check and apply default values ---
    If minProtectPotential = 0 Then minProtectPotential = -0.85
    If naturalPotential = 0 Then naturalPotential = -0.55
    If maxProtectPotential = 0 Then maxProtectPotential = -1.15
    If factorMutualInfluence = 0 Then factorMutualInfluence = 0.5
    If pipeLength = 0 Then pipeLength = 350000

    ' ================================================================
    ' 3. calculate potential shifts
    ' ================================================================
    Utzm = Abs(minProtectPotential) - Abs(naturalPotential)
    Utzo = Abs(maxProtectPotential) - Abs(naturalPotential)

    Debug.Print ">>> Utzm = " & Utzm
    Debug.Print ">>> Utzo = " & Utzo

    ' write shifts to cells (events already disabled globally)
    wsAnod.Range("pipeShiftPotentialMin").Value = Utzm
    wsAnod.Range("pipeShiftPotentialPoint").Value = Utzo

    ' validate the shifts
    If Utzm <= 0 Or Utzo <= 0 Then
        ' ошибка: смещения потенциала (utzm или utzo) равны нулю.
        ' error: potential shifts (utzm or utzo) are zero.
        ' проверьте значения minprotectpotential, maxprotectpotential и naturalpotential.
        ' проверьте зonченandя minprotectpotential, maxprotectpotential and naturalpotential.
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0441 043C 0435 0449 0435 043D 0438 044F 0020 043F 043E 0442 0435 043D 0446 0438 0430 043B 0430 0020 0028 0055 0074 007A 006D 0020 0438") & Ru("043B 0438 0020 0055 0074 007A 006F 0029 0020 0440 0430 0432 043D 044B 0020 043D 0443 043B 044E 002E") & vbCrLf & _
               Ru("043F 0440 043E 0432 0435 0440 044C 0442 0435 0020 0437 043D 0430 0447 0435 043D 0438 044F 0020 006D 0069 006E 0050 0072 006F 0074 0065 0063 0074 0050 006F 0074 0065 006E 0074") & Ru("0069 0061 006C 002C 0020 006D 0061 0078 0050 0072 006F 0074 0065 0063 0074 0050 006F 0074 0065 006E 0074 0069 0061 006C 0020 0438 0020 006E 0061 0074 0075 0072 0061 006C 0050") & Ru("006F 0074 0065 006E 0074 0069 0061 006C 002E"), vbExclamation
        GoTo CleanExit
    End If

    ' ================================================================
    ' 4. calculate Lz with formula (6.18)
    ' ================================================================
    Dim logArg As Double
    Dim denominator As Double

    denominator = factorMutualInfluence * Utzm
    If denominator <= 0 Then
        ' ошибка: denominator <= 0 (factormutualinfluence * utzm)
        ' error: denominator <= 0 (factormutualinfluence * utzm)
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 0064 0065 006E 006F 006D 0069 006E 0061 0074 006F 0072 0020 003C 003D 0020 0030 0020 0028 0066 0061 0063 0074 006F 0072 004D 0075 0074") & Ru("0075 0061 006C 0049 006E 0066 006C 0075 0065 006E 0063 0065 0020 002A 0020 0055 0074 007A 006D 0029"), vbExclamation
        GoTo CleanExit
    End If

    logArg = Utzo / denominator
    If logArg <= 0 Then
        ' ошибка: logarg <= 0 (utzo / (factormutualinfluence * utzm))
        ' error: logarg <= 0 (utzo / (factormutualinfluence * utzm))
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 006C 006F 0067 0041 0072 0067 0020 003C 003D 0020 0030 0020 0028 0055 0074 007A 006F 0020 002F 0020 0028 0066 0061 0063 0074 006F 0072") & Ru("004D 0075 0074 0075 0061 006C 0049 006E 0066 006C 0075 0065 006E 0063 0065 0020 002A 0020 0055 0074 007A 006D 0029 0029"), vbExclamation
        GoTo CleanExit
    End If

    Dim Lz As Double
    Lz = (2 / alphaMax) * Log(logArg)
    If Lz <= 0 Then
        ' расчетное lz <= 0, установлено в 0.
        ' calculated lz <= 0, set to 0.
        MsgBox Ru("0440 0430 0441 0447 0435 0442 043D 043E 0435 0020 004C 007A 0020 003C 003D 0020 0030 002C 0020 0443 0441 0442 0430 043D 043E 0432 043B 0435 043D 043E 0020 0432 0020 0030 002E"), vbExclamation
        Lz = 0
    End If

    ' ================================================================
    ' 5. clamp Lz
    ' ================================================================
    Dim originalLz As Double
    originalLz = Lz

    If Lz < Module_Constants.MIN_Lz Then
        Lz = Module_Constants.MIN_Lz
        ' lz ограничено снизу: 
        ' lz limited from below: 
        '  м
        Debug.Print Ru("004C 007A 0020 043E 0433 0440 0430 043D 0438 0447 0435 043D 043E 0020 0441 043D 0438 0437 0443 003A 0020") & originalLz & " -> " & Module_Constants.MIN_Lz & Ru("0020 043C")
    End If

    If Lz > Module_Constants.MAX_Lz Then
        Lz = Module_Constants.MAX_Lz
        ' lz ограничено сверху: 
        ' lz limited from above: 
        '  м
        Debug.Print Ru("004C 007A 0020 043E 0433 0440 0430 043D 0438 0447 0435 043D 043E 0020 0441 0432 0435 0440 0445 0443 003A 0020") & originalLz & " -> " & Module_Constants.MAX_Lz & Ru("0020 043C")
    End If

    If Lz > pipeLength Then
        Lz = pipeLength
        ' lz ограничено длиной трубопровода: 
        ' lz limited by pipeline length: 
        '  м
        Debug.Print Ru("004C 007A 0020 043E 0433 0440 0430 043D 0438 0447 0435 043D 043E 0020 0434 043B 0438 043D 043E 0439 0020 0442 0440 0443 0431 043E 043F 0440 043E 0432 043E 0434 0430 003A 0020") & originalLz & " -> " & pipeLength & Ru("0020 043C")
    End If

'    If Lz <> originalLz Then
'        MsgBox "protective zone Lz = " & Format(originalLz, "0.00") & " m" & vbCrLf & _
'               "limited to " & Format(Lz, "0.00") & " m", _
'               vbInformation, "Lz limit"
'    End If

    ' ================================================================
    ' 6. calculate pipeCount
    ' ================================================================
    Dim pipeCount As Long
    pipeCount = Application.RoundUp(pipeLength / Lz, 0)
    If pipeCount < 1 Then pipeCount = 1
    If pipeCount > Module_Constants.MAX_COL Then pipeCount = Module_Constants.MAX_COL

    Debug.Print ">>> pipeCount = " & pipeCount

    ' ================================================================
    ' 7. write Lz and pipeCountCP to cells (events disabled)
    ' ================================================================
    wsAnod.Range("lengthProtectiveZone").Value = Lz
    Application.EnableEvents = True
    wsAnod.Range("pipeCountCP").Value = pipeCount   ' D19, if that is the same cell
    DoEvents
' do NOT turn events off here, or Worksheet_Change may not run
' Application.EnableEvents = False
    colBr = pipeCount

    ' ================================================================
    ' 8. check: too many columns (> 1000)
    ' ================================================================
    If pipeCount > Module_Constants.MAX_COLUMNS_WARNING Then
        Dim response As VbMsgBoxResult
        ' внимание! будет создано 
        ' attention! will be created 
        '  колонок!
        '  columns!
        ' это более 
        ' this is more than 
        '  колонок.
        '  columns.
        ' операция может занять очень много времени.
        ' операцandя may take a very long time.
        ' продолжить?
        ' continue?
        ' слишком много колонок!
        ' too many columns!
        response = MsgBox(Ru("0432 043D 0438 043C 0430 043D 0438 0435 0021 0020 0431 0443 0434 0435 0442 0020 0441 043E 0437 0434 0430 043D 043E 0020") & pipeCount & Ru("0020 043A 043E 043B 043E 043D 043E 043A 0021") & vbCrLf & _
                         Ru("044D 0442 043E 0020 0431 043E 043B 0435 0435 0020") & Module_Constants.MAX_COLUMNS_WARNING & Ru("0020 043A 043E 043B 043E 043D 043E 043A 002E") & vbCrLf & vbCrLf & _
                         Ru("043E 043F 0435 0440 0430 0446 0438 044F 0020 043C 043E 0436 0435 0442 0020 0437 0430 043D 044F 0442 044C 0020 043E 0447 0435 043D 044C 0020 043C 043D 043E 0433 043E 0020 0432") & Ru("0440 0435 043C 0435 043D 0438 002E") & vbCrLf & _
                         Ru("043F 0440 043E 0434 043E 043B 0436 0438 0442 044C 003F"), _
                         vbYesNo + vbCritical + vbDefaultButton2, Ru("0441 043B 0438 0448 043A 043E 043C 0020 043C 043D 043E 0433 043E 0020 043A 043E 043B 043E 043D 043E 043A 0021"))

        If response = vbNo Then
            wsAnod.Range("lengthProtectiveZone").Value = originalLz
            wsAnod.Range("pipeCountCP").Value = 1
            ' операция отменена.
            ' operation cancelled.
            ' отмена
            ' cancel
            MsgBox Ru("043E 043F 0435 0440 0430 0446 0438 044F 0020 043E 0442 043C 0435 043D 0435 043D 0430 002E"), vbInformation, Ru("043E 0442 043C 0435 043D 0430")
            GoTo CleanExit
        End If
    End If

    ' ================================================================
    ' 11. recalculate formulas (Anod sheet only!)
    ' ================================================================
    Application.StatusBar = "recalculating formulas on Anod..."
    DoEvents

    Application.Calculate
    wsAnod.Calculate

    On Error Resume Next
    wsAnod.Range("currentCP").Calculate
    wsAnod.Range("currentEndLifeCP").Calculate
    wsAnod.Range("voltageEndLifeCP").Calculate
    wsAnod.Range("powerEndLifeCP").Calculate
    On Error GoTo CleanExit   ' restore the error handler

    ' ================================================================
    ' 12. calculate CP-unit currents
    ' ================================================================
    Application.StatusBar = "calculating CP currents..."
    DoEvents

    On Error Resume Next
    Dim Rstart As Double
    Dim Rend As Double
    Dim Nukz As Double
    Dim i As Double
    Dim it As Double

    Rstart = wsPipe.Range("pipeInputResistance").Value
    Rend = wsPipe.Range("pipeInputResistanceEndLife").Value
    Nukz = wsAnod.Range("pipeCountCP").Value

    If IsNumeric(Utzo) And IsNumeric(Rstart) And IsNumeric(Rend) And IsNumeric(Nukz) Then
        If Rstart > 0 And Rend > 0 And Nukz > 0 Then
            i = 2 * Utzo / Rstart / Nukz
            it = 2 * Utzo / Rend / Nukz

            wsAnod.Range("currentCP").Value = i
            wsAnod.Range("currentEndLifeCP").Value = it
        End If
    End If
    On Error GoTo CleanExit   ' restore the error handler

    ' ================================================================
    ' 13. final recalc (Anod sheet only!)
    ' ================================================================
    Application.StatusBar = "final recalculation..."
    DoEvents
    wsAnod.Calculate

    Application.StatusBar = "protective zone calculation finished"
    DoEvents
    ' >>> calcprotectivezone: завершено успешно
    ' >>> calcprotectivezone: completed successfully
    Debug.Print Ru("003E 003E 003E 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E 0065 003A 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E 0020 0443") & Ru("0441 043F 0435 0448 043D 043E")

CleanExit:
    ' Always clear the flag
    inProgress = False
    
    ' === restore original application state ===
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    If Err.Number <> 0 Then
        ' >>> ошибка calcprotectivezone: 
        ' >>> error calcprotectivezone: 
        Debug.Print Ru("003E 003E 003E 0020 043E 0448 0438 0431 043A 0430 0020 0043 0061 006C 0063 0050 0072 006F 0074 0065 0063 0074 0069 0076 0065 005A 006F 006E 0065 003A 0020") & Err.Number & " - " & Err.Description
        ' ошибка: 
        ' error: 
        ' код: 
        ' code: 
        ' ошибка макроса
        ' macro error
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description & vbCrLf & Ru("043A 043E 0434 003A 0020") & Err.Number, vbExclamation, Ru("043E 0448 0438 0431 043A 0430 0020 043C 0430 043A 0440 043E 0441 0430")
    End If
End Sub

' ================================================================
' calculate maximum protective potential
' ================================================================
Public Function CalcMaxProtectPotential() As Double
    On Error GoTo ErrorHandler
    
    Dim wsPipe As Worksheet
    Dim rngAlpha As Range
    Dim alphaMax As Double
    Dim naturalPot As Double
    Dim factorMutual As Double
    Dim minProtectPot As Double
    Dim pipeLen As Double
    Dim pipeCount As Double
    Dim result As Double
    Dim currentMaxPot As Double
    Dim limitedResult As Double
    
    ' --- read values ---
    On Error Resume Next
    naturalPot = Range("naturalPotential").Value
    factorMutual = Range("factorMutualInfluence").Value
    minProtectPot = Range("minProtectPotential").Value
    pipeLen = Range("pipeLength").Value
    pipeCount = Range("pipeCountCP").Value
    currentMaxPot = Range("maxProtectPotential").Value
    On Error GoTo 0
    
    ' --- default values ---
    If Not IsNumeric(naturalPot) Or naturalPot = 0 Then
        naturalPot = -0.55
    End If
    If Not IsNumeric(minProtectPot) Or minProtectPot = 0 Then
        minProtectPot = -0.85
    End If
    If Not IsNumeric(factorMutual) Or factorMutual = 0 Then
        factorMutual = 0.5
    End If
    If Not IsNumeric(pipeLen) Or pipeLen <= 0 Then
        pipeLen = 350000
    End If
    If Not IsNumeric(pipeCount) Or pipeCount <= 0 Then
        pipeCount = 1
    End If
    
    ' --- alphaMax ---
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If Not wsPipe Is Nothing Then
        On Error Resume Next
        Set rngAlpha = wsPipe.Range("factorPropagationCurrentAlongPipeEndLife")
        If Not rngAlpha Is Nothing Then
            On Error GoTo 0
            alphaMax = Application.WorksheetFunction.Max(rngAlpha)
        End If
    End If
    
    If Not IsNumeric(alphaMax) Or alphaMax <= 0 Then
        alphaMax = 0.0001
    End If
    
    ' --- calculation ---
    result = naturalPot - factorMutual * (naturalPot - minProtectPot) * Exp(alphaMax * pipeLen / (2 * pipeCount))
    
    If Not IsNumeric(result) Or Abs(result) > 1E+30 Then
        CalcMaxProtectPotential = currentMaxPot
        Exit Function
    End If
    
    ' --- clamp ---
    limitedResult = result
    If limitedResult < Module_Constants.MIN_PROTECT_POTENTIAL Then
        limitedResult = Module_Constants.MIN_PROTECT_POTENTIAL
        ' ограничение: вычисленное значение 
        ' limit: calculated value 
        '  заменено на 
        '  replaced with 
        Debug.Print Ru("043E 0433 0440 0430 043D 0438 0447 0435 043D 0438 0435 003A 0020 0432 044B 0447 0438 0441 043B 0435 043D 043D 043E 0435 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020") & result & Ru("0020 0437 0430 043C 0435 043D 0435 043D 043E 0020 043D 0430 0020") & Module_Constants.MIN_PROTECT_POTENTIAL
    ElseIf limitedResult > Module_Constants.MAX_PROTECT_POTENTIAL Then
        limitedResult = Module_Constants.MAX_PROTECT_POTENTIAL
        ' ограничение: вычисленное значение 
        ' limit: calculated value 
        '  заменено на 
        '  replaced with 
        Debug.Print Ru("043E 0433 0440 0430 043D 0438 0447 0435 043D 0438 0435 003A 0020 0432 044B 0447 0438 0441 043B 0435 043D 043D 043E 0435 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020") & result & Ru("0020 0437 0430 043C 0435 043D 0435 043D 043E 0020 043D 0430 0020") & Module_Constants.MAX_PROTECT_POTENTIAL
    End If
    
    CalcMaxProtectPotential = limitedResult
    Exit Function

ErrorHandler:
    CalcMaxProtectPotential = currentMaxPot
End Function
