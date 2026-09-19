Attribute VB_Name = "Module_calcPipe"

' ================================================================
' module: Module_calcPipe
' purpose: pipeline parameter calculation
' ================================================================
Option Explicit

' ================================================================
' helper to read arrays
' ================================================================
Private Function EnsureArray2(ByVal rangeName As String, ByVal cols As Long) As Variant
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim vals As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim wsPipe As Worksheet
    
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)

    
    ' check that the sheet exists
    If wsPipe Is Nothing Then
        ReDim arr(1 To 1, 1 To cols)
        EnsureArray2 = arr
        Exit Function
    End If

    On Error Resume Next
    Set rng = wsPipe.Range(rangeName)
    If rng Is Nothing Then
        Set rng = thisWorkbook.names(rangeName).RefersToRange
    End If
    On Error GoTo 0

    If rng Is Nothing Then
        ReDim arr(1 To 1, 1 To cols)
        EnsureArray2 = arr
        Exit Function
    End If

    vals = rng.Value

    If Not IsArray(vals) Then
        ReDim arr(1 To 1, 1 To cols)
        If Not IsError(vals) And Not IsEmpty(vals) Then
            arr(1, 1) = vals
        End If
        EnsureArray2 = arr
        Exit Function
    End If

    Dim rows As Long, currentCols As Long
    On Error Resume Next
    rows = UBound(vals, 1)
    currentCols = UBound(vals, 2)
    On Error GoTo 0

    If rows = 0 Or currentCols = 0 Then
        ReDim arr(1 To 1, 1 To cols)
        If IsArray(vals) Then
            Dim v As Variant
            Dim idx As Long
            idx = 1
            For Each v In vals
                If idx <= cols Then
                    arr(1, idx) = v
                    idx = idx + 1
                Else
                    Exit For
                End If
            Next v
        End If
        EnsureArray2 = arr
        Exit Function
    End If

    ReDim arr(1 To 1, 1 To cols)
    For i = 1 To Application.Min(currentCols, cols)
        arr(1, i) = vals(1, i)
    Next i
    EnsureArray2 = arr
    Exit Function

ErrorHandler:
    ReDim arr(1 To 1, 1 To cols)
    EnsureArray2 = arr
End Function

' ================================================================
' calculate pipeline parameters
' ================================================================
Public Sub btnPipeCalculate()
    On Error GoTo CleanExit
    
    ' guard against re-entry
    Static inProgress As Boolean
    If inProgress Then
        ' расчет уже выполняется. пожалуйста, подождите.
        ' calculation is already running. please wait.
        MsgBox Ru("0440 0430 0441 0447 0435 0442 0020 0443 0436 0435 0020 0432 044B 043F 043E 043B 043D 044F 0435 0442 0441 044F 002E 0020 043F 043E 0436 0430 043B 0443 0439 0441 0442 0430 002C") & Ru("0020 043F 043E 0434 043E 0436 0434 0438 0442 0435 002E"), vbExclamation
        Exit Sub
    End If
    inProgress = True

    ' save application settings
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant
    
    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    ' optimization
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.StatusBar = "starting calculation..."

    ' get a reference to SHEET_PIPE
    Dim wsPipe As Worksheet
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If wsPipe Is Nothing Then
        ' лист '
        ' лandст '
        ' ' не найден!
        ' ' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbCritical
        GoTo CleanExit
    End If

    ' determine the number of sections
    Dim colBr As Long
    Dim pipeCountValue As Variant
    Dim startCol As Long
    startCol = 4
    
    On Error Resume Next
    pipeCountValue = wsPipe.Range("pipeDifferentParametersNum").Value
    If Err.Number <> 0 Then
        pipeCountValue = 1
        Err.Clear
    End If
    On Error GoTo 0
    
    If IsNumeric(pipeCountValue) Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If
    
    If colBr < 1 Then colBr = 1
    
    ' ================================================================
    ' clear calculated values before writing
    ' ================================================================
    Dim clearStartRow As Long, clearEndRow As Long
    clearStartRow = 17
    clearEndRow = 29
    
    Dim clearRange As Range
    Set clearRange = wsPipe.Range(wsPipe.Cells(clearStartRow, startCol), _
                                   wsPipe.Cells(clearEndRow, startCol + colBr - 1))
    clearRange.ClearContents
    ' очищены строки 
    ' cleared rows 
    '  для 
    '  for 
    '  колонок
    '  columns
    Debug.Print Ru("043E 0447 0438 0449 0435 043D 044B 0020 0441 0442 0440 043E 043A 0438 0020") & clearStartRow & "-" & clearEndRow & Ru("0020 0434 043B 044F 0020") & colBr & Ru("0020 043A 043E 043B 043E 043D 043E 043A")

    ' warn when there are many sections
    If colBr > 100 Then
        ' количество плеч: 
        ' number of sections: 
        ' . это может занять много времени. продолжить?
        ' . это may take a long time. continue?
        ' подтверждение
        ' confirmation
        If MsgBox(Ru("043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 043F 043B 0435 0447 003A 0020") & colBr & Ru("002E 0020 044D 0442 043E 0020 043C 043E 0436 0435 0442 0020 0437 0430 043D 044F 0442 044C 0020 043C 043D 043E 0433 043E 0020 0432 0440 0435 043C 0435 043D 0438 002E 0020 043F") & Ru("0440 043E 0434 043E 043B 0436 0438 0442 044C 003F"), _
                  vbYesNo + vbExclamation, Ru("043F 043E 0434 0442 0432 0435 0440 0436 0434 0435 043D 0438 0435")) = vbNo Then
            GoTo CleanExit
        End If
    End If

    Application.StatusBar = "pipe calc for " & colBr & " sections... loading data 10%"

    ' load data from named ranges
    Dim arrSteelResist As Variant, arrDiameter As Variant, arrWallThick As Variant
    Dim arrSoilResistAvg As Variant
    Dim arrLayingDepth As Variant, arrInsulResistStart As Variant
    Dim arrChangeFactor As Variant, arrServiceLife As Variant

    arrSteelResist = EnsureArray2("pipeSteelResistivity", colBr)  ' <-- fixed
    arrDiameter = EnsureArray2("pipeDiameter", colBr)
    arrWallThick = EnsureArray2("pipeWallThickness", colBr)
    arrSoilResistAvg = EnsureArray2("soilResistivityAvg", colBr)
    arrLayingDepth = EnsureArray2("pipeLayingDepth", colBr)
    arrInsulResistStart = EnsureArray2("pipeInsulationResistivityStartLife", colBr)
    arrChangeFactor = EnsureArray2("pipeResistivityChangeFactor", colBr)
    arrServiceLife = EnsureArray2("serviceLifeDesigned", colBr)

    ' ================================================================
    ' check that section 1 has data
    ' ================================================================
    Dim missingMsg As String
    Dim paramArrays As Variant
    Dim paramLabels As Variant
    Dim p As Long
    
    ' input data only (do not clear calculated values!)
    paramArrays = Array(arrSteelResist, arrDiameter, arrWallThick, arrSoilResistAvg, _
                        arrLayingDepth, arrInsulResistStart, _
                        arrChangeFactor, arrServiceLife)
    ' удельное сопротивление стали
    ' steel resistivity
    ' диаметр трубы
    ' pipe diameter
    ' толщина стенки
    ' wall thickness
    ' среднее удельное сопротивление грунта
    ' average soil resistivity
    ' глубина заложения
    ' laying depth
    ' начальное сопротивление изоляции
    ' initial insulation resistance
    ' коэффициент изменения
    ' change factor
    ' срок службы
    ' service life
    paramLabels = Array(Ru("0443 0434 0435 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0441 0442 0430 043B 0438"), Ru("0434 0438 0430 043C 0435 0442 0440 0020 0442 0440 0443 0431 044B"), Ru("0442 043E 043B 0449 0438 043D 0430 0020 0441 0442 0435 043D 043A 0438"), _
                        Ru("0441 0440 0435 0434 043D 0435 0435 0020 0443 0434 0435 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0433 0440 0443 043D") & Ru("0442 0430"), _
                        Ru("0433 043B 0443 0431 0438 043D 0430 0020 0437 0430 043B 043E 0436 0435 043D 0438 044F"), Ru("043D 0430 0447 0430 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0438 0437 043E 043B 044F 0446 0438 0438"), _
                        Ru("043A 043E 044D 0444 0444 0438 0446 0438 0435 043D 0442 0020 0438 0437 043C 0435 043D 0435 043D 0438 044F"), Ru("0441 0440 043E 043A 0020 0441 043B 0443 0436 0431 044B"))
    
    missingMsg = ""
    For p = LBound(paramArrays) To UBound(paramArrays)
        On Error Resume Next
        Dim val As Variant
        val = paramArrays(p)(1, 1)
        If Err.Number <> 0 Or IsEmpty(val) Or Not IsNumeric(val) Or val <= 0 Then
            missingMsg = missingMsg & " - " & paramLabels(p) & vbCrLf
        End If
        Err.Clear
    Next p
    On Error GoTo 0
    
    If missingMsg <> "" Then
        ' введите исходные параметры в следующих диапазонах:
        ' enter source parameters in these ranges:
        ' для первой колонки (плечо 1) обязательно заполните все значения.
        ' for the first column (section 1) fill in all values.
        ' недостающие данные
        ' missing data
        MsgBox Ru("0432 0432 0435 0434 0438 0442 0435 0020 0438 0441 0445 043E 0434 043D 044B 0435 0020 043F 0430 0440 0430 043C 0435 0442 0440 044B 0020 0432 0020 0441 043B 0435 0434 0443 044E") & Ru("0449 0438 0445 0020 0434 0438 0430 043F 0430 0437 043E 043D 0430 0445 003A") & vbCrLf & vbCrLf & _
               missingMsg & vbCrLf & _
               Ru("0434 043B 044F 0020 043F 0435 0440 0432 043E 0439 0020 043A 043E 043B 043E 043D 043A 0438 0020 0028 043F 043B 0435 0447 043E 0020 0031 0029 0020 043E 0431 044F 0437 0430 0442") & Ru("0435 043B 044C 043D 043E 0020 0437 0430 043F 043E 043B 043D 0438 0442 0435 0020 0432 0441 0435 0020 0437 043D 0430 0447 0435 043D 0438 044F 002E"), _
               vbExclamation, Ru("043D 0435 0434 043E 0441 0442 0430 044E 0449 0438 0435 0020 0434 0430 043D 043D 044B 0435")
        GoTo CleanExit
    End If

    Application.StatusBar = "pipe calc for " & colBr & " sections... computing 30%"

    ' main calculations
    Dim pi As Double
    pi = Application.pi()

    ' result arrays
    Dim arrAlongResist() As Double
    Dim arrRp() As Double
    Dim arrRper() As Double
    Dim arrRp_length() As Double
    Dim arrRins0_length() As Double
    Dim arrRp0() As Double
    Dim arrRpt() As Double
    Dim arrAlpha() As Double
    Dim arrAlphaEnd() As Double
    Dim arrZ() As Double
    Dim arrZt() As Double

    ReDim arrAlongResist(1 To colBr)
    ReDim arrRp(1 To colBr)
    ReDim arrRper(1 To colBr)
    ReDim arrRp_length(1 To colBr)
    ReDim arrRins0_length(1 To colBr)
    ReDim arrRp0(1 To colBr)
    ReDim arrRpt(1 To colBr)
    ReDim arrAlpha(1 To colBr)
    ReDim arrAlphaEnd(1 To colBr)
    ReDim arrZ(1 To colBr)
    ReDim arrZt(1 To colBr)

    ' calculation variables
    Dim rho_steel As Double, D As Double, thick As Double
    Dim rho_avg As Double, rho_around As Double, h As Double
    Dim R_ins_start As Double, changeFactor As Double, serviceLife As Double
    Dim Rt As Double, Rp As Double, Rper As Double
    Dim Rp_length As Double, Rins0_length As Double
    Dim Rp0 As Double, Rpt As Double, alpha As Double, alphaEnd As Double
    Dim Z As Double, Zt As Double
    Dim denom As Double, ln_arg As Double
    Dim denominator As Double

    Dim i As Long
    
    ' main calculation loop
    For i = 1 To colBr
        ' read data for the current section
        On Error Resume Next
        rho_steel = CDbl(arrSteelResist(1, i))
        D = CDbl(arrDiameter(1, i))
        thick = CDbl(arrWallThick(1, i))
        rho_avg = CDbl(arrSoilResistAvg(1, i))
        ' ? rho_around is calculated; use default 500
        rho_around = 500 'CDbl(soilResistanceAroundPipe(1, i))
        h = CDbl(arrLayingDepth(1, i))
        R_ins_start = CDbl(arrInsulResistStart(1, i))
        changeFactor = CDbl(arrChangeFactor(1, i))
        serviceLife = CDbl(arrServiceLife(1, i))
        If Err.Number <> 0 Then
            Err.Clear
            GoTo SkipCalculation
        End If
        On Error GoTo 0

        ' ================================================================
        ' validate inputs (guard against division by zero)
        ' ================================================================
        If D <= 0 Or thick <= 0 Or h <= 0 Or rho_steel <= 0 Or rho_avg <= 0 Or rho_around <= 0 Then
            ' плечо 
            ' section 
            ' : некорректные входные данные
            ' : invalid input data
            Debug.Print Ru("043F 043B 0435 0447 043E 0020") & i & Ru("003A 0020 043D 0435 043A 043E 0440 0440 0435 043A 0442 043D 044B 0435 0020 0432 0445 043E 0434 043D 044B 0435 0020 0434 0430 043D 043D 044B 0435")
            GoTo SkipCalculation
        End If

        ' ================================================================
        ' calculate pipe resistance
        ' ================================================================
        denom = pi * (D - thick) * thick
        If denom = 0 Then
            Rt = 2.43E-06
        Else
            Rt = rho_steel / denom
        End If
        If Rt <= 0 Then Rt = 2.43E-06
        arrAlongResist(i) = Rt

' ================================================================
' iteratively calculate spreading resistance (Rp)
' ================================================================
' initial guess
Rp = 500
Dim Rp_prev As Double
Dim iter As Long
Const MAX_ITER_RP As Long = 10
Const TOL_RP As Double = 0.001 ' absolute convergence tolerance

If D <= 0 Or h <= 0 Or Rt <= 0 Or rho_avg <= 0 Then
    Rp = 500 ' guard against invalid data
Else
    denominator = D ^ 2 * h * Rt
    If denominator > 0 Then
        For iter = 1 To MAX_ITER_RP
            Rp_prev = Rp
            ln_arg = 0.4 * Rp_prev / denominator
            If ln_arg <= 0 Then
                Rp = 500
                Exit For
            End If
            Rp = rho_avg * D / 2 * Log(ln_arg)
            If Rp <= 0 Then
                Rp = 500
                Exit For
            End If
            ' check convergence
            If Abs(Rp - Rp_prev) < TOL_RP Then Exit For
        Next iter
    Else
        Rp = 500
    End If
End If
arrRp(i) = Rp

' ================================================================
' transition resistance (formula 6.3)
' ================================================================
Rper = Rp + R_ins_start
If Rper <= 0 Then Rper = 100
arrRper(i) = Rper

        ' ================================================================
        ' specific resistances
        ' ================================================================
        If D > 0 Then
            Rp_length = rho_around / (pi * D)
        Else
            Rp_length = 0
        End If
        arrRp_length(i) = Rp_length

        If D > 0 Then
            Rins0_length = R_ins_start / (pi * D)
        Else
            Rins0_length = 0
        End If
        arrRins0_length(i) = Rins0_length

        ' ================================================================
        ' total transition resistance
        ' ================================================================
        Rp0 = Rp_length + Rins0_length
        If Rp0 <= 0 Then Rp0 = 0.0001
        arrRp0(i) = Rp0

        ' ================================================================
        ' end-of-life transition resistance
        ' ================================================================
        If changeFactor <= 0 Then changeFactor = 0.001
        Rpt = Rp_length + Rins0_length * Exp(-changeFactor * serviceLife)
        If Rpt <= 0 Then Rpt = 0.0001
        arrRpt(i) = Rpt

        ' ================================================================
        ' attenuation coefficients
        ' ================================================================
        If Rt > 0 And Rp0 > 0 Then
            alpha = Sqr(Rt / Rp0)
        Else
            alpha = 0.0001
        End If
        arrAlpha(i) = alpha

        If Rt > 0 And Rpt > 0 Then
            alphaEnd = Sqr(Rt / Rpt)
        Else
            alphaEnd = 0.0001
        End If
        arrAlphaEnd(i) = alphaEnd

        ' ================================================================
        ' characteristic impedances
        ' ================================================================
        If Rt > 0 And Rp0 > 0 Then
            Z = Sqr(Rt * Rp0)
        Else
            Z = 0.0001
        End If
        arrZ(i) = Z

        If Rt > 0 And Rpt > 0 Then
            Zt = Sqr(Rt * Rpt)
        Else
            Zt = 0.0001
        End If
        arrZt(i) = Zt
        
        GoTo NextIteration
        
SkipCalculation:
        ' fill defaults on error
        arrAlongResist(i) = 2.43E-06
        arrRp(i) = 500
        arrRper(i) = 100
        arrRp_length(i) = 0
        arrRins0_length(i) = 0
        arrRp0(i) = 0.0001
        arrRpt(i) = 0.0001
        arrAlpha(i) = 0.0001
        arrAlphaEnd(i) = 0.0001
        
        arrZ(i) = 0.0001
        arrZt(i) = 0.0001
    

NextIteration:
    Next i

    ' ================================================================
    ' write results
    ' ================================================================
    Application.StatusBar = "pipe calc for " & colBr & " sections... writing results 80%"
    DoEvents

    On Error Resume Next
    wsPipe.Range("pipeAlongResistance").Resize(1, colBr).Value = arrAlongResist
    wsPipe.Range("soilResistanceAroundPipe").Resize(1, colBr).Value = arrRp
    wsPipe.Range("pipeTransientResistivity").Resize(1, colBr).Value = arrRper
    wsPipe.Range("soilResistivityAroundPipe").Resize(1, colBr).Value = arrRp_length
    wsPipe.Range("pipeInsulationResistanceStartLife").Resize(1, colBr).Value = arrRins0_length
    wsPipe.Range("pipeTransientResistance").Resize(1, colBr).Value = arrRp0
    wsPipe.Range("pipeTransientResistanceEndLife").Resize(1, colBr).Value = arrRpt
    wsPipe.Range("factorPropagationCurrentAlongPipe").Resize(1, colBr).Value = arrAlpha
    wsPipe.Range("factorPropagationCurrentAlongPipeEndLife").Resize(1, colBr).Value = arrAlphaEnd
    wsPipe.Range("pipeImpedance").Resize(1, colBr).Value = arrZ
    wsPipe.Range("pipeImpedanceEndLife").Resize(1, colBr).Value = arrZt
    On Error GoTo 0

    ' ================================================================
    ' calculate input resistances
    ' ================================================================
    Application.StatusBar = "pipe calc for " & colBr & " sections... resistances 90%"
    DoEvents

    Dim sumCondu As Double, sumConduEnd As Double
    sumCondu = 0#: sumConduEnd = 0#
    
    For i = 1 To colBr
        If arrZ(i) <> 0 Then
            On Error Resume Next
            sumCondu = sumCondu + 1# / arrZ(i)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
        If arrZt(i) <> 0 Then
            On Error Resume Next
            sumConduEnd = sumConduEnd + 1# / arrZt(i)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Next i

    On Error Resume Next
    If sumCondu <> 0 Then
        wsPipe.Range("pipeInputResistance").Value2 = 1# / sumCondu
    Else
        ' нет данных
        ' no data
        wsPipe.Range("pipeInputResistance").Value2 = Ru("043D 0435 0442 0020 0434 0430 043D 043D 044B 0445")
    End If
    
    If sumConduEnd <> 0 Then
        wsPipe.Range("pipeInputResistanceEndLife").Value2 = 1# / sumConduEnd
    Else
        ' нет данных
        ' no data
        wsPipe.Range("pipeInputResistanceEndLife").Value2 = Ru("043D 0435 0442 0020 0434 0430 043D 043D 044B 0445")
    End If
    On Error GoTo 0

    ' ================================================================
    ' recalculate formulas
    ' ================================================================
    Application.StatusBar = "pipe calc for " & colBr & " sections... recalc formulas 95%"
    DoEvents

    Application.Calculation = xlCalculationAutomatic
    wsPipe.Calculate
    
    ' recalculate the Anod calculation sheet
    On Error Resume Next
    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets(SHEET_ANOD)
    If Not wsAnod Is Nothing Then
        wsAnod.Calculate
    End If
    On Error GoTo 0

    Application.StatusBar = "done: pipe calc finished for " & colBr & " sections"
    
    ' ================================================================
    ' info message
    ' ================================================================
  '  MsgBox "calculation finished" & vbCrLf & _
  '         "sections: " & colBr & vbCrLf & _
  '         "input resistance: " & wsPipe.Range("pipeInputResistanceEndLife").value, _
  '         vbInformation, "done"

CleanExit:
    ' restore settings
    Application.Cursor = xlDefault
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = False
    
    If Err.Number <> 0 Then
        ' ошибка: 
        ' error: 
        ' код: 
        ' code: 
        ' ошибка макроса
        ' macro error
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description & vbCrLf & Ru("043A 043E 0434 003A 0020") & Err.Number, vbExclamation, Ru("043E 0448 0438 0431 043A 0430 0020 043C 0430 043A 0440 043E 0441 0430")
        Err.Clear
    End If
    
    inProgress = False
End Sub
