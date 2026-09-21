Attribute VB_Name = "Module_calcAnod"

' ================================================================
' module: Module_calcAnod
' purpose: anode-groundbed calculations (AG)
' ================================================================
Option Explicit

Public Sub btnAnodFullCalc()
    On Error GoTo CleanExit

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim appStateSaved As Boolean
    appStateSaved = False
    
    ' === btnanodfullcalc: начало ===
    ' === btnanodfullcalc: start ===
    If DEBUG_MODE Then Debug.Print Ru("003D 003D 003D 0020 0062 0074 006E 0041 006E 006F 0064 0046 0075 006C 006C 0043 0061 006C 0063 003A 0020 043D 0430 0447 0430 043B 043E 0020 003D 003D 003D")
    
    ' --- refresh Anod sheet named ranges before calculation ---
    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets("Anod")
    Dim colBr As Long
    Dim pipeVal As Variant
    pipeVal = wsAnod.Range("pipeCountCP").Value
    colBr = IIf(IsNumeric(pipeVal) And pipeVal > 0, CLng(pipeVal), 1)
    If DEBUG_MODE Then Debug.Print "  pipeCountCP = " & colBr
    
    ' ================================================================
    ' restore ranges directly (without calling UpdateAnodNamedRanges)
    ' ================================================================
    '   выполняем прямое восстановление диапазонов...
    '   restoring ranges directly...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0432 044B 043F 043E 043B 043D 044F 0435 043C 0020 043F 0440 044F 043C 043E 0435 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 0435 0020 0434") & Ru("0438 0430 043F 0430 0437 043E 043D 043E 0432 002E 002E 002E")
    Dim endColManual As Long
    endColManual = 4 + colBr - 1
    
    On Error Resume Next
    wsAnod.names("oneElectrodeResistanceAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, 4), wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, endColManual))
    wsAnod.names("oneElectrodeResistanceHorizAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, 4), wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, endColManual))
    wsAnod.names("numElectrodesAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_NUM_ELECTRODES_AG, 4), wsAnod.Cells(ROW_NUM_ELECTRODES_AG, endColManual))
    wsAnod.names("weightWithoutFillingAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, 4), wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, endColManual))
    wsAnod.names("serviceLifeAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_SERVICE_LIFE_AG, 4), wsAnod.Cells(ROW_SERVICE_LIFE_AG, endColManual))
    wsAnod.names("serviceLifeDeviation").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_SERVICE_LIFE_DEVIATION, 4), wsAnod.Cells(ROW_SERVICE_LIFE_DEVIATION, endColManual))
    wsAnod.names("correctResistanceAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_CORRECT_RESISTANCE_AG, 4), wsAnod.Cells(ROW_CORRECT_RESISTANCE_AG, endColManual))
    
    If Err.Number = 0 Then
        '   диапазоны восстановлены: 1x
        '   ranges restored: 1x
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0434 0438 0430 043F 0430 0437 043E 043D 044B 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 044B 003A 0020 0031 0078") & colBr
    Else
        '   !!! ошибка восстановления: 
        '   !!! restore error: 
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043E 0448 0438 0431 043A 0430 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 044F 003A 0020") & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' --- check the size of oneElectrodeResistanceAG ---
    Dim rngTest As Range
    On Error Resume Next
    Set rngTest = wsAnod.Range("oneElectrodeResistanceAG")
    If Err.Number <> 0 Then
        '   !!! oneelectroderesistanceag не существует!
        '   !!! oneelectroderesistanceag does not exist!
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 006F 006E 0065 0045 006C 0065 0063 0074 0072 006F 0064 0065 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 043D 0435 0020 0441") & Ru("0443 0449 0435 0441 0442 0432 0443 0435 0442 0021")
        Err.Clear
        GoTo CleanExit
    End If
    On Error GoTo 0
    
    If DEBUG_MODE Then Debug.Print "  oneElectrodeResistanceAG: rows=" & rngTest.rows.count & ", cols=" & rngTest.Columns.count
    
    If rngTest.rows.count <> 1 Then
        '   !!! oneelectroderesistanceag имеет 
        '   !!! oneelectroderesistanceag has
        '  строк, должно быть 1!
        '  rows, must be 1!
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 006F 006E 0065 0045 006C 0065 0063 0074 0072 006F 0064 0065 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0438 043C 0435 0435") & Ru("0442 0020") & rngTest.rows.count & Ru("0020 0441 0442 0440 043E 043A 002C 0020 0434 043E 043B 0436 043D 043E 0020 0431 044B 0442 044C 0020 0031 0021")
        Dim endColForce As Long
        endColForce = 4 + colBr - 1
        On Error Resume Next
        wsAnod.names("oneElectrodeResistanceAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, 4), wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_AG, endColForce))
        wsAnod.names("oneElectrodeResistanceHorizAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, 4), wsAnod.Cells(ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, endColForce))
        wsAnod.names("numElectrodesAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_NUM_ELECTRODES_AG, 4), wsAnod.Cells(ROW_NUM_ELECTRODES_AG, endColForce))
        wsAnod.names("weightWithoutFillingAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, 4), wsAnod.Cells(ROW_WEIGHT_WITHOUT_FILLING_AG, endColForce))
        wsAnod.names("serviceLifeAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_SERVICE_LIFE_AG, 4), wsAnod.Cells(ROW_SERVICE_LIFE_AG, endColForce))
        wsAnod.names("serviceLifeDeviation").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_SERVICE_LIFE_DEVIATION, 4), wsAnod.Cells(ROW_SERVICE_LIFE_DEVIATION, endColForce))
        wsAnod.names("correctResistanceAG").RefersTo = wsAnod.Range(wsAnod.Cells(ROW_CORRECT_RESISTANCE_AG, 4), wsAnod.Cells(ROW_CORRECT_RESISTANCE_AG, endColForce))
        If Err.Number = 0 Then
            '   диапазоны принудительно восстановлены
            '   rangeы force-restored
            If DEBUG_MODE Then Debug.Print Ru("0020 0020 0434 0438 0430 043F 0430 0437 043E 043D 044B 0020 043F 0440 0438 043D 0443 0434 0438 0442 0435 043B 044C 043D 043E 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432") & Ru("043B 0435 043D 044B")
        Else
            '   !!! ошибка принудительного восстановления: 
            '   !!! error force restore: 
            If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043E 0448 0438 0431 043A 0430 0020 043F 0440 0438 043D 0443 0434 0438 0442 0435 043B 044C 043D 043E 0433 043E 0020 0432 043E 0441 0441 0442 0430") & Ru("043D 043E 0432 043B 0435 043D 0438 044F 003A 0020") & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    '   проверка размерности пройдена
    '   dimension check passed
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 0438 0020 043F 0440 043E 0439 0434 0435 043D 0430")
    
    ' --- validate PipeCountCP ---
    Dim pipeCount As Long
    '   проверка pipecountcp...
    '   check pipecountcp...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 002E 002E 002E")
    
    With wsAnod.Range("PipeCountCP")
        If IsEmpty(.Value) Or Not IsNumeric(.Value) Then
            '   !!! pipecountcp пуст или не число
            '   !!! pipecountcp is empty or not a number
            If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 043F 0443 0441 0442 0020 0438 043B 0438 0020 043D 0435 0020 0447 0438 0441 043B 043E")
            ' pipecountcp пуст или не число.
            ' pipecountcp is empty or not a number.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 043F 0443 0441 0442 0020 0438 043B 0438 0020 043D 0435 0020 0447 0438 0441 043B 043E 002E")
            Exit Sub
        End If
        pipeCount = CLng(.Value)
        If pipeCount <= 0 Or pipeCount > MAX_COL Then
            '   !!! pipecountcp вне диапазона: 
            '   !!! pipecountcp out of range: 
            If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0432 043D 0435 0020 0434 0438 0430 043F 0430 0437 043E 043D 0430 003A 0020") & pipeCount
            ' pipecountcp должен быть от 1 до 254.
            ' pipecountcp must be from 1 to 254.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043E 0442 0020 0031 0020 0434 043E 0020") & MAX_COL & "."
            Exit Sub
        End If
    End With
    If DEBUG_MODE Then Debug.Print "  pipeCount = " & pipeCount

    ' --- check that the Pipe sheet exists ---
    Dim wsPipe As Worksheet
    '   проверка листа pipe...
    '   checking pipe sheet...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 043B 0438 0441 0442 0430 0020 0050 0069 0070 0065 002E 002E 002E")
    On Error Resume Next
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If wsPipe Is Nothing Then
        '   !!! лист pipe не найден!
        '   !!! pipe sheet not found!
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043B 0438 0441 0442 0020 0050 0069 0070 0065 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021")
        ' критическая ошибка: лист '
        ' critical error: лandст '
        ' ' не найден!
        ' ' not found!
        MsgBox Ru("043A 0440 0438 0442 0438 0447 0435 0441 043A 0430 044F 0020 043E 0448 0438 0431 043A 0430 003A 0020 043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021")
        Exit Sub
    End If
    On Error GoTo 0
    '   лист pipe найден
    '   pipe sheet found
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043B 0438 0441 0442 0020 0050 0069 0070 0065 0020 043D 0430 0439 0434 0435 043D")

    ' --- load named ranges ---
    '   загрузка массивов данных...
    '   loading data arrays...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0437 0430 0433 0440 0443 0437 043A 0430 0020 043C 0430 0441 0441 0438 0432 043E 0432 0020 0434 0430 043D 043D 044B 0445 002E 002E 002E")
    ' ... remaining code unchanged ...
    ' --- load named ranges ---
    '   загрузка массивов данных...
    '   loading data arrays...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0437 0430 0433 0440 0443 0437 043A 0430 0020 043C 0430 0441 0441 0438 0432 043E 0432 0020 0434 0430 043D 043D 044B 0445 002E 002E 002E")

    Dim arrFactorVoltageMarginCP As Variant
    Dim arrNominalOutputVoltageCP As Variant
    Dim arrCurrentEndLifeCP As Variant
    Dim arrWiresResistanceCPpipeAG As Variant
    Dim arrResistivity_i_layerDeepAG As Variant
    Dim arrResistivitySoilAG As Variant
    Dim arrLengthElectrodeAG As Variant
    Dim arrDiameterAG As Variant
    Dim arrDepthToMidAG As Variant
    Dim arrCokeBreezelengthElectrodeAG As Variant
    Dim arrCokeBreezeDiameterAG As Variant
    Dim arrCokeBreezeResistivityAG As Variant
    Dim arrResistivityMaterialAG As Variant
    Dim arrMassOneElectrodeAG As Variant
    Dim arrSpecificMaccOneMeterAG As Variant
    Dim arrFactorUseMassAG As Variant
    Dim arrDissolutionRateAG As Variant
    Dim arrAvgProtectionCurrentCPOverLife As Variant
    Dim arrFactorSoilHeterogeneity As Variant
    Dim arrTypeInstallationAG As Variant
    Dim rngVoltageEndLife As Range
    Dim rngPowerEndLife As Range

    On Error Resume Next
    arrFactorVoltageMarginCP = SafeGetArray(wsAnod.Range("factorVoltageMarginCP"))
    arrNominalOutputVoltageCP = SafeGetArray(wsAnod.Range("nominalOutputVoltageCP"))
    arrCurrentEndLifeCP = SafeGetArray(wsAnod.Range("currentEndLifeCP"))
    arrWiresResistanceCPpipeAG = SafeGetArray(wsAnod.Range("wiresResistanceCPpipeAG"))
    arrResistivity_i_layerDeepAG = SafeGetArray(wsAnod.Range("resistivity_i_layerDeepAG"))
    arrResistivitySoilAG = SafeGetArray(wsAnod.Range("resistivitySoilAG"))
    arrLengthElectrodeAG = SafeGetArray(wsAnod.Range("lengthElectrodeAG"))
    arrDiameterAG = SafeGetArray(wsAnod.Range("diameterAG"))
    arrDepthToMidAG = SafeGetArray(wsAnod.Range("depthToMidAG"))
    arrCokeBreezelengthElectrodeAG = SafeGetArray(wsAnod.Range("cokeBreezelengthElectrodeAG"))
    arrCokeBreezeDiameterAG = SafeGetArray(wsAnod.Range("cokeBreezeDiameterAG"))
    arrCokeBreezeResistivityAG = SafeGetArray(wsAnod.Range("cokeBreezeResistivityAG"))
    arrResistivityMaterialAG = SafeGetArray(wsAnod.Range("resistivityMaterialAG"))
    arrMassOneElectrodeAG = SafeGetArray(wsAnod.Range("massOneElectrodeAG"))
    arrSpecificMaccOneMeterAG = SafeGetArray(wsAnod.Range("specificMaccOneMeterAG"))
    arrFactorUseMassAG = SafeGetArray(wsAnod.Range("factorUseMassAG"))
    arrDissolutionRateAG = SafeGetArray(wsAnod.Range("dissolutionRateAG"))
    arrAvgProtectionCurrentCPOverLife = SafeGetArray(wsAnod.Range("avgProtectionCurrentCPOverLife"))
    arrFactorSoilHeterogeneity = SafeGetArray(wsAnod.Range("factorSoilHeterogeneity"))
    arrTypeInstallationAG = SafeGetArray(wsAnod.Range("typeInstallationAG"))
    Set rngVoltageEndLife = wsAnod.Range("voltageEndLifeCP")
    Set rngPowerEndLife = wsAnod.Range("powerEndLifeCP")
    If Err.Number <> 0 Then
        '   !!! ошибка чтения исходных данных: 
        '   !!! read error input data: 
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0438 0441 0445 043E 0434 043D 044B 0445 0020 0434 0430 043D 043D 044B 0445") & Ru("003A 0020") & Err.Description
        ' ошибка чтения исходных данных: 
        ' read error input data: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0438 0441 0445 043E 0434 043D 044B 0445 0020 0434 0430 043D 043D 044B 0445 003A 0020") & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    '   массивы загружены
    '   arrays loaded
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043C 0430 0441 0441 0438 0432 044B 0020 0437 0430 0433 0440 0443 0436 0435 043D 044B")

    ' --- check sizes of single-row ranges ---
    '   проверка размерностей однострочных диапазонов...
    '   check размерностей single-row ranges...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 0435 0439 0020 043E 0434 043D 043E 0441 0442 0440 043E 0447 043D 044B") & Ru("0445 0020 0434 0438 0430 043F 0430 0437 043E 043D 043E 0432 002E 002E 002E")
    If UBound(arrFactorVoltageMarginCP, 2) < pipeCount Then CheckFailed "factorVoltageMarginCP"
    If UBound(arrNominalOutputVoltageCP, 2) < pipeCount Then CheckFailed "nominalOutputVoltageCP"
    If UBound(arrCurrentEndLifeCP, 2) < pipeCount Then CheckFailed "currentEndLifeCP"
    If UBound(arrWiresResistanceCPpipeAG, 2) < pipeCount Then CheckFailed "wiresResistanceCPpipeAG"
    If UBound(arrResistivity_i_layerDeepAG, 2) < pipeCount Then CheckFailed "resistivity_i_layerDeepAG"
    If UBound(arrResistivitySoilAG, 2) < pipeCount Then CheckFailed "resistivitySoilAG"
    If UBound(arrLengthElectrodeAG, 2) < pipeCount Then CheckFailed "lengthElectrodeAG"
    If UBound(arrDiameterAG, 2) < pipeCount Then CheckFailed "diameterAG"
    If UBound(arrDepthToMidAG, 2) < pipeCount Then CheckFailed "depthToMidAG"
    If UBound(arrCokeBreezelengthElectrodeAG, 2) < pipeCount Then CheckFailed "cokeBreezelengthElectrodeAG"
    If UBound(arrCokeBreezeDiameterAG, 2) < pipeCount Then CheckFailed "cokeBreezeDiameterAG"
    If UBound(arrCokeBreezeResistivityAG, 2) < pipeCount Then CheckFailed "cokeBreezeResistivityAG"
    If UBound(arrResistivityMaterialAG, 2) < pipeCount Then CheckFailed "resistivityMaterialAG"
    If UBound(arrMassOneElectrodeAG, 2) < pipeCount Then CheckFailed "massOneElectrodeAG"
    If UBound(arrSpecificMaccOneMeterAG, 2) < pipeCount Then CheckFailed "specificMaccOneMeterAG"
    If UBound(arrFactorUseMassAG, 2) < pipeCount Then CheckFailed "factorUseMassAG"
    If UBound(arrDissolutionRateAG, 2) < pipeCount Then CheckFailed "dissolutionRateAG"
    If UBound(arrAvgProtectionCurrentCPOverLife, 2) < pipeCount Then CheckFailed "avgProtectionCurrentCPOverLife"
    If UBound(arrFactorSoilHeterogeneity, 2) < pipeCount Then CheckFailed "factorSoilHeterogeneity"
    If UBound(arrTypeInstallationAG, 2) < pipeCount Then CheckFailed "typeInstallationAG"

    ' check sizes of calculated single-row ranges
    If wsAnod.Range("resistanceEndLifeAG").Columns.count < pipeCount Then CheckFailed "resistanceEndLifeAG"
    If wsAnod.Range("lengthWorkPartDeepAG").Columns.count < pipeCount Then CheckFailed "lengthWorkPartDeepAG"
    If rngVoltageEndLife.Columns.count < pipeCount Then CheckFailed "voltageEndLifeCP"
    If rngPowerEndLife.Columns.count < pipeCount Then CheckFailed "powerEndLifeCP"

    ' проверка однострочных диапазонов: 1 x pipeCountCP
    ' check single-row ranges: 1 x pipeCountCP
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 043E 0434 043D 043E 0441 0442 0440 043E 0447 043D 044B 0445 0020 0434 0438 0430 043F 0430 0437 043E 043D 043E 0432") & Ru("002E 002E 002E")
    ' oneelectroderesistanceag (строки: ожидается 1)
    ' oneelectroderesistanceag (rows: expected 1)
    If wsAnod.Range("oneElectrodeResistanceAG").rows.count <> 1 Then CheckFailed Ru("006F 006E 0065 0045 006C 0065 0063 0074 0072 006F 0064 0065 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E") & Ru("0436 0438 0434 0430 0435 0442 0441 044F 0020 0031 0029")
    ' oneelectroderesistanceag (столбцы: ожидается pipecountcp)
    ' oneelectroderesistanceag (columns: expected pipecountcp)
    If wsAnod.Range("oneElectrodeResistanceAG").Columns.count <> pipeCount Then CheckFailed Ru("006F 006E 0065 0045 006C 0065 0063 0074 0072 006F 0064 0065 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020") & Ru("043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' numelectrodesag (строки: ожидается 1)
    ' numelectrodesag (rows: expected 1)
    If wsAnod.Range("numElectrodesAG").rows.count <> 1 Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & Ru("0031 0029")
    ' numelectrodesag (столбцы: ожидается pipecountcp)
    ' numelectrodesag (columns: expected pipecountcp)
    If wsAnod.Range("numElectrodesAG").Columns.count <> pipeCount Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F") & Ru("0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' weightwithoutfillingag (строки: ожидается 1)
    ' weightwithoutfillingag (rows: expected 1)
    If wsAnod.Range("weightWithoutFillingAG").rows.count <> 1 Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438") & Ru("0434 0430 0435 0442 0441 044F 0020 0031 0029")
    ' weightwithoutfillingag (столбцы: ожидается pipecountcp)
    ' weightwithoutfillingag (columns: expected pipecountcp)
    If wsAnod.Range("weightWithoutFillingAG").Columns.count <> pipeCount Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436") & Ru("0438 0434 0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifeag (строки: ожидается 1)
    ' servicelifeag (rows: expected 1)
    If wsAnod.Range("serviceLifeAG").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0031") & Ru("0029")
    ' servicelifeag (столбцы: ожидается pipecountcp)
    ' servicelifeag (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeAG").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0050") & Ru("0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifedeviation (строки: ожидается 1)
    ' servicelifedeviation (rows: expected 1)
    If wsAnod.Range("serviceLifeDeviation").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0031 0029")
    ' servicelifedeviation (столбцы: ожидается pipecountcp)
    ' servicelifedeviation (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeDeviation").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434") & Ru("0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' correctresistanceag (строки: ожидается 1)
    ' correctresistanceag (rows: expected 1)
    If wsAnod.Range("correctResistanceAG").rows.count <> 1 Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435") & Ru("0442 0441 044F 0020 0031 0029")
    ' correctresistanceag (столбцы: ожидается pipecountcp)
    ' correctresistanceag (columns: expected pipecountcp)
    If wsAnod.Range("correctResistanceAG").Columns.count <> pipeCount Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    '   все проверки размерности пройдены
    '   all dimension checks passed
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0432 0441 0435 0020 043F 0440 043E 0432 0435 0440 043A 0438 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 0438 0020 043F 0440 043E 0439 0434 0435 043D 044B")

    ' ============================================================
    ' turn off screen updating, events, and automatic calculation
    ' ============================================================
    '   начало основного расчета...
    '   starting main calculation...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 043D 0430 0447 0430 043B 043E 0020 043E 0441 043D 043E 0432 043D 043E 0433 043E 0020 0440 0430 0441 0447 0435 0442 0430 002E 002E 002E")
    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    appStateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim piVal As Double
    piVal = Application.Pi()

    Dim rngRzOut As Range
    Dim rngLzOut As Range
    Dim rngRp1Out As Range
    Dim rngRp1hOut As Range
    Dim rngNOut As Range
    Dim rngMassOut As Range
    Dim rngLifeOut As Range
    Dim rngKtOut As Range
    Dim rngRp1pOut As Range
    Set rngRzOut = wsAnod.Range("resistanceEndLifeAG")
    Set rngLzOut = wsAnod.Range("lengthWorkPartDeepAG")
    Set rngRp1Out = wsAnod.Range("oneElectrodeResistanceAG")
    On Error Resume Next
    Set rngRp1hOut = wsAnod.Range("oneElectrodeResistanceHorizAG")
    Err.Clear
    On Error GoTo 0
    Set rngNOut = wsAnod.Range("numElectrodesAG")
    Set rngMassOut = wsAnod.Range("weightWithoutFillingAG")
    Set rngLifeOut = wsAnod.Range("serviceLifeAG")
    Set rngKtOut = wsAnod.Range("serviceLifeDeviation")
    Set rngRp1pOut = wsAnod.Range("correctResistanceAG")

    ' clear calculated ranges
    Application.StatusBar = "clearing calculated ranges..."
    ClearRangePart rngRzOut, 1, 1, 1, pipeCount
    ClearRangePart rngVoltageEndLife, 1, 1, 1, pipeCount
    ClearRangePart rngPowerEndLife, 1, 1, 1, pipeCount
    ClearRangePart rngLzOut, 1, 1, 1, pipeCount
    ClearRangePart rngRp1Out, 1, 1, 1, pipeCount
    If Not rngRp1hOut Is Nothing Then ClearRangePart rngRp1hOut, 1, 1, 1, pipeCount
    ClearRangePart rngNOut, 1, 1, 1, pipeCount
    ClearRangePart rngMassOut, 1, 1, 1, pipeCount
    ClearRangePart rngLifeOut, 1, 1, 1, pipeCount
    ClearRangePart rngKtOut, 1, 1, 1, pipeCount
    ClearRangePart rngRp1pOut, 1, 1, 1, pipeCount

    ' --- scalar value from the Pipe sheet ---
    Dim pipeInputResistanceEndLife As Double
    On Error Resume Next
    pipeInputResistanceEndLife = GetNumeric(wsPipe.Range("pipeInputResistanceEndLife").Value, 0)
    If Err.Number <> 0 Then
        '   !!! ошибка чтения pipeinputresistanceendlife: 
        '   !!! read error pipeinputresistanceendlife: 
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0070 0069 0070 0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074") & Ru("0061 006E 0063 0065 0045 006E 0064 004C 0069 0066 0065 003A 0020") & Err.Description
        ' ошибка чтения pipeinputresistanceendlife: 
        ' read error pipeinputresistanceendlife: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0070 0069 0070 0065 0049 006E 0070 0075 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0045 006E") & Ru("0064 004C 0069 0066 0065 003A 0020") & Err.Description
        GoTo CleanExit
    End If
    On Error GoTo 0

    ' --- serviceLifeDesigned - design service life, years ---
    Dim serviceLifeDesignedMax As Double
    Dim arrServiceLifeDesigned As Variant
    Dim rngServiceLife As Range
    Dim serviceLifeVal As Variant

    On Error Resume Next
    Set rngServiceLife = wsPipe.Range("serviceLifeDesigned")
    If Not rngServiceLife Is Nothing Then
        serviceLifeVal = rngServiceLife.Value
        If Not IsArray(serviceLifeVal) Then
            ReDim arrServiceLifeDesigned(1 To 1, 1 To 1)
            arrServiceLifeDesigned(1, 1) = serviceLifeVal
        Else
            arrServiceLifeDesigned = ReadRangeAs1xN(rngServiceLife)
        End If
        serviceLifeDesignedMax = GetMaxFrom1xN(arrServiceLifeDesigned, 30)
    Else
        serviceLifeDesignedMax = 30
    End If
    If Err.Number <> 0 Then
        '   !!! ошибка чтения servicelifedesigned: 
        '   !!! read error servicelifedesigned: 
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 0021 0021 0021 0020 043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0073 0069") & Ru("0067 006E 0065 0064 003A 0020") & Err.Description
        ' ошибка чтения servicelifedesigned: 
        ' read error servicelifedesigned: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0073 0069 0067 006E 0065 0064 003A 0020") & Err.Description
        GoTo CleanExit
    End If
    On Error GoTo 0

    If serviceLifeDesignedMax <= 0 Then serviceLifeDesignedMax = 30
    If DEBUG_MODE Then Debug.Print "  serviceLifeDesignedMax = " & serviceLifeDesignedMax

    ' --- arrays for intermediate values ---
    Dim RzValues() As Double
    ReDim RzValues(1 To pipeCount)
    Dim Rp1Values() As Double
    ReDim Rp1Values(1 To pipeCount, 1 To 9)

    Dim outRz() As Variant
    Dim outVolt() As Variant
    Dim outPow() As Variant
    Dim outRp1() As Variant
    Dim outRp1h() As Variant
    Dim outN() As Variant
    Dim outMass() As Variant
    Dim outLz() As Variant
    Dim outLife() As Variant
    Dim outKt() As Variant
    Dim outRp1p() As Variant
    ReDim outRz(1 To 1, 1 To pipeCount)
    ReDim outVolt(1 To 1, 1 To pipeCount)
    ReDim outPow(1 To 1, 1 To pipeCount)
    ReDim outRp1(1 To 1, 1 To pipeCount)
    ReDim outRp1h(1 To 1, 1 To pipeCount)
    ReDim outN(1 To 1, 1 To pipeCount)
    ReDim outMass(1 To 1, 1 To pipeCount)
    ReDim outLz(1 To 1, 1 To pipeCount)
    ReDim outLife(1 To 1, 1 To pipeCount)
    ReDim outKt(1 To 1, 1 To pipeCount)
    ReDim outRp1p(1 To 1, 1 To pipeCount)

    ' ----------------------------------------------------------------
    ' calculate Rz and Rp1 for all columns
    ' ----------------------------------------------------------------
    Application.StatusBar = "calc Rz and Rp1..."
    Dim col As Long
    Dim factorV As Double, nomV As Double, curEnd As Double, wireRes As Double
    Dim rho_soil As Double, rho_layer As Double, rho_coke As Double, rho_mat As Double
    Dim l_el As Double, d_el As Double, h As Double, l_coke As Double, d_coke As Double
    Dim Rz As Double, voltage As Double, S_el As Double

    For col = 1 To pipeCount
        factorV = GetNumeric(arrFactorVoltageMarginCP(1, col), 0)
        nomV = GetNumeric(arrNominalOutputVoltageCP(1, col), 0)
        curEnd = GetNumeric(arrCurrentEndLifeCP(1, col), 0)
        wireRes = GetNumeric(arrWiresResistanceCPpipeAG(1, col), 0)
        rho_soil = GetNumeric(arrResistivitySoilAG(1, col), 0)
        rho_layer = GetNumeric(arrResistivity_i_layerDeepAG(1, col), 0)
        l_el = GetNumeric(arrLengthElectrodeAG(1, col), 0)
        d_el = GetNumeric(arrDiameterAG(1, col), 0)
        h = GetNumeric(arrDepthToMidAG(1, col), 0)
        l_coke = GetNumeric(arrCokeBreezelengthElectrodeAG(1, col), 0)
        d_coke = GetNumeric(arrCokeBreezeDiameterAG(1, col), 0)
        rho_coke = GetNumeric(arrCokeBreezeResistivityAG(1, col), 0)
        rho_mat = GetNumeric(arrResistivityMaterialAG(1, col), 0)

        ' formula 6.25
        If curEnd <> 0 And factorV > 0 And nomV > 0 Then
            Rz = (factorV * nomV / curEnd) - (pipeInputResistanceEndLife + wireRes)
            If Rz < 0 Then Rz = 0
            voltage = curEnd * (pipeInputResistanceEndLife + wireRes + Rz)
        Else
            Rz = 0
            voltage = 0
        End If

        RzValues(col) = Rz
        outRz(1, col) = Rz
        outVolt(1, col) = voltage
        outPow(1, col) = curEnd * voltage

        ' Rp1 by type
        ' vertical (6.28)
        If l_el > 0 And d_el > 0 And h > 0 And (4 * h - l_el) > 0 Then
            Rp1Values(col, 1) = (rho_soil / (2 * piVal * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
        End If
        ' horizontal le<h (6.29)
        If l_el > 0 And d_el > 0 Then
            Rp1Values(col, 2) = (rho_soil / (2 * piVal * l_el)) * Log(2 * l_el / d_el)
        End If
        ' horizontal with activator (6.31)
        If l_coke > 0 And d_coke > 0 And h > 0 And rho_coke > 0 And rho_soil > 0 And d_el > 0 Then
            Rp1Values(col, 4) = (rho_soil / (2 * piVal * l_coke)) * (Log(2 * l_coke / d_coke) + Log((l_coke + Sqr(l_coke ^ 2 + 16 * h ^ 2)) / (4 * h)) + (rho_coke / rho_soil) * Log(d_coke / d_el))
        End If
        ' extended (6.30)
        If l_el > 0 And d_el > 0 And h > 0 Then
            Rp1Values(col, 5) = (rho_soil / (piVal * l_el)) * Log(l_el / Sqr(d_el * h))
        End If
        ' extended with activator (6.32)
        If l_coke > 0 And d_coke > 0 And h > 0 And rho_coke > 0 And rho_soil > 0 And d_el > 0 Then
            Rp1Values(col, 6) = (rho_soil / (piVal * l_coke)) * (Log(l_coke / Sqr(d_coke * h)) + (rho_coke / (2 * rho_soil)) * Log(d_coke / d_el))
        End If
        ' deep without activator (6.28 with rho_layer)
        If l_el > 0 And d_el > 0 And h > 0 And (4 * h - l_el) > 0 Then
            Rp1Values(col, 7) = (rho_layer / (2 * piVal * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
        End If
        ' deep with end outlet (6.13)
        If l_coke > 0 And d_coke > 0 And rho_coke > 0 And rho_layer > 0 And d_el > 0 Then
            S_el = piVal * d_el ^ 2 / 4
            If S_el > 0 Then
                Rp1Values(col, 8) = (l_coke * rho_mat) / (2 * S_el) + (rho_layer / (2 * piVal * l_coke)) * (Log(4 * l_coke / d_coke) + (rho_coke / rho_layer) * Log(d_coke / d_el))
            End If
        End If
        ' deep without end outlet (6.14)
        If l_coke > 0 And d_coke > 0 And rho_coke > 0 And rho_layer > 0 And d_el > 0 And h > 0 And (4 * h - l_coke) > 0 Then
            S_el = piVal * d_el ^ 2 / 4
            If S_el > 0 Then
                Rp1Values(col, 9) = (l_coke * rho_mat) / (2 * S_el) + (rho_layer / (2 * piVal * l_coke)) * (Log(2 * l_coke / d_coke) + 0.5 * Log((4 * h + l_coke) / (4 * h - l_coke)) + (rho_coke / rho_layer) * Log(d_coke / d_el))
            End If
        End If
    Next col
    '   rz и rp1 рассчитаны
    '   rz and rp1 calculated
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0052 007A 0020 0438 0020 0052 0070 0031 0020 0440 0430 0441 0441 0447 0438 0442 0430 043D 044B")

    ' ----------------------------------------------------------------
    ' first loop: compute N and mass, write all parameters except service life and deviation
    ' ----------------------------------------------------------------
    Application.StatusBar = "calc N and mass..."
    Dim cellType As Variant
    Dim iType As Integer
    Dim n As Double, mass As Double, Rp1prime As Double
    Dim massOne As Double, specMass As Double
    Dim factorUse As Double, dissolRate As Double, avgCur As Double, factorHet As Double
    Dim lz As Double
    Dim R_vert As Double, R_horiz As Double

    For col = 1 To pipeCount
        cellType = arrTypeInstallationAG(1, col)
        If IsEmpty(cellType) Then GoTo NextCol1
        iType = Module_Visual.GetInstallationTypeIndex(CStr(cellType))
        If iType = 0 Then GoTo NextCol1

        factorV = GetNumeric(arrFactorVoltageMarginCP(1, col), 0)
        massOne = GetNumeric(arrMassOneElectrodeAG(1, col), 0)
        specMass = GetNumeric(arrSpecificMaccOneMeterAG(1, col), 0)
        factorUse = GetNumeric(arrFactorUseMassAG(1, col), 0)
        dissolRate = GetNumeric(arrDissolutionRateAG(1, col), 0)
        avgCur = GetNumeric(arrAvgProtectionCurrentCPOverLife(1, col), 0)
        factorHet = GetNumeric(arrFactorSoilHeterogeneity(1, col), 0)
        l_el = GetNumeric(arrLengthElectrodeAG(1, col), 0)
        Rz = RzValues(col)

        ' compute N
        Select Case iType
            Case 1
                If factorV > 0 And Rz > 0 And Rp1Values(col, 1) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 1) / (factorV * Rz), 0)
                Else: n = 0
                End If
            Case 2
                If factorV > 0 And Rz > 0 And Rp1Values(col, 2) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 2) / (factorV * Rz), 0)
                Else: n = 0
                End If
            Case 3
                R_vert = Rp1Values(col, 1)
                R_horiz = Rp1Values(col, 2)
                If factorV > 0 And Rz > 0 And R_horiz > 0 Then
                    n = Application.RoundUp((1.19 * R_vert * R_horiz - 0.98 * Rz * R_vert) / (factorV * Rz * R_horiz), 0)
                Else: n = 0
                End If
            Case 4
                If factorV > 0 And Rz > 0 And Rp1Values(col, 4) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 4) / (factorV * Rz), 0)
                Else: n = 0
                End If
            Case 5
                If factorV > 0 And Rz > 0 And Rp1Values(col, 5) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 5) / (factorV * Rz), 2)
                Else: n = 0
                End If
            Case 6
                If factorV > 0 And Rz > 0 And Rp1Values(col, 6) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 6) / (factorV * Rz), 2)
                Else: n = 0
                End If
            Case 7
                If factorV > 0 And Rz > 0 And Rp1Values(col, 7) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 7) / (factorV * Rz), 0)
                Else: n = 0
                End If
            Case 8
                If factorV > 0 And Rz > 0 And Rp1Values(col, 8) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 8) / (factorV * Rz), 0)
                Else: n = 0
                End If
            Case 9
                If factorV > 0 And Rz > 0 And Rp1Values(col, 9) > 0 Then
                    n = Application.RoundUp(Rp1Values(col, 9) / (factorV * Rz), 0)
                Else: n = 0
                End If
        End Select

        ' mass
        If iType = 5 Or iType = 6 Then
            mass = specMass * l_el * n
        Else
            mass = massOne * n
        End If

        ' Rp1'
        Rp1prime = n * factorV * Rz

        ' write results into the constant 1-row ranges
        If iType = 3 Then
            ' комбинированный: на лист пишем Rp1 вертикального и горизонтального электродов
            ' combined: write Rp1 of vertical and horizontal electrodes to the sheet
            outRp1(1, col) = Rp1Values(col, 1)
            outRp1h(1, col) = Rp1Values(col, 2)
        Else
            outRp1(1, col) = Rp1Values(col, iType)
        End If
        outN(1, col) = n
        outMass(1, col) = mass
        outRp1p(1, col) = Rp1prime

        ' working length of the deep anode
        If iType >= 7 And iType <= 9 Then
            If Rz > 0 Then
                lz = 3.5 * GetNumeric(arrResistivity_i_layerDeepAG(1, col), 0) / (piVal * Rz)
            Else: lz = 0
            End If
            outLz(1, col) = lz
        End If

NextCol1:
    Next col
    '   n и масса рассчитаны
    '   n and масса calculated
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 004E 0020 0438 0020 043C 0430 0441 0441 0430 0020 0440 0430 0441 0441 0447 0438 0442 0430 043D 044B")

    ' ----------------------------------------------------------------
    ' second loop: service life and deviation
    ' ----------------------------------------------------------------
    Application.StatusBar = "calc service life..."
    Dim life As Double, kt As Double
    Dim denom As Double
    Dim massFromSheet As Double

    For col = 1 To pipeCount
        cellType = arrTypeInstallationAG(1, col)
        If IsEmpty(cellType) Then GoTo NextCol2
        iType = Module_Visual.GetInstallationTypeIndex(CStr(cellType))
        If iType = 0 Then GoTo NextCol2

        massFromSheet = GetNumeric(outMass(1, col), 0)

        factorUse = GetNumeric(arrFactorUseMassAG(1, col), 0)
        dissolRate = GetNumeric(arrDissolutionRateAG(1, col), 0)
        avgCur = GetNumeric(arrAvgProtectionCurrentCPOverLife(1, col), 0)
        factorHet = GetNumeric(arrFactorSoilHeterogeneity(1, col), 0)

        ' service life
        If iType = 5 Or iType = 6 Or iType = 7 Or iType = 8 Or iType = 9 Then
            denom = dissolRate * avgCur * factorHet
        Else
            denom = dissolRate * avgCur
        End If

        If denom > 0 And massFromSheet > 0 Then
            life = Application.RoundDown((massFromSheet * factorUse) / denom, 0)
        Else
            life = 0
        End If

        ' deviation
        If serviceLifeDesignedMax > 0 Then
            kt = (serviceLifeDesignedMax - life) / serviceLifeDesignedMax
        Else
            kt = 0
        End If

        outLife(1, col) = life
        outKt(1, col) = kt

NextCol2:
    Next col
    '   срок службы рассчитан
    '   service life calculated
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0441 0440 043E 043A 0020 0441 043B 0443 0436 0431 044B 0020 0440 0430 0441 0441 0447 0438 0442 0430 043D")

    Application.StatusBar = "writing results..."
    Call Write1xN(rngRzOut, outRz, pipeCount)
    Call Write1xN(rngVoltageEndLife, outVolt, pipeCount)
    Call Write1xN(rngPowerEndLife, outPow, pipeCount)
    Call Write1xN(rngLzOut, outLz, pipeCount)
    Call Write1xN(rngRp1Out, outRp1, pipeCount)
    If Not rngRp1hOut Is Nothing Then Call Write1xN(rngRp1hOut, outRp1h, pipeCount)
    Call Write1xN(rngNOut, outN, pipeCount)
    Call Write1xN(rngMassOut, outMass, pipeCount)
    Call Write1xN(rngLifeOut, outLife, pipeCount)
    Call Write1xN(rngKtOut, outKt, pipeCount)
    Call Write1xN(rngRp1pOut, outRp1p, pipeCount)

    ' format visible calculation rows
    '   форматирование строк...
    '   formatting rows...
    If DEBUG_MODE Then Debug.Print Ru("0020 0020 0444 043E 0440 043C 0430 0442 0438 0440 043E 0432 0430 043D 0438 0435 0020 0441 0442 0440 043E 043A 002E 002E 002E")
    Call Module_Visual.UpdateVisibilityForAllColumns(wsAnod)
    Call Module_Visual.UpdateColorsForAllColumns(wsAnod)

    ' === btnanodfullcalc: успешно завершен ===
    ' === btnanodfullcalc: finished successfully ===
    If DEBUG_MODE Then Debug.Print Ru("003D 003D 003D 0020 0062 0074 006E 0041 006E 006F 0064 0046 0075 006C 006C 0043 0061 006C 0063 003A 0020 0443 0441 043F 0435 0448 043D 043E 0020 0437 0430 0432 0435 0440 0448") & Ru("0435 043D 0020 003D 003D 003D")
    GoTo CleanExit
    
CleanExit:
    Application.StatusBar = False
    If appStateSaved Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.Calculation = oldCalc
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
    If Err.Number <> 0 Then
        ' === btnanodfullcalc: ошибка ===
        ' === btnanodfullcalc: error ===
        If DEBUG_MODE Then Debug.Print Ru("003D 003D 003D 0020 0062 0074 006E 0041 006E 006F 0064 0046 0075 006C 006C 0043 0061 006C 0063 003A 0020 043E 0448 0438 0431 043A 0430 0020 003D 003D 003D")
        '   ошибка: 
        '   error: 
        If DEBUG_MODE Then Debug.Print Ru("0020 0020 043E 0448 0438 0431 043A 0430 003A 0020") & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Sub

' ================================================================
' 2. recalculate for one type (including N)
' ================================================================
Public Sub WriteFormulasToRow(relativeRow As Long, formulaType As String)
    On Error GoTo CleanExit
    Dim pipeCount As Long
    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets("Anod")
    
    With wsAnod.Range("PipeCountCP")
        If IsEmpty(.Value) Or Not IsNumeric(.Value) Then
            ' pipecountcp пуст или не число.
            ' pipecountcp is empty or not a number.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 043F 0443 0441 0442 0020 0438 043B 0438 0020 043D 0435 0020 0447 0438 0441 043B 043E 002E")
            Exit Sub
        End If
        pipeCount = CLng(.Value)
        If pipeCount <= 0 Or pipeCount > MAX_COL Then
            ' pipecountcp должен быть от 1 до 254.
            ' pipecountcp must be from 1 to 254.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043E 0442 0020 0031 0020 0434 043E 0020") & MAX_COL & "."
            Exit Sub
        End If
    End With

    ' load all input arrays
    Dim arrFactorVoltageMarginCP As Variant
    Dim arrNominalOutputVoltageCP As Variant
    Dim arrCurrentEndLifeCP As Variant
    Dim arrWiresResistanceCPpipeAG As Variant
    Dim arrResistivity_i_layerDeepAG As Variant
    Dim arrResistivitySoilAG As Variant
    Dim arrLengthElectrodeAG As Variant
    Dim arrDiameterAG As Variant
    Dim arrDepthToMidAG As Variant
    Dim arrCokeBreezelengthElectrodeAG As Variant
    Dim arrCokeBreezeDiameterAG As Variant
    Dim arrCokeBreezeResistivityAG As Variant
    Dim arrResistivityMaterialAG As Variant
    Dim arrMassOneElectrodeAG As Variant
    Dim arrSpecificMaccOneMeterAG As Variant
    Dim arrFactorUseMassAG As Variant
    Dim arrDissolutionRateAG As Variant
    Dim arrAvgProtectionCurrentCPOverLife As Variant
    Dim arrFactorSoilHeterogeneity As Variant

    On Error Resume Next
    arrFactorVoltageMarginCP = SafeGetArray(wsAnod.Range("factorVoltageMarginCP"))
    arrNominalOutputVoltageCP = SafeGetArray(wsAnod.Range("nominalOutputVoltageCP"))
    arrCurrentEndLifeCP = SafeGetArray(wsAnod.Range("currentEndLifeCP"))
    arrWiresResistanceCPpipeAG = SafeGetArray(wsAnod.Range("wiresResistanceCPpipeAG"))
    arrResistivity_i_layerDeepAG = SafeGetArray(wsAnod.Range("resistivity_i_layerDeepAG"))
    arrResistivitySoilAG = SafeGetArray(wsAnod.Range("resistivitySoilAG"))
    arrLengthElectrodeAG = SafeGetArray(wsAnod.Range("lengthElectrodeAG"))
    arrDiameterAG = SafeGetArray(wsAnod.Range("diameterAG"))
    arrDepthToMidAG = SafeGetArray(wsAnod.Range("depthToMidAG"))
    arrCokeBreezelengthElectrodeAG = SafeGetArray(wsAnod.Range("cokeBreezelengthElectrodeAG"))
    arrCokeBreezeDiameterAG = SafeGetArray(wsAnod.Range("cokeBreezeDiameterAG"))
    arrCokeBreezeResistivityAG = SafeGetArray(wsAnod.Range("cokeBreezeResistivityAG"))
    arrResistivityMaterialAG = SafeGetArray(wsAnod.Range("resistivityMaterialAG"))
    arrMassOneElectrodeAG = SafeGetArray(wsAnod.Range("massOneElectrodeAG"))
    arrSpecificMaccOneMeterAG = SafeGetArray(wsAnod.Range("specificMaccOneMeterAG"))
    arrFactorUseMassAG = SafeGetArray(wsAnod.Range("factorUseMassAG"))
    arrDissolutionRateAG = SafeGetArray(wsAnod.Range("dissolutionRateAG"))
    arrAvgProtectionCurrentCPOverLife = SafeGetArray(wsAnod.Range("avgProtectionCurrentCPOverLife"))
    arrFactorSoilHeterogeneity = SafeGetArray(wsAnod.Range("factorSoilHeterogeneity"))
    If Err.Number <> 0 Then
        ' ошибка чтения исходных данных: 
        ' read error input data: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0438 0441 0445 043E 0434 043D 044B 0445 0020 0434 0430 043D 043D 044B 0445 003A 0020") & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' dimension checks
    If UBound(arrFactorVoltageMarginCP, 2) < pipeCount Then CheckFailed "factorVoltageMarginCP"
    If UBound(arrNominalOutputVoltageCP, 2) < pipeCount Then CheckFailed "nominalOutputVoltageCP"
    If UBound(arrCurrentEndLifeCP, 2) < pipeCount Then CheckFailed "currentEndLifeCP"
    If UBound(arrWiresResistanceCPpipeAG, 2) < pipeCount Then CheckFailed "wiresResistanceCPpipeAG"
    If UBound(arrResistivity_i_layerDeepAG, 2) < pipeCount Then CheckFailed "resistivity_i_layerDeepAG"
    If UBound(arrResistivitySoilAG, 2) < pipeCount Then CheckFailed "resistivitySoilAG"
    If UBound(arrLengthElectrodeAG, 2) < pipeCount Then CheckFailed "lengthElectrodeAG"
    If UBound(arrDiameterAG, 2) < pipeCount Then CheckFailed "diameterAG"
    If UBound(arrDepthToMidAG, 2) < pipeCount Then CheckFailed "depthToMidAG"
    If UBound(arrCokeBreezelengthElectrodeAG, 2) < pipeCount Then CheckFailed "cokeBreezelengthElectrodeAG"
    If UBound(arrCokeBreezeDiameterAG, 2) < pipeCount Then CheckFailed "cokeBreezeDiameterAG"
    If UBound(arrCokeBreezeResistivityAG, 2) < pipeCount Then CheckFailed "cokeBreezeResistivityAG"
    If UBound(arrResistivityMaterialAG, 2) < pipeCount Then CheckFailed "resistivityMaterialAG"
    If UBound(arrMassOneElectrodeAG, 2) < pipeCount Then CheckFailed "massOneElectrodeAG"
    If UBound(arrSpecificMaccOneMeterAG, 2) < pipeCount Then CheckFailed "specificMaccOneMeterAG"
    If UBound(arrFactorUseMassAG, 2) < pipeCount Then CheckFailed "factorUseMassAG"
    If UBound(arrDissolutionRateAG, 2) < pipeCount Then CheckFailed "dissolutionRateAG"
    If UBound(arrAvgProtectionCurrentCPOverLife, 2) < pipeCount Then CheckFailed "avgProtectionCurrentCPOverLife"
    If UBound(arrFactorSoilHeterogeneity, 2) < pipeCount Then CheckFailed "factorSoilHeterogeneity"

    ' проверка однострочных диапазонов: 1 x pipeCountCP
    ' check single-row ranges: 1 x pipeCountCP
    ' numelectrodesag (строки: ожидается 1)
    ' numelectrodesag (rows: expected 1)
    If wsAnod.Range("numElectrodesAG").rows.count <> 1 Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & Ru("0031 0029")
    ' numelectrodesag (столбцы: ожидается pipecountcp)
    ' numelectrodesag (columns: expected pipecountcp)
    If wsAnod.Range("numElectrodesAG").Columns.count <> pipeCount Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F") & Ru("0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' weightwithoutfillingag (строки: ожидается 1)
    ' weightwithoutfillingag (rows: expected 1)
    If wsAnod.Range("weightWithoutFillingAG").rows.count <> 1 Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438") & Ru("0434 0430 0435 0442 0441 044F 0020 0031 0029")
    ' weightwithoutfillingag (столбцы: ожидается pipecountcp)
    ' weightwithoutfillingag (columns: expected pipecountcp)
    If wsAnod.Range("weightWithoutFillingAG").Columns.count <> pipeCount Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436") & Ru("0438 0434 0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifeag (строки: ожидается 1)
    ' servicelifeag (rows: expected 1)
    If wsAnod.Range("serviceLifeAG").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0031") & Ru("0029")
    ' servicelifeag (столбцы: ожидается pipecountcp)
    ' servicelifeag (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeAG").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0050") & Ru("0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifedeviation (строки: ожидается 1)
    ' servicelifedeviation (rows: expected 1)
    If wsAnod.Range("serviceLifeDeviation").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0031 0029")
    ' servicelifedeviation (столбцы: ожидается pipecountcp)
    ' servicelifedeviation (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeDeviation").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434") & Ru("0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' correctresistanceag (строки: ожидается 1)
    ' correctresistanceag (rows: expected 1)
    If wsAnod.Range("correctResistanceAG").rows.count <> 1 Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435") & Ru("0442 0441 044F 0020 0031 0029")
    ' correctresistanceag (столбцы: ожидается pipecountcp)
    ' correctresistanceag (columns: expected pipecountcp)
    If wsAnod.Range("correctResistanceAG").Columns.count <> pipeCount Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' clear the constant result rows (one row per parameter)
    ClearRangePart wsAnod.Range("numElectrodesAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("weightWithoutFillingAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("serviceLifeAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("serviceLifeDeviation"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("correctResistanceAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("oneElectrodeResistanceAG"), 1, 1, 1, pipeCount
    On Error Resume Next
    ClearRangePart wsAnod.Range("oneElectrodeResistanceHorizAG"), 1, 1, 1, pipeCount
    Err.Clear
    On Error GoTo 0
    If relativeRow >= 7 Then ClearRangePart wsAnod.Range("lengthWorkPartDeepAG"), 1, 1, 1, pipeCount

    ' read serviceLifeDesignedMax and pipeInputResistanceEndLife
    Dim wsPipe As Worksheet
    Dim serviceLifeDesignedMax As Double
    Dim arrServiceLifeDesigned As Variant
    Dim rngServiceLife As Range
    Dim serviceLifeVal As Variant
    Dim pipeInputResistanceEndLife As Double

    On Error Resume Next
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If Not wsPipe Is Nothing Then
        Set rngServiceLife = wsPipe.Range("serviceLifeDesigned")
        If Not rngServiceLife Is Nothing Then
            serviceLifeVal = rngServiceLife.Value
            If Not IsArray(serviceLifeVal) Then
                ReDim arrServiceLifeDesigned(1 To 1, 1 To 1)
                arrServiceLifeDesigned(1, 1) = serviceLifeVal
            Else
                arrServiceLifeDesigned = ReadRangeAs1xN(rngServiceLife)
            End If
            serviceLifeDesignedMax = GetMaxFrom1xN(arrServiceLifeDesigned, 30)
        Else
            serviceLifeDesignedMax = 30
        End If
    Else
        serviceLifeDesignedMax = 30
    End If
    pipeInputResistanceEndLife = GetNumeric(wsPipe.Range("pipeInputResistanceEndLife").Value, 0)
    If Err.Number <> 0 Then
        ' ошибка чтения данных с листа pipe: 
        ' read error данных с pipe sheet: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0434 0430 043D 043D 044B 0445 0020 0441 0020 043B 0438 0441 0442 0430 0020 0050 0069 0070 0065 003A 0020") & Err.Description
        GoTo CleanExit
    End If
    On Error GoTo 0
    If serviceLifeDesignedMax <= 0 Then serviceLifeDesignedMax = 30

    Dim col As Long
    Dim cellType As Variant, typeStr As String
    Dim iType As Integer
    Dim factorV As Double, nomV As Double, curEnd As Double, wireRes As Double
    Dim rho_soil As Double, rho_layer As Double, rho_coke As Double, rho_mat As Double
    Dim l_el As Double, d_el As Double, h As Double, l_coke As Double, d_coke As Double
    Dim Rz As Double, Rp1 As Double, S_el As Double
    Dim n As Double, mass As Double, life As Double, kt As Double, Rp1prime As Double
    Dim massOne As Double, specMass As Double
    Dim factorUse As Double, dissolRate As Double, avgCur As Double, factorHet As Double
    Dim denom As Double, lz As Double
    Dim R_vert As Double, R_horiz As Double

    For col = 1 To pipeCount
        cellType = wsAnod.Range("typeInstallationAG").Cells(1, col).Value
        If IsEmpty(cellType) Then GoTo NextCol
        iType = Module_Visual.GetInstallationTypeIndex(CStr(cellType))
        If iType <> relativeRow Then GoTo NextCol

        factorV = GetNumeric(arrFactorVoltageMarginCP(1, col), 0)
        nomV = GetNumeric(arrNominalOutputVoltageCP(1, col), 0)
        curEnd = GetNumeric(arrCurrentEndLifeCP(1, col), 0)
        wireRes = GetNumeric(arrWiresResistanceCPpipeAG(1, col), 0)
        rho_soil = GetNumeric(arrResistivitySoilAG(1, col), 0)
        rho_layer = GetNumeric(arrResistivity_i_layerDeepAG(1, col), 0)
        l_el = GetNumeric(arrLengthElectrodeAG(1, col), 0)
        d_el = GetNumeric(arrDiameterAG(1, col), 0)
        h = GetNumeric(arrDepthToMidAG(1, col), 0)
        l_coke = GetNumeric(arrCokeBreezelengthElectrodeAG(1, col), 0)
        d_coke = GetNumeric(arrCokeBreezeDiameterAG(1, col), 0)
        rho_coke = GetNumeric(arrCokeBreezeResistivityAG(1, col), 0)
        rho_mat = GetNumeric(arrResistivityMaterialAG(1, col), 0)

        If curEnd <> 0 Then
            Rz = (factorV * nomV / curEnd) - (pipeInputResistanceEndLife + wireRes)
            If Rz < 0 Then Rz = 0
        Else
            Rz = 0
        End If

        If iType <> 3 Then
            Select Case iType
                Case 1
                    If l_el > 0 And d_el > 0 And h > 0 And (4 * h - l_el) > 0 Then
                        Rp1 = (rho_soil / (2 * Application.pi() * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
                    End If
                Case 2
                    If l_el > 0 And d_el > 0 Then
                        Rp1 = (rho_soil / (2 * Application.pi() * l_el)) * Log(2 * l_el / d_el)
                    End If
                Case 4
                    If l_coke > 0 And d_coke > 0 And h > 0 And rho_coke > 0 And rho_soil > 0 And d_el > 0 Then
                        Rp1 = (rho_soil / (2 * Application.pi() * l_coke)) * (Log(2 * l_coke / d_coke) + Log((l_coke + Sqr(l_coke ^ 2 + 16 * h ^ 2)) / (4 * h)) + (rho_coke / rho_soil) * Log(d_coke / d_el))
                    End If
                Case 5
                    If l_el > 0 And d_el > 0 And h > 0 Then
                        Rp1 = (rho_soil / (Application.pi() * l_el)) * Log(l_el / Sqr(d_el * h))
                    End If
                Case 6
                    If l_coke > 0 And d_coke > 0 And h > 0 And rho_coke > 0 And rho_soil > 0 And d_el > 0 Then
                        Rp1 = (rho_soil / (Application.pi() * l_coke)) * (Log(l_coke / Sqr(d_coke * h)) + (rho_coke / (2 * rho_soil)) * Log(d_coke / d_el))
                    End If
                Case 7
                    If l_el > 0 And d_el > 0 And h > 0 And (4 * h - l_el) > 0 Then
                        Rp1 = (rho_layer / (2 * Application.pi() * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
                    End If
                Case 8
                    If l_coke > 0 And d_coke > 0 And rho_coke > 0 And rho_layer > 0 And d_el > 0 Then
                        S_el = Application.pi() * d_el ^ 2 / 4
                        If S_el > 0 Then
                            Rp1 = (l_coke * rho_mat) / (2 * S_el) + (rho_layer / (2 * Application.pi() * l_coke)) * (Log(4 * l_coke / d_coke) + (rho_coke / rho_layer) * Log(d_coke / d_el))
                        End If
                    End If
                Case 9
                    If l_coke > 0 And d_coke > 0 And rho_coke > 0 And rho_layer > 0 And d_el > 0 And h > 0 And (4 * h - l_coke) > 0 Then
                        S_el = Application.pi() * d_el ^ 2 / 4
                        If S_el > 0 Then
                            Rp1 = (l_coke * rho_mat) / (2 * S_el) + (rho_layer / (2 * Application.pi() * l_coke)) * (Log(2 * l_coke / d_coke) + 0.5 * Log((4 * h + l_coke) / (4 * h - l_coke)) + (rho_coke / rho_layer) * Log(d_coke / d_el))
                        End If
                    End If
            End Select
        End If

        If iType = 3 Then
            R_vert = 0: R_horiz = 0
            If l_el > 0 And d_el > 0 And h > 0 And (4 * h - l_el) > 0 Then
                R_vert = (rho_soil / (2 * Application.pi() * l_el)) * (Log(2 * l_el / d_el) + 0.5 * Log((4 * h + l_el) / (4 * h - l_el)))
            End If
            If l_el > 0 And d_el > 0 Then
                R_horiz = (rho_soil / (2 * Application.pi() * l_el)) * Log(2 * l_el / d_el)
            End If
            If factorV > 0 And Rz > 0 And R_horiz > 0 Then
                n = Application.RoundUp((1.19 * R_vert * R_horiz - 0.98 * Rz * R_vert) / (factorV * Rz * R_horiz), 0)
            Else
                n = 0
            End If
        Else
            If factorV > 0 And Rz > 0 And Rp1 > 0 Then
                Select Case formulaType
                    Case "simple": n = Application.RoundUp(Rp1 / (factorV * Rz), 0)
                    Case "extended": n = Application.RoundUp(Rp1 / (factorV * Rz), 2)
                    Case Else: n = 0
                End Select
            Else
                n = 0
            End If
        End If

        ' write single-electrode resistance
        If iType = 3 Then
            ' комбинированный: на лист пишем Rp1 вертикального и горизонтального электродов
            ' combined: write Rp1 of vertical and horizontal electrodes to the sheet
            wsAnod.Range("oneElectrodeResistanceAG").Cells(1, col).Value = R_vert
            wsAnod.Range("oneElectrodeResistanceHorizAG").Cells(1, col).Value = R_horiz
        Else
            wsAnod.Range("oneElectrodeResistanceAG").Cells(1, col).Value = Rp1
        End If

        massOne = GetNumeric(arrMassOneElectrodeAG(1, col), 0)
        specMass = GetNumeric(arrSpecificMaccOneMeterAG(1, col), 0)
        If iType = 5 Or iType = 6 Then
            mass = specMass * l_el * n
        Else
            mass = massOne * n
        End If

        factorUse = GetNumeric(arrFactorUseMassAG(1, col), 0)
        dissolRate = GetNumeric(arrDissolutionRateAG(1, col), 0)
        avgCur = GetNumeric(arrAvgProtectionCurrentCPOverLife(1, col), 0)
        factorHet = GetNumeric(arrFactorSoilHeterogeneity(1, col), 0)

        If iType = 5 Or iType = 6 Or iType = 7 Or iType = 8 Or iType = 9 Then
            denom = dissolRate * avgCur * factorHet
        Else
            denom = dissolRate * avgCur
        End If
        If denom > 0 And mass > 0 Then
            life = Application.RoundDown((mass * factorUse) / denom, 0)
        Else
            life = 0
        End If

        If serviceLifeDesignedMax > 0 Then
            kt = (serviceLifeDesignedMax - life) / serviceLifeDesignedMax
        Else
            kt = 0
        End If
        Rp1prime = n * factorV * Rz

        wsAnod.Range("numElectrodesAG").Cells(1, col).Value = n
        wsAnod.Range("weightWithoutFillingAG").Cells(1, col).Value = mass
        wsAnod.Range("serviceLifeAG").Cells(1, col).Value = life
        wsAnod.Range("serviceLifeDeviation").Cells(1, col).Value = kt
        wsAnod.Range("correctResistanceAG").Cells(1, col).Value = Rp1prime

        If iType >= 7 And iType <= 9 Then
            If Rz > 0 Then
                lz = 3.5 * rho_layer / (Application.pi() * Rz)
            Else
                lz = 0
            End If
            wsAnod.Range("lengthWorkPartDeepAG").Cells(1, col).Value = lz
        End If

NextCol:
    Next col

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ================================================================
' пересчёт производных (G, T, kт, Rp1') по текущему N, все колонки
' recalculate derived values from current N, all columns
' ================================================================
Public Sub RecalcDerivedForType(relativeRow As Long)
    On Error GoTo CleanExit
    Dim pipeCount As Long
    Dim wsAnod As Worksheet
    Set wsAnod = thisWorkbook.Worksheets("Anod")
    
    With wsAnod.Range("PipeCountCP")
        If IsEmpty(.Value) Or Not IsNumeric(.Value) Then
            ' pipecountcp пуст или не число.
            ' pipecountcp is empty or not a number.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 043F 0443 0441 0442 0020 0438 043B 0438 0020 043D 0435 0020 0447 0438 0441 043B 043E 002E")
            Exit Sub
        End If
        pipeCount = CLng(.Value)
        If pipeCount <= 0 Or pipeCount > MAX_COL Then
            ' pipecountcp должен быть от 1 до 254.
            ' pipecountcp must be from 1 to 254.
            MsgBox Ru("0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043E 0442 0020 0031 0020 0434 043E 0020") & MAX_COL & "."
            Exit Sub
        End If
    End With

    ' load required arrays
    Dim arrFactorVoltageMarginCP As Variant
    Dim arrMassOneElectrodeAG As Variant
    Dim arrSpecificMaccOneMeterAG As Variant
    Dim arrFactorUseMassAG As Variant
    Dim arrDissolutionRateAG As Variant
    Dim arrAvgProtectionCurrentCPOverLife As Variant
    Dim arrFactorSoilHeterogeneity As Variant
    Dim arrLengthElectrodeAG As Variant

    On Error Resume Next
    arrFactorVoltageMarginCP = SafeGetArray(wsAnod.Range("factorVoltageMarginCP"))
    arrMassOneElectrodeAG = SafeGetArray(wsAnod.Range("massOneElectrodeAG"))
    arrSpecificMaccOneMeterAG = SafeGetArray(wsAnod.Range("specificMaccOneMeterAG"))
    arrFactorUseMassAG = SafeGetArray(wsAnod.Range("factorUseMassAG"))
    arrDissolutionRateAG = SafeGetArray(wsAnod.Range("dissolutionRateAG"))
    arrAvgProtectionCurrentCPOverLife = SafeGetArray(wsAnod.Range("avgProtectionCurrentCPOverLife"))
    arrFactorSoilHeterogeneity = SafeGetArray(wsAnod.Range("factorSoilHeterogeneity"))
    arrLengthElectrodeAG = SafeGetArray(wsAnod.Range("lengthElectrodeAG"))
    If Err.Number <> 0 Then
        ' ошибка чтения исходных данных: 
        ' read error input data: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0438 0441 0445 043E 0434 043D 044B 0445 0020 0434 0430 043D 043D 044B 0445 003A 0020") & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' dimension checks
    If UBound(arrFactorVoltageMarginCP, 2) < pipeCount Then CheckFailed "factorVoltageMarginCP"
    If UBound(arrMassOneElectrodeAG, 2) < pipeCount Then CheckFailed "massOneElectrodeAG"
    If UBound(arrSpecificMaccOneMeterAG, 2) < pipeCount Then CheckFailed "specificMaccOneMeterAG"
    If UBound(arrFactorUseMassAG, 2) < pipeCount Then CheckFailed "factorUseMassAG"
    If UBound(arrDissolutionRateAG, 2) < pipeCount Then CheckFailed "dissolutionRateAG"
    If UBound(arrAvgProtectionCurrentCPOverLife, 2) < pipeCount Then CheckFailed "avgProtectionCurrentCPOverLife"
    If UBound(arrFactorSoilHeterogeneity, 2) < pipeCount Then CheckFailed "factorSoilHeterogeneity"
    If UBound(arrLengthElectrodeAG, 2) < pipeCount Then CheckFailed "lengthElectrodeAG"

    ' проверка однострочных диапазонов: 1 x pipeCountCP
    ' check single-row ranges: 1 x pipeCountCP
    ' numelectrodesag (строки: ожидается 1)
    ' numelectrodesag (rows: expected 1)
    If wsAnod.Range("numElectrodesAG").rows.count <> 1 Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020") & Ru("0031 0029")
    ' numelectrodesag (столбцы: ожидается pipecountcp)
    ' numelectrodesag (columns: expected pipecountcp)
    If wsAnod.Range("numElectrodesAG").Columns.count <> pipeCount Then CheckFailed Ru("006E 0075 006D 0045 006C 0065 0063 0074 0072 006F 0064 0065 0073 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F") & Ru("0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' weightwithoutfillingag (строки: ожидается 1)
    ' weightwithoutfillingag (rows: expected 1)
    If wsAnod.Range("weightWithoutFillingAG").rows.count <> 1 Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438") & Ru("0434 0430 0435 0442 0441 044F 0020 0031 0029")
    ' weightwithoutfillingag (столбцы: ожидается pipecountcp)
    ' weightwithoutfillingag (columns: expected pipecountcp)
    If wsAnod.Range("weightWithoutFillingAG").Columns.count <> pipeCount Then CheckFailed Ru("0077 0065 0069 0067 0068 0074 0057 0069 0074 0068 006F 0075 0074 0046 0069 006C 006C 0069 006E 0067 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436") & Ru("0438 0434 0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifeag (строки: ожидается 1)
    ' servicelifeag (rows: expected 1)
    If wsAnod.Range("serviceLifeAG").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0031") & Ru("0029")
    ' servicelifeag (столбцы: ожидается pipecountcp)
    ' servicelifeag (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeAG").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430 0435 0442 0441 044F 0020 0050") & Ru("0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' servicelifedeviation (строки: ожидается 1)
    ' servicelifedeviation (rows: expected 1)
    If wsAnod.Range("serviceLifeDeviation").rows.count <> 1 Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0031 0029")
    ' servicelifedeviation (столбцы: ожидается pipecountcp)
    ' servicelifedeviation (columns: expected pipecountcp)
    If wsAnod.Range("serviceLifeDeviation").Columns.count <> pipeCount Then CheckFailed Ru("0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0076 0069 0061 0074 0069 006F 006E 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434") & Ru("0430 0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")
    ' correctresistanceag (строки: ожидается 1)
    ' correctresistanceag (rows: expected 1)
    If wsAnod.Range("correctResistanceAG").rows.count <> 1 Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 0440 043E 043A 0438 003A 0020 043E 0436 0438 0434 0430 0435") & Ru("0442 0441 044F 0020 0031 0029")
    ' correctresistanceag (столбцы: ожидается pipecountcp)
    ' correctresistanceag (columns: expected pipecountcp)
    If wsAnod.Range("correctResistanceAG").Columns.count <> pipeCount Then CheckFailed Ru("0063 006F 0072 0072 0065 0063 0074 0052 0065 0073 0069 0073 0074 0061 006E 0063 0065 0041 0047 0020 0028 0441 0442 043E 043B 0431 0446 044B 003A 0020 043E 0436 0438 0434 0430") & Ru("0435 0442 0441 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0029")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearRangePart wsAnod.Range("weightWithoutFillingAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("serviceLifeAG"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("serviceLifeDeviation"), 1, 1, 1, pipeCount
    ClearRangePart wsAnod.Range("correctResistanceAG"), 1, 1, 1, pipeCount

    ' read serviceLifeDesignedMax
    Dim wsPipe As Worksheet
    Dim serviceLifeDesignedMax As Double
    Dim arrServiceLifeDesigned As Variant
    Dim rngServiceLife As Range
    Dim serviceLifeVal As Variant

    On Error Resume Next
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If Not wsPipe Is Nothing Then
        Set rngServiceLife = wsPipe.Range("serviceLifeDesigned")
        If Not rngServiceLife Is Nothing Then
            serviceLifeVal = rngServiceLife.Value
            If Not IsArray(serviceLifeVal) Then
                ReDim arrServiceLifeDesigned(1 To 1, 1 To 1)
                arrServiceLifeDesigned(1, 1) = serviceLifeVal
            Else
                arrServiceLifeDesigned = ReadRangeAs1xN(rngServiceLife)
            End If
            serviceLifeDesignedMax = GetMaxFrom1xN(arrServiceLifeDesigned, 30)
        Else
            serviceLifeDesignedMax = 30
        End If
    Else
        serviceLifeDesignedMax = 30
    End If
    If Err.Number <> 0 Then
        ' ошибка чтения servicelifedesigned: 
        ' read error servicelifedesigned: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0447 0442 0435 043D 0438 044F 0020 0073 0065 0072 0076 0069 0063 0065 004C 0069 0066 0065 0044 0065 0073 0069 0067 006E 0065 0064 003A 0020") & Err.Description
        GoTo CleanExit
    End If
    On Error GoTo 0
    If serviceLifeDesignedMax <= 0 Then serviceLifeDesignedMax = 30

    Dim col As Long
    Dim cellType As Variant, typeStr As String
    Dim iType As Integer
    Dim n As Double, mass As Double, life As Double, kt As Double, Rp1prime As Double
    Dim massOne As Double, specMass As Double
    Dim factorV As Double
    Dim factorUse As Double, dissolRate As Double, avgCur As Double, factorHet As Double
    Dim denom As Double
    Dim l_el As Double, Rz As Double

    For col = 1 To pipeCount
        cellType = wsAnod.Range("typeInstallationAG").Cells(1, col).Value
        If IsEmpty(cellType) Then GoTo NextCol
        iType = Module_Visual.GetInstallationTypeIndex(CStr(cellType))
        If iType = 0 Then GoTo NextCol
        n = GetNumeric(wsAnod.Range("numElectrodesAG").Cells(1, col).Value, 0)
        If n <= 0 Then GoTo NextCol

        factorV = GetNumeric(arrFactorVoltageMarginCP(1, col), 0)
        massOne = GetNumeric(arrMassOneElectrodeAG(1, col), 0)
        specMass = GetNumeric(arrSpecificMaccOneMeterAG(1, col), 0)
        factorUse = GetNumeric(arrFactorUseMassAG(1, col), 0)
        dissolRate = GetNumeric(arrDissolutionRateAG(1, col), 0)
        avgCur = GetNumeric(arrAvgProtectionCurrentCPOverLife(1, col), 0)
        factorHet = GetNumeric(arrFactorSoilHeterogeneity(1, col), 0)
        l_el = GetNumeric(arrLengthElectrodeAG(1, col), 0)
        Rz = GetNumeric(wsAnod.Range("resistanceEndLifeAG").Cells(1, col).Value, 0)

        If iType = 5 Or iType = 6 Then
            mass = specMass * l_el * n
        Else
            mass = massOne * n
        End If

        If iType = 5 Or iType = 6 Or iType = 7 Or iType = 8 Or iType = 9 Then
            denom = dissolRate * avgCur * factorHet
        Else
            denom = dissolRate * avgCur
        End If
        If denom > 0 And mass > 0 Then
            life = Application.RoundDown((mass * factorUse) / denom, 0)
        Else
            life = 0
        End If

        If serviceLifeDesignedMax > 0 Then
            kt = (serviceLifeDesignedMax - life) / serviceLifeDesignedMax
        Else
            kt = 0
        End If
        Rp1prime = n * factorV * Rz

        wsAnod.Range("weightWithoutFillingAG").Cells(1, col).Value = mass
        wsAnod.Range("serviceLifeAG").Cells(1, col).Value = life
        wsAnod.Range("serviceLifeDeviation").Cells(1, col).Value = kt
        wsAnod.Range("correctResistanceAG").Cells(1, col).Value = Rp1prime

NextCol:
    Next col

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ================================================================
' кнопка: пересчёт после изменения количества электродов (N)
' button: recalc after changing the electrode count (N)
' ================================================================
Sub btnRecalcDerivedType1(): Call RecalcDerivedForType(1): End Sub

' ================================================================
' helper functions
' ================================================================
Private Sub Write1xN(ByVal rng As Range, ByRef arr As Variant, ByVal nCols As Long)
    If rng Is Nothing Then Exit Sub
    If nCols < 1 Then Exit Sub
    rng.Cells(1, 1).Resize(1, nCols).Value = arr
End Sub

Private Function SafeGetArray(rng As Range) As Variant
    If rng Is Nothing Then Err.Raise vbObjectError + 1000, , "Range is Nothing"
    If rng.Cells.count = 1 Then
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = rng.Value
        SafeGetArray = arr
    Else
        SafeGetArray = rng.Value
    End If
End Function

Private Function ReadRangeAs1xN(rng As Range) As Variant
    Dim v As Variant
    v = rng.Value
    Dim rowsCnt As Long, colsCnt As Long
    If Not IsArray(v) Then
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = v
        ReadRangeAs1xN = arr
        Exit Function
    End If
    rowsCnt = UBound(v, 1)
    colsCnt = UBound(v, 2)
    Dim out() As Variant
    ReDim out(1 To 1, 1 To rowsCnt * colsCnt)
    Dim i As Long, j As Long, idx As Long
    idx = 1
    For i = 1 To rowsCnt
        For j = 1 To colsCnt
            out(1, idx) = v(i, j)
            idx = idx + 1
        Next
    Next
    ReadRangeAs1xN = out
End Function

Private Function GetMaxFrom1xN(arr As Variant, defaultValue As Double) As Double
    Dim maxVal As Double
    maxVal = defaultValue
    If Not IsArray(arr) Then
        If IsNumeric(arr) Then maxVal = CDbl(arr)
        GetMaxFrom1xN = maxVal
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        If IsNumeric(arr(1, i)) Then
            If arr(1, i) > maxVal Then maxVal = CDbl(arr(1, i))
        End If
    Next
    GetMaxFrom1xN = maxVal
End Function

Private Function GetNumeric(v As Variant, Optional defaultVal As Double = 0) As Double
    If IsNumeric(v) Then
        GetNumeric = CDbl(v)
    Else
        GetNumeric = defaultVal
    End If
End Function

Private Sub ClearRangePart(rng As Range, rowStart As Long, rowEnd As Long, colStart As Long, colEnd As Long)
    If rng Is Nothing Then Exit Sub
    Dim clearRows As Long, clearCols As Long
    clearRows = rowEnd - rowStart + 1
    clearCols = colEnd - colStart + 1
    If clearRows < 1 Or clearCols < 1 Then Exit Sub
    rng.Cells(rowStart, colStart).Resize(clearRows, clearCols).ClearContents
End Sub

Private Sub CheckFailed(ByVal rangeName As String)
    Dim pipeCount As Variant
    On Error Resume Next
    pipeCount = thisWorkbook.Worksheets("Anod").Range("PipeCountCP").Value
    If Err.Number <> 0 Then pipeCount = "?"
    On Error GoTo 0
    Debug.Print "!!! CheckFailed: " & rangeName & " (PipeCountCP = " & pipeCount & ")"
    ' диапазон '
    ' range '
    ' ' имеет недостаточную размерность для pipecountcp = 
    ' ' andмеет insufficient size for pipecountcp = 
    MsgBox Ru("0434 0438 0430 043F 0430 0437 043E 043D 0020 0027") & rangeName & Ru("0027 0020 0438 043C 0435 0435 0442 0020 043D 0435 0434 043E 0441 0442 0430 0442 043E 0447 043D 0443 044E 0020 0440 0430 0437 043C 0435 0440 043D 043E 0441 0442 044C 0020 0434") & Ru("043B 044F 0020 0050 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 003D 0020") & pipeCount & "."
    End
End Sub
