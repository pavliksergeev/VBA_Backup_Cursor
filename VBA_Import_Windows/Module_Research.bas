Attribute VB_Name = "Module_Research"

' ================================================================
' RESEARCH module: import from the Research Data sheet
' ================================================================
Option Explicit

' ================================================================
' insert pipe length from survey data
' ================================================================
Sub btnImportLength()
    On Error GoTo CleanExit
    
    Dim wsResearch As Worksheet
    Set wsResearch = thisWorkbook.Worksheets(SHEET_RESEARCH)
    
    If wsResearch Is Nothing Then
        ' лист 'research data' не найден!
        ' лandст 'research data' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027 0052 0065 0073 0065 0061 0072 0063 0068 0020 0064 0061 0074 0061 0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        Exit Sub
    End If
    
    If IsEmpty(wsResearch.Range("B3").Value) Then
        ' ячейка b3 на листе 'research data' пуста!
        ' cell b3 on sheet 'research data' is empty!
        MsgBox Ru("044F 0447 0435 0439 043A 0430 0020 0042 0033 0020 043D 0430 0020 043B 0438 0441 0442 0435 0020 0027 0052 0065 0073 0065 0061 0072 0063 0068 0020 0064 0061 0074 0061 0027 0020") & Ru("043F 0443 0441 0442 0430 0021"), vbExclamation
        Exit Sub
    End If
    
    ' note: extra space removed
    Range("pipeLength").Value = wsResearch.Range("B3").Value
    
    ' в разработке: тест-длина трубы скопирована: 
    ' under development: тест-длandon трубы скопandроваon: 
    '  м
    MsgBox Ru("0432 0020 0440 0430 0437 0440 0430 0431 043E 0442 043A 0435 003A 0020 0442 0435 0441 0442 002D 0434 043B 0438 043D 0430 0020 0442 0440 0443 0431 044B 0020 0441 043A 043E 043F") & Ru("0438 0440 043E 0432 0430 043D 0430 003A 0020") & Range("pipeLength").Value & Ru("0020 043C"), vbInformation

CleanExit:
End Sub

' ================================================================
' import the soil-heterogeneity factor
' ================================================================
Sub btnImportSoilHeterogeneity()
    ' в разработке импорт данных изысканий и расчет коэффициента неоднородности грунта
    ' under development survey data import and calc soil heterogeneity factor
    MsgBox Ru("0432 0020 0440 0430 0437 0440 0430 0431 043E 0442 043A 0435 0020 0438 043C 043F 043E 0440 0442 0020 0434 0430 043D 043D 044B 0445 0020 0438 0437 044B 0441 043A 0430 043D 0438") & Ru("0439 0020 0438 0020 0440 0430 0441 0447 0435 0442 0020 043A 043E 044D 0444 0444 0438 0446 0438 0435 043D 0442 0430 0020 043D 0435 043E 0434 043D 043E 0440 043E 0434 043D 043E") & Ru("0441 0442 0438 0020 0433 0440 0443 043D 0442 0430")
'    On Error GoTo CleanExit
'
'    Dim originalCalc As XlCalculation
'    originalCalc = Application.Calculation
'    Dim wsResearch As Worksheet
'    Set wsResearch = ThisWorkbook.Worksheets(SHEET_RESEARCH)
'
'    If wsResearch Is Nothing Then
'       MsgBox "sheet '" & SHEET_RESEARCH & "' not found", vbExclamation
'        Exit Sub
'    End If
'
'    If IsEmpty(wsResearch.Range("calcSoilHeterogeneity").Value) Then
'        MsgBox "range 'calcSoilHeterogeneity' is empty", vbExclamation
'        Exit Sub
'    End If
'
'    Range("factorSoilHeterogeneity").Value = wsResearch.Range("calcSoilHeterogeneity").Value
'
'CleanExit:
'    Application.Calculation = originalCalc
End Sub

' ================================================================
' import resistivity of the i-th soil layer
' ================================================================
Sub btnImportResistivityLayer()
    ' в разработке: импорт данных изысканий и расчет удельного электрического сопротивления i-того слоя земли, в котором располагается глубинный заземлитель (ri, ом*м)
    ' under development: survey data import and calc удельного электрandческого сопротandвленandя i-th soil layer, where the following is placed deep anode (ri, ом*м)
    MsgBox Ru("0432 0020 0440 0430 0437 0440 0430 0431 043E 0442 043A 0435 003A 0020 0438 043C 043F 043E 0440 0442 0020 0434 0430 043D 043D 044B 0445 0020 0438 0437 044B 0441 043A 0430 043D") & Ru("0438 0439 0020 0438 0020 0440 0430 0441 0447 0435 0442 0020 0443 0434 0435 043B 044C 043D 043E 0433 043E 0020 044D 043B 0435 043A 0442 0440 0438 0447 0435 0441 043A 043E 0433") & Ru("043E 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 044F 0020 0069 002D 0442 043E 0433 043E 0020 0441 043B 043E 044F 0020 0437 0435 043C 043B 0438 002C 0020") & Ru("0432 0020 043A 043E 0442 043E 0440 043E 043C 0020 0440 0430 0441 043F 043E 043B 0430 0433 0430 0435 0442 0441 044F 0020 0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0437") & Ru("0430 0437 0435 043C 043B 0438 0442 0435 043B 044C 0020 0028 0072 0069 002C 0020 043E 043C 002A 043C 0029")
'  On Error GoTo CleanExit
'
'  Dim originalCalc As XlCalculation
'  originalCalc = Application.Calculation
'
'  Dim wsResearch As Worksheet
'  Set wsResearch = ThisWorkbook.Worksheets(SHEET_RESEARCH)
'
'  If wsResearch Is Nothing Then
'      MsgBox "sheet 'Research data' not found", vbExclamation
'      Exit Sub
'  End If
'
'  If IsEmpty(wsResearch.Range("resistivity_ilayerDeepAG").Value) Then
'      MsgBox "range 'resistivity_ilayerDeepAG' is empty", vbExclamation
'      Exit Sub
'  End If
'
'  Range("resistivity_i_layerDeepAG").Value = wsResearch.Range("resistivity_ilayerDeepAG").Value
'
'CleanExit:
'    Application.Calculation = originalCalc
End Sub

' ================================================================
' (stub - not implemented yet)
' ================================================================
Sub button126_click()
    ' в разработке
    ' under development
    MsgBox Ru("0432 0020 0440 0430 0437 0440 0430 0431 043E 0442 043A 0435")
'    On Error GoTo CleanExit
'
'    Dim colBr As Long
'    Dim pipeCountValue As Variant
'
'    On Error Resume Next
'    pipeCountValue = Range("pipeDifferentParametersNum").Value
'    If Err.Number <> 0 Then pipeCountValue = 1
'    On Error GoTo 0
'
'    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
'        colBr = CLng(pipeCountValue)
'    Else
'        colBr = 1
'    End If
'    If colBr < 1 Then colBr = 1
'    If colBr > 16384 Then colBr = 16384
'
'    Dim i As Long
'    For i = 1 To colBr
'        On Error Resume Next
'        Dim srcVal As Variant
'        srcVal = Range("pipeWallThicknessAvg" & i).Value
'        If Err.Number = 0 Then
'            If Not IsEmpty(srcVal) And IsNumeric(srcVal) Then
'                Range("pipeWallThickness").Cells(1, i).Value = srcVal
'            End If
'        End If
'        On Error GoTo 0
'    Next i
'
'CleanExit:
End Sub

' ================================================================
' calculate mean soil resistivity from survey rows (popup form)
' writes to soilResistivityAvg of the active Pipe section
' ================================================================
Sub button112_click()
    Call ShowSoilAvgForm(SOILAVG_TARGET_PIPE)
End Sub
