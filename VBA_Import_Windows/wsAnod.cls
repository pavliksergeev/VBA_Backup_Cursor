VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsAnod"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ================================================================
' wsAnod Sheet Module
' Events and thin handlers only
' ================================================================

Private isUpdating As Boolean
Public isProgrammaticPipeCountChange As Boolean   ' <-- HERE
Private previousPipeCount As Variant

Private Sub Worksheet_Activate()
    ' No updates on activation
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If isUpdating Then Exit Sub
    isUpdating = True
    
    ' Ignore changes outside filters and key cells
    If Target.row < 34 Or Target.row > 38 Then
        If Intersect(Target, Me.Range("pipeCountCP")) Is Nothing And _
           Intersect(Target, Me.Range("maxProtectPotential")) Is Nothing Then
            isUpdating = False
            Exit Sub
        End If
    End If
    
    On Error GoTo CleanExit
    
    Debug.Print "Worksheet_Change: Target.Address = " & Target.address
    Debug.Print "Worksheet_Change: Target.CountLarge = " & Target.CountLarge
    Debug.Print "Worksheet_Change: Target.Column = " & Target.Column
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' ================================================================
    ' 1. Change of pipeCountCP
    ' ================================================================
    If Not Intersect(Target, Me.Range("pipeCountCP")) Is Nothing Or Not Intersect(Target, Me.Range("D19")) Is Nothing Then
        
        Dim newPipeCount As Variant
        newPipeCount = Me.Range("pipeCountCP").Value
        
         If IsNumeric(newPipeCount) And newPipeCount > 0 Then
            ' Round to integer if the value is fractional
            Dim newPipeCountInt As Long
            newPipeCountInt = CLng(Application.WorksheetFunction.Round(newPipeCount, 0))
            
            ' If rounding changed the value, write it back
            If newPipeCountInt <> CDbl(newPipeCount) Then
                Application.EnableEvents = False
                Me.Range("pipeCountCP").Value = newPipeCountInt
                Application.EnableEvents = True
                newPipeCount = newPipeCountInt
            End If
            
            ' If the rounded value matches the previous one, do nothing
    '        If Not IsEmpty(previousPipeCount) Then
    '            If previousPipeCount = newPipeCountInt Then
    '                GoTo CleanExit
    '            End If
    '        End If
            
            ' Remember the new value
            previousPipeCount = newPipeCountInt
            
            ' First recalculate max protective potential (manual change only)
            If Not isProgrammaticPipeCountChange Then
                Dim newMaxPot As Variant
                On Error Resume Next
                newMaxPot = Module_calcCP.CalcMaxProtectPotential()
                Debug.Print "newMaxPot = " & newMaxPot
                If IsNumeric(newMaxPot) And newMaxPot <> 0 Then
                    Application.EnableEvents = False
                    Me.Range("maxProtectPotential").Value = newMaxPot
                    Application.EnableEvents = True
                    ' >>> новое значение maxprotectpotential = 
                    ' >>> new value maxprotectpotential = 
                    Debug.Print Ru("003E 003E 003E 0020 041D 043E 0432 043E 0435 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020 006D 0061 0078 0050 0072 006F 0074 0065 0063 0074 0050 006F 0074 0065 006E 0074") & Ru("0069 0061 006C 0020 003D 0020") & newMaxPot
                End If
                On Error GoTo 0
            End If
            
            ' Then sync columns
            Call HandlePipeCountChange(newPipeCountInt)
        Else
            ' pipecountcp должен быть положительным числом!
            ' pipecountcp must be a positive number!
            MsgBox Ru("0070 0069 0070 0065 0063 006F 0075 006E 0074 0063 0070 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 043F 043E 043B 043E 0436 0438 0442 0435 043B 044C 043D") & Ru("044B 043C 0020 0447 0438 0441 043B 043E 043C 0021"), vbExclamation
            Application.EnableEvents = False
            Me.Range("pipeCountCP").Value = 1
            Application.EnableEvents = True
        End If
        
        GoTo CleanExit
    End If
    
    
    
    ' ================================================================
    ' 2. Change of maxProtectPotential
    ' ================================================================
    If Not Intersect(Target, Me.Range("maxProtectPotential")) Is Nothing Then
        Call HandleMaxProtectPotentialChange(Target)
        GoTo CleanExit
    End If
    
    ' ================================================================
    ' 3. Change in filter rows (34-38)
    ' ================================================================
    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = SafeRangeValue(Me.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)
    
    Dim startCol As Long
    startCol = START_COL
    Dim endCol As Long
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, Application.WorksheetFunction.Min(MAX_COL, endCol))
    
    Dim filterRange As Range
    Set filterRange = Me.Range(Me.Cells(FILTER_START_ROW, startCol), Me.Cells(FILTER_END_ROW, endCol))
    
    Dim filterIntersect As Range
    Set filterIntersect = Intersect(Target, filterRange)
    
    If Not filterIntersect Is Nothing Then
        If Target.CountLarge = 1 Then
            ' Single cell change - process only that column
            Dim changedCol As Long
            changedCol = Target.Column
            Debug.Print "Worksheet_Change: single cell, changedCol = " & changedCol
            If changedCol >= startCol And changedCol <= endCol Then
                Dim modelChanged As Boolean
                modelChanged = (Target.row = FILTER_END_ROW)
                Call ProcessFilterChange(Me, changedCol, modelChanged)
            End If
        Else
            ' Multiple cells change - process each affected column separately
            Debug.Print "Worksheet_Change: multiple cells"
            Dim colIdx As Long
            For colIdx = startCol To endCol
                If Not Intersect(Target, Me.Range(Me.Cells(FILTER_START_ROW, colIdx), Me.Cells(FILTER_END_ROW, colIdx))) Is Nothing Then
                    modelChanged = (Not Intersect(Target, Me.Cells(FILTER_END_ROW, colIdx)) Is Nothing)
                    Debug.Print "Worksheet_Change: processing colIdx = " & colIdx & ", modelChanged = " & modelChanged
                    Call ProcessFilterChange(Me, colIdx, modelChanged)
                End If
            Next colIdx
        End If
        
        ' Force-refresh validation for neighboring columns
        Dim colIdxAll As Long
        For colIdxAll = startCol To endCol
            If colIdxAll <> Target.Column Then
                Call Module_ValidationLogic.RefreshValidationForColumn(Me, colIdxAll)
            End If
        Next colIdxAll
        
        GoTo CleanExit
    End If
    ' ================================================================
    ' 4. Change of CP model (row 26)
    ' ================================================================
    Dim cpRange As Range
    Set cpRange = Me.Range(Me.Cells(CP_MODEL_ROW, startCol), Me.Cells(CP_MODEL_ROW, endCol))
    Dim cpIntersect As Range
    Set cpIntersect = Intersect(Target, cpRange)
    
    If Not cpIntersect Is Nothing And Target.CountLarge = 1 Then
        Dim cpCell As Range
        For Each cpCell In cpIntersect
            Dim cpCol As Long
            cpCol = cpCell.Column
            Dim cpValue As Variant
            cpValue = Me.Cells(CP_MODEL_ROW, cpCol).Value
            
            Dim rw As Long
            For rw = CP_HYPERLINK_ROW To CP_POWER_ROW
                Me.Cells(rw, cpCol).ClearContents
            Next rw
            
            If Not IsEmpty(cpValue) And cpValue <> "" Then
                Call UpdateCPData(Me, cpCol)
            End If
            
            Call Module_ValidationLogic.RefreshValidationForColumn(Me, cpCol)
        Next cpCell
        
        Me.Calculate
        GoTo CleanExit
    End If
    
    ' ================================================================
    ' 5. Change of CP current/voltage (rows 29-30)
    ' ================================================================
    Dim cpPowerRange As Range
    Set cpPowerRange = Me.Range(Me.Cells(CP_CURRENT_ROW, startCol), Me.Cells(CP_VOLTAGE_ROW, endCol))
    Dim cpPowerIntersect As Range
    Set cpPowerIntersect = Intersect(Target, cpPowerRange)
    
    If Not cpPowerIntersect Is Nothing Then
        Dim cpPowerCol As Long
        cpPowerCol = Target.Column
        If Target.count > 1 Then cpPowerCol = Target.Cells(1, 1).Column
        
        Dim currentVal As Variant
        Dim voltageVal As Variant
        currentVal = Me.Cells(CP_CURRENT_ROW, cpPowerCol).Value
        voltageVal = Me.Cells(CP_VOLTAGE_ROW, cpPowerCol).Value
        
        If IsNumeric(currentVal) And IsNumeric(voltageVal) Then
            Me.Cells(CP_POWER_ROW, cpPowerCol).Value = currentVal * voltageVal
        End If
        
        GoTo CleanExit
    End If
    
    ' ================================================================
    ' 6. Other changes - just recalculate
    ' ================================================================
    Me.Calculate

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    isUpdating = False
End Sub

' ================================================================
' Private handlers
' ================================================================

' Handles change of pipeCountCP
Private Sub HandlePipeCountChange(ByVal newCount As Long)
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Set ws = Me
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call Module_Sync.SyncAnodColumns(ws, newCount)
    Call Module_Sync.UpdateNamedRanges(ws, newCount)
    
    Dim colIdx As Long
    For colIdx = START_COL To START_COL + newCount - 1
        Call Module_ValidationLogic.RefreshValidationForColumn(ws, colIdx)
    Next colIdx
    
    Call Module_Visual.UpdateVisibilityForAllColumns(ws)
    Call Module_Visual.UpdateColorsForAllColumns(ws)
    
    ws.Calculate
    
CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' Handles change of maxProtectPotential
Private Sub HandleMaxProtectPotentialChange(ByVal Target As Range)
    Dim newVal As Variant
    newVal = Target.Value
    
    If IsNumeric(newVal) Then
        Dim limitedVal As Double
        
        If newVal < Module_Constants.MIN_PROTECT_POTENTIAL Then
            limitedVal = Module_Constants.MIN_PROTECT_POTENTIAL
            ' недопустимый защитный потенциал: введено значение 
            ' invalid protection potential: entered value 
            '  в
            ' установлено минимальное допустимое значение 
            ' set to minimum allowed value 
            '  в.
            ' ограничение потенциала
            ' potential limit
            MsgBox Ru("043D 0435 0434 043E 043F 0443 0441 0442 0438 043C 044B 0439 0020 0437 0430 0449 0438 0442 043D 044B 0439 0020 043F 043E 0442 0435 043D 0446 0438 0430 043B 003A 0020 0432 0432") & Ru("0435 0434 0435 043D 043E 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020") & Format(newVal, "0.000") & Ru("0020 0432") & vbCrLf & _
                   Ru("0443 0441 0442 0430 043D 043E 0432 043B 0435 043D 043E 0020 043C 0438 043D 0438 043C 0430 043B 044C 043D 043E 0435 0020 0434 043E 043F 0443 0441 0442 0438 043C 043E 0435 0020") & Ru("0437 043D 0430 0447 0435 043D 0438 0435 0020") & Format(Module_Constants.MIN_PROTECT_POTENTIAL, "0.000") & Ru("0020 0432 002E"), _
                   vbExclamation, Ru("043E 0433 0440 0430 043D 0438 0447 0435 043D 0438 0435 0020 043F 043E 0442 0435 043D 0446 0438 0430 043B 0430")
            Application.EnableEvents = False
            Target.Value = limitedVal
            Application.EnableEvents = True
        ElseIf newVal > Module_Constants.MAX_PROTECT_POTENTIAL Then
            limitedVal = Module_Constants.MAX_PROTECT_POTENTIAL
            ' недопустимый защитный потенциал: введено значение 
            ' invalid protection potential: entered value 
            '  в
            ' установлено максимальное допустимое значение 
            ' set to maximum allowed value 
            '  в.
            ' ограничение потенциала
            ' potential limit
            MsgBox Ru("043D 0435 0434 043E 043F 0443 0441 0442 0438 043C 044B 0439 0020 0437 0430 0449 0438 0442 043D 044B 0439 0020 043F 043E 0442 0435 043D 0446 0438 0430 043B 003A 0020 0432 0432") & Ru("0435 0434 0435 043D 043E 0020 0437 043D 0430 0447 0435 043D 0438 0435 0020") & Format(newVal, "0.000") & Ru("0020 0432") & vbCrLf & _
                   Ru("0443 0441 0442 0430 043D 043E 0432 043B 0435 043D 043E 0020 043C 0430 043A 0441 0438 043C 0430 043B 044C 043D 043E 0435 0020 0434 043E 043F 0443 0441 0442 0438 043C 043E 0435") & Ru("0020 0437 043D 0430 0447 0435 043D 0438 0435 0020") & Format(Module_Constants.MAX_PROTECT_POTENTIAL, "0.000") & Ru("0020 0432 002E"), _
                   vbExclamation, Ru("043E 0433 0440 0430 043D 0438 0447 0435 043D 0438 0435 0020 043F 043E 0442 0435 043D 0446 0438 0430 043B 0430")
            Application.EnableEvents = False
            Target.Value = limitedVal
            Application.EnableEvents = True
        End If
    End If
End Sub

' Processes filter change for a single column
Private Sub ProcessFilterChange(ByVal ws As Worksheet, ByVal col As Long, ByVal modelChanged As Boolean)
    Debug.Print "ProcessFilterChange: col = " & col & ", modelChanged = " & modelChanged
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If modelChanged Then
        Dim modelVal As Variant
        modelVal = ws.Cells(FILTER_END_ROW, col).Value
        
        If IsEmpty(modelVal) Or modelVal = "" Then
            Call ClearFiltersInColumn(ws, col)
            Call Module_AGData.ClearCalculatedDataForCol(ws, col)
            Call Module_ValidationLogic.RefreshValidationForColumn(ws, col)
        Else
            Call Module_AGData.InsertAGDataForColumn(ws, col)
            
            If Not IsEmpty(ws.Cells(CP_MODEL_ROW, col).Value) Then
                Call UpdateCPData(ws, col)
            End If
            
            Call Module_ValidationLogic.AutoFillFiltersFromModel(ws, col, modelVal)
            Call Module_ValidationLogic.UpdateAllFiltersCross(ws, col)
        End If
    Else
        Call Module_ValidationLogic.UpdateAllFiltersCross(ws, col)
        Call Module_ValidationLogic.RefreshValidationForColumn(ws, col)
    End If
    
    Call Module_Visual.UpdateVisibilityForAllColumns(ws)
    Call Module_Visual.UpdateColorsForAllColumns(ws)
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub btnCalcProtectiveZone()
    On Error GoTo CleanExit
    If Me Is Nothing Then Exit Sub
    On Error Resume Next
    isProgrammaticPipeCountChange = True
    Call Module_calcCP.CalcProtectiveZone(Me)
    isProgrammaticPipeCountChange = False
    ' ошибка расчета защитной зоны: 
    ' calculation error protective zone: 
    If Err.Number <> 0 Then MsgBox Ru("041E 0448 0438 0431 043A 0430 0020 0440 0430 0441 0447 0435 0442 0430 0020 0437 0430 0449 0438 0442 043D 043E 0439 0020 0437 043E 043D 044B 003A 0020") & Err.Description
    On Error GoTo CleanExit
CleanExit:
End Sub

Sub btnDistributeLength()
    On Error GoTo CleanExit
    If Me Is Nothing Then Exit Sub
    On Error Resume Next
    Call Module_calcCP.DistributeLength(Me)
    ' ошибка распрелеления укз& err.description
    ' error распрелеленandя укз& err.description
    If Err.Number <> 0 Then MsgBox Ru("041E 0448 0438 0431 043A 0430 0020 0440 0430 0441 043F 0440 0435 043B 0435 043B 0435 043D 0438 044F 0020 0443 043A 0437 0026 0020 0045 0072 0072 002E 0044 0065 0073 0063 0072") & Ru("0069 0070 0074 0069 006F 006E")
    On Error GoTo CleanExit
CleanExit:
End Sub


' Clears all filters in a column
Private Sub ClearFiltersInColumn(ByVal ws As Worksheet, ByVal col As Long)
    Dim rowIdx As Long
    For rowIdx = FILTER_START_ROW To FILTER_END_ROW
        ws.Cells(rowIdx, col).ClearContents
    Next rowIdx
End Sub

Public Sub SetPipeCountCPProgrammatically(ByVal newValue As Long)
    isProgrammaticPipeCountChange = True
    Application.EnableEvents = True
    Me.Range("pipeCountCP").Value = newValue
    DoEvents
    isProgrammaticPipeCountChange = False
End Sub

' Public variable - read inside the module: Public isProgrammaticPipeCountChange As Boolean
' Public procedure to set the mode
Public Sub SetProgrammaticMode(ByVal mode As Boolean)
    isProgrammaticPipeCountChange = mode
    If DEBUG_MODE Then Debug.Print ">>> programmatic mode = " & mode
End Sub
