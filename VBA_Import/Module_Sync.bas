Attribute VB_Name = "Module_Sync"

Option Explicit

' ================================================================
' Module_Sync
' Column synchronization and named range updates
' ================================================================

' Synchronizes columns (copy / clear) based on pipeCountCP, max 254 columns
Public Sub SyncAnodColumns(ByVal ws As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit
    
    Application.ScreenUpdating = True
    DoEvents
    Call SetStatusBar("syncing Anod columns (" & colBr & ")...")
    DoEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    colBr = SafeNumericValue(colBr, 1)
    If colBr > MAX_COL Then colBr = MAX_COL
    
    Dim actualCols As Long
    actualCols = 0
    Dim col As Long
    For col = START_COL To LAST_ANOD_COL
        If Not IsEmpty(ws.Cells(HEADER_ROW, col).Value) Then
            actualCols = actualCols + 1
        Else
            Exit For
        End If
    Next col
    If actualCols = 0 Then actualCols = 1
    
    Dim endCol As Long
    endCol = START_COL + colBr - 1
    endCol = Application.WorksheetFunction.Max(START_COL, Application.WorksheetFunction.Min(LAST_ANOD_COL, endCol))
    
    If colBr > actualCols Then
        Dim sourceCol As Long
        sourceCol = START_COL + actualCols - 1
        If sourceCol < START_COL Then sourceCol = START_COL
        
        Dim newCol As Long
        Dim totalSteps As Long
        totalSteps = colBr - actualCols
        Dim step As Long
        step = 0
        
        Call SetStatusBar("copying Anod columns 0 / " & totalSteps)
        DoEvents
        
        Dim firstNewCol As Long
        firstNewCol = sourceCol + 1
        For newCol = sourceCol + 1 To endCol
            step = step + 1
            Call PulseProgress("copying Anod columns", step, totalSteps)
            
            ws.Columns(sourceCol).Copy Destination:=ws.Columns(newCol)
            ws.Cells(HEADER_ROW, newCol).Value = ChrW(1059) & ChrW(1050) & ChrW(1047) & (newCol - START_COL + 1)
            
            Dim rowClear As Variant
            For Each rowClear In Array(3, 4, 5, 6, 7, 16, 17, 18, 19)
                ws.Cells(rowClear, newCol).ClearContents
                ws.Cells(rowClear, newCol).ClearFormats
                On Error Resume Next
                ws.Cells(rowClear, newCol).Validation.Delete
                On Error GoTo CleanExit
            Next rowClear
            
            Dim r As Long
            For r = 60 To 118
                ws.Cells(r, newCol).ClearContents
            Next r
            
            sourceCol = newCol
        Next newCol
        Call DeleteAnodLayerButtonsInColumns(ws, firstNewCol, endCol)
    End If
    
    If colBr < actualCols Then
        Dim clearStartCol As Long
        clearStartCol = START_COL + colBr
        Dim lastCol As Long
        lastCol = START_COL + actualCols - 1
        If lastCol > LAST_ANOD_COL Then lastCol = LAST_ANOD_COL
        If clearStartCol <= lastCol Then
            Call SetStatusBar("clearing extra Anod columns...")
            DoEvents
            ws.Range(ws.Columns(clearStartCol), ws.Columns(lastCol)).ClearContents
            ws.Range(ws.Columns(clearStartCol), ws.Columns(lastCol)).ClearFormats
            Call DeleteShapesInColumnRange(ws, clearStartCol, LAST_ANOD_COL)
        End If
    End If
    
    Call SetStatusBar("updating column visibility...")
    DoEvents
    
    Dim i As Long
    For i = START_COL To endCol
        ws.Columns(i).Hidden = False
        If ws.Columns(i).ColumnWidth = 0 Then ws.Columns(i).ColumnWidth = 8.43
    Next i
    
    If endCol < LAST_ANOD_COL Then
        ws.Range(ws.Columns(endCol + 1), ws.Columns(LAST_ANOD_COL)).EntireColumn.Hidden = False
    End If
    
    Call SetStatusBar("updating filters...")
    DoEvents
    Call UpdateFilterNamedRanges(ws, colBr)
    Application.ScreenUpdating = True
    DoEvents
    
    Call SetStatusBar("updating validation...")
    DoEvents
    Call Module_ValidationLogic.RefreshValidationForColumns(ws, START_COL, endCol)
    
    Call SetStatusBar("updating cp_list...")
    DoEvents
    Call UpdateCPValidationList
    Application.ScreenUpdating = True
    DoEvents
    
    Call SetStatusBar("updating typeCP...")
    DoEvents
    Call RefreshAllTypeCPValidation(ws)
    Application.ScreenUpdating = True
    DoEvents
    
    Call SetStatusBar("updating row visibility...")
    DoEvents
    Call Module_Visual.UpdateVisibilityForAllColumns(ws)
    Call SetStatusBar("updating colors...")
    DoEvents
    Call Module_Visual.UpdateColorsForAllColumns(ws)
    Application.ScreenUpdating = True
    DoEvents
    
    Call SetStatusBar("recalculating Anod...")
    DoEvents
    ws.Calculate
    Application.ScreenUpdating = True
    DoEvents
    
    Application.CutCopyMode = False
    DoEvents

CleanExit:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    DoEvents
End Sub

' Form Control / shapes survive Columns.Clear. Remove those sitting
' in dropped columns after a shrink sync (Pipe and Anod).
Public Sub DeleteShapesInColumnRange(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim iShp As Long
    Dim shp As Shape
    Dim colShp As Long
    Dim wasProt As Boolean

    If ws Is Nothing Then Exit Sub
    If firstCol > lastCol Then Exit Sub

    wasProt = ws.ProtectContents
    On Error Resume Next
    If wasProt Then ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    For iShp = ws.Shapes.Count To 1 Step -1
        Set shp = Nothing
        colShp = 0
        On Error Resume Next
        Set shp = ws.Shapes(iShp)
        If Not shp Is Nothing Then colShp = shp.TopLeftCell.Column
        If colShp >= firstCol And colShp <= lastCol Then shp.Delete
        On Error GoTo 0
    Next iShp

    If wasProt Then
        On Error Resume Next
        ws.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

' обновление именованных диапазонов листа Anod
' updates named ranges on Anod sheet
' все нескалярные имена Anod имеют размерность 1 x pipeCountCP
' all non-scalar Anod names are 1 x pipeCountCP
Public Sub UpdateNamedRanges(ByVal ws As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit

    Application.StatusBar = "updating named ranges..."
    DoEvents

    colBr = SafeNumericValue(colBr, 1)

    Dim endCol As Long
    endCol = START_COL + colBr - 1
    endCol = Application.WorksheetFunction.Max(START_COL, Application.WorksheetFunction.Min(LAST_ANOD_COL, endCol))

    Dim updatedCount As Long
    updatedCount = 0

    Dim arr As Variant
    Dim i As Long
    Dim nm1 As String
    Dim nmObj As name
    Dim rng As Range
    Dim newRng As Range

    arr = Split(ANOD_RANGE_NAMES_1ROW, "|")
    For i = LBound(arr) To UBound(arr)
        nm1 = Trim$(arr(i))
        If Len(nm1) = 0 Then GoTo NextFixedName

        Set rng = Nothing
        Set nmObj = Nothing
        On Error Resume Next
        Set nmObj = ws.names(nm1)
        If nmObj Is Nothing Then Set nmObj = thisWorkbook.names(nm1)
        If Not nmObj Is Nothing Then Set rng = nmObj.RefersToRange
        Err.Clear
        On Error GoTo CleanExit

        If rng Is Nothing Then GoTo NextFixedName
        If rng.Parent.name <> ws.name Then GoTo NextFixedName

        Set newRng = ws.Range(ws.Cells(rng.row, START_COL), ws.Cells(rng.row, endCol))
        On Error Resume Next
        nmObj.RefersTo = newRng
        If Err.Number = 0 Then
            updatedCount = updatedCount + 1
        Else
            Err.Clear
        End If
        On Error GoTo CleanExit
NextFixedName:
    Next i

    Call PinAnodResultNames1xN(ws, endCol)

    Application.StatusBar = "named ranges updated (" & updatedCount & ")"
    DoEvents

CleanExit:
End Sub

' фиксируем расчётные имена 1 x pipeCountCP на постоянных строках
' pin calculation-result names to 1 x pipeCountCP on their constant rows
Private Sub PinAnodResultNames1xN(ByVal ws As Worksheet, ByVal endCol As Long)
    On Error Resume Next
    Call SetAnodName1xN(ws, "oneElectrodeResistanceAG", ROW_ONE_ELECTRODE_RESISTANCE_AG, endCol)
    Call SetAnodName1xN(ws, "oneElectrodeResistanceHorizAG", ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, endCol)
    Call SetAnodName1xN(ws, "numElectrodesAG", ROW_NUM_ELECTRODES_AG, endCol)
    Call SetAnodName1xN(ws, "weightWithoutFillingAG", ROW_WEIGHT_WITHOUT_FILLING_AG, endCol)
    Call SetAnodName1xN(ws, "serviceLifeAG", ROW_SERVICE_LIFE_AG, endCol)
    Call SetAnodName1xN(ws, "serviceLifeDeviation", ROW_SERVICE_LIFE_DEVIATION, endCol)
    Call SetAnodName1xN(ws, "correctResistanceAG", ROW_CORRECT_RESISTANCE_AG, endCol)
    On Error GoTo 0
End Sub

Private Sub SetAnodName1xN(ByVal ws As Worksheet, ByVal rangeName As String, ByVal rowNum As Long, ByVal endCol As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, START_COL), ws.Cells(rowNum, endCol))
    On Error Resume Next
    ws.names(rangeName).RefersTo = rng
    If Err.Number <> 0 Then
        Err.Clear
        ws.names.Add name:=rangeName, RefersTo:=rng
    End If
    On Error GoTo 0
End Sub

' Updates filter named ranges (rows 34-38)
Public Sub UpdateFilterNamedRanges(ByVal ws As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit
    
    Dim startCol As Long
    startCol = START_COL
    Dim endCol As Long
    endCol = startCol + colBr - 1
    
    If startCol > endCol Then Exit Sub
    
    Dim filterNames As Variant
    Dim filterRows As Variant
    
    filterNames = Array( _
        "typeMaterial", _
        "typeMountingAG", _
        "typeInstallationAG", _
        "typeDeliveryAG", _
        "typeAG" _
    )
    
    filterRows = Array(34, 35, 36, 37, 38)
    
    Dim idx As Long
    Dim rowNum As Long
    Dim rng As Range
    Dim createdCount As Long
    Dim existsCount As Long
    createdCount = 0
    existsCount = 0
    
    For idx = LBound(filterNames) To UBound(filterNames)
        rowNum = filterRows(idx)
        
        Set rng = ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, endCol))
        
        Dim nameExists As Boolean
        nameExists = False
        Dim existingRng As Range
        
        On Error Resume Next
        Set existingRng = ws.names(filterNames(idx)).RefersToRange
        If Err.Number = 0 Then
            nameExists = True
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If nameExists Then
            existsCount = existsCount + 1
            If existingRng.address <> rng.address Then
                On Error Resume Next
                ws.names(filterNames(idx)).RefersTo = rng
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        Else
            On Error Resume Next
            ws.names.Add name:=filterNames(idx), RefersTo:=rng
            If Err.Number = 0 Then
                createdCount = createdCount + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next idx
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    DoEvents
    
CleanExit:
End Sub
