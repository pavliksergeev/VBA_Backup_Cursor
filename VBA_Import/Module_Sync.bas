Attribute VB_Name = "Module_Sync"

Option Explicit

' ================================================================
' Module_Sync
' Column synchronization and named range updates
' ================================================================

' Synchronizes columns (copy, delete, hide) based on pipeCountCP
Public Sub SyncAnodColumns(ByVal ws As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit
    
    Application.ScreenUpdating = True
    DoEvents
    Application.StatusBar = "syncing columns (" & colBr & ")..."
    DoEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    colBr = SafeNumericValue(colBr, 1)
    
    Dim actualCols As Long
    actualCols = 0
    Dim col As Long
    For col = START_COL To START_COL + 100
        If Not IsEmpty(ws.Cells(HEADER_ROW, col).Value) Then
            actualCols = actualCols + 1
        Else
            Exit For
        End If
    Next col
    If actualCols = 0 Then actualCols = 1
    
    Dim endCol As Long
    endCol = START_COL + colBr - 1
    endCol = Application.WorksheetFunction.Max(START_COL, Application.WorksheetFunction.Min(MAX_COL, endCol))
    
    If colBr > actualCols Then
        Dim sourceCol As Long
        sourceCol = START_COL + actualCols - 1
        If sourceCol < START_COL Then sourceCol = START_COL
        
        Dim newCol As Long
        Dim totalSteps As Long
        totalSteps = colBr - actualCols
        Dim step As Long
        step = 0
        
        Call Module_Constants.ShowProgress(0, totalSteps, "copying Anod columns...")
        
        For newCol = sourceCol + 1 To endCol
            step = step + 1
            
            If step Mod 5 = 0 Or step = totalSteps Then
                Call Module_Constants.ShowProgress(step, totalSteps, "copying Anod columns...")
                Application.ScreenUpdating = True
                DoEvents
            End If
            
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
        
        Call Module_Constants.ClearProgress
    End If
    
    If colBr < actualCols Then
        Dim clearStartCol As Long
        clearStartCol = START_COL + colBr
        Dim lastCol As Long
        lastCol = START_COL + actualCols - 1
        If clearStartCol <= lastCol Then
            ws.Range(ws.Columns(clearStartCol), ws.Columns(lastCol)).Clear
        End If
    End If
    
    Application.StatusBar = "updating column visibility..."
    DoEvents
    
    Dim i As Long
    For i = START_COL To endCol
        ws.Columns(i).Hidden = False
        If ws.Columns(i).ColumnWidth = 0 Then ws.Columns(i).ColumnWidth = 8.43
    Next i
    
    Dim lastUsedCol As Long
    On Error Resume Next
    lastUsedCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    If lastUsedCol = 0 Then lastUsedCol = MAX_COL
    If endCol < lastUsedCol Then
        ws.Range(ws.Columns(endCol + 1), ws.Columns(lastUsedCol)).EntireColumn.Hidden = True
    End If
    
    Application.StatusBar = "updating filters..."
    DoEvents
    Call UpdateFilterNamedRanges(ws, colBr)
    Application.ScreenUpdating = True
    DoEvents
    
    Application.StatusBar = "updating validation..."
    DoEvents
    Dim colIdx As Long
    For colIdx = START_COL To endCol
        Call Module_ValidationLogic.RefreshValidationForColumn(ws, colIdx)
        If colIdx Mod 10 = 0 Then
            Application.ScreenUpdating = True
            DoEvents
        End If
    Next colIdx
    
    Application.StatusBar = "updating cp_list..."
    DoEvents
    Call UpdateCPValidationList
    Application.ScreenUpdating = True
    DoEvents
    
    Application.StatusBar = "updating typeCP..."
    DoEvents
    Call RefreshAllTypeCPValidation(ws)
    Application.ScreenUpdating = True
    DoEvents
    
    Call Module_Visual.UpdateVisibilityForAllColumns(ws)
    Call Module_Visual.UpdateColorsForAllColumns(ws)
    Application.ScreenUpdating = True
    DoEvents
    
    Application.StatusBar = "recalculating..."
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

' Updates named ranges on Anod sheet
Public Sub UpdateNamedRanges(ByVal ws As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit

    Application.StatusBar = "updating named ranges..."
    DoEvents

    colBr = SafeNumericValue(colBr, 1)

    Dim endCol As Long
    endCol = START_COL + colBr - 1
    endCol = Application.WorksheetFunction.Max(START_COL, Application.WorksheetFunction.Min(MAX_COL, endCol))

    Dim updatedCount As Long
    updatedCount = 0

    Dim nm As name
    Dim rng As Range
    Dim newRng As Range
    Dim startRow As Long
    Dim endRow As Long
    Dim rowCount As Long
    Dim nameIndex As Long

    ' ================================================================
    ' 1. workbook-level names
    ' ================================================================
    nameIndex = 0
    For Each nm In thisWorkbook.names
        nameIndex = nameIndex + 1
        If nameIndex Mod 10 = 0 Then
            Application.StatusBar = "updating global names (" & nameIndex & ")"
            DoEvents
        End If

        On Error Resume Next
        Set rng = ws.Range(nm.name)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextGlobalName
        End If
        On Error GoTo 0

        If Not rng Is Nothing Then
            If rng.Parent.name = ws.name Then
                If Not Module_Constants.IsScalarName(nm.name) Then
                    startRow = rng.row
                    rowCount = rng.rows.count           ' ? actual row count
                    endRow = startRow + rowCount - 1
                    Set newRng = ws.Range(ws.Cells(startRow, START_COL), ws.Cells(endRow, endCol))

                    If Not newRng Is Nothing Then
                        On Error Resume Next
                        nm.RefersTo = newRng
                        If Err.Number = 0 Then
                            updatedCount = updatedCount + 1
                        Else
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
NextGlobalName:
    Next nm

    ' ================================================================
    ' 2. sheet-local names
    ' ================================================================
    nameIndex = 0
    For Each nm In ws.names
        nameIndex = nameIndex + 1
        If nameIndex Mod 10 = 0 Then
            Application.StatusBar = "updating local names (" & nameIndex & ")"
            DoEvents
        End If

        On Error Resume Next
        Set rng = ws.Range(nm.name)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextLocalName
        End If
        On Error GoTo 0

        If Not rng Is Nothing Then
            If Not Module_Constants.IsScalarName(nm.name) Then
                startRow = rng.row
                rowCount = rng.rows.count               ' ? actual row count
                endRow = startRow + rowCount - 1
                Set newRng = ws.Range(ws.Cells(startRow, START_COL), ws.Cells(endRow, endCol))

                If Not newRng Is Nothing Then
                    On Error Resume Next
                    nm.RefersTo = newRng
                    If Err.Number = 0 Then
                        updatedCount = updatedCount + 1
                    Else
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
NextLocalName:
    Next nm

    Application.StatusBar = "named ranges updated (" & updatedCount & ")"
    DoEvents

CleanExit:
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
        If idx Mod 2 = 0 Then
            Application.StatusBar = "restoring filters: " & filterNames(idx)
            Application.ScreenUpdating = True
            DoEvents
        End If
        
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
