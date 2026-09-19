Attribute VB_Name = "Module_DataCleanup"

' ================================================================
' module: Module_DataCleanup
' purpose: cleanup and normalization of data in reference tables
' ================================================================
Option Explicit

' ------------------------------------------------------------------
' clean TableAG (replace invalid characters with "_")
' ------------------------------------------------------------------
Sub CleanTableAG()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    
    ' list of columns to clean (exact names)
    Set targetColumns = New Collection
    ' материал анода
    ' anode material
    targetColumns.Add Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430")
    ' тип монтажа
    ' mounting type
    targetColumns.Add Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430")
    ' комплектация
    ' delivery set
    targetColumns.Add Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F")
    targetColumns.Add "model"
    
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
    If ws Is Nothing Then
        MsgBox "sheet '" & SHEET_LIST_AG & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(TABLE_AG)
    If tbl Is Nothing Then
        MsgBox "table '" & TABLE_AG & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' iterate over all table columns
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        
        ' check if this column should be cleaned
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            ' clean all cells in the column (except header)
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    cell.Value = CleanStringForTable(CStr(cell.Value))
                End If
            Next cell
        End If
    Next col
    
    Call Module_ValidationLists.InvalidateAGListCache
    MsgBox "table " & TABLE_AG & " cleaned of invalid characters." & vbCrLf & _
           "all invalid characters replaced with '_'.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ================================================================
' comprehensive cleanup and refresh (button on ListAG sheet)
' ================================================================
Sub FullDataCleanupAndRefresh()
    On Error GoTo CleanExit

    Dim answer As VbMsgBoxResult

    ' step 1: clean TableAG
    answer = MsgBox("clean table " & TABLE_AG & " of invalid characters?" & vbCrLf & _
                    "all invalid characters will be replaced with '_'.", _
                    vbYesNo + vbQuestion, "confirm TableAG cleanup")
    If answer = vbYes Then
        Call CleanTableAG
    Else
        MsgBox "table " & TABLE_AG & " cleanup skipped.", vbInformation
    End If

    ' step 2: clean TableMounting (if exists)
    answer = MsgBox("clean table TableMounting of invalid characters?" & vbCrLf & _
                    "if the table is not found, this step will be skipped.", _
                    vbYesNo + vbQuestion, "confirm TableMounting cleanup")
    If answer = vbYes Then
        Call CleanTableMounting
    Else
        MsgBox "TableMounting cleanup skipped.", vbInformation
    End If

    ' step 3: restore named ranges on Anod sheet
    answer = MsgBox("restore named ranges on Anod sheet?" & vbCrLf & _
                    "this will recreate all local names according to the current number of columns.", _
                    vbYesNo + vbQuestion, "confirm name restoration")
    If answer = vbYes Then
        Application.Run "RestoreAnodNamedRanges"
    Else
        MsgBox "name restoration skipped.", vbInformation
    End If

    ' step 4: refresh validation (dropdowns) on Anod sheet
    answer = MsgBox("refresh validation (dropdowns) on Anod sheet?" & vbCrLf & _
                    "dropdown lists will be recreated for all columns.", _
                    vbYesNo + vbQuestion, "confirm validation refresh")
    If answer = vbYes Then
        Application.Run "'" & SHEET_ANOD & "'!ForceRefreshAllValidation"
    Else
        MsgBox "validation refresh skipped.", vbInformation
    End If

    ' step 5: force synchronization of columns and named ranges
    answer = MsgBox("force refresh of columns and names on Anod sheet (sync)?" & vbCrLf & _
                    "this will recreate columns according to the current PipeCountCP and update all names.", _
                    vbYesNo + vbQuestion, "column synchronization")
    If answer = vbYes Then
        Application.Run "'" & SHEET_ANOD & "'!UpdateAllRanges"
    Else
        MsgBox "column synchronization skipped.", vbInformation
    End If

    ' force-enable events (in case they were disabled)
    Application.EnableEvents = True

    ' final reminder
    MsgBox "all selected operations completed." & vbCrLf & vbCrLf & _
           "important: check filters and data insertion in all columns!" & vbCrLf & _
           "make sure dropdowns appear and values are inserted correctly." & vbCrLf & vbCrLf & _
           "if columns do not update when PipeCountCP changes, run in the Immediate Window:" & vbCrLf & _
           "Application.EnableEvents = True", _
           vbInformation, "operation completed"

CleanExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "error: " & Err.Description, vbCritical
    End If
End Sub

' ------------------------------------------------------------------
' clean TableMounting (replace invalid characters with "_")
' ------------------------------------------------------------------
Sub CleanTableMounting()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    Dim found As Boolean
    
    ' list of columns to clean
    Set targetColumns = New Collection
    ' тип монтажа
    ' mounting type
    targetColumns.Add Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430")
    ' способ монтажа
    ' installation method
    targetColumns.Add Ru("0441 043F 043E 0441 043E 0431 0020 043C 043E 043D 0442 0430 0436 0430")
    
    ' try to find the table on sheet "ListAG"
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets("ListAG")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' if sheet "ListAG" is missing, try SHEET_LIST_AG
        Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
        If ws Is Nothing Then
            MsgBox "sheet for TableMounting not found!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' check for "TableMounting" existence
    On Error Resume Next
    Set tbl = ws.ListObjects("TableMounting")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "table 'TableMounting' not found on sheet " & ws.name, vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' iterate over all table columns
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    cell.Value = CleanStringForTable(CStr(cell.Value))
                End If
            Next cell
        End If
    Next col
    
    Call Module_ValidationLists.InvalidateAGListCache
    MsgBox "table TableMounting cleaned of invalid characters." & vbCrLf & _
           "all invalid characters replaced with '_'.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------
' clean TableCP (from control characters)
' ------------------------------------------------------------------
Sub CleanTableCP()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    
    Set targetColumns = New Collection
    targetColumns.Add "typeCP"
    ' модель
    ' model
    targetColumns.Add Ru("043C 043E 0434 0435 043B 044C")
    
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_CP)
    If ws Is Nothing Then
        MsgBox "sheet '" & SHEET_LIST_CP & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(TABLE_CP)
    If tbl Is Nothing Then
        MsgBox "table '" & TABLE_CP & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    Dim cleaned As String
                    cleaned = CleanStringForTable(CStr(cell.Value))
                    cleaned = Replace(cleaned, vbLf, "")
                    cleaned = Replace(cleaned, vbCr, "")
                    cleaned = Replace(cleaned, vbTab, " ")
                    cleaned = Application.WorksheetFunction.Trim(cleaned)
                    cell.Value = cleaned
                End If
            Next cell
        End If
    Next col
    
    MsgBox "table " & TABLE_CP & " cleaned of invalid characters.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
