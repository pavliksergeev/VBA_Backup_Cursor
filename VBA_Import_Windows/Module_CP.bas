Attribute VB_Name = "Module_CP"

Option Explicit

' ================================================================
' Module_CP
' CP data update, CP validation
' ================================================================

' Updates CP data (current, voltage, power, hyperlink) for a given column
Public Sub UpdateCPData(ByVal ws As Worksheet, ByVal col As Long)
    On Error GoTo CleanExit
    
    Dim eventsState As Boolean
    eventsState = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim modelValue As Variant
    modelValue = ws.Cells(CP_MODEL_ROW, col).Value
    
    If IsEmpty(modelValue) Or modelValue = "" Then
        ws.Cells(CP_HYPERLINK_ROW, col).Value = ""
        ws.Cells(CP_CURRENT_ROW, col).Value = ""
        ws.Cells(CP_VOLTAGE_ROW, col).Value = ""
        ws.Cells(CP_POWER_ROW, col).Value = ""
        GoTo CleanExit
    End If
    
    Dim wsListCP As Worksheet
    If Not WorksheetExists(SHEET_LIST_CP) Then GoTo CleanExit
    Set wsListCP = thisWorkbook.Worksheets(SHEET_LIST_CP)
    
    Dim tblCP As ListObject
    If Not TableExists(wsListCP, TABLE_CP) Then GoTo CleanExit
    Set tblCP = wsListCP.ListObjects(TABLE_CP)
    
    Dim headerRow As Range
    Set headerRow = tblCP.HeaderRowRange
    
    Dim typeColNum As Long
    Dim headerCell As Range
    
    On Error Resume Next
    Set headerCell = headerRow.Find(What:="typeCP", LookIn:=xlValues, LookAt:=xlWhole)
    If headerCell Is Nothing Then
        Set headerCell = headerRow.Find(What:="model", LookIn:=xlValues, LookAt:=xlPart)
    End If
    On Error GoTo CleanExit
    
    If headerCell Is Nothing Then GoTo CleanExit
    typeColNum = headerCell.Column - tblCP.HeaderRowRange.Column + 1
    
    Dim modelCol As Range
    Set modelCol = tblCP.ListColumns(typeColNum).DataBodyRange
    
    Dim foundRow As Variant
    On Error Resume Next
    foundRow = Application.Match(modelValue, modelCol, 0)
    On Error GoTo CleanExit
    
    If IsError(foundRow) Then
        ws.Cells(CP_HYPERLINK_ROW, col).Value = ""
        ws.Cells(CP_CURRENT_ROW, col).Value = ""
        ws.Cells(CP_VOLTAGE_ROW, col).Value = ""
        ws.Cells(CP_POWER_ROW, col).Value = ""
        GoTo CleanExit
    End If
    
    ' Find column numbers for needed fields
    Dim colNames As Variant
    colNames = Array("nominalOutputCurrenCP", "nominalOutputVoltageCP", "link", "power")
    Dim colNums As New Collection
    Dim i As Long
    For i = LBound(colNames) To UBound(colNames)
        On Error Resume Next
        Set headerCell = headerRow.Find(What:=colNames(i), LookIn:=xlValues, LookAt:=xlPart)
        If Err.Number <> 0 Or headerCell Is Nothing Then
            colNums.Add 0
        Else
            colNums.Add headerCell.Column - tblCP.HeaderRowRange.Column + 1
        End If
        On Error GoTo CleanExit
    Next i
    
    Dim currentColNum As Long, voltageColNum As Long, linkColNum As Long, powerColNum As Long
    currentColNum = colNums(1)
    voltageColNum = colNums(2)
    linkColNum = colNums(3)
    powerColNum = colNums(4)
    
    Dim curVal As Variant
    Dim voltVal As Variant
    
    If linkColNum > 0 Then
        Dim linkVal As Variant
        linkVal = tblCP.ListColumns(linkColNum).DataBodyRange.Cells(foundRow, 1).Value
        If Not IsEmpty(linkVal) Then ws.Cells(CP_HYPERLINK_ROW, col).Value = linkVal
    End If
    
    If currentColNum > 0 Then
        curVal = tblCP.ListColumns(currentColNum).DataBodyRange.Cells(foundRow, 1).Value
        If IsNumeric(curVal) Then ws.Cells(CP_CURRENT_ROW, col).Value = curVal
    End If
    
    If voltageColNum > 0 Then
        voltVal = tblCP.ListColumns(voltageColNum).DataBodyRange.Cells(foundRow, 1).Value
        If IsNumeric(voltVal) Then ws.Cells(CP_VOLTAGE_ROW, col).Value = voltVal
    End If
    
    If IsNumeric(curVal) And IsNumeric(voltVal) Then
        ws.Cells(CP_POWER_ROW, col).Value = curVal * voltVal
    Else
        If powerColNum > 0 Then
            Dim powVal As Variant
            powVal = tblCP.ListColumns(powerColNum).DataBodyRange.Cells(foundRow, 1).Value
            If IsNumeric(powVal) Then ws.Cells(CP_POWER_ROW, col).Value = powVal
        End If
    End If

    ws.Calculate

CleanExit:
    Application.EnableEvents = eventsState
End Sub

' Updates CP validation list (creates named range "cp_List")
Public Sub UpdateCPValidationList()
    On Error GoTo CleanExit

    Application.StatusBar = "updating cp_list..."
    DoEvents

    Dim wsSettings As Worksheet
    Dim wsAnod As Worksheet
    Dim cpCol As Collection
    Dim i As Long
    Dim arrData() As Variant

    On Error Resume Next
    Set wsSettings = thisWorkbook.Worksheets("_Settings")
    If wsSettings Is Nothing Then
        On Error GoTo 0
        Set wsSettings = thisWorkbook.Worksheets.Add(After:=thisWorkbook.Worksheets(thisWorkbook.Worksheets.count))
        wsSettings.name = "_Settings"
        wsSettings.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    Set wsAnod = thisWorkbook.Worksheets("Anod")
    If wsAnod Is Nothing Then Exit Sub

    Set cpCol = GetCPCollection()
    If cpCol.count = 0 Then Exit Sub

    wsSettings.Columns(1).Clear

    ReDim arrData(1 To cpCol.count, 1 To 1)
    For i = 1 To cpCol.count
        arrData(i, 1) = cpCol(i)
    Next i

    wsSettings.Range("A1:A" & cpCol.count).Value = arrData

    On Error Resume Next
    wsAnod.names("cp_List").Delete
    Err.Clear
    On Error GoTo 0

    Dim listRange As Range
    Set listRange = wsSettings.Range("A1:A" & cpCol.count)
    wsAnod.names.Add name:="cp_List", RefersTo:=listRange

    Application.StatusBar = "cp_list updated"
    DoEvents

    Exit Sub

CleanExit:
    Application.StatusBar = False
End Sub

' Refreshes CP validation for all columns
Public Sub RefreshAllTypeCPValidation(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Application.StatusBar = "updating typeCP..."
    DoEvents

    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = SafeRangeValue(ws.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)

    Dim startCol As Long
    startCol = START_COL
    Dim endCol As Long
    endCol = startCol + colBr - 1

    Dim cpListExists As Boolean
    Dim cpTestRng As Range
    On Error Resume Next
    Set cpTestRng = ws.names("cp_List").RefersToRange
    If Err.Number = 0 Then
        cpListExists = True
    Else
        cpListExists = False
        Err.Clear
    End If
    On Error GoTo 0

    If Not cpListExists Then Exit Sub

    Dim col As Long
    Dim validationFormula As String
    validationFormula = "=cp_List"

    For col = startCol To endCol
        On Error Resume Next
        With ws.Cells(CP_MODEL_ROW, col).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=validationFormula
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Next col

    Application.StatusBar = "typeCP updated"
    DoEvents
    Application.StatusBar = False

CleanExit:
    Application.StatusBar = False
End Sub

' Returns collection of unique CP models
Public Function GetCPCollection() As Collection
    On Error GoTo CleanExit

    Dim wsListCP As Worksheet
    Dim tblCP As ListObject
    Dim colIndex As Long
    Dim headerCell As Range
    Dim vals As Variant
    Dim col As New Collection
    Dim i As Long, val As Variant, item As Variant
    Dim exists As Boolean

    Set wsListCP = thisWorkbook.Worksheets(SHEET_LIST_CP)
    If wsListCP Is Nothing Then
        Set GetCPCollection = New Collection
        Exit Function
    End If

    If Not TableExists(wsListCP, TABLE_CP) Then
        Set GetCPCollection = New Collection
        Exit Function
    End If

    Set tblCP = wsListCP.ListObjects(TABLE_CP)

    Set headerCell = tblCP.HeaderRowRange.Find(What:="typeCP", LookIn:=xlValues, LookAt:=xlWhole)
    If headerCell Is Nothing Then
        Set headerCell = tblCP.HeaderRowRange.Find(What:="model", LookIn:=xlValues, LookAt:=xlPart)
    End If
    If headerCell Is Nothing Then
        Set GetCPCollection = New Collection
        Exit Function
    End If

    colIndex = headerCell.Column - tblCP.HeaderRowRange.Column + 1
    vals = tblCP.ListColumns(colIndex).DataBodyRange.Value

    If Not IsArray(vals) Then
        ReDim arr(1 To 1, 1 To 1): arr(1, 1) = vals: vals = arr
    End If

    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsError(val) And Not IsEmpty(val) And val <> "" Then
            Dim cleanVal As String
            cleanVal = CStr(val)
            cleanVal = Replace(cleanVal, vbLf, "")
            cleanVal = Replace(cleanVal, vbCr, "")
            cleanVal = Replace(cleanVal, vbTab, " ")
            cleanVal = Application.WorksheetFunction.Trim(cleanVal)
            If cleanVal = "" Then GoTo Skip

            exists = False
            For Each item In col
                If item = cleanVal Then
                    exists = True
                    Exit For
                End If
            Next item
            If Not exists Then col.Add cleanVal
        End If
Skip:
    Next i

    Set GetCPCollection = col
    Exit Function

CleanExit:
    Set GetCPCollection = New Collection
End Function
