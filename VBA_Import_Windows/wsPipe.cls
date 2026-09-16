VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsPipe"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ================================================================
' wsPipe sheet - full version
' Validation for steel grade, diameter, wall thickness,
' pipeResistivityChangeFactor, pipeInsulationResistivityStartLife
' ================================================================
Option Explicit

Private previousPipeCount As Variant
Private isUpdating As Boolean

' ================================================================
' helper: read array from a named range
' ================================================================
Private Function EnsureArray2(ByVal rangeName As String, ByVal cols As Long) As Variant
    Dim rng As Range
    Dim vals As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim wsPipe As Worksheet
    
    Set wsPipe = Me

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
End Function

' ================================================================
' get list of thickness header values for a given diameter from tableB
' ================================================================
Private Function GetThicknessHeadersForDiameter(ByVal diameter As Double) As Variant
    On Error GoTo CleanExit
    
    Dim wsRef As Worksheet
    Dim tblB As ListObject
    Dim diamRow As Long
    Dim col As Long
    Dim headerValues As Collection
    Dim arrResult() As Variant
    Dim i As Long
    
    On Error Resume Next
    Set wsRef = thisWorkbook.Worksheets(SHEET_STO_REF)
    If wsRef Is Nothing Then
        GetThicknessHeadersForDiameter = Array()
        Exit Function
    End If
    
    Set tblB = wsRef.ListObjects("TableB")
    If tblB Is Nothing Then
        GetThicknessHeadersForDiameter = Array()
        Exit Function
    End If
    On Error GoTo CleanExit
    
    Dim foundRow As Variant
    On Error Resume Next
    foundRow = Application.Match(diameter, tblB.ListColumns(1).DataBodyRange, 0)
    On Error GoTo CleanExit
    
    If IsError(foundRow) Then
        GetThicknessHeadersForDiameter = Array()
        Exit Function
    End If
    
    diamRow = CLng(foundRow)
    Set headerValues = New Collection
    
    For col = 3 To 24
        Dim cellValue As Variant
        Dim headerValue As Variant
        
        On Error Resume Next
        cellValue = tblB.DataBodyRange.Cells(diamRow, col).Value
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextCol
        End If
        On Error GoTo CleanExit
        
        If Not IsEmpty(cellValue) And cellValue <> "" Then
            If IsNumeric(cellValue) Then
                On Error Resume Next
                headerValue = tblB.HeaderRowRange.Cells(1, col).Value
                If Err.Number <> 0 Then
                    Err.Clear
                    GoTo NextCol
                End If
                On Error GoTo CleanExit
                
                If Not IsEmpty(headerValue) And headerValue <> "" Then
                    If IsNumeric(headerValue) Then
                        Dim exists As Boolean
                        exists = False
                        Dim item As Variant
                        For Each item In headerValues
                            If CDbl(item) = CDbl(headerValue) Then
                                exists = True
                                Exit For
                            End If
                        Next item
                        
                        If Not exists Then
                            headerValues.Add CDbl(headerValue)
                        End If
                    End If
                End If
            End If
        End If
        
NextCol:
    Next col
    
    If headerValues.count > 0 Then
        ReDim arrResult(1 To headerValues.count)
        For i = 1 To headerValues.count
            arrResult(i) = headerValues(i)
        Next i
        GetThicknessHeadersForDiameter = arrResult
    Else
        GetThicknessHeadersForDiameter = Array()
    End If

CleanExit:
    If IsEmpty(GetThicknessHeadersForDiameter) Then
        GetThicknessHeadersForDiameter = Array()
    End If
End Function

' ================================================================
' get list of Rins0 values from table Table6_1 (STO sheet)
' ================================================================
Private Function GetInsulationResistivityList() As String
    On Error GoTo CleanExit
    
    Dim wsSTO As Worksheet
    Set wsSTO = thisWorkbook.Worksheets(SHEET_STO_REF)
    If wsSTO Is Nothing Then
        Debug.Print "GetInsulationResistivityList: sheet " & SHEET_STO_REF & " not found"
        GetInsulationResistivityList = ""
        Exit Function
    End If
    
    Dim tbl As ListObject
    Set tbl = wsSTO.ListObjects("Table6_1")
    If tbl Is Nothing Then
        Debug.Print "GetInsulationResistivityList: table Table6_1 not found"
        GetInsulationResistivityList = ""
        Exit Function
    End If
    
    ' Column 2 contains the Rins0 values (1.00E+06, 1.00E+05, ...)
    Dim vals As Variant
    vals = tbl.ListColumns(2).DataBodyRange.Value
    
    If Not IsArray(vals) Then
        GetInsulationResistivityList = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            If IsNumeric(val) Then
                exists = False
                For Each item In col
                    If CDbl(item) = CDbl(val) Then exists = True: Exit For
                Next item
                If Not exists Then col.Add CDbl(val)
            End If
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetInsulationResistivityList = result
    Exit Function
    
CleanExit:
    GetInsulationResistivityList = ""
End Function

' ================================================================
' update wall thickness validation (preserve current value)
' ================================================================
Private Sub UpdateWallThicknessValidation(ByVal diamCell As Range, Optional ByVal keepValue As Boolean = False)
    On Error GoTo CleanExit
    
    If IsEmpty(diamCell.Value) Then
        Exit Sub
    End If
    
    If Not IsNumeric(diamCell.Value) Then
        Exit Sub
    End If
    
    Dim diameter As Double
    diameter = CDbl(diamCell.Value)
    
    If diameter <= 0 Then
        Exit Sub
    End If
    
    Dim wsPipe As Worksheet
    Set wsPipe = Me
    
    Dim col As Long
    col = diamCell.Column
    
    Dim thicknessList As Variant
    thicknessList = GetThicknessHeadersForDiameter(diameter)
    
    Dim thicknessCell As Range
    Set thicknessCell = wsPipe.Cells(8, col)
    
    Dim currentValue As Variant
    If keepValue Then
        currentValue = thicknessCell.Value
    End If
    
    If IsArray(thicknessList) Then
        Dim listSize As Long
        On Error Resume Next
        listSize = UBound(thicknessList) - LBound(thicknessList) + 1
        On Error GoTo 0
        
        If listSize > 0 Then
            Dim validationString As String
            Dim idx As Long
            validationString = ""
            For idx = LBound(thicknessList) To UBound(thicknessList)
                If validationString <> "" Then
                    validationString = validationString & ","
                End If
                validationString = validationString & CStr(thicknessList(idx))
            Next idx
            
            ' unprotect the sheet
            wsPipe.Unprotect Password:="ptm"
            
            On Error Resume Next
            thicknessCell.Validation.Delete
            On Error GoTo 0
            
            With thicknessCell.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=validationString
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = False
            End With
            
            If keepValue And Not IsEmpty(currentValue) Then
                Dim valueExists As Boolean
                valueExists = False
                Dim item As Variant
                For Each item In thicknessList
                    If CStr(item) = CStr(currentValue) Then
                        valueExists = True
                        Exit For
                    End If
                Next item
                
                If valueExists Then
                    thicknessCell.Value = currentValue
                Else
                    thicknessCell.Value = ""
                End If
            Else
                thicknessCell.Value = ""
            End If
            
            ' restore sheet protection
            wsPipe.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                            AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
        End If
    End If

CleanExit:
End Sub

' ================================================================
' fill steel resistivity based on steel grade
' ================================================================
Private Sub UpdateSteelResistivity(ByVal targetCell As Range)
    On Error GoTo CleanExit
    
    If IsEmpty(targetCell.Value) Then
        Exit Sub
    End If
    
    Dim steelGrade As String
    steelGrade = Trim(CStr(targetCell.Value))
    
    If steelGrade = "" Then
        Exit Sub
    End If
    
    Dim wsPipe As Worksheet
    Dim col As Long
    Dim wsRef As Worksheet
    Dim tblRef As ListObject
    Dim foundRow As Variant
    Dim resistivityValue As Variant
    
    Set wsPipe = Me
    col = targetCell.Column
    
    On Error Resume Next
    Set wsRef = thisWorkbook.Worksheets(SHEET_STO_REF)
    If wsRef Is Nothing Then
        ' лист 'data sto gazprom 9.2-003-2020' не найден!
        ' лandст 'data sto gazprom 9.2-003-2020' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027 0064 0061 0074 0061 0020 0073 0074 006F 0020 0067 0061 007A 0070 0072 006F 006D 0020 0039 002E 0032 002D 0030 0030 0033 002D 0032 0030 0032 0030") & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        Exit Sub
    End If
    
    Set tblRef = wsRef.ListObjects("Table6_2")
    If tblRef Is Nothing Then
        ' таблица 'table6.2' не найдена!
        ' table 'table6.2' not found!
        MsgBox Ru("0442 0430 0431 043B 0438 0446 0430 0020 0027 0074 0061 0062 006C 0065 0036 002E 0032 0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 0021"), vbExclamation
        Exit Sub
    End If
    On Error GoTo CleanExit
    
    On Error Resume Next
    foundRow = Application.Match(steelGrade, tblRef.ListColumns(1).DataBodyRange, 0)
    On Error GoTo CleanExit
    
    If IsError(foundRow) Then
        wsPipe.Cells(6, col).Value = ""
        Exit Sub
    End If
    
    resistivityValue = tblRef.ListColumns(2).DataBodyRange.Cells(foundRow, 1).Value
    
    If Not IsEmpty(resistivityValue) And IsNumeric(resistivityValue) Then
        wsPipe.Cells(6, col).Value = CDbl(resistivityValue)
    Else
        wsPipe.Cells(6, col).Value = ""
    End If

CleanExit:
End Sub

' ================================================================
' refresh thickness list when diameter changes
' ================================================================
Private Sub UpdateThicknessFromTableB(ByVal diamCell As Range)
    On Error GoTo CleanExit
    
    If IsEmpty(diamCell.Value) Then
        Exit Sub
    End If
    
    If Not IsNumeric(diamCell.Value) Then
        Exit Sub
    End If
    
    Dim diameter As Double
    diameter = CDbl(diamCell.Value)
    
    If diameter <= 0 Then
        Exit Sub
    End If
    
    Dim wsPipe As Worksheet
    Set wsPipe = Me
    
    Dim col As Long
    col = diamCell.Column
    
    wsPipe.Cells(8, col).Value = ""
    Call UpdateWallThicknessValidation(diamCell, False)

CleanExit:
End Sub

' ================================================================
' clear extra columns
' ================================================================
Private Sub ClearExtraColumns()
    On Error GoTo CleanExit
    
    Dim wsPipe As Worksheet
    Set wsPipe = Me
    
    Dim maxNamedCol As Long
    maxNamedCol = 0
    
    Dim nm As name
    Dim rng As Range
    
    For Each nm In wsPipe.names
        On Error Resume Next
        Set rng = wsPipe.Range(nm.name)
        If Not rng Is Nothing Then
            Dim rngLastCol As Long
            rngLastCol = rng.Column + rng.Columns.count - 1
            If rngLastCol > maxNamedCol Then
                maxNamedCol = rngLastCol
            End If
        End If
        On Error GoTo 0
    Next nm
    
    If maxNamedCol > 0 Then
        Dim clearStartCol As Long
        clearStartCol = maxNamedCol + 1
        
        Dim lastDataCol As Long
        On Error Resume Next
        lastDataCol = wsPipe.Cells.Find(What:="*", _
                                        After:=wsPipe.Cells(1, 1), _
                                        LookIn:=xlFormulas, _
                                        LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlPrevious, _
                                        MatchCase:=False).Column
        On Error GoTo 0
        
        If lastDataCol = 0 Then
            lastDataCol = wsPipe.Columns.count
        End If
        
        If clearStartCol <= lastDataCol Then
            Dim clearRange As Range
            Set clearRange = wsPipe.Range(wsPipe.Columns(clearStartCol), wsPipe.Columns(lastDataCol))
            clearRange.Clear
            clearRange.EntireColumn.Hidden = True
        End If
    End If

CleanExit:
End Sub

' ================================================================
' create validation for:
'   1. pipeSteelGrade (list from Table6_2)
'   2. pipeDiameter (list from TableB)
'   3. pipeResistivityChangeFactor (fixed list: 0.05, 0.08, 0.11)
'   4. pipeInsulationResistivityStartLife (list from Table6_1)
' ================================================================
Private Sub CreatePipeValidation()
    On Error GoTo CleanExit
    
    Dim wsPipe As Worksheet
    Set wsPipe = Me
    
    ' unprotect the sheet before modifying validation
    Dim wasProtected As Boolean
    wasProtected = wsPipe.ProtectContents
    If wasProtected Then
        wsPipe.Unprotect Password:="ptm"
    End If
    
    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = wsPipe.Range("pipeDifferentParametersNum").Value
    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If
    
    Dim startCol As Long
    startCol = 4
    Dim endCol As Long
    endCol = startCol + colBr - 1
    Dim col As Long
    
    ' --- 1. validation for pipeSteelGrade (row 5) ---
    Dim steelList As String
    steelList = GetSteelGradeList()
    If steelList <> "" Then
        For col = startCol To endCol
            On Error Resume Next
            With wsPipe.Cells(5, col).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=steelList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = False
            End With
            If Err.Number <> 0 Then
                Debug.Print "steel validation error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo CleanExit
        Next col
        Debug.Print "steel validation created for " & colBr & " columns"
    End If
    
    ' --- 2. validation for pipeDiameter (row 7) ---
    Dim diamList As String
    diamList = GetDiameterList()
    If diamList <> "" Then
        For col = startCol To endCol
            On Error Resume Next
            With wsPipe.Cells(7, col).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=diamList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = False
            End With
            If Err.Number <> 0 Then
                Debug.Print "diameter validation error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo CleanExit
        Next col
        Debug.Print "diameter validation created for " & colBr & " columns"
    End If
    
    ' --- 3. validation for pipeResistivityChangeFactor (fixed list: 0.05, 0.08, 0.11) ---
    Dim decSep As String
    decSep = Application.International(xlDecimalSeparator)
    
    Dim insListFixed As String
    insListFixed = "0" & decSep & "05," & _
                   "0" & decSep & "08," & _
                   "0" & decSep & "11"
    
    Dim changeFactorRow As Long
    On Error Resume Next
    changeFactorRow = wsPipe.Range("pipeResistivityChangeFactor").row
    On Error GoTo CleanExit
    If changeFactorRow < 1 Then changeFactorRow = 12
    
    For col = startCol To endCol
        On Error Resume Next
        With wsPipe.Cells(changeFactorRow, col).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=insListFixed
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        If Err.Number <> 0 Then
            Debug.Print "pipeResistivityChangeFactor validation error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo CleanExit
    Next col
    Debug.Print "pipeResistivityChangeFactor validation created for " & colBr & " columns (row=" & changeFactorRow & ")"
    
    ' --- 4. validation for pipeInsulationResistivityStartLife (Rins0 from Table6_1) ---
    Dim insListTable As String
    insListTable = GetInsulationResistivityList()
    If insListTable <> "" Then
        Dim insulationRow As Long
        On Error Resume Next
        insulationRow = wsPipe.Range("pipeInsulationResistivityStartLife").row
        On Error GoTo CleanExit
        If insulationRow < 1 Then insulationRow = 11
        
        For col = startCol To endCol
            On Error Resume Next
            With wsPipe.Cells(insulationRow, col).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=insListTable
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = False
                .ShowError = False
            End With
            If Err.Number <> 0 Then
                Debug.Print "pipeInsulationResistivityStartLife validation error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo CleanExit
        Next col
        Debug.Print "pipeInsulationResistivityStartLife validation created for " & colBr & " columns (row=" & insulationRow & ")"
    Else
        Debug.Print "GetInsulationResistivityList returned empty list - pipeInsulationResistivityStartLife validation not created"
    End If
    
    ' restore sheet protection
    If wasProtected Then
        wsPipe.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                        AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
    End If

CleanExit:
    If Err.Number <> 0 Then
        Debug.Print "error in createpipevalidation: " & Err.Description
    End If
End Sub

' ================================================================
' get steel grade list from STO reference
' ================================================================
Private Function GetSteelGradeList() As String
    On Error GoTo CleanExit
    
    Dim wsSTO As Worksheet
    Set wsSTO = thisWorkbook.Worksheets(SHEET_STO_REF)
    If wsSTO Is Nothing Then
        GetSteelGradeList = ""
        Exit Function
    End If
    
    Dim tbl As ListObject
    Set tbl = wsSTO.ListObjects("Table6_2")
    If tbl Is Nothing Then
        GetSteelGradeList = ""
        Exit Function
    End If
    
    Dim vals As Variant
    vals = tbl.ListColumns(1).DataBodyRange.Value
    
    If Not IsArray(vals) Then
        GetSteelGradeList = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetSteelGradeList = result
    Exit Function
    
CleanExit:
    GetSteelGradeList = ""
End Function

' ================================================================
' get diameter list from STO reference (tableB)
' ================================================================
Private Function GetDiameterList() As String
    On Error GoTo CleanExit
    
    Dim wsSTO As Worksheet
    Set wsSTO = thisWorkbook.Worksheets(SHEET_STO_REF)
    If wsSTO Is Nothing Then
        GetDiameterList = ""
        Exit Function
    End If
    
    Dim tbl As ListObject
    Set tbl = wsSTO.ListObjects("TableB")
    If tbl Is Nothing Then
        GetDiameterList = ""
        Exit Function
    End If
    
    Dim vals As Variant
    vals = tbl.ListColumns(1).DataBodyRange.Value
    
    If Not IsArray(vals) Then
        GetDiameterList = CStr(vals)
        Exit Function
    End If
    
    Dim col As New Collection
    Dim i As Long
    Dim val As Variant
    Dim item As Variant
    Dim result As String
    Dim exists As Boolean
    
    For i = 1 To UBound(vals, 1)
        val = vals(i, 1)
        If Not IsEmpty(val) And val <> "" Then
            exists = False
            For Each item In col
                If item = val Then exists = True: Exit For
            Next item
            If Not exists Then col.Add val
        End If
    Next i
    
    For Each item In col
        If result <> "" Then result = result & "," & item Else result = item
    Next item
    
    GetDiameterList = result
    Exit Function
    
CleanExit:
    GetDiameterList = ""
End Function

' ================================================================
' synchronize ranges, data, formats and validation
' ================================================================
Public Sub SyncPipeRangesAndColumns(ByVal colBr As Long)
    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wsPipe As Worksheet
    Set wsPipe = Me

    ' unprotect the sheet
    wsPipe.Unprotect Password:="ptm"

    If colBr < 1 Then colBr = 1
    If colBr > 254 Then colBr = 254

    Dim startCol As Long
    startCol = 4

    ' ================================================================
    ' 1. determine current number of columns
    ' ================================================================
    Dim actualCols As Long
    Dim impRng As Range
    
    On Error Resume Next
    Set impRng = wsPipe.Range("pipeImpedance")
    On Error GoTo 0
    
    If Not impRng Is Nothing Then
        actualCols = impRng.Columns.count
    Else
        actualCols = 1
    End If

    ' ================================================================
    ' 2. list of all named ranges
    ' ================================================================
    Dim rangeNames As Variant
    rangeNames = Array( _
        "pipeSteelGrade", "pipeSteelResistivity", "pipeDiameter", _
        "pipeWallThickness", "pipeInsulationResistivityStartLife", "pipeLayingDepth", _
        "soilResistivityAvg", "serviceLifeDesigned", "pipeResistivityChangeFactor", _
        "pipeAlongResistance", "soilResistanceAroundPipe", "pipeTransientResistivity", _
        "soilResistivityAroundPipe", "pipeSteelResistivity", "pipeInsulationResistanceStartLife", _
        "pipeTransientResistance", "pipeTransientResistanceEndLife", "factorPropagationCurrentAlongPipe", _
        "factorPropagationCurrentAlongPipeEndLife", "pipeImpedance", "pipeImpedanceEndLife" _
    )

    ' ================================================================
    ' 3. update all named ranges
    ' ================================================================
    Dim rngName As Variant
    Dim rng As Range
    Dim newRng As Range
    Dim startRow As Long
    Dim currentCols As Long
    Dim updatedCount As Long
    Dim totalCount As Long
    updatedCount = 0
    totalCount = UBound(rangeNames) - LBound(rangeNames) + 1

    For Each rngName In rangeNames
        If updatedCount Mod 2 = 0 And updatedCount > 0 Then
            Application.StatusBar = "sync " & colBr & " columns... " & _
                                    Format(updatedCount / totalCount * 100, "0") & "%"
            DoEvents
        End If
        
        On Error Resume Next
        Set rng = wsPipe.Range(rngName)
        On Error GoTo 0
        
        If Not rng Is Nothing Then
            startRow = rng.row
            currentCols = rng.Columns.count
            
            If currentCols <> colBr Then
                Set newRng = wsPipe.Range(wsPipe.Cells(startRow, startCol), _
                                          wsPipe.Cells(startRow, startCol + colBr - 1))
                
                On Error Resume Next
                wsPipe.names(rngName).Delete
                On Error GoTo 0
                
                wsPipe.names.Add name:=rngName, RefersTo:=newRng
                
                updatedCount = updatedCount + 1
            Else
                updatedCount = updatedCount + 1
            End If
        Else
            updatedCount = updatedCount + 1
        End If
    Next rngName
    
    Application.StatusBar = "sync " & colBr & " columns... 100%"
    DoEvents

    ' ================================================================
    ' 4. copy formats and clear data when growing
    ' ================================================================
    If colBr > actualCols Then
        Dim sourceCol As Long
        sourceCol = startCol + actualCols - 1
        If sourceCol < startCol Then sourceCol = startCol
        
        Dim newCol As Long
        Dim progressTotal As Long
        progressTotal = colBr - actualCols
        Dim progressStep As Long
        progressStep = 0
        
        For newCol = sourceCol + 1 To startCol + colBr - 1
            progressStep = progressStep + 1
            Application.StatusBar = "copying columns... " & _
                                    Format(progressStep / progressTotal * 100, "0") & "%"
            DoEvents
            
            ' 1. copy the whole column (data + formats)
            wsPipe.Columns(sourceCol).Copy Destination:=wsPipe.Columns(newCol)
            
            ' 2. update the header
            ' плечо 
            ' section 
            wsPipe.Cells(4, newCol).Value = Ru("043F 043B 0435 0447 043E 0020") & (newCol - startCol + 1)
            
            ' ================================================================
            ' 3. clear scalar cells (fully)
            ' ================================================================
            
            ' --- pipedifferentparametersnum (row 3) ---
            With wsPipe.Cells(3, newCol)
                .ClearContents
                .ClearFormats
                .Validation.Delete
            End With
            
            ' --- pipeinputresistance (row 28) ---
            With wsPipe.Cells(28, newCol)
                .ClearContents
                .ClearFormats
                .Validation.Delete
            End With
            
            ' --- pipeinputresistanceendlife (row 29) ---
            With wsPipe.Cells(29, newCol)
                .ClearContents
                .ClearFormats
                .Validation.Delete
            End With
            
            ' ================================================================
            ' 4. clear data only in rows 17-27 (keep formats and validation)
            ' ================================================================
            Dim rowIdx As Long
            For rowIdx = 17 To 27
                wsPipe.Cells(rowIdx, newCol).ClearContents
            Next rowIdx
            
            sourceCol = newCol
        Next newCol
        
        Application.StatusBar = "copying columns... 100%"
        DoEvents
    End If

    ' ================================================================
    ' 5. shrink: clear and hide extra columns
    ' ================================================================
    Dim endCol As Long
    endCol = startCol + colBr - 1
    Dim lastColAll As Long
    lastColAll = wsPipe.Columns.count
    
    If endCol < 1 Then endCol = 1
    If endCol > 254 Then endCol = 254
    If lastColAll < 1 Then lastColAll = 1
    If lastColAll > 254 Then lastColAll = 254
    
    If colBr < actualCols Then
        Application.StatusBar = "clearing extra columns..."
        DoEvents
        
        Dim clearStartCol As Long
        clearStartCol = startCol + colBr
        
        If clearStartCol < 1 Then clearStartCol = 1
        If clearStartCol > 254 Then clearStartCol = 254
        
        If clearStartCol <= lastColAll And clearStartCol >= 1 And clearStartCol <= 254 Then
            On Error Resume Next
            wsPipe.Range(wsPipe.Columns(clearStartCol), wsPipe.Columns(lastColAll)).Clear
            wsPipe.Range(wsPipe.Columns(clearStartCol), wsPipe.Columns(lastColAll)).EntireColumn.Hidden = True
            wsPipe.Range(wsPipe.Columns(clearStartCol), wsPipe.Columns(lastColAll)).EntireColumn.ColumnWidth = 8.43
            On Error GoTo 0
        End If
        
        Application.StatusBar = "clearing extra columns... done"
        DoEvents
    End If

    ' ================================================================
    ' 6. refresh wall thickness validation
    ' ================================================================
    Application.StatusBar = "updating wall thickness validation..."
    DoEvents
    
    Dim colIdx As Long
    For colIdx = startCol To endCol
        Dim diamCell As Range
        Set diamCell = wsPipe.Cells(7, colIdx)
        
        If Not IsEmpty(diamCell.Value) And IsNumeric(diamCell.Value) Then
            If CDbl(diamCell.Value) > 0 Then
                Call UpdateWallThicknessValidation(diamCell, True)
            End If
        End If
    Next colIdx

    ' ================================================================
    ' 7. update headers
    ' ================================================================
    Dim i As Long
    For i = 1 To colBr
        ' плечо 
        ' section 
        wsPipe.Cells(4, startCol + i - 1).Value = Ru("043F 043B 0435 0447 043E 0020") & i
    Next i

    ' ================================================================
    ' 8. update d3
    ' ================================================================
    On Error Resume Next
    wsPipe.Range("pipeDifferentParametersNum").Value = colBr
    On Error GoTo 0

    ' ================================================================
    ' 9. recalculate
    ' ================================================================
    Application.StatusBar = "recalculating..."
    DoEvents
    
    wsPipe.Calculate
    
    ' ================================================================
    ' 10. create validation for steel, diameter, Rins0
    ' ================================================================
    Application.StatusBar = "creating validation for steel, diameter, Riz0..."
    DoEvents
    Call CreatePipeValidation
    
    Application.StatusBar = "pipe calculation updated: " & colBr & " columns"

CleanExit:
    ' restore sheet protection
    On Error Resume Next
    wsPipe.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                    AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ================================================================
' procedure for external calls
' ================================================================
Public Sub ForceSyncFromD3()
    Dim wsPipe As Worksheet
    Set wsPipe = Me
    
    Dim colBr As Long
    Dim pipeVal As Variant
    
    On Error Resume Next
    pipeVal = wsPipe.Range("pipeDifferentParametersNum").Value
    On Error GoTo 0
    
    If Not IsNumeric(pipeVal) Or pipeVal <= 0 Then
        pipeVal = 1
    End If
    
    colBr = CLng(pipeVal)
    If colBr < 1 Then colBr = 1
    
    Call SyncPipeRangesAndColumns(colBr)
End Sub

' ================================================================
' sheet activate event
' ================================================================
Private Sub Worksheet_Activate()
    Call CreatePipeValidation
End Sub

' ================================================================
' sheet change event
' ================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    If Application.EnableEvents = False Then
        Application.EnableEvents = True
        Debug.Print "enableevents was false, restored in worksheet_change (pipe)"
    End If
    If Not Application.EnableEvents Then Application.EnableEvents = True
    On Error GoTo CleanExit
    
    If isUpdating Then Exit Sub

    ' ================================================================
    ' 1. check for change of pipedifferentparametersnum (d3)
    ' ================================================================
    Dim pipeCell As Range
    On Error Resume Next
    Set pipeCell = Me.Range("pipeDifferentParametersNum")
    If pipeCell Is Nothing Then
        Set pipeCell = Me.Range("D3")
    End If
    On Error GoTo 0

    If Not pipeCell Is Nothing Then
        If Not Intersect(Target, pipeCell) Is Nothing Then
            Dim newVal As Variant
            newVal = pipeCell.Value
            
            If Not IsNumeric(newVal) Or newVal = "" Or newVal <= 0 Then
                newVal = 1
                isUpdating = True
                Application.EnableEvents = False
                pipeCell.Value = 1
                Application.EnableEvents = True
                isUpdating = False
                GoTo CleanExit
            End If
            
            Dim colBr As Long
            colBr = CLng(newVal)
            If colBr < 1 Then colBr = 1
            If colBr > 254 Then colBr = 254
            
            ' ================================================================
            ' warn if the number of arms is > 1000
            ' ================================================================
            If colBr > 1000 Then
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
            
            isUpdating = True
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            
            ' progress indication
            Application.StatusBar = "starting synchronization of " & colBr & " columns..."
            DoEvents
            
            Call SyncPipeRangesAndColumns(colBr)
            
            ' --- create validation for steel, diameter, Rins0 ---
            Call CreatePipeValidation
            
            Application.StatusBar = "synchronization finished for " & colBr & " columns"
            DoEvents
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            isUpdating = False
            
            GoTo CleanExit
        End If
    End If

    ' ================================================================
    ' 2. steel grade change (row 5)
    ' ================================================================
    Dim pipeCountValue As Variant
    On Error Resume Next
    pipeCountValue = pipeCell.Value
    On Error GoTo 0

    Dim colBrOther As Long
    If IsNumeric(pipeCountValue) Then
        colBrOther = CLng(pipeCountValue)
    Else
        colBrOther = 1
    End If
    If colBrOther < 1 Then colBrOther = 1

    Dim startCol As Long, endCol As Long
    startCol = 4
    endCol = startCol + colBrOther - 1
    If endCol < startCol Then endCol = startCol
    If endCol > 254 Then endCol = 254

    Dim steelRange As Range
    Set steelRange = Me.Range(Me.Cells(5, startCol), Me.Cells(5, endCol))
    If Not Intersect(Target, steelRange) Is Nothing Then
        If Target.CountLarge > 1 Then
            Dim cellSteel As Range
            For Each cellSteel In Intersect(Target, steelRange)
                Call UpdateSteelResistivity(cellSteel)
            Next cellSteel
        Else
            Call UpdateSteelResistivity(Target)
        End If
        GoTo CleanExit
    End If

    ' ================================================================
    ' 3. diameter change (row 7)
    ' ================================================================
    Dim colBrDiam As Long
    On Error Resume Next
    pipeCountValue = pipeCell.Value
    On Error GoTo 0

    If IsNumeric(pipeCountValue) Then
        colBrDiam = CLng(pipeCountValue)
    Else
        colBrDiam = 1
    End If
    If colBrDiam < 1 Then colBrDiam = 1

    startCol = 4
    endCol = startCol + colBrDiam - 1
    If endCol < startCol Then endCol = startCol
    If endCol > 254 Then endCol = 254

    Dim diamRange As Range
    Set diamRange = Me.Range(Me.Cells(7, startCol), Me.Cells(7, endCol))
    If Not Intersect(Target, diamRange) Is Nothing Then
        If Target.CountLarge > 1 Then
            Dim cellDiam As Range
            For Each cellDiam In Intersect(Target, diamRange)
                Call UpdateThicknessFromTableB(cellDiam)
            Next cellDiam
        Else
            Call UpdateThicknessFromTableB(Target)
        End If
        GoTo CleanExit
    End If

    ' ================================================================
    ' 4. recalculate
    ' ================================================================
    Me.Calculate

CleanExit:
    Application.EnableEvents = True
    isUpdating = False
    Application.StatusBar = False
End Sub
