Attribute VB_Name = "Module_Helpers"

' ================================================================
' helper-functions module
' ================================================================
Option Explicit

' StatusBar on Windows Excel is ANSI, not Unicode.
' Mac-saved Russian literals become mojibake there. Keep StatusBar ASCII.
Public Sub SetStatusBar(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

' StatusBar progress for long loops. Updates on first, last, and every 5th step.
Public Sub PulseProgress(ByVal phase As String, ByVal current As Long, ByVal total As Long)
    On Error Resume Next
    If total <= 0 Then
        Application.StatusBar = phase
        DoEvents
        Exit Sub
    End If
    If current = 1 Or current = total Or (current Mod 5) = 0 Then
        Application.StatusBar = phase & " " & current & " / " & total
        DoEvents
    End If
    On Error GoTo 0
End Sub

' ================================================================
' global flag to prevent recalculating maxProtectPotential
' ================================================================
' Public gSkipMaxPotRecalc As Boolean

' safely get a numeric value
Public Function SafeNumericValue(Value As Variant, Optional defaultValue As Long = 1) As Long
    If IsNumeric(Value) Then
        If Value > 0 Then
            SafeNumericValue = CLng(Value)
        Else
            SafeNumericValue = defaultValue
        End If
    Else
        SafeNumericValue = defaultValue
    End If
    
    ' clamp to a range
    SafeNumericValue = Application.WorksheetFunction.Max(MIN_COL, _
                         Application.WorksheetFunction.Min(MAX_COL, SafeNumericValue))
End Function

' check that the sheet exists
Public Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not (thisWorkbook.Worksheets(wsName) Is Nothing)
    On Error GoTo 0
End Function

' check that the table exists
Public Function TableExists(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    TableExists = Not (ws.ListObjects(tableName) Is Nothing)
    On Error GoTo 0
End Function

' safely get a value from a range
Public Function SafeRangeValue(rng As Range) As Variant
    On Error Resume Next
    SafeRangeValue = rng.Value
    If Err.Number <> 0 Then
        SafeRangeValue = ""
    End If
    On Error GoTo 0
End Function

' clear a range with a safety check
Public Sub SafeClearRange(rng As Range)
    On Error Resume Next
    If Not rng Is Nothing Then
        rng.ClearContents
    End If
    On Error GoTo 0
End Sub

' ================================================================
' strip illegal characters from a string (replace with "_")
' ================================================================
Public Function CleanStringForTable(ByVal str As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    
    If IsNull(str) Or IsEmpty(str) Then
        CleanStringForTable = ""
        Exit Function
    End If
    
    result = ""
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        code = AscW(ch)
        
        Dim isAllowed As Boolean
        isAllowed = False
        
        ' letters: Latin (a-z, A-Z) and Cyrillic (A-Ya, a-ya)
        If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
           (code >= 1040 And code <= 1103) Or (code = 1025) Or (code = 1105) Then
            isAllowed = True
        ' digits
        ElseIf (code >= 48 And code <= 57) Then
            isAllowed = True
        ' allow spaces
        ElseIf code = 32 Then
            isAllowed = True
        ' allowed punctuation (no ! @ #)
        ElseIf InStr(1, "_-.,()/\+=*&$?<>", ch) > 0 Then
            isAllowed = True
        End If
        
        If isAllowed Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    
    ' --- collapse repeated spaces (but do not remove all spaces) ---
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' --- important: do not trim leading/trailing spaces ---
    CleanStringForTable = result
End Function

Sub EnableEventsOn()
    Application.EnableEvents = True
    ' события включены. теперь изменение pipecountcp должно обновлять колонки.
    ' events enabled. теперь ofмененandе pipecountcp должно обновлять колонкand.
    MsgBox Ru("0441 043E 0431 044B 0442 0438 044F 0020 0432 043A 043B 044E 0447 0435 043D 044B 002E 0020 0442 0435 043F 0435 0440 044C 0020 0438 0437 043C 0435 043D 0435 043D 0438 0435 0020") & Ru("0070 0069 0070 0065 0043 006F 0075 006E 0074 0043 0050 0020 0434 043E 043B 0436 043D 043E 0020 043E 0431 043D 043E 0432 043B 044F 0442 044C 0020 043A 043E 043B 043E 043D 043A") & Ru("0438 002E"), vbInformation
End Sub

' ================================================================
' safely read an array from a range (fixed for 1x1)
' ================================================================
Public Function SafeGetArray(ByVal rng As Range) As Variant
    On Error GoTo CleanExit
    
    Dim vals As Variant
    Dim arr() As Variant
    Dim rows As Long, cols As Long
    Dim i As Long
    
    If rng Is Nothing Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = Empty
        SafeGetArray = arr
        Exit Function
    End If
    
    vals = rng.Value
    
    ' if it is not an array (single cell), build a 1x1 array
    If Not IsArray(vals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = vals
        SafeGetArray = arr
        Exit Function
    End If
    
    ' determine the rank
    On Error Resume Next
    rows = UBound(vals, 1)
    cols = UBound(vals, 2)
    
    If Err.Number <> 0 Then
        ' this is a 1D array - convert to 1xN
        Err.Clear
        ReDim arr(1 To 1, 1 To UBound(vals))
        For i = LBound(vals) To UBound(vals)
            arr(1, i - LBound(vals) + 1) = vals(i)
        Next i
        SafeGetArray = arr
        Exit Function
    End If
    On Error GoTo CleanExit
    
    ' ================================================================
    ' !!! important: return a 1x1 array, not a scalar !!!
    ' ================================================================
    If rows = 1 And cols = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = vals(1, 1)
        SafeGetArray = arr
        Exit Function
    End If
    
    ' === if rows > 1 and cols = 1 - vertical array ===
    If cols = 1 And rows > 1 Then
        ' convert to 1xN (transpose)
        ReDim arr(1 To 1, 1 To rows)
        For i = 1 To rows
            arr(1, i) = vals(i, 1)
        Next i
        SafeGetArray = arr
        Exit Function
    End If
    
    ' === if rows = 1 and cols > 1 - horizontal array ===
    If rows = 1 And cols > 1 Then
        ' leave as is - 1xN
        SafeGetArray = vals
        Exit Function
    End If
    
    ' === multi-row array NxM ===
    ' return as is
    SafeGetArray = vals
    Exit Function
    
CleanExit:
    ' on error return an empty 1x1 array
    ReDim arr(1 To 1, 1 To 1)
    arr(1, 1) = Empty
    SafeGetArray = arr
End Function

' ================================================================
' generic helper: read a range as a 1xN array
' ================================================================
' converts any range into a 1xN array
' supports:
'   - scalar (one cell) -> 1x1 array
'   - horizontal 1xN range -> 1xN array
'   - vertical Nx1 range -> 1xN array (transpose)
'   - multi-row NxM range -> first row as a 1xM array
' ================================================================
Public Function ReadRangeAs1xN(ByVal rng As Range) As Variant
    On Error GoTo ErrorHandler
    
    Dim result As Variant
    Dim vals As Variant
    Dim i As Long, j As Long
    Dim rowsCount As Long, colsCount As Long
    
    ' check: the range exists
    If rng Is Nothing Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = Empty
        ReadRangeAs1xN = result
        Exit Function
    End If
    
    ' read values
    vals = rng.Value
    
    ' === case 1: scalar (one cell) ===
    If Not IsArray(vals) Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = vals
        ReadRangeAs1xN = result
        Exit Function
    End If
    
    ' determine the array rank
    On Error Resume Next
    rowsCount = UBound(vals, 1)
    colsCount = UBound(vals, 2)
    If Err.Number <> 0 Then
        ' this is a 1D array (rare) - convert it
        Err.Clear
        ReDim result(1 To 1, 1 To UBound(vals))
        For i = LBound(vals) To UBound(vals)
            result(1, i - LBound(vals) + 1) = vals(i)
        Next i
        ReadRangeAs1xN = result
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' === case 2: horizontal 1xN array ===
    If rowsCount = 1 And colsCount >= 1 Then
        ReadRangeAs1xN = vals
        Exit Function
    End If
    
    ' === case 3: vertical Nx1 array ===
    If colsCount = 1 And rowsCount >= 1 Then
        ReDim result(1 To 1, 1 To rowsCount)
        For i = 1 To rowsCount
            result(1, i) = vals(i, 1)
        Next i
        ReadRangeAs1xN = result
        Exit Function
    End If
    
    ' === case 4: multi-row, multi-column NxM ===
    ' take the first row as a 1xM array
    ReDim result(1 To 1, 1 To colsCount)
    For j = 1 To colsCount
        result(1, j) = vals(1, j)
    Next j
    ReadRangeAs1xN = result
    Exit Function
    
ErrorHandler:
    ' on error return an empty 1x1 array
    ReDim result(1 To 1, 1 To 1)
    result(1, 1) = Empty
    ReadRangeAs1xN = result
End Function

' ================================================================
' extra helper: find the maximum in a 1xN array
' ================================================================
Public Function GetMaxFrom1xN(ByVal arr As Variant, Optional ByVal defaultValue As Double = 30) As Double
    On Error GoTo ErrorHandler
    
    Dim maxVal As Double
    Dim i As Long
    Dim val As Variant
    
    ' check: is this an array
    If Not IsArray(arr) Then
        If IsNumeric(arr) Then
            GetMaxFrom1xN = CDbl(arr)
        Else
            GetMaxFrom1xN = defaultValue
        End If
        Exit Function
    End If
    
    ' check dimensions
    On Error Resume Next
    Dim rowsCount As Long, colsCount As Long
    rowsCount = UBound(arr, 1)
    colsCount = UBound(arr, 2)
    If Err.Number <> 0 Then
        ' 1D array
        Err.Clear
        maxVal = -1E+308
        For i = LBound(arr) To UBound(arr)
            val = arr(i)
            If IsNumeric(val) Then
                If CDbl(val) > maxVal Then maxVal = CDbl(val)
            End If
        Next i
        If maxVal = -1E+308 Then maxVal = defaultValue
        GetMaxFrom1xN = maxVal
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' 2D array - search in the first row
    maxVal = -1E+308
    For i = 1 To colsCount
        val = arr(1, i)
        If IsNumeric(val) Then
            If CDbl(val) > maxVal Then maxVal = CDbl(val)
        End If
    Next i
    
    If maxVal = -1E+308 Then maxVal = defaultValue
    GetMaxFrom1xN = maxVal
    Exit Function
    
ErrorHandler:
    GetMaxFrom1xN = defaultValue
End Function

' ================================================================
' diagnostics: check all named ranges (in Module_Helpers)
' ================================================================
Sub DiagnoseNamedRanges()
    On Error Resume Next
    
    ' === диагностика именованных диапазонов ===
    ' === named range diagnostics ===
    Debug.Print Ru("003D 003D 003D 0020 0434 0438 0430 0433 043D 043E 0441 0442 0438 043A 0430 0020 0438 043C 0435 043D 043E 0432 0430 043D 043D 044B 0445 0020 0434 0438 0430 043F 0430 0437 043E") & Ru("043D 043E 0432 0020 003D 003D 003D")
    ' время: 
    ' time: 
    Debug.Print Ru("0432 0440 0435 043C 044F 003A 0020") & Now
    Debug.Print ""
    
    Dim namesToCheck As Variant
    namesToCheck = Array( _
        "AG_Material", "AG_MountType", "AG_Completion", "AG_Model", _
        "AG_Diameter", "AG_Length", "AG_CokeDiam", "AG_CokeLength", _
        "AG_Mass", "AG_DissolutionRate", "AG_RatedCurrent", _
        "AG_ResistivityMaterial", "AG_CokeResistivity", _
        "AG_SpecificRatedCurrent", "AG_SpecificMass", _
        "typeMaterial", "typeMountingAG", "typeInstallationAG", _
        "typeDeliveryAG", "typeAG", "diameterAG", "lengthElectrodeAG", _
        "cokeBreezeDiameterAG", "cokeBreezelengthElectrodeAG", _
        "massOneElectrodeAG", "dissolutionRateAG", "ratedCurrent", _
        "resistivityMaterialAG", "cokeBreezeResistivityAG", _
        "specificRatedCurrent", "specificMaccOneMeterAG" _
    )
    
    Dim name As Variant
    Dim rng As Range
    Dim foundCount As Long
    Dim missingCount As Long
    Dim missingNames As String
    
    For Each name In namesToCheck
        Set rng = Nothing
        On Error Resume Next
        Set rng = thisWorkbook.names(name).RefersToRange
        If Err.Number <> 0 Then
            Err.Clear
            ' try as a local name on the Anod sheet
            On Error Resume Next
            Set rng = thisWorkbook.Worksheets("Anod").names(name).RefersToRange
            If Err.Number <> 0 Then
                Err.Clear
                rng = Nothing
            End If
            On Error GoTo 0
        End If
        On Error GoTo 0
        
        If rng Is Nothing Then
            '   ? не найден: 
            '   ? not found: 
            Debug.Print Ru("0020 0020 003F 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 003A 0020") & name
            missingCount = missingCount + 1
            If missingNames <> "" Then missingNames = missingNames & ", "
            missingNames = missingNames & name
        Else
            '   ? найден: 
            '   ? found: 
            '  (лист: 
            '  (лandст: 
            Debug.Print Ru("0020 0020 003F 0020 043D 0430 0439 0434 0435 043D 003A 0020") & name & " -> " & rng.address & Ru("0020 0028 043B 0438 0441 0442 003A 0020") & rng.Parent.name & ")"
            foundCount = foundCount + 1
        End If
    Next name
    
    Debug.Print ""
    ' === итог: найдено 
    ' === summary: found 
    ' , пропущено 
    ' , skipped 
    Debug.Print Ru("003D 003D 003D 0020 0438 0442 043E 0433 003A 0020 043D 0430 0439 0434 0435 043D 043E 0020") & foundCount & Ru("002C 0020 043F 0440 043E 043F 0443 0449 0435 043D 043E 0020") & missingCount & " ==="
    ' === пропущенные имена: 
    ' === missing names: 
    Debug.Print Ru("003D 003D 003D 0020 043F 0440 043E 043F 0443 0449 0435 043D 043D 044B 0435 0020 0438 043C 0435 043D 0430 003A 0020") & missingNames
    
    If missingCount > 0 Then
        ' === рекомендация: запустите createagcolumnnames для восстановления ===
        ' === recommendation: run createagcolumnnames to restore ===
        Debug.Print Ru("003D 003D 003D 0020 0440 0435 043A 043E 043C 0435 043D 0434 0430 0446 0438 044F 003A 0020 0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0043 0072 0065 0061 0074 0065 0041") & Ru("0047 0043 006F 006C 0075 006D 006E 004E 0061 006D 0065 0073 0020 0434 043B 044F 0020 0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 044F 0020 003D 003D 003D")
        ' найдено 
        ' found 
        '  пропущенных имён!
        '  missing names!
        ' список: 
        ' list: 
        ' запустите createagcolumnnames для восстановления.
        ' run createagcolumnnames to restore.
        MsgBox Ru("043D 0430 0439 0434 0435 043D 043E 0020") & missingCount & Ru("0020 043F 0440 043E 043F 0443 0449 0435 043D 043D 044B 0445 0020 0438 043C 0451 043D 0021") & vbCrLf & _
               Ru("0441 043F 0438 0441 043E 043A 003A 0020") & missingNames & vbCrLf & vbCrLf & _
               Ru("0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0043 0072 0065 0061 0074 0065 0041 0047 0043 006F 006C 0075 006D 006E 004E 0061 006D 0065 0073 0020 0434 043B 044F 0020 0432") & Ru("043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 044F 002E"), vbExclamation
    Else
        ' все имена найдены! (
        ' all names found! (
        '  шт.)
        '  pcs)
        MsgBox Ru("0432 0441 0435 0020 0438 043C 0435 043D 0430 0020 043D 0430 0439 0434 0435 043D 044B 0021 0020 0028") & foundCount & Ru("0020 0448 0442 002E 0029"), vbInformation
    End If
End Sub

' ================================================================
' procedure-visibility macro
' ================================================================
Sub PreviewVisibilityChanges()
    On Error Resume Next
    
    Dim wsAnod As Worksheet
    Dim vbComp As VBComponent
    Dim codeMod As CodeModule
    Dim lineNum As Long
    Dim lineText As String
    Dim procName As String
    Dim procType As String
    Dim isPublic As Boolean
    Dim needsPublic As Boolean
    Dim changesCount As Long
    
    Set wsAnod = thisWorkbook.Worksheets("Anod")
    If wsAnod Is Nothing Then Exit Sub
    
    Set vbComp = thisWorkbook.VBProject.VBComponents(wsAnod.CodeName)
    Set codeMod = vbComp.CodeModule
    
    Dim publicProcs As New Collection
    publicProcs.Add "GetMaterialListCross"
    publicProcs.Add "GetMountTypeListCross"
    publicProcs.Add "GetInstallTypeListCross"
    publicProcs.Add "GetDeliveryListCross"
    publicProcs.Add "GetModelListCross"
    publicProcs.Add "GetInstallationTypeFromMountType"
    publicProcs.Add "CheckInstallationExists"
    publicProcs.Add "DiagnoseNamedRanges"
    
    ' === предпросмотр изменений видимости ===
    ' === visibility change preview ===
    Debug.Print Ru("003D 003D 003D 0020 043F 0440 0435 0434 043F 0440 043E 0441 043C 043E 0442 0440 0020 0438 0437 043C 0435 043D 0435 043D 0438 0439 0020 0432 0438 0434 0438 043C 043E 0441 0442") & Ru("0438 0020 003D 003D 003D")
    
    For lineNum = 1 To codeMod.CountOfLines
        lineText = Trim(codeMod.Lines(lineNum, 1))
        
        If lineText Like "Private Sub *" Or lineText Like "Public Sub *" Or _
           lineText Like "Private Function *" Or lineText Like "Public Function *" Then
            
            If InStr(1, lineText, "Function", vbTextCompare) > 0 Then
                procType = "Function"
            Else
                procType = "Sub"
            End If
            
            isPublic = (InStr(1, lineText, "Public", vbTextCompare) > 0)
            procName = ExtractProcedureName(lineText, procType)
            
            If procName <> "" Then
                needsPublic = IsProcInCollection(publicProcs, procName)
                
                If (isPublic And Not needsPublic) Or (Not isPublic And needsPublic) Then
                    changesCount = changesCount + 1
                    ' строка 
                    ' line 
                    Debug.Print Ru("0441 0442 0440 043E 043A 0430 0020") & lineNum & ": " & lineText
                    If needsPublic Then
                        '   -> будет public (сейчас private)
                        '   -> will be public (now private)
                        Debug.Print Ru("0020 0020 002D 003E 0020 0431 0443 0434 0435 0442 0020 0050 0055 0042 004C 0049 0043 0020 0028 0441 0435 0439 0447 0430 0441 0020 0050 0072 0069 0076 0061 0074 0065 0029")
                    Else
                        '   -> будет private (сейчас public)
                        '   -> will be private (now public)
                        Debug.Print Ru("0020 0020 002D 003E 0020 0431 0443 0434 0435 0442 0020 0050 0072 0069 0076 0061 0074 0065 0020 0028 0441 0435 0439 0447 0430 0441 0020 0050 0075 0062 006C 0069 0063 0029")
                    End If
                End If
            End If
        End If
    Next lineNum
    
    ' === итог: 
    ' === summary: 
    '  изменений ===
    '  ofмененandй ===
    Debug.Print Ru("003D 003D 003D 0020 0438 0442 043E 0433 003A 0020") & changesCount & Ru("0020 0438 0437 043C 0435 043D 0435 043D 0438 0439 0020 003D 003D 003D")
    
    If changesCount = 0 Then
        ' все процедуры уже имеют правильную видимость!
        ' all procedures already have correct visibility!
        MsgBox Ru("0432 0441 0435 0020 043F 0440 043E 0446 0435 0434 0443 0440 044B 0020 0443 0436 0435 0020 0438 043C 0435 044E 0442 0020 043F 0440 0430 0432 0438 043B 044C 043D 0443 044E 0020") & Ru("0432 0438 0434 0438 043C 043E 0441 0442 044C 0021"), vbInformation
    Else
        ' найдено 
        ' found 
        '  процедур с неправильной видимостью.
        '  procedures with wrong visibility.
        ' запустите restoreprocedurevisibility для исправления.
        ' run restoreprocedurevisibility to fix.
        MsgBox Ru("043D 0430 0439 0434 0435 043D 043E 0020") & changesCount & Ru("0020 043F 0440 043E 0446 0435 0434 0443 0440 0020 0441 0020 043D 0435 043F 0440 0430 0432 0438 043B 044C 043D 043E 0439 0020 0432 0438 0434 0438 043C 043E 0441 0442 044C 044E") & Ru("002E") & vbCrLf & _
               Ru("0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0052 0065 0073 0074 006F 0072 0065 0050 0072 006F 0063 0065 0064 0075 0072 0065 0056 0069 0073 0069 0062 0069 006C 0069 0074") & Ru("0079 0020 0434 043B 044F 0020 0438 0441 043F 0440 0430 0432 043B 0435 043D 0438 044F 002E"), vbExclamation
    End If
End Sub

' ================================================================
' restore procedure visibility
' ================================================================
Sub RestoreProcedureVisibility()
    On Error GoTo CleanExit
    
    Dim wsAnod As Worksheet
    Dim vbComp As VBComponent
    Dim codeMod As CodeModule
    Dim lineNum As Long
    Dim lineText As String
    Dim procName As String
    Dim procType As String
    Dim isPublic As Boolean
    Dim needsPublic As Boolean
    Dim changesCount As Long
    Dim i As Long
    
    Set wsAnod = thisWorkbook.Worksheets("Anod")
    If wsAnod Is Nothing Then
        ' лист 'anod' не найден!
        ' лandст 'anod' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027 0041 006E 006F 0064 0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        Exit Sub
    End If
    
    ' get a reference to the Anod sheet module
    Set vbComp = thisWorkbook.VBProject.VBComponents(wsAnod.CodeName)
    Set codeMod = vbComp.CodeModule
    
    changesCount = 0
    
    ' --- 1. procedures that must be Public ---
    Dim publicProcs As New Collection
    publicProcs.Add "GetMaterialListCross"
    publicProcs.Add "GetMountTypeListCross"
    publicProcs.Add "GetInstallTypeListCross"
    publicProcs.Add "GetDeliveryListCross"
    publicProcs.Add "GetModelListCross"
    publicProcs.Add "GetInstallationTypeFromMountType"
    publicProcs.Add "CheckInstallationExists"
    publicProcs.Add "DiagnoseNamedRanges"    ' if present in the module
    
    ' --- 2. walk all procedures ---
    For lineNum = 1 To codeMod.CountOfLines
        lineText = Trim(codeMod.Lines(lineNum, 1))
        
        ' check whether the line is a procedure declaration
        If lineText Like "Private Sub *" Or lineText Like "Public Sub *" Or _
           lineText Like "Private Function *" Or lineText Like "Public Function *" Then
            
            ' determine the procedure type
            If InStr(1, lineText, "Function", vbTextCompare) > 0 Then
                procType = "Function"
            Else
                procType = "Sub"
            End If
            
            ' determine current visibility
            isPublic = (InStr(1, lineText, "Public", vbTextCompare) > 0)
            
            ' extract the procedure name
            procName = ExtractProcedureName(lineText, procType)
            
            If procName <> "" Then
                ' check whether the procedure must be Public
                needsPublic = IsProcInCollection(publicProcs, procName)
                
                ' if current visibility does not match the required one
                If (isPublic And Not needsPublic) Or (Not isPublic And needsPublic) Then
                    ' replace the line
                    Dim newLine As String
                    If needsPublic Then
                        newLine = Replace(lineText, "Private ", "Public ")
                        newLine = Replace(newLine, "Private", "Public")
                    Else
                        newLine = Replace(lineText, "Public ", "Private ")
                        newLine = Replace(newLine, "Public", "Private")
                    End If
                    
                    ' apply the change
                    codeMod.ReplaceLine lineNum, newLine
                    changesCount = changesCount + 1
                    ' изменено: 
                    ' changed: 
                    Debug.Print Ru("0438 0437 043C 0435 043D 0435 043D 043E 003A 0020") & lineText & " -> " & newLine
                End If
            End If
        End If
    Next lineNum
    
    ' --- 3. info message ---
    If changesCount > 0 Then
        ' восстановлена видимость 
        ' visibility restored 
        '  процедур.
        '  procedures.
        ' public оставлены только для функций фильтрации.
        ' public kept only for filter functions.
        ' все остальные процедуры сделаны private.
        ' all other procedures set to private.
        ' перезапустите excel, чтобы изменения вступили в силу.
        ' restart excel, for changes to take effect.
        MsgBox Ru("0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0430 0020 0432 0438 0434 0438 043C 043E 0441 0442 044C 0020") & changesCount & Ru("0020 043F 0440 043E 0446 0435 0434 0443 0440 002E") & vbCrLf & _
               Ru("0050 0075 0062 006C 0069 0063 0020 043E 0441 0442 0430 0432 043B 0435 043D 044B 0020 0442 043E 043B 044C 043A 043E 0020 0434 043B 044F 0020 0444 0443 043D 043A 0446 0438 0439") & Ru("0020 0444 0438 043B 044C 0442 0440 0430 0446 0438 0438 002E") & vbCrLf & _
               Ru("0432 0441 0435 0020 043E 0441 0442 0430 043B 044C 043D 044B 0435 0020 043F 0440 043E 0446 0435 0434 0443 0440 044B 0020 0441 0434 0435 043B 0430 043D 044B 0020 0050 0072 0069") & Ru("0076 0061 0074 0065 002E") & vbCrLf & vbCrLf & _
               Ru("043F 0435 0440 0435 0437 0430 043F 0443 0441 0442 0438 0442 0435 0020 0045 0078 0063 0065 006C 002C 0020 0447 0442 043E 0431 044B 0020 0438 0437 043C 0435 043D 0435 043D 0438") & Ru("044F 0020 0432 0441 0442 0443 043F 0438 043B 0438 0020 0432 0020 0441 0438 043B 0443 002E"), vbInformation
    Else
        ' все процедуры уже имеют правильную видимость!
        ' all procedures already have correct visibility!
        ' изменений не требуется.
        ' no changes needed.
        MsgBox Ru("0432 0441 0435 0020 043F 0440 043E 0446 0435 0434 0443 0440 044B 0020 0443 0436 0435 0020 0438 043C 0435 044E 0442 0020 043F 0440 0430 0432 0438 043B 044C 043D 0443 044E 0020") & Ru("0432 0438 0434 0438 043C 043E 0441 0442 044C 0021") & vbCrLf & _
               Ru("0438 0437 043C 0435 043D 0435 043D 0438 0439 0020 043D 0435 0020 0442 0440 0435 0431 0443 0435 0442 0441 044F 002E"), vbInformation
    End If
    
    Exit Sub

CleanExit:
    If Err.Number <> 0 Then
        ' ошибка: 
        ' error: 
        ' код: 
        ' code: 
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description & vbCrLf & Ru("043A 043E 0434 003A 0020") & Err.Number, vbExclamation
    End If
End Sub

' ================================================================
' helper: extract the procedure name
' ================================================================
Private Function ExtractProcedureName(ByVal lineText As String, ByVal procType As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim result As String
    
    ' find the name position (after Sub or Function)
    If procType = "Sub" Then
        startPos = InStr(1, lineText, "Sub", vbTextCompare) + 3
    Else
        startPos = InStr(1, lineText, "Function", vbTextCompare) + 8
    End If
    
    ' skip spaces
    Do While startPos <= Len(lineText) And Mid(lineText, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    ' find the end of the name (before a space or parenthesis)
    endPos = startPos
    Do While endPos <= Len(lineText)
        Dim ch As String
        ch = Mid(lineText, endPos, 1)
        If ch = " " Or ch = "(" Or ch = vbTab Then
            Exit Do
        End If
        endPos = endPos + 1
    Loop
    
    If startPos < endPos Then
        result = Mid(lineText, startPos, endPos - startPos)
        result = Trim(result)
    Else
        result = ""
    End If
    
    ExtractProcedureName = result
End Function

' ================================================================
' helper: lookup in a collection
' ================================================================
Private Function IsProcInCollection(ByVal col As Collection, ByVal procName As String) As Boolean
    On Error Resume Next
    Dim item As Variant
    For Each item In col
        If StrComp(item, procName, vbTextCompare) = 0 Then
            IsProcInCollection = True
            Exit Function
        End If
    Next item
    IsProcInCollection = False
End Function

' ================================================================
' diagnostics after a bulk paste
' ================================================================
Sub DiagnoseAfterPaste()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = thisWorkbook.Worksheets("Anod")
    
    ' === диагностика после групповой вставки ===
    ' === diagnostics after bulk paste ===
    Debug.Print Ru("003D 003D 003D 0020 0434 0438 0430 0433 043D 043E 0441 0442 0438 043A 0430 0020 043F 043E 0441 043B 0435 0020 0433 0440 0443 043F 043F 043E 0432 043E 0439 0020 0432 0441 0442") & Ru("0430 0432 043A 0438 0020 003D 003D 003D")
    
    ' 1. check that names exist
    Dim nameExists As Boolean
    nameExists = False
    On Error Resume Next
    Dim testRng As Range
    Set testRng = thisWorkbook.names("typeInstallationAG").RefersToRange
    If Err.Number = 0 Then
        nameExists = True
    End If
    On Error GoTo 0
    
    If Not nameExists Then
        ' ! typeinstallationag не найден
        ' ! typeinstallationag not found
        Debug.Print Ru("0021 0020 0074 0079 0070 0065 0049 006E 0073 0074 0061 006C 006C 0061 0074 0069 006F 006E 0041 0047 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D")
    Else
        ' 2. check typeInstallationAG
        Dim rng As Range
        Set rng = thisWorkbook.names("typeInstallationAG").RefersToRange
        Debug.Print "typeInstallationAG: " & rng.address
        
        Dim vals As Variant
        vals = rng.Value
        
        ' safe read
        If IsArray(vals) Then
            On Error Resume Next
            Dim cols As Long
            cols = UBound(vals, 2)
            If Err.Number = 0 Then
                '   колонок: 
                '   columns: 
                Debug.Print Ru("0020 0020 043A 043E 043B 043E 043D 043E 043A 003A 0020") & cols
                Dim i As Long
                For i = 1 To cols
                    Debug.Print "    [" & i & "] " & vals(1, i)
                Next i
            Else
                Err.Clear
                '   массив 1d: 
                '   array 1d: 
                Debug.Print Ru("0020 0020 043C 0430 0441 0441 0438 0432 0020 0031 0044 003A 0020") & UBound(vals)
                For i = 1 To UBound(vals)
                    Debug.Print "    [" & i & "] " & vals(i)
                Next i
            End If
            On Error GoTo 0
        Else
            '   значение: 
            '   value: 
            Debug.Print Ru("0020 0020 0437 043D 0430 0447 0435 043D 0438 0435 003A 0020") & vals
        End If
    End If
    
    Debug.Print ""
    
    ' 3. check cell values (row 36)
    Dim pipeCount As Long
    pipeCount = SafeNumericValue(ws.Range("pipeCountCP").Value, 1)
    Debug.Print "pipeCountCP: " & pipeCount
    
    Dim colIdx As Long
    For colIdx = 4 To 4 + pipeCount - 1
        ' колонка 
        ' column 
        ' : строка36=
        ' : line36=
        Debug.Print Ru("043A 043E 043B 043E 043D 043A 0430 0020") & colIdx & Ru("003A 0020 0441 0442 0440 043E 043A 0430 0033 0036 003D") & ws.Cells(36, colIdx).Value
    Next colIdx
    
    Debug.Print ""
    
    ' 4. check filter named ranges
    Dim filterNames As Variant
    filterNames = Array("typeMaterial", "typeMountingAG", "typeInstallationAG", "typeDeliveryAG", "typeAG")
    
    Dim nm As Variant
    For Each nm In filterNames
        On Error Resume Next
        Set rng = thisWorkbook.names(nm).RefersToRange
        If Err.Number = 0 Then
            Debug.Print nm & ": " & rng.address & " - OK"
        Else
            ' : не найден!
            ' : not found!
            Debug.Print nm & Ru("003A 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021")
            Err.Clear
        End If
        On Error GoTo 0
    Next nm
    
    Debug.Print ""
    
    ' 5. run btnAnodFullCalc
    ' === запуск btnanodfullcalc ===
    ' === start btnanodfullcalc ===
    Debug.Print Ru("003D 003D 003D 0020 0437 0430 043F 0443 0441 043A 0020 0062 0074 006E 0041 006E 006F 0064 0046 0075 006C 006C 0043 0061 006C 0063 0020 003D 003D 003D")
    On Error Resume Next
    Module_calcAnod.btnAnodFullCalc
    If Err.Number <> 0 Then
        ' ошибка btnanodfullcalc: 
        ' error btnanodfullcalc: 
        Debug.Print Ru("043E 0448 0438 0431 043A 0430 0020 0062 0074 006E 0041 006E 006F 0064 0046 0075 006C 006C 0043 0061 006C 0063 003A 0020") & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' === диагностика завершена ===
    ' === diagnostics finished ===
    Debug.Print Ru("003D 003D 003D 0020 0434 0438 0430 0433 043D 043E 0441 0442 0438 043A 0430 0020 0437 0430 0432 0435 0440 0448 0435 043D 0430 0020 003D 003D 003D")
End Sub
