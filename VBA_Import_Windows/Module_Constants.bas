Attribute VB_Name = "Module_Constants"

' ================================================================
' shared constants module
' ================================================================
Option Explicit
' ================================================================
' 0. sheet protection password
' ================================================================
Public Const SHEET_PASSWORD As String = "ptm"
' ================================================================
' 1. column constants
' ================================================================
Public Const START_COL As Long = 4
Public Const MAX_COL As Long = 254
Public Const MIN_COL As Long = 1

' ================================================================
' 2. row constants
' ================================================================
Public Const HEADER_ROW As Long = 2
Public Const MODEL_ROW As Long = 38
Public Const DATA_START_ROW As Long = 39
Public Const CP_MODEL_ROW As Long = 26
Public Const CP_HYPERLINK_ROW As Long = 27
Public Const CP_CURRENT_ROW As Long = 29
Public Const CP_VOLTAGE_ROW As Long = 30
Public Const CP_POWER_ROW As Long = 31

' ================================================================
' 3. filter constants
' ================================================================
Public Const FILTER_START_ROW As Long = 34      ' anode material
Public Const FILTER_END_ROW As Long = 38

' ================================================================
' 4. debug mode
' ================================================================
Public Const DEBUG_MODE As Boolean = True ' False/True

' ================================================================
' 5. sheet names (after rename)
' ================================================================
Public Const SHEET_ANOD As String = "Anod"
Public Const SHEET_PIPE As String = "Pipe"
Public Const SHEET_RESEARCH As String = "Research"
Public Const SHEET_LIST_AG As String = "ListAG"
Public Const SHEET_LIST_CP As String = "ListCP"
Public Const SHEET_STO As String = "STO"
Public Const SHEET_SETTINGS As String = "_Settings"
Public Const SHEET_GUIDE As String = "Guide"
Public Const SHEET_STO_REF As String = "STO"

' ================================================================
' 6. table names
' ================================================================
Public Const TABLE_AG As String = "TableAG"
Public Const TABLE_CP As String = "TableCP"
Public Const TABLE_ALL_TEST As String = "TableAllTest"

' ================================================================
' 7. scalar names (not variables) - notes:
' ================================================================
' note: these names must refer to a single cell!
' ================================================================
Public Const SCALAR_NAMES As String = _
    "Print_Area|" & _
    "pipeCountCP|" & _
    "pipeLength|" & _
    "lengthProtectiveZone|" & _
    "pipeDifferentParametersNum|" & _
    "minProtectPotential|" & _
    "maxProtectPotential|" & _
    "naturalPotential|" & _
    "factorMutualInfluence|" & _
    "pipeShiftPotentialMin|" & _
    "pipeShiftPotentialPoint|" & _
    "pipeInputResistance|" & _
    "pipeInputResistanceEndLife"

' ================================================================
' 7.1. scalar-name array for checks
' ================================================================
' note: GetScalarNamesArray() is in section 12
' (after all constants, because VBA cannot place a Function before all Const)

' ================================================================
' 7.2. check whether a name is scalar
' ================================================================
' note: IsScalarName() is in section 13
' (after all constants, because VBA cannot place a Function before all Const)

' ================================================================
' 8. protective-potential limits
' ================================================================
Public Const MIN_PROTECT_POTENTIAL As Double = -3.5
Public Const MAX_PROTECT_POTENTIAL As Double = -1.1

' ================================================================
' 9. Lz limits (protective-zone length)
' ================================================================
Public Const MIN_Lz As Double = 500      ' minimum protective-zone length (m)
Public Const MAX_Lz As Double = 100000   ' maximum protective-zone length (m)

' ================================================================
' 10. column-count warning
' ================================================================
Public Const MAX_COLUMNS_WARNING As Long = 250

' ================================================================
' 11. progress-bar constants
' ================================================================
Public Const PROGRESS_BAR_LENGTH As Long = 40
Public Const PROGRESS_UPDATE_INTERVAL As Long = 5

' ================================================================
' ================================================================
' all procedures (Sub and Function) must be below all constants!
' ================================================================
' ================================================================

' ================================================================
' 12. scalar-name array for checks (from section 7.1)
' ================================================================
Public Function GetScalarNamesArray() As Variant
    Dim namesArray() As String
    namesArray = Split(SCALAR_NAMES, "|")
    GetScalarNamesArray = namesArray
End Function

' ================================================================
' 13. check whether a name is scalar (from section 7.2)
' ================================================================
Public Function IsScalarName(ByVal name As String) As Boolean
    ' strip the sheet prefix if present
    Dim pureName As String
    pureName = name
    Dim pos As Long
    pos = InStr(pureName, "!")
    If pos > 0 Then
        pureName = Mid(pureName, pos + 1)
    End If
    
    Dim scalarArray() As String
    scalarArray = Split(SCALAR_NAMES, "|")
    
    Dim i As Long
    For i = LBound(scalarArray) To UBound(scalarArray)
        If StrComp(pureName, scalarArray(i), vbTextCompare) = 0 Then
            IsScalarName = True
            Exit Function
        End If
    Next i
    
    IsScalarName = False
End Function

' ================================================================
' 14. debug: extra check of all scalar names
' ================================================================
Public Sub CheckAllScalarNames()
    Dim scalarArray() As String
    scalarArray = Split(SCALAR_NAMES, "|")
    
    Dim i As Long
    Dim name As String
    Dim rng As Range
    Dim val As Variant
    Dim msg As String
    Dim wsAnod As Worksheet
    Dim wsPipe As Worksheet
    
    Set wsAnod = thisWorkbook.Worksheets(SHEET_ANOD)
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    
    ' проверка скалярных имен:
    ' scalar names check:
    msg = Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 0441 043A 0430 043B 044F 0440 043D 044B 0445 0020 0438 043C 0435 043D 003A") & vbCrLf & vbCrLf
    
    ' names that live on the Anod sheet
    Dim anodNames As Variant
    anodNames = Array( _
        "pipeCountCP", "pipeLength", "lengthProtectiveZone", _
        "minProtectPotential", "maxProtectPotential", "naturalPotential", _
        "factorMutualInfluence", "pipeShiftPotentialMin", "pipeShiftPotentialPoint" _
    )
    
    ' names that live on the Pipe sheet
    Dim pipeNames As Variant
    pipeNames = Array( _
        "pipeDifferentParametersNum", "pipeInputResistance", "pipeInputResistanceEndLife" _
    )
    
    For i = LBound(scalarArray) To UBound(scalarArray)
        name = scalarArray(i)
        If name = "" Or name = "Print_Area" Then GoTo NextName
        
        ' decide which sheet to search for the name
        Dim found As Boolean
        found = False
        Dim ws As Worksheet
        Dim nm As name
        
        ' check whether the name belongs to Anod
        Dim j As Long
        For j = LBound(anodNames) To UBound(anodNames)
            If StrComp(name, anodNames(j), vbTextCompare) = 0 Then
                Set ws = wsAnod
                found = True
                Exit For
            End If
        Next j
        
        ' if not found among Anod names, check Pipe
        If Not found Then
            For j = LBound(pipeNames) To UBound(pipeNames)
                If StrComp(name, pipeNames(j), vbTextCompare) = 0 Then
                    Set ws = wsPipe
                    found = True
                    Exit For
                End If
            Next j
        End If
        
        If Not found Or ws Is Nothing Then
            '  - не найден (неизвестный лист)!
            '  - not found (неofвестный лandст)!
            msg = msg & "? " & name & Ru("0020 002D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0020 0028 043D 0435 0438 0437 0432 0435 0441 0442 043D 044B 0439 0020 043B 0438 0441 0442 0029 0021") & vbCrLf
            GoTo NextName
        End If
        
        ' look up the name among the sheet's local names
        On Error Resume Next
        Set rng = Nothing
        For Each nm In ws.names
            Dim pureName As String
            pureName = nm.name
            Dim pos As Long
            pos = InStr(pureName, "!")
            If pos > 0 Then
                pureName = Mid(pureName, pos + 1)
            Else
                pureName = pureName
            End If
            
            If StrComp(pureName, name, vbTextCompare) = 0 Then
                Set rng = nm.RefersToRange
                Exit For
            End If
        Next nm
        
        If Err.Number <> 0 Or rng Is Nothing Then
            '  - не найден на листе 
            '  - not found on лandсте 
            msg = msg & "? " & name & Ru("0020 002D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0020 043D 0430 0020 043B 0438 0441 0442 0435 0020") & ws.name & "!" & vbCrLf
            Err.Clear
            GoTo NextName
        End If
        On Error GoTo 0
        
        val = rng.Value
        
        If IsArray(val) Then
            Dim rows As Long, cols As Long
            On Error Resume Next
            rows = UBound(val, 1)
            cols = UBound(val, 2)
            On Error GoTo 0
            If rows = 1 And cols = 1 Then
                '  (массив 1x1 на 
                '  (array 1x1 on 
                msg = msg & "? " & name & " = " & val(1, 1) & Ru("0020 0028 043C 0430 0441 0441 0438 0432 0020 0031 0078 0031 0020 043D 0430 0020") & ws.name & ")" & vbCrLf
            Else
                '  - массив 
                '  - array 
                '  на 
                '  on 
                msg = msg & "?? " & name & Ru("0020 002D 0020 043C 0430 0441 0441 0438 0432 0020") & rows & "x" & cols & Ru("0020 043D 0430 0020") & ws.name & " !!!" & vbCrLf
            End If
        Else
            msg = msg & "? " & name & " = " & val & " (" & ws.name & ")" & vbCrLf
        End If
        
NextName:
    Next i
    
    ' проверка скалярных имен
    ' scalar names check
    MsgBox msg, vbInformation, Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 0441 043A 0430 043B 044F 0440 043D 044B 0445 0020 0438 043C 0435 043D")
End Sub

' ================================================================
' 15. quick check
' ================================================================
Sub QuickCheck()
    Dim vbComp As VBComponent
    Dim line As Long
    Dim codeLine As String
    Dim found As Long
    Dim hardcoded As Variant
    hardcoded = Array("Anod calculation", "Pipe calculation", "Research data", "List AG", "List CP")
    
    ' === проверка ===
    ' === check ===
    Debug.Print Ru("003D 003D 003D 0020 043F 0440 043E 0432 0435 0440 043A 0430 0020 003D 003D 003D")
    found = 0
    
    For Each vbComp In thisWorkbook.VBProject.VBComponents
        If vbComp.name <> "Module_Constants" And vbComp.name <> "Module_Logging" Then
            For line = 1 To vbComp.CodeModule.CountOfLines
                codeLine = vbComp.CodeModule.Lines(line, 1)
                Dim name As Variant
                For Each name In hardcoded
                    If InStr(1, codeLine, """" & name & """", vbTextCompare) > 0 Then
                        found = found + 1
                        '  (строка 
                        '  (line 
                        Debug.Print "?? " & vbComp.name & Ru("0020 0028 0441 0442 0440 043E 043A 0430 0020") & line & "): " & Trim(codeLine)
                    End If
                Next name
            Next line
        End If
    Next vbComp
    
    If found = 0 Then
        ' устаревших имен не найдено!
        ' obsolete names not found!
        MsgBox Ru("0443 0441 0442 0430 0440 0435 0432 0448 0438 0445 0020 0438 043C 0435 043D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 0021"), vbInformation
    Else
        ' найдено: 
        ' found: 
        '  жестких ссылок. смотрите immediate window (ctrl+g)
        '  hard references. see the immediate window (ctrl+g)
        MsgBox Ru("043D 0430 0439 0434 0435 043D 043E 003A 0020") & found & Ru("0020 0436 0435 0441 0442 043A 0438 0445 0020 0441 0441 044B 043B 043E 043A 002E 0020 0441 043C 043E 0442 0440 0438 0442 0435 0020 0049 006D 006D 0065 0064 0069 0061 0074 0065") & Ru("0020 0057 0069 006E 0064 006F 0077 0020 0028 0043 0074 0072 006C 002B 0047 0029"), vbExclamation
    End If
End Sub

' ================================================================
' 16. test all constants
' ================================================================
Sub TestAllConstants()
    On Error Resume Next
    
    Dim msg As String
    ' проверка констант:
    ' constants check:
    msg = Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 043A 043E 043D 0441 0442 0430 043D 0442 003A") & vbCrLf & vbCrLf
    
    ' verify that all sheets exist
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array(SHEET_ANOD, SHEET_PIPE, SHEET_RESEARCH, SHEET_LIST_AG, SHEET_LIST_CP, SHEET_STO, SHEET_SETTINGS, SHEET_GUIDE)
    
    Dim name As Variant
    For Each name In sheetNames
        On Error Resume Next
        Set ws = thisWorkbook.Worksheets(name)
        If Err.Number = 0 Then
            '  - найден
            '  - found
            msg = msg & "? " & name & Ru("0020 002D 0020 043D 0430 0439 0434 0435 043D") & vbCrLf
        Else
            '  - не найден!
            '  - not found!
            msg = msg & "? " & name & Ru("0020 002D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021") & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next name
    
    ' verify that all tables exist
    ' проверка таблиц:
    ' table check:
    msg = msg & vbCrLf & Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 0442 0430 0431 043B 0438 0446 003A") & vbCrLf
    
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = thisWorkbook.Worksheets(SHEET_LIST_AG).ListObjects(TABLE_AG)
    If Err.Number = 0 Then
        ' ? table_ag - найдена
        ' ? table_ag - found
        msg = msg & Ru("003F 0020 0054 0041 0042 004C 0045 005F 0041 0047 0020 002D 0020 043D 0430 0439 0434 0435 043D 0430") & vbCrLf
    Else
        ' ? table_ag - не найдена!
        ' ? table_ag - not found!
        msg = msg & Ru("003F 0020 0054 0041 0042 004C 0045 005F 0041 0047 0020 002D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 0021") & vbCrLf
        Err.Clear
    End If
    
    Set tbl = thisWorkbook.Worksheets(SHEET_LIST_CP).ListObjects(TABLE_CP)
    If Err.Number = 0 Then
        ' ? table_cp - найдена
        ' ? table_cp - found
        msg = msg & Ru("003F 0020 0054 0041 0042 004C 0045 005F 0043 0050 0020 002D 0020 043D 0430 0439 0434 0435 043D 0430") & vbCrLf
    Else
        ' ? table_cp - не найдена!
        ' ? table_cp - not found!
        msg = msg & Ru("003F 0020 0054 0041 0042 004C 0045 005F 0043 0050 0020 002D 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 0021") & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
    
    ' check the new constants
    ' проверка констант ограничений:
    ' limit constants check:
    msg = msg & vbCrLf & Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 043A 043E 043D 0441 0442 0430 043D 0442 0020 043E 0433 0440 0430 043D 0438 0447 0435 043D 0438 0439 003A") & vbCrLf
    '  м
    msg = msg & "MIN_Lz = " & MIN_Lz & Ru("0020 043C") & vbCrLf
    '  м
    msg = msg & "MAX_Lz = " & MAX_Lz & Ru("0020 043C") & vbCrLf
    '  в
    msg = msg & "MIN_PROTECT_POTENTIAL = " & MIN_PROTECT_POTENTIAL & Ru("0020 0432") & vbCrLf
    '  в
    msg = msg & "MAX_PROTECT_POTENTIAL = " & MAX_PROTECT_POTENTIAL & Ru("0020 0432") & vbCrLf
    '  колонок
    '  columns
    msg = msg & "MAX_COLUMNS_WARNING = " & MAX_COLUMNS_WARNING & Ru("0020 043A 043E 043B 043E 043D 043E 043A") & vbCrLf
    
    ' check scalar names
    ' скалярные имена (
    ' scalar names (
    '  шт.):
    '  pcs):
    msg = msg & vbCrLf & Ru("0441 043A 0430 043B 044F 0440 043D 044B 0435 0020 0438 043C 0435 043D 0430 0020 0028") & UBound(Split(SCALAR_NAMES, "|")) + 1 & Ru("0020 0448 0442 002E 0029 003A") & vbCrLf
    Dim scalarArray() As String
    scalarArray = Split(SCALAR_NAMES, "|")
    Dim s As Variant
    For Each s In scalarArray
        If s <> "" Then
            msg = msg & "  - " & s & vbCrLf
        End If
    Next s
    
    ' тест констант
    ' constants test
    MsgBox msg, vbInformation, Ru("0442 0435 0441 0442 0020 043A 043E 043D 0441 0442 0430 043D 0442")
End Sub

' ================================================================
' 17. all sheet names
' ================================================================
Sub ShowAllSheetNames()
    Dim ws As Worksheet
    Dim msg As String
    ' список листов в книге:
    ' sheet list in workbook:
    msg = Ru("0441 043F 0438 0441 043E 043A 0020 043B 0438 0441 0442 043E 0432 0020 0432 0020 043A 043D 0438 0433 0435 003A") & vbCrLf & vbCrLf
    For Each ws In thisWorkbook.Worksheets
        msg = msg & "- " & ws.name & vbCrLf
    Next ws
    ' имена листов
    ' sheet names
    MsgBox msg, vbInformation, Ru("0438 043C 0435 043D 0430 0020 043B 0438 0441 0442 043E 0432")
End Sub

' ================================================================
' 18. progress-bar functions
' ================================================================

' show the progress bar on the status bar
Public Sub ShowProgress(ByVal current As Long, ByVal total As Long, ByVal message As String)
    On Error Resume Next
    
    If total <= 0 Then Exit Sub
    
    Dim percent As Double
    percent = current / total
    
    Dim filled As Long
    filled = Int(percent * PROGRESS_BAR_LENGTH)
    If filled > PROGRESS_BAR_LENGTH Then filled = PROGRESS_BAR_LENGTH
    
    Dim bar As String
    Dim i As Long
    
    bar = "["
    For i = 1 To PROGRESS_BAR_LENGTH
        If i <= filled Then
            bar = bar & "?"
        Else
            bar = bar & "?"
        End If
    Next i
    bar = bar & "] " & Format(percent * 100, "0") & "%"
    
    Application.StatusBar = message & " " & bar
    DoEvents
    
    On Error GoTo 0
End Sub

' reset the progress bar
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub
