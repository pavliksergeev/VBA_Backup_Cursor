Attribute VB_Name = "Module_PestorePipeNames"

' ================================================================
' module: Module_RestorePipeNames
' purpose: restore named ranges on the wsPipe sheet
' ================================================================
Option Explicit

' ================================================================
' main restore procedure
' ================================================================
Public Sub RestorePipeNamedRanges()
    On Error GoTo CleanExit
    
    ' save settings
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    
    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = "restoring Pipe calculation named ranges..."
    
    Dim wsPipe As Worksheet
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    
    If wsPipe Is Nothing Then
        ' лист '
        ' лandст '
        ' ' не найден!
        ' ' not found!
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbCritical
        GoTo CleanExit
    End If
    
    ' get the column count from D3 (pipeDifferentParametersNum)
    Dim colBr As Long
    Dim pipeCountValue As Variant
    
    On Error Resume Next
    pipeCountValue = wsPipe.Range("D3").Value
    If Err.Number <> 0 Then
        pipeCountValue = 1
        Err.Clear
    End If
    On Error GoTo 0
    
    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
        colBr = CLng(pipeCountValue)
    Else
        colBr = 1
    End If
    
    If colBr < 1 Then colBr = 1
    If colBr > 1024 Then colBr = 1024
    
    Application.StatusBar = "restoring: " & colBr & " columns..."
    DoEvents
    
    ' delete old names on the Pipe calculation sheet
    Call DeleteOldPipeNames(wsPipe)
    
    ' create new names
    Call CreatePipeNames(wsPipe, colBr)
    
    Application.StatusBar = "restore finished: " & colBr & " columns"
    
    'MsgBox "Pipe calculation named ranges restored" & vbCrLf & _
    '       "columns: " & colBr, vbInformation, "restore finished"
    
CleanExit:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc
    Application.StatusBar = False
End Sub

' ================================================================
' delete old names on the PIPE CALCULATION sheet
' ================================================================
Private Sub DeleteOldPipeNames(ByVal wsPipe As Worksheet)
    On Error Resume Next
    
    Application.StatusBar = "removing old names..."
    DoEvents
    
    Dim nm As name
    Dim deletedCount As Long
    deletedCount = 0
    
    ' names to restore (used for deletion)
    Dim namesToRestore As Variant
    namesToRestore = Array( _
        "pipeSteelGrade", _
        "pipeSteelResistivity", _
        "pipeDiameter", _
        "pipeWallThickness", _
        "pipeInsulationResistivityStartLife", _
        "pipeLayingDepth", _
        "soilResistivityAvg", _
        "serviceLifeDesigned", _
        "pipeResistivityChangeFactor", _
        "pipeAlongResistance", _
        "soilResistanceAroundPipe", _
        "pipeTransientResistivity", _
        "soilResistivityAroundPipe", _
        "pipeInsulationResistanceStartLife", _
        "pipeTransientResistance", _
        "pipeTransientResistanceEndLife", _
        "factorPropagationCurrentAlongPipe", _
        "factorPropagationCurrentAlongPipeEndLife", _
        "pipeImpedance", _
        "pipeImpedanceEndLife", _
        "pipeInputResistance", _
        "pipeInputResistanceEndLife", _
        "pipeDifferentParametersNum" _
    )
    
    Dim nameToDelete As Variant
    For Each nameToDelete In namesToRestore
        ' try to delete the global name
        On Error Resume Next
        thisWorkbook.names(nameToDelete).Delete
        If Err.Number = 0 Then
            deletedCount = deletedCount + 1
        Else
            Err.Clear
            ' try to delete the local name
            wsPipe.names(nameToDelete).Delete
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
            Else
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next nameToDelete
    
    ' удалено имен: 
    ' names deleted: 
    Debug.Print Ru("0443 0434 0430 043B 0435 043D 043E 0020 0438 043C 0435 043D 003A 0020") & deletedCount
End Sub

' ================================================================
' create new names
' ================================================================
Private Sub CreatePipeNames(ByVal wsPipe As Worksheet, ByVal colBr As Long)
    On Error GoTo CleanExit
    
    Application.StatusBar = "creating new names..."
    DoEvents
    
    Dim startCol As Long
    startCol = 4
    
    Dim endCol As Long
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, _
                     Application.WorksheetFunction.Min(1024, endCol))
    
    ' ================================================================
    ' create scalar names (single cell)
    ' ================================================================
    Call CreateScalarName(wsPipe, "pipeDifferentParametersNum", "D3")
    Call CreateScalarName(wsPipe, "pipeInputResistance", "D28")
    Call CreateScalarName(wsPipe, "pipeInputResistanceEndLife", "D29")
    
    ' ================================================================
    ' create range names
    ' ================================================================
    Dim rangeNames As Variant
    Dim rangeRows As Variant
    
    ' name array
    rangeNames = Array( _
        "pipeSteelGrade", _
        "pipeSteelResistivity", _
        "pipeDiameter", _
        "pipeWallThickness", _
        "pipeInsulationResistivityStartLife", _
        "pipeLayingDepth", _
        "soilResistivityAvg", _
        "serviceLifeDesigned", _
        "pipeResistivityChangeFactor", _
        "pipeAlongResistance", _
        "soilResistanceAroundPipe", _
        "pipeTransientResistivity", _
        "soilResistivityAroundPipe", _
        "pipeInsulationResistanceStartLife", _
        "pipeTransientResistance", _
        "pipeTransientResistanceEndLife", _
        "factorPropagationCurrentAlongPipe", _
        "factorPropagationCurrentAlongPipeEndLife", _
        "pipeImpedance", _
        "pipeImpedanceEndLife" _
    )
    
    ' row for each name (indexes match rangeNames)
    ' pipeSteelGrade = 5
    ' pipeSteelResistivity = 6
    ' pipeDiameter = 7
    ' pipeWallThickness = 8
    ' pipeInsulationResistivityStartLife = 9
    ' pipeLayingDepth = 10
    ' soilResistivityAvg = 11
    ' serviceLifeDesigned = 12
    ' pipeResistivityChangeFactor = 13
    ' pipeAlongResistance = 17
    ' soilResistanceAroundPipe = 18
    ' pipeTransientResistivity = 19
    ' soilResistivityAroundPipe = 20
    ' pipeInsulationResistanceStartLife = 21
    ' pipeTransientResistance = 22
    ' pipeTransientResistanceEndLife = 23
    ' factorPropagationCurrentAlongPipe = 24
    ' factorPropagationCurrentAlongPipeEndLife = 25
    ' pipeImpedance = 26
    ' pipeImpedanceEndLife = 27
    rangeRows = Array(5, 6, 7, 8, 9, 10, 11, 12, 13, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27)
    
    Dim nameIndex As Long
    Dim startRow As Long
    Dim newRng As Range
    Dim createdCount As Long
    createdCount = 0
    
    For nameIndex = LBound(rangeNames) To UBound(rangeNames)
        startRow = rangeRows(nameIndex)
        
        ' create a range from column 4 to endCol
        Set newRng = wsPipe.Range(wsPipe.Cells(startRow, startCol), _
                                  wsPipe.Cells(startRow, endCol))
        
        If Not newRng Is Nothing Then
            ' create the name (sheet-local scope only)
            On Error Resume Next
            wsPipe.names.Add name:=rangeNames(nameIndex), RefersTo:=newRng
            If Err.Number = 0 Then
                createdCount = createdCount + 1
                ' создано имя: 
                ' name created: 
                '  (строка 
                '  (line 
                ' , колонки 
                ' , колонкand 
                Debug.Print Ru("0441 043E 0437 0434 0430 043D 043E 0020 0438 043C 044F 003A 0020") & rangeNames(nameIndex) & Ru("0020 0028 0441 0442 0440 043E 043A 0430 0020") & startRow & Ru("002C 0020 043A 043E 043B 043E 043D 043A 0438 0020") & startCol & "-" & endCol & ")"
            Else
                ' не удалось создать имя 
                ' failed to create name 
                Debug.Print Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 0441 043E 0437 0434 0430 0442 044C 0020 0438 043C 044F 0020") & rangeNames(nameIndex)
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        ' update the status every 5 names
        If nameIndex Mod 5 = 0 Then
            Application.StatusBar = "creating names: " & Format((nameIndex + 1) / (UBound(rangeNames) + 1) * 100, "0") & "%"
            DoEvents
        End If
    Next nameIndex
    
    ' создано имен: 
    ' names created: 
    Debug.Print Ru("0441 043E 0437 0434 0430 043D 043E 0020 0438 043C 0435 043D 003A 0020") & createdCount
    
CleanExit:
End Sub

' ================================================================
' create a scalar name
' ================================================================
Private Sub CreateScalarName(ByVal wsPipe As Worksheet, ByVal name As String, ByVal address As String)
    On Error Resume Next
    
    ' delete the old name if it exists
    wsPipe.names(name).Delete
    Err.Clear
    
    ' create the new scalar name
    wsPipe.names.Add name:=name, RefersTo:=wsPipe.Range(address)
    
    If Err.Number = 0 Then
        ' создано скалярное имя: 
        ' scalar name created: 
        Debug.Print Ru("0441 043E 0437 0434 0430 043D 043E 0020 0441 043A 0430 043B 044F 0440 043D 043E 0435 0020 0438 043C 044F 003A 0020") & name & " -> " & address
    Else
        ' ошибка: не удалось создать скалярное имя 
        ' error: failed to create scalar name 
        Debug.Print Ru("043E 0448 0438 0431 043A 0430 003A 0020 043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 0441 043E 0437 0434 0430 0442 044C 0020 0441 043A 0430 043B 044F 0440 043D 043E") & Ru("0435 0020 0438 043C 044F 0020") & name
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' ================================================================
' button that runs restore
' ================================================================
Sub btnRestorePipeNames()
    Call RestorePipeNamedRanges
End Sub
