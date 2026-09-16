Attribute VB_Name = "Module_AGData"

Option Explicit

' ================================================================
' Module_AGData
' Insertion of AG data and clearing of calculated rows
' ================================================================

' Inserts data from named ranges into rows 39-49 for a given column
Public Sub InsertAGDataForColumn(ByVal ws As Worksheet, ByVal col As Long)
    On Error GoTo CleanExit
    
    Dim eventsState As Boolean
    eventsState = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim modelValue As Variant
    modelValue = ws.Cells(MODEL_ROW, col).Value
    
    Dim clearRow As Long
    For clearRow = DATA_START_ROW To DATA_START_ROW + 10
        ws.Cells(clearRow, col).ClearContents
    Next clearRow
    
    If IsEmpty(modelValue) Or modelValue = "" Then
        GoTo CleanExit
    End If
    
    Dim rngModel As Range
    On Error Resume Next
    Set rngModel = thisWorkbook.names("AG_Model").RefersToRange
    On Error GoTo 0
    
    If rngModel Is Nothing Then GoTo CleanExit
    
    Dim modelVals As Variant
    modelVals = rngModel.Value
    If Not IsArray(modelVals) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = modelVals
        modelVals = arr
    End If
    
    Dim modelValueClean As String
    modelValueClean = CleanStringForTable(CStr(modelValue))
    
    Dim foundRow As Long
    foundRow = 0
    Dim i As Long
    For i = 1 To UBound(modelVals, 1)
        Dim modelValClean As String
        modelValClean = CleanStringForTable(CStr(modelVals(i, 1)))
        If modelValClean = modelValueClean Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then GoTo CleanExit
    
    Dim paramNames As Variant
    Dim targetRows As Variant
    
    paramNames = Array( _
        "AG_Diameter", _
        "AG_Length", _
        "AG_CokeDiam", _
        "AG_CokeLength", _
        "AG_Mass", _
        "AG_DissolutionRate", _
        "AG_RatedCurrent", _
        "AG_ResistivityMaterial", _
        "AG_CokeResistivity", _
        "AG_SpecificRatedCurrent", _
        "AG_SpecificMass" _
    )
    
    targetRows = Array(39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49)
    
    Dim idx As Long
    Dim rngParam As Range
    Dim paramVals As Variant
    Dim targetRow As Long
    Dim valToWrite As Variant
    
    For idx = LBound(paramNames) To UBound(paramNames)
        targetRow = targetRows(idx)
        
        On Error Resume Next
        Set rngParam = thisWorkbook.names(paramNames(idx)).RefersToRange
        If Err.Number = 0 Then
            paramVals = rngParam.Value
            
            If Not IsArray(paramVals) Then
                ReDim arr(1 To 1, 1 To 1)
                arr(1, 1) = paramVals
                paramVals = arr
            End If
            
            If foundRow <= UBound(paramVals, 1) Then
                valToWrite = paramVals(foundRow, 1)
                If Not IsEmpty(valToWrite) Then ws.Cells(targetRow, col).Value = valToWrite
            End If
        Else
            Err.Clear
        End If
        On Error GoTo CleanExit
    Next idx

CleanExit:
    Application.EnableEvents = eventsState
End Sub

' Clears calculated data rows for a given column (rows 39-49, 59, 60-119)
Public Sub ClearCalculatedDataForCol(ByVal ws As Worksheet, ByVal targetCol As Long)
    Dim rw As Long
    ' Clear rows 39-49 (AG parameters)
    For rw = DATA_START_ROW To DATA_START_ROW + 10
        ws.Cells(rw, targetCol).ClearContents
    Next rw
    ' Clear row 59 and range 60-119 (calculated data)
    ws.Cells(59, targetCol).ClearContents
    ws.Range(ws.Cells(60, targetCol), ws.Cells(119, targetCol)).ClearContents
    ' Rows 50-57 and 58 are not cleared
End Sub

' Clears calculated data column (synonym for ClearCalculatedDataForCol)
Public Sub ClearCalculatedDataForColumn(ByVal ws As Worksheet, ByVal col As Long)
    Call ClearCalculatedDataForCol(ws, col)
End Sub
