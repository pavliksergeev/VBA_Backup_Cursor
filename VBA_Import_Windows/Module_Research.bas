Attribute VB_Name = "Module_Research"

' ================================================================
' RESEARCH module: import from the Research Data sheet
' ================================================================
Option Explicit

' ================================================================
' Anod Form Control "Длина по данным изысканий" (pipeLength row):
' L = sum of Pipe 1xN pipeLengthSegment (Ls of each section)
' ================================================================
Public Sub buttonPipeLengthSurvey_click()
    Call btnImportLength
End Sub

Public Sub btnImportLength()
    Dim wsAnod As Worksheet
    Dim wsPipe As Worksheet
    Dim rngSeg As Range
    Dim rngLen As Range
    Dim nOk As Long
    Dim total As Double
    Dim wasProt As Boolean
    Dim oldEvents As Boolean

    wasProt = False
    oldEvents = Application.EnableEvents
    On Error GoTo CleanExit

    Set wsAnod = thisWorkbook.Worksheets(SHEET_ANOD)
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    If wsAnod Is Nothing Then
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_ANOD & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        Exit Sub
    End If
    If wsPipe Is Nothing Then
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_PIPE & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set rngSeg = wsPipe.Range("pipeLengthSegment")
    Set rngLen = wsAnod.Range("pipeLength")
    On Error GoTo CleanExit
    If rngSeg Is Nothing Then
        MsgBox Ru("0438 043C 044F 0020 0027") & "pipeLengthSegment" & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 0021"), vbExclamation
        Exit Sub
    End If
    If rngLen Is Nothing Then
        MsgBox Ru("0438 043C 044F 0020 0027") & "pipeLength" & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 0021"), vbExclamation
        Exit Sub
    End If

    nOk = 0
    total = 0
    On Error Resume Next
    nOk = CLng(Application.WorksheetFunction.Count(rngSeg))
    If nOk > 0 Then total = CDbl(Application.WorksheetFunction.Sum(rngSeg))
    On Error GoTo CleanExit
    If nOk < 1 Then
        MsgBox Ru("041D 0435 0442 0020 0447 0438 0441 043B 043E 0432 044B 0445 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0434 043B 0438 043D 044B 0020 043F 043B 0435 0447 002E"), vbExclamation
        Exit Sub
    End If

    wasProt = wsAnod.ProtectContents
    oldEvents = Application.EnableEvents
    If wasProt Then wsAnod.Unprotect Password:=SHEET_PASSWORD
    Application.EnableEvents = False
    rngLen.Value = total
    Application.EnableEvents = oldEvents
    If wasProt Then
        wsAnod.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    End If
    If DEBUG_MODE Then
        Debug.Print "pipeLength = " & CStr(total) & " (" & CStr(nOk) & " pipeLengthSegment)"
    End If
    Exit Sub

CleanExit:
    On Error Resume Next
    Application.EnableEvents = True
    If Not wsAnod Is Nothing Then
        If wasProt Then
            wsAnod.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
        End If
    End If
    On Error GoTo 0
End Sub

' ================================================================
' import the soil-heterogeneity factor
' ================================================================
Sub btnImportSoilHeterogeneity()
    Call ShowSoilAvgFormLayer
End Sub

Public Sub buttonLayerDeep_click()
    Call ShowSoilAvgFormLayer
End Sub

' ================================================================
' import resistivity of the i-th soil layer
' ================================================================
Public Sub btnImportResistivityLayer()
    Call ShowSoilAvgFormLayer
End Sub

' ================================================================
' (stub - not implemented yet)
' ================================================================
Sub button126_click()
    ' в разработке
    ' under development
    MsgBox Ru("0432 0020 0440 0430 0437 0440 0430 0431 043E 0442 043A 0435")
'    On Error GoTo CleanExit
'
'    Dim colBr As Long
'    Dim pipeCountValue As Variant
'
'    On Error Resume Next
'    pipeCountValue = Range("pipeDifferentParametersNum").Value
'    If Err.Number <> 0 Then pipeCountValue = 1
'    On Error GoTo 0
'
'    If IsNumeric(pipeCountValue) And pipeCountValue > 0 Then
'        colBr = CLng(pipeCountValue)
'    Else
'        colBr = 1
'    End If
'    If colBr < 1 Then colBr = 1
'    If colBr > 16384 Then colBr = 16384
'
'    Dim i As Long
'    For i = 1 To colBr
'        On Error Resume Next
'        Dim srcVal As Variant
'        srcVal = Range("pipeWallThicknessAvg" & i).Value
'        If Err.Number = 0 Then
'            If Not IsEmpty(srcVal) And IsNumeric(srcVal) Then
'                Range("pipeWallThickness").Cells(1, i).Value = srcVal
'            End If
'        End If
'        On Error GoTo 0
'    Next i
'
'CleanExit:
End Sub

' ================================================================
' calculate mean soil resistivity from survey rows (sheet UI, no UserForm)
' writes to soilResistivityAvg of the active Pipe section
' ================================================================
Sub button112_click()
    Call ShowSoilAvgForm(SOILAVG_TARGET_PIPE)
End Sub
