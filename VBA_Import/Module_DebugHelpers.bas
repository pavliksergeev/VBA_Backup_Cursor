Attribute VB_Name = "Module_DebugHelpers"

Option Explicit

' ================================================================
' Module_DebugHelpers
' Debug and test helpers
' ================================================================

' Resets all flags to default values
Public Sub ResetFlags()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    DoEvents
End Sub

' Returns current flags state as string
Public Function GetFlagsState() As String
    GetFlagsState = "EnableEvents=" & Application.EnableEvents & _
                    ", ScreenUpdating=" & Application.ScreenUpdating & _
                    ", Calculation=" & Application.Calculation
End Function

' Clears a test column (filters and data rows 34-49)
Public Sub ClearTestColumn(ByVal ws As Worksheet, Optional ByVal col As Long = 4)
    On Error Resume Next
    
    Application.EnableEvents = False
    
    ws.Cells(34, col).ClearContents
    ws.Cells(35, col).ClearContents
    ws.Cells(36, col).ClearContents
    ws.Cells(37, col).ClearContents
    ws.Cells(38, col).ClearContents
    
    Dim i As Long
    For i = 39 To 49
        ws.Cells(i, col).ClearContents
    Next i
    
    Application.EnableEvents = True
End Sub

' Forces auto fill of filters from model for a column
Public Sub ForceAutoFillColumn(ByVal ws As Worksheet, Optional ByVal col As Long = 4, Optional ByVal modelValue As String = "")
    On Error GoTo CleanExit
    
    If modelValue = "" Then
        modelValue = ws.Cells(38, col).Value
    End If
    
    If modelValue = "" Or IsEmpty(modelValue) Then
        ' в ячейке d
        '  нет модели!
        '  no model!
        MsgBox Ru("0432 0020 044F 0447 0435 0439 043A 0435 0020 0064") & col & Ru("0020 043D 0435 0442 0020 043C 043E 0434 0435 043B 0438 0021"), vbExclamation
        Exit Sub
    End If
    
    Application.EnableEvents = True
    Call Module_ValidationLogic.AutoFillFiltersFromModel(ws, col, modelValue)

CleanExit:
End Sub

' Test validation inside (creates validation for cell D34)
Public Sub TestValidationInside(ByVal ws As Worksheet)
    Dim materialList As String
    materialList = Module_ValidationLists.GetAllMaterials()
    
    If materialList <> "" Then
        On Error Resume Next
        With ws.Cells(34, 4).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=materialList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        If Err.Number = 0 Then
            ' валидация создана для d34!
            ' validation created for d34!
            ' список: 
            ' list: 
            MsgBox Ru("0432 0430 043B 0438 0434 0430 0446 0438 044F 0020 0441 043E 0437 0434 0430 043D 0430 0020 0434 043B 044F 0020 0064 0033 0034 0021") & vbCrLf & _
                   Ru("0441 043F 0438 0441 043E 043A 003A 0020") & materialList, vbInformation
        Else
            ' ошибка создания валидации: 
            ' create error валandдацandand: 
            MsgBox Ru("043E 0448 0438 0431 043A 0430 0020 0441 043E 0437 0434 0430 043D 0438 044F 0020 0432 0430 043B 0438 0434 0430 0446 0438 0438 003A 0020") & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    Else
        ' список материалов пуст!
        ' material list is empty!
        MsgBox Ru("0441 043F 0438 0441 043E 043A 0020 043C 0430 0442 0435 0440 0438 0430 043B 043E 0432 0020 043F 0443 0441 0442 0021"), vbExclamation
    End If
End Sub
