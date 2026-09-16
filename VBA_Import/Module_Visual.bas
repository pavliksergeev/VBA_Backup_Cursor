Attribute VB_Name = "Module_Visual"

Option Explicit

' ================================================================
' Module_Visual
' Visibility, colors, installation type indices
' ================================================================

' Updates row visibility based on installation types in all columns
Public Sub UpdateVisibilityForAllColumns(ByVal ws As Worksheet)
    On Error GoTo CleanExit

    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = SafeRangeValue(ws.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)
    
    Dim startCol As Long
    startCol = START_COL
    Dim endCol As Long
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, Application.WorksheetFunction.Min(MAX_COL, endCol))

    Dim types As New Collection
    Dim colIdx As Long
    Dim typeStr As String
    Dim iType As Integer
    
    For colIdx = startCol To endCol
        typeStr = Trim(CStr(ws.Cells(FILTER_START_ROW + 2, colIdx).Value))
        If typeStr <> "" Then
            iType = GetInstallationTypeIndex(typeStr)
            If iType > 0 Then
                On Error Resume Next
                types.Add iType, CStr(iType)
                On Error GoTo 0
            End If
        End If
    Next colIdx

    ws.rows("59:120").Hidden = True

    If types.count > 0 Then
        Dim rowsToShow As New Collection
        Dim t As Variant
        For Each t In types
            Select Case t
                Case 1
                    rowsToShow.Add 60
                    rowsToShow.Add 70
                    rowsToShow.Add 80
                    rowsToShow.Add 90
                    rowsToShow.Add 100
                    rowsToShow.Add 110
                Case 2
                    rowsToShow.Add 61
                    rowsToShow.Add 71
                    rowsToShow.Add 81
                    rowsToShow.Add 91
                    rowsToShow.Add 101
                    rowsToShow.Add 111
                Case 3
                    rowsToShow.Add 60
                    rowsToShow.Add 61
                    rowsToShow.Add 72
                    rowsToShow.Add 82
                    rowsToShow.Add 92
                    rowsToShow.Add 102
                    rowsToShow.Add 112
                Case 4
                    rowsToShow.Add 63
                    rowsToShow.Add 73
                    rowsToShow.Add 83
                    rowsToShow.Add 93
                    rowsToShow.Add 103
                    rowsToShow.Add 113
                Case 5
                    rowsToShow.Add 64
                    rowsToShow.Add 74
                    rowsToShow.Add 84
                    rowsToShow.Add 94
                    rowsToShow.Add 104
                    rowsToShow.Add 114
                Case 6
                    rowsToShow.Add 65
                    rowsToShow.Add 75
                    rowsToShow.Add 85
                    rowsToShow.Add 95
                    rowsToShow.Add 105
                    rowsToShow.Add 115
                Case 7
                    rowsToShow.Add 59
                    rowsToShow.Add 66
                    rowsToShow.Add 76
                    rowsToShow.Add 86
                    rowsToShow.Add 96
                    rowsToShow.Add 106
                    rowsToShow.Add 116
                Case 8
                    rowsToShow.Add 59
                    rowsToShow.Add 67
                    rowsToShow.Add 77
                    rowsToShow.Add 87
                    rowsToShow.Add 97
                    rowsToShow.Add 107
                    rowsToShow.Add 117
                Case 9
                    rowsToShow.Add 59
                    rowsToShow.Add 68
                    rowsToShow.Add 78
                    rowsToShow.Add 88
                    rowsToShow.Add 98
                    rowsToShow.Add 108
                    rowsToShow.Add 118
            End Select
        Next t

        Dim r As Variant
        For Each r In rowsToShow
            ws.rows(r).Hidden = False
        Next r
    End If

CleanExit:
End Sub

' Updates colors for all columns based on installation types
Public Sub UpdateColorsForAllColumns(ByVal ws As Worksheet)
    On Error GoTo CleanExit
    
    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = SafeRangeValue(ws.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)
    
    Dim startCol As Long, endCol As Long
    startCol = START_COL
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, Application.WorksheetFunction.Min(MAX_COL, endCol))
    
    ' reset interior fill for rows 59-120
    ws.Range(ws.Cells(59, startCol), ws.Cells(120, endCol)).Interior.ColorIndex = xlNone
    ' reset interior and font for row 36 (typeInstallationAG)
    ws.Range(ws.Cells(FILTER_START_ROW + 2, startCol), ws.Cells(FILTER_START_ROW + 2, endCol)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(FILTER_START_ROW + 2, startCol), ws.Cells(FILTER_START_ROW + 2, endCol)).Font.ColorIndex = xlAutomatic
    
    Dim colIdx As Long
    For colIdx = startCol To endCol
        Call ApplyColorToInstallationType(ws, colIdx)
    Next colIdx

CleanExit:
End Sub

' Applies color to a single column based on installation type
' - interior fill and font color of typeInstallationAG (row 36) are copied
'   from the corresponding synchronized row (59-120)
Public Sub ApplyColorToInstallationType(ByVal ws As Worksheet, ByVal col As Long)
    On Error GoTo CleanExit
    
    Dim typeStr As String
    typeStr = Trim(CStr(ws.Cells(FILTER_START_ROW + 2, col).Value))
    
    ' empty value - reset interior and font colors
    If typeStr = "" Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.ColorIndex = xlNone
        ws.Cells(FILTER_START_ROW + 2, col).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(59, col), ws.Cells(120, col)).Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    Dim iType As Integer
    iType = GetInstallationTypeIndex(typeStr)
    
    ' unrecognized type - reset colors
    If iType = 0 Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.ColorIndex = xlNone
        ws.Cells(FILTER_START_ROW + 2, col).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(59, col), ws.Cells(120, col)).Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    Dim Color As Long
    Color = GetColorForType(iType)
    
    ' build list of synchronized rows and pick a sample row
    Dim rowsToShow As New Collection
    Dim sampleRow As Long
    sampleRow = 0
    
    Select Case iType
        Case 1
            rowsToShow.Add 60: rowsToShow.Add 70: rowsToShow.Add 80: rowsToShow.Add 90: rowsToShow.Add 100: rowsToShow.Add 110
            sampleRow = 60
        Case 2
            rowsToShow.Add 61: rowsToShow.Add 71: rowsToShow.Add 81: rowsToShow.Add 91: rowsToShow.Add 101: rowsToShow.Add 111
            sampleRow = 61
        Case 3
            rowsToShow.Add 60: rowsToShow.Add 61: rowsToShow.Add 72: rowsToShow.Add 82: rowsToShow.Add 92: rowsToShow.Add 102: rowsToShow.Add 112
            sampleRow = 60
        Case 4
            rowsToShow.Add 63: rowsToShow.Add 73: rowsToShow.Add 83: rowsToShow.Add 93: rowsToShow.Add 103: rowsToShow.Add 113
            sampleRow = 63
        Case 5
            rowsToShow.Add 64: rowsToShow.Add 74: rowsToShow.Add 84: rowsToShow.Add 94: rowsToShow.Add 104: rowsToShow.Add 114
            sampleRow = 64
        Case 6
            rowsToShow.Add 65: rowsToShow.Add 75: rowsToShow.Add 85: rowsToShow.Add 95: rowsToShow.Add 105: rowsToShow.Add 115
            sampleRow = 65
        Case 7
            rowsToShow.Add 59: rowsToShow.Add 66: rowsToShow.Add 76: rowsToShow.Add 86: rowsToShow.Add 96: rowsToShow.Add 106: rowsToShow.Add 116
            sampleRow = 59
        Case 8
            rowsToShow.Add 59: rowsToShow.Add 67: rowsToShow.Add 77: rowsToShow.Add 87: rowsToShow.Add 97: rowsToShow.Add 107: rowsToShow.Add 117
            sampleRow = 59
        Case 9
            rowsToShow.Add 59: rowsToShow.Add 68: rowsToShow.Add 78: rowsToShow.Add 88: rowsToShow.Add 98: rowsToShow.Add 108: rowsToShow.Add 118
            sampleRow = 59
    End Select
    
    ' set interior fill for synchronized rows
    Dim r As Variant
    For Each r In rowsToShow
        ws.Cells(r, col).Interior.Color = Color
    Next r
    
    ' copy fill and font color from the sample row to row 36 (typeInstallationAG)
    If sampleRow > 0 Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.Color = ws.Cells(sampleRow, col).Interior.Color
        ws.Cells(FILTER_START_ROW + 2, col).Font.Color = ws.Cells(sampleRow, col).Font.Color
        ws.Cells(FILTER_START_ROW + 2, col).Font.Bold = ws.Cells(sampleRow, col).Font.Bold
    End If

CleanExit:
End Sub

' Returns installation type index based on text description (case-insensitive, Cyrillic-safe)
Public Function GetInstallationTypeIndex(ByVal typeStr As String) As Integer
    Dim s As String
    s = LCase$(Trim$(typeStr))
    If s = "" Then
        GetInstallationTypeIndex = 0
        Exit Function
    End If
    
    Select Case True
        ' подповерхностный вертикальный
        ' near-surface vertical
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0432 0435 0440 0442 0438 043A 0430 043B 044C 043D 044B 0439"))
            GetInstallationTypeIndex = 1
        ' подповерхностный горизонтальный le<h
        ' near-surface horizontal le<h
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439 0020 006C 0065 003C") & Ru("0068"))
            GetInstallationTypeIndex = 2
        ' подповерхностный комбинированный
        ' near-surface combined
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043A 043E 043C 0431 0438 043D 0438 0440 043E 0432 0430 043D 043D 044B 0439"))
            GetInstallationTypeIndex = 3
        ' подповерхностный горизонтальный le>h
        ' near-surface horizontal le>h
        ' активатором
        ' activator
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439 0020 006C 0065 003E") & Ru("0068")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 4
        ' подповерхностный протяженный le>12h
        ' near-surface extended le>12h
        ' активатором
        ' activator
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439 0020 006C 0065 003E 0031 0032 0068")) And Not ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 5
        ' подповерхностный протяженный
        ' near-surface extended
        ' активатором
        ' activator
        Case ContainsText(s, Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C")) And ContainsText(s, "le>12h")
            GetInstallationTypeIndex = 6
        ' глубинный без активатора
        ' deep without activator
        Case ContainsText(s, Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0431 0435 0437 0020 0430 043A 0442 0438 0432 0430 0442 043E 0440 0430"))
            GetInstallationTypeIndex = 7
        ' глубинный с выходом торца
        ' deep with end at surface
        ' активатором
        ' activator
        Case ContainsText(s, Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0441 0020 0432 044B 0445 043E 0434 043E 043C 0020 0442 043E 0440 0446 0430")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 8
        ' глубинный без выхода торца
        ' deep without end at surface
        ' активатором
        ' activator
        Case ContainsText(s, Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0431 0435 0437 0020 0432 044B 0445 043E 0434 0430 0020 0442 043E 0440 0446 0430")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 9
        ' вертикальный
        ' vertical
        ' горизонтальный
        ' horizontal
        Case ContainsText(s, Ru("0432 0435 0440 0442 0438 043A 0430 043B 044C 043D 044B 0439")) And Not ContainsText(s, Ru("0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439"))
            GetInstallationTypeIndex = 1
        ' горизонтальный
        ' horizontal
        Case ContainsText(s, Ru("0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439")) And ContainsText(s, "le<h")
            GetInstallationTypeIndex = 2
        ' комбинированный
        ' combined
        Case ContainsText(s, Ru("043A 043E 043C 0431 0438 043D 0438 0440 043E 0432 0430 043D 043D 044B 0439"))
            GetInstallationTypeIndex = 3
        ' горизонтальный
        ' horizontal
        ' активатором
        ' activator
        Case ContainsText(s, Ru("0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439")) And ContainsText(s, "le>h") And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 4
        ' протяженный
        ' extended
        ' активатором
        ' activator
        Case ContainsText(s, Ru("043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439")) And ContainsText(s, "le>12h") And Not ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 5
        ' протяженный
        ' extended
        ' активатором
        ' activator
        Case ContainsText(s, Ru("043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C")) And ContainsText(s, "le>12h")
            GetInstallationTypeIndex = 6
        ' с выходом торца
        ' with end at surface
        ' активатором
        ' activator
        Case ContainsText(s, Ru("0441 0020 0432 044B 0445 043E 0434 043E 043C 0020 0442 043E 0440 0446 0430")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 8
        ' без выхода торца
        ' without end at surface
        ' активатором
        ' activator
        Case ContainsText(s, Ru("0431 0435 0437 0020 0432 044B 0445 043E 0434 0430 0020 0442 043E 0440 0446 0430")) And ContainsText(s, Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C"))
            GetInstallationTypeIndex = 9
        Case Else
            GetInstallationTypeIndex = 0
    End Select
End Function

Private Function ContainsText(ByVal haystack As String, ByVal needle As String) As Boolean
    ContainsText = (InStr(1, haystack, needle, vbTextCompare) > 0)
End Function

' Returns color for installation type index
Public Function GetColorForType(ByVal iType As Integer) As Long
    Select Case iType
        Case 1: GetColorForType = RGB(255, 200, 200)
        Case 2: GetColorForType = RGB(200, 255, 200)
        Case 3: GetColorForType = RGB(200, 200, 255)
        Case 4: GetColorForType = RGB(255, 255, 200)
        Case 5: GetColorForType = RGB(255, 200, 255)
        Case 6: GetColorForType = RGB(200, 255, 255)
        Case 7: GetColorForType = RGB(255, 150, 150)
        Case 8: GetColorForType = RGB(150, 255, 150)
        Case 9: GetColorForType = RGB(150, 150, 255)
        Case Else: GetColorForType = RGB(255, 255, 255)
    End Select
End Function
