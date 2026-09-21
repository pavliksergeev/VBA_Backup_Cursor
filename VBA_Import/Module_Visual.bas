Attribute VB_Name = "Module_Visual"

Option Explicit

' ================================================================
' Module_Visual
' Visibility, colors, installation type indices
' ================================================================

' Ru() needles for GetInstallationTypeIndex — built once per session
Private needlesReady As Boolean
Private nNearVert As String
Private nNearHorizLt As String
Private nNearComb As String
Private nNearHorizGt As String
Private nActivInstr As String
Private nNearExt12 As String
Private nNearExt As String
Private nDeepNoAct As String
Private nDeepEnd As String
Private nDeepNoEnd As String
Private nVert As String
Private nHoriz As String
Private nComb As String
Private nExtended As String
Private nWithEnd As String
Private nNoEnd As String

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
    endCol = Application.WorksheetFunction.Max(startCol, Application.WorksheetFunction.Min(LAST_ANOD_COL, endCol))

    Call ApplyAnodResultRowLabels(ws)

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

    ws.rows("58:120").Hidden = True
    Call HideSpareAnodCalcRows(ws)

    If types.count > 0 Then
        Dim rowsToShow As New Collection
        Dim t As Variant
        Dim typeRows As Variant
        Dim iRow As Long
        For Each t In types
            typeRows = ResultRowsForType(CInt(t))
            For iRow = LBound(typeRows) To UBound(typeRows)
                On Error Resume Next
                rowsToShow.Add typeRows(iRow), CStr(typeRows(iRow))
                On Error GoTo 0
            Next iRow
        Next t

        Dim r As Variant
        For Each r In rowsToShow
            ws.rows(r).Hidden = False
        Next r
    End If

CleanExit:
End Sub

' обновление цветов всех колонок по способу монтажа
' updates colors for all columns based on installation types
Public Sub UpdateColorsForAllColumns(ByVal ws As Worksheet)
    On Error GoTo CleanExit
    
    Dim colBr As Long
    Dim pipeCountValue As Variant
    pipeCountValue = SafeRangeValue(ws.Range("pipeCountCP"))
    colBr = SafeNumericValue(pipeCountValue, 1)
    
    Dim startCol As Long, endCol As Long
    startCol = START_COL
    endCol = startCol + colBr - 1
    endCol = Application.WorksheetFunction.Max(startCol, Application.WorksheetFunction.Min(LAST_ANOD_COL, endCol))
    
    ' сброс заливки расчётных строк
    ' reset interior fill for calculation result rows
    ws.Range(ws.Cells(58, startCol), ws.Cells(120, endCol)).Interior.ColorIndex = xlNone
    ' сброс заливки и шрифта строки 36 (typeInstallationAG)
    ' reset interior and font for row 36 (typeInstallationAG)
    ws.Range(ws.Cells(FILTER_START_ROW + 2, startCol), ws.Cells(FILTER_START_ROW + 2, endCol)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(FILTER_START_ROW + 2, startCol), ws.Cells(FILTER_START_ROW + 2, endCol)).Font.ColorIndex = xlAutomatic
    
    Dim colIdx As Long
    For colIdx = startCol To endCol
        Call PulseProgress("updating colors", colIdx - startCol + 1, endCol - startCol + 1)
        Call ApplyColorToInstallationType(ws, colIdx)
    Next colIdx

CleanExit:
End Sub

' заливка колонки по способу монтажа
' applies color to a single column based on installation type
' заливка и шрифт typeInstallationAG (строка 36) копируются с расчётной строки
' interior fill and font of typeInstallationAG (row 36) are copied from a result row
Public Sub ApplyColorToInstallationType(ByVal ws As Worksheet, ByVal col As Long)
    On Error GoTo CleanExit
    
    Dim typeStr As String
    typeStr = Trim(CStr(ws.Cells(FILTER_START_ROW + 2, col).Value))
    
    ' пустое значение - сброс заливки и шрифта
    ' empty value - reset interior and font colors
    If typeStr = "" Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.ColorIndex = xlNone
        ws.Cells(FILTER_START_ROW + 2, col).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(58, col), ws.Cells(120, col)).Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    Dim iType As Integer
    iType = GetInstallationTypeIndex(typeStr)
    
    ' нераспознанный тип - сброс цветов
    ' unrecognized type - reset colors
    If iType = 0 Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.ColorIndex = xlNone
        ws.Cells(FILTER_START_ROW + 2, col).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(58, col), ws.Cells(120, col)).Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    Dim Color As Long
    Color = GetColorForType(iType)
    
    Dim rowsToShow As Variant
    Dim sampleRow As Long
    rowsToShow = ResultRowsForType(iType)
    sampleRow = CLng(rowsToShow(LBound(rowsToShow)))
    
    ' заливка расчётных строк этого типа монтажа
    ' set interior fill for this installation type's result rows
    Dim r As Variant
    For Each r In rowsToShow
        ws.Cells(r, col).Interior.Color = Color
    Next r
    
    ' копируем заливку и шрифт с эталонной строки на строку 36 (typeInstallationAG)
    ' copy fill and font color from the sample row to row 36 (typeInstallationAG)
    If sampleRow > 0 Then
        ws.Cells(FILTER_START_ROW + 2, col).Interior.Color = ws.Cells(sampleRow, col).Interior.Color
        ws.Cells(FILTER_START_ROW + 2, col).Font.Color = ws.Cells(sampleRow, col).Font.Color
        ws.Cells(FILTER_START_ROW + 2, col).Font.Bold = ws.Cells(sampleRow, col).Font.Bold
    End If

CleanExit:
End Sub

' наименования расчётных строк без подстроки способа монтажа
' calculation row captions without the installation-type substring
Public Sub ApplyAnodResultRowLabels(ByVal ws As Worksheet)
    Dim wasProtected As Boolean
    Dim capRp1 As String

    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect Password:=SHEET_PASSWORD
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' переходное сопротивление растеканию тока аз,  из n электродов на последний год эксплуатации (rз(t), ом)
    ' spreading resistance of the ag from n electrodes in the last year of service (rz(t), ohm)
    Call SetAnodRowCaption(ws, ROW_RESISTANCE_END_LIFE_AG, _
        Ru("041F 0435 0440 0435 0445 043E 0434 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435") & _
        Ru("0020 0440 0430 0441 0442 0435 043A 0430 043D 0438 044E 0020 0442 043E 043A 0430 0020 0410 0417 002C 0020 0020 0438 0437") & _
        Ru("0020 004E 0020 044D 043B 0435 043A 0442 0440 043E 0434 043E 0432 0020 043D 0430 0020 043F 043E 0441 043B 0435 0434 043D") & _
        Ru("0438 0439 0020 0433 043E 0434 0020 044D 043A 0441 043F 043B 0443 0430 0442 0430 0446 0438 0438 0020 0028 0052 0437 0028") & _
        Ru("0074 0029 002C 0020 041E 043C 0029"))

    ' длина рабочей части глубинного анодного заземления (lз, м)
    ' working length of the deep anode groundbed (lz, m)
    Call SetAnodRowCaption(ws, ROW_LENGTH_WORK_PART_DEEP_AG, _
        Ru("0414 043B 0438 043D 0430 0020 0440 0430 0431 043E 0447 0435 0439 0020 0447 0430 0441 0442 0438 0020 0433 043B 0443 0431") & _
        Ru("0438 043D 043D 043E 0433 043E 0020 0430 043D 043E 0434 043D 043E 0433 043E 0020 0437 0430 0437 0435 043C 043B 0435 043D") & _
        Ru("0438 044F 0020 0028 006C 0437 002C 0020 043C 0029"))

    ' сопротивление растеканию тока одного аз (rр1, ом)
    ' spreading resistance of one ag (rp1, ohm)
    capRp1 = Ru("0421 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0440 0430 0441 0442 0435 043A 0430 043D 0438 044E") & _
             Ru("0020 0442 043E 043A 0430 0020 043E 0434 043D 043E 0433 043E 0020 0410 0417 0020 0028 0052 0440 0031 002C 0020 041E 043C") & _
             Ru("0029")
    Call SetAnodRowCaption(ws, ROW_ONE_ELECTRODE_RESISTANCE_AG, capRp1)
    ' то же наименование + "горизонтальный" (без способа монтажа)
    ' same caption plus "horizontal" (without installation type)
    Call SetAnodRowCaption(ws, ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, _
        capRp1 & " " & Ru("0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439"))

    ' количество электродов аз (nз, шт)
    ' number of ag electrodes (nz, pcs)
    Call SetAnodRowCaption(ws, ROW_NUM_ELECTRODES_AG, _
        Ru("041A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 044D 043B 0435 043A 0442 0440 043E 0434 043E 0432 0020 0410 0417") & _
        Ru("0020 0028 004E 0437 002C 0020 0448 0442 0029"))

    ' масса материала электрода аз без учета активатора (gз, кг)
    ' electrode mass without activator (gz, kg)
    Call SetAnodRowCaption(ws, ROW_WEIGHT_WITHOUT_FILLING_AG, _
        Ru("041C 0430 0441 0441 0430 0020 043C 0430 0442 0435 0440 0438 0430 043B 0430 0020 044D 043B 0435 043A 0442 0440 043E 0434") & _
        Ru("0430 0020 0410 0417 0020 0431 0435 0437 0020 0443 0447 0435 0442 0430 0020 0430 043A 0442 0438 0432 0430 0442 043E 0440") & _
        Ru("0430 0020 0028 0047 0437 002C 0020 043A 0433 0029"))

    ' срок службы аз (tр, год)
    ' ag service life (tp, year)
    Call SetAnodRowCaption(ws, ROW_SERVICE_LIFE_AG, _
        Ru("0421 0440 043E 043A 0020 0441 043B 0443 0436 0431 044B 0020 0410 0417 0020 0028 0054 0440 002C 0020 0433 043E 0434 0029"))

    ' коэффициент отклонения срока службы аз от проектного (kт, )
    ' service-life deviation from design (kt, )
    Call SetAnodRowCaption(ws, ROW_SERVICE_LIFE_DEVIATION, _
        Ru("041A 043E 044D 0444 0444 0438 0446 0438 0435 043D 0442 0020 043E 0442 043A 043B 043E 043D 0435 043D 0438 044F 0020 0441") & _
        Ru("0440 043E 043A 0430 0020 0441 043B 0443 0436 0431 044B 0020 0410 0417 0020 043E 0442 0020 043F 0440 043E 0435 043A 0442") & _
        Ru("043D 043E 0433 043E 0020 0028 006B 0442 002C 0020 0029"))

    ' сопротивление растеканию тока одного аз с откорректированным количеством электродов (rр1', ом)
    ' spreading resistance of one ag with corrected electrode count (rp1', ohm)
    Call SetAnodRowCaption(ws, ROW_CORRECT_RESISTANCE_AG, _
        Ru("0421 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0440 0430 0441 0442 0435 043A 0430 043D 0438 044E") & _
        Ru("0020 0442 043E 043A 0430 0020 043E 0434 043D 043E 0433 043E 0020 0410 0417 0020 0441 0020 043E 0442 043A 043E 0440 0440") & _
        Ru("0435 043A 0442 0438 0440 043E 0432 0430 043D 043D 044B 043C 0020 043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 043C") & _
        Ru("0020 044D 043B 0435 043A 0442 0440 043E 0434 043E 0432 0020 0028 0052 0440 0031 0027 002C 0020 041E 043C 0029"))

    If wasProtected Then
        On Error Resume Next
        ws.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

' пишем подпись в колонку C (наименования строк Anod)
' write the caption into column C (Anod row names)
Private Sub SetAnodRowCaption(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal caption As String)
    ws.Cells(rowNum, 3).Value = caption
    ' убрать ошибочную запись в колонку A с предыдущей версии
    ' remove the caption mistakenly written to column A by the previous version
    If InStr(1, CStr(ws.Cells(rowNum, 1).Value), Ru("0410 0417"), vbTextCompare) > 0 Then
        ws.Cells(rowNum, 1).ClearContents
    End If
End Sub

' постоянные расчётные строки для способа монтажа (не по одной строке на тип)
' constant calculation rows for an installation type (not one row per type)
' 58 Ra(t), 59 lз, 60 Rp1 верт., 61 Rp1 гориз., 70 Nз, 80 Gз, 90 Tр, 100 kт, 110 Rp1'
Public Function ResultRowsForType(ByVal iType As Integer) As Variant
    Select Case iType
        Case 3
            ' комбинированный: Rp1 вертикальный и горизонтальный (строки 60 и 61), без Ra(t), без lз
            ' combined: vertical and horizontal Rp1 (rows 60 and 61), no Ra(t), no lз
            ResultRowsForType = Array(ROW_ONE_ELECTRODE_RESISTANCE_AG, ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG, _
                                      ROW_NUM_ELECTRODES_AG, ROW_WEIGHT_WITHOUT_FILLING_AG, _
                                      ROW_SERVICE_LIFE_AG, ROW_SERVICE_LIFE_DEVIATION, ROW_CORRECT_RESISTANCE_AG)
        Case 7, 8, 9
            ' глубинный: Ra(t), lз, Rp1 и остальные
            ' deep: Ra(t), lз, Rp1 and the rest
            ResultRowsForType = Array(ROW_RESISTANCE_END_LIFE_AG, ROW_LENGTH_WORK_PART_DEEP_AG, _
                                      ROW_ONE_ELECTRODE_RESISTANCE_AG, ROW_NUM_ELECTRODES_AG, _
                                      ROW_WEIGHT_WITHOUT_FILLING_AG, ROW_SERVICE_LIFE_AG, _
                                      ROW_SERVICE_LIFE_DEVIATION, ROW_CORRECT_RESISTANCE_AG)
        Case Else
            ' остальные типы: Ra(t), Rp1 (без lз)
            ' other types: Ra(t), Rp1 (no lз)
            ResultRowsForType = Array(ROW_RESISTANCE_END_LIFE_AG, ROW_ONE_ELECTRODE_RESISTANCE_AG, _
                                      ROW_NUM_ELECTRODES_AG, ROW_WEIGHT_WITHOUT_FILLING_AG, _
                                      ROW_SERVICE_LIFE_AG, ROW_SERVICE_LIFE_DEVIATION, ROW_CORRECT_RESISTANCE_AG)
    End Select
End Function

' скрыть пустые копии по типам: в диапазоне 58-120 остаются только постоянные расчётные строки
' hide empty per-type copies: in rows 58-120 only the constant result rows remain
Public Sub HideSpareAnodCalcRows(ByVal ws As Worksheet)
    Dim r As Long
    Dim keepRow As Boolean
    For r = 58 To 120
        keepRow = (r = ROW_RESISTANCE_END_LIFE_AG) Or _
                  (r = ROW_LENGTH_WORK_PART_DEEP_AG) Or _
                  (r = ROW_ONE_ELECTRODE_RESISTANCE_AG) Or _
                  (r = ROW_ONE_ELECTRODE_RESISTANCE_HORIZ_AG) Or _
                  (r = ROW_NUM_ELECTRODES_AG) Or _
                  (r = ROW_WEIGHT_WITHOUT_FILLING_AG) Or _
                  (r = ROW_SERVICE_LIFE_AG) Or _
                  (r = ROW_SERVICE_LIFE_DEVIATION) Or _
                  (r = ROW_CORRECT_RESISTANCE_AG)
        If Not keepRow Then ws.rows(r).Hidden = True
    Next r
End Sub

' Returns installation type index based on text description (case-insensitive, Cyrillic-safe)
Public Function GetInstallationTypeIndex(ByVal typeStr As String) As Integer
    Dim s As String
    s = LCase$(Trim$(typeStr))
    If s = "" Then
        GetInstallationTypeIndex = 0
        Exit Function
    End If

    Call EnsureInstallNeedles
    
    Select Case True
        ' подповерхностный вертикальный
        ' near-surface vertical
        Case ContainsText(s, nNearVert)
            GetInstallationTypeIndex = 1
        ' подповерхностный горизонтальный le<h
        ' near-surface horizontal le<h
        Case ContainsText(s, nNearHorizLt)
            GetInstallationTypeIndex = 2
        ' подповерхностный комбинированный
        ' near-surface combined
        Case ContainsText(s, nNearComb)
            GetInstallationTypeIndex = 3
        ' подповерхностный горизонтальный le>h
        ' near-surface horizontal le>h
        ' активатором
        ' activator
        Case ContainsText(s, nNearHorizGt) And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 4
        ' подповерхностный протяженный le>12h
        ' near-surface extended le>12h
        ' активатором
        ' activator
        Case ContainsText(s, nNearExt12) And Not ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 5
        ' подповерхностный протяженный
        ' near-surface extended
        ' активатором
        ' activator
        Case ContainsText(s, nNearExt) And ContainsText(s, nActivInstr) And ContainsText(s, "le>12h")
            GetInstallationTypeIndex = 6
        ' глубинный без активатора
        ' deep without activator
        Case ContainsText(s, nDeepNoAct)
            GetInstallationTypeIndex = 7
        ' глубинный с выходом торца
        ' deep with end at surface
        ' активатором
        ' activator
        Case ContainsText(s, nDeepEnd) And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 8
        ' глубинный без выхода торца
        ' deep without end at surface
        ' активатором
        ' activator
        Case ContainsText(s, nDeepNoEnd) And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 9
        ' вертикальный
        ' vertical
        ' горизонтальный
        ' horizontal
        Case ContainsText(s, nVert) And Not ContainsText(s, nHoriz)
            GetInstallationTypeIndex = 1
        ' горизонтальный
        ' horizontal
        Case ContainsText(s, nHoriz) And ContainsText(s, "le<h")
            GetInstallationTypeIndex = 2
        ' комбинированный
        ' combined
        Case ContainsText(s, nComb)
            GetInstallationTypeIndex = 3
        ' горизонтальный
        ' horizontal
        ' активатором
        ' activator
        Case ContainsText(s, nHoriz) And ContainsText(s, "le>h") And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 4
        ' протяженный
        ' extended
        ' активатором
        ' activator
        Case ContainsText(s, nExtended) And ContainsText(s, "le>12h") And Not ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 5
        ' протяженный
        ' extended
        ' активатором
        ' activator
        Case ContainsText(s, nExtended) And ContainsText(s, nActivInstr) And ContainsText(s, "le>12h")
            GetInstallationTypeIndex = 6
        ' с выходом торца
        ' with end at surface
        ' активатором
        ' activator
        Case ContainsText(s, nWithEnd) And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 8
        ' без выхода торца
        ' without end at surface
        ' активатором
        ' activator
        Case ContainsText(s, nNoEnd) And ContainsText(s, nActivInstr)
            GetInstallationTypeIndex = 9
        Case Else
            GetInstallationTypeIndex = 0
    End Select
End Function

Private Sub EnsureInstallNeedles()
    If needlesReady Then Exit Sub

    ' подповерхностный вертикальный
    nNearVert = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0432 0435 0440 0442 0438 043A 0430 043B 044C 043D 044B 0439")
    ' подповерхностный горизонтальный le<h
    nNearHorizLt = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439 0020 006C 0065 003C") & Ru("0068")
    ' подповерхностный комбинированный
    nNearComb = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043A 043E 043C 0431 0438 043D 0438 0440 043E 0432 0430 043D 043D 044B 0439")
    ' подповерхностный горизонтальный le>h
    nNearHorizGt = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439 0020 006C 0065 003E") & Ru("0068")
    ' активатором
    nActivInstr = Ru("0430 043A 0442 0438 0432 0430 0442 043E 0440 043E 043C")
    ' подповерхностный протяженный le>12h
    nNearExt12 = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439 0020 006C 0065 003E 0031 0032 0068")
    ' подповерхностный протяженный
    nNearExt = Ru("043F 043E 0434 043F 043E 0432 0435 0440 0445 043D 043E 0441 0442 043D 044B 0439 0020 043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439")
    ' глубинный без активатора
    nDeepNoAct = Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0431 0435 0437 0020 0430 043A 0442 0438 0432 0430 0442 043E 0440 0430")
    ' глубинный с выходом торца
    nDeepEnd = Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0441 0020 0432 044B 0445 043E 0434 043E 043C 0020 0442 043E 0440 0446 0430")
    ' глубинный без выхода торца
    nDeepNoEnd = Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0431 0435 0437 0020 0432 044B 0445 043E 0434 0430 0020 0442 043E 0440 0446 0430")
    ' вертикальный
    nVert = Ru("0432 0435 0440 0442 0438 043A 0430 043B 044C 043D 044B 0439")
    ' горизонтальный
    nHoriz = Ru("0433 043E 0440 0438 0437 043E 043D 0442 0430 043B 044C 043D 044B 0439")
    ' комбинированный
    nComb = Ru("043A 043E 043C 0431 0438 043D 0438 0440 043E 0432 0430 043D 043D 044B 0439")
    ' протяженный
    nExtended = Ru("043F 0440 043E 0442 044F 0436 0435 043D 043D 044B 0439")
    ' с выходом торца
    nWithEnd = Ru("0441 0020 0432 044B 0445 043E 0434 043E 043C 0020 0442 043E 0440 0446 0430")
    ' без выхода торца
    nNoEnd = Ru("0431 0435 0437 0020 0432 044B 0445 043E 0434 0430 0020 0442 043E 0440 0446 0430")

    needlesReady = True
End Sub

Private Function ContainsText(ByVal haystack As String, ByVal needle As String) As Boolean
    ContainsText = (InStr(1, haystack, needle, vbTextCompare) > 0)
End Function

' цвет заливки для индекса способа монтажа (1-9)
' returns color for installation type index
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
