Attribute VB_Name = "Module_DataCleanup"

' ================================================================
' module: Module_DataCleanup
' ListAG button: FullDataCleanupAndRefresh -> AuditTableAG (read-only report)
' ListCP button: AuditTableCP (read-only report)
' CleanTable* still write to tables; the buttons do not call them.
' ================================================================
Option Explicit

Private auditErrCount As Long
Private auditWarnCount As Long

' ================================================================
' кнопка на ListAG: отчёт без изменения ячеек справочников
' ListAG button: report only, no writes to catalog tables
' ================================================================
Public Sub FullDataCleanupAndRefresh()
    Call AuditTableAG
End Sub

Public Sub AuditTableAG()
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As XlCalculation
    Dim oldAlerts As Boolean
    Dim wsReport As Worksheet
    Dim row As Long
    Dim tblAG As ListObject
    Dim tblMount As ListObject
    Dim colMat As Long, colMount As Long, colShip As Long, colModel As Long
    Dim nRow As Long
    Dim iRow As Long
    Dim body As Variant
    Dim matRaw As String, mountRaw As String, shipRaw As String, modelRaw As String
    Dim modelNorm As String, mountNorm As String
    Dim seenModels As Collection
    Dim mountSet As Collection
    Dim firstRow As Long
    Dim nmChk As Variant

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    On Error GoTo CleanExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Call SetStatusBar("audit TableAG...")

    auditErrCount = 0
    auditWarnCount = 0

    Set wsReport = EnsureAuditSheet("AuditAG", TABLE_AG & " / " & TABLE_MOUNTING)
    row = 5

    Set tblAG = GetListObjectOnSheet(SHEET_LIST_AG, TABLE_AG)
    If tblAG Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, 0, "-", "-", "-", _
                           Ru("0442 0430 0431 043B 0438 0446 0430 0020") & TABLE_AG & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430"), True)
        row = row + 1
        GoTo FinishReport
    End If
    If tblAG.DataBodyRange Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, 0, "-", "-", "-", _
                           Ru("043D 0435 0442 0020 0441 0442 0440 043E 043A 0020 0434 0430 043D 043D 044B 0445"), True)
        row = row + 1
        GoTo FinishReport
    End If

    colMat = ListColIndexByHeader(tblAG, Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430"))
    colMount = ListColIndexByHeader(tblAG, Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"))
    colShip = ListColIndexByHeader(tblAG, Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F"))
    colModel = ListColIndexByHeader(tblAG, "model")
    If colMat = 0 Or colMount = 0 Or colShip = 0 Or colModel = 0 Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, 0, "-", "-", "-", _
                           Ru("043D 0435 0020 043D 0430 0439 0434 0435 043D 044B 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 044B 0435 0020 043A 043E 043B 043E 043D 043A 0438"), True)
        row = row + 1
        GoTo FinishReport
    End If

    body = As2DAudit(tblAG.DataBodyRange.Value)
    nRow = UBound(body, 1)
    Set seenModels = New Collection
    Set tblMount = FindTableMountingAudit()
    Set mountSet = LoadMountTypeSet(tblMount)

    For iRow = 1 To nRow
        Call PulseProgress("audit TableAG", iRow, nRow)
        matRaw = CellTextAudit(body, iRow, colMat)
        mountRaw = CellTextAudit(body, iRow, colMount)
        shipRaw = CellTextAudit(body, iRow, colShip)
        modelRaw = CellTextAudit(body, iRow, colModel)

        If Len(Trim$(matRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                               Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430"), "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 043E 0435 0020 043F 043E 043B 0435"), True)
            row = row + 1
        Else
            row = row + WriteWhitespaceRow(wsReport, row, TABLE_AG, iRow, Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430"), matRaw)
        End If

        If Len(Trim$(mountRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                               Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"), "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 043E 0435 0020 043F 043E 043B 0435"), True)
            row = row + 1
        Else
            row = row + WriteWhitespaceRow(wsReport, row, TABLE_AG, iRow, Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"), mountRaw)
            mountNorm = Module_ValidationLists.NormalizeFilterText(mountRaw)
            If Not KeyExists(mountSet, mountNorm) Then
                Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                                   "typeMountingAG", mountRaw, mountNorm, _
                                   Ru("043D 0435 0442 0020 0432 0020") & TABLE_MOUNTING & Ru("0020 0028 043A 0430 0442 0430 043B 043E 0436 043D 044B 0439 0020 0442 0435 0441 0442 0020 043F 0440 043E 043F 0443 0441 043A 0430 0435 0442 0029"), True)
                row = row + 1
            End If
        End If

        If Len(Trim$(shipRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                               Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F"), "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 043E 0435 0020 043F 043E 043B 0435"), True)
            row = row + 1
        Else
            row = row + WriteWhitespaceRow(wsReport, row, TABLE_AG, iRow, Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F"), shipRaw)
        End If

        If Len(Trim$(modelRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                               "typeAG", "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 043E 0435 0020 043F 043E 043B 0435"), True)
            row = row + 1
        Else
            row = row + WriteWhitespaceRow(wsReport, row, TABLE_AG, iRow, "typeAG", modelRaw)
            modelNorm = Module_ValidationLists.NormalizeFilterText(modelRaw)
            If Len(modelNorm) > 0 Then
                firstRow = CollectionLongValue(seenModels, modelNorm)
                If firstRow = 0 Then
                    seenModels.Add CStr(iRow), modelNorm
                Else
                    Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_AG, iRow, _
                                       "typeAG", modelRaw, modelNorm, _
                                       Ru("0434 0443 0431 043B 0438 043A 0430 0442 0020 0074 0079 0070 0065 0041 0047 002C 0020 043F 0435 0440 0432 0430 044F 0020 0441 0442 0440 043E 043A 0430 0020") & firstRow, True)
                    row = row + 1
                End If
            End If
        End If
    Next iRow

    If tblMount Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_MOUNTING, 0, "-", "-", "-", _
                           Ru("0442 0430 0431 043B 0438 0446 0430 0020") & TABLE_MOUNTING & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430"), True)
        row = row + 1
    ElseIf tblMount.DataBodyRange Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_MOUNTING, 0, "-", "-", "-", _
                           Ru("043D 0435 0442 0020 0441 0442 0440 043E 043A 0020 0434 0430 043D 043D 044B 0445"), True)
        row = row + 1
    Else
        row = row + AuditMountingPairs(wsReport, row, tblMount)
    End If

    For Each nmChk In Array("AG_Material", "AG_MountType", "AG_Completion", "AG_Model")
        row = row + AuditNamedRangeRows(wsReport, row, CStr(nmChk), nRow)
    Next nmChk

FinishReport:
    Call FinalizeAuditSheet(wsReport, row, nRow)
    wsReport.Activate
    Application.ScreenUpdating = True
    Call ShowAuditDoneMsg("AuditAG")

CleanExit:
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    If Err.Number <> 0 Then
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description, vbCritical
    End If
End Sub

' ================================================================
' кнопка на ListCP: отчёт без изменения ячеек TableCP
' ListCP button: report only, no writes to TableCP
' ================================================================
Public Sub AuditTableCP()
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As XlCalculation
    Dim oldAlerts As Boolean
    Dim wsReport As Worksheet
    Dim row As Long
    Dim tblCP As ListObject
    Dim colType As Long, colModel As Long, colCur As Long, colVolt As Long
    Dim nRow As Long
    Dim iRow As Long
    Dim body As Variant
    Dim typeRaw As String, modelRaw As String, curRaw As String, voltRaw As String
    Dim typeNorm As String
    Dim seenTypes As Collection
    Dim firstRow As Long
    Dim nUnique As Long
    Dim rngList As Range
    Dim nList As Long

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    On Error GoTo CleanExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Call SetStatusBar("audit TableCP...")

    auditErrCount = 0
    auditWarnCount = 0

    Set wsReport = EnsureAuditSheet("AuditCP", TABLE_CP & " / typeCP")
    row = 5

    Set tblCP = GetListObjectOnSheet(SHEET_LIST_CP, TABLE_CP)
    If tblCP Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_CP, 0, "-", "-", "-", _
                           Ru("0442 0430 0431 043B 0438 0446 0430 0020") & TABLE_CP & Ru("0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430"), True)
        row = row + 1
        GoTo FinishReport
    End If
    If tblCP.DataBodyRange Is Nothing Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_CP, 0, "-", "-", "-", _
                           Ru("043D 0435 0442 0020 0441 0442 0440 043E 043A 0020 0434 0430 043D 043D 044B 0445"), True)
        row = row + 1
        GoTo FinishReport
    End If

    colType = ListColIndexByHeader(tblCP, "typeCP")
    If colType = 0 Then colType = ListColIndexByHeaderPart(tblCP, "model")
    colModel = ListColIndexByHeader(tblCP, "modelCP")
    colCur = ListColIndexByHeader(tblCP, "nominalOutputCurrenCP")
    colVolt = ListColIndexByHeader(tblCP, "nominalOutputVoltageCP")
    If colType = 0 Then
        Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_CP, 0, "typeCP", "-", "-", _
                           Ru("043A 043E 043B 043E 043D 043A 0430 0020 0074 0079 0070 0065 0043 0050 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430"), True)
        row = row + 1
        GoTo FinishReport
    End If

    body = As2DAudit(tblCP.DataBodyRange.Value)
    nRow = UBound(body, 1)
    Set seenTypes = New Collection

    For iRow = 1 To nRow
        Call PulseProgress("audit TableCP", iRow, nRow)
        typeRaw = CellTextAudit(body, iRow, colType)
        If colModel > 0 Then modelRaw = CellTextAudit(body, iRow, colModel) Else modelRaw = ""
        If colCur > 0 Then curRaw = CellTextAudit(body, iRow, colCur) Else curRaw = "x"
        If colVolt > 0 Then voltRaw = CellTextAudit(body, iRow, colVolt) Else voltRaw = "x"

        If Len(Trim$(typeRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_CP, iRow, _
                               "typeCP", "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043E 0431 044F 0437 0430 0442 0435 043B 044C 043D 043E 0435 0020 043F 043E 043B 0435"), True)
            row = row + 1
        Else
            row = row + WriteWhitespaceRow(wsReport, row, TABLE_CP, iRow, "typeCP", typeRaw)
            typeNorm = Module_ValidationLists.NormalizeFilterText(typeRaw)
            If Len(typeNorm) > 0 Then
                firstRow = CollectionLongValue(seenTypes, typeNorm)
                If firstRow = 0 Then
                    seenTypes.Add CStr(iRow), typeNorm
                Else
                    Call WriteAuditRow(wsReport, row, Ru("043E 0448 0438 0431 043A 0430"), TABLE_CP, iRow, _
                                       "typeCP", typeRaw, typeNorm, _
                                       Ru("0434 0443 0431 043B 0438 043A 0430 0442 0020 0074 0079 0070 0065 0043 0050 002C 0020 043F 0435 0440 0432 0430 044F 0020 0441 0442 0440 043E 043A 0430 0020") & firstRow, True)
                    row = row + 1
                End If
            End If
        End If

        If colModel > 0 Then
            If Len(Trim$(modelRaw)) = 0 Then
                Call WriteAuditRow(wsReport, row, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), TABLE_CP, iRow, _
                                   "modelCP", "", "", _
                                   Ru("043F 0443 0441 0442 043E 0435 0020 043F 043E 043B 0435"), False)
                row = row + 1
            Else
                row = row + WriteWhitespaceRow(wsReport, row, TABLE_CP, iRow, "modelCP", modelRaw)
            End If
        End If

        If colCur > 0 And Len(Trim$(curRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), TABLE_CP, iRow, _
                               "nominalOutputCurrenCP", "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043F 043E 043B 0435"), False)
            row = row + 1
        End If
        If colVolt > 0 And Len(Trim$(voltRaw)) = 0 Then
            Call WriteAuditRow(wsReport, row, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), TABLE_CP, iRow, _
                               "nominalOutputVoltageCP", "", "", _
                               Ru("043F 0443 0441 0442 043E 0435 0020 043F 043E 043B 0435"), False)
            row = row + 1
        End If
    Next iRow

    nUnique = seenTypes.Count
    nList = -1
    On Error Resume Next
    Set rngList = Nothing
    Set rngList = thisWorkbook.Worksheets(SHEET_ANOD).names("cp_List").RefersToRange
    If Not rngList Is Nothing Then nList = rngList.rows.count
    Err.Clear
    On Error GoTo CleanExit

    If nList < 0 Then
        Call WriteAuditRow(wsReport, row, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), "cp_List", 0, _
                           "cp_List", "-", CStr(nUnique), _
                           Ru("0438 043C 044F 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 002E 0020 0432 044B 043F 043E 043B 043D 0438 0442 0435 0020") & "UpdateCPValidationList", False)
        row = row + 1
    ElseIf nList < nUnique Then
        Call WriteAuditRow(wsReport, row, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), "cp_List", 0, _
                           "cp_List", CStr(nList), CStr(nUnique), _
                           Ru("0441 0442 0440 043E 043A 0020 0432 0020 0438 043C 0435 043D 0438 0020 043C 0435 043D 044C 0448 0435 002C 0020 0447 0435 043C 0020 0443 043D 0438 043A 0430 043B 044C 043D 044B 0445 0020 0074 0079 0070 0065 0043 0050 002E 0020 0432 044B 043F 043E 043B 043D 0438 0442 0435 0020") & "UpdateCPValidationList", False)
        row = row + 1
    End If

FinishReport:
    Call FinalizeAuditSheet(wsReport, row, nRow)
    wsReport.Activate
    Application.ScreenUpdating = True
    Call ShowAuditDoneMsg("AuditCP")

CleanExit:
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    If Err.Number <> 0 Then
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020") & Err.Description, vbCritical
    End If
End Sub

' ================================================================
' helpers
' ================================================================
Private Function EnsureAuditSheet(ByVal sheetName As String, ByVal subtitle As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = thisWorkbook.Worksheets.Add(After:=thisWorkbook.Worksheets(thisWorkbook.Worksheets.count))
        ws.name = sheetName
    End If
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone
    ws.Cells.Font.ColorIndex = xlAutomatic
    ws.Cells.Font.Bold = False
    ' отчёт о проверке
    ' audit report
    ws.Range("A1").Value = Ru("043E 0442 0447 0451 0442 0020 043E 0020 043F 0440 043E 0432 0435 0440 043A 0435") & " " & subtitle
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ' дата:
    ' date:
    ws.Range("A2").Value = Ru("0434 0430 0442 0430 003A 0020") & Now
    ' серьёзность
    ws.Range("A4").Value = Ru("0441 0435 0440 044C 0451 0437 043D 043E 0441 0442 044C")
    ' таблица
    ws.Range("B4").Value = Ru("0442 0430 0431 043B 0438 0446 0430")
    ' строка
    ws.Range("C4").Value = Ru("0441 0442 0440 043E 043A 0430")
    ' поле
    ws.Range("D4").Value = Ru("043F 043E 043B 0435")
    ' было
    ws.Range("E4").Value = Ru("0431 044B 043B 043E")
    ' при сравнении
    ws.Range("F4").Value = Ru("043F 0440 0438 0020 0441 0440 0430 0432 043D 0435 043D 0438 0438")
    ' замечание
    ws.Range("G4").Value = Ru("0437 0430 043C 0435 0447 0430 043D 0438 0435")
    ws.Range("A4:G4").Font.Bold = True
    Set EnsureAuditSheet = ws
End Function

Private Sub WriteAuditRow(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal severity As String, ByVal tableName As String, _
                          ByVal iRow As Long, ByVal fieldName As String, _
                          ByVal wasVal As Variant, ByVal cmpVal As Variant, _
                          ByVal note As String, ByVal isError As Boolean)
    ws.Cells(r, 1).Value = severity
    ws.Cells(r, 2).Value = tableName
    If iRow > 0 Then ws.Cells(r, 3).Value = iRow Else ws.Cells(r, 3).Value = "-"
    ws.Cells(r, 4).Value = fieldName
    ws.Cells(r, 5).Value = wasVal
    ws.Cells(r, 6).Value = cmpVal
    ws.Cells(r, 7).Value = note
    If isError Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(255, 199, 206)
        auditErrCount = auditErrCount + 1
    Else
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(255, 235, 156)
        auditWarnCount = auditWarnCount + 1
    End If
End Sub

Private Function WriteWhitespaceRow(ByVal ws As Worksheet, ByVal r As Long, _
                                    ByVal tableName As String, ByVal iRow As Long, _
                                    ByVal fieldName As String, ByVal raw As String) As Long
    Dim norm As String
    Dim note As String
    norm = Module_ValidationLists.NormalizeFilterText(raw)
    If raw = norm Then
        WriteWhitespaceRow = 0
        Exit Function
    End If
    note = WhitespaceNote(raw, norm)
    Call WriteAuditRow(ws, r, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), _
                       tableName, iRow, fieldName, raw, norm, note, False)
    WriteWhitespaceRow = 1
End Function

Private Function WhitespaceNote(ByVal raw As String, ByVal norm As String) As String
    Dim parts As String
    If InStr(raw, ChrW(160)) > 0 Then parts = parts & Ru("043D 0435 0440 0430 0437 0440 044B 0432 043D 044B 0439 0020 043F 0440 043E 0431 0435 043B") & "; "
    If InStr(raw, vbTab) > 0 Then parts = parts & Ru("0442 0430 0431 0443 043B 044F 0446 0438 044F") & "; "
    If InStr(raw, vbCr) > 0 Or InStr(raw, vbLf) > 0 Then parts = parts & Ru("043F 0435 0440 0435 0432 043E 0434 0020 0441 0442 0440 043E 043A 0438") & "; "
    If Len(raw) > 0 And (Left$(raw, 1) = " " Or Right$(raw, 1) = " ") Then
        parts = parts & Ru("043F 0440 043E 0431 0435 043B 044B 0020 043F 043E 0020 043A 0440 0430 044F 043C") & "; "
    End If
    If InStr(raw, "  ") > 0 Then parts = parts & Ru("0434 0432 043E 0439 043D 044B 0435 0020 043F 0440 043E 0431 0435 043B 044B") & "; "
    If InStr(raw, ChrW(1105)) > 0 Or InStr(raw, ChrW(1025)) > 0 Then parts = parts & "yo->e; "
    If Len(parts) = 0 Then parts = Ru("0437 0430 043C 0435 043D 0430 0020 0441 0438 043C 0432 043E 043B 043E 0432") & "; "
    WhitespaceNote = parts & Ru("044F 0447 0435 0439 043A 0438 0020 043D 0435 0020 043C 0435 043D 044F 044E 0442 0441 044F")
End Function

Private Function AuditMountingPairs(ByVal wsReport As Worksheet, ByVal startRow As Long, _
                                    ByVal tblMount As ListObject) As Long
    Dim written As Long
    Dim mVals As Variant, iVals As Variant
    Dim nRow As Long
    Dim i As Long
    Dim mountRaw As String, instRaw As String
    Dim r As Long

    mVals = As2DAudit(tblMount.ListColumns(4).DataBodyRange.Value)
    iVals = As2DAudit(tblMount.ListColumns(3).DataBodyRange.Value)
    nRow = UBound(mVals, 1)
    r = startRow
    written = 0

    For i = 1 To nRow
        mountRaw = CellTextAudit(mVals, i, 1)
        instRaw = CellTextAudit(iVals, i, 1)
        If Len(Trim$(mountRaw)) = 0 Or Len(Trim$(instRaw)) = 0 Then
            Call WriteAuditRow(wsReport, r, Ru("043E 0448 0438 0431 043A 0430"), TABLE_MOUNTING, i, _
                               "typeMountingAG / typeInstallationAG", mountRaw, instRaw, _
                               Ru("043D 0435 0442 0020 043F 0430 0440 044B 0020 043A 043E 043B 043E 043D 043E 043A 0020 0034 0020 0438 0020 0033"), True)
            r = r + 1
            written = written + 1
        Else
            written = written + WriteWhitespaceRow(wsReport, r, TABLE_MOUNTING, i, Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430"), mountRaw)
            r = startRow + written
            written = written + WriteWhitespaceRow(wsReport, r, TABLE_MOUNTING, i, Ru("0441 043F 043E 0441 043E 0431 0020 043C 043E 043D 0442 0430 0436 0430"), instRaw)
            r = startRow + written
        End If
    Next i
    AuditMountingPairs = written
End Function

Private Function AuditNamedRangeRows(ByVal wsReport As Worksheet, ByVal r As Long, _
                                     ByVal nm As String, ByVal nTbl As Long) As Long
    Dim rng As Range
    Dim nName As Long
    On Error Resume Next
    Set rng = thisWorkbook.names(nm).RefersToRange
    Err.Clear
    On Error GoTo 0
    If rng Is Nothing Then
        Call WriteAuditRow(wsReport, r, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), nm, 0, _
                           nm, "-", CStr(nTbl), _
                           Ru("0438 043C 044F 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E 002E 0020 0432 044B 043F 043E 043B 043D 0438 0442 0435 0020") & "CreateAGColumnNames", False)
        AuditNamedRangeRows = 1
        Exit Function
    End If
    nName = rng.rows.count
    If nName < nTbl Then
        Call WriteAuditRow(wsReport, r, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435"), nm, 0, _
                           nm, CStr(nName), CStr(nTbl), _
                           Ru("0441 0442 0440 043E 043A 0020 0432 0020 0438 043C 0435 043D 0438 0020 043C 0435 043D 044C 0448 0435 002C 0020 0447 0435 043C 0020 0432 0020 0442 0430 0431 043B 0438 0446 0435 002E 0020 0432 044B 043F 043E 043B 043D 0438 0442 0435 0020") & "CreateAGColumnNames", False)
        AuditNamedRangeRows = 1
    Else
        AuditNamedRangeRows = 0
    End If
End Function

Private Sub FinalizeAuditSheet(ByVal wsReport As Worksheet, ByVal row As Long, ByVal nChecked As Long)
    If row <= 5 Then
        wsReport.Cells(5, 1).Value = Ru("043F 0440 043E 0431 043B 0435 043C 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 043E")
        wsReport.Range("A5:G5").Interior.Color = RGB(198, 239, 206)
        row = 5
    End If
    ' проверено строк:
    ' rows checked:
    wsReport.Range("A3").Value = Ru("043F 0440 043E 0432 0435 0440 0435 043D 043E 0020 0441 0442 0440 043E 043A 003A 0020") & nChecked & _
                                 "  |  " & Ru("043E 0448 0438 0431 043E 043A 003A 0020") & auditErrCount & _
                                 "  |  " & Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0439 003A 0020") & auditWarnCount
    wsReport.Columns("A:G").AutoFit
    wsReport.Range("A4:G" & row).Borders.LineStyle = xlContinuous
End Sub

Private Sub ShowAuditDoneMsg(ByVal sheetName As String)
    Dim msg As String
    msg = Ru("043F 0440 043E 0432 0435 0440 043A 0430 0020 0437 0430 0432 0435 0440 0448 0435 043D 0430 002E") & vbCrLf & _
          Ru("043E 0448 0438 0431 043E 043A 003A 0020") & auditErrCount & _
          Ru("002C 0020 043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0439 003A 0020") & auditWarnCount & vbCrLf & _
          Ru("043B 0438 0441 0442 003A 0020") & sheetName
    If auditErrCount > 0 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

Private Function GetListObjectOnSheet(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set GetListObjectOnSheet = ws.ListObjects(tableName)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindTableMountingAudit() As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(TABLE_MOUNTING)
    If Not tbl Is Nothing Then
        Set FindTableMountingAudit = tbl
        On Error GoTo 0
        Exit Function
    End If
    For Each ws In thisWorkbook.Worksheets
        Set tbl = Nothing
        Set tbl = ws.ListObjects(TABLE_MOUNTING)
        If Not tbl Is Nothing Then
            Set FindTableMountingAudit = tbl
            On Error GoTo 0
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Private Function LoadMountTypeSet(ByVal tblMount As ListObject) As Collection
    Dim col As Collection
    Dim mVals As Variant
    Dim i As Long
    Dim k As String
    Set col = New Collection
    If tblMount Is Nothing Then
        Set LoadMountTypeSet = col
        Exit Function
    End If
    If tblMount.DataBodyRange Is Nothing Then
        Set LoadMountTypeSet = col
        Exit Function
    End If
    mVals = As2DAudit(tblMount.ListColumns(4).DataBodyRange.Value)
    For i = 1 To UBound(mVals, 1)
        k = Module_ValidationLists.NormalizeFilterText(CellTextAudit(mVals, i, 1))
        If Len(k) > 0 Then
            On Error Resume Next
            col.Add True, k
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    Set LoadMountTypeSet = col
End Function

Private Function CollectionLongValue(ByVal col As Collection, ByVal k As String) As Long
    Dim v As Variant
    If col Is Nothing Then Exit Function
    If Len(k) = 0 Then Exit Function
    On Error Resume Next
    v = col(k)
    If Err.Number <> 0 Then
        CollectionLongValue = 0
        Err.Clear
    Else
        CollectionLongValue = CLng(v)
    End If
    On Error GoTo 0
End Function

Private Function KeyExists(ByVal col As Collection, ByVal k As String) As Boolean
    Dim dummy As Variant
    If col Is Nothing Then Exit Function
    If Len(k) = 0 Then Exit Function
    On Error Resume Next
    dummy = col(k)
    KeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ListColIndexByHeader(ByVal tbl As ListObject, ByVal headerName As String) As Long
    Dim c As Long
    Dim h As String
    For c = 1 To tbl.ListColumns.count
        h = Trim$(CStr(tbl.ListColumns(c).name))
        If StrComp(h, headerName, vbTextCompare) = 0 Then
            ListColIndexByHeader = c
            Exit Function
        End If
    Next c
End Function

Private Function ListColIndexByHeaderPart(ByVal tbl As ListObject, ByVal part As String) As Long
    Dim c As Long
    Dim h As String
    For c = 1 To tbl.ListColumns.count
        h = Trim$(CStr(tbl.ListColumns(c).name))
        If InStr(1, h, part, vbTextCompare) > 0 Then
            ListColIndexByHeaderPart = c
            Exit Function
        End If
    Next c
End Function

Private Function As2DAudit(ByVal vals As Variant) As Variant
    If IsArray(vals) Then
        As2DAudit = vals
    Else
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = vals
        As2DAudit = arr
    End If
End Function

Private Function CellTextAudit(ByVal arr As Variant, ByVal iRow As Long, ByVal iCol As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = arr(iRow, iCol)
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then
        CellTextAudit = ""
    Else
        CellTextAudit = CStr(v)
    End If
End Function

' ------------------------------------------------------------------
' clean TableAG (replace invalid characters with "_")
' not used by the ListAG button
' ------------------------------------------------------------------
Sub CleanTableAG()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    
    ' list of columns to clean (exact names)
    Set targetColumns = New Collection
    ' материал анода
    ' anode material
    targetColumns.Add Ru("043C 0430 0442 0435 0440 0438 0430 043B 0020 0430 043D 043E 0434 0430")
    ' тип монтажа
    ' mounting type
    targetColumns.Add Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430")
    ' комплектация
    ' delivery set
    targetColumns.Add Ru("043A 043E 043C 043F 043B 0435 043A 0442 0430 0446 0438 044F")
    targetColumns.Add "model"
    
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
    If ws Is Nothing Then
        MsgBox "sheet '" & SHEET_LIST_AG & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(TABLE_AG)
    If tbl Is Nothing Then
        MsgBox "table '" & TABLE_AG & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' iterate over all table columns
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        
        ' check if this column should be cleaned
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            ' clean all cells in the column (except header)
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    cell.Value = CleanStringForTable(CStr(cell.Value))
                End If
            Next cell
        End If
    Next col
    
    Call Module_ValidationLists.InvalidateAGListCache
    MsgBox "table " & TABLE_AG & " cleaned of invalid characters." & vbCrLf & _
           "all invalid characters replaced with '_'.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------
' clean TableMounting (replace invalid characters with "_")
' ------------------------------------------------------------------
Sub CleanTableMounting()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    Dim found As Boolean
    
    ' list of columns to clean
    Set targetColumns = New Collection
    ' тип монтажа
    ' mounting type
    targetColumns.Add Ru("0442 0438 043F 0020 043C 043E 043D 0442 0430 0436 0430")
    ' способ монтажа
    ' installation method
    targetColumns.Add Ru("0441 043F 043E 0441 043E 0431 0020 043C 043E 043D 0442 0430 0436 0430")
    
    ' try to find the table on sheet "ListAG"
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets("ListAG")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' if sheet "ListAG" is missing, try SHEET_LIST_AG
        Set ws = thisWorkbook.Worksheets(SHEET_LIST_AG)
        If ws Is Nothing Then
            MsgBox "sheet for TableMounting not found!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' check for "TableMounting" existence
    On Error Resume Next
    Set tbl = ws.ListObjects("TableMounting")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "table 'TableMounting' not found on sheet " & ws.name, vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' iterate over all table columns
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    cell.Value = CleanStringForTable(CStr(cell.Value))
                End If
            Next cell
        End If
    Next col
    
    Call Module_ValidationLists.InvalidateAGListCache
    MsgBox "table TableMounting cleaned of invalid characters." & vbCrLf & _
           "all invalid characters replaced with '_'.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ------------------------------------------------------------------
' clean TableCP (from control characters)
' ------------------------------------------------------------------
Sub CleanTableCP()
    On Error GoTo CleanExit
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim headerVal As String
    Dim targetColumns As Collection
    Dim colName As Variant
    
    Set targetColumns = New Collection
    targetColumns.Add "typeCP"
    ' модель
    ' model
    targetColumns.Add Ru("043C 043E 0434 0435 043B 044C")
    
    Set ws = thisWorkbook.Worksheets(SHEET_LIST_CP)
    If ws Is Nothing Then
        MsgBox "sheet '" & SHEET_LIST_CP & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(TABLE_CP)
    If tbl Is Nothing Then
        MsgBox "table '" & TABLE_CP & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each col In tbl.ListColumns
        headerVal = Trim(CStr(col.Range.Cells(1, 1).Value))
        Dim shouldClean As Boolean
        shouldClean = False
        For Each colName In targetColumns
            If StrComp(headerVal, colName, vbTextCompare) = 0 Then
                shouldClean = True
                Exit For
            End If
        Next colName
        
        If shouldClean Then
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
                    Dim cleaned As String
                    cleaned = CleanStringForTable(CStr(cell.Value))
                    cleaned = Replace(cleaned, vbLf, "")
                    cleaned = Replace(cleaned, vbCr, "")
                    cleaned = Replace(cleaned, vbTab, " ")
                    cleaned = Application.WorksheetFunction.Trim(cleaned)
                    cell.Value = cleaned
                End If
            Next cell
        End If
    Next col
    
    MsgBox "table " & TABLE_CP & " cleaned of invalid characters.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
