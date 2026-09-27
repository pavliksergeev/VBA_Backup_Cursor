Attribute VB_Name = "Module_SoilAvg"

' ================================================================
' Survey editor for equivalent soil resistivity:
' rho_avg = (sum L)^2 / (sum (L / Sqr(rho)))^2
' No UserForm / MSForms: a Mac-saved .xlsm must open on Windows as-is.
' UI is sheet _SoilAvg (cells + Form Control buttons). Captions via Ru()
' at Show time. Storage table TableSoilAvg is on the same sheet, cols AA:AG.
' Write target is a named 1xN range + index: soilResistivityAvg (Pipe),
' resistivitySoilAG (Anod), or resistivity_i_layerDeepAG + factorSoilHeterogeneity
' (Anod row 53/54, layer mode). Survives SyncPipe / SyncAnodColumns.
' Mac paste: this .bas only (plus Module_Research / Module_Constants).
' ================================================================
Option Explicit

Public Const SOILAVG_TARGET_AUTO As Long = 0
Public Const SOILAVG_TARGET_PIPE As Long = 1
Public Const SOILAVG_TARGET_ANOD As Long = 2
Public Const SOILAVG_TARGET_LAYER As Long = 3

Private Const SOILAVG_DEFAULT_ROWS As Long = 5
Private Const UI_AVG_ROW As Long = 3
Private Const UI_BTN_ROW As Long = 5
Private Const UI_HEADER_ROW As Long = 6
Private Const UI_FIRST_ROW As Long = 7
Private Const UI_COLS As Long = 6
Private Const META_COL As Long = 7
Private Const STORE_COL As Long = 27
Private Const STORE_COLS As Long = 7
Private Const NM_PIPE As String = "soilResistivityAvg"
Private Const NM_ANOD As String = "resistivitySoilAG"
Private Const NM_LAYER As String = "resistivity_i_layerDeepAG"
Private Const NM_HET As String = "factorSoilHeterogeneity"
Private Const NM_WALL As String = "pipeWallThickness"
Private Const NM_LEN As String = "pipeLengthSegment"

Private mTarget As Range
Private mKey As String
Private mRangeName As String
Private mIdx As Long
Private mRowCount As Long
Private mPrevSheet As Worksheet
Private mFastDepth As Long
Private mOldCalc As XlCalculation
Private mOldUpd As Boolean
Private mOldEv As Boolean
Private mRhoMin As Double
Private mHetK As Double

' ================================================================
' public entry
' ================================================================
Public Sub ShowSoilAvgForm(Optional ByVal targetMode As Long = SOILAVG_TARGET_AUTO)
    Dim ws As Worksheet
    Dim wsPrev As Worksheet

    On Error Resume Next
    Set wsPrev = ActiveSheet
    On Error GoTo 0
    If Not wsPrev Is Nothing Then
        If StrComp(wsPrev.Name, SHEET_SOIL_AVG, vbTextCompare) = 0 Then
            Set wsPrev = Nothing
        End If
    End If

    Set mTarget = ResolveSoilAvgTarget(targetMode)
    If mTarget Is Nothing Then
        ' целевая ячейка не найдена.
        MsgBox Ru("0426 0435 043B 0435 0432 0430 044F 0020 044F 0447 0435 0439 043A 0430 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 002E"), vbExclamation
        Exit Sub
    End If
    Set mPrevSheet = wsPrev

    Set ws = EnsureSoilAvgSheet()
    If ws Is Nothing Then
        ' лист '  ' не найден!
        MsgBox Ru("043B 0438 0441 0442 0020 0027") & SHEET_SOIL_AVG & Ru("0027 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0021"), vbCritical
        Exit Sub
    End If

    Call SoilAvgUnprotect(ws)
    Call SoilAvgFastBegin
    On Error GoTo ShowFail
    Call PaintSoilAvgUi(ws)
    Call SaveSessionToSheet(ws)
    Call LoadKeyIntoGrid(ws)
    Call PlaceSoilAvgButtons(ws)
    Call SoilAvg_OnCalcSilent
    Call SoilAvgFastEnd

    ' Do not Protect _SoilAvg: ListRows.Add/Delete raise 1004 on a protected
    ' sheet (Windows and Mac). UserInterfaceOnly does not allow table edits.

    ws.Visible = xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("A" & CStr(UI_FIRST_ROW)).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    Exit Sub
ShowFail:
    Call SoilAvgFastEnd
End Sub

Public Sub ShowSoilAvgFormAnod()
    Call ShowSoilAvgForm(SOILAVG_TARGET_ANOD)
End Sub

Public Sub ShowSoilAvgFormLayer()
    Call ShowSoilAvgForm(SOILAVG_TARGET_LAYER)
End Sub

' ================================================================
' Form Control OnAction (ASCII names, captions set with Ru() each Show)
' ================================================================
Public Sub SoilAvg_OnCalc()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim wallAvg As Double
    Dim nWall As Long
    Dim errMsg As String
    Dim ws As Worksheet

    If Not RestoreSessionFromSheet(ws) Then Exit Sub
    Call SoilAvgFastBegin
    If Not TryCalcFromGrid(ws, avg, distSum, nRes, errMsg, wallAvg, nWall) Then
        Call SoilAvgFastEnd
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    If nRes < 1 Then
        Call SoilAvgFastEnd
        MsgBox Ru("041D 0435 0442 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 044F 0020 0434 043B 044F 0020") & _
               Ru("0441 0440 0435 0434 043D 0435 0433 043E 002E"), vbExclamation
        Exit Sub
    End If
    Call ShowCalcLabels(ws, avg, distSum, nRes, wallAvg, nWall)
    Call SoilAvgFastEnd
End Sub

Public Sub SoilAvg_OnApply()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim wallAvg As Double
    Dim nWall As Long
    Dim errMsg As String
    Dim ws As Worksheet
    Dim rngWall As Range
    Dim rngLen As Range

    If Not RestoreSessionFromSheet(ws) Then Exit Sub
    Set mTarget = LiveTargetCell()
    If mTarget Is Nothing Then
        MsgBox Ru("0426 0435 043B 0435 0432 0430 044F 0020 044F 0447 0435 0439 043A 0430 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 002E"), vbExclamation
        Exit Sub
    End If
    Call SoilAvgFastBegin
    If Not TryCalcFromGrid(ws, avg, distSum, nRes, errMsg, wallAvg, nWall) Then
        Call SoilAvgFastEnd
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    If nRes < 1 Then
        Call SoilAvgFastEnd
        MsgBox Ru("041D 0435 0442 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 044F 0020 0434 043B 044F 0020") & _
               Ru("0441 0440 0435 0434 043D 0435 0433 043E 002E"), vbExclamation
        Exit Sub
    End If

    If IsLayerMode() Then
        Call WriteSoilAvgTarget(mTarget, mRhoMin)
        Set rngLen = Named1xN(NM_HET, SHEET_ANOD)
        If Not rngLen Is Nothing Then
            Call WriteSoilAvgTarget(rngLen.Worksheet.Cells(rngLen.Row, mTarget.Column), mHetK)
        End If
    Else
        Call WriteSoilAvgTarget(mTarget, avg)
        If StrComp(mRangeName, NM_PIPE, vbTextCompare) = 0 Then
            If nWall >= 1 Then
                Set rngWall = Named1xN(NM_WALL, SHEET_PIPE)
                If Not rngWall Is Nothing Then
                    Call WriteSoilAvgTarget(rngWall.Worksheet.Cells(rngWall.Row, mTarget.Column), wallAvg)
                End If
            End If
            Set rngLen = Named1xN(NM_LEN, SHEET_PIPE)
            If Not rngLen Is Nothing Then
                Call WriteSoilAvgTarget(rngLen.Worksheet.Cells(rngLen.Row, mTarget.Column), distSum)
            End If
        End If
    End If
    Call PersistGridRows(ws)
    Call ShowCalcLabels(ws, avg, distSum, nRes, wallAvg, nWall)
    If StrComp(mRangeName, NM_PIPE, vbTextCompare) = 0 Then
        Call Module_calcPipe.btnPipeCalculate
    ElseIf IsLayerMode() Then
        Call btnAnodFullCalc
    End If
    If DEBUG_MODE Then
        Debug.Print "SoilAvg written to " & mRangeName & " (" & CStr(mIdx) & ")"
        If IsLayerMode() Then
            Debug.Print "resistivity_i_layerDeepAG (" & CStr(mIdx) & ") = " & CStr(mRhoMin)
            Debug.Print "factorSoilHeterogeneity (" & CStr(mIdx) & ") = " & CStr(mHetK)
        Else
            If nWall >= 1 Then Debug.Print "pipeWallThickness (" & CStr(mIdx) & ") = " & CStr(wallAvg)
            Debug.Print "pipeLengthSegment (" & CStr(mIdx) & ") = " & CStr(distSum)
        End If
    End If
    Call SoilAvgFastEnd
    Call GoToSoilAvgTargetCell
End Sub

Public Sub SoilAvg_OnSave()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim wallAvg As Double
    Dim nWall As Long
    Dim errMsg As String
    Dim ws As Worksheet

    If Not RestoreSessionFromSheet(ws) Then Exit Sub
    Call SoilAvgFastBegin
    If Not TryCalcFromGrid(ws, avg, distSum, nRes, errMsg, wallAvg, nWall, True) Then
        Call SoilAvgFastEnd
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    Call PersistGridRows(ws)
    Call SoilAvgFastEnd
    MsgBox Ru("0414 0430 043D 043D 044B 0435 0020 0438 0437 044B 0441 043A 0430 043D 0438 0439 0020 0441 043E 0445 0440 0430 043D 0435 043D 044B 002E"), vbInformation
End Sub

Public Sub SoilAvg_OnClose()
    Dim ws As Worksheet
    Dim wsBack As Worksheet

    Set ws = Nothing
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_SOIL_AVG)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Call RestoreSessionFromSheet(ws)
        ws.Visible = xlSheetVeryHidden
    End If

    Set wsBack = mPrevSheet
    If wsBack Is Nothing Then
        On Error Resume Next
        If IsLayerMode() Or StrComp(mRangeName, NM_ANOD, vbTextCompare) = 0 Then
            Set wsBack = thisWorkbook.Worksheets(SHEET_ANOD)
        Else
            Set wsBack = thisWorkbook.Worksheets(SHEET_PIPE)
        End If
        On Error GoTo 0
    End If
    If Not wsBack Is Nothing Then
        On Error Resume Next
        If wsBack.Visible = xlSheetVisible Then wsBack.Activate
        On Error GoTo 0
    End If
End Sub

' Hide the editor and select the named-range cell under the calling button.
Private Sub GoToSoilAvgTargetCell()
    Dim rng As Range
    Dim wsUi As Worksheet
    Dim wsDest As Worksheet

    Set rng = LiveTargetCell()
    If rng Is Nothing Then
        Call SoilAvg_OnClose
        Exit Sub
    End If
    Set wsDest = rng.Worksheet

    On Error Resume Next
    Set wsUi = thisWorkbook.Worksheets(SHEET_SOIL_AVG)
    On Error GoTo 0

    On Error Resume Next
    If wsDest.Visible <> xlSheetVisible Then wsDest.Visible = xlSheetVisible
    wsDest.Activate
    If Not wsUi Is Nothing Then
        If StrComp(ActiveSheet.Name, SHEET_SOIL_AVG, vbTextCompare) <> 0 Then
            wsUi.Visible = xlSheetVeryHidden
        End If
    End If
    Application.Goto rng, True
    On Error GoTo 0
End Sub

' ================================================================
' target: named 1xN range + 1-based index (survives column sync)
' ================================================================
Private Function ResolveSoilAvgTarget(ByVal targetMode As Long) As Range
    Dim rngPipe As Range
    Dim rngAnod As Range
    Dim rngLayer As Range
    Dim callerCell As Range
    Dim mode As Long
    Dim parent As Range

    Set rngPipe = Named1xN(NM_PIPE, SHEET_PIPE)
    Set rngAnod = Named1xN(NM_ANOD, SHEET_ANOD)
    Set rngLayer = Named1xN(NM_LAYER, SHEET_ANOD)

    mode = targetMode
    If Not IsSoilAvgUiSheet(ActiveSheet) Then
        Set callerCell = CellFromCaller()
    End If

    If mode = SOILAVG_TARGET_AUTO Then
        If Not callerCell Is Nothing Then
            If Not rngLayer Is Nothing Then
                If Not Intersect(callerCell, rngLayer) Is Nothing Then mode = SOILAVG_TARGET_LAYER
            End If
            If mode = SOILAVG_TARGET_AUTO And Not rngAnod Is Nothing Then
                If Not Intersect(callerCell, rngAnod) Is Nothing Then mode = SOILAVG_TARGET_ANOD
            End If
            If mode = SOILAVG_TARGET_AUTO And Not rngPipe Is Nothing Then
                If Not Intersect(callerCell, rngPipe) Is Nothing Then mode = SOILAVG_TARGET_PIPE
            End If
        End If
        If mode = SOILAVG_TARGET_AUTO Then
            If StrComp(ActiveSheet.Name, SHEET_ANOD, vbTextCompare) = 0 Then
                mode = SOILAVG_TARGET_ANOD
            Else
                mode = SOILAVG_TARGET_PIPE
            End If
        End If
    End If

    If mode = SOILAVG_TARGET_LAYER Then
        mRangeName = NM_LAYER
        Set parent = rngLayer
    ElseIf mode = SOILAVG_TARGET_ANOD Then
        mRangeName = NM_ANOD
        Set parent = rngAnod
    Else
        mRangeName = NM_PIPE
        Set parent = rngPipe
    End If
    mIdx = IdxIn1xN(parent, callerCell)
    mKey = KeyFromNameIdx(mRangeName, mIdx)
    Set ResolveSoilAvgTarget = CellByNameIdx(mRangeName, mIdx)
End Function

Private Function Named1xN(ByVal rangeName As String, ByVal sheetName As String) As Range
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Function
    Set Named1xN = ws.Range(rangeName)
    If Named1xN Is Nothing Then
        Set Named1xN = thisWorkbook.names(rangeName).RefersToRange
    End If
    On Error GoTo 0
End Function

Private Function CellFromCaller() As Range
    Dim v As Variant
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn As Button
    Dim i As Long
    Dim act As String
    Dim hit As Range
    Dim nHit As Long

    On Error Resume Next
    v = Application.Caller
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Function
    If VarType(v) = vbString Then
        Set shp = ws.Shapes(CStr(v))
        If Not shp Is Nothing Then Set CellFromCaller = ShapeAnchorCell(ws, shp)
        If CellFromCaller Is Nothing Then
            Set btn = ws.Buttons(CStr(v))
            If Not btn Is Nothing Then Set CellFromCaller = btn.TopLeftCell
        End If
    End If
    If CellFromCaller Is Nothing Then
        nHit = 0
        For i = 1 To ws.Shapes.Count
            act = ""
            Set shp = ws.Shapes(i)
            act = CStr(shp.OnAction)
            If InStr(1, act, "button112_click", vbTextCompare) > 0 Then
                nHit = nHit + 1
                Set hit = ShapeAnchorCell(ws, shp)
            End If
        Next i
        If nHit = 1 Then Set CellFromCaller = hit
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Cell that contains the button body (center), not the left border.
Private Function ShapeAnchorCell(ByVal ws As Worksheet, ByVal shp As Shape) As Range
    Dim cell As Range
    Dim cx As Double
    Dim n As Long
    If shp Is Nothing Then Exit Function
    On Error Resume Next
    Set cell = shp.TopLeftCell
    cx = shp.Left + shp.Width * 0.5
    n = 0
    Do While Not cell Is Nothing And n < 8
        If cx < cell.Left + cell.Width - 0.5 Then Exit Do
        Set cell = ws.Cells(cell.Row, cell.Column + 1)
        n = n + 1
    Loop
    Set ShapeAnchorCell = cell
    Err.Clear
    On Error GoTo 0
End Function

Private Function AnodLayerBtnLeft(ByVal cell As Range) As Double
    Dim dx As Double
    If cell Is Nothing Then Exit Function
    dx = 4
    If cell.Width * 0.12 > dx Then dx = cell.Width * 0.12
    If dx > 10 Then dx = 10
    AnodLayerBtnLeft = cell.Left + dx
End Function

Private Function IdxIn1xN(ByVal rng As Range, ByVal hint As Range) As Long
    Dim idx As Long
    Dim colHint As Long

    idx = 1
    If rng Is Nothing Then
        IdxIn1xN = 1
        Exit Function
    End If
    If Not hint Is Nothing Then
        If hint.Column >= rng.Column And hint.Column <= rng.Column + rng.Columns.Count - 1 Then
            idx = hint.Column - rng.Column + 1
            GoTo ClampIdx
        End If
    End If

    colHint = 0
    On Error Resume Next
    If StrComp(ActiveSheet.Name, rng.Worksheet.Name, vbTextCompare) = 0 Then
        colHint = ActiveCell.Column
    End If
    On Error GoTo 0

    If colHint >= rng.Column And colHint <= rng.Column + rng.Columns.Count - 1 Then
        idx = colHint - rng.Column + 1
    Else
        idx = 1
    End If
ClampIdx:
    If idx < 1 Then idx = 1
    If idx > rng.Columns.Count Then idx = rng.Columns.Count
    IdxIn1xN = idx
End Function

Private Function SheetForRangeName(ByVal rangeName As String) As String
    If StrComp(rangeName, NM_ANOD, vbTextCompare) = 0 Or _
       StrComp(rangeName, NM_LAYER, vbTextCompare) = 0 Or _
       StrComp(rangeName, NM_HET, vbTextCompare) = 0 Then
        SheetForRangeName = SHEET_ANOD
    Else
        SheetForRangeName = SHEET_PIPE
    End If
End Function

Private Function IsLayerMode() As Boolean
    IsLayerMode = (StrComp(mRangeName, NM_LAYER, vbTextCompare) = 0)
End Function

Private Function KeyFromNameIdx(ByVal rangeName As String, ByVal idx As Long) As String
    If StrComp(rangeName, NM_LAYER, vbTextCompare) = 0 Then
        KeyFromNameIdx = "LAYER|" & CStr(idx)
    ElseIf StrComp(rangeName, NM_ANOD, vbTextCompare) = 0 Then
        KeyFromNameIdx = "ANOD|" & CStr(idx)
    Else
        KeyFromNameIdx = "PIPE|" & CStr(idx)
    End If
End Function

Private Sub ParseSoilAvgKey(ByVal key As String, ByRef rangeName As String, ByRef idx As Long)
    Dim p As Long
    Dim prefix As String
    Dim rest As String

    rangeName = NM_PIPE
    idx = 1
    p = InStr(1, key, "|")
    If p < 1 Then Exit Sub
    prefix = UCase$(Trim$(Left$(key, p - 1)))
    rest = Trim$(Mid$(key, p + 1))
    If prefix = "LAYER" Then
        rangeName = NM_LAYER
    ElseIf prefix = "ANOD" Then
        rangeName = NM_ANOD
    End If
    If IsNumeric(rest) Then
        idx = CLng(Val(rest))
        If idx < 1 Then idx = 1
    End If
End Sub

Private Function CellByNameIdx(ByVal rangeName As String, ByVal idx As Long) As Range
    Dim parent As Range
    Set parent = Named1xN(rangeName, SheetForRangeName(rangeName))
    If parent Is Nothing Then Exit Function
    If idx < 1 Then idx = 1
    If idx > parent.Columns.Count Then idx = parent.Columns.Count
    Set CellByNameIdx = parent.Cells(1, idx)
End Function

Private Function LiveTargetCell() As Range
    Dim parent As Range

    Call ParseSoilAvgKey(mKey, mRangeName, mIdx)
    If Len(mRangeName) = 0 Then mRangeName = NM_PIPE
    If mIdx < 1 Then mIdx = 1
    Set parent = Named1xN(mRangeName, SheetForRangeName(mRangeName))
    If parent Is Nothing Then Exit Function
    If mIdx > parent.Columns.Count Then mIdx = parent.Columns.Count
    mKey = KeyFromNameIdx(mRangeName, mIdx)
    Set LiveTargetCell = parent.Cells(1, mIdx)
End Function

Private Function SoilAvgTargetCaption() As String
    If IsLayerMode() Then
        SoilAvgTargetCaption = Ru("0426 0435 043B 044C 003A 0020") & NM_LAYER & "; " & NM_HET & " (" & CStr(mIdx) & ")"
    Else
        SoilAvgTargetCaption = Ru("0426 0435 043B 044C 003A 0020") & mRangeName & " (" & CStr(mIdx) & ")"
    End If
End Function

' ================================================================
' UI sheet
' ================================================================
Private Sub PaintSoilAvgUi(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    ws.Range("A3:F5").ClearContents
    ws.Range("A45:F47").ClearContents
    Call ApplyRowVisibility(ws)

    ws.Range("A2").Value = SoilAvgTargetCaption()
    If IsLayerMode() Then
        ws.Range("A1").Value = Ru("0423 0434 0435 043B 044C 043D 043E 0435 0020 043C 0438 043D 0438 043C 0430 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0069 002D 0442 043E 0433 043E 0020 0441 043B 043E 044F 0020 0433 0440 0443 043D 0442 0430 002C 0020 0432 0020 043A 043E 0442 043E 0440 043E 043C 0020 0440 0430 0441 043F 043E 043B 0430 0433 0430 0435 0442 0441 044F 0020 0433 043B 0443 0431 0438 043D 043D 044B 0439 0020 0437 0430 0437 0435 043C 043B 0438 0442 0435 043B 044C")
        ws.Cells(UI_HEADER_ROW, 3).Value = Ru("0413 043B 0443 0431 0438 043D 0430 002C 0020 043C")
        ws.Cells(UI_HEADER_ROW, 4).Value = Ru("0422 043E 043B 0449 0438 043D 0430 0020 0441 043B 043E 044F 002C 0020 043C")
        ws.Cells(UI_HEADER_ROW, 6).Value = ""
        ws.Columns(6).Hidden = True
    Else
        ws.Range("A1").Value = Ru("0421 0440 0435 0434 043D 0435 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0433 0440 0443 043D 0442 0430")
        ws.Cells(UI_HEADER_ROW, 3).Value = Ru("0413 0435 043E 043F 043E 0437 0438 0446 0438 044F")
        ws.Cells(UI_HEADER_ROW, 4).Value = Ru("0414 0438 0441 0442 0430 043D 0446 0438 044F 002C 0020 043C")
        ws.Cells(UI_HEADER_ROW, 6).Value = Ru("0422 043E 043B 0449 0438 043D 0430 0020 0441 0442 0435 043D 043A 0438 002C 0020 043C")
        ws.Columns(6).Hidden = False
        ws.Columns(6).ColumnWidth = 18
    End If
    ws.Cells(UI_HEADER_ROW, 1).Value = Ru("0423 0447 0430 0441 0442 043E 043A")
    ws.Cells(UI_HEADER_ROW, 2).Value = Ru("041F 0438 043A 0435 0442")
    ws.Cells(UI_HEADER_ROW, 5).Value = Ru("0421 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 002C 0020 041E 043C 00B7 043C")

    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 22
    If Not IsLayerMode() Then ws.Columns(6).ColumnWidth = 18
    ws.Columns(META_COL).Hidden = True
    ws.Range(ws.Columns(STORE_COL), ws.Columns(STORE_COL + STORE_COLS - 1)).Hidden = True

    ws.Range("A1").Font.Bold = True
    If IsLayerMode() Then
        ws.Range(ws.Cells(UI_HEADER_ROW, 1), ws.Cells(UI_HEADER_ROW, 5)).Font.Bold = True
    Else
        ws.Range(ws.Cells(UI_HEADER_ROW, 1), ws.Cells(UI_HEADER_ROW, UI_COLS)).Font.Bold = True
    End If
End Sub

Private Sub LoadKeyIntoGrid(ByVal ws As Worksheet)
    Dim saved As Collection
    Dim iRow As Long
    Dim nFill As Long
    Dim dump() As Variant
    Dim arr As Variant

    Call ClearGrid(ws)
    Set saved = LoadSoilAvgRows(mKey)
    If saved.Count < 1 Then
        mRowCount = SOILAVG_DEFAULT_ROWS
    Else
        nFill = saved.Count
        mRowCount = nFill
        ReDim dump(1 To nFill, 1 To UI_COLS)
        For iRow = 1 To nFill
            arr = saved(iRow)
            dump(iRow, 1) = arr(0)
            dump(iRow, 2) = arr(1)
            dump(iRow, 3) = arr(2)
            dump(iRow, 4) = arr(3)
            dump(iRow, 5) = arr(4)
            If UBound(arr) >= 5 Then dump(iRow, 6) = arr(5)
        Next iRow
        ws.Range(ws.Cells(UI_FIRST_ROW, 1), ws.Cells(UI_FIRST_ROW + nFill - 1, UI_COLS)).Value = dump
    End If
    Call ApplyRowVisibility(ws)
    Call SaveSessionToSheet(ws)
End Sub

Private Sub ClearGrid(ByVal ws As Worksheet)
    Dim last As Long
    last = LastFilledSheetRow(ws)
    If last < UI_FIRST_ROW Then Exit Sub
    ws.Range(ws.Cells(UI_FIRST_ROW, 1), ws.Cells(last, UI_COLS)).ClearContents
End Sub

Private Sub ApplyRowVisibility(ByVal ws As Worksheet)
    Dim last As Long
    last = LastFilledSheetRow(ws)
    If last < UI_FIRST_ROW Then last = UI_FIRST_ROW
    If last < 47 Then last = 47
    ws.Rows("1:" & CStr(last)).Hidden = False
End Sub

Private Function LastFilledSheetRow(ByVal ws As Worksheet) As Long
    Dim c As Long
    Dim r As Long
    Dim last As Long
    last = 0
    For c = 1 To UI_COLS
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r >= UI_FIRST_ROW And r > last Then last = r
    Next c
    LastFilledSheetRow = last
End Function

Private Function GridRowCount(ByVal ws As Worksheet) As Long
    Dim last As Long
    last = LastFilledSheetRow(ws)
    If last < UI_FIRST_ROW Then
        GridRowCount = 0
    Else
        GridRowCount = last - UI_FIRST_ROW + 1
    End If
End Function

Private Function ReadGridValues(ByVal ws As Worksheet, ByRef nRow As Long) As Variant
    nRow = GridRowCount(ws)
    If nRow < 1 Then Exit Function
    ReadGridValues = ws.Range(ws.Cells(UI_FIRST_ROW, 1), ws.Cells(UI_FIRST_ROW + nRow - 1, UI_COLS)).Value
End Function

Private Function CellToTrim(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    CellToTrim = Trim$(CStr(v))
End Function

Private Function CoerceNumber(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = False
    CoerceNumber = 0
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then
        CoerceNumber = CDbl(v)
        ok = True
        Exit Function
    End If
    CoerceNumber = ParseDotNumber(CellToTrim(v), ok)
End Function

Private Function VariantRowEmpty(ByVal grid As Variant, ByVal iRow As Long) As Boolean
    Dim c As Long
    For c = 1 To UI_COLS
        If Len(CellToTrim(grid(iRow, c))) > 0 Then
            VariantRowEmpty = False
            Exit Function
        End If
    Next c
    VariantRowEmpty = True
End Function

Private Sub SoilAvg_OnCalcSilent()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim wallAvg As Double
    Dim nWall As Long
    Dim errMsg As String
    Dim ws As Worksheet
    If Not RestoreSessionFromSheet(ws) Then Exit Sub
    If TryCalcFromGrid(ws, avg, distSum, nRes, errMsg, wallAvg, nWall, True) Then
        Call ShowCalcLabels(ws, avg, distSum, nRes, wallAvg, nWall)
    End If
End Sub

Private Sub ShowCalcLabels(ByVal ws As Worksheet, ByVal avg As Double, ByVal distSum As Double, _
                           ByVal nRes As Long, ByVal wallAvg As Double, ByVal nWall As Long)
    Dim avgTxt As String
    Dim wallTxt As String
    Dim hetTxt As String
    If nRes < 1 Then
        avgTxt = "-"
        hetTxt = "-"
    Else
        avgTxt = FmtNum(avg)
        hetTxt = FmtNum(mHetK)
    End If
    If IsLayerMode() Then
        If nRes < 1 Then
            ws.Cells(UI_AVG_ROW, 1).Value = Ru("041C 0438 043D 0438 043C 0430 043B 044C 043D 043E 0435 0020 0443 0434 0435 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 003A") & " -"
        Else
            ws.Cells(UI_AVG_ROW, 1).Value = Ru("041C 0438 043D 0438 043C 0430 043B 044C 043D 043E 0435 0020 0443 0434 0435 043B 044C 043D 043E 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 003A") & " " & FmtNum(mRhoMin)
        End If
        ws.Cells(UI_AVG_ROW, 3).Value = Ru("0413 043B 0443 0431 0438 043D 0430 0020 0441 043A 0432 0430 0436 0438 043D 044B 003A") & " " & FmtNum(distSum)
        ws.Cells(UI_AVG_ROW, 5).Value = Ru("041A 043E 044D 0444 0444 0438 0446 0438 0435 043D 0442 003A") & " " & hetTxt
        ws.Cells(UI_AVG_ROW, 6).Value = ""
        Exit Sub
    End If
    If nWall < 1 Then
        wallTxt = "-"
    Else
        wallTxt = FmtNum(wallAvg)
    End If
    ws.Cells(UI_AVG_ROW, 1).Value = Ru("0421 0440 0435 0434 043D 0435 0435 003A") & " " & avgTxt
    ws.Cells(UI_AVG_ROW, 3).Value = Ru("0421 0443 043C 043C 0430 0020 0434 0438 0441 0442 0430 043D 0446 0438 0439 003A") & " " & FmtNum(distSum)
    ' Средняя толщина стенки:
    ws.Cells(UI_AVG_ROW, 6).Value = Ru("0421 0440 0435 0434 043D 044F 044F 0020 0442 043E 043B 0449 0438 043D 0430 0020 0441 0442 0435 043D 043A 0438 003A") & " " & wallTxt
End Sub

Private Function TryCalcFromGrid(ByVal ws As Worksheet, ByRef avg As Double, ByRef distSum As Double, _
                                 ByRef nRes As Long, ByRef errMsg As String, _
                                 ByRef wallAvg As Double, ByRef nWall As Long, _
                                 Optional ByVal skipInvalid As Boolean = False) As Boolean
    Dim iRow As Long
    Dim resOk As Boolean
    Dim distOk As Boolean
    Dim wallOk As Boolean
    Dim pkOk As Boolean
    Dim resVal As Double
    Dim distVal As Double
    Dim wallVal As Double
    Dim pkRaw As String
    Dim soilLenSum As Double
    Dim soilRatioSum As Double
    Dim layerInvSum As Double
    Dim rhoMin As Double
    Dim wallWSum As Double
    Dim wallDistSum As Double
    Dim grid As Variant
    Dim nGrid As Long

    avg = 0
    distSum = 0
    nRes = 0
    wallAvg = 0
    nWall = 0
    errMsg = ""
    soilLenSum = 0
    soilRatioSum = 0
    layerInvSum = 0
    rhoMin = 0
    mRhoMin = 0
    mHetK = 0
    wallWSum = 0
    wallDistSum = 0
    If ws Is Nothing Then Exit Function

    grid = ReadGridValues(ws, nGrid)
    If nGrid < 1 Then
        TryCalcFromGrid = True
        Exit Function
    End If

    For iRow = 1 To nGrid
        If VariantRowEmpty(grid, iRow) Then GoTo SoilAvg_NextRow
        pkRaw = CellToTrim(grid(iRow, 2))
        If Len(pkRaw) > 0 Then
            If Not IsWholeNumber(pkRaw, pkOk) Then
                If Not skipInvalid Then
                    errMsg = Ru("041F 0438 043A 0435 0442 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 0446 0435 043B 044B 043C 0020") & _
                             Ru("0447 0438 0441 043B 043E 043C 002E") & " #" & CStr(iRow)
                    Exit Function
                End If
            End If
        End If

        distVal = CoerceNumber(grid(iRow, 4), distOk)
        If distOk Then
            distSum = distSum + distVal
        ElseIf Len(CellToTrim(grid(iRow, 4))) > 0 Then
            If Not skipInvalid Then
                errMsg = BadNumberMsg(iRow)
                Exit Function
            End If
        End If

        resVal = CoerceNumber(grid(iRow, 5), resOk)
        If resOk Then
            If resVal <= 0 Then
                If Not skipInvalid Then
                    errMsg = BadNumberMsg(iRow)
                    Exit Function
                End If
            ElseIf distOk And distVal > 0 Then
                soilLenSum = soilLenSum + distVal
                soilRatioSum = soilRatioSum + distVal / Sqr(resVal)
                layerInvSum = layerInvSum + distVal / resVal
                nRes = nRes + 1
                If nRes = 1 Then
                    rhoMin = resVal
                ElseIf resVal < rhoMin Then
                    rhoMin = resVal
                End If
            End If
        ElseIf Len(CellToTrim(grid(iRow, 5))) > 0 Then
            If Not skipInvalid Then
                errMsg = BadNumberMsg(iRow)
                Exit Function
            End If
        End If

        If Not IsLayerMode() Then
            wallVal = CoerceNumber(grid(iRow, 6), wallOk)
            If wallOk Then
                If distOk Then
                    wallWSum = wallWSum + wallVal * distVal
                    wallDistSum = wallDistSum + distVal
                    nWall = nWall + 1
                End If
            ElseIf Len(CellToTrim(grid(iRow, 6))) > 0 Then
                If Not skipInvalid Then
                    errMsg = BadNumberMsg(iRow)
                    Exit Function
                End If
            End If
        End If
SoilAvg_NextRow:
    Next iRow

    If nRes > 0 And soilRatioSum <> 0 Then
        avg = Application.Round((soilLenSum / soilRatioSum) ^ 2, 2)
    End If
    If wallDistSum > 0 Then wallAvg = Application.Round(wallWSum / wallDistSum, 4)
    If nRes > 0 And rhoMin > 0 And layerInvSum <> 0 Then
        mRhoMin = Application.Round(rhoMin, 2)
        mHetK = Application.Round(soilLenSum / (rhoMin * layerInvSum), 3)
    End If
    TryCalcFromGrid = True
End Function

Private Function BadNumberMsg(ByVal iRow As Long) As String
    BadNumberMsg = Ru("041D 0435 043A 043E 0440 0440 0435 043A 0442 043D 043E 0435 0020 0447 0438 0441 043B 043E 0020 0432 0020 0441 0442 0440 043E 043A 0435 0020") & CStr(iRow)
End Function

' ================================================================
' session in G1:G4 so a VBA reset does not lose the target cell
' ================================================================
Private Sub SaveSessionToSheet(ByVal ws As Worksheet)
    ws.Cells(1, META_COL).Value = mKey
    ws.Cells(2, META_COL).Value = mRangeName
    ws.Cells(3, META_COL).Value = mIdx
    ws.Cells(4, META_COL).Value = mRowCount
    If Not mPrevSheet Is Nothing Then
        ws.Cells(5, META_COL).Value = mPrevSheet.Name
    End If
End Sub

Private Function RestoreSessionFromSheet(ByRef ws As Worksheet) As Boolean
    Dim prevName As String
    Dim savedName As String
    Dim savedIdx As Long

    Set ws = Nothing
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_SOIL_AVG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    mKey = CStr(ws.Cells(1, META_COL).Value)
    savedName = Trim$(CStr(ws.Cells(2, META_COL).Value))
    savedIdx = CLng(Val(CStr(ws.Cells(3, META_COL).Value)))
    mRowCount = CLng(Val(CStr(ws.Cells(4, META_COL).Value)))
    If mRowCount < 1 Then mRowCount = 1

    If InStr(1, mKey, "|") > 0 Then
        Call ParseSoilAvgKey(mKey, mRangeName, mIdx)
    ElseIf StrComp(savedName, NM_ANOD, vbTextCompare) = 0 Or _
           StrComp(savedName, NM_PIPE, vbTextCompare) = 0 Or _
           StrComp(savedName, NM_LAYER, vbTextCompare) = 0 Then
        mRangeName = savedName
        mIdx = savedIdx
        If mIdx < 1 Then mIdx = 1
        mKey = KeyFromNameIdx(mRangeName, mIdx)
    Else
        mRangeName = NM_PIPE
        mIdx = 1
        mKey = KeyFromNameIdx(mRangeName, mIdx)
    End If

    Set mTarget = LiveTargetCell()

    prevName = CStr(ws.Cells(5, META_COL).Value)
    If Len(prevName) > 0 Then
        On Error Resume Next
        Set mPrevSheet = thisWorkbook.Worksheets(prevName)
        On Error GoTo 0
    End If

    Call SoilAvgUnprotect(ws)
    RestoreSessionFromSheet = True
End Function

Private Function IsSoilAvgUiSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsSoilAvgUiSheet = (StrComp(ws.Name, SHEET_SOIL_AVG, vbTextCompare) = 0)
End Function

' ================================================================
' persist TableSoilAvg at AA:AG
' ================================================================
Private Sub PersistGridRows(ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim oldVals As Variant
    Dim grid As Variant
    Dim outArr() As Variant
    Dim nGrid As Long
    Dim nOld As Long
    Dim nKeep As Long
    Dim nNew As Long
    Dim nOut As Long
    Dim nCur As Long
    Dim i As Long
    Dim k As Long
    Dim distOk As Boolean
    Dim resOk As Boolean
    Dim wallOk As Boolean
    Dim distVal As Double
    Dim resVal As Double
    Dim wallVal As Double
    Dim hdr As Range
    Dim dest As Range

    Call SoilAvgUnprotect(ws)
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_SOIL_AVG)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    Call EnsureSoilAvgTableCols(tbl)

    grid = ReadGridValues(ws, nGrid)
    nNew = 0
    If nGrid >= 1 Then
        For i = 1 To nGrid
            If Not VariantRowEmpty(grid, i) Then nNew = nNew + 1
        Next i
    End If

    nOld = 0
    nKeep = 0
    If Not tbl.DataBodyRange Is Nothing Then
        oldVals = tbl.DataBodyRange.Value
        If IsArray(oldVals) Then
            nOld = UBound(oldVals, 1)
            For i = 1 To nOld
                If CStr(oldVals(i, 1)) <> mKey Then nKeep = nKeep + 1
            Next i
        End If
    End If

    nOut = nKeep + nNew
    Set hdr = tbl.HeaderRowRange
    If nOut < 1 Then
        If Not tbl.DataBodyRange Is Nothing Then
            On Error Resume Next
            tbl.DataBodyRange.Delete
            On Error GoTo 0
        End If
        Exit Sub
    End If

    ReDim outArr(1 To nOut, 1 To STORE_COLS)
    k = 0
    If nOld >= 1 Then
        For i = 1 To nOld
            If CStr(oldVals(i, 1)) <> mKey Then
                k = k + 1
                outArr(k, 1) = oldVals(i, 1)
                outArr(k, 2) = oldVals(i, 2)
                outArr(k, 3) = oldVals(i, 3)
                outArr(k, 4) = oldVals(i, 4)
                outArr(k, 5) = oldVals(i, 5)
                outArr(k, 6) = oldVals(i, 6)
                If UBound(oldVals, 2) >= 7 Then
                    outArr(k, 7) = oldVals(i, 7)
                Else
                    outArr(k, 7) = ""
                End If
            End If
        Next i
    End If
    If nGrid >= 1 Then
        For i = 1 To nGrid
            If Not VariantRowEmpty(grid, i) Then
                k = k + 1
                outArr(k, 1) = mKey
                outArr(k, 2) = CellToTrim(grid(i, 1))
                outArr(k, 3) = CellToTrim(grid(i, 2))
                outArr(k, 4) = CellToTrim(grid(i, 3))
                distVal = CoerceNumber(grid(i, 4), distOk)
                If distOk Then
                    outArr(k, 5) = distVal
                Else
                    outArr(k, 5) = ""
                End If
                resVal = CoerceNumber(grid(i, 5), resOk)
                If resOk Then
                    outArr(k, 6) = resVal
                Else
                    outArr(k, 6) = ""
                End If
                wallVal = CoerceNumber(grid(i, 6), wallOk)
                If wallOk Then
                    outArr(k, 7) = wallVal
                Else
                    outArr(k, 7) = ""
                End If
            End If
        Next i
    End If

    nCur = 0
    If Not tbl.DataBodyRange Is Nothing Then nCur = tbl.ListRows.Count
    Set dest = ws.Range(ws.Cells(hdr.Row, STORE_COL), ws.Cells(hdr.Row + nOut, STORE_COL + STORE_COLS - 1))
    If nOut <> nCur Then
        On Error Resume Next
        Application.DisplayAlerts = False
        tbl.Resize dest
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Value = outArr
End Sub

Private Function LoadSoilAvgRows(ByVal key As String) As Collection
    Dim c As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim v As Variant
    Dim iRow As Long
    Dim nRow As Long

    Set c = New Collection
    Set LoadSoilAvgRows = c
    If Not WorksheetExists(SHEET_SOIL_AVG) Then Exit Function
    Set ws = thisWorkbook.Worksheets(SHEET_SOIL_AVG)
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_SOIL_AVG)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Call EnsureSoilAvgTableCols(tbl)
    v = tbl.DataBodyRange.Value
    If Not IsArray(v) Then
        If CStr(tbl.DataBodyRange.Cells(1, 1).Value) = key Then
            c.Add Array(CStr(tbl.DataBodyRange.Cells(1, 2).Value), _
                        CStr(tbl.DataBodyRange.Cells(1, 3).Value), _
                        CStr(tbl.DataBodyRange.Cells(1, 4).Value), _
                        tbl.DataBodyRange.Cells(1, 5).Value, _
                        tbl.DataBodyRange.Cells(1, 6).Value, _
                        tbl.DataBodyRange.Cells(1, 7).Value)
        End If
        Exit Function
    End If
    nRow = UBound(v, 1)
    For iRow = 1 To nRow
        If CStr(v(iRow, 1)) = key Then
            If UBound(v, 2) >= 7 Then
                c.Add Array(CStr(v(iRow, 2)), CStr(v(iRow, 3)), CStr(v(iRow, 4)), v(iRow, 5), v(iRow, 6), v(iRow, 7))
            Else
                c.Add Array(CStr(v(iRow, 2)), CStr(v(iRow, 3)), CStr(v(iRow, 4)), v(iRow, 5), v(iRow, 6), Empty)
            End If
        End If
    Next iRow
End Function

Private Function EnsureSoilAvgSheet() As Worksheet
    Dim ws As Worksheet
    Dim hdr As Range
    Dim tbl As ListObject

    If WorksheetExists(SHEET_SOIL_AVG) Then
        Set ws = thisWorkbook.Worksheets(SHEET_SOIL_AVG)
    Else
        On Error Resume Next
        Application.DisplayAlerts = False
        Set ws = thisWorkbook.Worksheets.Add(Type:=xlWorksheet)
        ws.Name = SHEET_SOIL_AVG
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Function
    Set EnsureSoilAvgSheet = ws
    Call SoilAvgUnprotect(ws)

    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_SOIL_AVG)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        If tbl.Range.Column < STORE_COL Then
            On Error Resume Next
            Application.DisplayAlerts = False
            tbl.Delete
            Application.DisplayAlerts = True
            Set tbl = Nothing
            On Error GoTo 0
        Else
            Call EnsureSoilAvgTableCols(tbl)
            Exit Function
        End If
    End If

    Set hdr = ws.Range(ws.Cells(1, STORE_COL), ws.Cells(1, STORE_COL + STORE_COLS - 1))
    hdr.Value = Array("Key", "Section", "Picket", "Geo", "Distance", "Resistivity", "WallThickness")
    Set tbl = ws.ListObjects.Add(xlSrcRange, hdr, , xlYes)
    tbl.Name = TABLE_SOIL_AVG
End Function

Private Sub EnsureSoilAvgTableCols(ByVal tbl As ListObject)
    Dim lc As ListColumn
    If tbl Is Nothing Then Exit Sub
    On Error Resume Next
    Do While tbl.ListColumns.Count < STORE_COLS
        Set lc = tbl.ListColumns.Add
        If tbl.ListColumns.Count = 7 Then lc.Name = "WallThickness"
    Loop
    On Error GoTo 0
End Sub

' ListRows.Add / ListRows.Delete raise 1004 on a protected sheet.
Private Sub SoilAvgUnprotect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

Private Sub WriteSoilAvgTarget(ByVal rng As Range, ByVal avg As Double)
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim oldEvents As Boolean

    Set ws = rng.Worksheet
    wasProt = ws.ProtectContents
    oldEvents = Application.EnableEvents
    On Error Resume Next
    If wasProt Then ws.Unprotect Password:=SHEET_PASSWORD
    Application.EnableEvents = False
    rng.Value = avg
    Application.EnableEvents = oldEvents
    If wasProt Then
        If StrComp(ws.Name, SHEET_PIPE, vbTextCompare) = 0 Then
            ws.Protect Password:=SHEET_PASSWORD, AllowInsertingRows:=False, AllowDeletingRows:=False, _
                       AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
        Else
            ws.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
        End If
    End If
    On Error GoTo 0
End Sub

' Keep existing survey buttons (row, caption, font). Only add a button
' for a section column that does not already have one.
Public Sub PlacePipeSoilAvgButtons(ByVal ws As Worksheet, ByVal colBr As Long)
    Dim rng As Range
    Dim i As Long
    Dim shp As Shape
    Dim act As String
    Dim cap As String
    Dim w As Double
    Dim h As Double
    Dim relLeft As Double
    Dim relTop As Double
    Dim fontSize As Double
    Dim cell As Range
    Dim startCol As Long
    Dim tmplRow As Long
    Dim colIdx As Long
    Dim hasBtn() As Boolean

    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    Set rng = ws.Range(NM_PIPE)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    If colBr < 1 Then colBr = rng.Columns.Count
    If colBr < 1 Then Exit Sub
    startCol = rng.Column
    ReDim hasBtn(1 To colBr)

    cap = ""
    w = 36
    h = 18
    relLeft = 0
    relTop = 0
    fontSize = 0
    tmplRow = rng.Row

    For i = 1 To ws.Shapes.Count
        act = ""
        On Error Resume Next
        Set shp = ws.Shapes(i)
        act = CStr(shp.OnAction)
        On Error GoTo 0
        If InStr(1, act, "button112_click", vbTextCompare) > 0 Then
            colIdx = shp.TopLeftCell.Column - startCol + 1
            If colIdx >= 1 And colIdx <= colBr Then hasBtn(colIdx) = True
            If Len(cap) = 0 Then
                On Error Resume Next
                cap = shp.TextFrame.Characters.Text
                If Len(cap) = 0 Then cap = ws.Buttons(shp.Name).Text
                w = shp.Width
                h = shp.Height
                relLeft = shp.Left - shp.TopLeftCell.Left
                relTop = shp.Top - shp.TopLeftCell.Top
                tmplRow = shp.TopLeftCell.Row
                fontSize = shp.TextFrame.Characters.Font.Size
                If fontSize = 0 Then fontSize = ws.Buttons(shp.Name).Font.Size
                On Error GoTo 0
            End If
        End If
    Next i

    If Len(cap) = 0 Then cap = Ru("0438 0437 044B 0441 043A 002E")

    For i = 1 To colBr
        If Not hasBtn(i) Then
            Set cell = ws.Cells(tmplRow, startCol + i - 1)
            Call AddFormBtn(ws, "btnPipeSoilAvg" & CStr(i), _
                            cell.Left + relLeft, cell.Top + relTop, w, h, cap, "button112_click", fontSize)
        End If
    Next i
End Sub

Private Function IsLayerButtonAction(ByVal act As String) As Boolean
    If InStr(1, act, "buttonLayerDeep_click", vbTextCompare) > 0 Then
        IsLayerButtonAction = True
    ElseIf InStr(1, act, "btnImportResistivityLayer", vbTextCompare) > 0 Then
        IsLayerButtonAction = True
    ElseIf InStr(1, act, "btnImportSoilHeterogeneity", vbTextCompare) > 0 Then
        IsLayerButtonAction = True
    ElseIf InStr(1, act, "ShowSoilAvgFormLayer", vbTextCompare) > 0 Then
        IsLayerButtonAction = True
    End If
End Function

' Keep existing Anod layer-survey buttons (row 53). Place only when
' typeMountingAG contains "глубинный". Copied-on-sync buttons are stripped first.
Public Sub PlaceAnodLayerButtons(ByVal ws As Worksheet, ByVal colBr As Long)
    Dim rng As Range
    Dim rngMount As Range
    Dim i As Long
    Dim shp As Shape
    Dim act As String
    Dim cap As String
    Dim w As Double
    Dim h As Double
    Dim relLeft As Double
    Dim relTop As Double
    Dim fontSize As Double
    Dim cell As Range
    Dim startCol As Long
    Dim tmplRow As Long
    Dim colIdx As Long
    Dim hasBtn() As Boolean
    Dim wantBtn() As Boolean
    Dim mountVal As String
    Dim wasProt As Boolean
    Dim shpNew As Shape

    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:=SHEET_PASSWORD
    Set rng = ws.Range(NM_LAYER)
    Set rngMount = ws.Range("typeMountingAG")
    Err.Clear
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    If colBr < 1 Then colBr = rng.Columns.Count
    If rng.Columns.Count > 0 And colBr > rng.Columns.Count Then colBr = rng.Columns.Count
    If colBr < 1 Then Exit Sub
    startCol = rng.Column
    ReDim hasBtn(1 To colBr)
    ReDim wantBtn(1 To colBr)

    For i = 1 To colBr
        mountVal = ""
        On Error Resume Next
        If Not rngMount Is Nothing Then
            If i <= rngMount.Columns.Count Then mountVal = CStr(rngMount.Cells(1, i).Value)
        End If
        If Len(Trim$(mountVal)) = 0 Then
            mountVal = CStr(ws.Cells(FILTER_START_ROW + 1, startCol + i - 1).Value)
        End If
        Err.Clear
        On Error GoTo 0
        wantBtn(i) = MountTypeIsDeepAG(mountVal)
    Next i

    cap = Ru("0412 0432 043E 0434") & Chr(10) & _
          Ru("0434 0430 043D 043D 044B 0445") & Chr(10) & _
          Ru("0438 0437 044B 0441 043A 0430 043D 0438 0439")
    fontSize = 8
    relLeft = 0
    relTop = 0
    tmplRow = rng.Row
    h = ws.Range(ws.Cells(rng.Row, 1), ws.Cells(rng.Row + 1, 1)).Height
    If h < 1 Then h = ws.Rows(rng.Row).Height + ws.Rows(rng.Row + 1).Height
    w = MeasureCalibriWordWidth(ws, Ru("0438 0437 044B 0441 043A 0430 043D 0438 0439"), fontSize)
    If w < 8 Then w = fontSize * 0.55 * 9 + 6

    For i = ws.Shapes.Count To 1 Step -1
        act = ""
        colIdx = 0
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(i)
        If shp Is Nothing Then
            Err.Clear
            On Error GoTo 0
            GoTo PlaceAnod_NextShp
        End If
        act = CStr(shp.OnAction)
        Set cell = ShapeAnchorCell(ws, shp)
        If Not cell Is Nothing Then colIdx = cell.Column - startCol + 1
        Err.Clear
        On Error GoTo 0
        If Not IsLayerButtonAction(act) Then GoTo PlaceAnod_NextShp
        If colIdx >= 1 And colIdx <= colBr Then
            If wantBtn(colIdx) Then
                hasBtn(colIdx) = True
                Set cell = ws.Cells(tmplRow, startCol + colIdx - 1)
                Call ApplyAnodLayerBtnStyle(ws, shp, AnodLayerBtnLeft(cell), ws.Rows(tmplRow).Top, w, h, cap)
            Else
                On Error Resume Next
                shp.Delete
                Err.Clear
                On Error GoTo 0
            End If
        Else
            On Error Resume Next
            shp.Delete
            Err.Clear
            On Error GoTo 0
        End If
PlaceAnod_NextShp:
    Next i

    For i = 1 To colBr
        If wantBtn(i) And Not hasBtn(i) Then
            Set cell = ws.Cells(tmplRow, startCol + i - 1)
            Set shpNew = AddFormBtn(ws, "btnAnodLayerDeep" & CStr(i), _
                            AnodLayerBtnLeft(cell), ws.Rows(tmplRow).Top, w, h, cap, "buttonLayerDeep_click", fontSize, "Calibri")
            If Not shpNew Is Nothing Then
                Call ApplyAnodLayerBtnStyle(ws, shpNew, AnodLayerBtnLeft(cell), ws.Rows(tmplRow).Top, w, h, cap)
            End If
        End If
    Next i

    On Error Resume Next
    If wasProt Then ws.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    Err.Clear
    On Error GoTo 0
End Sub

Private Function MountTypeIsDeepAG(ByVal mountRaw As String) As Boolean
    Dim s As String
    Dim needle As String
    s = LCase$(Trim$(mountRaw))
    If Len(s) = 0 Then Exit Function
    needle = LCase$(Ru("0433 043B 0443 0431 0438 043D 043D 044B 0439"))
    MountTypeIsDeepAG = (InStr(1, s, needle, vbTextCompare) > 0)
End Function

Public Sub DeleteAnodLayerButtonsInColumns(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim i As Long
    Dim shp As Shape
    Dim act As String
    Dim colShp As Long
    Dim wasProt As Boolean
    Dim cellDel As Range

    If ws Is Nothing Then Exit Sub
    If firstCol > lastCol Then Exit Sub
    wasProt = ws.ProtectContents
    On Error Resume Next
    If wasProt Then ws.Unprotect Password:=SHEET_PASSWORD
    For i = ws.Shapes.Count To 1 Step -1
        act = ""
        colShp = 0
        Set shp = Nothing
        Set shp = ws.Shapes(i)
        If Not shp Is Nothing Then
            act = CStr(shp.OnAction)
            Set cellDel = Nothing
            Set cellDel = ShapeAnchorCell(ws, shp)
            If Not cellDel Is Nothing Then colShp = cellDel.Column
            If IsLayerButtonAction(act) Then
                If colShp >= firstCol And colShp <= lastCol Then shp.Delete
            End If
        End If
    Next i
    If wasProt Then
        ws.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' Form Control buttons (not ActiveX). Recreated each Show with Ru() captions.
' ================================================================
Private Sub PlaceSoilAvgButtons(ByVal ws As Worksheet)
    Dim topPos As Double
    Dim leftPos As Double
    Dim btnW As Double
    Dim gap As Double

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    Call DeleteSoilAvgButtons(ws)
    topPos = ws.Rows(UI_BTN_ROW).Top
    leftPos = ws.Columns(1).Left
    btnW = 90
    gap = 8

    Call AddFormBtn(ws, "btnSoilAvgCalc", leftPos, topPos, btnW, 22, _
                    Ru("0412 044B 0447 0438 0441 043B 0438 0442 044C"), "SoilAvg_OnCalc")
    Call AddFormBtn(ws, "btnSoilAvgApply", leftPos + (btnW + gap) * 1, topPos, btnW, 22, _
                    Ru("041F 043E 0434 0441 0442 0430 0432 0438 0442 044C"), "SoilAvg_OnApply")
    Call AddFormBtn(ws, "btnSoilAvgSave", leftPos + (btnW + gap) * 2, topPos, btnW, 22, _
                    Ru("0421 043E 0445 0440 0430 043D 0438 0442 044C"), "SoilAvg_OnSave")
    Call AddFormBtn(ws, "btnSoilAvgClose", leftPos + (btnW + gap) * 3, topPos, btnW, 22, _
                    Ru("0417 0430 043A 0440 044B 0442 044C"), "SoilAvg_OnClose")
End Sub

Private Sub DeleteSoilAvgButtons(ByVal ws As Worksheet)
    Dim i As Long
    Dim shp As Shape
    Dim act As String
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        act = ""
        On Error Resume Next
        act = CStr(shp.OnAction)
        On Error GoTo 0
        If InStr(1, act, "SoilAvg_On", vbTextCompare) > 0 Then
            shp.Delete
        End If
    Next i
End Sub

Private Function AddFormBtn(ByVal ws As Worksheet, ByVal btnName As String, _
                       ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                       ByVal caption As String, ByVal procName As String, _
                       Optional ByVal fontSize As Double = 0, _
                       Optional ByVal fontName As String = "") As Shape
    Dim shp As Shape
    Dim autoName As String
    Dim btn As Object
    Dim tf As Object

    ' Mac Excel 16.x: Form Control Name is read-only (error 70). Keep the auto name.
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, x, y, w, h)
    Set AddFormBtn = shp
    autoName = shp.Name
    shp.OnAction = "'" & Replace(thisWorkbook.Name, "'", "''") & "'!" & procName
    On Error Resume Next
    shp.AlternativeText = btnName
    Set tf = shp.TextFrame
    tf.Characters.Text = caption
    If fontSize > 0 Then tf.Characters.Font.Size = fontSize
    If Len(fontName) > 0 Then tf.Characters.Font.Name = fontName
    Set btn = ws.Buttons(autoName)
    btn.Text = caption
    If fontSize > 0 Then btn.Font.Size = fontSize
    If Len(fontName) > 0 Then btn.Font.Name = fontName
    Err.Clear
    On Error GoTo 0
End Function

Private Sub ApplyAnodLayerBtnStyle(ByVal ws As Worksheet, ByVal shp As Shape, _
                                   ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                                   ByVal cap As String)
    Dim btn As Object
    Dim tf As Object
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.Left = x
    shp.Top = y
    shp.Width = w
    shp.Height = h
    Set tf = shp.TextFrame
    tf.Characters.Text = cap
    tf.Characters.Font.Name = "Calibri"
    tf.Characters.Font.Size = 8
    tf.WordWrap = True
    tf.HorizontalAlignment = xlHAlignCenter
    tf.VerticalAlignment = xlVAlignCenter
    Set btn = ws.Buttons(shp.Name)
    btn.Text = cap
    btn.Font.Name = "Calibri"
    btn.Font.Size = 8
    Err.Clear
    On Error GoTo 0
End Sub

Private Function MeasureCalibriWordWidth(ByVal ws As Worksheet, ByVal word As String, ByVal fontSize As Double) As Double
    If fontSize < 1 Then fontSize = 8
    If Len(word) < 1 Then
        MeasureCalibriWordWidth = fontSize * 0.55 * 9 + 6
    Else
        MeasureCalibriWordWidth = fontSize * 0.55 * Len(word) + 6
    End If
End Function

Private Function ParseDotNumber(ByVal raw As String, ByRef ok As Boolean) As Double
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim seenDot As Boolean
    Dim seenDigit As Boolean
    Dim startAt As Long

    ok = False
    ParseDotNumber = 0
    s = Trim$(raw)
    If Len(s) = 0 Then Exit Function
    s = Replace(s, " ", "")
    s = Replace(s, ",", ".")
    startAt = 1
    If Mid$(s, 1, 1) = "+" Or Mid$(s, 1, 1) = "-" Then startAt = 2
    If startAt > Len(s) Then Exit Function
    For i = startAt To Len(s)
        c = Mid$(s, i, 1)
        If c = "." Then
            If seenDot Then Exit Function
            seenDot = True
        ElseIf c >= "0" And c <= "9" Then
            seenDigit = True
        Else
            Exit Function
        End If
    Next i
    If Not seenDigit Then Exit Function
    ParseDotNumber = Val(s)
    ok = True
End Function

Private Function IsWholeNumber(ByVal raw As String, ByRef ok As Boolean) As Boolean
    Dim n As Double
    n = ParseDotNumber(raw, ok)
    If Not ok Then
        IsWholeNumber = False
        Exit Function
    End If
    If n <> CLng(n) Then
        ok = False
        IsWholeNumber = False
        Exit Function
    End If
    IsWholeNumber = True
End Function

Private Function FmtNum(ByVal x As Double) As String
    FmtNum = Trim$(Str$(x))
End Function

Private Sub SoilAvgFastBegin()
    On Error Resume Next
    If mFastDepth = 0 Then
        mOldCalc = Application.Calculation
        mOldUpd = Application.ScreenUpdating
        mOldEv = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.Cursor = xlWait
    End If
    mFastDepth = mFastDepth + 1
    On Error GoTo 0
End Sub

Private Sub SoilAvgFastEnd()
    On Error Resume Next
    If mFastDepth > 0 Then mFastDepth = mFastDepth - 1
    If mFastDepth = 0 Then
        Application.Calculation = mOldCalc
        Application.EnableEvents = mOldEv
        Application.ScreenUpdating = mOldUpd
        Application.Cursor = xlDefault
        Application.StatusBar = False
    End If
    On Error GoTo 0
End Sub
