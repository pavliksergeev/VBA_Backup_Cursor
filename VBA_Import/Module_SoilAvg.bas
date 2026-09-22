Attribute VB_Name = "Module_SoilAvg"

' ================================================================
' Survey popup: rows of section / picket / geo / distance / resistivity.
' Average of non-empty resistivity is written to the calling named-range cell.
' Pipe button112 -> soilResistivityAvg. Anod: ShowSoilAvgForm SOILAVG_TARGET_ANOD
'   -> resistivitySoilAG. Raw rows persist on hidden sheet _SoilAvg.
' Mac: Insert UserForm, name it frmSoilAvg, paste frmSoilAvg.frm from Option Explicit.
' Windows: File Import frmSoilAvg.frm, then this module.
' ================================================================
Option Explicit

Public Const SOILAVG_TARGET_AUTO As Long = 0
Public Const SOILAVG_TARGET_PIPE As Long = 1
Public Const SOILAVG_TARGET_ANOD As Long = 2

Private Const SOILAVG_MAX_ROWS As Long = 40
Private Const SOILAVG_DEFAULT_ROWS As Long = 5
Private Const SOILAVG_ROW_H As Single = 22
Private Const FORM_NAME As String = "frmSoilAvg"

Private mFrm As Object
Private mTarget As Range
Private mKey As String
Private mRowCount As Long

' ================================================================
' public entry
' ================================================================
Public Sub ShowSoilAvgForm(Optional ByVal targetMode As Long = SOILAVG_TARGET_AUTO)
    Dim frm As Object

    Set mTarget = ResolveSoilAvgTarget(targetMode)
    If mTarget Is Nothing Then
        ' целевая ячейка не найдена.
        MsgBox Ru("0426 0435 043B 0435 0432 0430 044F 0020 044F 0447 0435 0439 043A 0430 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 002E"), vbExclamation
        Exit Sub
    End If
    mKey = MakeSoilAvgKey(mTarget)

    On Error Resume Next
    Set frm = UserForms.Add(FORM_NAME)
    On Error GoTo 0
    If frm Is Nothing Then
        MsgBox SoilAvgFormMissingMsg(), vbExclamation
        Exit Sub
    End If

    frm.Show vbModal
    On Error Resume Next
    Unload frm
    On Error GoTo 0
    Set mFrm = Nothing
    Set mTarget = Nothing
    mKey = ""
    mRowCount = 0
End Sub

Public Sub ShowSoilAvgFormAnod()
    Call ShowSoilAvgForm(SOILAVG_TARGET_ANOD)
End Sub

' ================================================================
' called from frmSoilAvg
' ================================================================
Public Sub SoilAvg_BuildUi(ByVal frm As Object)
    Dim fra As Object
    Dim iRow As Long
    Dim saved As Collection

    Set mFrm = frm
    frm.Caption = Ru("0421 0440 0435 0434 043D 0435 0435 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 0020 0433 0440 0443 043D 0442 0430")
    frm.Width = 660
    frm.Height = 420

    Call PlaceLabel(frm, "lblTitle", 8, 6, 400, 18, frm.Caption)
    Call PlaceLabel(frm, "lblTarget", 8, 24, 630, 16, SoilAvgTargetCaption())

    Call PlaceLabel(frm, "hdrSec", 14, 46, 110, 16, Ru("0423 0447 0430 0441 0442 043E 043A"))
    Call PlaceLabel(frm, "hdrPk", 128, 46, 60, 16, Ru("041F 0438 043A 0435 0442"))
    Call PlaceLabel(frm, "hdrGeo", 192, 46, 160, 16, Ru("0413 0435 043E 043F 043E 0437 0438 0446 0438 044F"))
    Call PlaceLabel(frm, "hdrDist", 356, 46, 90, 16, Ru("0414 0438 0441 0442 0430 043D 0446 0438 044F 002C 0020 043C"))
    Call PlaceLabel(frm, "hdrRes", 450, 46, 160, 16, _
                    Ru("0421 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 0435 002C 0020 041E 043C 00B7 043C"))

    Set fra = EnsureCtl(frm, "Forms.Frame.1", "fraRows")
    fra.Caption = ""
    fra.Left = 8
    fra.Top = 64
    fra.Width = 628
    fra.Height = 248
    On Error Resume Next
    fra.ScrollBars = 2
    fra.KeepScrollBarsVisible = 2
    On Error GoTo 0

    Call PlaceLabel(frm, "lblAvg", 8, 318, 310, 16, Ru("0421 0440 0435 0434 043D 0435 0435 003A") & " -")
    Call PlaceLabel(frm, "lblSum", 330, 318, 300, 16, Ru("0421 0443 043C 043C 0430 0020 0434 0438 0441 0442 0430 043D 0446 0438 0439 003A") & " -")

    Call PlaceButton(frm, "btnAdd", 8, 342, 96, 24, Ru("0414 043E 0431 0430 0432 0438 0442 044C"))
    Call PlaceButton(frm, "btnRemove", 108, 342, 96, 24, Ru("0423 0434 0430 043B 0438 0442 044C"))
    Call PlaceButton(frm, "btnCalc", 208, 342, 96, 24, Ru("0412 044B 0447 0438 0441 043B 0438 0442 044C"))
    Call PlaceButton(frm, "btnApply", 308, 342, 96, 24, Ru("041F 043E 0434 0441 0442 0430 0432 0438 0442 044C"))
    Call PlaceButton(frm, "btnSave", 408, 342, 96, 24, Ru("0421 043E 0445 0440 0430 043D 0438 0442 044C"))
    Call PlaceButton(frm, "btnClose", 508, 342, 96, 24, Ru("0417 0430 043A 0440 044B 0442 044C"))

    Set saved = LoadSoilAvgRows(mKey)
    If saved.Count < 1 Then
        mRowCount = 0
        For iRow = 1 To SOILAVG_DEFAULT_ROWS
            Call SoilAvg_AddRowSilent
        Next iRow
    Else
        mRowCount = 0
        For iRow = 1 To saved.Count
            If iRow > SOILAVG_MAX_ROWS Then Exit For
            Call SoilAvg_AddRowSilent
            Call FillRowFromArr(iRow, saved(iRow))
        Next iRow
    End If
    Call RefreshScroll
    Call SoilAvg_OnCalcSilent
End Sub

Public Sub SoilAvg_OnAdd()
    If mFrm Is Nothing Then Exit Sub
    If mRowCount >= SOILAVG_MAX_ROWS Then
        ' достигнут максимум строк.
        MsgBox Ru("0414 043E 0441 0442 0438 0433 043D 0443 0442 0020 043C 0430 043A 0441 0438 043C 0443 043C 0020 0441 0442 0440 043E 043A 002E"), vbInformation
        Exit Sub
    End If
    Call SoilAvg_AddRowSilent
    Call RefreshScroll
End Sub

Public Sub SoilAvg_OnRemove()
    Dim prefixes As Variant
    Dim p As Variant
    Dim fra As Object

    If mFrm Is Nothing Then Exit Sub
    If mRowCount <= 1 Then
        Call ClearRow(1)
        Call SoilAvg_OnCalcSilent
        Exit Sub
    End If

    Set fra = mFrm.Controls("fraRows")
    prefixes = Array("txtSec_", "txtPk_", "txtGeo_", "txtDist_", "txtRes_")
    On Error Resume Next
    For Each p In prefixes
        fra.Controls.Remove CStr(p) & CStr(mRowCount)
        mFrm.Controls.Remove CStr(p) & CStr(mRowCount)
    Next p
    On Error GoTo 0
    mRowCount = mRowCount - 1
    Call RefreshScroll
    Call SoilAvg_OnCalcSilent
End Sub

Public Sub SoilAvg_OnCalc()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim errMsg As String

    If Not TryCalcFromForm(avg, distSum, nRes, errMsg) Then
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    If nRes < 1 Then
        ' нет значений сопротивления для среднего.
        MsgBox Ru("041D 0435 0442 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 044F 0020 0434 043B 044F 0020") & _
               Ru("0441 0440 0435 0434 043D 0435 0433 043E 002E"), vbExclamation
        Exit Sub
    End If
    Call ShowCalcLabels(avg, distSum, nRes)
End Sub

Public Sub SoilAvg_OnApply()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim errMsg As String

    If mTarget Is Nothing Then
        MsgBox Ru("0426 0435 043B 0435 0432 0430 044F 0020 044F 0447 0435 0439 043A 0430 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 002E"), vbExclamation
        Exit Sub
    End If
    If Not TryCalcFromForm(avg, distSum, nRes, errMsg) Then
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    If nRes < 1 Then
        MsgBox Ru("041D 0435 0442 0020 0437 043D 0430 0447 0435 043D 0438 0439 0020 0441 043E 043F 0440 043E 0442 0438 0432 043B 0435 043D 0438 044F 0020 0434 043B 044F 0020") & _
               Ru("0441 0440 0435 0434 043D 0435 0433 043E 002E"), vbExclamation
        Exit Sub
    End If

    Call WriteSoilAvgTarget(mTarget, avg)
    Call PersistFormRows
    Call ShowCalcLabels(avg, distSum, nRes)
    ' среднее записано в
    MsgBox Ru("0421 0440 0435 0434 043D 0435 0435 0020 0437 0430 043F 0438 0441 0430 043D 043E 0020 0432 0020") & _
           mTarget.Address(False, False) & " (" & mKey & ")", vbInformation
End Sub

Public Sub SoilAvg_OnSave()
    Dim errMsg As String
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long

    If Not TryCalcFromForm(avg, distSum, nRes, errMsg, True) Then
        If Len(errMsg) > 0 Then MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    Call PersistFormRows
    ' данные изысканий сохранены.
    MsgBox Ru("0414 0430 043D 043D 044B 0435 0020 0438 0437 044B 0441 043A 0430 043D 0438 0439 0020 0441 043E 0445 0440 0430 043D 0435 043D 044B 002E"), vbInformation
End Sub

Public Sub SoilAvg_OnClose()
    If mFrm Is Nothing Then Exit Sub
    mFrm.Hide
End Sub

' ================================================================
' target cell: Caller shape, ActiveCell, or first cell of the named range
' ================================================================
Private Function ResolveSoilAvgTarget(ByVal targetMode As Long) As Range
    Dim rngPipe As Range
    Dim rngAnod As Range
    Dim callerCell As Range
    Dim mode As Long

    Set rngPipe = Named1xN("soilResistivityAvg", SHEET_PIPE)
    Set rngAnod = Named1xN("resistivitySoilAG", SHEET_ANOD)
    Set callerCell = CellFromCaller()

    mode = targetMode
    If mode = SOILAVG_TARGET_AUTO Then
        If Not callerCell Is Nothing Then
            If Not rngAnod Is Nothing Then
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

    If mode = SOILAVG_TARGET_ANOD Then
        Set ResolveSoilAvgTarget = PickCellIn1xN(rngAnod, callerCell)
    Else
        Set ResolveSoilAvgTarget = PickCellIn1xN(rngPipe, callerCell)
    End If
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

    On Error Resume Next
    v = Application.Caller
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set ws = ActiveSheet
    If VarType(v) = vbString Then
        Set shp = ws.Shapes(CStr(v))
        If Not shp Is Nothing Then Set CellFromCaller = shp.TopLeftCell
        If CellFromCaller Is Nothing Then
            Set btn = ws.Buttons(CStr(v))
            If Not btn Is Nothing Then Set CellFromCaller = btn.TopLeftCell
        End If
    End If
    On Error GoTo 0
End Function

Private Function PickCellIn1xN(ByVal rng As Range, ByVal hint As Range) As Range
    Dim idx As Long
    Dim colHint As Long

    If rng Is Nothing Then Exit Function
    If Not hint Is Nothing Then
        If Not Intersect(hint, rng) Is Nothing Then
            Set PickCellIn1xN = Intersect(hint, rng).Cells(1, 1)
            Exit Function
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
    If idx < 1 Then idx = 1
    If idx > rng.Columns.Count Then idx = rng.Columns.Count
    Set PickCellIn1xN = rng.Cells(1, idx)
End Function

Private Function MakeSoilAvgKey(ByVal rng As Range) As String
    Dim idx As Long
    Dim parent As Range
    Dim prefix As String

    prefix = "PIPE"
    If StrComp(rng.Worksheet.Name, SHEET_ANOD, vbTextCompare) = 0 Then
        prefix = "ANOD"
        Set parent = Named1xN("resistivitySoilAG", SHEET_ANOD)
    Else
        Set parent = Named1xN("soilResistivityAvg", SHEET_PIPE)
    End If
    idx = 1
    If Not parent Is Nothing Then
        idx = rng.Column - parent.Column + 1
        If idx < 1 Then idx = 1
    End If
    MakeSoilAvgKey = prefix & "|" & CStr(idx)
End Function

Private Function SoilAvgTargetCaption() As String
    Dim addr As String
    If mTarget Is Nothing Then
        addr = "-"
    Else
        addr = mTarget.Worksheet.Name & "!" & mTarget.Address(False, False)
    End If
    ' цель:
    SoilAvgTargetCaption = Ru("0426 0435 043B 044C 003A 0020") & addr & "  [" & mKey & "]"
End Function

' ================================================================
' form rows
' ================================================================
Private Sub SoilAvg_AddRowSilent()
    Dim fra As Object
    Dim topPos As Single
    Dim iRow As Long

    If mFrm Is Nothing Then Exit Sub
    If mRowCount >= SOILAVG_MAX_ROWS Then Exit Sub
    Set fra = mFrm.Controls("fraRows")
    mRowCount = mRowCount + 1
    iRow = mRowCount
    topPos = 4 + (iRow - 1) * SOILAVG_ROW_H
    Call PlaceText(fra, "txtSec_" & CStr(iRow), 4, topPos, 110, 18)
    Call PlaceText(fra, "txtPk_" & CStr(iRow), 118, topPos, 60, 18)
    Call PlaceText(fra, "txtGeo_" & CStr(iRow), 182, topPos, 160, 18)
    Call PlaceText(fra, "txtDist_" & CStr(iRow), 346, topPos, 80, 18)
    Call PlaceText(fra, "txtRes_" & CStr(iRow), 430, topPos, 90, 18)
End Sub

Private Sub FillRowFromArr(ByVal iRow As Long, ByVal arr As Variant)
    On Error Resume Next
    mFrm.Controls("txtSec_" & CStr(iRow)).Text = CStr(arr(0))
    mFrm.Controls("txtPk_" & CStr(iRow)).Text = CStr(arr(1))
    mFrm.Controls("txtGeo_" & CStr(iRow)).Text = CStr(arr(2))
    If Not IsEmpty(arr(3)) And Len(CStr(arr(3))) > 0 Then
        mFrm.Controls("txtDist_" & CStr(iRow)).Text = FmtNum(CDbl(arr(3)))
    End If
    If Not IsEmpty(arr(4)) And Len(CStr(arr(4))) > 0 Then
        mFrm.Controls("txtRes_" & CStr(iRow)).Text = FmtNum(CDbl(arr(4)))
    End If
    On Error GoTo 0
End Sub

Private Sub ClearRow(ByVal iRow As Long)
    On Error Resume Next
    mFrm.Controls("txtSec_" & CStr(iRow)).Text = ""
    mFrm.Controls("txtPk_" & CStr(iRow)).Text = ""
    mFrm.Controls("txtGeo_" & CStr(iRow)).Text = ""
    mFrm.Controls("txtDist_" & CStr(iRow)).Text = ""
    mFrm.Controls("txtRes_" & CStr(iRow)).Text = ""
    On Error GoTo 0
End Sub

Private Sub RefreshScroll()
    Dim fra As Object
    If mFrm Is Nothing Then Exit Sub
    Set fra = mFrm.Controls("fraRows")
    On Error Resume Next
    fra.ScrollTop = 0
    fra.ScrollHeight = 8 + mRowCount * SOILAVG_ROW_H
    On Error GoTo 0
End Sub

Private Sub SoilAvg_OnCalcSilent()
    Dim avg As Double
    Dim distSum As Double
    Dim nRes As Long
    Dim errMsg As String
    If TryCalcFromForm(avg, distSum, nRes, errMsg, True) Then
        Call ShowCalcLabels(avg, distSum, nRes)
    End If
End Sub

Private Sub ShowCalcLabels(ByVal avg As Double, ByVal distSum As Double, ByVal nRes As Long)
    Dim avgTxt As String
    Dim sumTxt As String
    If nRes < 1 Then
        avgTxt = "-"
    Else
        avgTxt = FmtNum(avg)
    End If
    sumTxt = FmtNum(distSum)
    mFrm.Controls("lblAvg").Caption = Ru("0421 0440 0435 0434 043D 0435 0435 003A") & " " & avgTxt
    mFrm.Controls("lblSum").Caption = Ru("0421 0443 043C 043C 0430 0020 0434 0438 0441 0442 0430 043D 0446 0438 0439 003A") & " " & sumTxt
End Sub

' skipInvalid=True: ignore bad cells (used for live labels / save of empty drafts)
Private Function TryCalcFromForm(ByRef avg As Double, ByRef distSum As Double, _
                                 ByRef nRes As Long, ByRef errMsg As String, _
                                 Optional ByVal skipInvalid As Boolean = False) As Boolean
    Dim iRow As Long
    Dim resOk As Boolean
    Dim distOk As Boolean
    Dim pkOk As Boolean
    Dim resVal As Double
    Dim distVal As Double
    Dim pkRaw As String
    Dim resSum As Double

    avg = 0
    distSum = 0
    nRes = 0
    errMsg = ""
    resSum = 0
    If mFrm Is Nothing Then Exit Function
    If mRowCount < 1 Then
        ' нужна хотя бы одна строка.
        errMsg = Ru("041D 0443 0436 043D 0430 0020 0445 043E 0442 044F 0020 0431 044B 0020 043E 0434 043D 0430 0020 0441 0442 0440 043E 043A 0430 002E")
        Exit Function
    End If

    For iRow = 1 To mRowCount
        pkRaw = Trim$(CStr(mFrm.Controls("txtPk_" & CStr(iRow)).Text))
        If Len(pkRaw) > 0 Then
            If Not IsWholeNumber(pkRaw, pkOk) Then
                If Not skipInvalid Then
                    ' пикет должен быть целым числом.
                    errMsg = Ru("041F 0438 043A 0435 0442 0020 0434 043E 043B 0436 0435 043D 0020 0431 044B 0442 044C 0020 0446 0435 043B 044B 043C 0020") & _
                             Ru("0447 0438 0441 043B 043E 043C 002E") & " #" & CStr(iRow)
                    Exit Function
                End If
            End If
        End If

        distVal = ParseDotNumber(CStr(mFrm.Controls("txtDist_" & CStr(iRow)).Text), distOk)
        If distOk Then
            distSum = distSum + distVal
        ElseIf Len(Trim$(CStr(mFrm.Controls("txtDist_" & CStr(iRow)).Text))) > 0 Then
            If Not skipInvalid Then
                errMsg = BadNumberMsg(iRow)
                Exit Function
            End If
        End If

        resVal = ParseDotNumber(CStr(mFrm.Controls("txtRes_" & CStr(iRow)).Text), resOk)
        If resOk Then
            resSum = resSum + resVal
            nRes = nRes + 1
        ElseIf Len(Trim$(CStr(mFrm.Controls("txtRes_" & CStr(iRow)).Text))) > 0 Then
            If Not skipInvalid Then
                errMsg = BadNumberMsg(iRow)
                Exit Function
            End If
        End If
    Next iRow

    If nRes > 0 Then avg = resSum / nRes
    TryCalcFromForm = True
End Function

Private Function BadNumberMsg(ByVal iRow As Long) As String
    ' некорректное число в строке
    BadNumberMsg = Ru("041D 0435 043A 043E 0440 0440 0435 043A 0442 043D 043E 0435 0020 0447 0438 0441 043B 043E 0020 0432 0020 0441 0442 0440 043E 043A 0435 0020") & CStr(iRow)
End Function

' ================================================================
' persist on _SoilAvg / TableSoilAvg
' ================================================================
Private Sub PersistFormRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim iRow As Long
    Dim lr As ListRow
    Dim resOk As Boolean
    Dim distOk As Boolean
    Dim resVal As Double
    Dim distVal As Double

    Set ws = EnsureSoilAvgSheet()
    If ws Is Nothing Then Exit Sub
    Set tbl = ws.ListObjects(TABLE_SOIL_AVG)
    Call DeleteRowsForKey(tbl, mKey)

    For iRow = 1 To mRowCount
        Set lr = tbl.ListRows.Add
        lr.Range.Cells(1, 1).Value = mKey
        lr.Range.Cells(1, 2).Value = Trim$(CStr(mFrm.Controls("txtSec_" & CStr(iRow)).Text))
        lr.Range.Cells(1, 3).Value = Trim$(CStr(mFrm.Controls("txtPk_" & CStr(iRow)).Text))
        lr.Range.Cells(1, 4).Value = Trim$(CStr(mFrm.Controls("txtGeo_" & CStr(iRow)).Text))
        distVal = ParseDotNumber(CStr(mFrm.Controls("txtDist_" & CStr(iRow)).Text), distOk)
        If distOk Then
            lr.Range.Cells(1, 5).Value = distVal
        Else
            lr.Range.Cells(1, 5).Value = ""
        End If
        resVal = ParseDotNumber(CStr(mFrm.Controls("txtRes_" & CStr(iRow)).Text), resOk)
        If resOk Then
            lr.Range.Cells(1, 6).Value = resVal
        Else
            lr.Range.Cells(1, 6).Value = ""
        End If
    Next iRow
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
    v = tbl.DataBodyRange.Value
    If Not IsArray(v) Then
        If CStr(tbl.DataBodyRange.Cells(1, 1).Value) = key Then
            c.Add Array(CStr(tbl.DataBodyRange.Cells(1, 2).Value), _
                        CStr(tbl.DataBodyRange.Cells(1, 3).Value), _
                        CStr(tbl.DataBodyRange.Cells(1, 4).Value), _
                        tbl.DataBodyRange.Cells(1, 5).Value, _
                        tbl.DataBodyRange.Cells(1, 6).Value)
        End If
        Exit Function
    End If
    nRow = UBound(v, 1)
    For iRow = 1 To nRow
        If CStr(v(iRow, 1)) = key Then
            c.Add Array(CStr(v(iRow, 2)), CStr(v(iRow, 3)), CStr(v(iRow, 4)), v(iRow, 5), v(iRow, 6))
        End If
    Next iRow
End Function

Private Sub DeleteRowsForKey(ByVal tbl As ListObject, ByVal key As String)
    Dim iRow As Long
    If tbl.ListRows.Count < 1 Then Exit Sub
    For iRow = tbl.ListRows.Count To 1 Step -1
        If CStr(tbl.ListRows(iRow).Range.Cells(1, 1).Value) = key Then
            tbl.ListRows(iRow).Delete
        End If
    Next iRow
End Sub

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
        ws.Visible = xlSheetVeryHidden
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Function
    Set EnsureSoilAvgSheet = ws

    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_SOIL_AVG)
    On Error GoTo 0
    If Not tbl Is Nothing Then Exit Function

    Set hdr = ws.Range("A1:F1")
    hdr.Value = Array("Key", "Section", "Picket", "Geo", "Distance", "Resistivity")
    Set tbl = ws.ListObjects.Add(xlSrcRange, hdr, , xlYes)
    tbl.Name = TABLE_SOIL_AVG
End Function

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

' ================================================================
' controls / parse
' ================================================================
Private Function EnsureCtl(ByVal parent As Object, ByVal progId As String, ByVal ctlName As String) As Object
    On Error Resume Next
    Set EnsureCtl = parent.Controls(ctlName)
    On Error GoTo 0
    If EnsureCtl Is Nothing Then
        Set EnsureCtl = parent.Controls.Add(progId, ctlName, True)
    End If
End Function

Private Sub PlaceLabel(ByVal parent As Object, ByVal ctlName As String, _
                       ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                       ByVal caption As String)
    Dim ctl As Object
    Set ctl = EnsureCtl(parent, "Forms.Label.1", ctlName)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
    ctl.Caption = caption
End Sub

Private Sub PlaceButton(ByVal parent As Object, ByVal ctlName As String, _
                        ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                        ByVal caption As String)
    Dim ctl As Object
    Set ctl = EnsureCtl(parent, "Forms.CommandButton.1", ctlName)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
    ctl.Caption = caption
    On Error Resume Next
    ctl.TakeFocusOnClick = False
    On Error GoTo 0
End Sub

Private Sub PlaceText(ByVal parent As Object, ByVal ctlName As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctl As Object
    Set ctl = EnsureCtl(parent, "Forms.TextBox.1", ctlName)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
End Sub

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

Private Function SoilAvgFormMissingMsg() As String
    ' форма frmSoilAvg не найдена. в редакторе vba: insert / userform ...
    SoilAvgFormMissingMsg = Ru("0424 043E 0440 043C 0430 0020 0066 0072 006D 0053 006F 0069 006C 0041 0076 0067 0020 043D 0435 0020 043D 0430 0439 0434 0435 043D 0430 002E 0020") & _
        Ru("0412 0020 0440 0435 0434 0430 043A 0442 043E 0440 0435 0020 0056 0042 0041 003A 0020 0049 006E 0073 0065 0072 0074 0020 002F 0020 0055 0073 0065 0072 0046 006F 0072 006D 002C 0020") & _
        Ru("0438 043C 044F 0020 0066 0072 006D 0053 006F 0069 006C 0041 0076 0067 002C 0020 0432 0441 0442 0430 0432 044C 0442 0435 0020 043A 043E 0434 0020 0438 0437 0020") & _
        Ru("0066 0072 006D 0053 006F 0069 006C 0041 0076 0067 002E 0066 0072 006D 0020 0028 0441 0020 004F 0070 0074 0069 006F 006E 0020 0045 0078 0070 006C 0069 0063 0069 0074 0029 002E")
End Function
