VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "thisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Workbook_Open()
    On Error GoTo CleanExit

    If DEBUG_MODE Then Debug.Print ">>> workbook_open: start"

    Application.EnableEvents = True
    Dim oldEnableEvents As Boolean
    oldEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' === get Anod sheet ===
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = thisWorkbook.Worksheets(SHEET_ANOD)
    On Error GoTo CleanExit
    If ws Is Nothing Then
        MsgBox "sheet " & SHEET_ANOD & " not found!"
        If DEBUG_MODE Then Debug.Print ">>> error: sheet " & SHEET_ANOD & " not found"
        GoTo CleanExit
    End If
    If DEBUG_MODE Then Debug.Print ">>> sheet " & SHEET_ANOD & " retrieved successfully"

    ' === unprotect Anod ===
    If DEBUG_MODE Then Debug.Print ">>> attempting to unprotect Anod sheet"
    On Error Resume Next
    ws.Unprotect Password:="ptm"
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> failed to unprotect: " & Err.Description
        MsgBox "unprotect error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> protection removed successfully"
    End If
    On Error GoTo 0

    ' === lock Anod columns ===
    If DEBUG_MODE Then Debug.Print ">>> locking columns a, b, c"
    Application.StatusBar = "locking columns a, b, c..."
    DoEvents
    On Error Resume Next
    ws.Cells.Locked = True
    ws.Columns("A:C").Locked = True
    ws.Range(ws.Cells(1, START_COL), ws.Cells(1, MAX_COL)).EntireColumn.Locked = False
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> column lock error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> column locking completed"
    End If
    On Error GoTo 0

    ' === unlock Anod named ranges ===
    Dim nm As name
    For Each nm In ws.names
        On Error Resume Next
        nm.RefersToRange.Locked = False
        On Error GoTo 0
    Next nm
    If DEBUG_MODE Then Debug.Print ">>> Anod named ranges unlocked"

    ' === decimal separator ===
    If DEBUG_MODE Then Debug.Print ">>> checking decimal separator"
    Application.StatusBar = "checking decimal separator..."
    DoEvents
    On Error Resume Next
    Call CheckAndSetDecimalSeparator
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> error in checkandsetdecimalseparator: " & Err.Description
        MsgBox "decimal separator check error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> decimal separator verified successfully"
    End If
    On Error GoTo 0

    ' === validation initialization ===
    If DEBUG_MODE Then Debug.Print ">>> validation initialization"
    Application.StatusBar = "initializing validation..."
    DoEvents
    On Error Resume Next

    Dim colBrInit As Long
    Dim pipeCountInit As Variant
    pipeCountInit = ws.Range("pipeCountCP").Value
    colBrInit = SafeNumericValue(pipeCountInit, 1)

    Dim colIdxInit As Long
    For colIdxInit = START_COL To START_COL + colBrInit - 1
        Call Module_ValidationLogic.RefreshValidationForColumn(ws, colIdxInit)
    Next colIdxInit

    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> validation initialization error: " & Err.Description
        MsgBox "validation initialization error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> validation initialization completed successfully"
    End If
    On Error GoTo 0

    ' === restore Anod protection ===
    If DEBUG_MODE Then Debug.Print ">>> restoring Anod sheet protection"
    On Error Resume Next
    ws.Protect Password:="ptm", UserInterfaceOnly:=True
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> failed to restore protection: " & Err.Description
        MsgBox "protection restore error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> protection restored"
    End If
    On Error GoTo 0

    ' === protect sheets STO, Guide, Pipe, ListCP, ListAG ===
    If DEBUG_MODE Then Debug.Print ">>> protecting sheets STO, Guide, Pipe, ListCP, ListAG"
    Dim wsSTO As Worksheet, wsGuide As Worksheet, wsPipe As Worksheet
    Dim wsListCP As Worksheet, wsListAG As Worksheet

    On Error Resume Next
    Set wsSTO = thisWorkbook.Worksheets(SHEET_STO)
    Set wsGuide = thisWorkbook.Worksheets(SHEET_GUIDE)
    Set wsPipe = thisWorkbook.Worksheets(SHEET_PIPE)
    Set wsListCP = thisWorkbook.Worksheets(SHEET_LIST_CP)
    Set wsListAG = thisWorkbook.Worksheets(SHEET_LIST_AG)
    On Error GoTo 0

    If Not wsSTO Is Nothing Then
        wsSTO.Unprotect Password:="ptm"
        wsSTO.Cells.Locked = True
        wsSTO.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                       AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
        If DEBUG_MODE Then Debug.Print ">>> STO sheet protected"
    End If

    If Not wsGuide Is Nothing Then
        wsGuide.Unprotect Password:="ptm"
        wsGuide.Cells.Locked = True
        wsGuide.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                         AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
        If DEBUG_MODE Then Debug.Print ">>> Guide sheet protected"
    End If

    If Not wsPipe Is Nothing Then
        wsPipe.Unprotect Password:="ptm"
        wsPipe.Cells.Locked = True

        ' Unlock Pipe named ranges
        Dim nmPipe As name
        For Each nmPipe In wsPipe.names
            On Error Resume Next
            nmPipe.RefersToRange.Locked = False
            On Error GoTo 0
        Next nmPipe

        wsPipe.Protect Password:="ptm", AllowInsertingRows:=False, AllowDeletingRows:=False, _
                        AllowSorting:=False, AllowFiltering:=False, UserInterfaceOnly:=True
        If DEBUG_MODE Then Debug.Print ">>> Pipe sheet protected"
    End If

    ' ListCP: allow editing values and inserting rows
    If Not wsListCP Is Nothing Then
        wsListCP.Unprotect Password:="ptm"
        wsListCP.Cells.Locked = False
        wsListCP.Protect Password:="ptm", AllowInsertingRows:=True, AllowDeletingRows:=False, _
                         AllowSorting:=False, AllowFiltering:=True, UserInterfaceOnly:=True
        If DEBUG_MODE Then Debug.Print ">>> ListCP sheet protected (editing and row insertion allowed)"
    End If

    ' ListAG: allow editing values and inserting rows
    If Not wsListAG Is Nothing Then
        wsListAG.Unprotect Password:="ptm"
        wsListAG.Cells.Locked = False
        wsListAG.Protect Password:="ptm", AllowInsertingRows:=True, AllowDeletingRows:=False, _
                         AllowSorting:=False, AllowFiltering:=True, UserInterfaceOnly:=True
        If DEBUG_MODE Then Debug.Print ">>> ListAG sheet protected (editing and row insertion allowed)"
    End If

    ' === recalculate and activate ===
    If DEBUG_MODE Then Debug.Print ">>> recalculate and activate sheet"
    On Error Resume Next
    ws.Calculate
    ws.Activate
    If Err.Number <> 0 Then
        If DEBUG_MODE Then Debug.Print ">>> calculate/activate error: " & Err.Description
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print ">>> recalculate and activate completed"
    End If
    On Error GoTo 0

    ' === restore parameters ===
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = oldEnableEvents

    If DEBUG_MODE Then Debug.Print ">>> workbook_open: finish"

    Exit Sub

CleanExit:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errTime As String

    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errTime = Format(Now, "yyyy-mm-dd hh:mm:ss")

    If DEBUG_MODE Then
        Debug.Print ">>> transitioning to cleanexit"
        Debug.Print ">>> error number: " & errNum
        Debug.Print ">>> description: " & errDesc
        Debug.Print ">>> source: " & errSource
    End If

    On Error Resume Next
    ws.Protect Password:="ptm", UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = oldEnableEvents

    If errNum <> 0 Then
        MsgBox "error in workbook_open: " & errDesc, vbCritical
    End If
End Sub
