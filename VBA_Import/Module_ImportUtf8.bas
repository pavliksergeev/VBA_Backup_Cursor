Attribute VB_Name = "Module_ImportUtf8"

Option Explicit

' Replace module source from UTF-8 .bas/.cls (Excel for Mac 16.59).
' File -> Import File cannot decode Russian. Do not remove modules first.
'
' Mac VBE ChrW of punctuation (em-dash etc.) becomes junk: bullets and euro.
' Run StripMacVbeGarbage on the open workbook to remove that junk.

Public Sub StripMacVbeGarbage()
    Dim comp As Object
    Dim cm As Object
    Dim i As Long
    Dim s As String
    Dim cleaned As String
    Dim nFix As Long
    Dim nDel As Long
    Dim prev As String

    On Error GoTo Fail

    For Each comp In thisWorkbook.VBProject.VBComponents
        If StrComp(comp.name, "Module_ImportUtf8", vbTextCompare) = 0 Then GoTo NextComp
        Set cm = comp.CodeModule
        If cm.CountOfLines < 1 Then GoTo NextComp

        For i = cm.CountOfLines To 1 Step -1
            s = cm.Lines(i, 1)
            cleaned = StripGarbageChars(s)
            If IsBlankOrGarbageLine(s) Then
                If i > 1 Then
                    prev = RTrim$(StripGarbageChars(cm.Lines(i - 1, 1)))
                    If Len(prev) > 0 Then
                        If Right$(prev, 1) = "_" Then
                            cm.DeleteLines i
                            nDel = nDel + 1
                            GoTo NextLine
                        End If
                    End If
                End If
                If ContainsGarbage(s) Then
                    cm.DeleteLines i
                    nDel = nDel + 1
                End If
            ElseIf cleaned <> s Then
                cm.ReplaceLine i, cleaned
                nFix = nFix + 1
            End If
NextLine:
        Next i
NextComp:
    Next comp

    MsgBox "StripMacVbeGarbage" & vbCrLf & _
           "Cleaned lines: " & nFix & vbCrLf & _
           "Deleted lines: " & nDel & vbCrLf & vbCrLf & _
           "Then Debug -> Compile VBAProject.", vbInformation
    Exit Sub

Fail:
    MsgBox "StripMacVbeGarbage failed: " & Err.Description, vbCritical
End Sub

Public Sub ImportUtf8Modules()
    Dim folderPath As String
    Dim names As Variant
    Dim i As Long
    Dim fileName As String
    Dim fullPath As String
    Dim okCount As Long
    Dim missCount As Long
    Dim errList As String
    Dim sample As String
    Dim picked As Variant

    On Error GoTo Fail

    names = ExpectedFileNames()

    folderPath = FindReadableImportFolder()
    If Len(folderPath) = 0 Then
        picked = Application.GetOpenFilename( _
            FileFilter:="All Files,*.*", _
            Title:="VBA_Import_Windows: Select All (Cmd+A)", _
            MultiSelect:=True)
        If VarType(picked) = vbBoolean Then
            MsgBox "Cancelled. Put the .xlsm next to folder VBA_Import_Windows and retry, or Select All files.", vbExclamation
            Exit Sub
        End If
        Application.StatusBar = "UTF-8 import..."
        If IsArray(picked) Then
            folderPath = FolderOfPath(CStr(picked(LBound(picked))))
            For i = LBound(picked) To UBound(picked)
                fileName = CStr(picked(i))
                On Error Resume Next
                ImportOneUtf8File fileName
                If Err.Number <> 0 Then
                    missCount = missCount + 1
                    errList = errList & fileName & ": " & Left$(Err.Description, 60) & vbCrLf
                    Err.Clear
                Else
                    okCount = okCount + 1
                End If
                On Error GoTo Fail
            Next i
        Else
            folderPath = FolderOfPath(CStr(picked))
            On Error Resume Next
            ImportOneUtf8File CStr(picked)
            If Err.Number <> 0 Then
                missCount = 1
                errList = Err.Description
            Else
                okCount = 1
            End If
            On Error GoTo Fail
        End If
        GoTo ShowResult
    End If

    Application.StatusBar = "UTF-8 import..."
    For i = LBound(names) To UBound(names)
        fileName = CStr(names(i))
        fullPath = JoinPath(folderPath, fileName)
        On Error Resume Next
        ImportOneUtf8File fullPath
        If Err.Number <> 0 Then
            missCount = missCount + 1
            errList = errList & fileName & ": " & Left$(Err.Description, 60) & vbCrLf
            Err.Clear
        Else
            okCount = okCount + 1
        End If
        On Error GoTo Fail
        Application.StatusBar = "UTF-8 import " & CStr(i - LBound(names) + 1)
    Next i

ShowResult:

    Application.StatusBar = False
    sample = CheckRussianSample("Module_RunAllTests")

    MsgBox "Folder: " & folderPath & vbCrLf & _
           "Updated: " & okCount & vbCrLf & _
           "Errors: " & missCount & vbCrLf & vbCrLf & _
           sample & vbCrLf & vbCrLf & errList, _
           IIf(missCount = 0, vbInformation, vbExclamation), _
           "ImportUtf8Modules"
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "ImportUtf8Modules failed: " & Err.Description, vbCritical
End Sub

Private Function ExpectedFileNames() As Variant
    ExpectedFileNames = Array( _
        "Module_RuStrings.bas", _
        "Module_AGData.bas", "Module_CP.bas", "Module_Constants.bas", _
        "Module_DataCleanup.bas", "Module_DebugHelpers.bas", "Module_DecimalSeparator.bas", _
        "Module_ExportAllToTXT.bas", "Module_FixCyrillic.bas", "Module_Helpers.bas", _
        "Module_Logging.bas", "Module_NamedRanges.bas", "Module_PestorePipeNames.bas", _
        "Module_Research.bas", "Module_RestoreAnodNames.bas", "Module_RunAllTests.bas", _
        "Module_Sync.bas", "Module_ValidationLists.bas", "Module_ValidationLogic.bas", _
        "Module_Visual.bas", "Module_calcAnod.bas", "Module_calcCP.bas", "Module_calcPipe.bas", _
        "wsAnod.cls", "wsPipe.cls", "thisWorkbook.cls")
End Function

Private Function FindReadableImportFolder() As String
    Dim cands As Variant
    Dim i As Long
    Dim probe As String
    Dim wbPath As String
    Dim homePath As String

    On Error Resume Next
    wbPath = thisWorkbook.Path
    homePath = Environ("HOME")
    On Error GoTo 0

    cands = Array( _
        JoinPath(wbPath, "VBA_Import_Windows"), _
        wbPath, _
        homePath & "/Downloads/VBA_Backup_Cursor/VBA_Import_Windows", _
        UnsandboxDownloadsPath(homePath & "/Library/Containers/com.microsoft.Excel/Data/Downloads/VBA_Backup_Cursor/VBA_Import_Windows"))

    For i = LBound(cands) To UBound(cands)
        probe = CStr(cands(i))
        If Len(probe) > 0 Then
            If CanOpenFile(JoinPath(probe, "Module_CP.bas")) Then
                FindReadableImportFolder = probe
                Exit Function
            End If
        End If
    Next i
End Function

Private Function UnsandboxDownloadsPath(ByVal p As String) As String
    Dim token As String
    Dim pos As Long
    token = "/Library/Containers/com.microsoft.Excel/Data/Downloads"
    pos = InStr(p, token)
    If pos > 0 Then
        UnsandboxDownloadsPath = Left$(p, pos - 1) & "/Downloads" & Mid$(p, pos + Len(token))
    Else
        UnsandboxDownloadsPath = p
    End If
End Function

Private Function CanOpenFile(ByVal fullPath As String) As Boolean
    Dim fileNum As Integer
    On Error Resume Next
    fileNum = FreeFile
    Open fullPath For Binary Access Read As #fileNum
    If Err.Number = 0 Then
        Close #fileNum
        CanOpenFile = True
    Else
        Err.Clear
        CanOpenFile = False
    End If
    On Error GoTo 0
End Function

Private Function FolderOfPath(ByVal fullPath As String) As String
    Dim pos As Long
    pos = InStrRev(fullPath, "/")
    If pos = 0 Then pos = InStrRev(fullPath, ":")
    If pos = 0 Then pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        FolderOfPath = Left$(fullPath, pos - 1)
    Else
        FolderOfPath = fullPath
    End If
End Function

Private Sub ImportOneUtf8File(ByVal fullPath As String)
    Dim raw As String
    Dim compName As String
    Dim comp As Object
    Dim cm As Object

    If StrComp(ComponentNameFromPath(fullPath), "Module_ImportUtf8", vbTextCompare) = 0 Then
        Exit Sub
    End If

    raw = ReadUtf8Binary(fullPath)
    raw = StripVbaExportHeader(raw)
    If Len(Trim$(raw)) = 0 Then Err.Raise vbObjectError + 2, , "empty file"

    compName = ComponentNameFromPath(fullPath)
    Set comp = FindComponent(compName)
    If comp Is Nothing Then
        Err.Raise vbObjectError + 3, , "module not found (do not delete it): " & compName
    End If

    Set cm = comp.CodeModule
    PutCode cm, raw
End Sub

Private Sub PutCode(ByVal cm As Object, ByVal raw As String)
    Dim lines() As String
    Dim i As Long
    Dim lineText As String

    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    lines = Split(raw, vbLf)

    Do While cm.CountOfLines > 0
        If cm.CountOfLines > 200 Then
            cm.DeleteLines 1, 200
        Else
            cm.DeleteLines 1, cm.CountOfLines
        End If
    Loop

    For i = LBound(lines) To UBound(lines)
        lineText = StripGarbageChars(lines(i))
        If Len(Trim$(lineText)) = 0 Then GoTo NextLine
        If cm.CountOfLines = 0 Then
            cm.InsertLines 1, lineText
        Else
            cm.InsertLines cm.CountOfLines + 1, lineText
        End If
NextLine:
    Next i
End Sub

Private Function StripGarbageChars(ByVal s As String) As String
    Dim i As Long
    Dim cp As Long
    Dim ch As String
    Dim out As String

    If Len(s) = 0 Then
        StripGarbageChars = ""
        Exit Function
    End If

    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + 65536
        If cp = 9 Then
            out = out & Chr$(9)
        ElseIf cp >= 32 And cp <= 126 Then
            out = out & ch
        ElseIf cp = 1025 Or cp = 1105 Then
            out = out & ch
        ElseIf cp >= 1040 And cp <= 1103 Then
            out = out & ch
        ElseIf cp = 160 Then
            out = out & " "
        ElseIf cp = 171 Or cp = 187 Then
            out = out & Chr$(34)
        ElseIf cp = 8211 Or cp = 8212 Or cp = 8722 Then
            out = out & "-"
        ElseIf cp = 8216 Or cp = 8217 Or cp = 8218 Then
            out = out & "'"
        ElseIf cp = 8220 Or cp = 8221 Or cp = 8222 Then
            out = out & Chr$(34)
        ElseIf cp = 8230 Then
            out = out & "..."
        End If
    Next i
    StripGarbageChars = out
End Function

Private Function ContainsGarbage(ByVal s As String) As Boolean
    Dim i As Long
    Dim cp As Long
    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp = 128 Or cp = 165 Or cp = 219 Then
            ContainsGarbage = True
            Exit Function
        End If
        If cp = 8226 Or cp = 8364 Or cp = 8212 Or cp = 8211 Then
            ContainsGarbage = True
            Exit Function
        End If
        If cp = 13 Or cp = 10 Or cp = 0 Then
            ContainsGarbage = True
            Exit Function
        End If
    Next i
End Function

Private Function IsBlankOrGarbageLine(ByVal s As String) As Boolean
    IsBlankOrGarbageLine = (Len(Trim$(StripGarbageChars(s))) = 0)
End Function

Private Function MacSafeChar(ByVal cp As Long) As String
    If cp < 0 Then
        MacSafeChar = ""
    ElseIf cp = 9 Then
        MacSafeChar = Chr$(9)
    ElseIf cp < 32 Then
        MacSafeChar = ""
    ElseIf cp <= 126 Then
        MacSafeChar = Chr$(cp)
    ElseIf cp = 1025 Or cp = 1105 Then
        MacSafeChar = ChrW$(cp)
    ElseIf cp >= 1040 And cp <= 1103 Then
        MacSafeChar = ChrW$(cp)
    ElseIf cp = 160 Then
        MacSafeChar = " "
    ElseIf cp = 171 Or cp = 187 Then
        MacSafeChar = Chr$(34)
    ElseIf cp = 8211 Or cp = 8212 Or cp = 8722 Then
        MacSafeChar = "-"
    ElseIf cp = 8216 Or cp = 8217 Or cp = 8218 Then
        MacSafeChar = "'"
    ElseIf cp = 8220 Or cp = 8221 Or cp = 8222 Then
        MacSafeChar = Chr$(34)
    ElseIf cp = 8226 Or cp = 183 Then
        MacSafeChar = "*"
    ElseIf cp = 8230 Then
        MacSafeChar = "..."
    ElseIf cp = 8364 Then
        MacSafeChar = "E"
    Else
        MacSafeChar = ""
    End If
End Function

Private Function ReadUtf8Binary(ByVal fullPath As String) As String
    Dim fileNum As Integer
    Dim n As Long
    Dim bytes() As Byte

    fileNum = FreeFile
    Open fullPath For Binary Access Read As #fileNum
    n = LOF(fileNum)
    If n <= 0 Then
        Close #fileNum
        Err.Raise vbObjectError + 4, , "empty file"
    End If
    ReDim bytes(1 To n)
    Get #fileNum, 1, bytes
    Close #fileNum
    ReadUtf8Binary = Utf8BytesToString(bytes)
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte) As String
    Dim i As Long
    Dim last As Long
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Dim cp As Long
    Dim out As String
    Dim parts() As String
    Dim p As Long
    Dim cap As Long

    i = LBound(bytes)
    last = UBound(bytes)
    cap = 512
    ReDim parts(1 To cap)
    p = 0

    If last - i >= 2 Then
        If (bytes(i) And &HFF) = &HEF And (bytes(i + 1) And &HFF) = &HBB And (bytes(i + 2) And &HFF) = &HBF Then
            i = i + 3
        End If
    End If

    Do While i <= last
        b0 = bytes(i) And &HFF
        If b0 < &H80 Then
            cp = b0
            i = i + 1
        ElseIf (b0 And &HE0) = &HC0 Then
            If i + 1 > last Then Exit Do
            b1 = bytes(i + 1) And &HFF
            cp = (b0 And &H1F) * 64 + (b1 And &H3F)
            i = i + 2
        ElseIf (b0 And &HF0) = &HE0 Then
            If i + 2 > last Then Exit Do
            b1 = bytes(i + 1) And &HFF
            b2 = bytes(i + 2) And &HFF
            cp = (b0 And &HF) * 4096 + (b1 And &H3F) * 64 + (b2 And &H3F)
            i = i + 3
        ElseIf (b0 And &HF8) = &HF0 Then
            If i + 3 > last Then Exit Do
            b1 = bytes(i + 1) And &HFF
            b2 = bytes(i + 2) And &HFF
            b3 = bytes(i + 3) And &HFF
            cp = (b0 And 7) * 262144 + (b1 And &H3F) * 4096 + (b2 And &H3F) * 64 + (b3 And &H3F)
            cp = cp - 65536
            p = p + 1
            If p > cap Then
                cap = cap * 2
                ReDim Preserve parts(1 To cap)
            End If
            parts(p) = "?"
            i = i + 4
            GoTo ContinueLoop
        Else
            cp = 63
            i = i + 1
        End If
        p = p + 1
        If p > cap Then
            cap = cap * 2
            ReDim Preserve parts(1 To cap)
        End If
        If cp < 0 Then cp = 0
        parts(p) = MacSafeChar(cp)
ContinueLoop:
    Loop

    If p = 0 Then
        Utf8BytesToString = ""
    Else
        ReDim Preserve parts(1 To p)
        Utf8BytesToString = Join(parts, "")
    End If
End Function

Private Function CheckRussianSample(ByVal compName As String) As String
    Dim comp As Object
    Dim t As String
    Dim n As Long
    Dim p As Long
    Dim ch As String
    Dim code As Long
    Set comp = FindComponent(compName)
    If comp Is Nothing Then
        CheckRussianSample = compName & ": module not found"
        Exit Function
    End If
    n = comp.CodeModule.CountOfLines
    If n < 1 Then
        CheckRussianSample = compName & ": empty"
        Exit Function
    End If
    If n > 250 Then n = 250
    t = comp.CodeModule.Lines(1, n)
    p = InStr(t, "MsgBox " & Chr$(34))
    If p = 0 Then
        CheckRussianSample = "Check: no MsgBox in first lines of " & compName
        Exit Function
    End If
    ch = Mid$(t, p + 8, 1)
    code = AscW(ch)
    If code >= 1040 And code <= 1103 Then
        CheckRussianSample = "Check: Russian OK in " & compName & " (AscW=" & CStr(code) & ")"
    Else
        CheckRussianSample = "Check: NOT Russian in " & compName & " (AscW=" & CStr(code) & "). Replace Module_ImportUtf8 and run again, folder VBA_Import_Windows."
    End If
End Function

Private Function FindComponent(ByVal compName As String) As Object
    Dim comp As Object
    For Each comp In thisWorkbook.VBProject.VBComponents
        If StrComp(comp.name, compName, vbTextCompare) = 0 Then
            Set FindComponent = comp
            Exit Function
        End If
    Next comp
End Function

Private Function ComponentNameFromPath(ByVal fullPath As String) As String
    Dim fileName As String
    Dim pos As Long
    fileName = fullPath
    pos = InStrRev(fileName, "/")
    If pos = 0 Then pos = InStrRev(fileName, ":")
    If pos = 0 Then pos = InStrRev(fileName, "\")
    If pos > 0 Then fileName = Mid$(fileName, pos + 1)
    pos = InStrRev(fileName, ".")
    If pos > 0 Then fileName = Left$(fileName, pos - 1)
    ComponentNameFromPath = fileName
End Function

Private Function JoinPath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim f As String
    f = folderPath
    Do While Len(f) > 0
        If Right$(f, 1) = "/" Or Right$(f, 1) = "\" Or Right$(f, 1) = ":" Then
            f = Left$(f, Len(f) - 1)
        Else
            Exit Do
        End If
    Loop
    If InStr(folderPath, "/") > 0 Then
        JoinPath = f & "/" & fileName
    ElseIf InStr(folderPath, ":") > 0 Then
        JoinPath = f & ":" & fileName
    Else
        JoinPath = f & Application.PathSeparator & fileName
    End If
End Function

Private Function StripVbaExportHeader(ByVal text As String) As String
    Dim lines() As String
    Dim i As Long
    Dim started As Boolean
    Dim out As String
    Dim line As String
    Dim s As String

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    lines = Split(text, vbLf)

    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        s = LTrim$(line)
        If Not started Then
            If Left$(s, 9) = "Attribute" Then GoTo NextLine
            If Left$(UCase$(s), 7) = "VERSION" Then GoTo NextLine
            If UCase$(s) = "BEGIN" Then GoTo NextLine
            If Left$(s, 8) = "MultiUse" Then GoTo NextLine
            If UCase$(s) = "END" And Len(s) = 3 Then GoTo NextLine
            started = True
        End If
        If Len(out) > 0 Then out = out & vbLf
        out = out & line
NextLine:
    Next i
    StripVbaExportHeader = out
End Function
