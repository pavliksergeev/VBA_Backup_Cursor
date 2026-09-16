Attribute VB_Name = "Module_ExportAllToTXT"

' ================================================================
' macro: ExportAllModulesToTXT
' purpose: export all project modules to a TXT file
' save to the Downloads folder if available
' supported: Windows 11, macOS (Apple Silicon / Intel)
' ================================================================
Sub ExportAllModulesToTXT()
    ' check VBProject access
    On Error Resume Next
    Dim testProj As Object
    Set testProj = thisWorkbook.VBProject
    If testProj Is Nothing Then
                ' нет доступа к объектной модели vba!
                ' no access to the object model vba!
                '  включите доверие в меню: файл - па
                '  включandте доверandе в меню: file - па
                ' раметры - центр управления безопасн
                ' раметры - центр управленandя безопасн
                ' остью - параметры макросов, отметьт
                ' остью - macro settings, отметьт
                ' е галочку 'доверять доступ к объект
                ' е галочку 'trust access к объект
                ' ной модели проектов vba'.
                ' ной моделand проектов vba'.
                ' ошибка доступа
                ' access error
                MsgBox Ru("043D 0435 0442 0020 0434 043E 0441 0442 0443 043F 0430 0020 043A 0020 043E 0431 044A 0435 043A 0442 043D 043E 0439 0020 043C 043E 0434 0435 043B 0438 0020 0056 0042 0041 0021") & _
            Ru("0020 0432 043A 043B 044E 0447 0438 0442 0435 0020 0434 043E 0432 0435 0440 0438 0435 0020 0432 0020 043C 0435 043D 044E 003A 0020 0444 0430 0439 043B 0020 002D 0020 043F 0430") & _
            Ru("0440 0430 043C 0435 0442 0440 044B 0020 002D 0020 0446 0435 043D 0442 0440 0020 0443 043F 0440 0430 0432 043B 0435 043D 0438 044F 0020 0431 0435 0437 043E 043F 0430 0441 043D") & _
            Ru("043E 0441 0442 044C 044E 0020 002D 0020 043F 0430 0440 0430 043C 0435 0442 0440 044B 0020 043C 0430 043A 0440 043E 0441 043E 0432 002C 0020 043E 0442 043C 0435 0442 044C 0442") & _
            Ru("0435 0020 0433 0430 043B 043E 0447 043A 0443 0020 0027 0434 043E 0432 0435 0440 044F 0442 044C 0020 0434 043E 0441 0442 0443 043F 0020 043A 0020 043E 0431 044A 0435 043A 0442") & _
            Ru("043D 043E 0439 0020 043C 043E 0434 0435 043B 0438 0020 043F 0440 043E 0435 043A 0442 043E 0432 0020 0056 0042 0041 0027 002E"), vbCritical, Ru("043E 0448 0438 0431 043A 0430 0020 0434 043E 0441 0442 0443 043F 0430")
        Exit Sub
    End If
    On Error GoTo 0

    ' --- resolve save folder (Downloads or Desktop) ---
    Dim folderPath As String
    Dim pathSep As String
    pathSep = Application.PathSeparator

    ' get the Downloads path
    folderPath = GetDownloadsPath()
    
    ' if Downloads is missing or not writable, use Desktop
    If Not FolderExists(folderPath) Or Not CanWriteToFolder(folderPath) Then
        folderPath = GetDesktopPath()
    End If

    ' if Desktop is also unavailable, fall back to the current folder
    If Not FolderExists(folderPath) Or Not CanWriteToFolder(folderPath) Then
        folderPath = CurDir
    End If

    ' --- build the file name ---
    Dim filePath As String
    filePath = folderPath & pathSep & "AllModules_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"

    ' --- open the file with error handling ---
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo FileError
    Open filePath For Output As #fileNum
    On Error GoTo 0

    ' --- header ---
    Print #fileNum, "================================================================"
    ' все модули проекта
    ' all project modules
    Print #fileNum, Ru("0432 0441 0435 0020 043C 043E 0434 0443 043B 0438 0020 043F 0440 043E 0435 043A 0442 0430")
    Print #fileNum, "================================================================"
    ' дата экспорта: 
    ' export date: 
    Print #fileNum, Ru("0434 0430 0442 0430 0020 044D 043A 0441 043F 043E 0440 0442 0430 003A 0020") & Now
    ' файл: 
    ' file: 
    Print #fileNum, Ru("0444 0430 0439 043B 003A 0020") & thisWorkbook.fullName
    ' всего листов: 
    ' total sheets: 
    Print #fileNum, Ru("0432 0441 0435 0433 043E 0020 043B 0438 0441 0442 043E 0432 003A 0020") & thisWorkbook.Worksheets.count
    ' платформа: 
    ' platform: 
    Print #fileNum, Ru("043F 043B 0430 0442 0444 043E 0440 043C 0430 003A 0020") & Application.OperatingSystem
    Print #fileNum, "================================================================"
    Print #fileNum, ""

    ' --- loop over modules ---
    Dim vbComp As Object
    Dim count As Long, totalLines As Long, line As Long
    Dim errorMsg As String
    count = 0
    totalLines = 0

    Application.ScreenUpdating = False
    Application.StatusBar = "exporting modules..."

    For Each vbComp In thisWorkbook.VBProject.VBComponents
        count = count + 1
        errorMsg = ""

        Print #fileNum, ""
        Print #fileNum, "================================================================"
        ' модуль: 
        ' module: 
        Print #fileNum, Ru("043C 043E 0434 0443 043B 044C 003A 0020") & vbComp.name
        ' тип: 
        ' type: 
        Print #fileNum, Ru("0442 0438 043F 003A 0020") & GetComponentType(vbComp.Type)
        Print #fileNum, "================================================================"
        Print #fileNum, ""

        On Error Resume Next
        Dim lineCount As Long
        lineCount = vbComp.CodeModule.CountOfLines
        If Err.Number <> 0 Then
            ' ошибка получения количества строк: 
            ' get error колandчества lines: 
            errorMsg = Ru("043E 0448 0438 0431 043A 0430 0020 043F 043E 043B 0443 0447 0435 043D 0438 044F 0020 043A 043E 043B 0438 0447 0435 0441 0442 0432 0430 0020 0441 0442 0440 043E 043A 003A 0020") & Err.Description
            Err.Clear
            Print #fileNum, "' !!! " & errorMsg
            GoTo EndModule
        End If

        If lineCount > 0 Then
            For line = 1 To lineCount
                Dim codeLine As String
                codeLine = vbComp.CodeModule.Lines(line, 1)
                If Err.Number <> 0 Then
                    ' ошибка на строке 
                    ' error on linesе 
                    errorMsg = Ru("043E 0448 0438 0431 043A 0430 0020 043D 0430 0020 0441 0442 0440 043E 043A 0435 0020") & line & ": " & Err.Description
                    Err.Clear
                    Print #fileNum, "' !!! " & errorMsg
                    Exit For
                Else
                    Print #fileNum, codeLine
                    totalLines = totalLines + 1
                End If
            Next line
        Else
            ' ' (пустой модуль)
            ' ' (empty module)
            Print #fileNum, Ru("0027 0020 0028 043F 0443 0441 0442 043E 0439 0020 043C 043E 0434 0443 043B 044C 0029")
        End If

EndModule:
        On Error GoTo 0

        If errorMsg <> "" Then
            ' ' === модуль выгружен с ошибками ===
            ' ' === module exported with errors ===
            Print #fileNum, Ru("0027 0020 003D 003D 003D 0020 043C 043E 0434 0443 043B 044C 0020 0432 044B 0433 0440 0443 0436 0435 043D 0020 0441 0020 043E 0448 0438 0431 043A 0430 043C 0438 0020 003D 003D") & Ru("003D")
        Else
            ' === конец модуля 
            ' === end of module 
            Print #fileNum, Ru("003D 003D 003D 0020 043A 043E 043D 0435 0446 0020 043C 043E 0434 0443 043B 044F 0020") & vbComp.name & " ==="
        End If
        Print #fileNum, ""

        Application.StatusBar = "export: " & count & " modules, " & totalLines & " lines..."
        DoEvents
    Next vbComp

    ' --- footer ---
    Print #fileNum, ""
    Print #fileNum, "================================================================"
    ' конец выгрузки
    ' end of export
    Print #fileNum, Ru("043A 043E 043D 0435 0446 0020 0432 044B 0433 0440 0443 0437 043A 0438")
    Print #fileNum, "================================================================"
    ' всего модулей: 
    ' total modules: 
    Print #fileNum, Ru("0432 0441 0435 0433 043E 0020 043C 043E 0434 0443 043B 0435 0439 003A 0020") & count
    ' всего строк кода: 
    ' total code lines: 
    Print #fileNum, Ru("0432 0441 0435 0433 043E 0020 0441 0442 0440 043E 043A 0020 043A 043E 0434 0430 003A 0020") & totalLines
    ' дата: 
    ' date: 
    Print #fileNum, Ru("0434 0430 0442 0430 003A 0020") & Now
    Print #fileNum, "================================================================"

    Close #fileNum

    Application.StatusBar = "export finished!"
    Application.ScreenUpdating = True

    ' --- message ---
    Dim msg As String
    ' все модули выгружены!
    ' all modules exported!
    msg = Ru("0432 0441 0435 0020 043C 043E 0434 0443 043B 0438 0020 0432 044B 0433 0440 0443 0436 0435 043D 044B 0021") & vbCrLf & vbCrLf
    ' файл: 
    ' file: 
    msg = msg & Ru("0444 0430 0439 043B 003A 0020") & filePath & vbCrLf
    ' модулей: 
    ' modules: 
    msg = msg & Ru("043C 043E 0434 0443 043B 0435 0439 003A 0020") & count & vbCrLf
    ' строк кода: 
    ' code lines: 
    msg = msg & Ru("0441 0442 0440 043E 043A 0020 043A 043E 0434 0430 003A 0020") & totalLines & vbCrLf & vbCrLf
    ' файл сохранён в папке:
    ' file saved in folder:
    msg = msg & Ru("0444 0430 0439 043B 0020 0441 043E 0445 0440 0430 043D 0451 043D 0020 0432 0020 043F 0430 043F 043A 0435 003A") & vbCrLf & folderPath

    ' выгрузка завершена
    ' export finished
    MsgBox msg, vbInformation, Ru("0432 044B 0433 0440 0443 0437 043A 0430 0020 0437 0430 0432 0435 0440 0448 0435 043D 0430")
    Exit Sub

FileError:
    ' не удалось создать файл: 
    ' failed to create file: 
    ' ошибка записи
    ' write error
    MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 0441 043E 0437 0434 0430 0442 044C 0020 0444 0430 0439 043B 003A 0020") & filePath & vbCrLf & Err.Description, vbCritical, Ru("043E 0448 0438 0431 043A 0430 0020 0437 0430 043F 0438 0441 0438")
    Resume CleanExit

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' ================================================================
' helper functions
' ================================================================

Private Function GetComponentType(ByVal compType As Long) As String
    Select Case compType
        ' стандартный модуль
        ' standard module
        Case 1:   GetComponentType = Ru("0441 0442 0430 043D 0434 0430 0440 0442 043D 044B 0439 0020 043C 043E 0434 0443 043B 044C")
        ' класс-модуль
        ' class module
        Case 2:   GetComponentType = Ru("043A 043B 0430 0441 0441 002D 043C 043E 0434 0443 043B 044C")
        ' форма
        ' form
        Case 3:   GetComponentType = Ru("0444 043E 0440 043C 0430")
        ' лист/книга
        ' sheet/workbook
        Case 100: GetComponentType = Ru("043B 0438 0441 0442 002F 043A 043D 0438 0433 0430")
        ' тип 
        ' type 
        Case Else: GetComponentType = Ru("0442 0438 043F 0020") & compType
    End Select
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Function GetDesktopPath() As String
    If InStr(Application.OperatingSystem, "Mac") > 0 Then
        GetDesktopPath = Environ("HOME") & "/Desktop"
    Else
        GetDesktopPath = Environ("USERPROFILE") & "\Desktop"
    End If
End Function

Private Function GetDownloadsPath() As String
    If InStr(Application.OperatingSystem, "Mac") > 0 Then
        GetDownloadsPath = Environ("HOME") & "/Downloads"
    Else
        GetDownloadsPath = Environ("USERPROFILE") & "\Downloads"
    End If
End Function

Private Function CanWriteToFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim testFile As String
    testFile = folderPath & Application.PathSeparator & "test_write.tmp"
    Open testFile For Output As #1
    If Err.Number = 0 Then
        Close #1
        Kill testFile
        CanWriteToFolder = True
    Else
        CanWriteToFolder = False
        Err.Clear
    End If
    On Error GoTo 0
End Function
