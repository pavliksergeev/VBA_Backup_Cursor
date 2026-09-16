Attribute VB_Name = "Module_FixCyrillic"

Option Explicit

Sub FixBrokenCyrillic()
    ' ================================================================
    ' single macro for Windows and Mac
    ' detects the OS and applies the matching table
    ' ================================================================

    ' --- detect the operating system ---
    Dim os As String
    os = Application.OperatingSystem
    Dim isWindows As Boolean, isMac As Boolean
    isWindows = (InStr(1, os, "Windows", vbTextCompare) > 0)
    isMac = (InStr(1, os, "Mac", vbTextCompare) > 0)

    If Not isWindows And Not isMac Then
        ' неизвестная операционная система: 
        ' unknown operating system: 
        ' макрос поддерживает только windows и macos.
        ' macro supports only windows and macos.
        MsgBox Ru("043D 0435 0438 0437 0432 0435 0441 0442 043D 0430 044F 0020 043E 043F 0435 0440 0430 0446 0438 043E 043D 043D 0430 044F 0020 0441 0438 0441 0442 0435 043C 0430 003A 0020") & os & vbCrLf & _
               Ru("043C 0430 043A 0440 043E 0441 0020 043F 043E 0434 0434 0435 0440 0436 0438 0432 0430 0435 0442 0020 0442 043E 043B 044C 043A 043E 0020 0077 0069 006E 0064 006F 0077 0073 0020") & Ru("0438 0020 006D 0061 0063 006F 0073 002E"), vbExclamation
        Exit Sub
    End If

    ' ================================================================
    ' create a backup of the current workbook
    ' ================================================================
    Dim backupPath As String
    Dim originalPath As String
    Dim answer As VbMsgBoxResult
    Dim backupSuccess As Boolean
    backupSuccess = False

    originalPath = thisWorkbook.fullName
    
    If originalPath <> "" Then
        Dim folder As String, nameBase As String, ext As String
        folder = Left(originalPath, InStrRev(originalPath, Application.PathSeparator))
        nameBase = Left(originalPath, InStrRev(originalPath, ".") - 1)
        ext = Mid(originalPath, InStrRev(originalPath, "."))
        backupPath = folder & nameBase & "_backup_" & Format(Now, "YYYYMMDD_HHMMSS") & ext
        
        On Error Resume Next
        thisWorkbook.SaveCopyAs backupPath
        If Err.Number = 0 Then
            backupSuccess = True
            ' резервная копия создана:
            ' backup copy создаon:
            ' резервирование
            ' backup
            MsgBox Ru("0440 0435 0437 0435 0440 0432 043D 0430 044F 0020 043A 043E 043F 0438 044F 0020 0441 043E 0437 0434 0430 043D 0430 003A") & vbCrLf & backupPath, vbInformation, Ru("0440 0435 0437 0435 0440 0432 0438 0440 043E 0432 0430 043D 0438 0435")
        Else
            ' не удалось создать резервную копию:
            ' failed to create резервную копandю:
            ' ошибка
            ' error
            MsgBox Ru("043D 0435 0020 0443 0434 0430 043B 043E 0441 044C 0020 0441 043E 0437 0434 0430 0442 044C 0020 0440 0435 0437 0435 0440 0432 043D 0443 044E 0020 043A 043E 043F 0438 044E 003A") & vbCrLf & Err.Description, vbExclamation, Ru("043E 0448 0438 0431 043A 0430")
            Err.Clear
            ' продолжить без резервной копии?
            ' continue without backup copy?
            ' подтверждение
            ' confirmation
            answer = MsgBox(Ru("043F 0440 043E 0434 043E 043B 0436 0438 0442 044C 0020 0431 0435 0437 0020 0440 0435 0437 0435 0440 0432 043D 043E 0439 0020 043A 043E 043F 0438 0438 003F"), vbYesNo + vbQuestion, Ru("043F 043E 0434 0442 0432 0435 0440 0436 0434 0435 043D 0438 0435"))
            If answer = vbNo Then Exit Sub
        End If
        On Error GoTo 0
    Else
        ' книга ещё не сохранена. резервная копия не создана.
        ' workbook is not saved yet. backup copy не создаon.
        ' рекомендуется сохранить файл перед запуском макроса.
        ' save the file first before running the macro.
        ' предупреждение
        ' warning
        MsgBox Ru("043A 043D 0438 0433 0430 0020 0435 0449 0451 0020 043D 0435 0020 0441 043E 0445 0440 0430 043D 0435 043D 0430 002E 0020 0440 0435 0437 0435 0440 0432 043D 0430 044F 0020 043A") & Ru("043E 043F 0438 044F 0020 043D 0435 0020 0441 043E 0437 0434 0430 043D 0430 002E") & vbCrLf & _
               Ru("0440 0435 043A 043E 043C 0435 043D 0434 0443 0435 0442 0441 044F 0020 0441 043E 0445 0440 0430 043D 0438 0442 044C 0020 0444 0430 0439 043B 0020 043F 0435 0440 0435 0434 0020") & Ru("0437 0430 043F 0443 0441 043A 043E 043C 0020 043C 0430 043A 0440 043E 0441 0430 002E"), vbExclamation, Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435")
        ' продолжить без резервной копии?
        ' continue without backup copy?
        ' подтверждение
        ' confirmation
        answer = MsgBox(Ru("043F 0440 043E 0434 043E 043B 0436 0438 0442 044C 0020 0431 0435 0437 0020 0440 0435 0437 0435 0440 0432 043D 043E 0439 0020 043A 043E 043F 0438 0438 003F"), vbYesNo + vbQuestion, Ru("043F 043E 0434 0442 0432 0435 0440 0436 0434 0435 043D 0438 0435"))
        If answer = vbNo Then Exit Sub
    End If

    ' ================================================================
    ' main part of the macro
    ' ================================================================
    Const SKIP_MODULE As String = "Module_FixCyrillic"   ' this module's name
    Dim vbComp As Object
    Dim codeMod As Object
    Dim lineNum As Long
    Dim lineText As String
    Dim newLine As String
    Dim totalFixed As Long
    Dim unknownChars As New Collection          ' collects unknown characters
    Dim i As Long, j As Long
    
    ' ========== mapping table (depends on OS) ==========
    Dim brokenChars As Variant
    Dim correctChars As Variant
    
    If isWindows Then
        ' --- Windows table (Mac OS Roman -> windows-1251) ---
        brokenChars = Array( _
            ' њ
            ' ћ
            ' ќ
            ' љ
            ' џ
            ' ѓ
            ' ђ
            ' ѕ
            ' і
            ' ј
            ' ї
            ' ў
            Ru("040A"), Ru("040B"), "„", "‘", "’", "“", "”", Ru("040C"), Ru("0409"), Ru("040F"), Ru("0403"), Ru("0402"), _
            Ru("0405"), Ru("0406"), Ru("0408"), Ru("0407"), "?", Ru("040E"), _
            Ru("0453"), Ru("045F") _
        )
        correctChars = Array( _
            ' м
            ' о
            ' д
            ' с
            ' т
            ' р
            ' у
            ' н
            ' л
            ' п
            ' г
            ' б
            ' ѕ
            ' і
            ' ј
            ' ї
            ' ў
            ' ы
            Ru("043C"), Ru("043E"), Ru("0434"), Ru("0441"), Ru("0442"), Ru("0440"), Ru("0443"), Ru("043D"), Ru("043B"), Ru("043F"), Ru("0433"), Ru("0431"), _
            Ru("0455"), Ru("0456"), Ru("0458"), Ru("0457"), "?", Ru("045E"), _
            Ru("0433"), Ru("044B") _
        )
    Else ' isMac
        ' --- Mac table (windows-1251 -> Mac OS Roman) ---
        ' this table is approximate; after the first run you will see the real characters
        ' and can extend it
        brokenChars = Array( _
            "?", "?", "?", "?", "?", "?", "?", "?", "?", "?", _
            "?", "?", "?", "?", "?", "?", "?", "?", _
            "?", "?", "™", "?", "?", "?", "?" _
        )
        correctChars = Array( _
            ' а
            ' е
            ' о
            ' у
            ' б
            ' и
            ' and
            ' д
            ' п
            ' с
            ' т
            ' р
            Ru("0430"), Ru("0435"), Ru("043E"), Ru("0443"), Ru("0431"), Ru("0430"), Ru("0435"), Ru("0438"), Ru("043E"), Ru("0443"), _
            Ru("0435"), Ru("0430"), Ru("043E"), Ru("0430"), Ru("043E"), Ru("0430"), Ru("0434"), Ru("043F"), _
            Ru("0441"), Ru("0442"), Ru("0440"), Ru("0441"), Ru("0442"), Ru("0440"), Ru("0443") _
        )
    End If
    
    ' check array lengths
    If UBound(brokenChars) <> UBound(correctChars) Then
        ' ошибка: количество элементов в массивах не совпадает!
        ' error: element count в arrayах does not match!
        MsgBox Ru("043E 0448 0438 0431 043A 0430 003A 0020 043A 043E 043B 0438 0447 0435 0441 0442 0432 043E 0020 044D 043B 0435 043C 0435 043D 0442 043E 0432 0020 0432 0020 043C 0430 0441 0441") & Ru("0438 0432 0430 0445 0020 043D 0435 0020 0441 043E 0432 043F 0430 0434 0430 0435 0442 0021"), vbCritical
        Exit Sub
    End If
    ' ==================================================================================
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For Each vbComp In thisWorkbook.VBProject.VBComponents
        If vbComp.name = SKIP_MODULE Then
            ' пропущен модуль: 
            ' skipped module: 
            Debug.Print Ru("043F 0440 043E 043F 0443 0449 0435 043D 0020 043C 043E 0434 0443 043B 044C 003A 0020") & vbComp.name
            GoTo NextModule
        End If
        
        Set codeMod = vbComp.CodeModule
        If codeMod.CountOfLines > 0 Then
            For lineNum = 1 To codeMod.CountOfLines
                lineText = codeMod.Lines(lineNum, 1)
                ' step 1: fix broken characters
                newLine = ReplaceBrokenChars(lineText, brokenChars, correctChars, unknownChars)
                ' step 2: lowercase all Russian letters
                newLine = ConvertRussianToLower(newLine)
                If newLine <> lineText Then
                    codeMod.ReplaceLine lineNum, newLine
                    totalFixed = totalFixed + 1
                End If
            Next lineNum
        End If
NextModule:
    Next vbComp
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' --- print results ---
    Dim msg As String
    ' исправлено строк: 
    ' lines fixed: 
    msg = Ru("0438 0441 043F 0440 0430 0432 043B 0435 043D 043E 0020 0441 0442 0440 043E 043A 003A 0020") & totalFixed & vbCrLf
    
    If unknownChars.count > 0 Then
        ' найдены неизвестные символы (их нужно добавить в таблицу):
        ' unknown characters found (add them to the table):
        msg = msg & Ru("043D 0430 0439 0434 0435 043D 044B 0020 043D 0435 0438 0437 0432 0435 0441 0442 043D 044B 0435 0020 0441 0438 043C 0432 043E 043B 044B 0020 0028 0438 0445 0020 043D 0443 0436") & Ru("043D 043E 0020 0434 043E 0431 0430 0432 0438 0442 044C 0020 0432 0020 0442 0430 0431 043B 0438 0446 0443 0029 003A") & vbCrLf
        Dim ch As Variant
        For Each ch In unknownChars
            ' ' (код 
            ' ' (code 
            msg = msg & "  '" & ch & Ru("0027 0020 0028 043A 043E 0434 0020") & AscW(ch) & ")" & vbCrLf
        Next ch
        ' список также выведен в immediate window (ctrl+g).
        ' the list is also printed в immediate window (ctrl+g).
        msg = msg & vbCrLf & Ru("0441 043F 0438 0441 043E 043A 0020 0442 0430 043A 0436 0435 0020 0432 044B 0432 0435 0434 0435 043D 0020 0432 0020 0069 006D 006D 0065 0064 0069 0061 0074 0065 0020 0077 0069") & Ru("006E 0064 006F 0077 0020 0028 0063 0074 0072 006C 002B 0067 0029 002E")
        
        ' === неизвестные символы (
        ' === unknown characters (
        Debug.Print Ru("003D 003D 003D 0020 043D 0435 0438 0437 0432 0435 0441 0442 043D 044B 0435 0020 0441 0438 043C 0432 043E 043B 044B 0020 0028") & IIf(isWindows, "windows", "mac") & ") ==="
        For Each ch In unknownChars
            ' символ: '
            ' character: '
            ' '  код: 
            ' '  code: 
            Debug.Print Ru("0441 0438 043C 0432 043E 043B 003A 0020 0027") & ch & Ru("0027 0020 0020 043A 043E 0434 003A 0020") & AscW(ch)
        Next ch
        Debug.Print "========================================"
    Else
        ' все битые символы успешно заменены!
        ' все broken characters replaced successfully!
        msg = msg & Ru("0432 0441 0435 0020 0431 0438 0442 044B 0435 0020 0441 0438 043C 0432 043E 043B 044B 0020 0443 0441 043F 0435 0448 043D 043E 0020 0437 0430 043C 0435 043D 0435 043D 044B 0021")
    End If
    
    If backupSuccess Then
        ' резервная копия сохранена: 
        ' backup copy сохранеon: 
        msg = msg & vbCrLf & Ru("0440 0435 0437 0435 0440 0432 043D 0430 044F 0020 043A 043E 043F 0438 044F 0020 0441 043E 0445 0440 0430 043D 0435 043D 0430 003A 0020") & backupPath
    End If
    
    ' восстановление завершено
    ' restore finished
    MsgBox msg, vbInformation, Ru("0432 043E 0441 0441 0442 0430 043D 043E 0432 043B 0435 043D 0438 0435 0020 0437 0430 0432 0435 0440 0448 0435 043D 043E")
End Sub

' ================================================================
' replace broken characters one by one
' ================================================================
Private Function ReplaceBrokenChars(ByVal text As String, _
    ByVal broken As Variant, ByVal correct As Variant, ByVal unknown As Collection) As String
    
    Dim result As String
    Dim i As Long, j As Long
    Dim ch As String
    Dim code As Long
    Dim replaced As Boolean
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        code = AscW(ch)
        replaced = False
        
        ' look up the character in the mapping table
        For j = LBound(broken) To UBound(broken)
            If ch = broken(j) Then
                result = result & correct(j)
                replaced = True
                Exit For
            End If
        Next j
        
        If Not replaced Then
            ' if the character is not found and is neither ASCII nor a Russian letter,
            ' add it to the unknown-character collection (if not already there)
            If code > 127 And Not IsRussianChar(code) Then
                Dim exists As Boolean
                exists = False
                Dim item As Variant
                For Each item In unknown
                    If item = ch Then
                        exists = True
                        Exit For
                    End If
                Next item
                If Not exists Then
                    unknown.Add ch
                End If
            End If
            result = result & ch   ' leave as is
        End If
    Next i
    
    ReplaceBrokenChars = result
End Function

' ================================================================
' check whether the character is a Russian letter (by code)
' ================================================================
Private Function IsRussianChar(ByVal code As Long) As Boolean
    ' Unicode ranges for Cyrillic:
    ' A-Ya: 1040-1071, a-ya: 1072-1103, Yo: 1025, yo: 1105
    If (code >= 1040 And code <= 1103) Or code = 1025 Or code = 1105 Then
        IsRussianChar = True
    Else
        IsRussianChar = False
    End If
End Function

' ================================================================
' lowercase Russian letters
' (leave English and other characters unchanged)
' ================================================================
Private Function ConvertRussianToLower(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        code = AscW(ch)
        
        ' A-Ya (1040-1071) -> a-ya (code + 32)
        If code >= 1040 And code <= 1071 Then
            result = result & ChrW(code + 32)
        ' Yo (1025) -> yo (1105)
        ElseIf code = 1025 Then
            result = result & ChrW(1105)
        ' leave all other characters as is
        Else
            result = result & ch
        End If
    Next i
    
    ConvertRussianToLower = result
End Function
