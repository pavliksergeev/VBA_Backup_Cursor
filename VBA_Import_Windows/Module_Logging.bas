Attribute VB_Name = "Module_Logging"

' ================================================================
' Module_Logging
' ================================================================
Option Explicit

Private logFile As String
Private logEnabled As Boolean

' initialize the logger
Public Sub InitLogger(Optional enableLog As Boolean = True)
    logEnabled = enableLog
    If logEnabled Then
        logFile = thisWorkbook.Path & "\Anod_Log_" & Format(Now, "YYYYMMDD") & ".txt"
        LogMessage "========================================"
        ' лог запущен: 
        ' log started: 
        LogMessage Ru("043B 043E 0433 0020 0437 0430 043F 0443 0449 0435 043D 003A 0020") & Now
        LogMessage "========================================"
    End If
End Sub

' write a message to the log
Public Sub LogMessage(message As String)
    If Not logEnabled Then Exit Sub
    
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    Print #fileNum, Format(Now, "HH:MM:SS") & " - " & message
    Close #fileNum
    
    ' also print to the Immediate Window for debugging
    Debug.Print Format(Now, "HH:MM:SS") & " - " & message
    On Error GoTo 0
End Sub

' write an error to the log
Public Sub LogError(procName As String, errNumber As Long, errDescription As String)
    ' ошибка в 
    ' error in 
    LogMessage Ru("043E 0448 0438 0431 043A 0430 0020 0432 0020") & procName & ": #" & errNumber & " - " & errDescription
End Sub

' write a warning
Public Sub LogWarning(message As String)
    ' предупреждение: 
    ' warning: 
    LogMessage Ru("043F 0440 0435 0434 0443 043F 0440 0435 0436 0434 0435 043D 0438 0435 003A 0020") & message
End Sub

' write an info message
Public Sub LogInfo(message As String)
    ' инфо: 
    ' info: 
    LogMessage Ru("0438 043D 0444 043E 003A 0020") & message
End Sub

' clear the log
Public Sub ClearLog()
    If logEnabled Then
        On Error Resume Next
        Kill logFile
        On Error GoTo 0
        ' лог очищен
        ' log cleared
        LogMessage Ru("043B 043E 0433 0020 043E 0447 0438 0449 0435 043D")
    End If
End Sub
