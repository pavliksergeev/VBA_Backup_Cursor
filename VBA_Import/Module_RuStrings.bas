Attribute VB_Name = "Module_RuStrings"

Option Explicit

' Unicode strings from UTF-16 hex. Source stays ASCII, so a Mac-saved .xlsm
' opens on Windows without re-importing .bas/.cls.
' Example: Ru("0442 0435 0441 0442") -> Russian "test".
' Codes: 0410-042F = A-Ya, 0430-044F = a-ya, 0020 = space, 002E = period.
Public Function Ru(ByVal hexUtf16 As String) As String
    Dim parts() As String
    Dim i As Long
    Dim token As String
    Dim cp As Long
    Dim out As String
    Dim ch As String

    hexUtf16 = Trim$(hexUtf16)
    If Len(hexUtf16) = 0 Then
        Ru = ""
        Exit Function
    End If

    parts = Split(hexUtf16, " ")
    out = ""
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) = 0 Then GoTo NextTok
        cp = CLng("&H" & token)
        ch = CharFromCode(cp)
        out = out & ch
NextTok:
    Next i
    Ru = out
End Function

Private Function CharFromCode(ByVal cp As Long) As String
    Dim ch As String
    ch = ""
    On Error Resume Next
    ch = Application.WorksheetFunction.Unichar(cp)
    On Error GoTo 0
    If Len(ch) = 0 Then
        If cp >= 0 And cp <= 127 Then
            ch = Chr$(cp)
        Else
            ch = ChrW$(cp)
        End If
    End If
    CharFromCode = ch
End Function
