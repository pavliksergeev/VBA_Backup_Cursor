Attribute VB_Name = "Module_DecimalSeparator"

Option Explicit

' True if this workbook turned off system separators for the Excel session.
Public gDecimalSepOverridden As Boolean

' ================================================================
' Set Excel decimal separator to "." when the OS uses comma.
' Inform only; no Yes/No. VBA CStr() still uses the OS locale;
' number lists must use NumberToValidationText / Str$ instead of CStr.
' ================================================================
Public Sub CheckAndSetDecimalSeparator()
    Dim currentSep As String
    Dim targetSep As String

    currentSep = Application.International(xlDecimalSeparator)
    targetSep = "."

    If currentSep = targetSep Then Exit Sub

    On Error Resume Next
    Application.UseSystemSeparators = False
    Application.DecimalSeparator = targetSep
    Application.ThousandsSeparator = ","
    If Err.Number = 0 Then
        gDecimalSepOverridden = True
        MsgBox "Decimal separator changed to '.' (dot)." & vbCrLf & vbCrLf & _
               "This applies to all open Excel workbooks for this session." & vbCrLf & _
               "System settings are not changed." & vbCrLf & vbCrLf & _
               "The separator returns to the system default when this workbook is closed.", _
               vbInformation, "Decimal Separator"
    Else
        gDecimalSepOverridden = False
        MsgBox "decimal separator check error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub RestoreDecimalSeparator()
    On Error Resume Next
    If gDecimalSepOverridden Then
        Application.UseSystemSeparators = True
        gDecimalSepOverridden = False
    End If
    Err.Clear
End Sub

' Invariant decimal point for xlValidateList Formula1 items.
' CStr(0.325) is "0,325" on Russian Windows and splits the comma-delimited list.
Public Function NumberToValidationText(ByVal v As Variant) As String
    On Error GoTo Fallback
    If IsError(v) Then
        NumberToValidationText = ""
        Exit Function
    End If
    If IsEmpty(v) Then
        NumberToValidationText = ""
        Exit Function
    End If
    If IsNumeric(v) Then
        NumberToValidationText = WithLeadingZero(Trim$(Str$(CDbl(v))))
        Exit Function
    End If
Fallback:
    On Error Resume Next
    NumberToValidationText = CStr(v)
    If Err.Number <> 0 Then
        NumberToValidationText = ""
        Err.Clear
    End If
End Function

' Str$(0.325) is " .325" on some VBA hosts; dropdowns then show .325 not 0.325.
Private Function WithLeadingZero(ByVal s As String) As String
    If Left$(s, 1) = "." Then
        WithLeadingZero = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        WithLeadingZero = "-0" & Mid$(s, 2)
    Else
        WithLeadingZero = s
    End If
End Function
