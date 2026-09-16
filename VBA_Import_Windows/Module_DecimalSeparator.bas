Attribute VB_Name = "Module_DecimalSeparator"

Option Explicit

' ================================================================
' CHECK AND SET DECIMAL SEPARATOR
' ================================================================
Sub CheckAndSetDecimalSeparator()
    Dim currentSep As String
    Dim targetSep As String
    Dim answer As VbMsgBoxResult
    Dim msg As String
    currentSep = Application.International(xlDecimalSeparator)
    targetSep = "."

    If currentSep = targetSep Then Exit Sub

    msg = "Current decimal separator in Excel: '" & currentSep & "'" & vbCrLf & vbCrLf
    msg = msg & "For correct operation of this file, it is recommended to use '.' (dot)." & vbCrLf & vbCrLf
    msg = msg & "Change separator to '.' for all open workbooks?" & vbCrLf
    msg = msg & "(This will affect all Excel, but not system settings!)" & vbCrLf & vbCrLf
    msg = msg & "Click 'Yes' - change separator" & vbCrLf
    msg = msg & "Click 'No' - continue without changes"

    answer = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton1, "Excel Regional Settings")

    If answer = vbYes Then
        Application.UseSystemSeparators = False
        Application.DecimalSeparator = targetSep
        Application.ThousandsSeparator = ","
        MsgBox "Separator changed to: '" & targetSep & "'" & vbCrLf & _
               "Note: This affected all open Excel workbooks." & vbCrLf & vbCrLf & _
               "Settings will return to system defaults when Excel is closed.", _
               vbInformation, "Separator Changed"
    Else
        MsgBox "Separator not changed." & vbCrLf & _
               "Some functions may not work correctly." & vbCrLf & _
               "You can manually change the separator:" & vbCrLf & _
               "File -> Options -> Advanced -> Decimal separator", _
               vbExclamation, "Warning"
    End If
End Sub
