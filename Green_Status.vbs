Dim objShell, lngMinutes, boolValid

Set objShell = CreateObject("WScript.Shell")
lngMinutes = InputBox("How Long?" & Replace(Space(5), " ", vbNewLine) & "Enter Minutes:", "working")
If lngMinutes = vbEmpty Then

Else
        On Error Resume Next
        Err.Clear
        boolValid = False
        lngMinutes = CLng(lngMinutes)
        If Err.Number = O Then
                If lngMinutes > 0 Then
                        For i = 1 To lngMinutes
                                WScript.Sleep 60000
                                objShell.SendKeys "{SCROLLOCK}"
                        Next
                        boolValid = True
                        MsgBox "Back to work", vbOKOnly+vbInfomation, "Done"
                End If
        End If
        On Error Goto 0
        If boolValid = False Then
                MsgBox " Re-Enter" & vbNewLine & "Number please", vbOKOnly+vbInformation, "Not Working"
        End If
End If

Set objShell = Nothing
WScript.Quit 0
