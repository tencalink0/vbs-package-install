Dim message, title

If WScript.Arguments.Count > 1 Then
    message = WScript.Arguments(0)
    title = WScript.Arguments(1)
    Do
        MsgBox message, 4096, title
    Loop
ElseIf WScript.Arguments.Count > 0 Then
    message = WScript.Arguments(0)
    Do
        MsgBox message, 4096
    Loop
End If