Dim objShell, objFSO
Set objShell = Wscript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim title, currentPath, barOutputPath, barFilePath
title = "Downloader"
currentPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

barOutputPath = "argout.vbs"
barFilePath = "bar.txt"

Dim tickSpeed, tickLimit
tickSpeed = 10 '10
tickLimit = 1000 '1000

BarLoop()

Function ReadBar()
    Dim file, line
    Set file = objFSO.OpenTextFile(currentPath + "\" + barFilePath, 1)
    line = file.ReadLine
    file.Close

    If InStr(line, "/") = 0 Then
        ReadBar = line
    Else
        ReadBar = Split(line, "/")
    End If
End Function

Function CompareArrays(arr1, arr2)
    Dim i
    If UBound(arr1) <> UBound(arr2) Then
        CompareArrays = False
        Exit Function
    End If
    For i = 0 To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            CompareArrays = False
            Exit Function
        End If
    Next
    CompareArrays = True
End Function

Function BarLoop()
    Dim ended, change, tick, bar
    Dim scriptPath, objProcess, barString
    Dim barMem, finalMessage

    Set objProcess = Nothing
    ended = False
    tick = 0
    
    ReDim barMem(1)
    finalMessage = "Completed"

    Do Until ended = True Or tick >= tickLimit
        bar = ReadBar()

        If VarType(bar) = 8 Then
            If bar = "Done" Then
                ended = True
                Exit Do
            End If
        ElseIf IsArray(bar) Then
            If UBound(bar) + 1 < 2 Then
                ended = True
                Exit Do
            ElseIf Not CompareArrays(barMem, bar) Then
                ReDim barMem(UBound(bar))
                For i = 0 To UBound(bar)
                    barMem(i) = bar(i)
                Next

                scriptPath = currentPath & "\" & barOutputPath
                barString = DrawBar(bar(0), bar(1))

                If Not objProcess Is Nothing Then
                    On Error Resume Next
                    objProcess.Terminate
                    On Error GoTo 0
                End If

                Set objProcess = objShell.Exec("wscript """ & scriptPath & """ """ & barString & """ """ & title & """ ")
                tick = 0
            Else
                tick = tick + 1
            End If
        Else
            ended = True
            Exit Do
        End If

        If tick >= tickLimit Then
            finalMessage = "Timeout!"
        End If

        WScript.Sleep tickSpeed
    Loop
    On Error Resume Next
    objProcess.Terminate
    On Error GoTo 0
    scriptPath = currentPath & "\" & barOutputPath
    Set objProcess = objShell.Exec("wscript """ & scriptPath & """ """ & "Completed!" & """ """ & title & """ ")
    Wscript.Sleep 1000
    On Error Resume Next
    objProcess.Terminate
    On Error GoTo 0
End Function

Function DrawBar(current, limit)
    Dim bar, i

    bar = ""

    For i = 1 To current
        bar = bar & "#"
    Next

    For i = 1 To (limit - current)
        bar = bar & "-"
    Next

    DrawBar = bar
End Function

Function RenderBar(timeout, bar)
    Dim scriptPath, objProcess, barString
    scriptPath = currentPath + "\" + barOutputPath

    barString = DrawBar(bar(0), bar(1))

    Set objProcess = objShell.Exec("wscript """ & scriptPath & """ """ & barString & """ """ & title & """ ")
    WScript.Sleep timeout
    objProcess.Terminate
End Function