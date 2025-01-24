Dim objShell, objFSO
Set objShell = Wscript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim title, progName, progDest, currentPath, desktopPath, barFilePath, loadingBarCode
title = "Downloader"
progName = "BChat"
progFilename = "main\main.vbs"
iconFilename = "main\main.ico"

barFilePath = "bar\bar.txt"
loadingBarCode = "bar\loadingbar.vbs"

currentPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
desktopPath = objShell.SpecialFolders("Desktop")

Sub WriteToFile(filePath, content)
    Dim file
    Set file = objFSO.OpenTextFile(currentPath + "\" + filePath, 2, True)
    file.WriteLine(content)
    file.Close
End Sub

Function CreateShortcut()
    Dim desktopPath, programPath, shortcutPath, iconPath
    desktopPath = currentPath

    WriteToFile barFilePath, "0/8"
    Set objProcess = objShell.Exec("wscript """ & loadingBarCode & """")

    programPath = currentPath + "\" + progFilename
    shortcutPath = desktopPath + "\" + progName + ".lnk"
    iconPath = currentPath + "\" + iconFilename
    WriteToFile barFilePath, "3/8"
    Wscript.Sleep 500

    Set shortcut = objShell.CreateShortcut(shortcutPath)
    shortcut.TargetPath = programPath
    shortcut.IconLocation = iconPath
    WriteToFile barFilePath, "6/8"
    shortcut.Save
    WriteToFile barFilePath, "8/8"
    Wscript.Sleep 500 'Buffer time

    WriteToFile barFilePath, "Done"
End Function

Dim answer
answer = msgbox("Are you sure you would like to continue installing " + progName, 4096 + 48 + 4, title)
If answer = 6 Then
    CreateShortcut()
Else
    msgbox "Creation failed", 4096 + 16, title
End If