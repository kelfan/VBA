# vb run python 
```vb
Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string'

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function

Sub TestShellRun()
    Dim str As String
    str1 = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ThisWorkbook.FullName)
    MsgBox ShellRun("python " & str1 & "\print.py")
End Sub
```

# vb run window bat File 
```vb
Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string'

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function

Sub TestShellRun()
    MsgBox ShellRun("E:\workspace\VBA\externalAccess\echo.bat")
End Sub
```

# 相对地址 Excel当前路径 
```vb
Sub TestShellRun()
    Dim str As String
    str1 = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ThisWorkbook.FullName)
    MsgBox ShellRun(str1 & "\echo.bat")
End Sub
```

# resources 
- [Control External Processes using VBA (FTP example) (cc)](https://www.youtube.com/watch?v=fMWWcoXnzHc)
- [Excel VBA Introduction Part 49 - Downloading Files from Websites](https://www.youtube.com/watch?v=JPezrWwvsJM)
- [Webinar: Python for Excel with PyXLL](https://www.youtube.com/watch?v=0RzTsvBIhaE)

- [Excel VBA: Insert Picture from Directory on Cell Value Change](https://www.youtube.com/watch?v=VUl3l9wB51M)