' Outlook ↔ Google 캘린더 동기화 — 창 없이 백그라운드 실행
' Windows Task Scheduler에서 이 파일을 실행하도록 등록하세요.
'
' 프로그램: wscript.exe
' 인수:    "...\Programs\run_sync.vbs"
' 시작 위치: ...\Programs

Option Explicit

Dim fso, shell, scriptDir, cmd, pythonw

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

pythonw = FindPythonw()
If pythonw = "" Then
    WScript.Echo "pythonw.exe를 찾을 수 없습니다. Python 설치 경로를 확인하세요."
    WScript.Quit 1
End If

shell.CurrentDirectory = scriptDir
' 0 = 창 숨김, False = wscript는 즉시 종료(동기화는 백그라운드)
cmd = """" & pythonw & """ """ & scriptDir & "\main.py"" --quiet"
shell.Run cmd, 0, False

Function FindPythonw()
    Dim exe, candidates, i, p
    ' pyw 런처 (Python 공식 설치)
    exe = FindOnPath("pyw.exe")
    If exe <> "" Then
        FindPythonw = exe
        Exit Function
    End If
    exe = FindOnPath("pythonw.exe")
    If exe <> "" Then
        FindPythonw = exe
        Exit Function
    End If
    ' python.exe 옆 pythonw.exe
    exe = FindOnPath("python.exe")
    If exe <> "" Then
        p = fso.GetParentFolderName(exe) & "\pythonw.exe"
        If fso.FileExists(p) Then
            FindPythonw = p
            Exit Function
        End If
    End If
    FindPythonw = ""
End Function

Function FindOnPath(name)
    Dim stream, line
    On Error Resume Next
    Set stream = shell.Exec("cmd /c where " & name)
    If Err.Number <> 0 Then
        FindOnPath = ""
        Exit Function
    End If
    Do While Not stream.StdOut.AtEndOfStream
        line = Trim(stream.StdOut.ReadLine())
        If line <> "" And fso.FileExists(line) Then
            FindOnPath = line
            Exit Function
        End If
    Loop
    FindOnPath = ""
End Function
