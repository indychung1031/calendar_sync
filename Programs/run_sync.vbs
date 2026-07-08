' Outlook ↔ Google 캘린더 동기화 — 창 없이 백그라운드 실행
' Windows Task Scheduler에서 이 파일을 실행하도록 등록하세요.
'
' 프로그램: wscript.exe
' 인수:    "...\Programs\run_sync.vbs"
' 시작 위치: ...\Programs
'
' 주의: 여러 Python이 설치된 경우 pyw/py 기본값(예: 3.14)에는 의존성이
'       없을 수 있다. 따라서 후보 인터프리터마다 필수 패키지 import를
'       실제로 검증해, 의존성이 설치된 pythonw.exe만 사용한다.

Option Explicit

Dim fso, shell, scriptDir, cmd, pythonw

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

pythonw = FindPythonw()
If pythonw = "" Then
    WScript.Echo "의존성이 설치된 pythonw.exe를 찾을 수 없습니다. requirements.txt 설치 여부를 확인하세요."
    WScript.Quit 1
End If

shell.CurrentDirectory = scriptDir
' 0 = 창 숨김, False = wscript는 즉시 종료(동기화는 백그라운드)
cmd = """" & pythonw & """ """ & scriptDir & "\main.py"" --quiet"
shell.Run cmd, 0, False

' 필수 의존성이 설치된 pythonw.exe 경로를 반환. 없으면 "".
Function FindPythonw()
    ' 후보를 삽입 순서대로 보관 (Dictionary는 중복 방지 + 순서 유지)
    Dim cand
    Set cand = CreateObject("Scripting.Dictionary")

    ' 1) 명시적으로 알려진 설치 경로 (최우선)
    AddCandidate cand, shell.ExpandEnvironmentStrings( _
        "%LocalAppData%\Programs\Python\Python313\pythonw.exe")

    ' 2) PATH 상의 pythonw.exe 전부
    AddAllOnPath cand, "pythonw.exe"

    ' 3) PATH 상의 python.exe 옆 pythonw.exe 전부
    AddSiblingPythonw cand, "python.exe"

    ' 4) 최후의 폴백: 런처
    AddAllOnPath cand, "pyw.exe"

    ' 의존성 검증 통과하는 첫 후보 사용
    Dim key
    For Each key In cand.Keys
        If HasDeps(key) Then
            FindPythonw = key
            Exit Function
        End If
    Next

    FindPythonw = ""
End Function

' 후보 추가 (실제 파일이 존재할 때만, 중복 제외)
Sub AddCandidate(dict, path)
    If path = "" Then Exit Sub
    If Not fso.FileExists(path) Then Exit Sub
    If Not dict.Exists(path) Then dict.Add path, True
End Sub

' `where <name>` 결과의 각 경로를 후보로 추가
Sub AddAllOnPath(dict, name)
    Dim stream, line
    On Error Resume Next
    Set stream = shell.Exec("cmd /c where " & name)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Do While Not stream.StdOut.AtEndOfStream
        line = Trim(stream.StdOut.ReadLine())
        AddCandidate dict, line
    Loop
    On Error GoTo 0
End Sub

' `where python.exe` 각 경로 옆의 pythonw.exe를 후보로 추가
Sub AddSiblingPythonw(dict, name)
    Dim stream, line, sib
    On Error Resume Next
    Set stream = shell.Exec("cmd /c where " & name)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Do While Not stream.StdOut.AtEndOfStream
        line = Trim(stream.StdOut.ReadLine())
        If line <> "" Then
            sib = fso.GetParentFolderName(line) & "\pythonw.exe"
            AddCandidate dict, sib
        End If
    Loop
    On Error GoTo 0
End Sub

' 해당 인터프리터에 필수 패키지가 설치돼 있으면 True
Function HasDeps(exe)
    Dim rc
    On Error Resume Next
    ' 창 숨김(0) + 완료 대기(True). import 실패 시 종료 코드 != 0
    rc = shell.Run("""" & exe & """ -c ""import google.auth, googleapiclient, win32com.client""", 0, True)
    If Err.Number <> 0 Then
        On Error GoTo 0
        HasDeps = False
        Exit Function
    End If
    On Error GoTo 0
    HasDeps = (rc = 0)
End Function
