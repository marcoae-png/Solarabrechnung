Option Explicit
On Error Resume Next
Dim fso, shell, folder, script, root, subfolder, python, pyExec, args
Dim command, appUrl, ready, attempt
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
script = fso.BuildPath(folder, "solar_server.py")
shell.CurrentDirectory = folder
appUrl = "http://127.0.0.1:8774/"
python = ""
args = ""

' Prefer a registered Python launcher; do not pass launcher flags to python.exe.
Set pyExec = shell.Exec("cmd /c where py.exe")
If Err.Number = 0 Then
  If pyExec.ExitCode = 0 Then
    python = Trim(pyExec.StdOut.ReadLine)
    If LCase(fso.GetFileName(python)) = "py.exe" Then args = "-3 "
  End If
End If
Err.Clear

' This machine can use the Python bundled with Codex.
If python = "" Then
  root = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\WindowsApps"
  If fso.FolderExists(root) Then
    For Each subfolder In fso.GetFolder(root).SubFolders
      If InStr(1, subfolder.Name, "OpenAI.CodexPrimaryRuntime", vbTextCompare) = 1 Then
        If fso.FileExists(fso.BuildPath(subfolder.Path, "dependencies\python\python.exe")) Then
          python = fso.BuildPath(subfolder.Path, "dependencies\python\python.exe")
        End If
      End If
    Next
  End If
End If

' Direct check for WindowsApps installations that block folder enumeration.
If python = "" Then
  python = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\WindowsApps\OpenAI.CodexPrimaryRuntime.v26-1105-1297-0_26.1105.1297.0_x64__3k8sg7r9htsxt\dependencies\python\python.exe"
  If Not fso.FileExists(python) Then python = ""
End If
Err.Clear
On Error GoTo 0

If python = "" Then
  MsgBox "Python 3 wurde nicht gefunden. Installiere Python 3.11 oder neuer und führe py -3 -m pip install -r requirements.txt im Programmordner aus.", vbExclamation, "Solarabrechnung"
  WScript.Quit 1
End If

command = Chr(34) & python & Chr(34) & " " & args & Chr(34) & script & Chr(34)
On Error Resume Next
shell.Run command, 0, False
On Error GoTo 0

ready = False
For attempt = 1 To 30
  WScript.Sleep 500
  If ServerReady() Then
    ready = True
    Exit For
  End If
Next

If ready Then
  shell.Run appUrl, 1, False
Else
  MsgBox "Der Server antwortet nicht. Prüfe, ob Python 3.11 und die Pakete aus requirements.txt installiert sind. Falls der Port 8774 bereits verwendet wird, schliesse die andere Solarabrechnung zuerst.", vbCritical, "Solarabrechnung"
End If

Function ServerReady()
  Dim request
  ServerReady = False
  On Error Resume Next
  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
  request.SetTimeouts 500, 500, 500, 500
  request.Open "GET", "http://127.0.0.1:8774/api/status", False
  request.Send
  If Err.Number = 0 Then ServerReady = (request.Status = 200)
  Err.Clear
  On Error GoTo 0
End Function
