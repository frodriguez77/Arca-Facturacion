Option Explicit

Dim objShell, objFSO
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

Dim appDir
appDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

Const URL = "http://localhost:5000"

If Not ServidorActivo() Then
    objShell.Run "py """ & appDir & "\app.py""", 0, False

    Dim i
    For i = 1 To 15
        WScript.Sleep 1000
        If ServidorActivo() Then Exit For
    Next
End If

Dim chrome
chrome = BuscarChrome()
If chrome <> "" Then
    objShell.Run """" & chrome & """ --new-window " & URL, 1, False
Else
    objShell.Run URL, 1, False
End If


Function ServidorActivo()
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", URL & "/", False
    http.SetTimeouts 300, 300, 1000, 1000
    http.Send
    ServidorActivo = (Err.Number = 0 And http.Status > 0)
    On Error GoTo 0
End Function

Function BuscarChrome()
    Dim paths(2)
    paths(0) = objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe"
    paths(1) = objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe"
    paths(2) = objShell.ExpandEnvironmentStrings("%LocalAppData%") & "\Google\Chrome\Application\chrome.exe"
    Dim j
    For j = 0 To 2
        If objFSO.FileExists(paths(j)) Then
            BuscarChrome = paths(j)
            Exit Function
        End If
    Next
    BuscarChrome = ""
End Function
