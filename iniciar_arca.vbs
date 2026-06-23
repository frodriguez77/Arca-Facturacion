Option Explicit

Dim objShell, objFSO
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

Dim appDir
appDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

Const URL        = "http://localhost:5000"
Const REMOTE_URL = "https://raw.githubusercontent.com/frodriguez77/Arca-Facturacion/claude/new-pc-download-setup-5AF1K/VERSION"

' --- Control de actualizaciones (antes de iniciar el servidor) ---
Dim localVer, remoteVer
localVer  = GetLocalVersion()
remoteVer = GetRemoteVersion()

If remoteVer <> "" And remoteVer <> localVer Then
    ' Verificar que no estemos en un loop post-actualizacion
    Dim flagFile : flagFile = appDir & "\update_done.flag"
    If objFSO.FileExists(flagFile) Then
        ' Ya se actualizo, borrar flag y continuar normalmente
        objFSO.DeleteFile flagFile, True
    Else
        ' Marcar que vamos a actualizar y lanzar el actualizador
        Dim ts2 : Set ts2 = objFSO.CreateTextFile(flagFile, True)
        ts2.Write "updated"
        ts2.Close
        objShell.Run "powershell -ExecutionPolicy Bypass -File """ & appDir & "\actualizar.ps1"" -AutoReiniciar", 1, True
        WScript.Quit
    End If
End If

' --- Preparar arca.exe (copia de pythonw.exe en la carpeta de Python) ---
Dim arcaExe
arcaExe = PrepararArcaExe()

' --- Iniciar servidor si no esta corriendo ---
If Not ServidorActivo() Then
    If arcaExe <> "" Then
        objShell.Run """" & arcaExe & """ """ & appDir & "\app.py""", 0, False
    Else
        objShell.Run "py """ & appDir & "\app.py""", 0, False
    End If

    Dim i
    For i = 1 To 15
        WScript.Sleep 1000
        If ServidorActivo() Then Exit For
    Next
End If

' --- Abrir Chrome ---
Dim chrome
chrome = BuscarChrome()
If chrome <> "" Then
    objShell.Run """" & chrome & """ --new-window " & URL, 1, False
Else
    objShell.Run URL, 1, False
End If


' ================================================================
Function GetLocalVersion()
    On Error Resume Next
    Dim f : f = appDir & "\VERSION"
    If Not objFSO.FileExists(f) Then
        GetLocalVersion = "0.0.0"
        Exit Function
    End If
    Dim ts : Set ts = objFSO.OpenTextFile(f, 1)
    GetLocalVersion = LimpiarVer(ts.ReadAll())
    ts.Close
    On Error GoTo 0
End Function

Function GetRemoteVersion()
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", REMOTE_URL, False
    http.SetTimeouts 500, 500, 4000, 4000
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        GetRemoteVersion = LimpiarVer(http.ResponseText)
    Else
        GetRemoteVersion = ""
    End If
    On Error GoTo 0
End Function

' Trim() en VBScript NO elimina Chr(10)/Chr(13), solo espacios.
' Esta funcion los elimina para poder comparar versiones correctamente.
Function LimpiarVer(s)
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(10), "")
    LimpiarVer = Trim(s)
End Function

Function PrepararArcaExe()
    On Error Resume Next

    ' Encontrar pythonw.exe via "where python"
    Dim oExec
    Set oExec = objShell.Exec("cmd /c where python")
    Dim pythonPath : pythonPath = Trim(oExec.StdOut.ReadAll())
    pythonPath = Trim(Split(pythonPath, vbCrLf)(0))

    If pythonPath = "" Or Not objFSO.FileExists(pythonPath) Then
        PrepararArcaExe = ""
        Exit Function
    End If

    Dim pythonDir   : pythonDir   = objFSO.GetParentFolderName(pythonPath)
    Dim pythonwPath : pythonwPath = pythonDir & "\pythonw.exe"
    Dim arcaPath    : arcaPath    = pythonDir & "\arca.exe"

    ' Copiar pythonw.exe como arca.exe (solo la primera vez)
    If Not objFSO.FileExists(arcaPath) Then
        If objFSO.FileExists(pythonwPath) Then
            objFSO.CopyFile pythonwPath, arcaPath
        End If
    End If

    If objFSO.FileExists(arcaPath) Then
        PrepararArcaExe = arcaPath
    Else
        PrepararArcaExe = ""
    End If

    On Error GoTo 0
End Function

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
