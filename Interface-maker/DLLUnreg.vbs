'Unregister an ActiveX DLL or OCX.
'
'RUN THIS AS AN ADMIN USER (on Vista or later you will
'be prompted for elevation).
'

Option Explicit

Private Const DllName = "IRegexp.dll"

Private WinVer

  With WScript.CreateObject("WScript.Shell")
    WinVer = .RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    If Fix(CSng(replace(WinVer,".",","))) < 6 then
      'Win2K or XP (run by admin user).
      .Run "regsvr32 /u """ & WScript.Arguments(0) & """"
    Else
      'Vista or later, request elevation.
      With CreateObject("Shell.Application")
        .ShellExecute "regsvr32", "/u """ & CreateObject ("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\" & DllName & """", , "runas"
      End With
    End If
  End With
