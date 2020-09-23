Attribute VB_Name = "modIPWizard"
Option Explicit
'***********************************************
'*                IP Wizard                    *
'*  Copyright © 2000-2001 , Kemtech Software   *
'*             Michael J. Kempf                *
'***********************************************

'API to Get a Windows Version
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                     (ByRef lpVersionInformation As OSVERSIONINFO) As Long



Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' Returns Version of Windows as a String
' NOTE: Win95 returns "4.00"
        'Win98 returns "4.10"
        'WinNT returns ""
        'Win2000 returns "5.00"

Function WindowsVersion() As String
    Dim osInfo As OSVERSIONINFO
    
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    GetVersionEx osInfo
    
    WindowsVersion = osInfo.dwMajorVersion & "." & Right$("0" & Format$ _
        (osInfo.dwMinorVersion), 2)
        
End Function


Sub Main()
'Check Windows Version to make sure Windows 2000 is running
If Not WindowsVersion = "5.00" Then
    MsgBox "IP Wizard is only designed for the Windows 2000 Operating System !", vbCritical, "IP Wizard"
    End
Else
    Load frmIPWizard
    frmIPWizard.Show
End If

End Sub
