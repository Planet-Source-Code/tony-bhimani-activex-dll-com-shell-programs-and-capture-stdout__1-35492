Attribute VB_Name = "modOSVer"
' Module to find out what version operating system this DLL
' is installed on. Since Windows 9x uses command.com as the
' command interpreter and Windows NT uses cmd.exe, we have
' to know what we're working with here to be able to shell
' DOS programs.

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


' Check what our OS is and send it to the caller
' either 3.x, 9x, or NT
Public Function CheckOS() As Integer
  Dim OSInfo As OSVERSIONINFO
    
  OSInfo.dwOSVersionInfoSize = Len(OSInfo)
  GetVersionEx OSInfo
  CheckOS = OSInfo.dwPlatformId
End Function
