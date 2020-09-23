VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Project: myASPx.dll (COM DLL for use by apps or a
'          webserver. (can be used by ASP or ColdFusion)
'
' Description: Very similar to a Free server addon called
'              ASPExec which executes DOS or Windows apps
'              and returns their stdout. This can be used
'              for many purposes. e.g. CGI, ping, tracert,
'              shelling system tasks, etc... and displaying
'              the output to a webpage.
'
' Author: Tony Bhimani
' Email: TonyBhimani@hotmail.com
' Homepage: http://www.xenocafe.com
' Demo scripts: http://demos.xenocafe.com
'
Option Explicit

Dim strAppName As String
Dim strParams As String
Dim nShowWindow As Integer

' Set the application name property (e.g. c:\path\whatever.exe)
Public Property Let AppName(ByVal tmpAppName As String)
  strAppName = tmpAppName
End Property

' Return the application name property
Public Property Get AppName() As String
  AppName = strAppName
End Property

' Set the application parameters property (e.g. arguments like /w /p or whatever)
Public Property Let Params(ByVal tmpParams As String)
  strParams = tmpParams
End Property

' Return the application parameters
Public Property Get Params() As String
  Params = strParams
End Property

' Set the window state (e.g. 0 = Hidden, 1 = Normal, etc...)
Public Property Let ShowWindow(ByVal tmpShowWindow As Integer)
  nShowWindow = tmpShowWindow
End Property

' Return the window state
Public Property Get ShowWindow() As Integer
  ShowWindow = nShowWindow
End Property

' Shell a DOS program and return its stdout if any
Public Function DosExec() As String
  Dim strCmd As String
  
  ' make sure the AppName() property has been set
  ' if it hasn't, fail
  If Len(strAppName) = 0 Then
    DosExec = "AppName() property not set."
    Exit Function
  End If
  
  ' check what OS we are running on so we can use the
  ' correct method of creating the process
  '
  ' From Mattias Sjögren's source - reason we use a 32-bit console for Win9x
  '
  ' If we are running Windows 9x, we have to launch the command using an
  ' intermediate Win32 console application (RedirStub.exe in this case),
  ' since Command.com is a 16-bit program. See KB article Q150956.
  Select Case CheckOS
    Case VER_PLATFORM_WIN32s
      'Windows 3.x (FAIL FAIL FAIL - Get with the times)
      DosExec = "Ancient OS detected. It's time to upgrade."
      Exit Function
    Case VER_PLATFORM_WIN32_WINDOWS
      ' Windows 9x
      strCmd = "RedirStub " & strCmd
    Case VER_PLATFORM_WIN32_NT
      ' Windows NT/2000/XP
      strCmd = Environ$("COMSPEC") & " /c " & strAppName & " " & strParams
  End Select
  
  ' shell it and return it
  DosExec = GetCommandOutput(strCmd, True, True, , nShowWindow)
End Function

' Shell a Windows program and return its stdout if any
Public Function WinExec() As String
  Dim strCmd As String
  
  ' make sure the AppName() property has been set
  ' if it hasn't, fail
  If Len(strAppName) = 0 Then
    WinExec = "AppName() property not set."
    Exit Function
  End If
  
  ' Build the shell string
  strCmd = strAppName & " " & strParams
  
  ' shell it and return it
  WinExec = GetCommandOutput(strCmd, True, True, , nShowWindow)
End Function

