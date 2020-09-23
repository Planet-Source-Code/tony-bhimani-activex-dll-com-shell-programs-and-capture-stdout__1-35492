VERSION 5.00
Begin VB.Form frmVBShell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shell Execute using myASPx.dll"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExecute 
      Caption         =   "E&xecute"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkWinProg 
      Caption         =   "&Windows Program"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkShowProg 
      Caption         =   "&Show Program"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtParams 
      Height          =   350
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtProgName 
      Height          =   350
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "P&arameters:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "&Program:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmVBShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project: VBShell.exe
'
' Description: Example of using the myASPx.dll in an EXE
'              to capture program STDOUT
'
' Author: Tony Bhimani
' Email: TonyBhimani@hotmail.com
' Homepage: http://www.xenocafe.com
' Demo scripts: http://demos.xenocafe.com

Option Explicit

' Declare an instance of our object
Dim myASPx As myASPx.Exec

Private Sub cmdExecute_Click()
  Dim strOut As String
  
  ' set the parameters
  myASPx.AppName = txtProgName.Text
  myASPx.Params = txtParams.Text
  myASPx.ShowWindow = chkShowProg.Value
  ' execute according to DOS or Windows program and grab the STDOUT
  If chkWinProg.Value = 1 Then
    strOut = myASPx.WinExec
  Else
    strOut = myASPx.DosExec
  End If
  
  ' Display our STDOUT if any (remember, Windows programs may not
  ' return anything. Depends on the program.
  If Len(strOut) > 0 Then
    MsgBox strOut, , "Shell Execute using myASPx.dll"
  End If
End Sub

Private Sub Form_Load()
  ' Create an instance of our DLL
  Set myASPx = CreateObject("myASPx.Exec")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Destroy the instance of our DLL
  Set myASPx = Nothing
End Sub
