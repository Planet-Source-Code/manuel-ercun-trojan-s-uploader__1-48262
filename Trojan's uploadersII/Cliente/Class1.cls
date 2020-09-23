VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200

Dim rn As NOTIFYICONDATA
Private Sub Class_Initialize()
rn.cbSize = Len(rn)
rn.hIcon = Form1.Picture1.Picture
rn.hwnd = Form1.Picture1.hwnd
rn.szTip = "Trojan's Downloader"
rn.uID = 1
rn.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
rn.uCallbackMessage = WM_MOUSEMOVE
End Sub
Public Function ShellAdd()
Form1.Hide
Shell_NotifyIcon NIM_ADD, rn
End Function

Public Function ShellDel()
Shell_NotifyIcon NIM_DELETE, rn
Form1.WindowState = vbNormal
Form1.Show
End Function

