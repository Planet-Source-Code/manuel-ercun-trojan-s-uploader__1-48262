Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpsubkey As String) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type


Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const PROCESS_TERMINATE As Long = &H1


Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)

Dim Process As PROCESSENTRY32
Dim pids As Long
Dim fso As New FileSystemObject

Public Sub Processs()
Form1.Text1 = Form1.Text1 & "Detecting process" & vbCrLf
Dim res&, ant&
Dim u As Integer
res = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
If res <> 0 Then
Process.dwSize = Len(Process)
ant = Process32First(res, Process)
u = 0
Do

If InStr(Process.szexeFile, "Mgsdll32.exe") Then pids = Process.th32ProcessID: Form1.Text1 = Form1.Text1 & "Process found" & vbCrLf
ant = Process32Next(res, Process)
u = u + 1
Loop Until ant = 0
End If
If pids = 0 Then Form1.Text1 = Form1.Text1 & "the Remover does not find process" & vbCrLf: Exit Sub
killprocess pids
Form1.Text1.SelStart = Len(Form1.Text1)
End Sub

Private Sub killprocess(pid As Long)
Dim res&
Dim uexit As Long
res = OpenProcess(PROCESS_TERMINATE, 0&, pid)
If res <> 0 Then
Call TerminateProcess(res, uexit)
Call CloseHandle(res)
Form1.Text1 = Form1.Text1 & "Kill Process" & vbCrLf
End If
Form1.Text1.SelStart = Len(Form1.Text1)
End Sub
Public Function GetSystem() As String
Dim s As String
Dim res As Long
s = String(255, vbNullChar)
res = GetSystemDirectory(s, Len(s))
If res <> 0 Then GetSystem = Left(s, InStr(s, vbNullChar) - 1) & "\"
End Function
Public Sub exe()
DoEvents
Form1.Text1 = Form1.Text1 & "Detecting binary" & vbCrLf
If fso.FileExists(GetSystem & "Mgsdll32.exe") = True Then
Kill GetSystem & "Mgsdll32.exe"
Form1.Text1 = Form1.Text1 & "Delete binary" & vbCrLf
Else
Form1.Text1 = Form1.Text1 & "The remover does not find the binary" & vbCrLf
End If
Form1.Text1.SelStart = Len(Form1.Text1)
End Sub
Public Sub reg()
On Error Resume Next
Form1.Text1 = Form1.Text1 & "Find Regedit" & vbCrLf

Dim fs
Set fs = CreateObject("wscript.shell")
fs.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\KernelDo"
RegDeleteKey HKEY_CLASSES_ROOT, ".troj"
RegDeleteKey HKEY_CLASSES_ROOT, "trojfile\Shell\Open\Command"
RegDeleteKey HKEY_CLASSES_ROOT, "trojfile\Shell\Open"
RegDeleteKey HKEY_CLASSES_ROOT, "trojfile\Shell"
RegDeleteKey HKEY_CLASSES_ROOT, "trojfile"
Form1.Text1 = Form1.Text1 & "Delete Regedit" & vbCrLf
Form1.Text1.SelStart = Len(Form1.Text1)
End Sub
