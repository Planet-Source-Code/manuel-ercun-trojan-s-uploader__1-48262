Attribute VB_Name = "Module1"
Option Explicit

Public i As Integer
Public v
Public b
Public rec, pah As String




Public Function Icons(name As String) As Integer
Dim jun As Integer
If Right(name, 3) = "exe" Or Right(name, 3) = UCase("exe") Then
jun = 10
ElseIf Right(name, 3) = "bat" Or Right(name, 3) = UCase("ucase") Then
jun = 7
ElseIf Right(name, 3) = "bmp" Or Right(name, 3) = UCase("bmp") Then
jun = 8
ElseIf Right(name, 3) = "doc" Or Right(name, 3) = UCase("doc") Then
jun = 9
ElseIf Right(name, 3) = "gif" Or Right(name, 3) = UCase("gif") Then
jun = 11
ElseIf Right(name, 3) = "htm" Or Right(name, 3) = UCase("htm") Or Right(name, 4) = "html" Or Right(name, 4) = UCase("html") Then
jun = 16
ElseIf Right(name, 3) = "ini" Or Right(name, 3) = UCase("ini") Then
jun = 15
ElseIf Right(name, 3) = "jpg" Or Right(name, 3) = UCase("jpg") Then
jun = 12
ElseIf Right(name, 3) = "sys" Or Right(name, 3) = UCase("sys") Or Right(name, 3) = "dll" Or Right(name, 3) = UCase("dll") Then
jun = 17
ElseIf Right(name, 3) = "vbp" Or Right(name, 3) = UCase("vbp") Or Right(name, 3) = "frm" Or Right(name, 3) = UCase("frm") Then
jun = 28
ElseIf Right(name, 3) = "txt" Or Right(name, 3) = UCase("txt") Then
jun = 18
ElseIf Right(name, 3) = "wav" Or Right(name, 3) = UCase("wav") Or Right(name, 3) = "mp3" Or Right(name, 3) = UCase("mp3") Then
jun = 13
ElseIf Right(name, 3) = "mpg" Or Right(name, 3) = UCase("mpg") Or Right(name, 3) = "avi" Or Right(name, 3) = UCase("avi") Then
jun = 14
ElseIf Right(name, 3) = "zip" Or Right(name, 3) = UCase("zip") Then
jun = 19
ElseIf Right(name, 3) = "rar" Or Right(name, 3) = UCase("rar") Then
jun = 20

Else
jun = 6
End If
Icons = jun
End Function
Public Function Barrs(str As String) As String
If InStr(str, "\\") Then Barrs = Replace(str, "\\", "\")
End Function




