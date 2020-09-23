Attribute VB_Name = "Module1"
Option Explicit

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
Public ras As Integer
Public Edit As rserver

Public Sub Save(ByRef path As String)
Edit.taa = "taa"
Edit.ser = Form1.Text1
Edit.mai = Form1.Text2
Edit.asu = Form1.Text3
Edit.ras = ras
Edit.mes = Form1.Text7
Edit.tit = Form1.Text8
Edit.por = Form1.Text4
Edit.pro = Form1.Text5
Open path For Binary As #1
Seek #1, LOF(1) - 246
Put #1, , Edit.taa
Put #1, , Edit.ser
Put #1, , Edit.mai
Put #1, , Edit.asu
Put #1, , Edit.ras
Put #1, , Edit.mes
Put #1, , Edit.tit
Put #1, , Edit.por
Put #1, , Edit.pro
Close #1

End Sub

