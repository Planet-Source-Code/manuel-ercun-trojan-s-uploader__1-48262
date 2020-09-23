VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Trojan's Downloader"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   4560
         TabIndex        =   3
         Top             =   840
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
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   4335
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   600
         Left            =   4560
         TabIndex        =   4
         Top             =   1560
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
               ImageIndex      =   2
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   495
         Left            =   1800
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
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
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":14BE
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
exe
Timer1.Enabled = False
Form1.Text1.SelStart = Len(Text1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Text1.Text = ""
Processs
reg
DoEvents
Form1.Text1 = Form1.Text1 & "Detecting binary" & vbCrLf

Timer1.Enabled = True

Form1.Text1.SelStart = Len(Text1)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
End
End Sub
