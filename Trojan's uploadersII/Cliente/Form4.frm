VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1155
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1440
      Picture         =   "Form4.frx":0000
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting Opening the file."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Move Form1.Left + (Form1.Width / 2) - (Me.Width / 2), Form1.Top + (Form1.Height / 2) - (Me.Height / 2)

End Sub

Private Sub Timer1_Timer()
If Image1.Tag = "up" Then
Image1.Visible = True
Image1.Tag = "down"
Else
Image1.Visible = False
Image1.Tag = "up"
End If
End Sub
