VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1440
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TrojanII.UserControl1 UserControl11 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Unload Form4
Me.Move Form1.Left + (Form1.Width / 2) - (Me.Width / 2), Form1.Top + (Form1.Height / 2) - (Me.Height / 2)
UserControl11.Move 0, 0

End Sub

Private Sub Form_Resize()
Form3.Width = UserControl11.Width
Form3.Height = UserControl11.Height
End Sub

Private Sub UserControl11_Salir()
Form1.TreeView1_DblClick
DoEvents
Unload Form4
Unload Me
End Sub
