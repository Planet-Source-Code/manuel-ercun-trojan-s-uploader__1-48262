VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   360
      ScaleHeight     =   390
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   1320
      Width           =   3810
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   3765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   2
         Top             =   105
         Width           =   525
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Image Image1 
      Height          =   410
      Left            =   4170
      Picture         =   "UserControl1.ctx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim imax, imin, ivalue As Long

Event Salir()
Private Sub Image1_Click()
RaiseEvent Salir
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub





Private Sub UserControl_Initialize()
Picture1.Move 0, 0
Image1.Move Picture1.Width, 0
imin = 0
imax = 100
ivalue = 1
End Sub

Public Property Get Max() As Long
Max = imax
End Property
Public Property Let Max(ByVal new_max As Long)
imax = new_max
If imax < ivalue Then imax = ivalue
If imax < imin Then imax = imin
PropertyChanged "Max"
End Property

Public Property Get Min() As Long
Min = imin
End Property

Public Property Let Min(ByVal new_min As Long)
imin = new_min
If imin > imax Then imin = imax
If imin > ivalue Then imin = ivalue
PropertyChanged "Min"
End Property

Public Property Get Value() As Long
Value = ivalue
End Property
Public Property Let Value(ByVal new_value As Long)
ivalue = new_value
If ivalue > imax Then ivalue = imax
If ivalue < imin Then ivalue = imin
Label1.Width = Int(ivalue - imin) / Int(imax - imin) * Picture1.ScaleWidth

Label2.Caption = Format(ivalue / imax, "0.0%")


PropertyChanged "Value"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Max = PropBag.ReadProperty("Max", 100)
Min = PropBag.ReadProperty("Min", 0)
Value = PropBag.ReadProperty("Value", 1)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Picture1.Height
Picture1.Width = UserControl.Width - Image1.Width
Shape1.Width = Picture1.Width - 40
Label2.Left = (Picture1.Width / 2) - (Label2.Width / 2) + 100
Image1.Left = Picture1.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Max", imax, 100)
Call PropBag.WriteProperty("Min", imin, 0)
Call PropBag.WriteProperty("Value", ivalue, 1)
End Sub



