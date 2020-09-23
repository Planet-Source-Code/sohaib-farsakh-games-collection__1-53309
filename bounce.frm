VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2880
      Top             =   1440
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   500
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   4500
      Width           =   500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim max
Dim hor, ver
Private Sub Form_Load()
max = 2500
hor = 1
ver = 1
End Sub

Private Sub Timer1_Timer()
If max < 7000 Then
max = max + 10
End If
If hor = 1 Then
Shape1.Left = Shape1.Left + 50
Else
Shape1.Left = Shape1.Left - 50
End If
b = (Shape1.Top - max) / 20 + 50
b = Abs(b)
If ver = 1 Then
Shape1.Top = Shape1.Top + b
Else
Shape1.Top = Shape1.Top - b
End If
If Shape1.Top <= max Then
ver = 1
End If
If Shape1.Top >= 7500 Then
ver = 2
If Shape1.Left <= 0 Then
hor = 1
End If
If Shape1.Left >= 10300 Then
hor = 2
End If
End If

End Sub
