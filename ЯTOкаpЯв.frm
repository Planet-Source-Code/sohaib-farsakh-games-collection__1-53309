VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   7920
      Max             =   20
      TabIndex        =   1
      Top             =   2880
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3840
      Top             =   5160
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   5295
      Left            =   8160
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   2280
      Shape           =   1  'Square
      Top             =   4560
      Width           =   720
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   5040
      Shape           =   1  'Square
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   720
      Shape           =   1  'Square
      Top             =   2160
      Width           =   720
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   4320
      Shape           =   1  'Square
      Top             =   3480
      Width           =   720
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   5520
      Shape           =   1  'Square
      Top             =   3600
      Width           =   720
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   0
      Shape           =   1  'Square
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()
Select Case Text1.Text
Case "2"
Shape1.Top = Shape1.Top + 720
Case "8"
Shape1.Top = Shape1.Top - 720
Case "4"
Shape1.Left = Shape1.Left - 720
Case "6"
Shape1.Left = Shape1.Left + 720
End Select
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
Shape2.Left = Shape2.Left - 250
Shape3.Left = Shape3.Left - 250
Shape4.Left = Shape4.Left - 250
Shape5.Left = Shape5.Left - 250
Shape6.Left = Shape6.Left - 250
If Shape2.Left < 0 Then
Randomize
a = Int(Rnd * 5300)
Shape2.Top = a
Shape2.Left = 7800
End If
If Shape6.Left < 0 Then
Randomize
a = Int(Rnd * 5300)
Shape6.Top = a
Shape6.Left = 7800
End If
If Shape3.Left < 0 Then
Randomize
a = Int(Rnd * 5300)
Shape3.Top = a
Shape3.Left = 7800
End If
If Shape4.Left < 0 Then
Randomize
a = Int(Rnd * 5300)
Shape4.Top = a
Shape4.Left = 7800
End If
If Shape5.Left < 0 Then
Randomize
a = Int(Rnd * 5300)
Shape5.Top = a
Shape5.Left = 7800
End If
On Error Resume Next
If Shape1.Left > Shape2.Left - 720 And Shape1.Left < Shape2.Left + 720 And Shape1.Top > Shape2.Top - 720 And Shape1.Top < Shape2.Top + 720 Then
Shape7.Height = Shape7.Height - 100
Shape7.Top = Shape7.Top + 100
End If
If Shape1.Left > Shape3.Left - 720 And Shape1.Left < Shape3.Left + 720 And Shape1.Top > Shape3.Top - 720 And Shape1.Top < Shape3.Top + 720 Then
Shape7.Height = Shape7.Height + 50
Shape7.Top = Shape7.Top - 50
End If
If Shape1.Left > Shape4.Left - 720 And Shape1.Left < Shape4.Left + 720 And Shape1.Top > Shape4.Top - 720 And Shape1.Top < Shape4.Top + 720 Then
Shape7.Height = Shape7.Height - 100
Shape7.Top = Shape7.Top + 100
End If
If Shape1.Left > Shape5.Left - 720 And Shape1.Left < Shape5.Left + 720 And Shape1.Top > Shape5.Top - 720 And Shape1.Top < Shape5.Top + 720 Then
Shape7.Height = Shape7.Height - 100
Shape7.Top = Shape7.Top + 100
End If
If Shape1.Left > Shape6.Left - 720 And Shape1.Left < Shape6.Left + 720 And Shape1.Top > Shape6.Top - 720 And Shape1.Top < Shape6.Top + 720 Then
Shape7.Height = Shape7.Height - 100
Shape7.Top = Shape7.Top + 100
End If
Timer1.Interval = 200 / VScroll1.Value
If Shape7.Height < 0 Then
End
End If
Label1.Caption = Val(Label1.Caption) + 1
End Sub
