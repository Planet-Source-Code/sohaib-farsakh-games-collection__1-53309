VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáßÑÉ"
   ClientHeight    =   7920
   ClientLeft      =   3570
   ClientTop       =   675
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "ÅíÞÇÝ"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   6120
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   120
      Top             =   1320
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ÇáßÑÉ.frx":0000
      Left            =   2040
      List            =   "ÇáßÑÉ.frx":0010
      TabIndex        =   6
      Text            =   "ãÊæÓØ"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "íÓÇÑ"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "íãíä"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   240
      Top             =   1560
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   400
      Left            =   3480
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   400
      Left            =   2400
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   400
      Left            =   1320
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   240
      Top             =   480
      Width           =   1005
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   2400
      Top             =   0
      Width           =   1005
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   1320
      Top             =   0
      Width           =   1005
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   240
      Top             =   0
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   3480
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "3"
      Height          =   375
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÃÑæÇÍ"
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "ÇáãÑÍáÉ"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "ÇáÏÑÌÉ"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   6960
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   3240
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
If Combo1.Text = "ãÈÊÏìÁ" Then
Timer1.Interval = 150
End If
If Combo1.Text = "ãÊæÓØ" Then
Timer1.Interval = 80
End If
If Combo1.Text = "ãÊÞÏã" Then
Timer1.Interval = 40
End If
If Combo1.Text = "ãÍÊÑÝ" Then
Timer1.Interval = 20
End If

End Sub

Private Sub Command1_Click()
Line1.X1 = Line1.X1 + 400
Line1.X2 = Line1.X2 + 400
End Sub

Private Sub Command2_Click()
Line1.X1 = Line1.X1 - 400
Line1.X2 = Line1.X2 - 400
End Sub

Private Sub Command3_Click()
Select Case Command3.Caption
Case "ÅíÞÇÝ"
Command3.Caption = "ÅßãÇá"
Case "ÅßãÇá"
Command3.Caption = "ÅíÞÇÝ"
End Select
End Sub

Private Sub Timer1_Timer()
If Command3.Caption = "ÅíÞÇÝ" Then
If Label1.Caption = "1" Then
Shape1.Top = Shape1.Top + 200
End If
If Label1.Caption = "2" Then
Shape1.Top = Shape1.Top - 200
End If
If Label2.Caption = "1" Then
Shape1.Left = Shape1.Left - 50
End If
If Label2.Caption = "2" Then
Shape1.Left = Shape1.Left + 50
End If
If Line1.X1 < (Shape1.Left + 250) And Line1.X2 > (Shape1.Left + 250) Then
If Shape1.Top > 5250 Then
Label1.Caption = "2"
Select Case Timer1.Interval
Case 150
Label4.Caption = Val(Label4.Caption) + 5
Case 80
Label4.Caption = Val(Label4.Caption) + 10
Case 40
Label4.Caption = Val(Label4.Caption) + 25
Case 20
Label4.Caption = Val(Label4.Caption) + 50
End Select
End If
End If
End If
If Shape1.Top > 5580 And Shape1.Top < 5780 Then
Label7.Caption = Val(Label7.Caption) - 0.5
End If
If Shape1.Top > Shape2.Top - 400 And Shape1.Top < Shape2.Top + 500 And Shape1.Left > Shape2.Left - 500 And Shape1.Left < Shape2.Left + 1000 Then
If Shape2.Visible = True Then
Shape2.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape3.Top - 400 And Shape1.Top < Shape3.Top + 500 And Shape1.Left > Shape3.Left - 500 And Shape1.Left < Shape3.Left + 1000 Then
If Shape3.Visible = True Then
Shape3.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape4.Top - 400 And Shape1.Top < Shape4.Top + 500 And Shape1.Left > Shape4.Left - 500 And Shape1.Left < Shape4.Left + 1000 Then
If Shape4.Visible = True Then
Shape4.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape5.Top - 400 And Shape1.Top < Shape5.Top + 500 And Shape1.Left > Shape5.Left - 500 And Shape1.Left < Shape5.Left + 1000 Then
If Shape5.Visible = True Then
Label4.Caption = Val(Label4.Caption) + 10
Shape5.Visible = False
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape6.Top - 400 And Shape1.Top < Shape6.Top + 500 And Shape1.Left > Shape6.Left - 500 And Shape1.Left < Shape6.Left + 1000 Then
If Shape6.Visible = True Then
Shape6.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape7.Top - 400 And Shape1.Top < Shape7.Top + 500 And Shape1.Left > Shape7.Left - 500 And Shape1.Left < Shape7.Left + 1000 Then
If Shape7.Visible = True Then
Shape7.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape8.Top - 400 And Shape1.Top < Shape8.Top + 500 And Shape1.Left > Shape8.Left - 500 And Shape1.Left < Shape8.Left + 1000 Then
If Shape8.Visible = True Then
Shape8.Visible = False
Label4.Caption = Val(Label4.Caption) + 10
Label1.Caption = "1"
End If
End If
If Shape1.Top > Shape9.Top - 400 And Shape1.Top < Shape9.Top + 500 And Shape1.Left > Shape9.Left - 500 And Shape1.Left < Shape9.Left + 1000 Then
If Shape9.Visible = True Then
Label4.Caption = Val(Label4.Caption) + 10
Shape9.Visible = False
Label1.Caption = "1"
End If
End If
End Sub

Private Sub Timer2_Timer()
If Shape1.Left < 0 Then
Label2.Caption = "2"
End If
If Shape1.Left > 4300 Then
Label2.Caption = "1"
End If
If Shape1.Top < 0 Then
Label1.Caption = "1"
End If
If Shape1.Top > 6500 Then
Label1.Caption = "2"
End If
If Shape2.Visible = False And Shape3.Visible = False And Shape4.Visible = False And Shape5.Visible = False And Shape6.Visible = False And Shape7.Visible = False And Shape8.Visible = False And Shape9.Visible = False Then
Shape2.Visible = True
Shape3.Visible = True
Shape4.Visible = True
Shape5.Visible = True
Shape6.Visible = True
Shape7.Visible = True
Shape8.Visible = True
Shape9.Visible = True
End If
End Sub

Private Sub Timer3_Timer()
If Combo1.Text = "ãÈÊÏìÁ" Then
Timer1.Interval = 150
End If
If Combo1.Text = "ãÊæÓØ" Then
Timer1.Interval = 80
End If
If Combo1.Text = "ãÊÞÏã" Then
Timer1.Interval = 40
End If
If Combo1.Text = "ãÍÊÑÝ" Then
Timer1.Interval = 20
End If
If Shape1.Top > 5500 Then
Timer1.Interval = 0
End If
If Val(Label7.Caption) < 0 Then
End
End If
End Sub
