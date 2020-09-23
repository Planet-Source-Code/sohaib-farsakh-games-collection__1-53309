VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ßÓÑ ÇáãÑÈÚÇÊ"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   Tag             =   "6600"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Tag             =   "1"
      Top             =   7560
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇáãÑÍáÉ:"
      Height          =   375
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   7680
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   76
      Left            =   6600
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   75
      Left            =   5520
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   74
      Left            =   4440
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   73
      Left            =   3360
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   72
      Left            =   2280
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   71
      Left            =   1200
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   70
      Left            =   120
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   69
      Left            =   7680
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   68
      Left            =   10920
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   67
      Left            =   8760
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   66
      Left            =   9840
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   65
      Left            =   6600
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   64
      Left            =   5520
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   63
      Left            =   4440
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   62
      Left            =   3360
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   61
      Left            =   2280
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   60
      Left            =   1200
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   59
      Left            =   120
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   58
      Left            =   7680
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   57
      Left            =   10920
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   56
      Left            =   8760
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   55
      Left            =   9840
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   54
      Left            =   6600
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   53
      Left            =   5520
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   52
      Left            =   4440
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   51
      Left            =   3360
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   50
      Left            =   2280
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   49
      Left            =   1200
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   48
      Left            =   120
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   47
      Left            =   7680
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   46
      Left            =   10920
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   45
      Left            =   8760
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   405
      Index           =   44
      Left            =   9840
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   375
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇáßÑÇÊ ÇáÅÖÇÝíÉ:"
      Height          =   375
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   495
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇáÏÑÌÉ:"
      Height          =   495
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   7680
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   43
      Left            =   9840
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   42
      Left            =   8760
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   41
      Left            =   10920
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   40
      Left            =   1200
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   39
      Left            =   2280
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   38
      Left            =   3360
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   37
      Left            =   4440
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   36
      Left            =   8760
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   35
      Left            =   1200
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   34
      Left            =   2280
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   33
      Left            =   3360
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   32
      Left            =   4440
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   31
      Left            =   5520
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   30
      Left            =   5520
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   29
      Left            =   7680
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   28
      Left            =   1200
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   27
      Left            =   2280
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   26
      Left            =   3360
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   25
      Left            =   4440
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   24
      Left            =   5520
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   23
      Left            =   6600
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   22
      Left            =   6600
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   21
      Left            =   6600
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   20
      Left            =   120
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   19
      Left            =   120
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   18
      Left            =   7680
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   17
      Left            =   7680
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   16
      Left            =   7680
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   15
      Left            =   8760
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   14
      Left            =   8760
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   13
      Left            =   9840
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   12
      Left            =   10920
      Top             =   960
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   11
      Left            =   9840
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   10
      Left            =   9840
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   9
      Left            =   10920
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   8
      Left            =   10920
      Top             =   480
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   7
      Left            =   120
      Top             =   0
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   6
      Left            =   120
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   5
      Left            =   1200
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   4
      Left            =   2280
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   3
      Left            =   3360
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   2
      Left            =   4440
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   1
      Left            =   5520
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   0
      Left            =   6600
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   5520
      Shape           =   3  'Circle
      Tag             =   "1"
      Top             =   6840
      Width           =   150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   6
      Tag             =   "40"
      X1              =   4560
      X2              =   6600
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Menu options 
      Caption         =   "&ÎíÇÑÇÊ"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
      Begin VB.Menu new 
         Caption         =   "&ÌÏíÏ"
         Shortcut        =   ^N
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu slow 
         Caption         =   "&ÈØíÁ"
      End
      Begin VB.Menu medium 
         Caption         =   "&ãÊæÓØ"
      End
      Begin VB.Menu fast 
         Caption         =   "&ÓÑíÚ"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu pause 
         Caption         =   "&ÅíÞÇÝ"
         Shortcut        =   {F2}
      End
      Begin VB.Menu resume 
         Caption         =   "&ÅßãÇá"
         Shortcut        =   {F3}
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu demo 
         Caption         =   "&áÚÈ ÇáßãÈíæÊÑ"
      End
      Begin VB.Menu stop 
         Caption         =   "&ÅíÞÇÝ áÚÈ ÇáßãÈíæÊÑ"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub demo_Click()
Label2.Tag = 2
End Sub

Private Sub fast_Click()
Timer1.Interval = 1
End Sub

Private Sub Form_Click()
Select Case Timer1.Enabled
Case False
Timer1.Enabled = True
Case True
Timer1.Enabled = False
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.Tag = 1 Then
Line1.X1 = X - 750
Line1.X2 = X + 750
End If
End Sub

Private Sub medium_Click()
Timer1.Interval = 60
End Sub

Private Sub new_Click()
For a = 0 To 76 Step 1
Shape2(a).Visible = True
Next a
Timer1.Enabled = False
Label2.Caption = "0"
Line1.Tag = 40
Shape1.Tag = 1
Timer1.Tag = 1
Label4.Caption = "3"
Shape1.Top = Line1.Y1 - 200
Shape1.Left = (Line1.X1 + Line1.X2) / 2
Label6.Caption = "1"
For X = 0 To 76
Shape2(X).FillColor = &H80
Next X
End Sub

Private Sub pause_Click()
Timer1.Enabled = False
End Sub

Private Sub resume_Click()
Timer1.Enabled = True
End Sub

Private Sub slow_Click()
Timer1.Interval = 250
End Sub

Private Sub stop_Click()
Label2.Tag = 1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label4.Tag = 0
For X = 0 To 76 Step 1
If Shape2(X).Visible = False Then
Label4.Tag = Val(Label4.Tag) + 1
End If
Next X
If Label4.Tag >= 77 Then
For a = 0 To 76 Step 1
Shape2(a).Visible = True
Next a
Label6.Caption = Val(Label6.Caption) + 1
For X = 0 To 76
Shape2(X).FillColor = &H80
Next X
End If
If Label6.Caption = "4" Then
Label6.Caption = "3"
Timer1.Enabled = False
End If
Select Case Shape1.Tag
Case 1
Shape1.Top = Shape1.Top - (7000 / Line1.Tag)
Case 2
Shape1.Top = Shape1.Top + (7000 / Line1.Tag)
End Select
Select Case Timer1.Tag
Case 1
Shape1.Left = Shape1.Left + Line1.Tag
Case 2
Shape1.Left = Shape1.Left - Line1.Tag
End Select
If Shape1.Top <= 0 Then
Shape1.Tag = 2
End If
If Shape1.Left <= 0 Then
Timer1.Tag = 1
End If
If Shape1.Left >= Form1.Width - 150 Then
Timer1.Tag = 2
End If
On Error Resume Next
If Shape1.Top >= Line1.Y1 - 160 And Shape1.Left >= Line1.X1 - 75 And Shape1.Left <= Line1.X2 - 75 Then
Shape1.Tag = 1
If Label2.Tag = 1 Then
a = Shape1.Left - ((Line1.X1 + Line1.X2) / 2)
a = Abs(a)
Line1.Tag = Sqr(a) * 5
End If
End If
For i = 0 To 76 Step 1
If Shape1.Top >= Shape2(i).Top - 150 And Shape1.Top <= Shape2(i).Top + 400 And Shape1.Left >= Shape2(i).Left - 150 And Shape1.Left <= Shape2(i).Left + 1000 And Shape2(i).Visible = True Then
If Label6.Caption = "1" Then
Shape2(i).Visible = False
End If
If Label6.Caption = "2" Then
Select Case Shape2(i).FillColor
Case &H80&
Shape2(i).FillColor = &H8080FF
Case &H8080FF
Shape2(i).Visible = False
End Select
End If
If Label6.Caption = "3" Then
Select Case Shape2(i).FillColor
Case &H80&
Shape2(i).FillColor = &HFF&
Case &HFF&
Shape2(i).FillColor = &H8080FF
Case &H8080FF
Shape2(i).FillColor = &HC0C0FF
Case &HC0C0FF
Shape2(i).Visible = False
End Select
End If
Shape1.Tag = 2

If Label2.Tag = 1 Then
Label2.Caption = Val(Label2.Caption) + 10

If Val(Label2.Caption) / 1000 = Int(Val(Label2.Caption) / 1000) Then
Label4.Caption = Val(Label4.Caption) + 1
End If
End If
End If
Next i
If Shape1.Top > Line1.Y1 + 400 Then
Label4.Caption = Val(Label4.Caption) - 1
Shape1.Top = 4500
Timer1.Enabled = False
End If
If Label2.Tag = 2 Then
Line1.Tag = 40
Line1.X1 = Shape1.Left - 750
Line1.X2 = Line1.X1 + 1500
End If
If Val(Label4.Caption) < 0 Then End

End Sub

