VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rods"
   ClientHeight    =   7470
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   35
      Left            =   360
      TabIndex        =   38
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   34
      Left            =   720
      TabIndex        =   37
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   33
      Left            =   1080
      TabIndex        =   36
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   32
      Left            =   1440
      TabIndex        =   35
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   31
      Left            =   1800
      TabIndex        =   34
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   30
      Left            =   2160
      TabIndex        =   33
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   29
      Left            =   2520
      TabIndex        =   32
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   28
      Left            =   2880
      TabIndex        =   31
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   27
      Left            =   3240
      TabIndex        =   30
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   26
      Left            =   3600
      TabIndex        =   29
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   25
      Left            =   3960
      TabIndex        =   28
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   24
      Left            =   4320
      TabIndex        =   27
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   23
      Left            =   4680
      TabIndex        =   26
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   22
      Left            =   5040
      TabIndex        =   25
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   21
      Left            =   5400
      TabIndex        =   24
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   20
      Left            =   5760
      TabIndex        =   23
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   19
      Left            =   6120
      TabIndex        =   22
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   18
      Left            =   6480
      TabIndex        =   21
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Blue player turn"
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   17
      Left            =   6480
      TabIndex        =   18
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   16
      Left            =   6120
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   15
      Left            =   5760
      TabIndex        =   16
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   14
      Left            =   5400
      TabIndex        =   15
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   13
      Left            =   5040
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   12
      Left            =   4680
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   11
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   10
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   9
      Left            =   3600
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   8
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   7
      Left            =   2880
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   5
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   4
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   3135
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu one 
         Caption         =   "One Player"
         Begin VB.Menu easy 
            Caption         =   "Easy"
         End
         Begin VB.Menu normal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu diff 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu two 
         Caption         =   "Two Players"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu set 
         Caption         =   "Game Setup"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim number As Integer
Private Sub gamelost()
If Form1.Command1.Tag = 1 Then
If Label3.Caption = "Blue player turn" Then
d = MsgBox("Red Wins!")
Else
d = MsgBox("Blue Wins!")
End If
End If

End Sub
Private Sub draw(ByVal num As Integer)
drawed = 0
For j = 0 To 35
If Label2(j).Enabled = True Then
If Val(Form1.Label1.Tag) < num And Label2(j).BackColor = vbYellow Then
If Form1.Tag = True Then
Label2(j).BackColor = vbBlue
Else
Label2(j).BackColor = vbRed
End If
Form1.Label1.Tag = Form1.Label1.Tag + 1
Form1.Command1.Tag = Form1.Command1.Tag - 1
Label2(j).Enabled = False
End If
If Form1.Label1.Tag = num Then
If Form1.Tag = False Then
Form1.Tag = True
Else
Form1.Tag = False
End If
Exit Sub
End If
End If
Next j
End Sub



Private Sub play()
Label3.Caption = "Blue Player turn"
gamelost
If easy.Checked = True Then n = Val(Form2.Text2.Text + 1) + 1
If normal.Checked = True Then n = Int(2.2 * Val(Form2.Text2.Text + 1)) + 1
If diff.Checked = True Then n = Form2.Text1.Text
If Val(Command1.Tag) > n Then
Randomize
d = Int(Rnd * Val(Form2.Text2.Text)) + 1
draw (d)
Else
a = (Val(Command1.Tag) - 1) Mod (Val(Form2.Text2.Text) + 1)
If a = 0 Then a = 1
draw (a)
End If
End Sub


Private Sub Command1_Click()
If Form1.Tag = False Then
Form1.Tag = True
Else
Form1.Tag = False
End If
Form1.Label1.Tag = 0
If two.Checked = False Then
play
End If
Command1.Enabled = False
Form1.Label1.Tag = 0
If Label3.Caption = "Blue player turn" Then
Label3.Caption = "Red player turn"
Else
Label3.Caption = "Blue player turn"
End If
gamelost
End Sub

Private Sub diff_Click()
easy.Checked = False
normal.Checked = False
diff.Checked = True
two.Checked = False

End Sub

Private Sub easy_Click()
easy.Checked = True
normal.Checked = False
diff.Checked = False
two.Checked = False
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Label1.Height = 3400
Form1.Height = Form1.Height - 3400
Form1.Tag = True
Form1.Label1.Tag = 0
For i = 1 To 35
Label2(i).BackColor = vbYellow
Next i
For i = 18 To 35
Label2(i).Visible = False
Next i
Command1.Top = Command1.Top - 3300
Label3.Top = Label3.Top - 3300
Form1.Command1.Tag = 18
Form1.Label3.Tag = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Label2_Click(Index As Integer)
If Form1.Label1.Tag < Form1.Label3.Tag And Label2(Index).BackColor = vbYellow Then
If Form1.Tag = True Then
Label2(Index).BackColor = vbBlue
Else
Label2(Index).BackColor = vbRed
End If
Form1.Label1.Tag = Form1.Label1.Tag + 1
Form1.Command1.Tag = Form1.Command1.Tag - 1
Label2(Index).Enabled = False
Command1.Enabled = True
Else
d = MsgBox("Sorry,you can't take any more rods!")
End If
End Sub

Private Sub new_Click()
Call newgame
End Sub

Private Sub normal_Click()
easy.Checked = False
normal.Checked = True
diff.Checked = False
two.Checked = False
End Sub


Private Sub set_Click()
Load Form2
Form2.Show vbModal
End Sub

Private Sub two_Click()
easy.Checked = False
normal.Checked = False
diff.Checked = False
two.Checked = True

End Sub
