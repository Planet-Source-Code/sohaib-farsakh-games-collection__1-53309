VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rods"
   ClientHeight    =   4170
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Blue player turn"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   3720
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
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
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim number As Integer
Dim first As Boolean
Dim rods As Integer
Private Sub gamelost()
If rods = 1 Then
If Label3.Caption = "Blue player turn" Then
d = MsgBox("Red Wins!")
Else
d = MsgBox("Blue Wins!")
End If
End If

End Sub
Private Sub drawone()
drawed = 0
For j = 0 To 17
If Label2(j).Enabled = True Then
If number < 3 And Label2(j).BackColor = vbYellow < 3 Then
If first = True Then
Label2(j).BackColor = vbBlue
Else
Label2(j).BackColor = vbRed
End If
number = number + 1
rods = rods - 1
Label2(j).Enabled = False
End If
If number = 1 Then
first = Not (first)
Exit Sub
End If
End If
Next j
End Sub
Private Sub drawtwo()
drawed = 0
For j = 0 To 17
If Label2(j).Enabled = True Then
If number < 3 And Label2(j).BackColor = vbYellow < 3 Then
If first = True Then
Label2(j).BackColor = vbBlue
Else
Label2(j).BackColor = vbRed
End If
number = number + 1
rods = rods - 1
Label2(j).Enabled = False
End If
If number = 2 Then
first = Not (first)
Exit Sub
End If
End If
Next j
End Sub
Private Sub drawthree()
drawed = 0
For j = 0 To 17
If Label2(j).Enabled = True Then
If number < 3 And Label2(j).BackColor = vbYellow < 3 Then
If first = True Then
Label2(j).BackColor = vbBlue
Else
Label2(j).BackColor = vbRed
End If
number = number + 1
rods = rods - 1
Label2(j).Enabled = False
End If
If number = 3 Then
first = Not (first)
Exit Sub
End If
End If
Next j
End Sub



Private Sub play()
Label3.Caption = "Blue Player turn"
gamelost
If easy.Checked = True Then
Select Case rods
Case 4
drawthree
Case 3
drawtwo
Case 2
drawone
Case 1
drawone
Case 5
drawone
Case Else
Randomize
a = Int(Rnd * 3)
If a = 0 Then drawone
If a = 1 Then drawtwo
If a = 2 Then drawthree
End Select
End If

If normal.Checked = True Then
Select Case rods
Case 4, 8
drawthree
Case 3, 7
drawtwo
Case 2, 6, 10
drawone
Case 1
drawone
Case 5, 9
drawone
Case Else
Randomize
a = Int(Rnd * 3)
If a = 0 Then drawone
If a = 1 Then drawtwo
If a = 2 Then drawthree
End Select
End If


If diff.Checked = True Then
Select Case rods
Case 4, 8, 12, 16
drawthree
Case 3, 7, 11, 15
drawtwo
Case 2, 6, 10, 14, 18
drawone
Case 1
drawone
Case 5, 9, 13, 17
drawone
Case Else
Randomize
a = Int(Rnd * 3)
If a = 0 Then drawone
If a = 1 Then drawtwo
If a = 2 Then drawthree
End Select
End If
End Sub


Private Sub Command1_Click()
first = Not (first)
number = 0
If two.Checked = False Then
play
End If
Command1.Enabled = False
number = 0
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
first = True
number = 0
For i = 1 To 17
Label2(i).BackColor = vbYellow
Next i
rods = 18
End Sub

Private Sub Label2_Click(Index As Integer)
If number < 3 And Label2(Index).BackColor = vbYellow < 3 Then
If first = True Then
Label2(Index).BackColor = vbBlue
Else
Label2(Index).BackColor = vbRed
End If
number = number + 1
rods = rods - 1
Label2(Index).Enabled = False
Command1.Enabled = True
Else
d = MsgBox("Sorry,you can't take any more rods!")
End If
End Sub

Private Sub new_Click()
first = True
number = 0
For i = 0 To 17
Label2(i).BackColor = vbYellow
Label2(i).Enabled = True
Next i
rods = 18
Label3.Caption = "Blue player turn"
End Sub

Private Sub normal_Click()
easy.Checked = False
normal.Checked = True
diff.Checked = False
two.Checked = False

End Sub

Private Sub Timer1_Timer()
gamelost
End Sub

Private Sub two_Click()
easy.Checked = False
normal.Checked = False
diff.Checked = False
two.Checked = True

End Sub
