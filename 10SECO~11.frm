VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "10 Seconds"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Timer tmrEndGame 
      Interval        =   100
      Left            =   5760
      Top             =   4200
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   3960
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   3960
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   3960
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   2640
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   2640
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   1320
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare countdown variables for each button.  These will begin at 10
' and countdown by one with each buttons timer event.
Dim fiCount1 As Integer
Dim fiCount2 As Integer
Dim fiCount3 As Integer
Dim fiCount4 As Integer
Dim fiCount5 As Integer
Dim fiCount6 As Integer
Dim fiCount7 As Integer
Dim fiCount8 As Integer
Dim fiCount9 As Integer

'Declare countdown speed variables for each button.  Each button will have its
'own random countdown speed which will decrease when the button is clicked.
Dim fiSpeed1 As Integer
Dim fiSpeed2 As Integer
Dim fiSpeed3 As Integer
Dim fiSpeed4 As Integer
Dim fiSpeed5 As Integer
Dim fiSpeed6 As Integer
Dim fiSpeed7 As Integer
Dim fiSpeed8 As Integer
Dim fiSpeed9 As Integer

'Declare variable for game score
Dim fiScore As Integer
Dim fiNumLeft As Integer


'Declare variable to check that same button has not been pressed twice so as to
'disable cheating by just clicking repeatedly on the button.
Dim fiLastPressed As Integer


Private Sub cmd1_Click()
If fiLastPressed <> 1 Then
fiCount1 = 10
cmd1.Caption = 10
Timer1.Interval = Timer1.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 1
End If
End Sub

Private Sub cmd2_Click()
If fiLastPressed <> 2 Then
fiCount2 = 10
cmd2.Caption = 10
Timer2.Interval = Timer2.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 2
End If
End Sub

Private Sub cmd3_Click()
If fiLastPressed <> 3 Then
fiCount3 = 10
cmd3.Caption = 10
Timer3.Interval = Timer3.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 3
End If

End Sub

Private Sub cmd4_Click()
If fiLastPressed <> 4 Then
fiCount4 = 10
cmd4.Caption = 10
Timer4.Interval = Timer4.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 4
End If
End Sub

Private Sub cmd5_Click()
If fiLastPressed <> 5 Then
fiCount5 = 10
cmd5.Caption = 10
Timer5.Interval = Timer5.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 5
End If
End Sub

Private Sub cmd6_Click()
If fiLastPressed <> 6 Then
fiCount6 = 10
cmd6.Caption = 10
Timer6.Interval = Timer6.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 6
End If
End Sub

Private Sub cmd7_Click()
If fiLastPressed <> 7 Then
fiCount7 = 10
cmd7.Caption = 10
Timer7.Interval = Timer7.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 7
End If
End Sub

Private Sub cmd8_Click()
If fiLastPressed <> 8 Then
fiCount8 = 10
cmd8.Caption = 10
Timer8.Interval = Timer8.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 8
End If
End Sub

Private Sub cmd9_Click()
If fiLastPressed <> 9 Then
fiCount9 = 10
cmd9.Caption = 10
Timer9.Interval = Timer9.Interval - 50
fiScore = fiScore + 10
txtScore.Text = Str(fiScore)
fiLastPressed = 9
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()

cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = True
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = True
cmd7.Enabled = True
cmd8.Enabled = True
cmd9.Enabled = True


fiCount1 = 11
fiCount2 = 11
fiCount3 = 11
fiCount4 = 11
fiCount5 = 11
fiCount6 = 11
fiCount7 = 11
fiCount8 = 11
fiCount9 = 11
fiScore = 0
txtScore.Text = "0"
fiLastPressed = 0
fiNumLeft = 8

'set initial start speeds for each timer
Timer1.Interval = 1000
Timer2.Interval = 1000
Timer3.Interval = 1000
Timer4.Interval = 1000
Timer5.Interval = 1000
Timer6.Interval = 1000
Timer7.Interval = 1000
Timer8.Interval = 1000
Timer9.Interval = 1000

'Start all timers counting down
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
tmrEndGame.Enabled = True


End Sub

Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
fiCount1 = 10
fiCount2 = 10
fiCount3 = 10
fiCount4 = 10
fiCount5 = 10
fiCount6 = 10
fiCount7 = 10
fiCount8 = 10
fiCount9 = 10
fiScore = 0
txtScore.Text = "0"
fiLastPressed = 0
fiNumLeft = 8

'set initial start speeds for each timer
Timer1.Interval = 1000
Timer2.Interval = 1000
Timer3.Interval = 1000
Timer4.Interval = 1000
Timer5.Interval = 1000
Timer6.Interval = 1000
Timer7.Interval = 1000
Timer8.Interval = 1000
Timer9.Interval = 1000

End Sub

Private Sub Timer1_Timer()
If fiCount1 = 0 Then
    Timer1.Enabled = False
    cmd1.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount1 = fiCount1 - 1
cmd1.Caption = Str(fiCount1)
End If
End Sub

Private Sub Timer2_Timer()
If fiCount2 = 0 Then
    Timer2.Enabled = False
    cmd2.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount2 = fiCount2 - 1
cmd2.Caption = Str(fiCount2)
End If
End Sub

Private Sub Timer3_Timer()
If fiCount3 = 0 Then
    Timer3.Enabled = False
    cmd3.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount3 = fiCount3 - 1
cmd3.Caption = Str(fiCount3)
End If
End Sub

Private Sub Timer4_Timer()
If fiCount4 = 0 Then
    Timer4.Enabled = False
    cmd4.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount4 = fiCount4 - 1
cmd4.Caption = Str(fiCount4)
End If
End Sub

Private Sub Timer5_Timer()
If fiCount5 = 0 Then
    Timer5.Enabled = False
    cmd5.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount5 = fiCount5 - 1
cmd5.Caption = Str(fiCount5)
End If
End Sub

Private Sub Timer6_Timer()
If fiCount6 = 0 Then
    Timer6.Enabled = False
    cmd6.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount6 = fiCount6 - 1
cmd6.Caption = Str(fiCount6)
End If
End Sub

Private Sub Timer7_Timer()
If fiCount7 = 0 Then
    Timer7.Enabled = False
    cmd7.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount7 = fiCount7 - 1
cmd7.Caption = Str(fiCount7)
End If
End Sub

Private Sub Timer8_Timer()
If fiCount8 = 0 Then
    Timer8.Enabled = False
    cmd8.Enabled = False
    fiNumLeft = fiNumLeft - 1
Else
fiCount8 = fiCount8 - 1
cmd8.Caption = Str(fiCount8)
End If
End Sub

Private Sub Timer9_Timer()
If fiCount9 = 0 Then
    Timer9.Enabled = False
    cmd9.Enabled = False
    fiNumLeft = fiNumLeft - 1
    
Else
fiCount9 = fiCount9 - 1
cmd9.Caption = Str(fiCount9)
End If
End Sub



Private Sub tmrEndGame_Timer()
If fiNumLeft = 0 Then
cmd1.Caption = "X"
cmd2.Caption = "X"
cmd3.Caption = "X"
cmd4.Caption = "X"
cmd5.Caption = "X"
cmd6.Caption = "X"
cmd7.Caption = "X"
cmd8.Caption = "X"
cmd9.Caption = "X"

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False

frmGameOver.Visible = True
tmrEndGame.Enabled = False
End If
End Sub
