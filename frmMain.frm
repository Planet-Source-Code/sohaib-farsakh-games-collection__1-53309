VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6450
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   5
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C00000&
      Height          =   4500
      Left            =   240
      ScaleHeight     =   27
      ScaleMode       =   0  'User
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   360
      Width           =   4500
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   " Apples:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   " lives:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   " level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   " score:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Left            =   4800
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu action 
         Caption         =   "Start"
         Shortcut        =   {F3}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu high 
         Caption         =   "Highscores..."
      End
      Begin VB.Menu loadgraphics 
         Caption         =   "Load Additional Graphics..."
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu gamemode 
         Caption         =   "Game Mode"
         Begin VB.Menu complete 
            Caption         =   "Completing Levels"
         End
         Begin VB.Menu crash 
            Caption         =   "Playing Until Crash"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu gowall 
         Caption         =   "Go through walls"
         Checked         =   -1  'True
      End
      Begin VB.Menu seperate1 
         Caption         =   "-"
      End
      Begin VB.Menu speed 
         Caption         =   "Speed"
         Begin VB.Menu slow 
            Caption         =   "Slow"
         End
         Begin VB.Menu normal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu fast 
            Caption         =   "Fast"
         End
         Begin VB.Menu insane 
            Caption         =   "Insane!"
         End
      End
      Begin VB.Menu maze 
         Caption         =   "Maze"
         Begin VB.Menu mazestyle 
            Caption         =   "None"
            Index           =   0
         End
      End
      Begin VB.Menu increase 
         Caption         =   "Increase in length"
         Begin VB.Menu amount 
            Caption         =   "2"
            Index           =   0
         End
      End
      Begin VB.Menu seperate2 
         Caption         =   "-"
      End
      Begin VB.Menu gotolvl 
         Caption         =   "Go to Level..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub action_Click()
Label1_Click
End Sub

Private Sub amount_Click(Index As Integer)
For i = 0 To 5
amount(i).Checked = False
Next
amount(Index).Checked = True
lengthincrease = Val(amount(Index).Caption)
End Sub

Private Sub Command1_Click()
Load frmPic
frmPic.Show
End Sub

Private Sub complete_Click()
If complete.Checked = False Then newgame
complete.Checked = True
crash.Checked = False
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
speed.Enabled = False
maze.Enabled = False
increase.Enabled = False
gotolvl.Enabled = True
level = 1
livescore = 200
gotolevel (1)
frmMain.Label6.Caption = "3"
score = 0
Label3.Caption = "0"
End Sub

Private Sub crash_Click()

If slow.Checked = True Then s_interval = 180
If normal.Checked = True Then s_interval = 75
If fast.Checked = True Then s_interval = 36
If insane.Checked = True Then s_interval = 15

If mazestyle(0).Checked = True Then mazenum = 0
If mazestyle(1).Checked = True Then mazenum = 1
If mazestyle(2).Checked = True Then mazenum = 2
If mazestyle(3).Checked = True Then mazenum = 3
If mazestyle(4).Checked = True Then mazenum = 4

If amount(0).Checked = True Then lengthincrease = 2
If amount(1).Checked = True Then lengthincrease = 3
If amount(2).Checked = True Then lengthincrease = 5
If amount(3).Checked = True Then lengthincrease = 8
If amount(4).Checked = True Then lengthincrease = 12
If amount(5).Checked = True Then lengthincrease = 17
If crash.Checked = False Then newgame

complete.Checked = False
crash.Checked = True
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
speed.Enabled = True
maze.Enabled = True
increase.Enabled = True
gotolvl.Enabled = False
score = 0
Label3.Caption = "0"

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub fast_Click()
slow.Checked = False
normal.Checked = False
fast.Checked = True
insane.Checked = False
s_interval = 36

End Sub

Private Sub Form_Activate()
drawframe
End Sub


Private Sub Form_Load()
framecolor = vbBlue

newgame
For i = 1 To 4
Load mazestyle(i)
mazestyle(i).Visible = True
mazestyle(i).Caption = "Maze" + Str(i)
Next
mazestyle(0).Checked = True
lengthincrease = 3
For i = 1 To 5
Load amount(i)
amount(i).Visible = True
Next
amount(1).Caption = "3"
amount(2).Caption = "5"
amount(3).Caption = "8"
amount(4).Caption = "12"
amount(5).Caption = "17"

amount(1).Checked = True

LoadHighScores

password(2) = "yurf"
password(3) = "trwa"
password(4) = "nfra"
password(5) = "byuj"
password(6) = "vtdf"
password(7) = "glsw"
password(8) = "gtxm"
password(9) = "mcty"
password(10) = "vtsk"
s_interval = 75

End Sub
Public Sub LoadHighScores()
On Error Resume Next
Dim i As Integer, strName As String, strScore As String
Load frmscores
Open App.Path & ".\data.hst" For Input As #1
For i = 0 To 19
  Input #1, strName, strScore
  a = strName
  b = strScore
  If i < 10 Then
  scorescomplete(i, 1) = a
  scorescomplete(i, 2) = b
  Else
  scorescrash(i - 10, 1) = a
  scorescrash(i - 10, 2) = b
  End If
Next i
Close #1
End Sub
Private Sub SaveHighScores()
On Error Resume Next
Dim i As Integer
Open App.Path & ".\data.hst" For Output As #1
For i = 0 To 9
  Write #1, scorescomplete(i, 1), scorescomplete(i, 2)
Next i
For i = 0 To 9
  Write #1, scorescrash(i, 1), scorescrash(i, 2)
Next i
Close #1
Unload frmscores
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveHighScores
End
End Sub

Private Sub gotolvl_Click()
Load frmlevel
frmlevel.Show , frmMain
frmlevel.Text1.Text = ""
frmlevel.Text2.Text = ""
End Sub

Private Sub gowall_Click()
gowall.Checked = Not (gowall.Checked)
End Sub

Private Sub high_Click()
frmscores.Show , frmMain
If complete.Checked = True Then
frmscores.Option1.Value = True
displayscores ("complete")
Else
frmscores.Option2.Value = True
displayscores ("crash")
End If
End Sub

Private Sub insane_Click()
slow.Checked = False
normal.Checked = False
fast.Checked = False
insane.Checked = True
s_interval = 15

End Sub

Private Sub Label1_Click()
Select Case Label1.Caption
Case "Start"
Label1.Caption = "Pause"
Delay (0.4)
gamerun = True
Case "Restart"
If crash.Checked = True Then
newgame
Else
complete_Click
End If
Case "Pause"
Label1.Caption = "Resume"
gamerun = False
Case "Resume"
Label1.Caption = "Pause"
Delay (0.4)
gamerun = True
End Select
action.Caption = Label1.Caption

End Sub

Private Sub loadgraphics_Click()
CommonDialog1.Filter = "Bmp Files|*.bmp| Gif Files|*.gif| All Files|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
frmPic.Picture2.Picture = LoadPicture(CommonDialog1.FileName)
'frmPic.Picture2.Picture = frmPic.Picture2.Image
For i = 0 To 9
frmPic.Picture1(i).PaintPicture frmPic.Picture2.Picture, 0, 0, 150, 150, i * 150, 0, 150, 150
frmPic.Picture1(i).Picture = frmPic.Picture1(i).Image
frmPic.Picture1(i).Height = 210
frmPic.Picture1(i).Width = 210
Next
drawsnake
Load frmPic
frmPic.Refresh
Call Module1.applymaze(mazenum)
frmMain.Picture1.PaintPicture frmPic.Picture1(pictureindex).Picture, foodx, foody, 1, 1
frmMain.BackColor = frmPic.Picture2.Point(1501, 1)
Label1.BackColor = frmPic.Picture2.Point(1516, 1)
Label10.BackColor = frmPic.Picture2.Point(1516, 1)
framecolor = (frmPic.Picture2.Point(1531, 1))
drawframe
End Sub

Private Sub mazestyle_Click(Index As Integer)
d = MsgBox("This will start a new game, continue?", vbYesNo)
If d = vbNo Then Exit Sub
For i = 0 To 4
mazestyle(i).Checked = False
Next
mazestyle(Index).Checked = True
'If Index <> 0 Then applymaze (Index)
mazenum = Index
newgame
End Sub

Private Sub new_Click()
If crash.Checked = True Then
newgame
Else
complete_Click
End If
End Sub

Private Sub normal_Click()
slow.Checked = False
normal.Checked = True
fast.Checked = False
insane.Checked = False
s_interval = 75

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
If lastdir <> "down" Then direction = "up"
Case vbKeyDown
If lastdir <> "up" Then direction = "down"
Case vbKeyRight
If lastdir <> "left" Then direction = "right"
Case vbKeyLeft
If lastdir <> "right" Then direction = "left"
Case vbKeyEscape
End
Case vbKeyPause, vbKeySpace
Label1_Click
End Select

End Sub

Private Sub slow_Click()
slow.Checked = True
normal.Checked = False
fast.Checked = False
insane.Checked = False
s_interval = 180
End Sub

