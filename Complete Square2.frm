VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00007F00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complete Square V.3.16"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9480
   Icon            =   "Complete Square2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00007F00&
      Caption         =   "Back to the game"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   1200
      Top             =   6240
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   720
      Top             =   6240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      Caption         =   "Players"
      Height          =   2415
      Left            =   7440
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   1180
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Difficult"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Easy"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "V.Easy"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "V. Easy"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "next turn:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         Height          =   255
         Left            =   -120
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Difficult"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Easy"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   40
         TabIndex        =   12
         Top             =   460
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Score"
      Height          =   2535
      Left            =   7440
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "wins by red"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "wins by blue"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   310
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Label27"
      Height          =   495
      Left            =   480
      TabIndex        =   31
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu pauseresume 
         Caption         =   "Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu suggest 
         Caption         =   "Suggest Move"
         Shortcut        =   ^M
      End
      Begin VB.Menu undo 
         Caption         =   "Undo Move"
         Shortcut        =   ^Z
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu blue 
         Caption         =   "Blue player"
         Begin VB.Menu human1 
            Caption         =   "Human"
            Checked         =   -1  'True
         End
         Begin VB.Menu very1 
            Caption         =   "Very Easy"
         End
         Begin VB.Menu easy1 
            Caption         =   "Easy"
         End
         Begin VB.Menu normal1 
            Caption         =   "Normal"
         End
         Begin VB.Menu difficult1 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu red 
         Caption         =   "Red player"
         Begin VB.Menu human2 
            Caption         =   "Human"
         End
         Begin VB.Menu very2 
            Caption         =   "Very Easy"
         End
         Begin VB.Menu easy2 
            Caption         =   "Easy"
         End
         Begin VB.Menu normal2 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu difficult2 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu separateline 
         Caption         =   "-"
      End
      Begin VB.Menu speed 
         Caption         =   "Computer Playing Speed"
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
      End
      Begin VB.Menu animate 
         Caption         =   "Animate Colors"
      End
      Begin VB.Menu seperateline2 
         Caption         =   "-"
      End
      Begin VB.Menu gridsize 
         Caption         =   "Grid Size"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu intetnet 
      Caption         =   "Internet"
      Begin VB.Menu host 
         Caption         =   "Host a game"
      End
      Begin VB.Menu join 
         Caption         =   "Join"
      End
      Begin VB.Menu chat 
         Caption         =   "Show Chat Window"
      End
      Begin VB.Menu seperateline3 
         Caption         =   "-"
      End
      Begin VB.Menu endnet 
         Caption         =   "End Internet Mode"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu contenents 
         Caption         =   "Playing the game"
         Shortcut        =   {F1}
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim turn As String
Dim grid(12, 12) As Integer
Dim grid2(12, 12) As Integer
Dim recentwin As Boolean
Dim winned As Integer
Dim lbl2(160) As Boolean
Dim lbl3(160) As Boolean
Dim enabled1 As Boolean, enabled2 As Boolean
Dim lastblue As Integer, lastred As Integer
Dim firstbluemove As Boolean, firstredmove As Boolean
Dim allmoves(320, 2) As Variant
Dim movenum As Integer
Dim removedwins As Integer
Dim boardsize As Integer
Dim verlinenum As Integer, horlinenum As Integer
Dim loaded1 As Integer, loaded2 As Integer
Dim horlinewidth, verlineheight
Dim animatecolors As Boolean
Dim testing As Boolean, colorsteps As Long
Dim freezed As Boolean
Dim isexit As Boolean
Dim internet As Boolean, netstatus As String


Private Sub writehelp()
Dim words As String
words = "Complete Square is a strategy game for two players, each one tries to close as many squares as possible(closing a square means closing its last side), you close a side by clicking on it. There is a red player and a blue player, each one of them can be a human player or a computer player of any level of the four levels available(the very easy label is too easy to beat(anyone can beat it), but the difficult level needs an expert player to beat it!). When you take a square a circle of the player color appears in it, and the computer shows how many squares are taken by each player and how many times did each player win in the score frame, and it shows the players and turns in players frame, you can adjust the computer playing speed from options>computer playing speed, and you can adjust the grid size(number of squares on a side)from options>grid size and then type a number from(2-12) then click OK. the The Suggest Move command shows you a move according to the level you are playing with"
words = words + "(only if one player is human and the other is computer), the Undo Move command takes you back to a previous point in the game(doesn't work when both players are computers), F3 pauses and resumes the game especially when both players are computers, and you can enable or disable the animation in colors by clicking Options>Animate Colors. The game needs a screen resolution (640 by 480) or higher. Thank you for trying the game!"
Text1.Text = words

End Sub

Private Sub changecolor(label As label, color As Long)
On Error Resume Next
freezed = True
If animate.Checked = True Then
Dim oldcolor As Long
Dim curr As Integer, curg As Integer, curb As Integer
Dim R, G, B, R2, G2, B2 As Integer
oldcolor = label.BackColor
temp = (oldcolor And 255)
R = temp And 255
temp = Int(oldcolor / 256)
G = temp And 255
temp = Int(oldcolor / 65536)
B = temp And 255

temp = (color And 255)
R2 = temp And 255
temp = Int(color / 256)
G2 = temp And 255
temp = Int(color / 65536)
B2 = temp And 255
start = Timer
For i = 1 To colorsteps
curr = Int(R + ((R2 - R) / colorsteps * i))
curg = Int(G + ((G2 - G) / colorsteps * i))
curb = Int(B + ((B2 - B) / colorsteps * i))
label.BackColor = RGB(curr, curg, curb)
DoEvents
Delay (0.001)
Next i
duration = ((Timer - start) * 1000) / 1000
If testing = True Then
colorsteps = (0.4 * colorsteps) / duration
End If
Else
label.BackColor = color
End If
testing = False
freezed = False
End Sub

Private Sub ordinarylast(thecolor As String)
Select Case thecolor
Case "blue"
If lastblue < (horlinenum + 1) Then
Call changecolor(Label2(lastblue), &HFF0000)
Else
Call changecolor(Label3(lastblue - (horlinenum + 1)), &HFF0000)

End If
Case "red"
If lastred < (horlinenum + 1) Then
Call changecolor(Label2(lastred), &HFF&)
Else
Call changecolor(Label3(lastred - (horlinenum + 1)), &HFF&)
End If
End Select
End Sub
Sub Delay(secs)
Dim start
start = Timer
While (Timer < (start + secs))
DoEvents
Wend
End Sub

Private Sub saveenabled()
For i = 0 To horlinenum
lbl2(i) = Label2(i).Enabled
Label2(i).Enabled = False
Next i
For i = 0 To verlinenum
lbl3(i) = Label3(i).Enabled
Label3(i).Enabled = False
Next i
End Sub
Private Sub revertenabled()
For i = 0 To horlinenum
Label2(i).Enabled = lbl2(i)
Next i
For i = 0 To verlinenum
Label3(i).Enabled = lbl3(i)
Next i
End Sub

Private Sub endgame()
If Val(Label5.Caption) > Val(Label7.Caption) Then
MsgBox ("The blue player wins the game!")
Label16.Caption = Val(Label16.Caption) + 1
End If
If Val(Label5.Caption) < Val(Label7.Caption) Then
MsgBox ("The red player wins the game!")
Label18.Caption = Val(Label18.Caption) + 1
End If
If Val(Label5.Caption) = Val(Label7.Caption) Then
MsgBox ("The game is a draw")
End If
Timer1.Enabled = False
Timer2.Enabled = False
End Sub
Private Sub unmakemove(linenum As Integer)
removedwins = 0
Dim indexx As Integer
If linenum <= horlinenum Then
indexx = linenum
If Label2(indexx).Enabled = True Then Exit Sub
Label2(indexx).BackColor = &HC000&
Label2(indexx).Enabled = True
a = Int(Label2(indexx).index / boardsize) + 1
B = (Label2(indexx).index Mod boardsize) + 1
If a > 1 Then
If grid(a - 1, B) = 5 Then
grid(a - 1, B) = 4
removedwins = removedwins + 1
End If
grid(a - 1, B) = grid(a - 1, B) - 1
End If
If a < (boardsize + 1) Then
If grid(a, B) = 5 Then
grid(a, B) = 4
removedwins = removedwins + 1
End If
grid(a, B) = grid(a, B) - 1
End If



Else


indexx = linenum - (horlinenum + 1)
If Label3(indexx).Enabled = True Then Exit Sub
Label3(indexx).BackColor = &HC000&
Label3(indexx).Enabled = True

a = Int(Label3(indexx).index / (boardsize + 1)) + 1
B = (Label3(indexx).index Mod (boardsize + 1)) + 1
If B > 1 Then
If grid(a, B - 1) = 5 Then
grid(a, B - 1) = 4
removedwins = removedwins + 1
End If
grid(a, B - 1) = grid(a, B - 1) - 1
End If
If B < (boardsize + 1) Then
If grid(a, B) = 5 Then
grid(a, B) = 4
removedwins = removedwins + 1
End If
grid(a, B) = grid(a, B) - 1
End If


End If

End Sub
Private Sub undomove()
If movenum < 1 Then Exit Sub
If Val(Label5.Caption) + Val(Label7.Caption) = boardsize ^ 2 Then Exit Sub
On Error GoTo err
Dim player As String, opponent As String
If (turn = "blue" And human1.Checked = True) Or (turn = "red" And human2.Checked = True) Then
player = "human"
Else
player = "computer"
End If

If (turn = "red" And human1.Checked = True) Or (turn = "blue" And human2.Checked = True) Then
opponent = "human"
Else
opponent = "computer"
End If

If player = "human" And opponent = "human" Then
unmakemove (allmoves(movenum, 0))
If allmoves(movenum, 2) = True Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
If removedwins = 2 Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
End If
If allmoves(movenum, 1) = "blue" Then
Label5.Caption = Val(Label5.Caption) - removedwins
Else
Label7.Caption = Val(Label7.Caption) - removedwins
End If
End If
If turn <> allmoves(movenum, 1) Then
changeturn
End If
movenum = movenum - 1
lastblue = 0
lastred = 0
For i = movenum To 0 Step -1
If allmoves(i, 1) = "blue" And lastblue = 0 Then lastblue = allmoves(i, 0)
If allmoves(i, 1) = "red" And lastred = 0 Then lastred = allmoves(i, 0)
Next i
If lastblue = 0 Then firstbluemove = True
If lastred = 0 Then firstredmove = True

End If
If player = "human" And opponent = "computer" Then
changes = 0
saveturn = turn
Do
If allmoves(movenum, 1) <> saveturn Then
changes = changes + 1
saveturn = allmoves(movenum, 1)
End If
unmakemove (allmoves(movenum, 0))
If allmoves(movenum, 2) = True Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
If removedwins = 2 Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
End If

If allmoves(movenum, 1) = "blue" Then
Label5.Caption = Val(Label5.Caption) - removedwins
Else
Label7.Caption = Val(Label7.Caption) - removedwins
End If
End If
movenum = movenum - 1
Loop While changes < 2
lastblue = 0
lastred = 0
For i = movenum To 0 Step -1
If allmoves(i, 1) = "blue" And lastblue = 0 Then lastblue = allmoves(i, 0)
If allmoves(i, 1) = "red" And lastred = 0 Then lastred = allmoves(i, 0)
Next i
If lastblue = 0 Then firstbluemove = True
If lastred = 0 Then firstredmove = True

End If
err:
Exit Sub
End Sub
Private Sub takeline(linenum As Integer)
Dim indexx As Integer
If linenum <= horlinenum Then
indexx = linenum
If Label2(indexx).Enabled = False Then Exit Sub
recentwin = False
firstmove = False
Select Case turn
Case "blue"
testing = True
Call changecolor(Label2(indexx), &HFF8080)
If firstbluemove = False Then ordinarylast ("blue")
lastblue = indexx
firstbluemove = False

Case "red"
Call changecolor(Label2(indexx), &H8080FF)

If firstredmove = False Then ordinarylast ("red")
lastred = indexx
firstredmove = False
End Select
Label2(indexx).Enabled = False
a = Int(Label2(indexx).index / boardsize) + 1
B = (Label2(indexx).index Mod boardsize) + 1
If a > 1 Then
grid(a - 1, B) = grid(a - 1, B) + 1
End If
If a < (boardsize + 1) Then
grid(a, B) = grid(a, B) + 1
End If



Else


indexx = linenum - (horlinenum + 1)
recentwin = False
firstmove = False
If Label3(indexx).Enabled = False Then Exit Sub

Select Case turn
Case "blue"
testing = True
Call changecolor(Label3(indexx), &HFF8080)
If firstbluemove = False Then ordinarylast ("blue")
lastblue = indexx + (horlinenum + 1)
firstbluemove = False
Case "red"
Call changecolor(Label3(indexx), &H8080FF)
If firstredmove = False Then ordinarylast ("red")
lastred = indexx + (horlinenum + 1)
firstredmove = False
End Select
Label3(indexx).Enabled = False
a = Int(Label3(indexx).index / (boardsize + 1)) + 1
B = (Label3(indexx).index Mod (boardsize + 1)) + 1
If B > 1 Then
grid(a, B - 1) = grid(a, B - 1) + 1
End If
If B < (boardsize + 1) Then
grid(a, B) = grid(a, B) + 1
End If


End If
checkwin
movenum = movenum + 1
allmoves(movenum, 0) = linenum
allmoves(movenum, 1) = turn
allmoves(movenum, 2) = recentwin
If recentwin = False Then
changeturn
End If

End Sub
Private Sub resizeboard()

For i = horlinenum To 1 Step -1
Label2(i).Visible = False
Next i

For i = verlinenum To 1 Step -1
Label3(i).Visible = False
Next i

horlinenum = boardsize * (boardsize + 1) - 1
verlinenum = (boardsize + 1) * boardsize - 1

horlinewidth = (6825 - (105 * (boardsize + 1))) / boardsize
verlineheight = (5145 - (105 * (boardsize + 1))) / boardsize
For i = 0 To horlinenum
If i > loaded1 Then
Load Label2(i)
End If
Label2(i).Visible = True
rownum = Int(i / boardsize)
colnum = i Mod boardsize

Label2(i).Top = 960 + (rownum * (verlineheight + 105))
Label2(i).Height = 105
Label2(i).Width = horlinewidth
Label2(i).Left = 360 + (colnum * horlinewidth) + (105 * (colnum + 1))
Next i

For i = 0 To verlinenum
If i > loaded2 Then
Load Label3(i)
End If
Label3(i).Visible = True
rownum = Int(i / (boardsize + 1))
colnum = i Mod (boardsize + 1)

Label3(i).Top = 960 + (rownum * (verlineheight + 105)) + 105
Label3(i).Height = verlineheight
Label3(i).Width = 105
Label3(i).Left = 360 + (horlinewidth + 105) * colnum
Next i


If horlinenum > loaded1 Then
loaded1 = horlinenum
End If

If verlinenum > loaded2 Then
loaded2 = verlinenum
End If
End Sub


Private Function safelines() As Boolean
For i = 1 To boardsize - 1
For j = 1 To boardsize - 1
If grid(i, j) < 2 Then
If grid(i + 1, j) < 2 Or grid(i, j + 1) < 2 Then
safelines = True
Exit Function
End If
End If
Next
Next
safelines = False
End Function




Private Sub play(player As String, suggest As Boolean)
Dim diff As Integer
If suggest = False Then
Select Case player
Case "blue"
If very1.Checked = True Then diff = 1
If easy1.Checked = True Then diff = 2
If normal1.Checked = True Then diff = 3
If difficult1.Checked = True Then diff = 4
Case "red"
If very2.Checked = True Then diff = 1
If easy2.Checked = True Then diff = 2
If normal2.Checked = True Then diff = 3
If difficult2.Checked = True Then diff = 4
End Select
Else
Select Case turn
Case "red"
If very1.Checked = True Then diff = 1
If easy1.Checked = True Then diff = 2
If normal1.Checked = True Then diff = 3
If difficult1.Checked = True Then diff = 4
Case "blue"
If very2.Checked = True Then diff = 1
If easy2.Checked = True Then diff = 2
If normal2.Checked = True Then diff = 3
If difficult2.Checked = True Then diff = 4
End Select

End If


Dim plays(320) As Integer
Dim element As Integer
element = 0



k = "l"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim stopwin As Boolean
Dim takelinee As Boolean
Dim cells(145, 1)
Dim wins(0 To 20) As Integer
Dim lines As Integer

If safelines = False And diff = 4 Then


For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i


ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardsize Step 1
If found = True Then Exit For
For j = 1 To boardsize Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True Then 'And i > 1 Then
If i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
End If
cells(numberr, 0) = i
cells(numberr, 1) = j
If i > 1 Then
If grid2(i - 1, j) = 4 Then
numberr = numberr + 1
cells(numberr, 0) = i - 1
cells(numberr, 1) = j
End If
End If
End If
If Label2(B).Enabled = True Then ' And i < boardsize Then
If i < boardsize Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
End If
cells(numberr, 0) = i
cells(numberr, 1) = j
If i < boardsize Then
If grid2(i + 1, j) = 4 Then
numberr = numberr + 1
cells(numberr, 0) = i + 1
cells(numberr, 1) = j
End If
End If
End If
If Label3(c).Enabled = True Then ' And j > 1 Then
If j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
End If
cells(numberr, 0) = i
cells(numberr, 1) = j
If j > 1 Then
If grid2(i, j - 1) = 4 Then
numberr = numberr + 1
cells(numberr, 0) = i
cells(numberr, 1) = j - 1
End If
End If
End If
If Label3(d).Enabled = True Then ' And j < boardsize Then
If j < boardsize Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
End If
cells(numberr, 0) = i
cells(numberr, 1) = j
If j < boardsize Then
If grid2(i, j + 1) = 4 Then
numberr = numberr + 1
cells(numberr, 0) = i
cells(numberr, 1) = j + 1
End If
End If
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop


If numberr >= 2 Then
lines = 0
For i = 1 To boardsize
For j = 1 To boardsize
If grid(i, j) = 3 Then
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
If Label2(a).Enabled = True Then
lines = lines + 1
wins(lines) = a
End If
If Label2(B).Enabled = True Then
lines = lines + 1
wins(lines) = B
End If
If Label3(c).Enabled = True Then
lines = lines + 1
wins(lines) = c + (horlinenum + 1)
End If
If Label3(d).Enabled = True Then
lines = lines + 1
wins(lines) = d + (horlinenum + 1)
End If
End If
Next j
Next i


stopwin = True
For i = 1 To lines

If stopwin = False Then Exit For

For y = 1 To boardsize
For l = 1 To boardsize
grid2(y, l) = grid(y, l)
Next l
Next y

If Val(wins(i)) < horlinenum Then
u = Int(wins(i) / boardsize) + 1
v = (wins(i) Mod boardsize) + 1
If u > 1 Then
grid2(u - 1, v) = grid2(u - 1, v) + 1
End If
If u < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
Else
u = Int((wins(i) - (horlinenum + 1)) / (boardsize + 1)) + 1
v = ((wins(i) - (horlinenum + 1)) Mod (boardsize + 1)) + 1
If v > 1 Then
' MsgBox (Str(u))
' MsgBox (Str(v))

 grid2(u, v - 1) = grid2(u, v - 1) + 1
End If
If v < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
End If
takelinee = False

For j = 1 To numberr

a = cells(j, 0)
B = cells(j, 1)
If grid2(a, B) < 3 Then

If a > 1 Then
If Label2((a - 1) * boardsize + B - 1).Enabled = True And grid2(a - 1, B) < 3 Then
takelinee = True
Exit For
End If
End If

If a < (boardsize) Then
If Label2((a) * boardsize + B - 1).Enabled = True And grid2(a + 1, B) < 3 Then
takelinee = True
Exit For
End If
End If

If B > 1 Then
If Label3((a - 1) * (boardsize + 1) + B - 1).Enabled = True And grid2(a, B - 1) < 3 Then
takelinee = True
Exit For
End If
End If

If B < (boardsize) Then
If Label3((a - 1) * (boardsize + 1) + B).Enabled = True And grid2(a, B + 1) < 3 Then
takelinee = True
Exit For
End If
End If

End If

Next j

If takelinee = True Then
element = 1
plays(1) = wins(i)
k = "p"
stopwin = False
End If
If stopwin = False Then Exit For

Next i


For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i
If stopwin = True Then

For j = 1 To numberr
a = cells(j, 0)
B = cells(j, 1)
If grid2(a, B) < 3 Then

If a > 1 Then
If Label2((a - 1) * boardsize + B - 1).Enabled = True And grid2(a - 1, B) < 3 Then
plays(1) = (a - 1) * boardsize + B - 1
element = 1
'Exit For
End If
End If

If a < (boardsize) Then
If Label2((a) * boardsize + B - 1).Enabled = True And grid2(a + 1, B) < 3 Then
plays(1) = (a) * boardsize + B - 1
element = 1
'Exit For
End If
End If

If B > 1 Then
If Label3((a - 1) * (boardsize + 1) + B - 1).Enabled = True And grid2(a, B - 1) < 3 Then
plays(1) = (a - 1) * (boardsize + 1) + B - 1 + (horlinenum + 1)
element = 1
'Exit For
End If
End If

If B < (boardsize) Then
If Label3((a - 1) * (boardsize + 1) + B).Enabled = True And grid2(a, B + 1) < 3 Then
plays(1) = (a - 1) * (boardsize + 1) + B + (horlinenum + 1)
element = 1
'Exit For
End If
End If

End If
Next j


End If



End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Form1.Caption = Str(plays(1))
Randomize
yy = Int(Rnd * 10) + 1
If (diff > 1 Or yy > 5) And element = 0 Then
k = "d"
For i = 1 To boardsize
For j = 1 To boardsize
If grid(i, j) = 3 Then
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
If Label2(a).Enabled = True Then
element = element + 1
plays(element) = a
End If
If Label2(B).Enabled = True Then
element = element + 1
plays(element) = B
End If
If Label3(c).Enabled = True Then
element = element + 1
plays(element) = c + (horlinenum + 1)
End If
If Label3(d).Enabled = True Then
element = element + 1
plays(element) = d + (horlinenum + 1)
End If
End If
Next j
Next i
End If
If diff > 1 Then
If element = 0 Then
For i = 0 To horlinenum
If Label2(i).Enabled = True Then
a1 = 0
a2 = 0
a = Int(i / boardsize) + 1
B = (i Mod boardsize) + 1
If a > 1 Then
If grid(a - 1, B) < 2 Then a1 = 1
Else
a1 = 1
End If
If a < (boardsize + 1) Then
If grid(a, B) < 2 Then a2 = 1
Else
a2 = 1
End If
If a1 = 1 And a2 = 1 And Label2(i).Enabled = True Then
element = element + 1
plays(element) = i
End If
End If
Next i




For i = 0 To verlinenum
If Label3(i).Enabled = True Then

a1 = 0
a2 = 0
a = Int(i / (boardsize + 1)) + 1
B = (i Mod (boardsize + 1)) + 1
If B > 1 Then
If grid(a, B - 1) < 2 Then a1 = 1
Else
a1 = 1
End If
If B < (boardsize + 1) Then
If grid(a, B) < 2 Then a2 = 1
Else
a2 = 1
End If
If a1 = 1 And a2 = 1 Then
If Label3(i).Enabled = True Then
element = element + 1
plays(element) = i + (horlinenum + 1)
End If
End If
End If
Next i

End If
End If




If diff > 2 And element = 0 Then

leastnum = 500

For z = 0 To horlinenum

For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i


If Label2(z).Enabled = True Then
Label2(z).Enabled = False
u = Int(z / boardsize) + 1
v = (z Mod boardsize) + 1
If u > 1 Then
grid2(u - 1, v) = grid2(u - 1, v) + 1
End If
If u < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardsize Step 1
If found = True Then Exit For
For j = 1 To boardsize Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True And i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
End If
If Label2(B).Enabled = True And i < boardsize Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
End If
If Label3(c).Enabled = True And j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
End If
If Label3(d).Enabled = True And j < boardsize Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop



If numberr < leastnum Then
leastnum = numberr
chosen = z
End If
Label2(z).Enabled = True

End If
Next z







For z = 0 To verlinenum

For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i


If Label3(z).Enabled = True Then
Label3(z).Enabled = False
u = Int(z / (boardsize + 1)) + 1
v = (z Mod (boardsize + 1)) + 1
If v > 1 Then
grid2(u, v - 1) = grid2(u, v - 1) + 1
End If
If v < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardsize Step 1
If found = True Then Exit For
For j = 1 To boardsize Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True And i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
End If
If Label2(B).Enabled = True And i < boardsize Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
End If
If Label3(c).Enabled = True And j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
End If
If Label3(d).Enabled = True And j < boardsize Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop



If numberr < leastnum Then
leastnum = numberr
chosen = z + (horlinenum + 1)
End If
Label3(z).Enabled = True

End If
Next z

element = 1
plays(1) = chosen
End If





If element = 0 Then
For i = 0 To horlinenum
If Label2(i).Enabled = True Then
element = element + 1
plays(element) = i
End If
Next i
For i = 0 To verlinenum
If Label3(i).Enabled = True Then
element = element + 1
plays(element) = i + (horlinenum + 1)
End If
Next i
End If

If element <> 0 Then
d = Int(Rnd * element) + 1
e = plays(d)
End If
If suggest = False Then
takeline (e)
Else

If e < (horlinenum + 1) Then
Call changecolor(Label2(e), &H40&)
Delay (0.5)
Call changecolor(Label2(e), &HC000&)
Else
Call changecolor(Label3(e - (horlinenum + 1)), &H40&)
Delay (0.5)
Call changecolor(Label3(e - (horlinenum + 1)), &HC000&)

End If

End If
'Form1.Caption = k

End Sub
Private Sub checkwin()
For i = 1 To boardsize
    For j = 1 To boardsize
        If grid(i, j) = 4 Then
        recentwin = False
        grid(i, j) = 5
        Select Case turn
        Case "blue"
        Label5.Caption = Val(Label5.Caption) + 1
        Case "red"
        Label7.Caption = Val(Label7.Caption) + 1
        End Select
        recentwin = True
        winned = winned + 1
        Load Shape1(winned)
        Shape1(winned).Visible = True
        Select Case turn
        Case "blue"
        Shape1(winned).FillColor = &HFF0000
        Case "red"
        Shape1(winned).FillColor = &HFF&
        End Select
        Shape1(winned).Width = horlinewidth * 0.42
        Shape1(winned).Height = horlinewidth * 0.42
        centerx = 420 + (horlinewidth + 105) * (j - 0.5)
        centery = 1040 + (verlineheight + 105) * (i - 0.5)
        Shape1(winned).Top = centery - 0.5 * Shape1(winned).Width
        Shape1(winned).Left = centerx - 0.5 * Shape1(winned).Height
        If Val(Label5.Caption) + Val(Label7.Caption) = (boardsize * boardsize) Then
        endgame
        End If
        End If
    Next j
Next i

End Sub
Private Sub newgame()
On Error Resume Next
For i = 1 To winned
Shape1(i).Visible = False
Unload Shape1(i)
Next i

turn = "blue"
recentwin = False
winned = 0
For i = 1 To boardsize
    For j = 1 To boardsize
        grid(i, j) = 0
    Next j
Next i
For i = 0 To horlinenum
Label2(i).BackColor = &HC000&
Label2(i).Enabled = True
lbl2(i) = True
Next i
For i = 0 To verlinenum
Label3(i).BackColor = &HC000&
Label3(i).Enabled = True
lbl3(i) = True
Next i
Label5.Caption = "0"
Label7.Caption = "0"
Label14.Caption = "Blue"
Label14.ForeColor = &HFF0000
Timer1.Enabled = False
Timer1.Enabled = True
Timer2.Enabled = False
enabled1 = True
enabled2 = False
firstbluemove = True
firstredmove = True
movenum = 0
freezed = False
pauseresume.Caption = "Pause"
End Sub
Private Sub changeturn()
If turn = "blue" Then
turn = "red"
Label14.Caption = "Red"
Label14.ForeColor = &HFF&
Timer1.Enabled = False
Timer2.Enabled = True
Else
turn = "blue"
Label14.Caption = "Blue"
Label14.ForeColor = &HFF0000
Timer1.Enabled = True
Timer2.Enabled = False

End If
End Sub
Private Sub checkgamerunning()
If movenum > 1 And movenum < (horlinenum + verlinenum) Then
d = MsgBox("Can't change players during the game, do you want to start a new game?", vbYesNo)
If d = vbYes Then
newgame
isexit = False
Else
isexit = True
End If
End If
If freezed = True Then isexit = True
End Sub

Private Sub about_Click()
MsgBox ("Complete Square is a strategy game,programmed by:Sohaib Abu Farsakh")
End Sub

Private Sub animate_Click()
If animate.Checked = False Then
d = MsgBox("It is NOT recomended to enable this option, do you still want to enable it?", vbYesNo)
If d = vbNo Then Exit Sub
End If
animate.Checked = Not (animate.Checked)
End Sub

Private Sub chat_Click()
If chat.Checked = False Then
chat.Checked = True
Form2.Show , Form1
Else
chat.Checked = False
Form2.Hide
End If
End Sub

Private Sub Command1_Click()
Label16.Caption = "0"
Label18.Caption = "0"
End Sub



Private Sub Command2_Click()
Text1.Visible = False
Command2.Visible = False
End Sub

Private Sub contenents_Click()
Text1.Visible = True
Text1.Width = 8775
Text1.Height = 5350
Command2.Visible = True
End Sub

Private Sub difficult1_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label12.Top = 1420
human1.Checked = False
very1.Checked = False
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = True

End Sub

Private Sub difficult2_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label26.Top = 1420
human2.Checked = False
very2.Checked = False
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = True

End Sub

Private Sub easy1_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label12.Top = 950
human1.Checked = False
very1.Checked = False
easy1.Checked = True
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub easy2_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label26.Top = 950
human2.Checked = False
very2.Checked = False
easy2.Checked = True
normal2.Checked = False
difficult2.Checked = False

End Sub

Private Sub endnet_Click()
Winsock1.Close
internet = False
newgame
End Sub

Private Sub exit_Click()
Unload Form2
Unload Form1
End
End Sub

Private Sub fast_Click()
slow.Checked = False
normal.Checked = False
fast.Checked = True
Timer1.Interval = 70
Timer2.Interval = 70
End Sub

Private Sub Form_Load()
Load Form2
newgame
writehelp
horlinenum = 0
verlinenum = 0
loaded1 = 0
loaded2 = 0
colorsteps = 15
boardsize = 7
resizeboard
internet = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End
End Sub

Private Sub gridsize_Click()
On Error Resume Next

d = InputBox("please enter the board size (2 - 12)", "board size", Str(boardsize))
If d < 2 Then d = 2
If d > 12 Then d = 12
boardsize = d
newgame
resizeboard
End Sub

Private Sub host_Click()
Winsock1.LocalPort = 3584
Winsock1.Listen
Label27.Caption = "Your IP Address is " + Winsock1.LocalIP
internet = True
netstatus = "host"
newgame
End Sub

Private Sub human1_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label12.Top = 460
human1.Checked = True
very1.Checked = False
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub human2_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label26.Top = 460
human2.Checked = True
very2.Checked = False
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = False

End Sub

Private Sub join_Click()
j = InputBox("Enter your partner's IP here: ")
Winsock1.Connect j, 3584
internet = True
netstatus = "join"
newgame
End Sub

Private Sub Label1_Click()
very1_Click
End Sub

Private Sub Label10_Click()
difficult1_Click
End Sub

Private Sub Label11_Click()
human1_Click
End Sub

Private Sub Label19_Click()
human2_Click
End Sub

Private Sub Label2_Click(index As Integer)
Dim a As String
If Label2(index).Enabled = False Then Exit Sub
If internet = False Then
If (turn = "blue" And human1.Checked = False) Or (turn = "red" And human2.Checked = False) Then Exit Sub
Else
If (turn = "blue" And netstatus = "join") Or (turn = "red" And netstatus = "host") Then Exit Sub
End If
If freezed = True Then Exit Sub
takeline (index)
If internet = True Then
a = "P" + (Str(index))
Winsock1.SendData (a)
End If
End Sub

Private Sub Label20_Click()
very2_Click
End Sub

Private Sub Label21_Click()
easy2_Click
End Sub

Private Sub Label22_Click()
normal2_Click
End Sub

Private Sub Label23_Click()
difficult2_Click
End Sub

Private Sub Label3_Click(index As Integer)
Dim a As String
If Label3(index).Enabled = False Then Exit Sub
If internet = False Then
If (turn = "blue" And human1.Checked = False) Or (turn = "red" And human2.Checked = False) Then Exit Sub
Else
If (turn = "blue" And netstatus = "join") Or (turn = "red" And netstatus = "host") Then Exit Sub
End If
If freezed = True Then Exit Sub
takeline (index + (horlinenum + 1))
If internet = True Then
a = "P" + Str((index + (horlinenum + 1)))
Winsock1.SendData (a)
End If

End Sub

Private Sub Label8_Click()
easy1_Click
End Sub

Private Sub Label9_Click()
normal1_Click
End Sub

Private Sub new_Click()
newgame
End Sub


Private Sub normal_Click()
slow.Checked = False
normal.Checked = True
fast.Checked = False
Timer1.Interval = 200
Timer2.Interval = 200

End Sub

Private Sub normal1_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label12.Top = 1180
human1.Checked = False
very1.Checked = False
easy1.Checked = False
normal1.Checked = True
difficult1.Checked = False

End Sub

Private Sub normal2_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label26.Top = 1180
human2.Checked = False
very2.Checked = False
easy2.Checked = False
normal2.Checked = True
difficult2.Checked = False

End Sub

Private Sub pauseresume_Click()
Select Case pauseresume.Caption
Case "Pause"
pauseresume.Caption = "Resume"
saveenabled
enabled1 = Timer1.Enabled
enabled2 = Timer2.Enabled
Timer1.Enabled = False
Timer2.Enabled = False
Case "Resume"
pauseresume.Caption = "Pause"
revertenabled
Timer1.Enabled = enabled1
Timer2.Enabled = enabled2
End Select
End Sub

Private Sub slow_Click()
slow.Checked = True
normal.Checked = False
fast.Checked = False
Timer1.Interval = 500
Timer2.Interval = 500

End Sub

Private Sub suggest_Click()
suggest.Enabled = False
If human1.Checked = True And human2.Checked = True Then
d = MsgBox("Can't suggest move if two humans are playing!", vbInformation)
Exit Sub
End If
If (turn = "blue" And human1.Checked = True) Or (turn = "red" And human2.Checked = True) Then
Call play("blue", True)
End If
suggest.Enabled = True
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If internet = True Then Exit Sub
If human1.Checked = False And turn = "blue" Then
Call play("blue", False)
End If
End Sub



Private Sub Timer2_Timer()
If internet = True Then Exit Sub

If human2.Checked = False And turn = "red" Then
Call play("red", False)
End If

End Sub

Private Sub undo_Click()
undomove
End Sub

Private Sub very1_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label12.Top = 700
human1.Checked = False
very1.Checked = True
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub very2_Click()
isexit = False
checkgamerunning
If isexit = True Then Exit Sub

Label26.Top = 700
human2.Checked = False
very2.Checked = True
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = False

End Sub



Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
If internet = True Then
Winsock1.SendData "G" + (Str(boardsize))
End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Winsock1.GetData a, vbString
If Mid(a, 1, 1) = "P" Then
aa = Mid(a, 2, (Len(a) - 1))
takeline (Val(aa))
End If
If Mid(a, 1, 1) = "G" Then
aa = Mid(a, 2, (Len(a) - 1))
newgame
boardsize = Val(aa)
resizeboard
End If
If Mid(a, 1, 1) = "C" Then
aa = Mid(a, 2, (Len(a) - 1))
Form2.List1.AddItem aa
End If

End Sub

