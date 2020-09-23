VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect4"
   ClientHeight    =   5535
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   4335
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "42"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining cells:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Normal"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Game mode:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Wins by red"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Wins by yellow"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Next turn:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   2445
      Left            =   600
      Picture         =   "Connect4.frx":0000
      ScaleHeight     =   2385
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4395
      Left            =   240
      Picture         =   "Connect4.frx":523E
      ScaleHeight     =   4335
      ScaleWidth      =   7020
      TabIndex        =   0
      Top             =   600
      Width           =   7080
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   6
         Left            =   5280
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   5
         Left            =   4560
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   4
         Left            =   3840
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   3
         Left            =   3240
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   2
         Left            =   2520
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   1
         Left            =   1800
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   3615
         Index           =   0
         Left            =   1080
         Top             =   120
         Width           =   615
      End
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
         Begin VB.Menu diff 
            Caption         =   "Beginner"
            Index           =   0
         End
         Begin VB.Menu diff 
            Caption         =   "Easy"
            Index           =   1
         End
         Begin VB.Menu diff 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu diff 
            Caption         =   "Difficult"
            Index           =   3
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
Dim rowcells(1 To 7) As Byte
Dim turn As String
Dim grid(1 To 7, 1 To 6) As Byte, grid2(1 To 7, 1 To 6) As Byte
Dim rowvalue(1 To 7) As Long
Dim startdepth As Byte, currow As Byte
Dim searchplayer As Byte
Dim remainingcells As Byte
Dim difficulty As Byte
Dim running As Boolean
Private Sub newgame()
Picture1.Cls
For i = 1 To 7
rowcells(i) = 0
Next
turn = "yellow"
Label2.Caption = "Yellow"
Label2.ForeColor = vbYellow
remainingcells = 42
Label10.Caption = "42"
running = True
For i = 0 To 6
Image1(i).Enabled = True
Next
For i = 1 To 7
For j = 1 To 6
grid(i, j) = 0
Next
Next
End Sub
Private Sub changeturn()
Select Case turn
Case "yellow"
turn = "red"
Label2.Caption = "Red"
Label2.ForeColor = vbRed
Case "red"
turn = "yellow"
Label2.Caption = "Yellow"
Label2.ForeColor = vbYellow
End Select
End Sub
Private Sub takemove(row As Byte)
If rowcells(row) >= 6 Then Exit Sub
Call drawcell(row, rowcells(row) + 1, turn)
rowcells(row) = rowcells(row) + 1
Select Case turn
Case "yellow"
grid(row, rowcells(row)) = 1
Case "red"
grid(row, rowcells(row)) = 2
End Select
remainingcells = remainingcells - 1
Label10.Caption = Str(remainingcells)
checkwin
changeturn
End Sub
Private Sub checkwin()
Dim winning As Boolean
winning = False
For i = 1 To 4
If winning = True Then Exit For
For j = 1 To 6
If winning = True Then Exit For
If grid(i, j) = grid(i + 1, j) And grid(i, j) = grid(i + 2, j) And grid(i, j) = grid(i + 3, j) And grid(i, j) <> 0 Then
winning = True
End If
Next
Next
For i = 1 To 7
If winning = True Then Exit For
For j = 1 To 3
If winning = True Then Exit For
If grid(i, j) = grid(i, j + 1) And grid(i, j) = grid(i, j + 2) And grid(i, j) = grid(i, j + 3) And grid(i, j) <> 0 Then
winning = True
End If
Next
Next
For i = 1 To 4
If winning = True Then Exit For
For j = 1 To 3
If winning = True Then Exit For
If grid(i, j) = grid(i + 1, j + 1) And grid(i, j) = grid(i + 2, j + 2) And grid(i, j) = grid(i + 3, j + 3) And grid(i, j) <> 0 Then
winning = True
End If
Next
Next
For i = 4 To 7
If winning = True Then Exit For
For j = 1 To 3
If winning = True Then Exit For
If grid(i, j) = grid(i - 1, j + 1) And grid(i, j) = grid(i - 2, j + 2) And grid(i, j) = grid(i - 3, j + 3) And grid(i, j) <> 0 Then
winning = True
End If
Next
Next
If winning = True Then
running = False
For i = 0 To 6
Image1(i).Enabled = False
Next
s = "The " + turn + " player wins the game"
MsgBox (s)
If turn = "yellow" Then
Label4.Caption = Val(Label4.Caption) + 1
Else
Label6.Caption = Val(Label6.Caption) + 1
End If
End If
End Sub
Private Sub drawcell(x As Byte, y As Byte, color As String)
Select Case color
Case "white"
Y2 = 1800
Case "red"
Y2 = 1200
Case "yellow"
Y2 = 600
End Select
Picture1.PaintPicture Picture2.Picture, (x - 1) * 690 + 1133, (6 - y) * 600 + 180, 600, 600, 0, Y2, 600, 600, vbSrcCopy
End Sub
Private Sub equalgrids()
For i = 1 To 7
For j = 1 To 6
grid2(i, j) = grid(i, j)
Next
Next
End Sub
Private Sub lineatt()
For i = 1 To 4
For j = 1 To 6
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i + 1, j) = 0 Then emptyy = emptyy + 1
If grid(i + 1, j) = 1 Then yellow = yellow + 1
If grid(i + 1, j) = 2 Then red = red + 1
If grid(i + 2, j) = 0 Then emptyy = emptyy + 1
If grid(i + 2, j) = 1 Then yellow = yellow + 1
If grid(i + 2, j) = 2 Then red = red + 1
If grid(i + 3, j) = 0 Then emptyy = emptyy + 1
If grid(i + 3, j) = 1 Then yellow = yellow + 1
If grid(i + 3, j) = 2 Then red = red + 1

If red = 2 And yellow = 0 Then
For k = i To i + 3
If grid(k, j) = 0 And rowcells(k) = j - 1 Then
rowvalue(k) = rowvalue(k) + 90
End If
Next k
End If

If red = 1 And yellow = 0 And difficulty > 2 Then
For k = i To i + 3
If grid(k, j) = 0 And rowcells(k) = j - 1 Then
rowvalue(k) = rowvalue(k) + 2
End If
Next k
End If

Next j
Next i


For i = 1 To 4
For j = 1 To 3
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i + 1, j + 1) = 0 Then emptyy = emptyy + 1
If grid(i + 1, j + 1) = 1 Then yellow = yellow + 1
If grid(i + 1, j + 1) = 2 Then red = red + 1
If grid(i + 2, j + 2) = 0 Then emptyy = emptyy + 1
If grid(i + 2, j + 2) = 1 Then yellow = yellow + 1
If grid(i + 2, j + 2) = 2 Then red = red + 1
If grid(i + 3, j + 3) = 0 Then emptyy = emptyy + 1
If grid(i + 3, j + 3) = 1 Then yellow = yellow + 1
If grid(i + 3, j + 3) = 2 Then red = red + 1

If red = 2 And yellow = 0 Then
For k = 0 To 3
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 1 Then
rowvalue(i + k) = rowvalue(i + k) + 90
End If
Next k
End If

If red = 1 And yellow = 0 And difficulty > 2 Then
For k = 0 To 3
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 1 Then
rowvalue(i + k) = rowvalue(i + k) + 2
End If
Next k
End If

Next j
Next i



For i = 4 To 7
For j = 1 To 3
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i - 1, j + 1) = 0 Then emptyy = emptyy + 1
If grid(i - 1, j + 1) = 1 Then yellow = yellow + 1
If grid(i - 1, j + 1) = 2 Then red = red + 1
If grid(i - 2, j + 2) = 0 Then emptyy = emptyy + 1
If grid(i - 2, j + 2) = 1 Then yellow = yellow + 1
If grid(i - 2, j + 2) = 2 Then red = red + 1
If grid(i - 3, j + 3) = 0 Then emptyy = emptyy + 1
If grid(i - 3, j + 3) = 1 Then yellow = yellow + 1
If grid(i - 3, j + 3) = 2 Then red = red + 1

If red = 2 And yellow = 0 Then
For k = 0 To 3
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 1 Then
rowvalue(i - k) = rowvalue(i - k) + 90
End If
Next k
End If

If red = 1 And yellow = 0 And difficulty > 2 Then
For k = 0 To 3
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 1 Then
rowvalue(i - k) = rowvalue(i - k) + 2
End If
Next k
End If

Next j
Next i

End Sub
Private Sub linedef()
Dim i As Byte


For i = 1 To 7
num = rowcells(i)
If num < 6 And num >= 2 Then
If grid(i, num) = 1 And grid(i, num - 1) = 1 Then
rowvalue(i) = rowvalue(i) + 140
End If
End If
Next i


For i = 1 To 4
For j = 1 To 6
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i + 1, j) = 0 Then emptyy = emptyy + 1
If grid(i + 1, j) = 1 Then yellow = yellow + 1
If grid(i + 1, j) = 2 Then red = red + 1
If grid(i + 2, j) = 0 Then emptyy = emptyy + 1
If grid(i + 2, j) = 1 Then yellow = yellow + 1
If grid(i + 2, j) = 2 Then red = red + 1
If grid(i + 3, j) = 0 Then emptyy = emptyy + 1
If grid(i + 3, j) = 1 Then yellow = yellow + 1
If grid(i + 3, j) = 2 Then red = red + 1

If red = 0 And yellow = 2 Then
For k = i To i + 3
If grid(k, j) = 0 And rowcells(k) = j - 1 Then
rowvalue(k) = rowvalue(k) + 180
If (rowcells(k) + 1) = 3 Or (rowcells(k) + 1) = 5 Then
rowvalue(k) = rowvalue(k) + 1000
End If
End If
If grid(k, j) = 0 And rowcells(k) = j - 2 And diffculty > 2 Then
rowvalue(k) = rowvalue(k) - 45
If (rowcells(k) + 2) = 3 Or (rowcells(k) + 2) = 5 Then
rowvalue(k) = rowvalue(k) - 250
End If

End If

Next k
End If


If red = 0 And yellow = 1 And difficulty > 2 Then
For k = i To i + 3
If grid(k, j) = 0 And rowcells(k) = j - 1 Then
rowvalue(k) = rowvalue(k) + 4
End If
If grid(k, j) = 0 And rowcells(k) = j - 2 And diffculty > 2 Then
rowvalue(k) = rowvalue(k) - 1
End If

Next k
End If

Next j
Next i


For i = 1 To 4
For j = 1 To 3
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i + 1, j + 1) = 0 Then emptyy = emptyy + 1
If grid(i + 1, j + 1) = 1 Then yellow = yellow + 1
If grid(i + 1, j + 1) = 2 Then red = red + 1
If grid(i + 2, j + 2) = 0 Then emptyy = emptyy + 1
If grid(i + 2, j + 2) = 1 Then yellow = yellow + 1
If grid(i + 2, j + 2) = 2 Then red = red + 1
If grid(i + 3, j + 3) = 0 Then emptyy = emptyy + 1
If grid(i + 3, j + 3) = 1 Then yellow = yellow + 1
If grid(i + 3, j + 3) = 2 Then red = red + 1

If red = 0 And yellow = 2 Then
For k = 0 To 3
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 1 Then
rowvalue(i + k) = rowvalue(i + k) + 180
End If
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 2 And diffculty > 2 Then
rowvalue(i + k) = rowvalue(i + k) - 45
End If

Next k
End If


If red = 0 And yellow = 1 And difficulty > 2 Then
For k = 0 To 3
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 1 Then
rowvalue(i + k) = rowvalue(i + k) + 4
End If
If grid(i + k, j + k) = 0 And rowcells(i + k) = j + k - 2 And diffculty > 2 Then
rowvalue(i + k) = rowvalue(i + k) - 1
End If

Next k
End If

Next j
Next i



For i = 4 To 7
For j = 1 To 3
yellow = 0
red = 0
emptyy = 0
If grid(i, j) = 0 Then emptyy = emptyy + 1
If grid(i, j) = 1 Then yellow = yellow + 1
If grid(i, j) = 2 Then red = red + 1
If grid(i - 1, j + 1) = 0 Then emptyy = emptyy + 1
If grid(i - 1, j + 1) = 1 Then yellow = yellow + 1
If grid(i - 1, j + 1) = 2 Then red = red + 1
If grid(i - 2, j + 2) = 0 Then emptyy = emptyy + 1
If grid(i - 2, j + 2) = 1 Then yellow = yellow + 1
If grid(i - 2, j + 2) = 2 Then red = red + 1
If grid(i - 3, j + 3) = 0 Then emptyy = emptyy + 1
If grid(i - 3, j + 3) = 1 Then yellow = yellow + 1
If grid(i - 3, j + 3) = 2 Then red = red + 1

If red = 0 And yellow = 2 Then
For k = 0 To 3
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 1 Then
rowvalue(i - k) = rowvalue(i - k) + 180
End If
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 2 And diffculty > 2 Then
rowvalue(i - k) = rowvalue(i - k) - 45
End If

Next k
End If


If red = 0 And yellow = 1 And difficulty > 2 Then
For k = 0 To 3
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 1 Then
rowvalue(i - k) = rowvalue(i - k) + 4
End If
If grid(i - k, j + k) = 0 And rowcells(i - k) = j + k - 2 And diffculty > 2 Then
rowvalue(i - k) = rowvalue(i - k) - 1
End If

Next k
End If

Next j
Next i
End Sub
Private Function causewin(x As Byte, y As Byte, player As Byte) As Boolean ' the player can be either1(yellow) or 2(red)
On Error Resume Next
Dim temp As Byte
causewin = False
If y >= 4 Then
If grid2(x, y - 1) = grid2(x, y - 2) And grid2(x, y - 1) = grid2(x, y - 3) And grid2(x, y - 1) = player Then
causewin = True
Exit Function
End If
End If
temp = grid2(x, y)

grid2(x, y) = player
For i = 1 To 4
If grid2(i, y) = player And grid2(i + 1, y) = player And grid2(i + 2, y) = player And grid2(i + 3, y) = player Then
causewin = True
grid2(x, y) = temp
Exit Function
End If
Next
For i = 1 To 4
For j = 1 To 3
If grid2(i, j) = grid2(i + 1, j + 1) And grid2(i, j) = grid2(i + 2, j + 2) And grid2(i, j) = grid2(i + 3, j + 3) And grid2(i, j) = player Then
causewin = True
grid2(x, y) = temp
Exit Function
End If
Next
Next

For i = 4 To 7
For j = 1 To 3
If grid2(i, j) = grid2(i - 1, j + 1) And grid2(i, j) = grid2(i - 2, j + 2) And grid2(i, j) = grid2(i - 3, j + 3) And grid2(i, j) = player Then
causewin = True
grid2(x, y) = temp
Exit Function
End If
Next
Next
grid2(x, y) = temp


End Function
Private Function rowwin(roww As Byte, player As Byte) As Boolean

rowwin = causewin(roww, rowcells(roww) + 1, player)
End Function
Private Function getplayer(depth As Byte) As Byte
If (startdepth - depth) Mod 2 = 0 Then
getplayer = 2
Else
getplayer = 1
End If

End Function
Private Sub search(depth As Byte)
On Error Resume Next
Dim counter As Byte, ii As Byte
For ii = 1 To 7
If depth = startdepth Then currow = ii
If rowcells(ii) < 6 Then
grid2(ii, rowcells(ii) + 1) = getplayer(depth)
rowcells(ii) = rowcells(ii) + 1
If getplayer(depth) = 1 Then

For counter = 1 To 7
'If rowwin(j, 2) = True Then
If causewin(counter, rowcells(counter) + 1, 2) = True Then
a = 60 / (5 ^ (1 - depth))
rowvalue(currow) = rowvalue(currow) + a
End If
Next
End If
If getplayer(depth) = 2 Then

For counter = 1 To 7
If causewin(counter, rowcells(counter) + 1, 1) = True Then
a = -60 / (5 ^ (1 - depth))
rowvalue(currow) = rowvalue(currow) + a
 'Form1.Caption = Str(counter)
End If
Next
End If
If depth > 0 Then
search (depth - 1)
End If
rowcells(ii) = rowcells(ii) - 1
grid2(ii, rowcells(ii) + 1) = 0
End If
'Form1.Caption = Str(depth)
DoEvents
Next
End Sub
Private Function selmove() As Byte
Dim topnum As Integer
topnum = -10000
selmove = 0
For i = 1 To 7
If rowvalue(i) = topnum Then
Randomize
a = Int(Rnd * 4) + 1
If a = 1 Then selmove = i
End If
If rowvalue(i) > topnum Then
selmove = i
topnum = rowvalue(i)
'Form1.Caption = Str(topnum) + Str(i)

End If
Next i
End Function
Private Sub play()
Dim i As Byte
equalgrids
For i = 1 To 7
rowvalue(i) = 0
If rowcells(i) <= 4 And difficulty > 0 Then
If causewin(i, rowcells(i) + 2, 1) = True Then rowvalue(i) = -5000
End If
If rowcells(i) <= 5 Then
If difficulty > 0 Then
If causewin(i, rowcells(i) + 1, 1) = True Then rowvalue(i) = 5000
End If
If causewin(i, rowcells(i) + 1, 2) = True Then rowvalue(i) = 10000
End If
If rowcells(i) >= 6 Then rowvalue(i) = -100000
Next
If difficulty > 1 Then
lineatt
linedef
rowvalue(2) = rowvalue(2) + 16
rowvalue(3) = rowvalue(3) + 10
rowvalue(4) = rowvalue(4) + 25
rowvalue(5) = rowvalue(5) + 10
rowvalue(6) = rowvalue(6) + 16
End If
If difficulty > 2 Then
For i = 1 To 7
If causewin(i, rowcells(i) + 2, 2) = True Then rowvalue(i) = rowvalue(i) - 50
If causewin(i, rowcells(i) + 2, 2) = True And causewin(i, rowcells(i) + 3, 2) = True Then rowvalue(i) = rowvalue(i) + 200
Next
startdepth = 1
search (1)
End If


takemove (selmove)

End Sub

Private Sub Command1_Click()
For i = 1 To 7
MsgBox (Str(rowvalue(i)))
Next
End Sub

Private Sub Command2_Click()
d = InputBox("")
Form1.Caption = rowcells(Val(d))
End Sub

Private Sub diff_Click(Index As Integer)
two.Checked = False
For i = 0 To 3
diff(i).Checked = False
Next
diff(Index).Checked = True
difficulty = Index
Label8.Caption = diff(Index).Caption
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
newgame
difficulty = 2
End Sub

Private Sub Image1_Click(Index As Integer)
If rowcells(Index + 1) >= 6 Then Exit Sub

Call takemove(Index + 1)
If two.Checked = False And remainingcells > 0 And running = True Then
play
End If
End Sub

Private Sub new_Click()
newgame
End Sub

Private Sub two_Click()
For i = 0 To 3
diff(i).Checked = False
Next
two.Checked = True
Label8.Caption = "Two players"
End Sub
