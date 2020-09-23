VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link4 Game"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   6
      Left            =   4920
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   5
      Left            =   4320
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   4
      Left            =   3600
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   3
      Left            =   2880
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   2
      Left            =   2160
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   1
      Left            =   1440
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5175
      Index           =   0
      Left            =   600
      Tag             =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   48
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   47
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   46
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   45
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   44
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   43
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   42
      Left            =   720
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   41
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   40
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   39
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   38
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   37
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   36
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   35
      Left            =   720
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   34
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   33
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   32
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   31
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   30
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   29
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   28
      Left            =   720
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   27
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   26
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   25
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   24
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   23
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   22
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   21
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   20
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   19
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   18
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   17
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   16
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   15
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   720
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   600
      Top             =   720
      Width           =   5175
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
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
         End
         Begin VB.Menu doff 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu two 
         Caption         =   "Two Players"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim order As Integer
Option Base 1
Dim board(7, 7) As Integer
Dim diff As String

Private Sub play()

End Sub
Private Sub win()

End Sub

Private Sub doff_Click()
diff = 3
two.Checked = False
easy.Checked = False
normal.Checked = False
doff.Checked = True

End Sub

Private Sub easy_Click()
diff = 1
two.Checked = False
easy.Checked = True
normal.Checked = False
doff.Checked = False

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
order = 1
End Sub
Private Sub checkwin()
For b = 1 To 4
For a = 1 To 7
i = b
j = a
If board(i, j) < 3 And board(i, j) > 0 Then
If board(i, j) = board(i + 1, j) And board(i, j) = board(i + 2, j) And board(i, j) = board(i + 3, j) Then
If board(i, j) = 1 Then
d = MsgBox("yellow wins")
Else
d = MsgBox("red wins")
End If
For h = 0 To 6
Image1(h).Enabled = False
Next h
Exit Sub
End If
End If
Next a
Next b
For b = 1 To 7
For a = 1 To 4
i = b
j = a
If board(i, j) < 3 And board(i, j) > 0 Then
If board(i, j) = board(i, j + 1) And board(i, j) = board(i, j + 2) And board(i, j) = board(i, j + 3) Then
If board(i, j) = 1 Then
d = MsgBox("yellow wins")
Else
d = MsgBox("red wins")
End If
For h = 0 To 6
Image1(h).Enabled = False
Next h

Exit Sub
End If
End If
Next a
Next b
For b = 1 To 4
For a = 1 To 4
i = b
j = a
If board(i, j) < 3 And board(i, j) > 0 Then
If board(i, j) = board(i + 1, j + 1) And board(i, j) = board(i + 2, j + 2) And board(i, j) = board(i + 3, j + 3) Then
If board(i, j) = 1 Then
d = MsgBox("yellow wins")
Else
d = MsgBox("red wins")
End If
For h = 0 To 6
Image1(h).Enabled = False
Next h

Exit Sub
End If
End If
Next a
Next b
For b = 4 To 7
For a = 1 To 4
i = b
j = a
If board(i, j) < 3 And board(i, j) > 0 Then
If board(i, j) = board(i - 1, j + 1) And board(i, j) = board(i - 2, j + 2) And board(i, j) = board(i - 3, j + 3) Then
If board(i, j) = 1 Then
d = MsgBox("yellow wins")
Else
d = MsgBox("red wins")
End If
For h = 0 To 6
Image1(h).Enabled = False
Next h

Exit Sub
End If
End If
Next a
Next b

End Sub
Private Sub Image1_Click(Index As Integer)
On Error GoTo g
If (Val(Image1(Index).Tag) < 8) Then
a = Image1(Index).Tag * 7 + Image1(Index).Index
Select Case order
Case 1
Shape2(a).FillColor = vbYellow
board(Image1(Index).Index + 1, Image1(Index).Tag + 1) = order
If order = 1 Then
order = 2
Else
order = 1
End If
Case 2
Shape2(a).FillColor = vbRed
board(Image1(Index).Index + 1, Image1(Index).Tag + 1) = order

If order = 1 Then
order = 2
Else
order = 1
End If
End Select
Image1(Index).Tag = Image1(Index).Tag + 1
End If
checkwin
play
g:
Exit Sub
End Sub

Private Sub new_Click()
Unload Form1
Load Form1
Form1.Show
For i = 1 To 7
For j = 1 To 7
board(i, j) = 0
Next j
Next i


End Sub

Private Sub noraml_Click()
diff = 2
two.Checked = False
easy.Checked = False
normal.Checked = True
doff.Checked = False

End Sub

Private Sub two_Click()
diff = two
two.Checked = True
easy.Checked = False
normal.Checked = False
doff.Checked = False
End Sub
