VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matching"
   ClientHeight    =   5475
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Tag             =   "0"
      Top             =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1215
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00C00000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1215
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu gametype 
         Caption         =   "Game Type"
         Begin VB.Menu color 
            Caption         =   "Match Color"
         End
         Begin VB.Menu value 
            Caption         =   "Match Value"
         End
         Begin VB.Menu both 
            Caption         =   "Match Color And Value"
         End
      End
      Begin VB.Menu num 
         Caption         =   "Number  Of Cards"
         Begin VB.Menu numnum 
            Caption         =   "8"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unfolded As Integer
Dim cardnum As Integer
Dim gtype As Integer
Dim matched As Boolean
Dim indexx As Integer

Private Sub both_Click()
Unload Form1
Load Form1
Form1.Show
gtype = 3

End Sub

Private Sub color_Click()
Unload Form1
Load Form1
Form1.Show
gtype = 1

End Sub

Private Sub Form_Load()
For k = 1 To 7
Load numnum(k)
numnum(k).Caption = (k + 1) * 8
Next k
gtype = 1
unfolded = 0
a = 120
b = 220
For i = 1 To 63
Load Shape1(i)
Shape1(i).Visible = True
Load Image1(i)
Image1(i).Visible = True
Load Label1(i)

a = a + 1080
Shape1(i).Left = a
Shape1(i).Top = b
Image1(i).Left = a
Image1(i).Top = b
Label1(i).Left = a
Label1(i).Top = b

Label1(i).Caption = (i) \ 8
l = (i + 1) Mod 8
Select Case l
Case 1, 2
Label1(i).BackColor = vbRed
Case 3, 4
Label1(i).BackColor = vbBlue
Case 5, 6
Label1(i).BackColor = vbbrown
Case 7, 0
Label1(i).BackColor = vbGreen
End Select
If a > 10500 Then
a = 120 - 1080
b = b + 1400
End If
Next i
For j = 1 To 300
Randomize
d = Int(Rnd * 63)
e = Int(Rnd * 63)
x = Label1(d).BackColor
y = Label1(d).Caption
Label1(d).Caption = Label1(e).Caption
Label1(d).BackColor = Label1(e).BackColor
Label1(e).Caption = y
Label1(e).BackColor = x
Next j
End Sub

Private Sub Image1_Click(Index As Integer)
Image1(Index).Enabled = False
If unfolded = 1 Then
Select Case gtype
Case 1
If Label1(cardnum).BackColor = Label1(Index).BackColor Then
matched = True
Else
matched = False
End If
Case 2
If Label1(cardnum).Caption = Label1(Index).Caption Then
matched = True
Else
matched = False
End If
Case 3
If Label1(cardnum).Caption = Label1(Index).Caption And Label1(cardnum).BackColor = Label1(Index).BackColor Then
matched = True
Else
matched = False
End If
End Select
If matched = True Then
Image1(Index).Visible = False
Image1(cardnum).Visible = False
Label1(Index).Visible = True
Label1(cardnum).Visible = True
Shape1(Index).Visible = False
Shape1(cardnum).Visible = False
indexx = Index
Timer1.Enabled = True
unfolded = 0
Else
Label1(Index).Visible = True
Shape1(Index).Visible = False
Image1(Index).Enabled = True
Image1(cardnum).Enabled = True
indexx = Index
Timer1.Enabled = True
unfolded = 0
End If
Exit Sub
End If
If unfolded = 0 Then
Shape1(Index).Visible = False
Label1(Index).Visible = True
unfolded = 1
cardnum = Image1(Index).Index
End If

End Sub

Private Sub Timer1_Timer()
Select Case Timer1.Tag
Case 0
Timer1.Tag = 1
Case 1
Label1(indexx).Visible = False
Label1(cardnum).Visible = False
If matched = False Then
Shape1(indexx).Visible = True
Shape1(cardnum).Visible = True
End If
Timer1.Tag = 0

Timer1.Enabled = False
End Select
End Sub

Private Sub value_Click()
Unload Form1
Load Form1
Form1.Show
gtype = 2

End Sub
