VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Squares"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5595
   Begin VB.CommandButton Command20 
      Caption         =   "letters(english)"
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command19 
      Caption         =   "letters(arabic)"
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command18 
      Caption         =   "numbers"
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3240
   End
   Begin VB.CommandButton Command17 
      Caption         =   "mix"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Height          =   700
      Left            =   2400
      TabIndex        =   15
      Top             =   2400
      Width           =   700
   End
   Begin VB.CommandButton Command15 
      Caption         =   "15"
      Height          =   700
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Width           =   700
   End
   Begin VB.CommandButton Command14 
      Caption         =   "14"
      Height          =   700
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Width           =   700
   End
   Begin VB.CommandButton Command13 
      Caption         =   "13"
      Height          =   700
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   700
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      Height          =   700
      Left            =   2400
      TabIndex        =   11
      Top             =   1700
      Width           =   700
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      Height          =   700
      Left            =   1680
      TabIndex        =   10
      Top             =   1700
      Width           =   700
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      Height          =   700
      Left            =   960
      TabIndex        =   9
      Top             =   1700
      Width           =   700
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   700
      Left            =   240
      TabIndex        =   8
      Top             =   1700
      Width           =   700
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   700
      Left            =   2400
      TabIndex        =   7
      Top             =   1000
      Width           =   700
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   700
      Left            =   1680
      TabIndex        =   6
      Top             =   1000
      Width           =   700
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   700
      Left            =   960
      TabIndex        =   5
      Top             =   1000
      Width           =   700
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   700
      Left            =   240
      TabIndex        =   4
      Top             =   1000
      Width           =   700
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   700
      Left            =   2400
      TabIndex        =   3
      Top             =   300
      Width           =   700
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   700
      Left            =   1680
      TabIndex        =   2
      Top             =   300
      Width           =   700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   700
      Left            =   960
      TabIndex        =   1
      Top             =   300
      Width           =   700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   700
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   700
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Score:"
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   3600
      Picture         =   "ÊÑÊíÈ.frx":0000
      Top             =   3360
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Moves:"
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Time:"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Command1.Top >= (Command16.Top - 720) And Command1.Top <= (Command16.Top + 720) And Command1.Left >= (Command16.Left - 720) And Command1.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command1.Top
Command16.Left = Command1.Left
Command1.Top = a
Command1.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command10_Click()
If Command10.Top >= (Command16.Top - 720) And Command10.Top <= (Command16.Top + 720) And Command10.Left >= (Command16.Left - 720) And Command10.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command10.Top
Command16.Left = Command10.Left
Command10.Top = a
Command10.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command11_Click()
If Command11.Top >= (Command16.Top - 720) And Command11.Top <= (Command16.Top + 720) And Command11.Left >= (Command16.Left - 720) And Command11.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command11.Top
Command16.Left = Command11.Left
Command11.Top = a
Command11.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command12_Click()
If Command12.Top >= (Command16.Top - 720) And Command12.Top <= (Command16.Top + 720) And Command12.Left >= (Command16.Left - 720) And Command12.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command12.Top
Command16.Left = Command12.Left
Command12.Top = a
Command12.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command13_Click()
If Command13.Top >= (Command16.Top - 720) And Command13.Top <= (Command16.Top + 720) And Command13.Left >= (Command16.Left - 720) And Command13.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command13.Top
Command16.Left = Command13.Left
Command13.Top = a
Command13.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command14_Click()
If Command14.Top >= (Command16.Top - 720) And Command14.Top <= (Command16.Top + 720) And Command14.Left >= (Command16.Left - 720) And Command14.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command14.Top
Command16.Left = Command14.Left
Command14.Top = a
Command14.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command15_Click()
If Command15.Top >= (Command16.Top - 720) And Command15.Top <= (Command16.Top + 720) And Command15.Left >= (Command16.Left - 720) And Command15.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command15.Top
Command16.Left = Command15.Left
Command15.Top = a
Command15.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command17_Click()
For x = 1 To 100 Step 1
Randomize
a = Int(Rnd * 15) + 1
Select Case a
Case 1
b = Command10.Top
c = Command10.Left
Command10.Top = Command1.Top
Command10.Left = Command1.Left
Command1.Top = b
Command1.Left = c
Case 2
b = Command11.Top
c = Command11.Left
Command11.Top = Command4.Top
Command11.Left = Command4.Left
Command4.Top = b
Command4.Left = c
Case 3
b = Command1.Top
c = Command1.Left
Command1.Top = Command5.Top
Command1.Left = Command5.Left
Command5.Top = b
Command5.Left = c
Case 4
b = Command14.Top
c = Command14.Left
Command14.Top = Command2.Top
Command14.Left = Command2.Left
Command2.Top = b
Command2.Left = c
Case 5
b = Command16.Top
c = Command16.Left
Command16.Top = Command6.Top
Command16.Left = Command6.Left
Command6.Top = b
Command6.Left = c
Case 6
b = Command8.Top
c = Command8.Left
Command8.Top = Command15.Top
Command8.Left = Command15.Left
Command15.Top = b
Command15.Left = c
Case 7
b = Command7.Top
c = Command7.Left
Command7.Top = Command13.Top
Command7.Left = Command13.Left
Command13.Top = b
Command13.Left = c
Case 8
b = Command3.Top
c = Command3.Left
Command3.Top = Command7.Top
Command3.Left = Command7.Left
Command7.Top = b
Command7.Left = c
Case 9
b = Command2.Top
c = Command2.Left
Command2.Top = Command7.Top
Command2.Left = Command7.Left
Command7.Top = b
Command7.Left = c
Case 10
b = Command3.Top
c = Command3.Left
Command3.Top = Command9.Top
Command3.Left = Command9.Left
Command9.Top = b
Command9.Left = c
Case 11
b = Command15.Top
c = Command15.Left
Command15.Top = Command11.Top
Command15.Left = Command11.Left
Command11.Top = b
Command11.Left = c
Case 12
b = Command1.Top
c = Command1.Left
Command1.Top = Command9.Top
Command1.Left = Command9.Left
Command9.Top = b
Command9.Left = c
Case 13
b = Command8.Top
c = Command8.Left
Command8.Top = Command3.Top
Command8.Left = Command3.Left
Command3.Top = b
Command3.Left = c
Case 14
b = Command6.Top
c = Command6.Left
Command6.Top = Command15.Top
Command6.Left = Command15.Left
Command15.Top = b
Command15.Left = c
Case 15
b = Command5.Top
c = Command5.Left
Command5.Top = Command7.Top
Command5.Left = Command7.Left
Command7.Top = b
Command7.Left = c
Case 15
b = Command16.Top
c = Command16.Left
Command16.Top = Command10.Top
Command16.Left = Command10.Left
Command10.Top = b
Command10.Left = c
End Select
Next x
Timer1.Interval = 1000
Label2.Caption = "0"
Label4.Caption = "0"
Image1.Visible = False
End Sub

Private Sub Command18_Click()
Command1.Caption = "1"
Command2.Caption = "2"
Command3.Caption = "3"
Command4.Caption = "4"
Command5.Caption = "5"
Command6.Caption = "6"
Command7.Caption = "7"
Command8.Caption = "8"
Command9.Caption = "9"
Command10.Caption = "10"
Command11.Caption = "11"
Command12.Caption = "12"
Command13.Caption = "13"
Command14.Caption = "14"
Command15.Caption = "15"
End Sub

Private Sub Command19_Click()
Command1.Caption = "Ç"
Command2.Caption = "È"
Command3.Caption = "Ê"
Command4.Caption = "Ë"
Command5.Caption = "Ì"
Command6.Caption = "Í"
Command7.Caption = "Î"
Command8.Caption = "Ï"
Command9.Caption = "Ð"
Command10.Caption = "Ñ"
Command11.Caption = "Ò"
Command12.Caption = "Ó"
Command13.Caption = "Ô"
Command14.Caption = "Õ"
Command15.Caption = "Ö"
End Sub

Private Sub Command2_Click()
If Command2.Top >= (Command16.Top - 720) And Command2.Top <= (Command16.Top + 720) And Command2.Left >= (Command16.Left - 720) And Command2.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command2.Top
Command16.Left = Command2.Left
Command2.Top = a
Command2.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command20_Click()
Command1.Caption = "A"
Command2.Caption = "B"
Command3.Caption = "C"
Command4.Caption = "D"
Command5.Caption = "E"
Command6.Caption = "F"
Command7.Caption = "G"
Command8.Caption = "H"
Command9.Caption = "I"
Command10.Caption = "J"
Command11.Caption = "K"
Command12.Caption = "L"
Command13.Caption = "M"
Command14.Caption = "N"
Command15.Caption = "O"
End Sub

Private Sub Command3_Click()
If Command3.Top >= (Command16.Top - 720) And Command3.Top <= (Command16.Top + 720) And Command3.Left >= (Command16.Left - 720) And Command3.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command3.Top
Command16.Left = Command3.Left
Command3.Top = a
Command3.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command4_Click()
If Command4.Top >= (Command16.Top - 720) And Command4.Top <= (Command16.Top + 720) And Command4.Left >= (Command16.Left - 720) And Command4.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command4.Top
Command16.Left = Command4.Left
Command4.Top = a
Command4.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command5_Click()
If Command5.Top >= (Command16.Top - 720) And Command5.Top <= (Command16.Top + 720) And Command5.Left >= (Command16.Left - 720) And Command5.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command5.Top
Command16.Left = Command5.Left
Command5.Top = a
Command5.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command6_Click()
If Command6.Top >= (Command16.Top - 720) And Command6.Top <= (Command16.Top + 720) And Command6.Left >= (Command16.Left - 720) And Command6.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command6.Top
Command16.Left = Command6.Left
Command6.Top = a
Command6.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command7_Click()
If Command7.Top >= (Command16.Top - 720) And Command7.Top <= (Command16.Top + 720) And Command7.Left >= (Command16.Left - 720) And Command7.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command7.Top
Command16.Left = Command7.Left
Command7.Top = a
Command7.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command8_Click()
If Command8.Top >= (Command16.Top - 720) And Command8.Top <= (Command16.Top + 720) And Command8.Left >= (Command16.Left - 720) And Command8.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command8.Top
Command16.Left = Command8.Left
Command8.Top = a
Command8.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Command9_Click()
If Command9.Top >= (Command16.Top - 720) And Command9.Top <= (Command16.Top + 720) And Command9.Left >= (Command16.Left - 720) And Command9.Left <= (Command16.Left + 720) Then
a = Command16.Top
b = Command16.Left
Command16.Top = Command9.Top
Command16.Left = Command9.Left
Command9.Top = a
Command9.Left = b
Label4.Caption = Val(Label4.Caption) + 1
End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Val(Label2.Caption) + 1
If Command1.Top = 300 And Command2.Top = 300 And Command3.Top = 300 And Command4.Top = 300 And Command5.Top = 1000 And Command6.Top = 1000 And Command7.Top = 1000 And Command8.Top = 1000 And Command9.Top = 1700 And Command10.Top = 1700 And Command11.Top = 1700 And Command12.Top = 1700 And Command13.Top = 2400 And Command14.Top = 2400 And Command15.Top = 2400 And Command1.Left = 240 And Command5.Left = 240 And Command9.Left = 240 And Command13.Left = 240 And Command2.Left = 960 And Command6.Left = 960 And Command10.Left = 960 And Command14.Left = 960 And Command3.Left = 1680 And Command7.Left = 1680 And Command11.Left = 1680 And Command15.Left = 1680 And Command4.Left = 2400 And Command8.Left = 2400 And Command12.Left = 2400 Then
Timer1.Interval = 0
Image1.Visible = True
a = 1000 - (Val(Label2.Caption) * 2) - (Val(Label4.Caption) * 5)
Label6.Caption = Val(Label6.Caption) + a
End If
End Sub
