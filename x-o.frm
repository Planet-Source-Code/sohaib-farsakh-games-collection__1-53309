VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X-O game"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton Option5 
      Caption         =   "master"
      Height          =   495
      Left            =   3000
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "expert"
      Height          =   495
      Left            =   3000
      TabIndex        =   21
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "reset"
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Tag             =   "X"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "play again"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Tag             =   "X"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   1320
   End
   Begin VB.OptionButton Option3 
      Caption         =   "hard"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "medium"
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "easy"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "X                O"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "" And Label1.Caption = "X" Then
Command1.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command10_Click()
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
Select Case Label2.Caption
Case "X"
Label2.Caption = "O"
Case "O"
Label2.Caption = "X"
End Select
Label1.Caption = Label2.Caption
Label3.Caption = "0"
Label7.Caption = ""
End Sub

Private Sub Command11_Click()
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
Label2.Caption = "X"
Label1.Caption = Label2.Caption
Label3.Caption = "0"
Label7.Caption = ""
Label5.Caption = "0"
Label6.Caption = "0"
End Sub

Private Sub Command2_Click()
If Command2.Caption = "" And Label1.Caption = "X" Then
Command2.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "" And Label1.Caption = "X" Then
Command3.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "" And Label1.Caption = "X" Then
Command4.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command5_Click()
If Command5.Caption = "" And Label1.Caption = "X" Then
Command5.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command6_Click()
If Command6.Caption = "" And Label1.Caption = "X" Then
Command6.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
End If
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End Sub

Private Sub Command7_Click()
If Command7.Caption = "" And Label1.Caption = "X" Then
Command7.Caption = Label1.Caption
Label1.Caption = "O"
Label3.Caption = Val(Label3.Caption) + 1
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
End If
End Sub

Private Sub Command8_Click()
If Command8.Caption = "" And Label1.Caption = "X" Then
Command8.Caption = Label1.Caption
Label3.Caption = Val(Label3.Caption) + 1
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
Label1.Caption = "O"
End If
End Sub

Private Sub Command9_Click()
If Command9.Caption = "" And Label1.Caption = "X" Then
Command9.Caption = Label1.Caption
Label3.Caption = Val(Label3.Caption) + 1
If ((Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command5.Caption = "X" And Command2.Caption = "X" And Command8.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "X") Or (Command9.Caption = "X" And Command6.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X")) And Label7.Caption = "" Then
Label7.Caption = "X wins"
Label5.Caption = Val(Label5.Caption) + 1
End If
Label1.Caption = "O"
End If
End Sub

Private Sub Timer1_Timer()


If Option2.Value = True Or Option3.Value = True Or Option4.Value = True Or Option5.Value = True Then
If ((Command2.Caption = "O" And Command3.Caption = "O") Or (Command5.Caption = "O" And Command9.Caption = "O") Or (Command4.Caption = "O" And Command7.Caption = "O")) And Label1.Caption = "O" And Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "O" And Command1.Caption = "O") Or (Command5.Caption = "O" And Command7.Caption = "O") Or (Command6.Caption = "O" And Command9.Caption = "O")) And Label1.Caption = "O" And Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If ((Command3.Caption = "O" And Command1.Caption = "O") Or (Command5.Caption = "O" And Command8.Caption = "O")) And Label1.Caption = "O" And Command2.Caption = "" Then
Command2.Caption = "O"
Label1.Caption = "X"
End If
If ((Command7.Caption = "O" And Command1.Caption = "O") Or (Command5.Caption = "O" And Command6.Caption = "O")) And Label1.Caption = "O" And Command4.Caption = "" Then
Command4.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "O" And Command8.Caption = "O") Or (Command1.Caption = "O" And Command9.Caption = "O") Or (Command3.Caption = "O" And Command7.Caption = "O") Or (Command4.Caption = "O" And Command6.Caption = "O")) And Label1.Caption = "O" And Command5.Caption = "" Then
Command5.Caption = "O"
Label1.Caption = "X"
End If
If ((Command4.Caption = "O" And Command5.Caption = "O") Or (Command3.Caption = "O" And Command9.Caption = "O")) And Label1.Caption = "O" And Command6.Caption = "" Then
Command6.Caption = "O"
Label1.Caption = "X"
End If
If ((Command3.Caption = "O" And Command5.Caption = "O") Or (Command1.Caption = "O" And Command4.Caption = "O") Or (Command8.Caption = "O" And Command9.Caption = "O")) And Label1.Caption = "O" And Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If ((Command1.Caption = "O" And Command5.Caption = "O") Or (Command3.Caption = "O" And Command6.Caption = "O") Or (Command8.Caption = "O" And Command7.Caption = "O")) And Label1.Caption = "O" And Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "O" And Command5.Caption = "O") Or (Command7.Caption = "O" And Command9.Caption = "O")) And Label1.Caption = "O" And Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If





If ((Command2.Caption = "X" And Command3.Caption = "X") Or (Command5.Caption = "X" And Command9.Caption = "X") Or (Command4.Caption = "X" And Command7.Caption = "X")) And Label1.Caption = "O" And Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "X" And Command1.Caption = "X") Or (Command5.Caption = "X" And Command7.Caption = "X") Or (Command6.Caption = "X" And Command9.Caption = "X")) And Label1.Caption = "O" And Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If ((Command3.Caption = "X" And Command1.Caption = "X") Or (Command5.Caption = "X" And Command8.Caption = "X")) And Label1.Caption = "O" And Command2.Caption = "" Then
Command2.Caption = "O"
Label1.Caption = "X"
End If
If ((Command7.Caption = "X" And Command1.Caption = "X") Or (Command5.Caption = "X" And Command6.Caption = "X")) And Label1.Caption = "O" And Command4.Caption = "" Then
Command4.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "X" And Command8.Caption = "X") Or (Command1.Caption = "X" And Command9.Caption = "X") Or (Command3.Caption = "X" And Command7.Caption = "X") Or (Command4.Caption = "X" And Command6.Caption = "X")) And Label1.Caption = "O" And Command5.Caption = "" Then
Command5.Caption = "O"
Label1.Caption = "X"
End If
If ((Command4.Caption = "X" And Command5.Caption = "X") Or (Command3.Caption = "X" And Command9.Caption = "X")) And Label1.Caption = "O" And Command6.Caption = "" Then
Command6.Caption = "O"
Label1.Caption = "X"
End If
If ((Command3.Caption = "X" And Command5.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X") Or (Command8.Caption = "X" And Command9.Caption = "X")) And Label1.Caption = "O" And Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If ((Command1.Caption = "X" And Command5.Caption = "X") Or (Command3.Caption = "X" And Command6.Caption = "X") Or (Command8.Caption = "X" And Command7.Caption = "X")) And Label1.Caption = "O" And Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If ((Command2.Caption = "X" And Command5.Caption = "X") Or (Command7.Caption = "X" And Command9.Caption = "X")) And Label1.Caption = "O" And Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If
End If

 


If Option4.Value = True Or Option5.Value = True Then
If Command1.Caption = "O" And (Command3.Caption = "X" Or Command7.Caption = "X") And Label1.Caption = "O" And Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If

If Command9.Caption = "O" And (Command3.Caption = "X" Or Command7.Caption = "X") And Label1.Caption = "O" And Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If

If Command7.Caption = "O" And (Command1.Caption = "X" Or Command9.Caption = "X") And Label1.Caption = "O" And Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If

If Command3.Caption = "O" And (Command1.Caption = "X" Or Command9.Caption = "X") And Label1.Caption = "O" And Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If



If Command1.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "" And Command4.Caption = "" And Command7.Caption = "" And Label1.Caption = "O" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption = "X" And Command8.Caption = "X" And Command6.Caption = "" And Command7.Caption = "" And Command9.Caption = "" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command6.Caption = "X" And Command8.Caption = "X" And Command3.Caption = "" And Command7.Caption = "" And Command9.Caption = "" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command6.Caption = "X" And Command7.Caption = "X" And Command3.Caption = "" And Command8.Caption = "" And Command9.Caption = "" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If




If Command5.Caption = "O" And Command9.Caption = "O" And Command8.Caption = "" And Command3.Caption = "" And Command7.Caption = "" And Label1.Caption = "O" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If Command5.Caption = "O" And Command7.Caption = "O" And Command8.Caption = "" And Command1.Caption = "" And Command9.Caption = "" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If


If Command1.Caption = "O" And Command6.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If Command1.Caption = "O" And Command8.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption = "O" And Command4.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption = "O" And Command8.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command7.Caption = "O" And Command2.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
If Command7.Caption = "O" And Command6.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command9.Caption = "O" And Command2.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If Command9.Caption = "O" And Command4.Caption = "X" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If


If Command1.Caption <> "" And Command9.Caption <> "" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption <> "" And Command7.Caption <> "" And Label3.Caption = "1" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If


If Command1.Caption = "O" And Command5.Caption = "O" And Command4.Caption = "" And Command7.Caption = "" And Command3.Caption = "" And Label1.Caption = "O" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption = "O" And Command5.Caption = "O" And Command6.Caption = "" And Command1.Caption = "" And Command9.Caption = "" And Label1.Caption = "O" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If



End If



If Option5.Value = True Then
If Label3.Caption = "0" And Label1.Caption = "O" Then
Randomize
b = Int(Rnd * 20) + 1
Select Case b
Case 1 To 10

a = Int(Rnd * 4) + 1
Select Case a
Case 1
Command1.Caption = "O"
Case 2
Command3.Caption = "O"
Case 3
Command7.Caption = "O"
Case 4
Command9.Caption = "O"

End Select
Case 11 To 16
Command5.Caption = "O"
Case 17 To 20
a = Int(Rnd * 4) + 1
Select Case a
Case 1
Command2.Caption = "O"
Case 2
Command4.Caption = "O"
Case 3
Command6.Caption = "O"
Case 4
Command8.Caption = "O"
End Select
End Select
Label1.Caption = "X"
End If


If Command1.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command9.Caption = "" And Command8.Caption = "" And Command6.Caption = "" Then
Randomize
d = Int(Rnd * 10) + 1
Select Case d
Case 1 To 4
Command9.Caption = "O"
Case 5 To 7
Command8.Caption = "O"
Case 8 To 10
Command6.Caption = "O"
End Select
Label1.Caption = "X"
End If
If Command3.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command7.Caption = "" And Command8.Caption = "" And Command4.Caption = "" Then
Randomize
d = Int(Rnd * 10) + 1
Select Case d
Case 1 To 4
Command7.Caption = "O"
Case 5 To 7
Command8.Caption = "O"
Case 8 To 10
Command4.Caption = "O"
End Select
Label1.Caption = "X"
End If
If Command7.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command3.Caption = "" And Command2.Caption = "" And Command6.Caption = "" Then
Randomize
d = Int(Rnd * 10) + 1
Select Case d
Case 1 To 4
Command3.Caption = "O"
Case 5 To 7
Command2.Caption = "O"
Case 8 To 10
Command6.Caption = "O"
End Select
Label1.Caption = "X"
End If
If Command9.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command1.Caption = "" And Command4.Caption = "" And Command2.Caption = "" Then
Randomize
d = Int(Rnd * 10) + 1
Select Case d
Case 1 To 4
Command1.Caption = "O"
Case 5 To 7
Command2.Caption = "O"
Case 8 To 10
Command4.Caption = "O"
End Select
Label1.Caption = "X"
End If


If Command5.Caption = "O" And Command1.Caption = "X" And Label1.Caption = "O" And Label2.Caption = "O" And Label3.Caption = "1" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command5.Caption = "O" And Command3.Caption = "X" And Label1.Caption = "O" And Label2.Caption = "O" And Label3.Caption = "1" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
If Command5.Caption = "O" And Command7.Caption = "X" And Label1.Caption = "O" And Label2.Caption = "O" And Label3.Caption = "1" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If



End If




If Option3.Value = True Or Option4.Value = True Or Option5.Value = True Then
If Command1.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
If Command9.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
If Command7.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
If Command3.Caption = "O" And Command5.Caption = "X" And Label1.Caption = "O" And Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If

If Command3.Caption = "X" And Command5.Caption = "O" And Command7.Caption = "X" And Label1.Caption = "O" And Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If
If Command1.Caption = "X" And Command5.Caption = "O" And Command9.Caption = "X" And Label1.Caption = "O" And Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If
If Label3.Caption = "0" And Label1.Caption = "O" Then
Randomize
a = Int(Rnd * 4) + 1
Select Case a
Case 1
Command1.Caption = "O"
Case 2
Command3.Caption = "O"
Case 3
Command7.Caption = "O"
Case 4
Command9.Caption = "O"

End Select
Label1.Caption = "X"
End If




If Label1.Caption = "O" Then
If Command5.Caption = "" Then
Command5.Caption = "O"
Label1.Caption = "X"
End If
End If
If Label1.Caption = "O" Then
If Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
End If

If Label1.Caption = "O" Then
If Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
End If



If Label1.Caption = "O" Then
If Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
End If
If Label1.Caption = "O" Then
If Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
End If
If Label1.Caption = "O" Then
If Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If
End If

If Label1.Caption = "O" Then
If Command2.Caption = "" Then
Command2.Caption = "O"
Label1.Caption = "X"
End If
End If
If Label1.Caption = "O" Then
If Command4.Caption = "" Then
Command4.Caption = "O"
Label1.Caption = "X"
End If
End If
If Label1.Caption = "O" Then
If Command6.Caption = "" Then
Command6.Caption = "O"
Label1.Caption = "X"
End If
End If
End If

For x = 1 To 50 Step 1
If Label1.Caption = "O" Then
Randomize
b = Int(Rnd * 9) + 1
Select Case b
Case 1
If Command1.Caption = "" Then
Command1.Caption = "O"
Label1.Caption = "X"
End If
Case 2
If Command2.Caption = "" Then
Command2.Caption = "O"
Label1.Caption = "X"
End If
Case 3
If Command3.Caption = "" Then
Command3.Caption = "O"
Label1.Caption = "X"
End If
Case 4
If Command4.Caption = "" Then
Command4.Caption = "O"
Label1.Caption = "X"
End If
Case 5
If Command5.Caption = "" Then
Command5.Caption = "O"
Label1.Caption = "X"
End If
Case 6
If Command6.Caption = "" Then
Command6.Caption = "O"
Label1.Caption = "X"
End If
Case 7
If Command7.Caption = "" Then
Command7.Caption = "O"
Label1.Caption = "X"
End If
Case 8
If Command8.Caption = "" Then
Command8.Caption = "O"
Label1.Caption = "X"
End If
Case 9
If Command9.Caption = "" Then
Command9.Caption = "O"
Label1.Caption = "X"
End If
End Select
End If
Next x
If ((Command1.Caption = "O" And Command2.Caption = "O" And Command3.Caption = "O") Or (Command1.Caption = "O" And Command4.Caption = "O" And Command7.Caption = "O") Or (Command1.Caption = "O" And Command5.Caption = "O" And Command9.Caption = "O") Or (Command5.Caption = "O" And Command2.Caption = "O" And Command8.Caption = "O") Or (Command5.Caption = "O" And Command7.Caption = "O" And Command3.Caption = "O") Or (Command9.Caption = "O" And Command6.Caption = "O" And Command3.Caption = "O") Or (Command4.Caption = "O" And Command5.Caption = "O" And Command6.Caption = "O") Or (Command7.Caption = "O" And Command8.Caption = "O" And Command9.Caption = "O")) And Label7.Caption = "" Then
Label7.Caption = "O wins"
Label6.Caption = Val(Label6.Caption) + 1
End If
On Error Resume Next
Form1.Height = 3285
Form1.Width = 6435
End Sub
