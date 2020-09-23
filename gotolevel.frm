VERSION 5.00
Begin VB.Form frmlevel 
   Caption         =   "Go to level"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the password:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the level number:"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmlevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
level = Int(Val(Text1.Text))
If level = 1 Then
gotolevel (1)
frmMain.Label6.Caption = "3"
livescore = 200
score = 0
frmMain.Label3.Caption = "0"
Else
If level < 1 Or level > 10 Then
MsgBox ("Levels are from 1 to 10")
Exit Sub
Else
If Text2.Text = password(level) Then
gotolevel (level)
livescore = 200
score = 0
frmMain.Label3.Caption = "0"
frmMain.Label6.Caption = "3"
Else
MsgBox ("The password doesn't match the level number!")
Exit Sub
End If
End If
End If
frmlevel.Hide
End Sub

Private Sub Command2_Click()
frmlevel.Hide
End Sub

