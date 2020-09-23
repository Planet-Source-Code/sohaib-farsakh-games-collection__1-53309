VERSION 5.00
Begin VB.Form frmGameOver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Game Over!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmGameOver.Visible = False
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
frmGame.txtScore.Text = "0"
fiLastPressed = 0
fiNumLeft = 9

'set initial start speeds for each timer
frmGame.Timer1.Interval = 1000
frmGame.Timer2.Interval = 1000
frmGame.Timer3.Interval = 1000
frmGame.Timer4.Interval = 1000
frmGame.Timer5.Interval = 1000
frmGame.Timer6.Interval = 1000
frmGame.Timer7.Interval = 1000
frmGame.Timer8.Interval = 1000
frmGame.Timer9.Interval = 1000


End Sub
