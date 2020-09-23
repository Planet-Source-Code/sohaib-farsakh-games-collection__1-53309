VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Card Back"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Picture from file"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   5040
      Picture         =   "deck.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   11
      Left            =   5040
      Picture         =   "deck.frx":3C06
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   10
      Left            =   4080
      Picture         =   "deck.frx":780C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   9
      Left            =   3120
      Picture         =   "deck.frx":B412
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   8
      Left            =   2160
      Picture         =   "deck.frx":F018
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   7
      Left            =   1200
      Picture         =   "deck.frx":12C1E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   6
      Left            =   240
      Picture         =   "deck.frx":16824
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   5
      Left            =   5040
      Picture         =   "deck.frx":1A42A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   4
      Left            =   4080
      Picture         =   "deck.frx":1E030
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   3
      Left            =   3120
      Picture         =   "deck.frx":21C36
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   2
      Left            =   2160
      Picture         =   "deck.frx":2583C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   1
      Left            =   1200
      Picture         =   "deck.frx":29442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   0
      Left            =   240
      Picture         =   "deck.frx":2D048
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.Filter = "Gif Files|*.gif|Bmp Files|*.bmp|All Files|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Image2.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub Command2_Click()
For i = 0 To 47
If Form1.Image2(i).Picture = Form1.Image1(0).Picture Then
Form1.Image2(i).Picture = Image2.Picture
End If
Next i
Form1.Image1(0).Picture = Image2.Picture

Form2.Hide
End Sub

Private Sub Command3_Click()
Form2.Hide
End Sub

Private Sub Image1_Click(Index As Integer)
Image2.Picture = Image1(Index).Picture
End Sub

