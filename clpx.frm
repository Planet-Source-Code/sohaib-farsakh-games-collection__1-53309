VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complex equation editor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text23 
      Height          =   495
      Left            =   4320
      TabIndex        =   29
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text22 
      Height          =   495
      Left            =   3840
      TabIndex        =   28
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text21 
      Height          =   495
      Left            =   2160
      TabIndex        =   26
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text20 
      Height          =   495
      Left            =   1680
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text19 
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   4320
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004000&
      Caption         =   "y=ax^b"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004000&
      Caption         =   "y=a^(bx+c)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004000&
      Caption         =   "y=a*Abs(bx+c)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004000&
      Caption         =   "y=a*Tan(bx+c)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004000&
      Caption         =   "y=a*Cos(bx+c)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Caption         =   "y=a*Sin(bx+c)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "y=ax^5+bx^4+cx^3+dx^2+ex+f"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form11.Option1.Tag = Text7.Text
Form11.Option2.Tag = Text8.Text
Form11.Option3.Tag = Text9.Text
Form11.Option4.Tag = Text10.Text
Form11.Option5.Tag = Text11.Text
Form11.Option6.Tag = Text12.Text
Form11.Option7.Tag = Text13.Text
Form11.Option8.Tag = Text14.Text
Form11.Option9.Tag = Text15.Text
Form11.Option10.Tag = Text22.Text
Form11.Option11.Tag = Text23.Text
Form11.Command1.Tag = Text1.Text
Form11.Command2.Tag = Text2.Text
Form11.Command3.Tag = Text3.Text
Form11.Command4.Tag = Text4.Text
Form11.Command5.Tag = Text5.Text
Form11.Command6.Tag = Text6.Text
Form11.Label9.Caption = Text19.Text
Form11.Command8.Tag = Text20.Text
Form11.Command9.Tag = Text21.Text
Form11.Command10.Tag = Text16.Text
Form11.Command11.Tag = Text17.Text
Form11.Command12.Tag = Text18.Text
Form2.Hide
End Sub

Private Sub Command2_Click()
Form2.Hide
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
KeyPreview = True
If KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub
