VERSION 5.00
Begin VB.Form frmscores 
   Caption         =   "Top Scores"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "Playing until crash"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Completing levels"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "The top ten scores in snake"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmscores.Hide
End Sub

Private Sub Option1_Click()
displayscores ("complete")
End Sub

Private Sub Option2_Click()
displayscores ("crash")

End Sub
