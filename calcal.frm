VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "Backspace"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   2760
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "C"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1560
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "AC"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "log"
      Height          =   495
      Index           =   3
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "ln"
      Height          =   495
      Index           =   2
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "+/-"
      Height          =   495
      Index           =   1
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "Exp"
      Height          =   495
      Index           =   0
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton trig 
      BackColor       =   &H00C0C000&
      Caption         =   "Tan"
      Height          =   495
      Index           =   2
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton trig 
      BackColor       =   &H00C0C000&
      Caption         =   "Cos"
      Height          =   495
      Index           =   1
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton trig 
      BackColor       =   &H00C0C000&
      Caption         =   "Sin"
      Height          =   495
      Index           =   0
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton pow 
      BackColor       =   &H00C0C000&
      Caption         =   "x^y"
      Height          =   495
      Index           =   2
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton pow 
      BackColor       =   &H00C0C000&
      Caption         =   "x^3"
      Height          =   495
      Index           =   1
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton pow 
      BackColor       =   &H00C0C000&
      Caption         =   "x^2"
      Height          =   495
      Index           =   0
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton brck 
      BackColor       =   &H00C0C000&
      Caption         =   ")"
      Height          =   495
      Index           =   1
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton brck 
      BackColor       =   &H00C0C000&
      Caption         =   "("
      Height          =   495
      Index           =   0
      Left            =   1200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00C0C000&
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00C0C000&
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00C0C000&
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton operation 
      BackColor       =   &H00C0C000&
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "="
      Height          =   495
      Left            =   5160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdpoint 
      BackColor       =   &H00808000&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   5160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   5160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   5160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   3480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label3 
      Height          =   15
      Left            =   600
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastpress As String

Private Sub brck_Click(Index As Integer)
Select Case Index
Case 0
Label2.Caption = Label2.Caption + "("
Label3.Caption = Label3.Caption + "("

Label1.Caption = "0"
Case 1
Label2.Caption = Label2.Caption + Label1.Caption + ")"
Label1.Caption = "0"
lastpress = "brck"
End Select
End Sub

Private Sub cmdpoint_Click()
Label1.Caption = Label1.Caption + "."
cmdpoint.Enabled = False

End Sub

Private Sub Command1_Click(Index As Integer)
If Label1.Enabled = True Then
If Label1.Caption = "0" Then
Label1.Caption = Command1(Index).Caption
Else
Label1.Caption = Label1.Caption + Command1(Index).Caption
End If
End If
End Sub

Private Sub Command2_Click()
cmdpoint.Enabled = True

End Sub

Private Sub Command3_Click(Index As Integer)
Select Case Index
Case 1
Label1.Caption = -Val(Label1.Caption)
Case 0
Label1.Caption = Label1.Caption + "E+"
Case 2
Label2.Caption = Label2.Caption + "ln"
Case 3
Label2.Caption = Label2.Caption + "log"
End Select
End Sub

Private Sub Command4_Click()
cmdpoint.Enabled = True

End Sub

Private Sub Command5_Click()
cmdpoint.Enabled = True

End Sub

Private Sub Command6_Click()
cmdpoint.Enabled = True
D = Len(Label1.Caption)
Select Case Len(Label1.Caption)
Case 1
Label1.Caption = "0"
Case Else
Label1.Caption = Left(Label1.Caption, D - 1)
End Select
End Sub

Private Sub operation_Click(Index As Integer)
cmdpoint.Enabled = True
If lastpress <> "brck" Then
Label2.Caption = Label2.Caption + Label1.Caption + operation(Index).Caption
If lastpress = "trig" Then
a = Val(Label1.Caption) * 180 / 3.14159265358979
Label3.Caption = Label3.Caption + "(" + Str(a) + ")" + operation(Index).Caption
Else
Label3.Caption = Label3.Caption + Label1.Caption + operation(Index).Caption
End If
Label1.Caption = "0"

Else
Label2.Caption = Label2.Caption + operation(Index).Caption
Label3.Caption = Label3.Caption + operation(Index).Caption

Label1.Caption = "0"
lastpress = ""
End If
End Sub

Private Sub pow_Click(Index As Integer)

cmdpoint.Enabled = True

Select Case Index
Case 0
Label2.Caption = Label2.Caption + Label1.Caption + "^2"
Label1.Caption = "0"
Case 1
Label2.Caption = Label2.Caption + Label1.Caption + "^3"
Label1.Caption = "0"
Case 2
lastpress = "^"
Label2.Caption = Label2.Caption + Label1.Caption + "^"
Label1.Caption = "0"
End Select
End Sub

Private Sub trig_Click(Index As Integer)
Select Case Index
Case 0
b = "sin"
Case 1
b = "cos"
Case 2
b = "tan"
End Select
Label2.Caption = Label2.Caption + b
lastpress = "trig"
End Sub
