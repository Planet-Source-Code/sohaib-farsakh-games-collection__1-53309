VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Setup"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "3"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "18"
      Top             =   120
      Width           =   735
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Set maximum number of rods to take at a turn"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Set the number of rods"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Val(Form2.Text1.Text) > 18 Then
Form1.Label1.Height = 6800
Form1.Height = 8130
Form1.Command1.Top = 7000
Form1.Label3.Top = 7000

Else
Form1.Label1.Height = 3400
Form1.Height = 8130 - 3400
Form1.Label3.Top = 3600
Form1.Command1.Top = 3600

End If
newgame
For i = 1 To 35
Form1.Label2(i).Visible = False
Next i
For i = 1 To Val(Text1.Text - 1)
Form1.Label2(i).Visible = True
Next i
Form1.Label3.Tag = Val(Text2.Text)
Form1.Command1.Tag = Val(Form2.Text1.Text)
Form2.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub

Private Sub UpDown1_DownClick()
If Val(Text1.Text) > 2 Then
Text1.Text = Val(Text1.Text) - 1
End If
End Sub

Private Sub UpDown1_UpClick()
If Val(Text1.Text) < 36 Then
Text1.Text = Val(Text1.Text) + 1
End If
End Sub
Private Sub UpDown2_DownClick()
If Val(Text2.Text) > 1 Then
Text2.Text = Val(Text2.Text) - 1
End If
End Sub

Private Sub UpDown2_UpClick()
If Val(Text2.Text) < 9 Then
Text2.Text = Val(Text2.Text) + 1
End If
End Sub
