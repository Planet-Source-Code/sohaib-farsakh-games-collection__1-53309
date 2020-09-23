VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "create gradient"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "vertical"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "horizontal"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color#2"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color#1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.color
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowColor
Picture2.BackColor = CommonDialog1.color
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim Color1, Color2 As Long
Dim R, G, B, R2, G2, B2 As Integer
Dim temp As Long
Color1 = Picture1.BackColor
Color2 = Picture2.BackColor
temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
B = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255
Form1.Timer1.Enabled = False
Form1.Picture1.DrawWidth = 1
If Option1.Value = True Then
a = (R2 - R) / (Form1.tg3.Tag - Form1.tg1.Tag)
c = (G2 - G) / (Form1.tg3.Tag - Form1.tg1.Tag)
d = (B2 - B) / (Form1.tg3.Tag - Form1.tg1.Tag)
For i = Form1.tg1.Tag To Form1.tg3.Tag
Form1.Picture1.Line (i, Form1.tg2.Tag)-(i, Form1.tg4.Tag), RGB(R, G, B)
R = R + a
G = G + c
B = B + d
Next i
End If
If Option2.Value = True Then
a = (R2 - R) / (Form1.tg4.Tag - Form1.tg2.Tag)
c = (G2 - G) / (Form1.tg4.Tag - Form1.tg2.Tag)
d = (B2 - B) / (Form1.tg4.Tag - Form1.tg2.Tag)
For i = Form1.tg2.Tag To Form1.tg4.Tag
Form1.Picture1.Line (Form1.tg1.Tag, i)-(Form1.tg3.Tag, i), RGB(R, G, B)
R = R + a
G = G + c
B = B + d
Next i
End If
Form1.Timer1.Enabled = True
Unload Form2
End Sub

Private Sub Command4_Click()
Unload Form2
End Sub
