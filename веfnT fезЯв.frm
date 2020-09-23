VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ÊÍæíá æÍÏÇÊ"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo11 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0000
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0016
      TabIndex        =   18
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0042
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0058
      TabIndex        =   17
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0084
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0091
      TabIndex        =   16
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":00AB
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":00B8
      TabIndex        =   15
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":00D2
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":00F1
      TabIndex        =   14
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0129
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0148
      TabIndex        =   13
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "äÝÐ"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0180
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":019F
      TabIndex        =   11
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":01F3
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0212
      TabIndex        =   10
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Íæá"
      Height          =   615
      Left            =   3720
      TabIndex        =   9
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0266
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0288
      TabIndex        =   5
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":02C5
      Left            =   2880
      List            =   "ÊÍæíá æÍÏÇÊ.frx":02E7
      TabIndex        =   2
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ÊÍæíá æÍÏÇÊ.frx":0324
      Left            =   1320
      List            =   "ÊÍæíá æÍÏÇÊ.frx":0337
      TabIndex        =   0
      Text            =   "ÊÍæíá æÍÏÇÊ"
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   ":ÇáÌæÇÈ"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Åáì"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Íæá ãä"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Combo1.Text = "ÇáØæá" Then
If Combo2.Text = "ßáã" Then
Label5.Caption = Val(Text1.Text) * 1000000
End If
If Combo2.Text = "ãÊÑ" Then
Label5.Caption = Val(Text1.Text) * 1000
End If
If Combo2.Text = "ÏÓã" Then
Label5.Caption = Val(Text1.Text) * 100
End If
If Combo2.Text = "Óã" Then
Label5.Caption = Val(Text1.Text) * 10
End If
If Combo2.Text = "ãã" Then
Label5.Caption = Val(Text1.Text)
End If
If Combo2.Text = "ãíá" Then
Label5.Caption = Val(Text1.Text) * 1609000
End If
If Combo2.Text = "íÇÑÏÉ" Then
Label5.Caption = Val(Text1.Text) * 914.4
End If
If Combo2.Text = "ÞÏã" Then
Label5.Caption = Val(Text1.Text) * 304.8
End If
If Combo2.Text = "ÈæÕÉ" Then
Label5.Caption = Val(Text1.Text) * 25.4
End If
If Combo2.Text = "ÓäÉ ÖæÆíÉ" Then
Label5.Caption = Val(Text1.Text) * 9.4608E+18
End If
If Combo3.Text = "ßáã" Then
Label4.Caption = Val(Label5.Caption) / 1000000
End If
If Combo3.Text = "ãÊÑ" Then
Label4.Caption = Val(Label5.Caption) / 1000
End If
If Combo3.Text = "ÏÓã" Then
Label4.Caption = Val(Label5.Caption) / 100
End If
If Combo3.Text = "Óã" Then
Label4.Caption = Val(Label5.Caption) / 10
End If
If Combo3.Text = "ãã" Then
Label4.Caption = Val(Label5.Caption)
End If
If Combo3.Text = "ãíá" Then
Label4.Caption = Val(Label5.Caption) / 1609000
End If
If Combo3.Text = "íÇÑÏÉ" Then
Label4.Caption = Val(Label5.Caption) / 914.4
End If
If Combo3.Text = "ÞÏã" Then
Label4.Caption = Val(Label5.Caption) / 304.8
End If
If Combo3.Text = "ÈæÕÉ" Then
Label4.Caption = Val(Label5.Caption) / 25.4
End If
If Combo3.Text = "ÓäÉ ÖæÆíÉ" Then
Label4.Caption = Val(Label5.Caption) / 9.4608E+18
End If
End If
If Combo1.Text = "ÇáãÓÇÍÉ" Then
If Combo4.Text = "Óã ãÑÈÚ" Then
Label5.Caption = Val(Text1.Text)
End If
If Combo4.Text = "ãÊÑ ãÑÈÚ" Then
Label5.Caption = Val(Text1.Text) * 10000
End If
If Combo4.Text = "Ïæäã" Then
Label5.Caption = Val(Text1.Text) * 10000000
End If
If Combo4.Text = "ßáã ãÑÈÚ" Then
Label5.Caption = Val(Text1.Text) * 10000000000#
End If
If Combo4.Text = "åßÊÇÑ" Then
Label5.Caption = Val(Text1.Text) * 100000000
End If
If Combo4.Text = "ãíá ãÑÈÚ" Then
Label5.Caption = Val(Text1.Text) * 25900000000#
End If
If Combo4.Text = "ÈæÕÉ ãÑÈÚÉ" Then
Label5.Caption = Val(Text1.Text) * 6.4516
End If
If Combo4.Text = "ÞÏã ãÑÈÚ" Then
Label5.Caption = Val(Text1.Text) * 929.316
End If
If Combo4.Text = "ÝÏÇä" Then
Label5.Caption = Val(Text1.Text) * 40468790.47
End If
If Combo5.Text = "ßáã ãÑÈÚ" Then
Label4.Caption = Val(Label5.Caption) / 10000000000#
End If
If Combo5.Text = "Ïæäã" Then
Label4.Caption = Val(Label5.Caption) / 10000000
End If
If Combo5.Text = "ãÊÑ ãÑÈÚ" Then
Label4.Caption = Val(Label5.Caption) / 10000
End If
If Combo5.Text = "Óã ãÑÈÚ" Then
Label4.Caption = Val(Label5.Caption)
End If
If Combo5.Text = "åßÊÇÑ" Then
Label4.Caption = Val(Label5.Caption) / 100000000
End If
If Combo5.Text = "ãíá ãÑÈÚ" Then
Label4.Caption = Val(Label5.Caption) / 25900000000#
End If
If Combo5.Text = "ÈæÕÉ ãÑÈÚÉ" Then
Label4.Caption = Val(Label5.Caption) / 6.4516
End If
If Combo5.Text = "ÞÏã ãÑÈÚ" Then
Label4.Caption = Val(Label5.Caption) / 929.316
End If
If Combo5.Text = "ÝÏÇä" Then
Label4.Caption = Val(Label5.Caption) / 40468790.47
End If
End If
If Combo1.Text = "ÇáÒãä" Then
If Combo6.Text = "íæã" Then
Label5.Caption = Val(Text1.Text) * 86400
End If
If Combo6.Text = "ÓÇÚÉ" Then
Label5.Caption = Val(Text1.Text) * 3600
End If
If Combo6.Text = "ÏÞíÞÉ" Then
Label5.Caption = Val(Text1.Text) * 60
End If
If Combo6.Text = "ËÇäíÉ" Then
Label5.Caption = Val(Text1.Text)
End If
If Combo6.Text = "ÃÓÈæÚ" Then
Label5.Caption = Val(Text1.Text) * 604800
End If
If Combo6.Text = "ÔåÑ" Then
Label5.Caption = Val(Text1.Text) * 2592000
End If
If Combo6.Text = "ÓäÉ" Then
Label5.Caption = Val(Text1.Text) * 31557600
End If
If Combo6.Text = "ÚÞÏ" Then
Label5.Caption = Val(Text1.Text) * 315576000
End If
If Combo6.Text = "ÞÑä" Then
Label5.Caption = Val(Text1.Text) * 3155760000#
End If
If Combo7.Text = "íæã" Then
Label4.Caption = Val(Label5.Caption) / 86400
End If
If Combo7.Text = "ÓÇÚÉ" Then
Label4.Caption = Val(Label5.Caption) / 3600
End If
If Combo7.Text = "ÏÞíÞÉ" Then
Label4.Caption = Val(Label5.Caption) / 60
End If
If Combo7.Text = "ËÇäíÉ" Then
Label4.Caption = Val(Label5.Caption)
End If
If Combo7.Text = "ÃÓÈæÚ" Then
Label4.Caption = Val(Label5.Caption) / 604800
End If
If Combo7.Text = "ÔåÑ" Then
Label4.Caption = Val(Label5.Caption) / 2592000
End If
If Combo7.Text = "ÓäÉ" Then
Label4.Caption = Val(Label5.Caption) / 31557600
End If
If Combo7.Text = "ÚÞÏ" Then
Label4.Caption = Val(Label5.Caption) / 315576000
End If
If Combo7.Text = "ÞÑä" Then
Label4.Caption = Val(Label5.Caption) / 3155760000#
End If
End If
If Combo1.Text = "ÇáÍÑÇÑÉ" Then
If Combo8.Text = "ãÆæí" Then
Label5.Caption = Val(Text1.Text)
End If
If Combo8.Text = "ÝåÑäåÇíÊ" Then
Label5.Caption = (Val(Text1.Text) - 32) / 1.8
End If
If Combo8.Text = "ßáÝä" Then
Label5.Caption = Val(Text1.Text) - 273
End If
If Combo9.Text = "ãÆæí" Then
Label4.Caption = Val(Label5.Caption)
End If
If Combo9.Text = "ÝåÑäåÇíÊ" Then
Label4.Caption = Val(Label5.Caption) * 1.8 + 32
End If
If Combo9.Text = "ßáÝä" Then
Label4.Caption = Val(Label5.Caption) + 273
End If
End If
If Combo1.Text = "ÇáæÒä" Then
If Combo10.Text = "Øä" Then
Label5.Caption = Val(Text1.Text)
End If
If Combo10.Text = "ßíáæ ÛÑÇã" Then
Label5.Caption = Val(Text1.Text) / 1000
End If
If Combo10.Text = "ÛÑÇã" Then
Label5.Caption = Val(Text1.Text) / 1000000
End If
If Combo10.Text = "ÑØá" Then
Label5.Caption = Val(Text1.Text) / 2205
End If
If Combo10.Text = "ÃæÞíÉ" Then
Label5.Caption = Val(Text1.Text) / 35000
End If
If Combo10.Text = "ÃæäÕÉ" Then
Label5.Caption = Val(Text1.Text) / 35280
End If
If Combo11.Text = "Øä" Then
Label4.Caption = Val(Label5.Caption)
End If
If Combo11.Text = "ßíáæ ÛÑÇã" Then
Label4.Caption = Val(Label5.Caption) * 1000
End If
If Combo11.Text = "ÛÑÇã" Then
Label4.Caption = Val(Label5.Caption) * 1000000
End If
If Combo11.Text = "ÑØá" Then
Label4.Caption = Val(Label5.Caption) * 2205
End If
If Combo11.Text = "ÃæÞíÉ" Then
Label4.Caption = Val(Label5.Caption) * 35000
End If
If Combo11.Text = "ÃæäÕÉ" Then
Label4.Caption = Val(Label5.Caption) * 35280
End If
End If
End Sub

Private Sub Command2_Click()
If Combo1.Text = "ÇáØæá" Then
Combo2.Visible = True
Combo3.Visible = True
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
Combo7.Visible = False
Combo8.Visible = False
Combo9.Visible = False
Combo10.Visible = False
Combo11.Visible = False
End If
If Combo1.Text = "ÇáãÓÇÍÉ" Then
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = True
Combo5.Visible = True
Combo6.Visible = False
Combo7.Visible = False
Combo8.Visible = False
Combo9.Visible = False
Combo10.Visible = False
Combo11.Visible = False
End If
If Combo1.Text = "ÇáÒãä" Then
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = True
Combo7.Visible = True
Combo8.Visible = False
Combo9.Visible = False
Combo10.Visible = False
Combo11.Visible = False
End If
If Combo1.Text = "ÇáÍÑÇÑÉ" Then
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
Combo7.Visible = False
Combo8.Visible = True
Combo9.Visible = True
Combo10.Visible = False
Combo11.Visible = False
End If
If Combo1.Text = "ÇáæÒä" Then
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo6.Visible = False
Combo7.Visible = False
Combo8.Visible = False
Combo9.Visible = False
Combo10.Visible = True
Combo11.Visible = True
End If
End Sub

Private Sub Form_Load()
Combo2.Visible = True
Combo3.Visible = True
If Combo1.Text = "ÇáØæá" Then
Combo2.Visible = True
Combo3.Visible = True
Combo4.Visible = False
Combo5.Visible = False
End If
If Combo1.Text = "ÇáãÓÇÍÉ" Then
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = True
Combo5.Visible = True
End If
End Sub

