VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Drawing Program"
   ClientHeight    =   8310
   ClientLeft      =   510
   ClientTop       =   1110
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   7800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   9
      Height          =   6975
      Left            =   1320
      ScaleHeight     =   600
      ScaleMode       =   0  'User
      ScaleWidth      =   900
      TabIndex        =   13
      Tag             =   "2"
      Top             =   720
      Width           =   9735
      Begin VB.Frame Frame2 
         Caption         =   "image size"
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2400
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "OK"
            Height          =   375
            Left            =   2400
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   840
            TabIndex        =   22
            Text            =   "900"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   840
            TabIndex        =   21
            Text            =   "600"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "width:"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "height:"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "paint.frx":0000
      Left            =   120
      List            =   "paint.frx":0025
      TabIndex        =   12
      Text            =   "1"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fore Color"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back Color"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ToolBox"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
      Begin MSForms.ToggleButton tg9 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "circle"
         Top             =   4080
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":004F
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg8 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "rectangle"
         Top             =   3600
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":0859
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg7 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "line"
         Top             =   3120
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":1063
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg6 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "air brush"
         Top             =   2640
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":186D
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg5 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "brush"
         Top             =   2160
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":2077
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg4 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Tag             =   "0"
         ToolTipText     =   "pen"
         Top             =   1680
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":2881
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg3 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Tag             =   "0"
         ToolTipText     =   "pick color"
         Top             =   1200
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":308B
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg2 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         ToolTipText     =   "eraser"
         Top             =   720
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":3895
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton tg1 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Tag             =   "0"
         ToolTipText     =   "select area"
         Top             =   240
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1508;873"
         Value           =   "0"
         PicturePosition =   262148
         Picture         =   "paint.frx":409F
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Label Label3 
      Caption         =   "0 , 0"
      Height          =   495
      Left            =   9000
      TabIndex        =   17
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "drawing width"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut Image"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy Image"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu linee 
         Caption         =   "-"
      End
      Begin VB.Menu clr 
         Caption         =   "&clear Selected Area"
      End
      Begin VB.Menu pnt 
         Caption         =   "&Paint Selected Area"
      End
      Begin VB.Menu clrim 
         Caption         =   "&clear Image"
      End
      Begin VB.Menu sa 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu image 
      Caption         =   "&Image"
      Begin VB.Menu ic 
         Caption         =   "&Invert colors"
         Shortcut        =   ^I
      End
      Begin VB.Menu is 
         Caption         =   "&Image size"
         Shortcut        =   ^E
      End
      Begin VB.Menu cg 
         Caption         =   "&Create gradient"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton6_Click()

End Sub

Private Sub cg_Click()
Load Form2
Form2.Show
End Sub

Private Sub clr_Click()
If tg2.Tag > tg4.Tag Then
d = tg2.Tag
tg2.Tag = tg4.Tag
tg4.Tag = d
End If
For i = tg2.Tag To tg4.Tag
Picture1.DrawWidth = 1
Picture1.Line (tg1.Tag, i)-(tg3.Tag, i), Picture1.BackColor
Next i
End Sub

Private Sub clrim_Click()
Picture1.Picture = LoadPicture()

End Sub

Private Sub Command1_Click()
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.color
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowColor
Picture1.ForeColor = CommonDialog1.color
End Sub

Private Sub Command3_Click()
Picture1.Picture = LoadPicture()

End Sub

Private Sub fh_Click()
Dim a As Single, B As Single, c As Single, d As Single
If tg1.Tag > tg3.Tag Then
d = tg1.Tag
tg1.Tag = tg3.Tag
tg3.Tag = d
End If
If tg2.Tag > tg4.Tag Then
d = tg2.Tag
tg2.Tag = tg4.Tag
tg4.Tag = d
End If
Picture2.ScaleWidth = Val(tg3.Tag - tg1.Tag)
Picture2.ScaleHeight = Val(tg4.Tag - tg2.Tag)
For a = 1 To (tg3.Tag - tg1.Tag)
For B = 1 To (tg4.Tag - tg2.Tag)
Picture2.Point(a, B) = Picture1.Point(c, d)
Next B
Next a
End Sub

Private Sub Command4_Click()
Picture1.Height = Val(Text1.Text) / 600 * Picture1.Height
Picture1.Width = Val(Text2.Text) / 900 * Picture1.Width
Frame2.Visible = False
End Sub

Private Sub Command5_Click()
Frame2.Visible = False
End Sub

Private Sub copy_Click()
Clipboard.Clear
  Clipboard.SetData Picture1.Picture
End Sub

Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetData Picture1.Picture
Picture1.Picture = LoadPicture()
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub ic_Click()

Timer1.Enabled = False
Picture1.DrawWidth = 1
Dim a, c As Long
Dim temp, color As Long
Dim R, G, O As Integer
For a = tg1.Tag To tg3.Tag Step 3
For c = tg2.Tag To tg4.Tag Step 3
color = Picture1.Point(a, c)
temp = (color And 255)
R = temp And 255
temp = Int(color / 256)
G = temp And 255
temp = Int(color / 65536)
O = temp And 255
Picture1.PSet (a, c), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a + 1, c), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a, c + 1), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a + 1, c + 1), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a - 1, c - 1), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a - 1, c), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a, c - 1), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a - 1, c + 1), RGB(255 - R, 255 - G, 255 - O)
Picture1.PSet (a + 1, c - 1), RGB(255 - R, 255 - G, 255 - O)
Next c
Next a
Timer1.Enabled = True
End Sub

Private Sub is_Click()
Frame2.Visible = True
End Sub

Private Sub new_Click()
Picture1.Picture = LoadPicture()
End Sub

Private Sub open_Click()
CommonDialog1.Filter = "Gif Files|*.gif|Bmp Files|*.bmp|All Files|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Picture1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub paste_Click()
Picture1.Picture = Clipboard.GetData
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Tag = 1
If Label1.Caption = "7" Then
Select Case tg1.Tag
Case "0"
tg1.Tag = X
tg2.Tag = Y
Case Else
a = tg1.Tag
B = tg2.Tag
Picture1.Line (a, B)-(X, Y)
tg1.Tag = "0"
tg2.Tag = "0"
End Select
End If
If Label1.Caption = "8" Then
Select Case tg1.Tag
Case "0"
tg1.Tag = X
tg2.Tag = Y
Case Else
a = tg1.Tag
B = tg2.Tag
Picture1.Line (a, B)-(a, Y)
Picture1.Line (X, Y)-(a, Y)
Picture1.Line (X, Y)-(X, B)
Picture1.Line (a, B)-(X, B)
tg1.Tag = "0"
tg2.Tag = "0"
End Select
End If
If Label1.Caption = "9" Then
Select Case tg1.Tag
Case "0"
tg1.Tag = X
tg2.Tag = Y
Case Else
a = tg1.Tag
B = tg2.Tag
c = (a + X) / 2
d = (B + Y) / 2
e = Sqr(((Y - B) ^ 2) + ((X - a) ^ 2)) / 2
Picture1.Circle (c, d), e
tg1.Tag = "0"
tg2.Tag = "0"
End Select
End If
If Label1.Caption = "3" Then
Picture1.ForeColor = Picture1.Point(X, Y)
End If
If Label1.Caption = "6" Then
Picture1.DrawWidth = 1
For i = 1 To 50
Randomize
d = Int(Rnd * ((20 * Combo1.Text + 1))) - ((10 * Combo1.Text + 1) + 1)
e = Int(Rnd * ((20 * Combo1.Text + 1))) - ((10 * Combo1.Text + 1) + 1)
f = X + d
G = Y + e
Picture1.PSet (f, G)
Next i
End If
If Label1.Caption = "1" Then
Select Case tg1.Tag
Case 0
tg1.Tag = X
tg2.Tag = Y
Case Else
Select Case tg3.Tag
Case 0
tg3.Tag = X
tg4.Tag = Y
Case Else
tg1.Tag = X
tg2.Tag = Y
tg3.Tag = 0
tg4.Tag = 0
End Select
End Select
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Label1.Caption = "5" And Picture1.Tag = 1 Then
Picture1.PSet (X, Y)
Select Case tg1.Tag
Case 0
tg1.Tag = X
tg2.Tag = Y
Case Else
Picture1.Line (tg1.Tag, tg2.Tag)-(X, Y)

tg1.Tag = X
tg2.Tag = Y
End Select
End If
If Label1.Caption = "4" And Picture1.Tag = 1 Then
Picture1.DrawWidth = 1
Picture1.PSet (X, Y)

Select Case tg1.Tag
Case 0
tg1.Tag = X
tg2.Tag = Y
Case Else
Picture1.Line (tg1.Tag, tg2.Tag)-(X, Y)

tg1.Tag = X
tg2.Tag = Y
End Select
End If
If Label1.Caption = "2" And Picture1.Tag = 1 Then
Picture1.PSet (X, Y), Picture1.BackColor
Select Case tg1.Tag
Case 0
tg1.Tag = X
tg2.Tag = Y
Case Else
Picture1.Line (tg1.Tag, tg2.Tag)-(X, Y), Picture1.BackColor
tg1.Tag = X
tg2.Tag = Y
End Select
End If
Label3.Caption = Int(X) & " , " & Int(Y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Tag = 2
If Label1.Caption = "5" Or Label1.Caption = "4" Or Label1.Caption = "2" Then
tg1.Tag = 0
tg2.Tag = 0
End If
End Sub

Private Sub pnt_Click()
If tg2.Tag > tg4.Tag Then
d = tg2.Tag
tg2.Tag = tg4.Tag
tg4.Tag = d
End If
For i = tg2.Tag To tg4.Tag
Picture1.DrawWidth = 1
Picture1.Line (tg1.Tag, i)-(tg3.Tag, i)
Next i
End Sub

Private Sub sa_Click()
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = Picture1.Width
tg4.Tag = Picture1.Height
End Sub

Private Sub tg1_Click()
If tg1.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 1
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg2_Click()
If tg2.Value = True Then
tg1.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 2
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg3_Click()
If tg3.Value = True Then
tg2.Value = False
tg1.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 3
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg4_Click()
If tg4.Value = True Then
tg2.Value = False
tg3.Value = False
tg1.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 4
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg5_Click()
If tg5.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg1.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 5
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg6_Click()
If tg6.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg1.Value = False
tg7.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 6
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg7_Click()
If tg7.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg1.Value = False
tg8.Value = False
tg9.Value = False
End If
Label1.Caption = 7
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg8_Click()
If tg8.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg1.Value = False
tg9.Value = False
End If
Label1.Caption = 8
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub tg9_Click()
If tg9.Value = True Then
tg2.Value = False
tg3.Value = False
tg4.Value = False
tg5.Value = False
tg6.Value = False
tg7.Value = False
tg8.Value = False
tg1.Value = False
End If
Label1.Caption = 9
tg1.Tag = 0
tg2.Tag = 0
tg3.Tag = 0
tg4.Tag = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Label1.Caption <> "4" And Label1.Caption <> "6" Then
Picture1.DrawWidth = Val(Combo1.Text) * 1
End If
End Sub
