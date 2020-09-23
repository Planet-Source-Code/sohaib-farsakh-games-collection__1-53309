VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vectors"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   0  'User
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Zoom out"
      Height          =   375
      Left            =   1320
      TabIndex        =   44
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zoom In"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   1875
      TabIndex        =   27
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1440
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operations on vectors"
      Height          =   3015
      Left            =   1080
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command9 
         Caption         =   "Back"
         Height          =   375
         Left            =   4440
         TabIndex        =   40
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Components of the vector"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Sum of selected vectors"
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Tag             =   "0"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<<"
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">>"
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   1920
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "angle:"
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "value:"
         Height          =   255
         Left            =   3000
         TabIndex        =   45
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "y :"
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "x :"
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   6000
      Left            =   2880
      ScaleHeight     =   400
      ScaleMode       =   0  'User
      ScaleWidth      =   556.098
      TabIndex        =   0
      Top             =   0
      Width           =   6044
      Begin VB.Frame Frame1 
         Caption         =   "New Vector"
         Height          =   3015
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1560
            TabIndex        =   15
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   2400
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Draw the vector"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   480
            TabIndex        =   10
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ends at point"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2640
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "value and angle"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "vector name:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "y:"
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   1560
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "x:"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "angle:"
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "length:"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   133.798
         X2              =   256.446
         Y1              =   371.717
         Y2              =   347.475
      End
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   26
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   25
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   22
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "angle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "end point:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu zoom 
         Caption         =   "zoom"
         Visible         =   0   'False
         Begin VB.Menu in 
            Caption         =   "In"
            Visible         =   0   'False
         End
         Begin VB.Menu out 
            Caption         =   "Out"
         End
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu vect 
      Caption         =   "Vectors"
      Begin VB.Menu newv 
         Caption         =   "New Vector"
      End
      Begin VB.Menu oper 
         Caption         =   "Operations on vectors"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim enabledraw As Boolean
Dim xpos As Double, ypos As Double
Dim vectors(1000, 7) As Variant
Dim selected(200) As Variant
Dim vectornum As Integer
Const pi = 3.14159265358979
Dim selnum As Integer
Dim zm
Private Sub drawall()
List1.clear
For i = 1 To vectornum
            If vectors(vectornum, 5) > 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) - 180
        End If
        
       ' If vectors(vectornum, 5) < 0 And vectors(vectornum, 2) < 0 Then
       ' vectors(vectornum, 5) = vectors(vectornum, 5) + 180
       ' End If

drawvector (i)
            If vectors(vectornum, 5) > 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If
        
        If vectors(vectornum, 5) < 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If

Next i
End Sub
Private Sub createcoordinates(zoom As Double)
Picture1.Scale (-10, 10)-(10, -10)

For i = -9 To 9
Picture1.Line (i, 10)-(i, -10), &HC0C0C0
If i = 0 Then Picture1.Line (0, 10)-(0, -10), vbRed
If (i) Mod 2 = 0 And Abs(i) <> 10 Then
Picture1.CurrentX = i + 0.055
Picture1.CurrentY = 0.7
Picture1.Print i * zoom
End If
Next i
For i = -9 To 9
Picture1.Line (10, i)-(-10, i), &HC0C0C0
If i = 0 Then Picture1.Line (10, 0)-(-10, 0), vbRed
If i Mod 2 = 0 And i <> 0 Then
Picture1.CurrentX = 0.1
Picture1.CurrentY = i + 0.5
Picture1.Print i * zoom
End If

Next i
Picture1.Scale (-10 * zoom, 10 * zoom)-(10 * zoom, -10 * zoom)

End Sub
Private Sub drawvector(num As Integer)
Picture1.ForeColor = vectors(num, 6)
Picture1.Line (0, 0)-(vectors(num, 2), vectors(num, 3))
List1.AddItem (vectors(num, 1))
a = Cos(Val(vectors(num, 5)) / 180 * pi) * 0.2
b = Sin(Val(vectors(num, 5)) / 180 * pi) * 0.2
If vectors(vectornum, 2) < 0 Then
a = -a
b = -b
End If
c = Val(vectors(num, 2)) - a
d = Val(vectors(num, 3)) - b
e = c + b
f = d - a
g = c - b
h = d + a
i = Val(vectors(num, 2))
j = Val(vectors(num, 3))
Picture1.DrawWidth = 2
Picture1.Line (e, f)-(g, h)
Picture1.Line (i, j)-(e, f)
Picture1.Line (i, j)-(g, h)
Picture1.DrawWidth = 1
'Form1.Caption = Str(g) + "," + Str(h)
End Sub

Private Sub clear_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text5.Text = "" Then
MsgBox ("vector must have a name")
Exit Sub
End If
For i = 1 To vectornum
If Text5.Text = vectors(i, 1) Then
MsgBox ("there is already a vector with this name")
Exit Sub
End If
Next i
vectornum = vectornum + 1
If Option3.Value = False Then
    If Option1.Value = True Then
        vectors(vectornum, 2) = Val(Text1.Text) * Cos(pi / 180 * Val(Text2.Text))
        vectors(vectornum, 3) = Val(Text1.Text) * Sin(pi / 180 * Val(Text2.Text))
        vectors(vectornum, 4) = Val(Text1.Text)
        vectors(vectornum, 5) = Val(Text2.Text)

    End If
    If Option2.Value = True Then
        vectors(vectornum, 2) = Val(Text3.Text)
        vectors(vectornum, 3) = Val(Text4.Text)
        vectors(vectornum, 4) = Sqr(Val(Text3.Text) ^ 2 + Val(Text4.Text) ^ 2)
        vectors(vectornum, 5) = Atn(Val(Text4.Text) / Val(Text3.Text)) * 180 / pi

    End If
    vectors(vectornum, 1) = Text5.Text
    vectors(vectornum, 6) = Picture1.ForeColor
    drawvector (vectornum)
            If vectors(vectornum, 5) > 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If
        
        If vectors(vectornum, 5) < 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If

Else
    enabledraw = True
    Line1.Visible = True
End If
Frame1.Visible = False
End Sub

Private Sub Command10_Click()
Picture1.Cls
zm = zm * 2
createcoordinates (zm)
drawall

End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub

Private Sub Command3_Click()
If List2.ListIndex <> -1 Then
a = 0
b = 1
Do Until a = 1
If selected(b) = "" Then
a = 1
selected(b) = vectors(List2.ListIndex + 1, 1)
List3.AddItem (vectors(List2.ListIndex + 1, 1))
End If
b = b + 1
Loop
selnum = selnum + 1
End If
End Sub

Private Sub Command4_Click()
Picture1.Cls
zm = zm / 2
createcoordinates (zm)
drawall
End Sub

Private Sub Command5_Click()
For i = 1 To 200
selected(i) = ""
Next i
List3.clear
For i = 1 To vectornum
selected(i) = vectors(i, 1)
List3.AddItem (vectors(i, 1))
Next i
selnum = vectornum
End Sub

Private Sub Command6_Click()
For i = 1 To 200
selected(i) = ""
Next i
List3.clear
selnum = 0
End Sub

Private Sub Command7_Click()
On Error Resume Next
If List3.ListCount > 0 Then
Dim xx As Double, yy As Double
xx = 0
yy = 0
For i = 1 To selnum


a = 0
b = 1
Do Until b = 2
a = a + 1

If vectors(a, 1) = selected(i) Then
b = 2
End If
Loop

'Form1.Caption = Str(a)
xx = xx + vectors(a, 2)
yy = yy + vectors(a, 3)

Next i

Label9.Caption = Int(Sqr(xx ^ 2 + yy ^ 2) * 100) / 100
Tann = yy / xx
Label10.Caption = Atn(Tann) * 180 / pi
Label10.Caption = Int(Val(Label10.Caption) * 100) / 100

Text3.Text = xx
Text4.Text = yy
Option1.Value = False
Option2.Value = True
Option3.Value = False
Command7.Tag = Command7.Tag + 1
Text5.Text = "Displacement" + Command7.Tag
Picture1.ForeColor = vbRed
Command1_Click
If Val(Label10.Caption) > 0 And xx < 0 Then
Label10.Caption = Val(Label10.Caption) + 180
End If
        
If Val(Label10.Caption) < 0 And xx < 0 Then
Label10.Caption = Val(Label10.Caption) + 180
End If

Picture1.ForeColor = vbBlack
End If
End Sub

Private Sub Command8_Click()
If List2.ListIndex <> -1 Then
Label11.Caption = (vectors(List2.ListIndex + 1, 2))
Label12.Caption = (vectors(List2.ListIndex + 1, 3))
End If
End Sub

Private Sub Command9_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
enabledraw = False
vectornum = 0
selnum = 0
zm = 0.5
End Sub


Private Sub in_Click()
Command4_Click
End Sub

Private Sub newv_Click()
Frame1.Visible = True
Text1.Text = ""
Text2.Text = ""
Text5.Text = ""
End Sub

Private Sub oper_Click()
List2.clear
List3.clear
Frame2.Visible = True
For i = 1 To vectornum
List2.AddItem (vectors(i, 1))
Next i
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""

End Sub

Private Sub out_Click()
Command10_Click
End Sub

Private Sub Picture1_Click()
If enabledraw = True Then

vectors(vectornum, 2) = xpos
vectors(vectornum, 3) = ypos
vectors(vectornum, 4) = Sqr(Val(xpos) ^ 2 + Val(ypos) ^ 2)
vectors(vectornum, 5) = Atn(Val(ypos) / Val(xpos)) * 180 / pi
vectors(vectornum, 1) = Text5.Text
vectors(vectornum, 6) = Picture1.ForeColor
enabledraw = False
drawvector (vectornum)
Line1.Visible = False
End If
            If vectors(vectornum, 5) > 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If
        
        If vectors(vectornum, 5) < 0 And vectors(vectornum, 2) < 0 Then
        vectors(vectornum, 5) = vectors(vectornum, 5) + 180
        End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.Caption = "X= " + Str(Int(X * 100) / 100) + ",Y= " + Str(Int(Y * 100) / 100)
xpos = X
ypos = Y
If enabledraw = True Then
Line1.Visible = True
Line1.X1 = 0
Line1.Y1 = 0
Line1.X2 = X
Line1.Y2 = Y
End If
End Sub

Private Sub Timer1_Timer()
createcoordinates (1)
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label7(0).Caption = vectors(List1.ListIndex + 1, 1)
Label7(1).Caption = Str(Int(vectors(List1.ListIndex + 1, 2) * 100) / 100) + " , " + Str(Int(vectors(List1.ListIndex + 1, 3) * 100) / 100)
Label7(2).Caption = Str(Int(vectors(List1.ListIndex + 1, 4) * 100) / 100)
Label7(3).Caption = Str(Int(vectors(List1.ListIndex + 1, 5) * 100) / 100)
Picture2.BackColor = vectors(List1.ListIndex + 1, 6)
End Sub
