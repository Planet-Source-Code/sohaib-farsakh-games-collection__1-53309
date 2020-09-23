VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "equations"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   8040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Option12"
      Height          =   255
      Left            =   4080
      TabIndex        =   48
      Top             =   8040
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "equations.frx":0000
      Left            =   1080
      List            =   "equations.frx":000D
      TabIndex        =   46
      Text            =   "Radians"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Clear all"
      Height          =   495
      Left            =   2160
      TabIndex        =   44
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00004000&
      Caption         =   "Show dividers"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   41
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Drawing color"
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Height          =   500
      Left            =   2760
      Picture         =   "equations.frx":002D
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command10 
      Height          =   500
      Left            =   2160
      Picture         =   "equations.frx":051F
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command9 
      Height          =   500
      Left            =   2760
      Picture         =   "equations.frx":0A11
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command8 
      Height          =   500
      Left            =   2160
      Picture         =   "equations.frx":0F03
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command7 
      Height          =   500
      Left            =   120
      Picture         =   "equations.frx":13F5
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command6 
      Height          =   500
      Left            =   720
      Picture         =   "equations.frx":18AF
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2040
      Width           =   500
   End
   Begin VB.CommandButton Command5 
      Height          =   500
      Left            =   1320
      Picture         =   "equations.frx":1D69
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command4 
      Height          =   500
      Left            =   720
      Picture         =   "equations.frx":2223
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   500
      Left            =   720
      TabIndex        =   31
      Top             =   1440
      Width           =   250
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   500
      Left            =   960
      TabIndex        =   30
      Top             =   1440
      Width           =   250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   27
      Top             =   7200
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Period"
      Height          =   375
      Left            =   1680
      TabIndex        =   24
      Top             =   6600
      Width           =   1335
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Abs"
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   5400
      Width           =   1335
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Complex"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option9 
      Caption         =   "SinX"
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton Option8 
      Caption         =   "CosX"
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
   End
   Begin VB.OptionButton Option7 
      Caption         =   "TanX"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Exponential"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Fifth"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Second"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Fourth"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Third"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "First"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Line Line5 
      Index           =   11
      X1              =   4000
      X2              =   12000
      Y1              =   7000
      Y2              =   7000
   End
   Begin VB.Line Line5 
      Index           =   10
      X1              =   4000
      X2              =   12000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line5 
      Index           =   9
      X1              =   4000
      X2              =   12000
      Y1              =   5000
      Y2              =   5000
   End
   Begin VB.Line Line5 
      Index           =   8
      X1              =   4000
      X2              =   12000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      Index           =   7
      X1              =   4000
      X2              =   12000
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line5 
      Index           =   6
      X1              =   4000
      X2              =   12000
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line5 
      Index           =   5
      X1              =   11000
      X2              =   11000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   4
      X1              =   10000
      X2              =   10000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   3
      X1              =   9000
      X2              =   9000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   2
      X1              =   7000
      X2              =   7000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   6000
      X2              =   6000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   5000
      X2              =   5000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   5040
      TabIndex        =   49
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4560
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line4 
      X1              =   4000
      X2              =   12000
      Y1              =   8000
      Y2              =   8000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004000&
      Caption         =   "Angle unit:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004000&
      Caption         =   "(0,0)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004000&
      Caption         =   "pointer location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004000&
      Caption         =   "to:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   26
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004000&
      Caption         =   "from:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004000&
      Height          =   8655
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-60"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   8040
      TabIndex        =   11
      Tag             =   "-60"
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-20"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   8040
      TabIndex        =   10
      Tag             =   "-20"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-40"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   8040
      TabIndex        =   9
      Tag             =   "-40"
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   8
      Tag             =   "20"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   8040
      TabIndex        =   7
      Tag             =   "60"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   8040
      TabIndex        =   6
      Tag             =   "40"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   10920
      TabIndex        =   5
      Tag             =   "60"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   4
      Tag             =   "40"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   3
      Tag             =   "20"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-20"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   2
      Tag             =   "-20"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-40"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   1
      Tag             =   "-40"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-60"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Tag             =   "-60"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   11
      X1              =   7900
      X2              =   8100
      Y1              =   7000
      Y2              =   7000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   10
      X1              =   7900
      X2              =   8100
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   9
      X1              =   7900
      X2              =   8100
      Y1              =   5000
      Y2              =   5000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   8
      X1              =   7900
      X2              =   8100
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   7
      X1              =   7900
      X2              =   8100
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   6
      X1              =   7900
      X2              =   8100
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   5
      X1              =   11000
      X2              =   11000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   4
      X1              =   10000
      X2              =   10000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   9000
      X2              =   9000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   7000
      X2              =   7000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   6000
      X2              =   6000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   5000
      X2              =   5000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4000
      X2              =   12000
      Y1              =   4000
      Y2              =   4000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   8000
      X2              =   8000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu zoomto 
         Caption         =   "Zoom to"
         Begin VB.Menu a 
            Caption         =   "25%"
         End
         Begin VB.Menu b 
            Caption         =   "50%"
         End
         Begin VB.Menu c 
            Caption         =   "100%"
         End
         Begin VB.Menu d 
            Caption         =   "200%"
         End
         Begin VB.Menu e 
            Caption         =   "400%"
         End
         Begin VB.Menu line 
            Caption         =   "-"
         End
         Begin VB.Menu custom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu move 
         Caption         =   "Move"
         Begin VB.Menu left 
            Caption         =   "Left"
         End
         Begin VB.Menu right 
            Caption         =   "Right"
         End
         Begin VB.Menu up 
            Caption         =   "Up"
         End
         Begin VB.Menu down 
            Caption         =   "Down"
         End
         Begin VB.Menu liine 
            Caption         =   "-"
         End
         Begin VB.Menu go 
            Caption         =   "Go to point"
         End
      End
   End
   Begin VB.Menu equation 
      Caption         =   "Equation"
      Begin VB.Menu mul 
         Caption         =   "Multi Parts"
         Begin VB.Menu first 
            Caption         =   "First"
         End
         Begin VB.Menu second 
            Caption         =   "Second"
         End
         Begin VB.Menu third 
            Caption         =   "Third"
         End
         Begin VB.Menu fourth 
            Caption         =   "Fourth"
         End
         Begin VB.Menu fifth 
            Caption         =   "Fifth"
         End
      End
      Begin VB.Menu exp 
         Caption         =   "Exponential"
      End
      Begin VB.Menu sin 
         Caption         =   "SinX"
      End
      Begin VB.Menu cos 
         Caption         =   "CosX"
      End
      Begin VB.Menu tan 
         Caption         =   "TanX"
      End
      Begin VB.Menu abs 
         Caption         =   "Absolute Value"
      End
      Begin VB.Menu linee 
         Caption         =   "-"
      End
      Begin VB.Menu cplx 
         Caption         =   "Complex Equation Editor"
      End
      Begin VB.Menu period 
         Caption         =   "Period"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu drawing 
         Caption         =   "Drawing color..."
      End
      Begin VB.Menu pointer 
         Caption         =   "Show Pointer Location"
      End
      Begin VB.Menu show 
         Caption         =   "Show Dividers"
      End
      Begin VB.Menu angle 
         Caption         =   "Angle Unit"
         Begin VB.Menu deg 
            Caption         =   "Degrees"
         End
         Begin VB.Menu rad 
            Caption         =   "Radians"
         End
         Begin VB.Menu grad 
            Caption         =   "Gradians"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Clearall()
Command1.Tag = "0"
Command2.Tag = "0"
Command3.Tag = "0"
Command4.Tag = "0"
Command5.Tag = "0"
Command6.Tag = "0"
Command8.Tag = "0"
Command9.Tag = "0"
Command10.Tag = "0"
Command11.Tag = "0"
Command12.Tag = "0"
Label9.Caption = "0"
Option1.Tag = "0"
Option2.Tag = "0"
Option3.Tag = "0"
Option4.Tag = "0"
Option5.Tag = "0"
Option6.Tag = "0"
Option7.Tag = "0"
Option8.Tag = "0"
Option9.Tag = "0"
Option10.Tag = "0"
Option11.Tag = "0"

End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
For i = 0 To 11
Line5(i).Visible = False
Next i
Else
For i = 0 To 11
Line5(i).Visible = True
Next i
End If
End Sub

Private Sub Command1_Click()
For i = Text1.Tag To Text2.Tag Step (Text2.Tag - Text1.Tag) / 4000
j = i - ((Text2.Tag - Text1.Tag) / 4000)
a = (Val(Command1.Tag) * (i ^ 5)) + (Val(Command2.Tag) * (i ^ 4)) + (Val(Command3.Tag) * (i ^ 3)) + (Val(Command4.Tag) * (i ^ 2)) + (Val(Command5.Tag) * (i ^ 1)) + Val(Command6.Tag) + Val(Command10.Tag) * Abs(Val(Command11.Tag) * i + Val(Command12.Tag)) + Val(Option1.Tag) * sin(Val(Option2.Tag) * i + Val(Option3.Tag)) + Val(Option4.Tag) * cos(Val(Option5.Tag) * i + Val(Option6.Tag)) + Val(Option7.Tag) * tan(Val(Option8.Tag) * i + Val(Option9.Tag))
t = (Val(Command1.Tag) * (j ^ 5)) + (Val(Command2.Tag) * (j ^ 4)) + (Val(Command3.Tag) * (j ^ 3)) + (Val(Command4.Tag) * (j ^ 2)) + (Val(Command5.Tag) * (j ^ 1)) + Val(Command6.Tag) + Val(Command10.Tag) * Abs(Val(Command11.Tag) * j + Val(Command12.Tag)) + Val(Option1.Tag) * sin(Val(Option2.Tag) * j + Val(Option3.Tag)) + Val(Option4.Tag) * cos(Val(Option5.Tag) * j + Val(Option6.Tag)) + Val(Option7.Tag) * tan(Val(Option8.Tag) * j + Val(Option9.Tag))

If Label9.Caption <> 0 Then
a = a + Val(Label9.Caption) ^ (Val(Command8.Tag) * i + Val(Command9.Tag))
t = t + Val(Label9.Caption) ^ (Val(Command8.Tag) * j + Val(Command9.Tag))
End If
b = ((j * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + 8000
c = ((i * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + 8000
y = (Val(Label2(7).Caption) - Val(Label2(11).Caption))
e = 4000 - (((a * 6000 / y)))
f = 4000 - (((t * 6000 / y)))
If b > 4000 And b < 12000 And f > 0 And f < 8000 And c > 4000 And c < 12000 And e > 0 And e < 8000 Then
Form1.Line (b, f)-(c, e)
End If
Next i

End Sub

Private Sub Command10_Click()
For i = 6 To 11
Label2(i).Caption = Val(Label2(i).Caption) / 4
Next i
End Sub

Private Sub Command11_Click()
For i = 6 To 11
Label2(i).Caption = Val(Label2(i).Caption) * 4
Next i
End Sub

Private Sub Command12_Click()
CommonDialog1.ShowColor
Form1.ForeColor = CommonDialog1.Color
End Sub

Private Sub Command13_Click()
Clearall
Form1.Picture = LoadPicture()
End Sub

Private Sub Command2_Click()
For i = 0 To 11
Label2(i).Caption = Val(Label2(i).Caption) / 4
Next i
End Sub

Private Sub Command3_Click()
For i = 0 To 11
Label2(i).Caption = Val(Label2(i).Caption) * 4
Next i
End Sub

Private Sub Command8_Click()
For i = 0 To 5
Label2(i).Caption = Val(Label2(i).Caption) / 4
Next i
End Sub

Private Sub Command9_Click()
For i = 0 To 5
Label2(i).Caption = Val(Label2(i).Caption) * 4
Next i
End Sub

Private Sub Form_Load()
For i = 0 To 11
Line5(i).Visible = False
Next i
End Sub

Private Sub Option1_Click()
Clearall
Label1.Caption = "1"
d = InputBox("y=ax+b, enter(a)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax+b, enter(b)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option10_Click()
Label1.Caption = "11"
End Sub

Private Sub Option11_Click()
Clearall
Label1.Caption = "10"
d = InputBox("y=a*Abs(bx+c)+d, enter(a)", "entering data", "1")
Command10.Tag = d
d = InputBox("y=a*Abs(bx+c)+d, enter(b)", "entering data", "1")
Command11.Tag = d
d = InputBox("y=a*Abs(bx+c)+d, enter(c)", "entering data", "0")
Command12.Tag = d
d = InputBox("y=a*Abs(bx+c)+d, enter(d)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option2_Click()
Clearall
Label1.Caption = "3"
d = InputBox("y=ax^3 + bx^2 + cx + d, enter(a)", "entering data", "1")
Command3.Tag = d
d = InputBox("y=ax^3 + bx^2 + cx + d, enter(b)", "entering data", "1")
Command4.Tag = d
d = InputBox("y=ax^3 + bx^2 + cx + d, enter(c)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax^3 + bx^2 + cx + d, enter(d)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option3_Click()
Clearall
Label1.Caption = "4"
d = InputBox("y=ax^4 + bx^3 + cx^2 + dx + e, enter(a)", "entering data", "1")
Command2.Tag = d
d = InputBox("y=ax^4 + bx^3 + cx^2 + dx + e, enter(b)", "entering data", "1")
Command3.Tag = d
d = InputBox("y=ax^4 + bx^3 + cx^2 + dx + e, enter(c)", "entering data", "1")
Command4.Tag = d
d = InputBox("y=ax^4 + bx^3 + cx^2 + dx + e, enter(d)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax^4 + bx^3 + cx^2 + dx + e, enter(e)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option4_Click()
Clearall
Label1.Caption = "2"
d = InputBox("y=ax^2 + bx + c, enter(a)", "entering data", "1")
Command4.Tag = d
d = InputBox("y=ax^2 + bx + c, enter(b)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax^2 + bx + c, enter(c)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option5_Click()
Clearall
Label1.Caption = "5"
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(a)", "entering data", "1")
Command1.Tag = d
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(b)", "entering data", "1")
Command2.Tag = d
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(c)", "entering data", "1")
Command3.Tag = d
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(d)", "entering data", "1")
Command4.Tag = d
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(e)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax^5 + bx^4 + cx^3 + dx^2 + ex + f, enter(f)", "entering data", "0")
Command6.Tag = d
End Sub

Private Sub Option6_Click()
Clearall
Label1.Caption = "6"
d = InputBox("y=a^(bx + c), enter(a)", "entering data", "1")
Label9.Caption = d
d = InputBox("y=a^(bx + c), enter(b)", "entering data", "1")
Command8.Tag = d
d = InputBox("y=a^(bx + c), enter(c)", "entering data", "0")
Command9.Tag = d
End Sub

Private Sub Option7_Click()
Clearall
Label1.Caption = "9"
d = InputBox("y=a*Tan(bx+c)+d, enter(a)", "entering data", "1")
Option7.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(b)", "entering data", "1")
Option8.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(c)", "entering data", "0")
Option9.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(d)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Option8_Click()
Clearall
Label1.Caption = "8"
d = InputBox("y=a*Cos(bx+c)+d, enter(a)", "entering data", "1")
Option4.Tag = d
d = InputBox("y=a*Cos(bx+c)+d, enter(b)", "entering data", "1")
Option5.Tag = d
d = InputBox("y=a*Cos(bx+c)+d, enter(c)", "entering data", "0")
Option6.Tag = d
d = InputBox("y=a*Cos(bx+c)+d, enter(d)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Option9_Click()
Clearall
Label1.Caption = "7"
d = InputBox("y=a*Sin(bx+c)+d, enter(a)", "entering data", "1")
Option1.Tag = d
d = InputBox("y=a*Sin(bx+c)+d, enter(b)", "entering data", "1")
Option2.Tag = d
d = InputBox("y=a*Sin(bx+c)+d, enter(c)", "entering data", "0")
Option3.Tag = d
d = InputBox("y=a*Sin(bx+c)+d, enter(d)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Timer1_Timer()
If Check1.Value = 0 Then
Text1.Tag = Val(Label2(0).Caption) * 4 / 3
Text2.Tag = Val(Label2(5).Caption) * 4 / 3
Else
Text1.Tag = Val(Text1.Text)
Text2.Tag = Val(Text2.Text)
End If
End Sub
