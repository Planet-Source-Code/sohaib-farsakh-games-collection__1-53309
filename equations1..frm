VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form11 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Equations"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   6720
      TabIndex        =   79
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3120
      TabIndex        =   78
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00AED8E6&
      Caption         =   "calculation bar>"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AED8E6&
      Caption         =   "Draw"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00AED8E6&
      Caption         =   "x=y+...."
      Height          =   375
      Left            =   1800
      TabIndex        =   76
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   75
      Text            =   "0"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00AED8E6&
      Caption         =   "find y"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00AED8E6&
      Caption         =   "slope="
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2640
      TabIndex        =   68
      Text            =   "0"
      Top             =   5040
      Width           =   855
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   6000
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   480
      TabIndex        =   64
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Zoom"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   8160
      TabIndex        =   55
      Top             =   2160
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1440
         TabIndex        =   63
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1440
         TabIndex        =   62
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H000080FF&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H000080FF&
         Caption         =   "OK"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "vertical zoom:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "horizontal zoom:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "zoom:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5000
      TabIndex        =   54
      Top             =   3360
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00AED8E6&
      Caption         =   "First"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00AED8E6&
      Caption         =   "Third"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00AED8E6&
      Caption         =   "Fourth"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00AED8E6&
      Caption         =   "Second"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00AED8E6&
      Caption         =   "Exponential"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00AED8E6&
      Caption         =   "TanX"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00AED8E6&
      Caption         =   "Abs"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00AED8E6&
      Caption         =   "Period"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AED8E6&
      Caption         =   "+"
      Height          =   500
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   250
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00AED8E6&
      Caption         =   "-"
      Height          =   500
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   250
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   720
      Picture         =   "equations1..frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   1320
      Picture         =   "equations1..frx":04BA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   720
      Picture         =   "equations1..frx":0974
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   500
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   120
      Picture         =   "equations1..frx":0E2E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   2160
      Picture         =   "equations1..frx":12E8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   2760
      Picture         =   "equations1..frx":17DA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   840
      Width           =   500
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   2160
      Picture         =   "equations1..frx":1CCC
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00AED8E6&
      Height          =   500
      Left            =   2760
      Picture         =   "equations1..frx":21BE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Width           =   500
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00AED8E6&
      Caption         =   "Drawing color"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00004000&
      Caption         =   "Show gridlines"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00AED8E6&
      Caption         =   "Clear all"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "equations1..frx":26B0
      Left            =   6240
      List            =   "equations1..frx":26BD
      TabIndex        =   25
      Tag             =   "0"
      Text            =   "Radians"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Option12"
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   8040
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   8040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Equation Solver"
      Height          =   3735
      Left            =   4560
      TabIndex        =   50
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command15 
         BackColor       =   &H000080FF&
         Caption         =   "Close"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H000080FF&
         Caption         =   "Solve"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   53
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FFFF&
         Caption         =   "x="
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FFFF&
         Caption         =   "Solve equation when y="
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H000F6200&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   73
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackColor       =   &H000F6200&
      Caption         =   "Find y when x="
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   72
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label22 
      BackColor       =   &H000F6200&
      Height          =   1335
      Left            =   240
      TabIndex        =   71
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label21 
      BackColor       =   &H000F6200&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   70
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H000F6200&
      Caption         =   "Find slope of equation when x="
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   67
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H000F6200&
      Height          =   1335
      Left            =   240
      TabIndex        =   66
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004000&
      Caption         =   "to:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00004000&
      Caption         =   "f(x)="
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   3120
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   4000
      X2              =   12000
      Y1              =   4000
      Y2              =   4000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   8000
      X2              =   8000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   13
      X1              =   7900
      X2              =   8100
      Y1              =   4000
      Y2              =   4000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   12
      X1              =   8000
      X2              =   8000
      Y1              =   3900
      Y2              =   4100
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   49
      Tag             =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   8040
      TabIndex        =   48
      Tag             =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   5000
      X2              =   5000
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
      Index           =   2
      X1              =   7000
      X2              =   7000
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
      Index           =   4
      X1              =   10000
      X2              =   10000
      Y1              =   3900
      Y2              =   4100
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
      Index           =   6
      X1              =   7900
      X2              =   8100
      Y1              =   1000
      Y2              =   1000
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
      Index           =   8
      X1              =   7900
      X2              =   8100
      Y1              =   3000
      Y2              =   3000
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
      Index           =   10
      X1              =   7900
      X2              =   8100
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   11
      X1              =   7900
      X2              =   8100
      Y1              =   7000
      Y2              =   7000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-15"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   47
      Tag             =   "-60"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-10"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   6040
      TabIndex        =   46
      Tag             =   "-40"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-5"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   7040
      TabIndex        =   45
      Tag             =   "-20"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   9040
      TabIndex        =   44
      Tag             =   "20"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   10040
      TabIndex        =   43
      Tag             =   "40"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   11040
      TabIndex        =   42
      Tag             =   "60"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   8040
      TabIndex        =   41
      Tag             =   "40"
      Top             =   1720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   8040
      TabIndex        =   40
      Tag             =   "60"
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   39
      Tag             =   "20"
      Top             =   2720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-10"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   8040
      TabIndex        =   38
      Tag             =   "-40"
      Top             =   5720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-5"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   8040
      TabIndex        =   37
      Tag             =   "-20"
      Top             =   4720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-15"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   8040
      TabIndex        =   36
      Tag             =   "-60"
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004000&
      Caption         =   "from:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004000&
      Caption         =   "pointer location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004000&
      Caption         =   "(0,0)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   30
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   4000
      X2              =   12000
      Y1              =   8000
      Y2              =   8000
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   5040
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   5000
      X2              =   5000
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
      Index           =   2
      X1              =   7000
      X2              =   7000
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
      Index           =   4
      X1              =   10000
      X2              =   10000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   5
      X1              =   11000
      X2              =   11000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   6
      X1              =   4000
      X2              =   12000
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line5 
      Index           =   7
      X1              =   4000
      X2              =   12000
      Y1              =   2000
      Y2              =   2000
   End
   Begin VB.Line Line5 
      Index           =   8
      X1              =   4000
      X2              =   12000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      Index           =   9
      X1              =   4005
      X2              =   12005
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Line Line5 
      Index           =   10
      X1              =   4005
      X2              =   12005
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line5 
      Index           =   11
      X1              =   4000
      X2              =   12000
      Y1              =   7000
      Y2              =   7000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004000&
      Height          =   8655
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   3975
   End
   Begin VB.Line Line5 
      Index           =   12
      Visible         =   0   'False
      X1              =   8000
      X2              =   8000
      Y1              =   0
      Y2              =   8000
   End
   Begin VB.Line Line5 
      Index           =   13
      X1              =   4005
      X2              =   12005
      Y1              =   4000
      Y2              =   4000
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const pi = 3.14159265358979
Option Base 1
Dim equnum As Integer
Dim equations(5000, 7) As String
Dim lastpress As String
Dim firsts As Boolean
Private Sub drawold()
Dim X As Double, Y As Double
Dim err As Boolean
Dim istan As Boolean, iscut As Boolean
On Error Resume Next

 Text11_Change
 
For i = 1 To equnum

istan = False
Check1.Value = equations(i, 2)
Check3.Value = equations(i, 3)
Text1.Text = equations(i, 4)
Text2.Text = equations(i, 5)
Form11.ForeColor = equations(i, 6)

n = 50000

k = Val((Label2(5).Caption) - Val(Label2(0).Caption)) / 6
If Check1.Value = 0 Then
Text1.Tag = Val(Label2(0).Caption) - k
Text2.Tag = Val(Label2(5).Caption) + k
Else
Text1.Tag = Val(Text1.Text)
Text2.Tag = Val(Text2.Text)
n = 50000 * (Val(Text2.Text) - Val(Text1.Text)) / ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) * 4 / 3)

End If
k = ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) / 30) ^ (1 / 3)
n = n * k


If Check3.Value = 0 Then
For h = 1 To Len(Text8.Text) - 3
If Mid(Text8.Text, h, 3) = "tan" Then istan = True
If Mid(Text8.Text, h, 3) = "mod" Or Mid(Text8.Text, h, 3) = "int" Or Mid(Text8.Text, h, 3) = "fix" Or Mid(Text8.Text, h, 3) = "log" Then iscut = True
Next h

For h = 1 To Len(Text8.Text) - 4
If Mid(Text8.Text, h, 4) = "/sin" Or Mid(Text8.Text, h, 4) = "/cos" Then istan = True
Next h

If istan = True Then n = n * 1.4
If n < 10000 Then n = 10000
If n > 250000 Then n = 250000

ProgressBar1.Visible = True
X = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi)
slope = 1
b = 4000
f = 4000


draw = True
For t = 1 To equnum - 1
If equations(equnum, 1) = equations(t, 1) And equations(t, 7) = True Then
draw = False
End If
Next t
If draw = True Then


Do Until X >= Val(Text2.Tag)
err = False
On Error GoTo error
ScriptControl1.ExecuteStatement ("x = " & X)
a = ScriptControl1.Eval(Trim(equations(i, 1)))


If iscut = True Then
factor1 = 0
factor2 = 0
xx = X
X = (xlast + xx) / 2
ScriptControl1.ExecuteStatement ("x = " & X)
aa = ScriptControl1.Eval(Trim(Text8.Text))
X = xx
aaa = (alast + a) / 2
If (aa >= aaa * 0.9999 And aa <= aaa * 1.0001) Or (aa <= aaa * 0.99 And aa >= aaa * 1.01) Then factor1 = 1
If (aa > alast And aa < a) Or (aa < alast And aa > a) Then factor2 = 1
If factor1 = 0 And factor2 = 0 Then err = True
End If


c = ((X * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + Val(Label6.Tag) * 1000 + 8000
Y = (Val(Label2(7).Caption) - Val(Label2(11).Caption))
e = 4000 - (((a * 6000 / Y))) - 1000 * Val(Label7.Tag)

If b >= 4000 And b <= 12000 And c >= 4000 And c <= 12000 And ((e >= 0 And e <= 8000) Or (f >= 0 And f <= 8000)) And err = False Then
If X = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi) Then
Form11.PSet (c, e)
Else
Form11.Line (b, f)-(c, e)
slope = (e - f) / (c - b)
End If
Else
If istan = False Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi))
End If
End If
slope = (e - f) / (c - b)

ProgressBar1.Value = (X - Text1.Tag) / (Text2.Tag - Text1.Tag) * 100
b = c
f = e
slope = Abs(slope)
xlast = X
alast = a
If slope < 1 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 30
End If
If slope >= 1 And slope < 25 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 10
End If
If slope >= 25 And slope < 100 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 4
End If
If slope >= 100 And slope < 1000 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) / 2
End If
If slope >= 1000 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) / 9
End If

Loop
equations(i, 7) = True
End If

ProgressBar1.Visible = False



Else


n = 50000
ProgressBar1.Visible = True
Y = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi)
slope = 1
b = 4000
f = 4000




Do Until Y >= Val(Text2.Tag)

ScriptControl1.ExecuteStatement ("y = " & Y)
a = ScriptControl1.Eval(Trim(equations(i, 1)))

c = ((a * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + Val(Label6.Tag) * 1000 + 8000
q = (Val(Label2(7).Caption) - Val(Label2(11).Caption))
e = 4000 - (((Y * 6000 / q))) - 1000 * Val(Label7.Tag)

If b >= 4000 And b <= 12000 And c >= 4000 And c <= 12000 And ((e >= 0 And e <= 8000) Or (f >= 0 And f <= 8000)) Then
If Y = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi) Then
Form11.PSet (c, e)
Else
Form11.Line (b, f)-(c, e)
slope = (c - b) / (e - f)
End If
Else
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 20
End If
slope = (c - b) / (e - f)

ProgressBar1.Value = (Y - Text1.Tag) / (Text2.Tag - Text1.Tag) * 100
b = c
f = e
slope = Abs(slope)
If slope < 1 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 30
End If
If slope >= 1 And slope < 25 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 10
End If
If slope >= 25 And slope < 100 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 4
End If
If slope >= 100 And slope < 1000 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) / 2
End If
If slope >= 1000 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) / 9
End If

Loop
ProgressBar1.Visible = False
End If

Next i
Combo1.Tag = 1
error:
err = True
Resume Next

End Sub
Private Sub Clearall()

Form11.Picture = LoadPicture()
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
Combo1.Tag = 0
End Sub

Private Sub Check1_Click()
If Check2.Value = 0 Then
For X = 0 To 11
Line5(X).Visible = False
Next X
Else
For X = 0 To 11
Line5(X).Visible = True
Next X
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
For X = 0 To 13
Line5(X).Visible = True
Next X
MDIForm1.sd.Checked = True
Else
For X = 0 To 13
Line5(X).Visible = False
Next X
MDIForm1.sd.Checked = False
End If
End Sub

Private Sub Command1_Click()
Dim X As Double, Y As Double
Dim istan As Boolean, iscut As Boolean
istan = False
Dim err As Boolean
On Error Resume Next
equnum = equnum + 1
If MDIForm1.clrchg.Checked = True Then
Form11.ForeColor = RGB(Rnd * 180, Rnd * 180, Rnd * 180)
End If

equations(equnum, 1) = Text8.Text
equations(equnum, 2) = Check1.Value
equations(equnum, 3) = Check3.Value
equations(equnum, 4) = Text1.Text
equations(equnum, 5) = Text2.Text
equations(equnum, 6) = Form11.ForeColor
iscut = False
For i = 1 To Len(Text8.Text) - 3
If Mid(Text8.Text, i, 3) = "tan" Then istan = True
If Mid(Text8.Text, i, 3) = "mod" Or Mid(Text8.Text, i, 3) = "int" Or Mid(Text8.Text, i, 3) = "fix" Or Mid(Text8.Text, i, 3) = "log" Then iscut = True
Next i

For i = 1 To Len(Text8.Text) - 4
If Mid(Text8.Text, i, 4) = "/sin" Or Mid(Text8.Text, i, 4) = "/cos" Then istan = True
Next i

k = Val((Label2(5).Caption) - Val(Label2(0).Caption)) / 6
If Check1.Value = 0 Then
Text1.Tag = Val(Label2(0).Caption) - k
Text2.Tag = Val(Label2(5).Caption) + k
n = 50000
Else
Text1.Tag = Val(Text1.Text)
Text2.Tag = Val(Text2.Text)
n = 50000 * (Val(Text2.Text) - Val(Text1.Text)) / ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) * 4 / 3)
End If
If istan = True Then n = n * 1.4
k = ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) / 30) ^ (1 / 3)
n = n * k
draw = True
For t = 1 To equnum - 1
If equations(equnum, 1) = equations(t, 1) And equations(t, 7) = True Then
draw = False
End If
Next t
If draw = True Then

If Check3.Value = 0 Then
ProgressBar1.Visible = True
slope = 1
b = 4000
f = 4000
If n < 10000 Then n = 10000
If n > 250000 Then n = 250000
X = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi)


Do Until X >= Val(Text2.Tag)
err = False
On Error GoTo error
ScriptControl1.ExecuteStatement ("x = " & X)
a = ScriptControl1.Eval(Trim(Text8.Text))

If iscut = True Then
factor1 = 0
factor2 = 0
xx = X
X = (xlast + xx) / 2
ScriptControl1.ExecuteStatement ("x = " & X)
aa = ScriptControl1.Eval(Trim(Text8.Text))
X = xx
aaa = (alast + a) / 2
If (aa >= aaa * 0.9999 And aa <= aaa * 1.0001) Or (aa <= aaa * 0.99 And aa >= aaa * 1.01) Then factor1 = 1
If (aa > alast And aa < a) Or (aa < alast And aa > a) Then factor2 = 1
If factor1 = 0 And factor2 = 0 Then err = True
End If

c = ((X * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + Val(Label6.Tag) * 1000 + 8000
Y = (Val(Label2(7).Caption) - Val(Label2(11).Caption))
e = 4000 - (((a * 6000 / Y))) - 1000 * Val(Label7.Tag)

If b >= 4000 And b <= 12000 And c >= 4000 And c <= 12000 And ((e >= 0 And e <= 8000) Or (f >= 0 And f <= 8000)) And err = False Then
If X = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi) Then
Form11.PSet (c, e)
Else
Form11.Line (b, f)-(c, e)
slope = (e - f) / (c - b)
End If
Else
If istan = False Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi))
End If
End If
slope = (e - f) / (c - b)

ProgressBar1.Value = (X - Text1.Tag) / (Text2.Tag - Text1.Tag) * 100
b = c
f = e
slope = Abs(slope)
xlast = X
alast = a
If slope < 1 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 30
End If
If slope >= 1 And slope < 25 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 10
End If
If slope >= 25 And slope < 100 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) * 4
End If
If slope >= 100 And slope < 1000 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) / 2
End If
If slope >= 1000 Then
X = X + (((Text2.Tag - Text1.Tag) / n * pi)) / 9
End If

Loop
equations(equnum, 7) = True
ProgressBar1.Visible = False


Else


ProgressBar1.Visible = True
Y = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi)
slope = 1
b = 4000
f = 4000




Do Until Y >= Val(Text2.Tag)
On Error GoTo error
ScriptControl1.ExecuteStatement ("y = " & Y)
a = ScriptControl1.Eval(Trim(Text8.Text))

c = ((a * 6000 / (Val(Label2(5).Caption) - Val(Label2(0).Caption)))) + Val(Label6.Tag) * 1000 + 8000
q = (Val(Label2(7).Caption) - Val(Label2(11).Caption))
e = 4000 - (((Y * 6000 / q))) - 1000 * Val(Label7.Tag)

If b >= 4000 And b <= 12000 And c >= 4000 And c <= 12000 And ((e >= 0 And e <= 8000) Or (f >= 0 And f <= 8000)) And err = False Then
If Y = Text1.Tag + ((Text2.Tag - Text1.Tag) / n * pi) Then
Form11.PSet (c, e)
Else
Form11.Line (b, f)-(c, e)
slope = (c - b) / (e - f)
End If
Else
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 20
End If
slope = (c - b) / (e - f)

ProgressBar1.Value = (Y - Text1.Tag) / (Text2.Tag - Text1.Tag) * 100
b = c
f = e
slope = Abs(slope)
If slope < 1 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 30
End If
If slope >= 1 And slope < 25 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 10
End If
If slope >= 25 And slope < 100 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) * 4
End If
If slope >= 100 And slope < 1000 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) / 2
End If
If slope >= 1000 Then
Y = Y + (((Text2.Tag - Text1.Tag) / n * pi)) / 9
End If

Loop
ProgressBar1.Visible = False
End If
Combo1.Tag = 1
error:
err = True
Resume Next
End If

End Sub

Private Sub Command10_Click()
For X = 6 To 11
Label2(X).Caption = Val(Label2(X).Caption) / 4
Next X
Label2(13).Caption = Val(Label2(13).Caption) / 4

If Combo1.Tag = 1 Then
Form11.cls
drawold
End If

End Sub

Private Sub Command11_Click()
For X = 6 To 11
Label2(X).Caption = Val(Label2(X).Caption) * 4
Next X
Label2(13).Caption = Val(Label2(13).Caption) * 4

If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub


Private Sub Command12_Click()
CommonDialog1.ShowColor
Form11.ForeColor = CommonDialog1.Color

End Sub

Private Sub Command13_Click()
Clearall
Form11.Picture = LoadPicture()
Unload Form11
Load Form11
Form11.Show
lastpress = ""
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim a, b, c As Double
Select Case Label1.Caption
Case "1"
a = Val(Text3.Text) - Val(Command6.Tag)
Label12.Caption = a / Val(Command5.Tag)
Case "2"
c = Val(Command6.Tag) - Val(Text3.Text)
b = Command5.Tag
a = Command4.Tag
e = (b ^ 2 - 4 * a * c)
d = (-b + Sqr(e)) / (2 * a)
f = (-b - Sqr(e)) / (2 * a)
d = d * 10000
d = Int(d)
d = d / 10000
f = f * 10000
f = Int(f)
f = f / 10000

If e = 0 Then
Label12.Caption = Str$(d)
Else
Label12.Caption = Str$(d) + " , " + Str$(f)
End If
Case "6"
a = (Log(Val(Text3.Text)) / Log(10)) / (Log(Val(Label9.Caption)) / Log(10))
Label12.Caption = (a - Val(Command9.Tag)) / Command8.Tag
Case "3"
Dim isfound As Integer
a1 = Command3.Tag
a2 = Command4.Tag
a3 = Command5.Tag
a4 = Command6.Tag - Val(Text3.Text)
If a2 = 0 And a3 = 0 And a4 = 0 Then Label12.Caption = "0,0,0"
If a3 <> 0 And a4 = 0 Then
r1 = (-a2 + Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r2 = (-a2 - Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r1 = Int(r1 * 100) / 100
r2 = Int(r2 * 100) / 100
Label12.Caption = "0," + Str(r1) + "," + Str(r2)
End If
If a2 <> 0 And a3 = 0 And a4 = 0 Then
rt = -a2 / a1
Label12.Caption = "0,0," + Str(rt)
End If
If a4 <> 0 Then
Root = 0
isfound = 0
test = 0
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 1
Loop
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 1
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 0.0002
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 0.0002
Loop
End If
b1 = a1
c1 = Root * b1
b2 = a2 + c1
c2 = Root * b2
b3 = c2 + a3
delta = b2 ^ 2 - (4 * b1 * b3)
Root = Int(Root * 1000) / 1000
If delta >= 0 Then
r1 = (-b2 + Sqr(delta)) / (2 * b1)
r2 = (-b2 - Sqr(delta)) / (2 * b1)
r1 = Int(r1 * 1000) / 1000
r2 = Int(r2 * 1000) / 1000
Label12.Caption = Str(Root) + "," + Str(r1) + "," + Str(r2)
Else
Label12.Caption = Str(Root)
End If
End If
Case "9"
a1 = Option1.Tag
a2 = Option2.Tag
a3 = Option3.Tag
a4 = Command6.Tag - Val(Text3.Text)
lirst = -a4 / a1
arc = Atn(lirst)
sol = (arc - a3) / a2
Select Case Combo1.Text
Case "Degrees"
Label12.Caption = sol * 180 / pi
Case "Radians"
Label12.Caption = Str$(sol)
Case "Gradians"
Label12.Caption = sol * 200 / pi
End Select
Case "10"
a1 = Command10.Tag
a2 = Command11.Tag
a3 = Command12.Tag
a4 = Command6.Tag - Val(Text3.Text)
a5 = -a4 / a1
a1 = Command11.Tag
a2 = Command12.Tag
a3 = a5
sol = (a3 - a2) / a1
Label12.Caption = Str(sol)
a1 = Command11.Tag
a2 = Command12.Tag
a3 = -a5
sol = (a3 - a2) / a1
Label12.Caption = Label12.Caption + "," + Str(sol)



Case "4"
a1 = Command2.Tag
a2 = Command3.Tag
a3 = Command4.Tag
a4 = Command5.Tag
a5 = Command6.Tag - Val(Text3.Text)
If a2 = 0 And a3 = 0 And a4 = 0 And a5 = 0 Then Label12.Caption = "0,0,0,0"
If a3 <> 0 And a4 = 0 And a5 = 0 Then
r1 = (-a2 + Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r2 = (-a2 - Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r1 = Int(r1 * 1000) / 1000
r2 = Int(r2 * 1000) / 1000
Label12.Caption = "0,0," + Str(r1) + "," + Str(r2)
End If
If a2 <> 0 And a3 = 0 And a4 = 0 And a5 = 0 Then
rt = -a2 / a1
Label12.Caption = "0,0,0," + Str(rt)
End If




If a4 <> 0 And a5 = 0 Then
If a2 = 0 And a3 = 0 And a4 = 0 Then Label12.Caption = "0,0,0"
If a3 <> 0 And a4 = 0 Then
r1 = (-a2 + Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r2 = (-a2 - Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r1 = Int(r1 * 100) / 100
r2 = Int(r2 * 100) / 100
Label12.Caption = "0," + Str(r1) + "," + Str(r2)
End If
If a2 <> 0 And a3 = 0 And a4 = 0 Then
rt = -a2 / a1
Label12.Caption = "0,0," + Str(rt)
End If
If a4 <> 0 Then
Root = 0
isfound = 0
test = 0
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 1
Loop
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 1
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 0.0002
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 0.0002
Loop
End If
b1 = a1
c1 = Root * b1
b2 = a2 + c1
c2 = Root * b2
b3 = c2 + a3
delta = b2 ^ 2 - (4 * b1 * b3)
Root = Int(Root * 1000) / 1000
If delta >= 0 Then
r1 = (-b2 + Sqr(delta)) / (2 * b1)
r2 = (-b2 - Sqr(delta)) / (2 * b1)
r1 = Int(r1 * 1000) / 1000
r2 = Int(r2 * 1000) / 1000
Label12.Caption = "0 , " + Str(Root) + "," + Str(r1) + "," + Str(r2)
Else
Label12.Caption = "0 , " + Str(Root)
End If
End If

End If







If a5 <> 0 Then
Root = 0
isfound = 0
test = 0
Do Until isfound = 1
result = a1 * test ^ 4 + a2 * test ^ 3 + a3 * test ^ 2 + a4 * test + a5
If result = 0 Then
isfound = 1
Root = test
End If
If test >= Abs(a5) Then isfound = 1
test = test + 1
Loop
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 4 + a2 * test ^ 3 + a3 * test ^ 2 + a4 * test + a5
If result = 0 Then
isfound = 1
Root = test
End If
If test <= -Abs(a5) Then isfound = 1
test = test - 1
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 4 + a2 * test ^ 3 + a3 * test ^ 2 + a4 * test + a5
If result > -0.001 And result < 0.001 Then
isfound = 1
Root = test
End If
If test >= Abs(a5) Then isfound = 1
test = test + 0.00002
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 4 + a2 * test ^ 3 + a3 * test ^ 2 + a4 * test + a5
If result > -0.001 And result < 0.001 Then
isfound = 1
Root = test
End If
If test <= -Abs(a5) Then isfound = 1
test = test - 0.00002
Loop
End If
Root = Int(Root * 1000) / 1000
Label12.Caption = Str(Root)
b1 = a1
c1 = Root * b1
b2 = a2 + c1
c2 = Root * b2
b3 = a3 + c2
c3 = Root * b3
b4 = c3 + a4
a1 = b1
a2 = b2
a3 = b3
a4 = b4
If a2 = 0 And a3 = 0 And a4 = 0 Then Label12.Caption = Label12.Caption + "," + "0,0,0"
If a3 <> 0 And a4 = 0 Then
r1 = (-a2 + Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r2 = (-a2 - Sqr(a2 ^ 2 - 4 * a1 * a3)) / (2 * a1)
r1 = Int(r1 * 100) / 100
r2 = Int(r2 * 100) / 100
Label12.Caption = Label12.Caption + "," + "0," + Str(r1) + "," + Str(r2)
End If
If a2 <> 0 And a3 = 0 And a4 = 0 Then
rt = -a2 / a1
Label12.Caption = Label12.Caption + "," + "0,0," + Str(rt)
End If
If a4 <> 0 Then
Root = 0
isfound = 0
test = 0
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 1
Loop
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result = 0 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 1
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test >= Abs(a4) Then isfound = 1
test = test + 0.0002
Loop
End If
test = 0
isfound = 0

If Root = 0 Then
Do Until isfound = 1
result = a1 * test ^ 3 + a2 * test ^ 2 + a3 * test + a4
If result > -0.01 And result < 0.01 Then
isfound = 1
Root = test
End If
If test <= -Abs(a4) Then isfound = 1
test = test - 0.0002
Loop
End If
b1 = a1
c1 = Root * b1
b2 = a2 + c1
c2 = Root * b2
b3 = c2 + a3
delta = b2 ^ 2 - (4 * b1 * b3)
Root = Int(Root * 1000) / 1000
If delta >= 0 Then
r1 = (-b2 + Sqr(delta)) / (2 * b1)
r2 = (-b2 - Sqr(delta)) / (2 * b1)
r1 = Int(r1 * 1000) / 1000
r2 = Int(r2 * 1000) / 1000
Label12.Caption = Label12.Caption + "," + Str(Root) + "," + Str(r1) + "," + Str(r2)
Else
Label12.Caption = Label12.Caption + "," + Str(Root)
End If
End If
End If




End Select

End Sub

Private Sub Command15_Click()
Frame1.Visible = False
Form11.Option1.Visible = False
Form11.Option2.Visible = False
Form11.Option3.Visible = False
Form11.Option4.Visible = False
Form11.Option5.Visible = False
Form11.Option6.Visible = False
Form11.Option7.Visible = False
Combo1.Visible = False
End Sub

Private Sub Command16_Click()
If Command16.Caption = "calculation bar>" Then
Text10.Visible = True
Text10.Text = ""
Command16.Caption = "calculation bar<"
Else
Command16.Caption = "calculation bar>"
Text10.Visible = False
End If
End Sub

Private Sub Command17_Click()
On Error Resume Next
X = Text4.Text
ScriptControl1.ExecuteStatement ("x = " & X)
a = ScriptControl1.Eval(Trim(Text8.Text))
a = a * 1000000000000#
t = Int(a)
c = a - t
If c >= 0.5 Then
t = t + 1
End If
t = t / 1000000000000#
a = t
Label13.Caption = Str$(a)
End Sub

Private Sub Command18_Click()
If Text6.Text = "0" Or Text6.Text = "" Or Text7.Text = "0" Or Text7.Text = "" Then
d = MsgBox("Can't set zoom to zero", vbOKOnly & vbCritical, "Error")
Else
On Error Resume Next
For X = 0 To 5
Label2(X).Caption = 100 / Val(Text6.Text) * Label2(X).Tag
Next X
Label2(12).Caption = 100 / Val(Text6.Text) * Label2(12).Tag
For X = 6 To 11
Label2(X).Caption = 100 / Val(Text7.Text) * Label2(X).Tag
Next X
Label2(13).Caption = 100 / Val(Text7.Text) * Label2(13).Tag
Frame3.Visible = False
End If
If Combo1.Tag = 1 Then
Form11.cls
drawold
End If
End Sub

Private Sub Command19_Click()
Frame3.Visible = False
End Sub

Private Sub Command2_Click()
For X = 0 To 13
Label2(X).Caption = Val(Label2(X).Caption) / 4
Next X
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Command20_Click()
On Error Resume Next
X = Val(Text9.Text)
ScriptControl1.ExecuteStatement ("x = " & X)
g = ScriptControl1.Eval(Trim(Text8.Text))

X = Val(Text9.Text) + 0.0000000001
ScriptControl1.ExecuteStatement ("x = " & X)
h = ScriptControl1.Eval(Trim(Text8.Text))

X = Val(Text9.Text) - 0.0000000001
ScriptControl1.ExecuteStatement ("x = " & X)
v = ScriptControl1.Eval(Trim(Text8.Text))

X = Val(Text9.Text)
Y = Val(Text9.Text) + 0.0000000001
z = Val(Text9.Text) - 0.0000000001

slope1 = (h - g) / (Y - X)
slope2 = (g - v) / (X - z)

slope12 = (slope1 + slope2) / 2
a = slope12
a = a * 10
t = Int(a)
c = a - t
If c >= 0.5 Then
t = t + 1
End If
t = t / 10
a = t

Label21.Caption = Str(a)
End Sub

Private Sub Command3_Click()
For X = 0 To 13
Label2(X).Caption = Val(Label2(X).Caption) * 4
Next X
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Command4_Click()
a = ((Val(Label2(7).Caption) - Val(Label2(11).Caption)) / 3)
For X = 6 To 11
Label2(X).Caption = Val(Label2(X).Caption) - a
Next X
Label2(13).Caption = Val(Label2(13).Caption) - a

For X = 6 To 11
Label2(X).Tag = Val(Label2(X).Tag) - 40
Next X
Label2(13).Tag = Val(Label2(13).Tag) - 40

Line2.Y1 = Line2.Y1 - 2000
Line2.Y2 = Line2.Y2 - 2000
For X = 0 To 5
Label2(X).Top = Label2(X).Top - 2000
Line3(X).Y1 = Line3(X).Y1 - 2000
Line3(X).Y2 = Line3(X).Y2 - 2000
Next X
Label2(12).Top = Label2(12).Top - 2000
Line3(12).Y1 = Line3(12).Y1 - 2000
Line3(12).Y2 = Line3(12).Y2 - 2000
Label7.Tag = Val(Label7.Tag) + 2
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Command5_Click()
a = ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) / 3)
For X = 0 To 5
Label2(X).Caption = Val(Label2(X).Caption) - a
Next X
Label2(12) = Val(Label2(12).Caption) - a

For X = 0 To 5
Label2(X).Tag = Val(Label2(X).Tag) - 40
Next X
Label2(12).Tag = Val(Label2(12).Tag) - 40

Line1.X1 = Line1.X1 + 2000
Line1.X2 = Line1.X2 + 2000
For X = 6 To 11
Label2(X).Left = Label2(X).Left + 2000
Line3(X).X1 = Line3(X).X1 + 2000
Line3(X).X2 = Line3(X).X2 + 2000
Next X
Label2(13).Left = Label2(13).Left + 2000
Line3(13).X1 = Line3(13).X1 + 2000
Line3(13).X2 = Line3(13).X2 + 2000
Label6.Tag = Val(Label6.Tag) + 2
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If


End Sub

Private Sub Command6_Click()
a = ((Val(Label2(7).Caption) - Val(Label2(11).Caption)) / 3)
For X = 6 To 11
Label2(X).Caption = Val(Label2(X).Caption) + a
Next X

For X = 6 To 11
Label2(X).Tag = Val(Label2(X).Tag) + 40
Next X
Label2(13).Tag = Val(Label2(13).Tag) + 40

Label2(13).Caption = Val(Label2(13).Caption) + a
Line2.Y1 = Line2.Y1 + 2000
Line2.Y2 = Line2.Y2 + 2000
For X = 0 To 5
Label2(X).Top = Label2(X).Top + 2000
Line3(X).Y1 = Line3(X).Y1 + 2000
Line3(X).Y2 = Line3(X).Y2 + 2000
Next X
Label2(12).Top = Label2(12).Top + 2000
Line3(12).Y1 = Line3(12).Y1 + 2000
Line3(12).Y2 = Line3(12).Y2 + 2000
Label7.Tag = Val(Label7.Tag) - 2
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Command7_Click()
a = ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) / 3)
For X = 0 To 5
Label2(X).Caption = Val(Label2(X).Caption) + a
Next X
Label2(12) = Val(Label2(12).Caption) + a

For X = 0 To 5
Label2(X).Tag = Val(Label2(X).Tag) + 40
Next X
Label2(12).Tag = Val(Label2(12).Tag) + 40

Line1.X1 = Line1.X1 - 2000
Line1.X2 = Line1.X2 - 2000
For X = 6 To 11
Label2(X).Left = Label2(X).Left - 2000
Line3(X).X1 = Line3(X).X1 - 2000
Line3(X).X2 = Line3(X).X2 - 2000
Next X
Label2(13).Left = Label2(13).Left - 2000
Line3(13).X1 = Line3(13).X1 - 2000
Line3(13).X2 = Line3(13).X2 - 2000
Label6.Tag = Val(Label6.Tag) - 2
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If
End Sub

Private Sub Command8_Click()
For X = 0 To 5
Label2(X).Caption = Val(Label2(X).Caption) / 4
Next X
Label2(12).Caption = Val(Label2(12).Caption) / 4
If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Command9_Click()
For X = 0 To 5
Label2(X).Caption = Val(Label2(X).Caption) * 4
Next X
Label2(12).Caption = Val(Label2(12).Caption) * 4

If Combo1.Tag = 1 Then
Form11.cls

drawold
End If

End Sub

Private Sub Form_Click()
On Error Resume Next
If X >= 4000 And X <= 12000 And Y >= 0 And Y <= 8000 Then
t = X - 4000
b = Val(Label2(5).Caption) - Val(Label2(0).Caption)
t = t * b / 6000
t = t + (Val((Label2(5).Caption) + Val(Label2(0).Caption)) / 2)
t = t - ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) * 2 / 3)
t = t * 100
t = Int(t)
t = t / 100
v = 8000 - Y
c = Val(Label2(7).Caption) - Val(Label2(11).Caption)
v = v * c / 6000
v = v + (Val((Label2(7).Caption) + Val(Label2(11).Caption)) / 2)
v = v - ((Val(Label2(7).Caption) - Val(Label2(11).Caption)) * 2 / 3)
v = v * 100
v = Int(v)
v = v / 100

Label7.Caption = "( " + Str$(t) + " , " + Str$(v) + " )"
End If
End Sub

Private Sub Form_Load()
equnum = 0
For X = 0 To 13
Label2(X).Tag = Label2(X).Caption
Line5(X).Visible = False
Line5(X).BorderStyle = 3
Next X
Clearall
firsts = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If X >= 4000 And X <= 12000 And Y >= 0 And Y <= 8000 Then
t = X - 4000
b = Val(Label2(5).Caption) - Val(Label2(0).Caption)
t = t * b / 6000
t = t + (Val((Label2(5).Caption) + Val(Label2(0).Caption)) / 2)
t = t - ((Val(Label2(5).Caption) - Val(Label2(0).Caption)) * 2 / 3)
t = t * 100
t = Int(t)
t = t / 100
v = 8000 - Y
c = Val(Label2(7).Caption) - Val(Label2(11).Caption)
v = v * c / 6000
v = v + (Val((Label2(7).Caption) + Val(Label2(11).Caption)) / 2)
v = v - ((Val(Label2(7).Caption) - Val(Label2(11).Caption)) * 2 / 3)
v = v * 100
v = Int(v)
v = v / 100

Label7.Caption = "( " + Str$(t) + " , " + Str$(v) + " )"
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form11
Form11.Show
MDIForm1.Timer1.Enabled = True
End Sub

Private Sub Option1_Click()
Label1.Caption = "1"
d = InputBox("y=ax+b, enter(a)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax+b, enter(b)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Option1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option1.Value = True Then

Option1_Click
Option1.Value = True
End If
End Sub






Private Sub Option2_Click()
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

Private Sub Option2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option2.Value = True Then

Option2_Click
Option2.Value = True
End If
End Sub

Private Sub Option3_Click()
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

Private Sub Option3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option3.Value = True Then

Option3_Click
Option3.Value = True
End If
End Sub

Private Sub Option4_Click()
Label1.Caption = "2"
d = InputBox("y=ax^2 + bx + c, enter(a)", "entering data", "1")
Command4.Tag = d
d = InputBox("y=ax^2 + bx + c, enter(b)", "entering data", "1")
Command5.Tag = d
d = InputBox("y=ax^2 + bx + c, enter(c)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Option4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option4.Value = True Then

Option4_Click
Option4.Value = True
End If
End Sub

Private Sub Option5_Click()
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

Private Sub Option5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option5.Value = True Then

Option5_Click
Option5.Value = True
End If
End Sub

Private Sub Option6_Click()
Label1.Caption = "6"
d = InputBox("y=a^(bx + c), enter(a)", "entering data", "1")
Label9.Caption = d
d = InputBox("y=a^(bx + c), enter(b)", "entering data", "1")
Command8.Tag = d
d = InputBox("y=a^(bx + c), enter(c)", "entering data", "0")
Command9.Tag = d

End Sub

Private Sub Option6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option6.Value = True Then
Option6_Click
Option6.Value = True
End If
End Sub

Private Sub Option7_Click()
Label1.Caption = "9"
d = InputBox("y=a*Tan(bx+c)+d, enter(a)", "entering data", "1")
Option1.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(b)", "entering data", "1")
Option2.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(c)", "entering data", "0")
Option3.Tag = d
d = InputBox("y=a*Tan(bx+c)+d, enter(d)", "entering data", "0")
Command6.Tag = d

End Sub

Private Sub Option7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Option7.Value = True Then
Option7_Click
Option7.Value = True
End If
End Sub



Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text10.Text = ScriptControl1.Eval(Text10.Text)
End If
End Sub

Private Sub Text11_Change()
For i = 1 To equnum
equations(i, 7) = False
Next i
End Sub

Private Sub Text5_Change()
Text6.Text = Text5.Text
Text7.Text = Text5.Text
End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
'KeyPreview = True
'Select Case KeyAscii
'Case 48 To 57
'lastpress = "number"
'Case 88, 120
'If firsts = True Then
'KeyAscii = 42
'Text8_KeyPress (88)
'Else
'KeyAscii = 88
'End If
'a = Mid(Text8.Text, Len(Text8.Text), 1)
'Select Case a
'Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
'Text8.Text = Text8.Text + "*x"
'KeyAscii = 0
''End Select
'lastpress = ""
'Case Else
'lastpress = ""
'End Select

End Sub

Private Sub Timer1_Timer()
k = Val((Label2(5).Caption) - Val(Label2(0).Caption)) / 6
If Check1.Value = 0 Then
Text1.Tag = Val(Label2(0).Caption) - k
Text2.Tag = Val(Label2(5).Caption) + k
Else
Text1.Tag = Val(Text1.Text)
Text2.Tag = Val(Text2.Text)
End If
Select Case Check3.Value
Case 0
Label19.Caption = "f(x)="
Case 1
Label19.Caption = "f(y)="
End Select
End Sub
