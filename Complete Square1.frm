VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00007F00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complete Square V.2.3"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   1920
      Top             =   6240
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   720
      Top             =   6240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      Caption         =   "Players"
      Height          =   2415
      Left            =   7440
      TabIndex        =   117
      Top             =   3720
      Width           =   1695
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   138
         Top             =   1180
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   137
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Difficult"
         Height          =   255
         Left            =   960
         TabIndex        =   135
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   960
         TabIndex        =   134
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Easy"
         Height          =   255
         Left            =   1080
         TabIndex        =   133
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "V.Easy"
         Height          =   255
         Left            =   960
         TabIndex        =   132
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         Height          =   255
         Left            =   960
         TabIndex        =   131
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "V. Easy"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "next turn:"
         Height          =   255
         Left            =   240
         TabIndex        =   123
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         Height          =   255
         Left            =   -120
         TabIndex        =   121
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Difficult"
         Height          =   255
         Left            =   0
         TabIndex        =   120
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   0
         TabIndex        =   119
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Easy"
         Height          =   255
         Left            =   0
         TabIndex        =   118
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   40
         TabIndex        =   122
         Top             =   460
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Score"
      Height          =   2535
      Left            =   7440
      TabIndex        =   112
      Top             =   960
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   255
         Left            =   240
         TabIndex        =   129
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   128
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "wins by red"
         Height          =   255
         Left            =   240
         TabIndex        =   127
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   126
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "wins by blue"
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   116
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   310
         TabIndex        =   115
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   114
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   55
      Left            =   7080
      TabIndex        =   111
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   54
      Left            =   6120
      TabIndex        =   110
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   53
      Left            =   5160
      TabIndex        =   109
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   52
      Left            =   4200
      TabIndex        =   108
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   51
      Left            =   3240
      TabIndex        =   107
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   50
      Left            =   2280
      TabIndex        =   106
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   49
      Left            =   1320
      TabIndex        =   105
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   48
      Left            =   360
      TabIndex        =   104
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   47
      Left            =   7080
      TabIndex        =   103
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   46
      Left            =   6120
      TabIndex        =   102
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   45
      Left            =   5160
      TabIndex        =   101
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   44
      Left            =   4200
      TabIndex        =   100
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   43
      Left            =   3240
      TabIndex        =   99
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   42
      Left            =   2280
      TabIndex        =   98
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   41
      Left            =   1320
      TabIndex        =   97
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   40
      Left            =   360
      TabIndex        =   96
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   39
      Left            =   7080
      TabIndex        =   95
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   38
      Left            =   6120
      TabIndex        =   94
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   37
      Left            =   5160
      TabIndex        =   93
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   36
      Left            =   4200
      TabIndex        =   92
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   35
      Left            =   3240
      TabIndex        =   91
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   34
      Left            =   2280
      TabIndex        =   90
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   33
      Left            =   1320
      TabIndex        =   89
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   32
      Left            =   360
      TabIndex        =   88
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   31
      Left            =   7080
      TabIndex        =   87
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   30
      Left            =   6120
      TabIndex        =   86
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   29
      Left            =   5160
      TabIndex        =   85
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   28
      Left            =   4200
      TabIndex        =   84
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   27
      Left            =   3240
      TabIndex        =   83
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   26
      Left            =   2280
      TabIndex        =   82
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   25
      Left            =   1320
      TabIndex        =   81
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   24
      Left            =   360
      TabIndex        =   80
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   23
      Left            =   7080
      TabIndex        =   79
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   22
      Left            =   6120
      TabIndex        =   78
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   21
      Left            =   5160
      TabIndex        =   77
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   20
      Left            =   4200
      TabIndex        =   76
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   19
      Left            =   3240
      TabIndex        =   75
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   18
      Left            =   2280
      TabIndex        =   74
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   17
      Left            =   1320
      TabIndex        =   73
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   16
      Left            =   360
      TabIndex        =   72
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   15
      Left            =   7080
      TabIndex        =   71
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   14
      Left            =   6120
      TabIndex        =   70
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   13
      Left            =   5160
      TabIndex        =   69
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   12
      Left            =   4200
      TabIndex        =   68
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   11
      Left            =   3240
      TabIndex        =   67
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   10
      Left            =   2280
      TabIndex        =   66
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   9
      Left            =   1320
      TabIndex        =   65
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   64
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   7
      Left            =   7080
      TabIndex        =   63
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   6
      Left            =   6120
      TabIndex        =   62
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   5
      Left            =   5160
      TabIndex        =   61
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   4
      Left            =   4200
      TabIndex        =   60
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   3
      Left            =   3240
      TabIndex        =   59
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   2
      Left            =   2280
      TabIndex        =   58
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   57
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   56
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   55
      Left            =   6240
      TabIndex        =   55
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   54
      Left            =   5280
      TabIndex        =   54
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   53
      Left            =   4320
      TabIndex        =   53
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   52
      Left            =   3360
      TabIndex        =   52
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   51
      Left            =   2400
      TabIndex        =   51
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   50
      Left            =   1440
      TabIndex        =   50
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   49
      Left            =   480
      TabIndex        =   49
      Top             =   6000
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   48
      Left            =   6240
      TabIndex        =   48
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   47
      Left            =   5280
      TabIndex        =   47
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   46
      Left            =   4320
      TabIndex        =   46
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   45
      Left            =   3360
      TabIndex        =   45
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   44
      Left            =   2400
      TabIndex        =   44
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   43
      Left            =   1440
      TabIndex        =   43
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   42
      Left            =   480
      TabIndex        =   42
      Top             =   5280
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   41
      Left            =   6240
      TabIndex        =   41
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   40
      Left            =   5280
      TabIndex        =   40
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   39
      Left            =   4320
      TabIndex        =   39
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   38
      Left            =   3360
      TabIndex        =   38
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   37
      Left            =   2400
      TabIndex        =   37
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   36
      Left            =   1440
      TabIndex        =   36
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   35
      Left            =   480
      TabIndex        =   35
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   34
      Left            =   6240
      TabIndex        =   34
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   33
      Left            =   5280
      TabIndex        =   33
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   32
      Left            =   4320
      TabIndex        =   32
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   31
      Left            =   3360
      TabIndex        =   31
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   30
      Left            =   2400
      TabIndex        =   30
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   29
      Left            =   1440
      TabIndex        =   29
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   28
      Left            =   480
      TabIndex        =   28
      Top             =   3840
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   27
      Left            =   6240
      TabIndex        =   27
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   26
      Left            =   5280
      TabIndex        =   26
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   25
      Left            =   4320
      TabIndex        =   25
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   24
      Left            =   3360
      TabIndex        =   24
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   23
      Left            =   2400
      TabIndex        =   23
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   22
      Left            =   1440
      TabIndex        =   22
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   21
      Left            =   480
      TabIndex        =   21
      Top             =   3120
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   20
      Left            =   6240
      TabIndex        =   20
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   19
      Left            =   5280
      TabIndex        =   19
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   18
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   17
      Left            =   3360
      TabIndex        =   17
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   16
      Left            =   2400
      TabIndex        =   16
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   15
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   14
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   13
      Left            =   6240
      TabIndex        =   13
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   12
      Left            =   5280
      TabIndex        =   12
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   11
      Left            =   4320
      TabIndex        =   11
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   10
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   9
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   6
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   5
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   105
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu pauseresume 
         Caption         =   "Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu suggest 
         Caption         =   "Suggest Move"
         Shortcut        =   ^M
      End
      Begin VB.Menu undo 
         Caption         =   "Undo Move"
         Shortcut        =   ^Z
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu blue 
         Caption         =   "Blue player"
         Begin VB.Menu human1 
            Caption         =   "Human"
            Checked         =   -1  'True
         End
         Begin VB.Menu very1 
            Caption         =   "Very Easy"
         End
         Begin VB.Menu easy1 
            Caption         =   "Easy"
         End
         Begin VB.Menu normal1 
            Caption         =   "Normal"
         End
         Begin VB.Menu difficult1 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu red 
         Caption         =   "Red player"
         Begin VB.Menu human2 
            Caption         =   "Human"
         End
         Begin VB.Menu very2 
            Caption         =   "Very Easy"
         End
         Begin VB.Menu easy2 
            Caption         =   "Easy"
         End
         Begin VB.Menu normal2 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu difficult2 
            Caption         =   "Difficult"
         End
      End
      Begin VB.Menu separateline 
         Caption         =   "-"
      End
      Begin VB.Menu speed 
         Caption         =   "Computer Playing Speed"
         Begin VB.Menu slow 
            Caption         =   "Slow"
         End
         Begin VB.Menu normal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu fast 
            Caption         =   "Fast"
         End
      End
      Begin VB.Menu seperateline2 
         Caption         =   "-"
      End
      Begin VB.Menu gridsize 
         Caption         =   "Grid Size"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim turn As String
Dim grid(20, 20) As Integer
Dim grid2(20, 20) As Integer
Dim recentwin As Boolean
Dim winned As Integer
Dim lbl2(450) As Boolean
Dim lbl3(450) As Boolean
Dim enabled1 As Boolean, enabled2 As Boolean
Dim lastblue As Integer, lastred As Integer
Dim firstbluemove As Boolean, firstredmove As Boolean
Dim allmoves(1000, 2) As Variant
Dim movenum As Integer
Dim removedwins As Integer
Dim boardheight As Integer, boardwidth As Integer
Dim verlinenum As Integer, horlinenum As Integer
Dim loaded1 As Integer, loaded2 As Integer
Dim horlinewidth, verlineheight


Private Sub ordinarylast(color As String)
Select Case color
Case "blue"
If lastblue < (horlinenum + 1) Then
Label2(lastblue).BackColor = vbBlue
Else
Label3(lastblue - (horlinenum + 1)).BackColor = vbBlue
End If
Case "red"
If lastred < (horlinenum + 1) Then
Label2(lastred).BackColor = vbRed
Else
Label3(lastred - (horlinenum + 1)).BackColor = vbRed
End If
End Select
End Sub
Sub Delay(secs)
Dim Start
Start = Timer
While (Timer < (Start + secs))
DoEvents
Wend
End Sub

Private Sub saveenabled()
For i = 0 To horlinenum
lbl2(i) = Label2(i).Enabled
Label2(i).Enabled = False
Next i
For i = 1 To verlinenum
lbl3(i) = Label3(i).Enabled
Label3(i).Enabled = False
Next i
End Sub
Private Sub revertenabled()
For i = 0 To horlinenum
Label2(i).Enabled = lbl2(i)
Next i
For i = 1 To verlinenum
Label3(i).Enabled = lbl3(i)
Next i
End Sub

Private Sub endgame()
If Val(Label5.Caption) > Val(Label7.Caption) Then
MsgBox ("The blue player wins the game!")
Label16.Caption = Val(Label16.Caption) + 1
End If
If Val(Label5.Caption) < Val(Label7.Caption) Then
MsgBox ("The red player wins the game!")
Label18.Caption = Val(Label18.Caption) + 1
End If
If Val(Label5.Caption) = Val(Label7.Caption) Then
MsgBox ("The game is a draw")
End If
End Sub
Private Sub unmakemove(linenum As Integer)
removedwins = 0
Dim indexx As Integer
If linenum <= horlinenum Then
indexx = linenum
If Label2(indexx).Enabled = True Then Exit Sub
Label2(indexx).BackColor = &HC000&
Label2(indexx).Enabled = True
a = Int(Label2(indexx).index / boardwidth) + 1
b = (Label2(indexx).index Mod boardwidth) + 1
If a > 1 Then
If grid(a - 1, b) = 5 Then
grid(a - 1, b) = 4
removedwins = removedwins + 1
End If
grid(a - 1, b) = grid(a - 1, b) - 1
End If
If a < (boardheight + 1) Then
If grid(a, b) = 5 Then
grid(a, b) = 4
removedwins = removedwins + 1
End If
grid(a, b) = grid(a, b) - 1
End If



Else


indexx = linenum - (horlinenum + 1)
If Label3(indexx).Enabled = True Then Exit Sub
Label3(indexx).BackColor = &HC000&
Label3(indexx).Enabled = True

a = Int(Label3(indexx).index / (boardwidth + 1)) + 1
b = (Label3(indexx).index Mod (boardwidth + 1)) + 1
If b > 1 Then
If grid(a, b - 1) = 5 Then
grid(a, b - 1) = 4
removedwins = removedwins + 1
End If
grid(a, b - 1) = grid(a, b - 1) - 1
End If
If b < (boardwidth + 1) Then
If grid(a, b) = 5 Then
grid(a, b) = 4
removedwins = removedwins + 1
End If
grid(a, b) = grid(a, b) - 1
End If


End If

End Sub
Private Sub undomove()
On Error GoTo err
Dim player As String, opponent As String
If (turn = "blue" And human1.Checked = True) Or (turn = "red" And human2.Checked = True) Then
player = "human"
Else
player = "computer"
End If

If (turn = "red" And human1.Checked = True) Or (turn = "blue" And human2.Checked = True) Then
opponent = "human"
Else
opponent = "computer"
End If

If player = "human" And opponent = "human" Then
unmakemove (allmoves(movenum, 0))
If allmoves(movenum, 2) = True Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
If removedwins = 2 Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
End If
If allmoves(movenum, 1) = "blue" Then
Label5.Caption = Val(Label5.Caption) - removedwins
Else
Label7.Caption = Val(Label7.Caption) - removedwins
End If
End If
If turn <> allmoves(movenum, 1) Then
changeturn
End If
movenum = movenum - 1
lastblue = 0
lastred = 0
For i = movenum To 0 Step -1
If allmoves(i, 1) = "blue" And lastblue = 0 Then lastblue = allmoves(i, 0)
If allmoves(i, 1) = "red" And lastred = 0 Then lastred = allmoves(i, 0)
Next i
If lastblue = 0 Then firstbluemove = True
If lastred = 0 Then firstredmove = True

End If
If player = "human" And opponent = "computer" Then
changes = 0
saveturn = turn
Do
If allmoves(movenum, 1) <> saveturn Then
changes = changes + 1
saveturn = allmoves(movenum, 1)
End If
unmakemove (allmoves(movenum, 0))
If allmoves(movenum, 2) = True Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
If removedwins = 2 Then
Shape1(winned).Visible = False
Unload Shape1(winned)
winned = winned - 1
End If

If allmoves(movenum, 1) = "blue" Then
Label5.Caption = Val(Label5.Caption) - removedwins
Else
Label7.Caption = Val(Label7.Caption) - removedwins
End If
End If
movenum = movenum - 1
Loop While changes < 2
lastblue = 0
lastred = 0
For i = movenum To 0 Step -1
If allmoves(i, 1) = "blue" And lastblue = 0 Then lastblue = allmoves(i, 0)
If allmoves(i, 1) = "red" And lastred = 0 Then lastred = allmoves(i, 0)
Next i
If lastblue = 0 Then firstbluemove = True
If lastred = 0 Then firstredmove = True

End If
err:
Exit Sub
End Sub
Private Sub takeline(linenum As Integer)
Dim indexx As Integer
If linenum <= horlinenum Then
indexx = linenum
If Label2(indexx).Enabled = False Then Exit Sub
recentwin = False
firstmove = False
Select Case turn
Case "blue"
Label2(indexx).BackColor = &HFF8080
If firstbluemove = False Then ordinarylast ("blue")
lastblue = indexx
firstbluemove = False

Case "red"
Label2(indexx).BackColor = &H8080FF
If firstredmove = False Then ordinarylast ("red")
lastred = indexx
firstredmove = False
End Select
Label2(indexx).Enabled = False
a = Int(Label2(indexx).index / boardwidth) + 1
b = (Label2(indexx).index Mod boardwidth) + 1
If a > 1 Then
grid(a - 1, b) = grid(a - 1, b) + 1
End If
If a < (boardheight + 1) Then
grid(a, b) = grid(a, b) + 1
End If



Else


indexx = linenum - (horlinenum + 1)
recentwin = False
firstmove = False
If Label3(indexx).Enabled = False Then Exit Sub

Select Case turn
Case "blue"
Label3(indexx).BackColor = &HFF8080
If firstbluemove = False Then ordinarylast ("blue")
lastblue = indexx + (horlinenum + 1)
firstbluemove = False
Case "red"
Label3(indexx).BackColor = &H8080FF
If firstredmove = False Then ordinarylast ("red")
lastred = indexx + (horlinenum + 1)
firstredmove = False
End Select
Label3(indexx).Enabled = False
a = Int(Label3(indexx).index / (boardwidth + 1)) + 1
b = (Label3(indexx).index Mod (boardwidth + 1)) + 1
If b > 1 Then
grid(a, b - 1) = grid(a, b - 1) + 1
End If
If b < (boardwidth + 1) Then
grid(a, b) = grid(a, b) + 1
End If


End If
checkwin
movenum = movenum + 1
allmoves(movenum, 0) = linenum
allmoves(movenum, 1) = turn
allmoves(movenum, 2) = recentwin
If recentwin = False Then
changeturn
End If

End Sub
Private Sub resizeboard()

For i = horlinenum To 1 Step -1
Label2(i).Visible = False
Next i

For i = verlinenum To 1 Step -1
Label3(i).Visible = False
Next i

horlinenum = boardwidth * (boardheight + 1) - 1
verlinenum = (boardwidth + 1) * boardheight - 1

horlinewidth = (6825 - (105 * (boardwidth + 1))) / boardwidth
verlineheight = (5145 - (105 * (boardheight + 1))) / boardheight
For i = 0 To horlinenum
If i > loaded1 Then
Load Label2(i)
End If
Label2(i).Visible = True
rownum = Int(i / boardwidth)
colnum = i Mod boardwidth

Label2(i).Top = 960 + (rownum * (verlineheight + 105))
Label2(i).Height = 105
Label2(i).Width = horlinewidth
Label2(i).Left = 360 + (colnum * horlinewidth) + (105 * (colnum + 1))
Next i

For i = 0 To verlinenum
If i > loaded2 Then
Load Label3(i)
End If
Label3(i).Visible = True
rownum = Int(i / (boardwidth + 1))
colnum = i Mod (boardwidth + 1)

Label3(i).Top = 960 + (rownum * (verlineheight + 105)) + 105
Label3(i).Height = verlineheight
Label3(i).Width = 105
Label3(i).Left = 360 + (horlinewidth + 105) * colnum
Next i


If horlinenum > loaded1 Then
loaded1 = horlinenum
End If

If verlinenum > loaded2 Then
loaded2 = verlinenum
End If
End Sub
Private Sub play(player As String, suggest As Boolean)
Dim diff As Integer
If suggest = False Then
Select Case player
Case "blue"
If very1.Checked = True Then diff = 1
If easy1.Checked = True Then diff = 2
If normal1.Checked = True Then diff = 3
If difficult1.Checked = True Then diff = 4
Case "red"
If very2.Checked = True Then diff = 1
If easy2.Checked = True Then diff = 2
If normal2.Checked = True Then diff = 3
If difficult2.Checked = True Then diff = 4
End Select
Else
Select Case turn
Case "red"
If very1.Checked = True Then diff = 1
If easy1.Checked = True Then diff = 2
If normal1.Checked = True Then diff = 3
If difficult1.Checked = True Then diff = 4
Case "blue"
If very2.Checked = True Then diff = 1
If easy2.Checked = True Then diff = 2
If normal2.Checked = True Then diff = 3
If difficult2.Checked = True Then diff = 4
End Select

End If
Dim plays(900) As Integer
Dim element As Integer
element = 0

If diff > 1 Then
For i = 1 To boardwidth
For j = 1 To boardheight
If grid(i, j) = 3 Then
a = (i - 1) * boardwidth + j - 1
b = a + boardwidth
c = (i - 1) * (boardwidth + 1) + j - 1
d = c + 1
If Label2(a).Enabled = True Then
element = element + 1
plays(element) = a
End If
If Label2(b).Enabled = True Then
element = element + 1
plays(element) = b
End If
If Label3(c).Enabled = True Then
element = element + 1
plays(element) = c + (horlinenum + 1)
End If
If Label3(d).Enabled = True Then
element = element + 1
plays(element) = d + (horlinenum + 1)
End If
End If
Next j
Next i
End If
If diff > 2 Then
If element = 0 Then
For i = 0 To horlinenum
If Label2(i).Enabled = True Then
a1 = 0
a2 = 0
a = Int(i / boardwidth) + 1
b = (i Mod boardwidth) + 1
If a > 1 Then
If grid(a - 1, b) < 2 Then a1 = 1
Else
a1 = 1
End If
If a < (boardheight + 1) Then
If grid(a, b) < 2 Then a2 = 1
Else
a2 = 1
End If
If a1 = 1 And a2 = 1 And Label2(i).Enabled = True Then
element = element + 1
plays(element) = i
End If
End If
Next i




For i = 0 To verlinenum
If Label3(i).Enabled = True Then

a1 = 0
a2 = 0
a = Int(i / (boardwidth + 1)) + 1
b = (i Mod (boardwidth + 1)) + 1
If b > 1 Then
If grid(a, b - 1) < 2 Then a1 = 1
Else
a1 = 1
End If
If b < (boardwidth + 1) Then
If grid(a, b) < 2 Then a2 = 1
Else
a2 = 1
End If
If a1 = 1 And a2 = 1 Then
If Label3(i).Enabled = True Then
element = element + 1
plays(element) = i + (horlinenum + 1)
End If
End If
End If
Next i

End If
End If




If diff = 4 And element = 0 Then

leastnum = 500

For z = 0 To horlinenum

For i = 1 To boardwidth
For j = 1 To boardheight
grid2(i, j) = grid(i, j)
Next j
Next i


If Label2(z).Enabled = True Then
Label2(z).Enabled = False
u = Int(z / boardwidth) + 1
v = (z Mod boardwidth) + 1
If u > 1 Then
grid2(u - 1, v) = grid2(u - 1, v) + 1
End If
If u < (boardheight + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardwidth Step 1
If found = True Then Exit For
For j = 1 To boardheight Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardwidth + j - 1
b = a + boardwidth
c = (i - 1) * (boardwidth + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True And i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
End If
If Label2(b).Enabled = True And i < boardheight Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
End If
If Label3(c).Enabled = True And j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
End If
If Label3(d).Enabled = True And j < boardwidth Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop



If numberr < leastnum Then
leastnum = numberr
chosen = z
End If
Label2(z).Enabled = True

End If
Next z







For z = 0 To verlinenum

For i = 1 To boardwidth
For j = 1 To boardheight
grid2(i, j) = grid(i, j)
Next j
Next i


If Label3(z).Enabled = True Then
Label3(z).Enabled = False
u = Int(z / (boardwidth + 1)) + 1
v = (z Mod (boardwidth + 1)) + 1
If v > 1 Then
grid2(u, v - 1) = grid2(u, v - 1) + 1
End If
If v < (boardwidth + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardwidth Step 1
If found = True Then Exit For
For j = 1 To boardheight Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardwidth + j - 1
b = a + boardwidth
c = (i - 1) * (boardwidth + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True And i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
End If
If Label2(b).Enabled = True And i < boardheight Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
End If
If Label3(c).Enabled = True And j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
End If
If Label3(d).Enabled = True And j < boardwidth Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop



If numberr < leastnum Then
leastnum = numberr
chosen = z + (horlinenum + 1)
End If
Label3(z).Enabled = True

End If
Next z

element = 1
plays(1) = chosen
End If





If element = 0 Then
For i = 0 To horlinenum
If Label2(i).Enabled = True Then
element = element + 1
plays(element) = i
End If
Next i
For i = 0 To verlinenum
If Label3(i).Enabled = True Then
element = element + 1
plays(element) = i + (horlinenum + 1)
End If
Next i
End If

If element <> 0 Then
d = Int(Rnd * element) + 1
e = plays(d)
End If
If suggest = False Then
takeline (e)
Else

If e < (horlinenum + 1) Then
Label2(e).BackColor = vbbrown
Delay (0.5)
Label2(e).BackColor = &HC000&
Else
Label3(e - (horlinenum + 1)).BackColor = vbbrown
Delay (0.5)
Label3(e - (horlinenum + 1)).BackColor = &HC000&
End If

End If
End Sub
Private Sub checkwin()
For i = 1 To boardwidth
    For j = 1 To boardheight
        If grid(i, j) = 4 Then
        recentwin = False
        grid(i, j) = 5
        Select Case turn
        Case "blue"
        Label5.Caption = Val(Label5.Caption) + 1
        Case "red"
        Label7.Caption = Val(Label7.Caption) + 1
        End Select
        recentwin = True
        winned = winned + 1
        Load Shape1(winned)
        Shape1(winned).Visible = True
        Select Case turn
        Case "blue"
        Shape1(winned).FillColor = vbBlue
        Case "red"
        Shape1(winned).FillColor = vbRed
        End Select
        Shape1(winned).Width = horlinewidth * 0.42
        Shape1(winned).Height = horlinewidth * 0.42
        centerx = 420 + (horlinewidth + 105) * (j - 0.5)
        centery = 1040 + (verlineheight + 105) * (i - 0.5)
        Shape1(winned).Top = centery - 0.5 * Shape1(winned).Width
        Shape1(winned).Left = centerx - 0.5 * Shape1(winned).Height
        If Val(Label5.Caption) + Val(Label7.Caption) = (boardwidth * boardheight) Then
        endgame
        End If
        End If
    Next j
Next i

End Sub
Private Sub newgame()
For i = 1 To winned
Shape1(i).Visible = False
Unload Shape1(i)
Next i

turn = "blue"
recentwin = False
winned = 0
For i = 1 To boardwidth
    For j = 1 To boardheight
        grid(i, j) = 0
    Next j
Next i
For i = 0 To horlinenum
Label2(i).BackColor = &HC000&
Label2(i).Enabled = True
lbl2(i) = True
Next i
For i = 0 To verlinenum
Label3(i).BackColor = &HC000&
Label3(i).Enabled = True
lbl3(i) = True
Next i
Label5.Caption = "0"
Label7.Caption = "0"
Label14.Caption = "Blue"
Label14.ForeColor = vbBlue
Timer1.Enabled = True
Timer2.Enabled = True
enabled1 = True
enabled2 = True
firstbluemove = True
firstredmove = True
movenum = 0
End Sub
Private Sub changeturn()
If turn = "blue" Then
turn = "red"
Label14.Caption = "Red"
Label14.ForeColor = vbRed
Timer1.Enabled = False
Timer2.Enabled = True
Else
turn = "blue"
Label14.Caption = "Blue"
Label14.ForeColor = vbBlue
Timer1.Enabled = True
Timer2.Enabled = False

End If
End Sub

Private Sub Command1_Click()
Label16.Caption = "0"
Label18.Caption = "0"
End Sub



Private Sub difficult1_Click()
Label12.Top = 1420
human1.Checked = False
very1.Checked = False
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = True

End Sub

Private Sub difficult2_Click()
Label26.Top = 1420
human2.Checked = False
very2.Checked = False
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = True

End Sub

Private Sub easy1_Click()
Label12.Top = 950
human1.Checked = False
very1.Checked = False
easy1.Checked = True
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub easy2_Click()
Label26.Top = 950
human2.Checked = False
very2.Checked = False
easy2.Checked = True
normal2.Checked = False
difficult2.Checked = False

End Sub

Private Sub exit_Click()
Unload Form1
End
End Sub

Private Sub fast_Click()
slow.Checked = False
normal.Checked = False
fast.Checked = True
Timer1.Interval = 70
Timer2.Interval = 70
End Sub

Private Sub Form_Load()
newgame
boardwidth = 7
boardheight = 7
horlinenum = 55
verlinenum = 55
loaded1 = 55
loaded2 = 55
horlinewidth = 855
verlineheight = 615
End Sub

Private Sub gridsize_Click()
On Error Resume Next
'd = InputBox("please enter the horizontal cells (2 - 12)", "board size")
'boardwidth = d
'If boardwidth > 12 Then boardwidth = 12
'If boardwidth < 2 Then boardwidth = 2

'd = InputBox("please enter the vertical cells (2 - 12)", "board size")
'boardheight = d
'If boardheight > 12 Then boardheight = 12
'If boardheight < 2 Then boardheight = 2

d = InputBox("please enter the board size (2 - 12)", "board size", Str(boardwidth))
If d < 2 Then d = 2
If d > 12 Then d = 12
boardheight = d
boardwidth = d
newgame
resizeboard
End Sub

Private Sub human1_Click()
Label12.Top = 460
human1.Checked = True
very1.Checked = False
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub human2_Click()
Label26.Top = 460
human2.Checked = True
very2.Checked = False
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = False

End Sub

Private Sub Label2_Click(index As Integer)
If Label2(index).Enabled = False Then Exit Sub
If (turn = "blue" And human1.Checked = False) Or (turn = "red" And human2.Checked = False) Then Exit Sub
takeline (index)
End Sub

Private Sub Label3_Click(index As Integer)
If Label3(index).Enabled = False Then Exit Sub
If (turn = "blue" And human1.Checked = False) Or (turn = "red" And human2.Checked = False) Then Exit Sub
takeline (index + (horlinenum + 1))
End Sub

Private Sub new_Click()
newgame
End Sub

Private Sub normal_Click()
slow.Checked = False
normal.Checked = True
fast.Checked = False
Timer1.Interval = 200
Timer2.Interval = 200

End Sub

Private Sub normal1_Click()
Label12.Top = 1180
human1.Checked = False
very1.Checked = False
easy1.Checked = False
normal1.Checked = True
difficult1.Checked = False

End Sub

Private Sub normal2_Click()
Label26.Top = 1180
human2.Checked = False
very2.Checked = False
easy2.Checked = False
normal2.Checked = True
difficult2.Checked = False

End Sub

Private Sub pauseresume_Click()
Select Case pauseresume.Caption
Case "Pause"
pauseresume.Caption = "Resume"
saveenabled
enabled1 = Timer1.Enabled
enabled2 = Timer2.Enabled
Timer1.Enabled = False
Timer2.Enabled = False
Case "Resume"
pauseresume.Caption = "Pause"
revertenabled
Timer1.Enabled = enabled1
Timer2.Enabled = enabled2
End Select
End Sub

Private Sub slow_Click()
slow.Checked = True
normal.Checked = False
fast.Checked = False
Timer1.Interval = 500
Timer2.Interval = 500

End Sub

Private Sub suggest_Click()
If human1.Checked = True And human2.Checked = True Then
d = MsgBox("Can't suggest move if two humans are playing!", vbInformation)
Exit Sub
End If
If (turn = "blue" And human1.Checked = True) Or (turn = "red" And human2.Checked = True) Then
Call play("blue", True)
End If
End Sub

Private Sub Timer1_Timer()
If human1.Checked = False And turn = "blue" Then
Call play("blue", False)
End If
End Sub



Private Sub Timer2_Timer()
If human2.Checked = False And turn = "red" Then
Call play("red", False)
End If

End Sub

Private Sub undo_Click()
undomove
End Sub

Private Sub very1_Click()
Label12.Top = 700
human1.Checked = False
very1.Checked = True
easy1.Checked = False
normal1.Checked = False
difficult1.Checked = False

End Sub

Private Sub very2_Click()
Label26.Top = 700
human2.Checked = False
very2.Checked = True
easy2.Checked = False
normal2.Checked = False
difficult2.Checked = False

End Sub
