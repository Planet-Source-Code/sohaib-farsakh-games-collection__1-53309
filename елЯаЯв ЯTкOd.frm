VERSION 5.00
Object = "{A37D0E58-B6D0-11D2-971F-EC500970267D}#7.0#0"; "ANALOGCLOCK.OCX"
Begin VB.Form Form1 
   Caption         =   "ÍÓÇÈÇÊ ÇáÒãä"
   ClientHeight    =   6495
   ClientLeft      =   735
   ClientTop       =   2010
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "ÎÑæÌ"
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "<"
      Height          =   495
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   101
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   ">"
      Height          =   495
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   100
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "ÇáÂáÉ ÇáÍÇÓÈÉ"
      Height          =   1935
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   4560
      Width           =   4335
      Begin VB.CommandButton Command10 
         Caption         =   "^2"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "sqr"
         Height          =   255
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "/"
         Height          =   255
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "*"
         Height          =   255
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "-"
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   255
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáäÇÊÌ"
         Height          =   375
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÚãáíÉ"
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÑÞã ÇáËÇäí"
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÑÞã ÇáÃæá"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ÍÓÇÈÇÊ ÇáÓÑÚÉ"
      Height          =   1815
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   2760
      Width           =   3255
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0000
         Left            =   960
         List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":000D
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Text            =   "ËÇäíÉ"
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0025
         Left            =   960
         List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0035
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Text            =   "ãÊÑ/ËÇäíÉ"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0062
         Left            =   960
         List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":006F
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Text            =   "ãÊÑ"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ÇáÌæÇÈ"
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label56 
         Height          =   255
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÒãä ÈæÍÏÉ"
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÓÑÚÉ ÈæÍÏÉ"
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÓÇÝÉ ÈæÍÏÉ"
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0086
      Left            =   4560
      List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":00A5
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "ÊÍæíá æÍÏÇÊ ÇáÒãä"
      Height          =   1815
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   2760
      Width           =   3135
      Begin VB.CommandButton Command4 
         Caption         =   "Íæá"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":00DD
         Left            =   1200
         List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":00FC
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Text            =   "ÇÎÊÑ ÇáæÍÏÉ"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Åáì:"
         Height          =   255
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Íæá ãä:"
         Height          =   255
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇáæÞÊ æÇÊÇÑíÎ"
      Height          =   2655
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   0
      Width           =   2055
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   2160
      End
      Begin AnalogClockControl.AnalogClock AnalogClock1 
         Height          =   1575
         Left            =   240
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         BackColor       =   16777215
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label49"
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label48"
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " ÍÓÇÈ ÇáÒãä ÇáããÖì Èíä áÍÙÊíä"
      Height          =   2655
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0134
         Left            =   240
         List            =   "ÍÓÇÈÇÊ ÇáÒãä.frx":0147
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Text            =   "ÓáÓáÉ"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ÍÓÇÈ ÇáÒãä"
         Height          =   375
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "ÅÚØÇÁ ÇáÌæÇÈ ÈÕæÑÉ:"
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "ÇáÏÞÇÆÞ"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÓÇÚÉ(ãä 24)"
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÇÑíÎ(ÓäÉ/ÔåÑ/íæã)"
         Height          =   375
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2040
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÇáÑÒäÇãÉ"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Çáíæã"
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   3255
      Left            =   6600
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   48
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   47
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6720
      TabIndex        =   46
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7320
      TabIndex        =   45
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8040
      TabIndex        =   44
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8760
      TabIndex        =   43
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9480
      TabIndex        =   42
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   41
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   40
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8760
      TabIndex        =   36
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9480
      TabIndex        =   35
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   34
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   33
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6720
      TabIndex        =   32
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8040
      TabIndex        =   30
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8760
      TabIndex        =   29
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9480
      TabIndex        =   28
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   27
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8040
      TabIndex        =   23
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8760
      TabIndex        =   22
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "ÃÍÏ         ÇËäíä         ËáÇËÇÁ       ÃÑÈÚÇÁ       ÎãíÓ       ÌãÚÉ       ÓÈÊ"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Çáíæã                          ÇáÔåÑ                        ÇáÚÇã                    "
      Height          =   255
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text4.Text = Text5.Text
Label3.Caption = ((Val(Text3.Text) - 1) * 365) + Int((Val(Text3.Text) - 1) / 4)
Select Case Text2.Text
Case "1"
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text)
Case "2"
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 31
Case "3"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 60
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 59
End If
Case "4"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 91
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 90
End If
Case "5"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 121
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 120
End If
Case "6"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 152
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 151
End If
Case "7"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 182
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 181
End If
Case "8"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 213
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 212
End If
Case "9"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 244
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 243
End If
Case "10"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 274
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 273
End If
Case "11"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 305
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 304
End If
Case "12"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 335
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 334
End If
End Select
a = Val(Label3.Caption) / 7
b = Int(Val(Label3.Caption) / 7)
c = a - b
Label4.Caption = 7 * c
d = Label4.Caption
Label4.Caption = Int(Label4.Caption)
If d - Val(Label4.Caption) > 0.5 Then
Label4.Caption = Val(Label4.Caption) + 1
End If
Select Case Label4.Caption
Case "0"
Label2.Caption = "ÇáÓÈÊ"

Case "1"
Label2.Caption = "ÇáÃÍÏ"
Case "2"
Label2.Caption = "ÇáÇËäíä"
Case "3"
Label2.Caption = "ÇáËáÇËÇÁ"
Case "4"
Label2.Caption = "ÇáÃÑÈÚÇÁ"
Case "5"
Label2.Caption = "ÇáÎãíÓ"
Case "6"
Label2.Caption = "ÇáÌãÚÉ"
End Select
End Sub

Private Sub Command10_Click()
Text18.Text = Val(Text16.Text) ^ 2
End Sub

Private Sub Command11_Click()
Text18.Text = Sqr(Text16.Text)
End Sub

Private Sub Command12_Click()
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 13 Then
Text3.Text = Val(Text3.Text) + 1
Text2.Text = "1"
End If
Command2_Click
End Sub

Private Sub Command13_Click()
Text2.Text = Val(Text2.Text) - 1
If Text2.Text = "0" Then
Text2.Text = "12"
Text3.Text = Val(Text3.Text) - 1
End If
Command2_Click
End Sub

Private Sub Command14_Click()
End
End Sub

Private Sub Command2_Click()
Text4.Text = "1"
Label3.Caption = ((Val(Text3.Text) - 1) * 365) + Int((Val(Text3.Text) - 1) / 4)
Select Case Text2.Text
Case "1"
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text)
Case "2"
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 31
Case "3"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 60
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 59
End If
Case "4"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 91
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 90
End If
Case "5"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 121
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 120
End If
Case "6"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 152
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 151
End If
Case "7"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 182
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 181
End If
Case "8"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 213
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 212
End If
Case "9"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 244
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 243
End If
Case "10"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 274
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 273
End If
Case "11"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 305
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 304
End If
Case "12"
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 335
Else
Label3.Caption = Val(Label3.Caption) + Val(Text1.Text) + 334
End If
End Select
a = Val(Label3.Caption) / 7
b = Int(Val(Label3.Caption) / 7)
c = a - b
Label4.Caption = 7 * c
d = Label4.Caption
Label4.Caption = Int(Label4.Caption)
If d - Val(Label4.Caption) > 0.5 Then
Label4.Caption = Val(Label4.Caption) + 1
End If
Select Case Label4.Caption
Case "0"

Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = Val(Label11.Caption) + 1
Case "1"

Label6.Caption = "1"
Label7.Caption = Val(Label6.Caption) + 1
Label8.Caption = Val(Label7.Caption) + 1
Label9.Caption = Val(Label8.Caption) + 1
Label10.Caption = Val(Label9.Caption) + 1
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
Case "2"
Label6.Caption = ""
Label7.Caption = Val(Label6.Caption) + 1
Label8.Caption = Val(Label7.Caption) + 1
Label9.Caption = Val(Label8.Caption) + 1
Label10.Caption = Val(Label9.Caption) + 1
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
Case "3"
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = Val(Label7.Caption) + 1
Label9.Caption = Val(Label8.Caption) + 1
Label10.Caption = Val(Label9.Caption) + 1
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
Case "4"
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = Val(Label8.Caption) + 1
Label10.Caption = Val(Label9.Caption) + 1
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
Case "5"
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = Val(Label9.Caption) + 1
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
Case "6"
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = Val(Label10.Caption) + 1
Label12.Caption = Val(Label11.Caption) + 1
End Select
Label12.Caption = Val(Label11.Caption) + 1
Label13.Caption = Val(Label12.Caption) + 1
Label14.Caption = Val(Label13.Caption) + 1
Label15.Caption = Val(Label14.Caption) + 1
Label16.Caption = Val(Label15.Caption) + 1
Label17.Caption = Val(Label16.Caption) + 1
Label18.Caption = Val(Label17.Caption) + 1
Label19.Caption = Val(Label18.Caption) + 1
Label20.Caption = Val(Label19.Caption) + 1
Label21.Caption = Val(Label20.Caption) + 1
Label22.Caption = Val(Label21.Caption) + 1
Label23.Caption = Val(Label22.Caption) + 1
Label24.Caption = Val(Label23.Caption) + 1
Label25.Caption = Val(Label24.Caption) + 1
Label26.Caption = Val(Label25.Caption) + 1
Label27.Caption = Val(Label26.Caption) + 1
Label28.Caption = Val(Label27.Caption) + 1
Label29.Caption = Val(Label28.Caption) + 1
Label30.Caption = Val(Label29.Caption) + 1
Label31.Caption = Val(Label30.Caption) + 1
Label32.Caption = Val(Label31.Caption) + 1
Label33.Caption = Val(Label32.Caption) + 1
Label34.Caption = Val(Label33.Caption) + 1
Label35.Caption = Val(Label34.Caption) + 1
Label36.Caption = Val(Label35.Caption) + 1
Label37.Caption = Val(Label36.Caption) + 1
Label38.Caption = Val(Label37.Caption) + 1
Label39.Caption = Val(Label38.Caption) + 1
Label40.Caption = Val(Label39.Caption) + 1
Label41.Caption = Val(Label40.Caption) + 1
Label42.Caption = Val(Label41.Caption) + 1
If Val(Label37.Caption) > 31 Then
Label37.Caption = ""
End If
If Val(Label38.Caption) > 31 Then
Label38.Caption = ""
End If
If Val(Label39.Caption) > 31 Then
Label39.Caption = ""
End If
If Val(Label40.Caption) > 31 Then
Label40.Caption = ""
End If
If Val(Label41.Caption) > 31 Then
Label41.Caption = ""
End If
If Val(Label42.Caption) > 31 Then
Label42.Caption = ""
End If
If Text2.Text = "4" Or Text2.Text = "6" Or Text2.Text = "9" Or Text2.Text = "11" Then
If Val(Label37.Caption) > 30 Then
Label37.Caption = ""
End If
If Val(Label38.Caption) > 30 Then
Label38.Caption = ""
End If
If Val(Label39.Caption) > 30 Then
Label39.Caption = ""
End If
If Val(Label40.Caption) > 30 Then
Label40.Caption = ""
End If
If Val(Label41.Caption) > 30 Then
Label41.Caption = ""
End If
If Val(Label36.Caption) > 30 Then
Label36.Caption = ""
End If
If Val(Label42.Caption) > 30 Then
Label42.Caption = ""
End If
End If
If Text2.Text = "2" Then
If Val(Text3.Text) / 4 = Int(Val(Text3.Text) / 4) Then
If Val(Label37.Caption) > 29 Then
Label37.Caption = ""
End If
If Val(Label38.Caption) > 29 Then
Label38.Caption = ""
End If
If Val(Label39.Caption) > 29 Then
Label39.Caption = ""
End If
If Val(Label40.Caption) > 29 Then
Label40.Caption = ""
End If
If Val(Label41.Caption) > 29 Then
Label41.Caption = ""
End If
If Val(Label36.Caption) > 29 Then
Label36.Caption = ""
End If
If Val(Label35.Caption) > 29 Then
Label35.Caption = ""
End If
If Val(Label42.Caption) > 29 Then
Label42.Caption = ""
End If
End If
If Val(Text3.Text) / 4 <> Int(Val(Text3.Text) / 4) Then
If Val(Label37.Caption) > 28 Then
Label37.Caption = ""
End If
If Val(Label38.Caption) > 28 Then
Label38.Caption = ""
End If
If Val(Label39.Caption) > 28 Then
Label39.Caption = ""
End If
If Val(Label40.Caption) > 28 Then
Label40.Caption = ""
End If
If Val(Label41.Caption) > 28 Then
Label41.Caption = ""
End If
If Val(Label36.Caption) > 28 Then
Label36.Caption = ""
End If
If Val(Label35.Caption) > 28 Then
Label35.Caption = ""
End If
If Val(Label34.Caption) > 28 Then
Label34.Caption = ""
End If
If Val(Label42.Caption) > 28 Then
Label42.Caption = ""
End If
End If
End If

End Sub

Private Sub Command3_Click()
On Error Resume Next
a = Val(Text11.Text) - Val(Text8.Text)
b = (Val(Text10.Text) - Val(Text7.Text)) * 60
c = (DateValue(Text9.Text) - DateValue(Text6.Text)) * 1440
Label43.Tag = a + b + c
If Val(Label43.Tag) < 0 Then
Label43.Tag = -Val(Label43.Tag)
End If
Select Case Combo1.Text
Case "ÃíÇã"
Label43.Caption = Val(Label43.Tag) / 1440
Case "ÃÓÇÈíÚ"
Label43.Caption = Val(Label43.Tag) / 10080
Case "ÓÇÚÇÊ"
Label43.Caption = Val(Label43.Tag) / 60
Case "ÏÞÇÆÞ"
Label43.Caption = Label43.Tag
Case "ÓáÓáÉ"
a = Label43.Tag / 10080
b = Int(Label43.Tag / 10080)
c = a - b
w = c * 10080
d = Int(w)
x = w - d
If x > 0.5 Then
d = d + 1
End If
Label43.Caption = Str$(b) + "ÃÓÇÈíÚ"
e = d / 1440
f = Int(d / 1440)
g = e - f
z = g * 1440
h = Int(z)
y = z - h
If y > 0.5 Then
h = h + 1
End If

Label43.Caption = Label43.Caption + " - " + Str$(f) + "ÃíÇã"

i = h / 60
j = Int(h / 60)
k = i - j
l = k * 60
m = Int(l)
n = l - m
If n > 0.5 Then
m = m + 1
End If
Label43.Caption = Label43.Caption + " - " + Str$(j) + "ÓÇÚÇÊ" + " - " + Str$(m) + "ÏÞÇÆÞ"
End Select
End Sub

Private Sub Command4_Click()
If Combo2.Text = "íæã" Then
Label52.Tag = Val(Text12.Text) * 86400
End If
If Combo2.Text = "ÓÇÚÉ" Then
Label52.Tag = Val(Text12.Text) * 3600
End If
If Combo2.Text = "ÏÞíÞÉ" Then
Label52.Tag = Val(Text12.Text) * 60
End If
If Combo2.Text = "ËÇäíÉ" Then
Label52.Tag = Val(Text12.Text)
End If
If Combo2.Text = "ÃÓÈæÚ" Then
Label52.Tag = Val(Text12.Text) * 604800
End If
If Combo2.Text = "ÔåÑ" Then
Label52.Tag = Val(Text12.Text) * 2592000
End If
If Combo2.Text = "ÓäÉ" Then
Label52.Tag = Val(Text12.Text) * 31557600
End If
If Combo2.Text = "ÚÞÏ" Then
Label52.Tag = Val(Text12.Text) * 315576000
End If
If Combo2.Text = "ÞÑä" Then
Label52.Tag = Val(Text12.Text) * 3155760000#
End If
If Combo3.Text = "íæã" Then
Label52.Caption = Val(Label52.Tag) / 86400
End If
If Combo3.Text = "ÓÇÚÉ" Then
Label52.Caption = Val(Label52.Tag) / 3600
End If
If Combo3.Text = "ÏÞíÞÉ" Then
Label52.Caption = Val(Label52.Tag) / 60
End If
If Combo3.Text = "ËÇäíÉ" Then
Label52.Caption = Val(Label52.Tag)
End If
If Combo3.Text = "ÃÓÈæÚ" Then
Label52.Caption = Val(Label52.Tag) / 604800
End If
If Combo3.Text = "ÔåÑ" Then
Label52.Caption = Val(Label52.Tag) / 2592000
End If
If Combo3.Text = "ÓäÉ" Then
Label52.Caption = Val(Label52.Tag) / 31557600
End If
If Combo3.Text = "ÚÞÏ" Then
Label52.Caption = Val(Label52.Tag) / 315576000
End If
If Combo3.Text = "ÞÑä" Then
Label52.Caption = Val(Label52.Tag) / 3155760000#
End If

End Sub

Private Sub Command5_Click()
Select Case Combo4.Text
Case "ãÊÑ"
Text13.Tag = Text13.Text
Case "ßíáæãÊÑ"
Text13.Tag = Val(Text13.Text) * 1000
Case "ãíá"
Text13.Tag = Val(Text13.Text) * 1609
End Select
Select Case Combo5.Text
Case "ãÊÑ/ËÇäíÉ"
Text14.Tag = Text14.Text
Case "ßíáæãÊÑ/ÓÇÚÉ"
Text14.Tag = Val(Text14.Text) / 3.6
Case "ãíá/ÓÇÚÉ"
Text14.Tag = Val(Text14.Text) / 3.6 * 1.609
Case "ÚÞÏÉ"
Text14.Tag = Val(Text14.Text) * 0.515
End Select
Select Case Combo6.Text
Case "ËÇäíÉ"
Text15.Tag = Text15.Text
Case "ÏÞíÞÉ"
Text15.Tag = Val(Text15.Text) * 60
Case "ÓÇÚÉ"
Text15.Tag = Val(Text15.Text) * 3600
End Select
If Text13.Text <> "" And Text14.Text <> "" And Text15.Text = "" Then
Label56.Caption = Val(Text13.Tag) / Val(Text14.Tag)
Select Case Combo6.Text
Case "ÏÞíÞÉ"
Label56.Caption = Val(Label56.Caption) / 60
Case "ÓÇÚÉ"
Label56.Caption = Val(Label56.Caption) / 3600
End Select
Label56.Caption = Val(Label56.Caption) * 1000000
a = Int(Val(Label56.Caption))
b = Val(Label56.Caption) - a
If b > 0.5 Then
a = a + 1
End If
Label56.Caption = Str$(a)
Label56.Caption = Val(Label56.Caption) / 1000000
End If
If Text13.Text <> "" And Text15.Text <> "" And Text14.Text = "" Then
Label56.Caption = Val(Text13.Tag) / Val(Text15.Tag)
Select Case Combo5.Text
Case "ßíáæãÊÑ/ÓÇÚÉ"
Label56.Caption = Val(Label56.Caption) * 3.6
Case "ãíá/ÓÇÚÉ"
Label56.Caption = Val(Label56.Caption) * 3.6 / 1.609
Case "ÚÞÏÉ"
Label56.Caption = Val(Label56.Caption) / 0.515
End Select
Label56.Caption = Val(Label56.Caption) * 1000000
a = Int(Val(Label56.Caption))
b = Val(Label56.Caption) - a
If b > 0.5 Then
a = a + 1
End If
Label56.Caption = Str$(a)
Label56.Caption = Val(Label56.Caption) / 1000000
End If
If Text13.Text = "" And Text15.Text <> "" And Text14.Text <> "" Then
Label56.Caption = Val(Text14.Tag) * Val(Text15.Tag)
Select Case Combo4.Text
Case "ßíáæãÊÑ"
Label56.Caption = Val(Label56.Caption) / 1000
Case "ãíá"
Label56.Caption = Val(Label56.Caption) / 1000 / 1.609
End Select
Label56.Caption = Val(Label56.Caption) * 1000000
a = Int(Val(Label56.Caption))
b = Val(Label56.Caption) - a
If b > 0.5 Then
a = a + 1
End If
Label56.Caption = Str$(a)
Label56.Caption = Val(Label56.Caption) / 1000000
End If
End Sub

Private Sub Command6_Click()
Text18.Text = Val(Text16.Text) + Val(Text17.Text)
End Sub

Private Sub Command7_Click()
Text18.Text = Val(Text16.Text) - Val(Text17.Text)
End Sub

Private Sub Command8_Click()
Text18.Text = Val(Text16.Text) * Val(Text17.Text)
End Sub

Private Sub Command9_Click()
On Error Resume Next
Text18.Text = Val(Text16.Text) / Val(Text17.Text)
End Sub

Private Sub Text4_Change()
Text1.Text = Text4.Text
End Sub

Private Sub Text5_Change()
Text4.Text = Text5.Text
End Sub

Private Sub Timer1_Timer()
Label48.Caption = Time
Label49.Caption = Date
AnalogClock1.Value = Time
End Sub
