VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00007F00&
   Caption         =   "Concentration"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrplayer2 
      Interval        =   800
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer tmrplayer1 
      Interval        =   800
      Left            =   2040
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1320
      Top             =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "player1:0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "player2:0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "player1 turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   47
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   46
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   45
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   44
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   43
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   42
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   41
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   40
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   39
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   38
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   37
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   36
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   35
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   34
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   33
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   32
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   31
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   30
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   29
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   28
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   27
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   26
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   25
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   24
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   23
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   22
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   21
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   20
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   19
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   18
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   17
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   16
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   15
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   14
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   13
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   12
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   11
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   10
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   9
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   8
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   7
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   6
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   5
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   4
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   1455
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   48
      Left            =   0
      Picture         =   "concencration.frx":0000
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   47
      Left            =   0
      Picture         =   "concencration.frx":0FAA
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   46
      Left            =   0
      Picture         =   "concencration.frx":1F54
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   45
      Left            =   0
      Picture         =   "concencration.frx":2EFE
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   44
      Left            =   0
      Picture         =   "concencration.frx":3EA8
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   43
      Left            =   0
      Picture         =   "concencration.frx":4E52
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   42
      Left            =   0
      Picture         =   "concencration.frx":5DFC
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   41
      Left            =   0
      Picture         =   "concencration.frx":6DA6
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   40
      Left            =   0
      Picture         =   "concencration.frx":7D50
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   39
      Left            =   0
      Picture         =   "concencration.frx":8CFA
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   38
      Left            =   0
      Picture         =   "concencration.frx":9CA4
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   37
      Left            =   0
      Picture         =   "concencration.frx":AC4E
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   36
      Left            =   0
      Picture         =   "concencration.frx":BBF8
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   35
      Left            =   0
      Picture         =   "concencration.frx":CBA2
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   34
      Left            =   0
      Picture         =   "concencration.frx":DB4C
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   33
      Left            =   0
      Picture         =   "concencration.frx":EAF6
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   32
      Left            =   0
      Picture         =   "concencration.frx":FAA0
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   31
      Left            =   0
      Picture         =   "concencration.frx":10A4A
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   30
      Left            =   0
      Picture         =   "concencration.frx":119F4
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   29
      Left            =   0
      Picture         =   "concencration.frx":1299E
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   28
      Left            =   0
      Picture         =   "concencration.frx":13948
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   27
      Left            =   0
      Picture         =   "concencration.frx":148F2
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   26
      Left            =   0
      Picture         =   "concencration.frx":1589C
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   25
      Left            =   0
      Picture         =   "concencration.frx":16846
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   24
      Left            =   0
      Picture         =   "concencration.frx":177F0
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   23
      Left            =   0
      Picture         =   "concencration.frx":1879A
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   22
      Left            =   0
      Picture         =   "concencration.frx":19744
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   21
      Left            =   0
      Picture         =   "concencration.frx":1A6EE
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   20
      Left            =   0
      Picture         =   "concencration.frx":1B698
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   19
      Left            =   0
      Picture         =   "concencration.frx":1C642
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   18
      Left            =   0
      Picture         =   "concencration.frx":1D5EC
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   17
      Left            =   0
      Picture         =   "concencration.frx":1E596
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   16
      Left            =   0
      Picture         =   "concencration.frx":1F540
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   15
      Left            =   0
      Picture         =   "concencration.frx":204EA
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   14
      Left            =   0
      Picture         =   "concencration.frx":21494
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   13
      Left            =   0
      Picture         =   "concencration.frx":2243E
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   12
      Left            =   0
      Picture         =   "concencration.frx":233E8
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   11
      Left            =   0
      Picture         =   "concencration.frx":24392
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   10
      Left            =   0
      Picture         =   "concencration.frx":2533C
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   9
      Left            =   0
      Picture         =   "concencration.frx":262E6
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   8
      Left            =   0
      Picture         =   "concencration.frx":27290
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   7
      Left            =   0
      Picture         =   "concencration.frx":2823A
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   6
      Left            =   0
      Picture         =   "concencration.frx":291E4
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   5
      Left            =   0
      Picture         =   "concencration.frx":2A18E
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   4
      Left            =   0
      Picture         =   "concencration.frx":2B138
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   3
      Left            =   0
      Picture         =   "concencration.frx":2C0E2
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   2
      Left            =   0
      Picture         =   "concencration.frx":2D08C
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   1
      Left            =   0
      Picture         =   "concencration.frx":2E036
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   0
      Left            =   0
      Picture         =   "concencration.frx":2EFE0
      Stretch         =   -1  'True
      Top             =   9500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu type 
         Caption         =   "Game Type"
         Begin VB.Menu suite 
            Caption         =   "Match Suite"
         End
         Begin VB.Menu value 
            Caption         =   "Match Value"
         End
         Begin VB.Menu both 
            Caption         =   "Match Color and Value"
         End
      End
      Begin VB.Menu spp3 
         Caption         =   "-"
      End
      Begin VB.Menu ply1 
         Caption         =   "Player1"
         Begin VB.Menu human1 
            Caption         =   "Human"
         End
         Begin VB.Menu spp5 
            Caption         =   "-"
         End
         Begin VB.Menu novice1 
            Caption         =   "Novice Computer"
         End
         Begin VB.Menu good1 
            Caption         =   "Good Computer"
         End
         Begin VB.Menu expert1 
            Caption         =   "Expert Computer"
         End
         Begin VB.Menu master1 
            Caption         =   "Master Computer"
         End
      End
      Begin VB.Menu ply2 
         Caption         =   "Player2"
         Begin VB.Menu none 
            Caption         =   "(none)"
         End
         Begin VB.Menu spp1 
            Caption         =   "-"
         End
         Begin VB.Menu human2 
            Caption         =   "Human"
         End
         Begin VB.Menu spp2 
            Caption         =   "-"
         End
         Begin VB.Menu novice2 
            Caption         =   "Novice Computer"
         End
         Begin VB.Menu good2 
            Caption         =   "Good Computer"
         End
         Begin VB.Menu expert2 
            Caption         =   "Expert Computer"
         End
         Begin VB.Menu master2 
            Caption         =   "Master Computer"
         End
      End
      Begin VB.Menu spp4 
         Caption         =   "-"
      End
      Begin VB.Menu deck 
         Caption         =   "Deck"
      End
      Begin VB.Menu background 
         Caption         =   "Background Color"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shown As Integer
Dim shown1 As Integer, shown2 As Integer
Dim gametype As Integer
Dim matched As Boolean
Dim positions(47, 2) As Integer
Dim turns As Integer
Dim score1 As Integer, score2 As Integer
Dim player1 As String, player2 As String
Private Sub enable()
For k = 0 To 47
Image2(k).Enabled = True
Next k
End Sub
Private Sub newgame()
For i = 1 To 48
Image2(i - 1).Visible = True
Image2(i - 1).Picture = Image1(0).Picture
Image2(i - 1).Tag = 0

Next i
For h = 0 To 47
 Image2(h).Top = positions(h, 1)
 Image2(h).Left = positions(h, 2)
Next h

shuffle
shown1 = 0
shown2 = 0

shown = 0
matched = False
Label1.Caption = "player1 turn"
Label2.Caption = "player1:0"
Label3.Caption = "player2:0"
score1 = 0
score2 = 0
End Sub
Private Sub shuffle()
For k = 0 To 65
Randomize
m = Int(Rnd * 47)
v = Int(Rnd * 47)
If Image2(m).Visible = True And Image2(v).Visible = True Then
u = Image2(m).Top
i = Image2(m).Left
Image2(m).Top = Image2(v).Top
Image2(m).Left = Image2(v).Left
Image2(v).Top = u
Image2(v).Left = i
End If
Next k
End Sub


Private Sub background_Click()
CommonDialog1.Color = Form1.BackColor
CommonDialog1.ShowColor
Form1.BackColor = CommonDialog1.Color
End Sub

Private Sub both_Click()
gametype = 3
newgame
End Sub


Private Sub Command1_Click()
shuffle
End Sub

Private Sub deck_Click()
Load Form2
Form2.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub expert1_Click()
player1 = "computer"

End Sub

Private Sub expert2_Click()
turns = 2
player2 = "computer"
End Sub

Private Sub Form_Load()
For h = 0 To 47
positions(h, 1) = Image2(h).Top
positions(h, 2) = Image2(h).Left
Next h
newgame
gametype = 2
shuffle
turns = 1
player1 = "human"
player2 = "none"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub good1_Click()
player1 = "computer"
End Sub

Private Sub good2_Click()
turns = 2
player2 = "computer"
End Sub

Private Sub human1_Click()
player1 = "human"
End Sub

Private Sub human2_Click()
turns = 2
player2 = "human"
End Sub

Private Sub Image2_Click(Index As Integer)
If Image2(Index).Picture = Image1(0).Picture Then
Image2(Index).Picture = Image1(Index + 1).Picture
Image1(Index).Tag = 1
shown = shown + 1
If shown = 1 Then
shown1 = Image2(Index).Index
End If
If shown = 2 Then
shown2 = Image2(Index).Index
For j = 0 To 47
Image2(j).Enabled = False
Next j

Select Case gametype
Case 1
If Int(Image2(shown1).Index / 12) = Int(Image2(shown2).Index / 12) Then
matched = True
Else
matched = False
End If
Case 2
If (Image2(shown1).Index) Mod 12 = (Image2(shown2).Index) Mod 12 Then
matched = True
Else
matched = False
End If
Case 3
If ((Image2(shown1).Index) Mod 12 = (Image2(shown2).Index) Mod 12) And Int(Image2(shown1).Index / 24) = Int(Image2(shown2).Index / 24) Then
matched = True
Else
matched = False
End If

End Select
Timer1.Enabled = True

End If
End If

End Sub


Private Sub master1_Click()
player1 = "computer"
End Sub

Private Sub master2_Click()
turns = 2
player2 = "computer"
End Sub

Private Sub new_Click()
newgame
End Sub

Private Sub none_Click()
turns = 1
player2 = "none"
End Sub

Private Sub novice1_Click()
player1 = "computer"
End Sub

Private Sub novice2_Click()
turns = 2
player2 = "computer"
End Sub

Private Sub suite_Click()
gametype = 1
newgame
End Sub

Private Sub Timer1_Timer()
If matched = True Then
Image2(shown1).Visible = False
Image2(shown2).Visible = False
If Label1.Caption = "player1 turn" Then
score1 = score1 + 5
Label2.Caption = "player1:" + Str(score1)
Else
score2 = score2 + 5
Label3.Caption = "player2:" + Str(score2)
End If

Else
Image2(shown1).Picture = Image1(0).Picture
Image2(shown2).Picture = Image1(0).Picture
If Label1.Caption = "player1 turn" Then
score1 = score1 - 1
Label2.Caption = "player1:" + Str(score1)
Else
score2 = score2 - 1
Label3.Caption = "player2:" + Str(score2)
End If

If turns = 2 Then
If Label1.Caption = "player1 turn" Then
Label1.Caption = "player2 turn"
Else
Label1.Caption = "player1 turn"
End If
End If
End If
shown = 0

Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
enable
Timer2.Enabled = False
End Sub

Private Sub tmrplayer1_Timer()
If Label1.Caption = "player1 turn" And player1 = "computer" Then
Select Case shown
Case 0
Image2_Click (Rnd * 47)
Case 1
Image2_Click (Rnd * 47)
End Select
End If
End Sub

Private Sub value_Click()
gametype = 2
newgame
End Sub
