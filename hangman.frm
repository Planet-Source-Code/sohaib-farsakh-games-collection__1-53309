VERSION 5.00
Begin VB.Form HangMan 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Hang Man"
   ClientHeight    =   5895
   ClientLeft      =   330
   ClientTop       =   660
   ClientWidth     =   6570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   3300
      TabIndex        =   2
      Top             =   2805
      Visible         =   0   'False
      Width           =   3120
      Begin VB.CommandButton Command1 
         Caption         =   "Quit Game"
         Height          =   465
         Index           =   1
         Left            =   1605
         TabIndex        =   5
         Top             =   1995
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Word"
         Height          =   465
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   1980
         Width           =   1365
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   285
         TabIndex        =   4
         Top             =   360
         Width           =   2580
      End
   End
   Begin VB.Line Line2 
      Index           =   8
      Visible         =   0   'False
      X1              =   2100
      X2              =   1665
      Y1              =   4335
      Y2              =   4575
   End
   Begin VB.Line Line2 
      Index           =   7
      Visible         =   0   'False
      X1              =   2550
      X2              =   2130
      Y1              =   4650
      Y2              =   4320
   End
   Begin VB.Line Line2 
      Index           =   6
      Visible         =   0   'False
      X1              =   2130
      X2              =   2130
      Y1              =   4320
      Y2              =   4005
   End
   Begin VB.Line Line2 
      Index           =   5
      Visible         =   0   'False
      X1              =   1875
      X2              =   2130
      Y1              =   5175
      Y2              =   4935
   End
   Begin VB.Line Line2 
      Index           =   2
      Visible         =   0   'False
      X1              =   2370
      X2              =   2145
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line Line2 
      Index           =   3
      Visible         =   0   'False
      X1              =   2190
      X2              =   2325
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      Index           =   4
      Visible         =   0   'False
      X1              =   1890
      X2              =   2025
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      Index           =   1
      Visible         =   0   'False
      X1              =   1875
      X2              =   1605
      Y1              =   5175
      Y2              =   4905
   End
   Begin VB.Line Line2 
      Index           =   0
      Visible         =   0   'False
      X1              =   2385
      X2              =   2610
      Y1              =   5160
      Y2              =   4875
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   2070
      Shape           =   2  'Oval
      Top             =   4005
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Word 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Index           =   1
      Left            =   315
      TabIndex        =   1
      Top             =   1860
      Width           =   585
   End
   Begin VB.Line Line1 
      Index           =   4
      Visible         =   0   'False
      X1              =   2130
      X2              =   1590
      Y1              =   4320
      Y2              =   4035
   End
   Begin VB.Line Line1 
      Index           =   5
      Visible         =   0   'False
      X1              =   2130
      X2              =   2700
      Y1              =   4290
      Y2              =   3960
   End
   Begin VB.Line Line1 
      Index           =   7
      Visible         =   0   'False
      X1              =   2130
      X2              =   2715
      Y1              =   4920
      Y2              =   5475
   End
   Begin VB.Line Line1 
      Index           =   6
      Visible         =   0   'False
      X1              =   2100
      X2              =   1605
      Y1              =   4950
      Y2              =   5445
   End
   Begin VB.Line Line1 
      Index           =   3
      Visible         =   0   'False
      X1              =   2130
      X2              =   2130
      Y1              =   4080
      Y2              =   4965
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   2040
      X2              =   2235
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Shape Shape2 
      Height          =   315
      Index           =   3
      Left            =   2055
      Shape           =   2  'Oval
      Top             =   3645
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape Shape2 
      Height          =   165
      Index           =   2
      Left            =   2175
      Shape           =   3  'Circle
      Top             =   3525
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape Shape2 
      Height          =   165
      Index           =   1
      Left            =   1890
      Shape           =   3  'Circle
      Top             =   3540
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Line Line1 
      Index           =   2
      Visible         =   0   'False
      X1              =   2085
      X2              =   2085
      Y1              =   3045
      Y2              =   3390
   End
   Begin VB.Line Line1 
      Index           =   1
      Visible         =   0   'False
      X1              =   930
      X2              =   2085
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line1 
      Index           =   0
      Visible         =   0   'False
      X1              =   915
      X2              =   915
      Y1              =   3030
      Y2              =   5670
   End
   Begin VB.Label Letter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Index           =   0
      Left            =   1545
      Shape           =   3  'Circle
      Top             =   3375
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Menu mnuNW 
      Caption         =   "&New Word"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "HangMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As Integer, C As Integer, TopOfLtr As Single, LtrNumb As Integer, num As Integer, Hits As Integer
Dim LeftOfLtr As Single, LetterWidth As Integer, FileNum As Integer, dataarray(), TheWord As String
Dim Misses As Integer, I As Integer
Sub GameLost()
    Dim Speed As Integer
    Speed = 10250
    Shape2(0).FillColor = &HFFC0FF
    For R = 1 To 10
        Line1(4).Visible = False
        Line1(5).Visible = False
        Line1(6).Visible = False
        Line1(7).Visible = False
        For I = 0 To Line2.Count - 1
            Line2(I).Visible = True
        Next
        For I = 0 To Speed
            DoEvents
        Next
        For I = 0 To Line2.Count - 1
            Line2(I).Visible = False
        Next
        Line1(4).Visible = True
        Line1(5).Visible = True
        Line1(6).Visible = True
        Line1(7).Visible = True
        For I = 0 To Speed
            DoEvents
        Next
    Next
    Shape2(0).FillColor = &H800080
    Shape1.Visible = True
    Label1.Caption = "You've Lost! " & Chr(10) & Chr(13) & "The Word Was " & TheWord
    Frame1.Visible = True
End Sub
Sub SelectWord()
    Dim WORDLEFT As Single, WordNum As Integer
    Misses = 0
    Hits = 0
    For I = 0 To Line1.Count - 1
        Line1(I).Visible = False
    Next
    For I = 0 To Shape2.Count - 1
        Shape2(I).Visible = False
    Next
    Line4.Visible = False
    Shape2(0).FillColor = &HFFFFFF
    Randomize
    WORDLEFT = Word(1).Left + Word(1).Width
    WordNum = Int((20 - 0 + 1) * Rnd + 0)
    TheWord = dataarray(WordNum)
    For R = 1 To Len(TheWord)
        If R > 1 Then
            Load Word(R)
            Word(R).Left = WORDLEFT
            WORDLEFT = WORDLEFT + Word(1).Width
        End If
        Word(R).Visible = True
    Next
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame1.Visible = False
            mnuNW_Click
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii > 64 And KeyAscii < 91 Then
        If Letter(KeyAscii - 65).Enabled Then Letter_Click KeyAscii - 65
    End If
    If KeyAscii > 96 And KeyAscii < 123 Then
        If Letter(KeyAscii - 97).Enabled Then Letter_Click KeyAscii - 97
    End If
End Sub

Private Sub Form_Load()
    ReDim dataarray(1)
    FileNum = FreeFile
    LtrNumb = 65
    TopOfLtr = Letter(0).Top
    Letter(0) = Chr(LtrNumb)
    LetterWidth = Letter(0).Width
    LeftOfLtr = Letter(0).Left + LetterWidth
    LtrNumb = LtrNumb + 1
    For C = 1 To 25
        Load Letter(C)
        If C = 13 Then
            TopOfLtr = TopOfLtr + Letter(0).Height
            LeftOfLtr = Letter(0).Left
        End If
        Letter(C).Top = TopOfLtr
        Letter(C).Left = LeftOfLtr
        Letter(C) = Chr(LtrNumb)
        LeftOfLtr = LeftOfLtr + LetterWidth
        Letter(C).Visible = True
        LtrNumb = LtrNumb + 1
    Next
    num = 1
    Open App.Path & "\HangMan.dat" For Input As FileNum
    Do While Not EOF(FileNum)
        Input #FileNum, dataarray(num - 1)
        num = num + 1
        ReDim Preserve dataarray(num - 1)
    Loop
    Close #FileNum
    SelectWord
End Sub

Private Sub Letter_Click(Index As Integer)
    Dim Di%, I%
    Di% = InStr(TheWord, Letter(Index))
    If Di% > 0 Then
        Word(Di%) = Letter(Index)
        Hits = Hits + 1
        Do While Di% <> 0
            Di% = InStr(Di% + 1, TheWord, Letter(Index))
            If Di% > 0 Then
                Word(Di%) = Letter(Index)
                Hits = Hits + 1
            End If
        Loop
    Else
        'MsgBox Letter(Index) & " Not in the Word"
        If Misses < 3 Then Line1(Misses).Visible = True
        If Misses = 3 Then
            For I = 0 To Shape2.Count - 1
                Shape2(I).Visible = True
            Next
            Line4.Visible = True
        End If
        If Misses > 3 Then Line1(Misses - 1).Visible = True
        Misses = Misses + 1
        If Misses = 9 Then
            GameLost
            For R = 0 To 25
                Letter(R).Enabled = False
            Next
        End If
    End If
    If Hits = Len(TheWord) Then
        For R = 0 To 25
            Letter(R).Enabled = True
        Next
        Label1.Caption = "You've Won!"
        Frame1.Visible = True
    End If
    Letter(Index).Enabled = False
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNW_Click()
    For R = 2 To Word.Count
        Unload Word(R)
    Next
    For R = 0 To 25
        Letter(R).Enabled = True
    Next
    Shape1.Visible = False
    Word(1) = ""
    SelectWord
End Sub

