VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Plot 
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Plot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15
      TabIndex        =   6
      Top             =   30
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Text            =   "3*Cos(3*x)*Sin(5*x)*Sin(x)"
      Top             =   15
      Width           =   2640
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5475
      TabIndex        =   1
      Text            =   "0.001"
      Top             =   15
      Width           =   840
   End
   Begin VB.TextBox txtScale 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      TabIndex        =   0
      Text            =   "10"
      Top             =   15
      Width           =   390
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   5880
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Accuracy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4710
      TabIndex        =   5
      Top             =   45
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Expression"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1185
      TabIndex        =   4
      Top             =   45
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Scale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6360
      TabIndex        =   3
      Top             =   60
      Width           =   600
   End
End
Attribute VB_Name = "Plot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Components "Microsoft Script Control 1.0"
Dim XMin As Integer
Dim XMax  As Integer
Dim YMin As Integer
Dim YMax As Integer
Private Sub Command1_Click()
        On Error GoTo errhandler:
Command1.Enabled = False
    Dim X As Double, Y As Double
    Cls
    ScaleLeft = XMin
    ScaleTop = YMax
    ScaleWidth = XMax - XMin
    ScaleHeight = -(YMax - YMin)
    ForeColor = vbBlack
    Call DrawLine(Val(txtScale))
    Call TikMark(Val(txtScale))
    DrawStyle = 0
    ForeColor = vbBlue
    Line (XMin, 0)-(XMax, 0)
    Line (0, YMin)-(0, YMax)
    ForeColor = vbRed
    ScriptControl1.Reset
    For X = XMin To XMax Step Val(Text2)
        ScriptControl1.ExecuteStatement ("X = " & X)

        Y = ScriptControl1.Eval(Trim(Text1.Text))
        PSet (X, Y)
    Next
    Command1.Enabled = True
Exit Sub

errhandler:

MsgBox "Invalid function"
Command1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = "X=" & X & ", " & "Y=" & Y
End Sub

Private Sub Form_Resize()
    XMin = -Val(txtScale)
    XMax = Val(txtScale)
    YMin = -Val(txtScale)
    YMax = Val(txtScale)
    ScaleLeft = XMin
    ScaleTop = YMax
    ScaleWidth = XMax - XMin
    ScaleHeight = -(YMax - YMin)
    Refresh
End Sub
Function DrawLine(Distance As Double)
    Dumy = IIf(XMax < 0, -XMax / Distance, XMax / Distance)
    XorY = IIf(XMax < 0, -XMax, XMax)
    While XorY >= XMin
        If XMax < 0 Then
            Line (XorY, -XMax)-(XorY, -XMax)
            Line (-XMax, XorY)-(-XMax, XorY)
        Else
            Line (XorY, XMax)-(XorY, -XMax)
            Line (XMax, XorY)-(-XMax, XorY)
        End If
        XorY = XorY - Dumy
    Wend
End Function
Function TikMark(Distance As Double)
    Dumy = IIf(XMax < 0, (-XMax) / Distance, XMax / Distance)
    XorY = IIf(XMax < 0, -XMax, XMax)
        While XorY >= XMin
            If YMax < 0 Then
                Line (XorY, (-YMax) / 100)-(XorY, -((-YMax) / 100))
            Else
                Line (XorY, YMax / 100)-(XorY, -(YMax) / 100)
            End If
            If XorY <> 0 Then
                Print Format(Round(XorY, 0), "0")
            End If
            XorY = XorY - Dumy
        Wend
    Dumy = IIf(YMax < 0, (-YMax) / Distance, YMax / Distance)
    XorY = IIf(YMax < 0, -YMax, YMax)
        While XorY >= YMin
            If XMax < 0 Then
                Line ((-XMax) / 25, XorY)-(-((-XMax) / 25), XorY)
            Else
                Line (XMax / 25, XorY)-(-(XMax) / 25, XorY)
            End If
            If XorY <> 0 Then
                Print Format(Round(XorY, 0), "0")
            End If
            XorY = XorY - Dumy
        Wend
End Function

Private Sub txtScale_Change()
    XMin = -Val(txtScale)
    XMax = Val(txtScale)
    YMin = -Val(txtScale)
    YMax = Val(txtScale)
End Sub
