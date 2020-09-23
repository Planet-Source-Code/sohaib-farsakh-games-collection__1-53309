VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form1 
   Caption         =   "3D Graph"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "axis"
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   1440
      ScaleHeight     =   3195
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
   Begin MSScriptControlCtl.ScriptControl scc 
      Left            =   120
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xang As Single, yang As Single, zang  As Single
Dim x As Single, y As Single, z As Single
Private Sub Command1_Click()
x = 5
y = 6
scc.ExecuteStatement ("x=" & x)
scc.ExecuteStatement ("y=" & y)
a = scc.Eval(Trim(Text1.Text))
Text2.Text = Str(a)
End Sub

Private Sub rotabouty(ByVal xx As Single, ByVal yy As Single, ByVal zz As Single, angle As Single)
    oldx = xx
    oldy = yy
    oldz = zz
    
    x = (oldx * Cos(angle)) - (oldz * Sin(angle))
    y = yy
    z = (oldx * Sin(angle)) + (oldz * Cos(angle))
End Sub

Private Sub rotaboutx(ByVal xx As Single, ByVal yy As Single, ByVal zz As Single, angle As Single)
    oldx = xx
    oldy = yy
    oldz = zz
    
    x = xx
    y = (oldz * Sin(angle)) + (oldy * Cos(angle))
    z = (oldz * Cos(angle)) - (oldy * Sin(angle))
End Sub

Private Sub rotaboutz(ByVal xx As Single, ByVal yy As Single, ByVal zz As Single, angle As Single)
    oldx = xx
    oldy = yy
    oldz = zz
    
    x = (oldx * Sin(angle)) + (oldy * Cos(angle))
    y = (oldx * Cos(angle)) - (oldy * Sin(angle))
    z = zz
End Sub

Private Function Rads(deg As Single) As Single
    Rads = deg * 3.1459 / 180 ' convert the angle into radians
End Function

Private Sub Command2_Click()
On Error Resume Next
Dim xres(1 To 3) As Variant, yres(1 To 3) As Variant
For xxx = -2 To 1.5 Step 0.25
For yyy = -2 To 1.5 Step 0.25
scc.Reset
scc.ExecuteStatement ("x=" & xxx)
scc.ExecuteStatement ("y=" & yyy)
zzz = -scc.Eval(Trim(Text1.Text))
Call rotaboutx(xxx, yyy, zzz, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(1) = x
yres(1) = y
'Form1.Caption = Str(xres(1))

X2 = xxx + 0.25
scc.Reset
scc.ExecuteStatement ("x=" & X2)
scc.ExecuteStatement ("y=" & yyy)
zzz = -scc.Eval(Trim(Text1.Text))
Call rotaboutx(X2, yyy, zzz, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(2) = x
yres(2) = y

Y2 = yyy + 0.25
scc.Reset
scc.ExecuteStatement ("x=" & xxx)
scc.ExecuteStatement ("y=" & Y2)
zzz = -scc.Eval(Trim(Text1.Text))
Call rotaboutx(xxx, Y2, zzz, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(3) = x
yres(3) = y

Picture1.Line (xres(1), yres(1))-(xres(2), yres(2))
Picture1.Line (xres(1), yres(1))-(xres(3), yres(3))

Next
Next
End Sub

Private Sub Command3_Click()
coor
End Sub

Private Sub coor()
Dim xres(1 To 6), yres(1 To 6)
Call rotaboutx(-5, 0, 0, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(1) = x
yres(1) = y

Call rotaboutx(5, 0, 0, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(2) = x
yres(2) = y

Call rotaboutx(0, -5, 0, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(3) = x
yres(3) = y

Call rotaboutx(0, 5, 0, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(4) = x
yres(4) = y

Call rotaboutx(0, 0, -5, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(5) = x
yres(5) = y

Call rotaboutx(0, 0, 5, xang)
Call rotabouty(x, y, z, yang)
Call rotaboutz(x, y, z, zang)
xres(6) = x
yres(6) = y
Picture1.ForeColor = vbBlack
Picture1.Line (xres(1), yres(1))-(xres(2), yres(2))
Picture1.ForeColor = vbRed
Picture1.Line (xres(3), yres(3))-(xres(4), yres(4))
Picture1.ForeColor = vbBlue
Picture1.Line (xres(5), yres(5))-(xres(6), yres(6))

End Sub

Private Sub Command4_Click()
Picture1.Cls
End Sub

Private Sub Form_Load()
xang = Rads(56.5)
yang = Rads(60)
zang = Rads(60)
Picture1.Scale (-5, 5)-(5, -5)
End Sub
