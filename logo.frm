VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGO"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "logo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   6600
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4691
      Left            =   360
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   65.86
         X2              =   65.86
         Y1              =   46.602
         Y2              =   58.252
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   61.985
         X2              =   61.985
         Y1              =   46.602
         Y2              =   58.252
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   58.111
         X2              =   58.111
         Y1              =   46.602
         Y2              =   58.252
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xpos As Double, ypos As Double, angle As Double, pentype As String
Dim commands(1 To 100) As String
Const pi = 3.14159265358979
Dim reptimes As Integer
Dim linesnum As Integer

Private Sub fillarray()
On Error Resume Next
Dim elements As Integer, numlast As Boolean
statement = Text1.Text
elements = 0
For i = 1 To Len(statement)
If Mid(statement, i, 2) = "fd" Or Mid(statement, i, 2) = "bk" Or Mid(statement, i, 2) = "rt" Or Mid(statement, i, 2) = "lt" Then
elements = elements + 1
commands(elements) = Mid(statement, i, 2)
numlast = False
End If
If Mid(statement, i, 2) = "pu" Or Mid(statement, i, 2) = "pd" Then
elements = elements + 1
commands(elements) = Mid(statement, i, 2)
numlast = False
End If
If Mid(statement, i, 2) = "hd" Or Mid(statement, i, 2) = "st" Then
elements = elements + 1
commands(elements) = Mid(statement, i, 2)
numlast = False
End If

If Mid(statement, i, 1) = "(" Or Mid(statement, i, 1) = "[" Then
elements = elements + 1
commands(elements) = "("
numlast = False
End If
If Mid(statement, i, 1) = ")" Or Mid(statement, i, 1) = "]" Then
elements = elements + 1
commands(elements) = ")"
numlast = False
End If

If Mid(statement, i, 6) = "repeat" Then
elements = elements + 1
commands(elements) = "rp"
numlast = False
End If

If Mid(statement, i, 6) = "circle" Then
elements = elements + 1
commands(elements) = "cr"
numlast = False
End If

If Mid(statement, i, 4) = "home" Then
elements = elements + 1
commands(elements) = "home"
numlast = False
End If
If Mid(statement, i, 4) = "seth" Then
elements = elements + 1
commands(elements) = "seth"
numlast = False
End If
If Mid(statement, i, 10) = "setpensize" Then
elements = elements + 1
commands(elements) = "size"
numlast = False
End If

If isnumber(Mid(statement, i, 1)) = True Then
If numlast = True Then
commands(elements) = commands(elements) + Mid(statement, i, 1)
Else
numlast = True
elements = elements + 1
commands(elements) = Mid(statement, i, 1)
End If
End If
Next i
linesnum = elements
End Sub
Private Sub dodrawing(startp As Integer, endp As Integer)
Dim startpp As Integer, endpp As Integer, depth As Integer
On Error Resume Next
i = startp
Do
If commands(i) = "fd" Then
Call drawline("fd", Val(commands(i + 1)))
End If
If commands(i) = "bk" Then
Call drawline("bk", Val(commands(i + 1)))
End If
If commands(i) = "rt" Then
Call turn("rt", Val(commands(i + 1)))
End If
If commands(i) = "lt" Then
Call turn("lt", Val(commands(i + 1)))
End If
If commands(i) = "pu" Then
pentype = "pu"
End If
If commands(i) = "pd" Then
pentype = "pd"
End If
If commands(i) = "hd" Then
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
End If
If commands(i) = "st" Then
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
End If

If commands(i) = "cr" Then
Picture1.Circle (xpos, ypos), Val(commands(i + 1))
End If
If commands(i) = "home" Then
angle = 270
xpos = 100
ypos = 75
drawarrow
End If
If commands(i) = "seth" Then
angle = 270
Call turn("rt", Val(commands(i + 1)))
drawarrow
End If
If commands(i) = "size" Then
Picture1.DrawWidth = Val(commands(i + 1))
End If

If commands(i) = "rp" Then
reptimes = Val(commands(i + 1))

depth = 1
startpp = i + 3

s = i + 3
Do
If commands(s) = "(" Then depth = depth + 1
If commands(s) = ")" Then depth = depth - 1
If depth = 0 Then
endpp = s
i = endpp
Exit Do
End If
DoEvents
s = s + 1
Loop
For g = 1 To reptimes
Call dodrawing(startpp, endpp)
Next

End If
DoEvents
i = i + 1
Loop Until i > endp

End Sub

Private Sub drawarrow()
deltax = 7 * Sin(rad(angle))
deltay = 7 * Cos(rad(angle))
point1x = xpos + deltax
point2x = xpos - deltax
point1y = ypos - deltay
point2y = ypos + deltay
point3x = xpos + 5 * Cos(rad(angle))
point3y = ypos + 5 * Sin(rad(angle))
Line1.X1 = point1x
Line1.Y1 = point1y
Line1.X2 = point2x
Line1.Y2 = point2y
Line2.X1 = point2x
Line2.Y1 = point2y
Line2.X2 = point3x
Line2.Y2 = point3y
Line3.X1 = point3x
Line3.Y1 = point3y
Line3.X2 = point1x
Line3.Y2 = point1y

End Sub
Private Sub drawline(direction As String, length As Double)
oldx = xpos
oldy = ypos
xx = length * Cos(rad(angle))
yy = length * Sin(rad(angle))
If direction = "fd" Then
xpos = xpos + xx
ypos = ypos + yy
Else
xpos = xpos - xx
ypos = ypos - yy
End If
If pentype = "pd" Then
Picture1.Line (oldx, oldy)-(xpos, ypos)
End If
drawarrow
End Sub
Private Sub turn(direction As String, value As Double)
If direction = "lt" Then
angle = angle - value
Else
angle = angle + value
End If
angle = angle Mod 360
drawarrow
End Sub
Private Function isnumber(a) As Boolean
If a = "0" Or a = "1" Or a = "2" Or a = "3" Or a = "4" Or a = "5" Or a = "6" Or a = "7" Or a = "8" Or a = "9" Or a = "." Then
isnumber = True
Else
isnumber = False
End If
End Function
Private Function rad(c) As Double
rad = c * pi / 180
End Function
Private Sub Command1_Click()
fillarray
Call dodrawing(1, linesnum)
List1.AddItem (Text1.Text)
Text1.Text = ""
End Sub




Private Sub Command2_Click()
Picture1.Cls
xpos = 100
ypos = 75
angle = 270
pentype = "pd"
drawarrow
List1.Clear
linesnum = 0
Picture1.DrawWidth = 1
End Sub

Private Sub Form_Initialize()
On Error Resume Next
Form1.Line (0, 0)-(2000, 2000)
For i = 1 To Val(Form1.Height) Step 15
Form1.Line (0, i)-(Val(Form1.Width), i), RGB(i * 255 / Val(Form1.Height) * 0.8, i * 255 / Val(Form1.Height), (i * 255 / Val(Form1.Height)) * 0.4)
Next

End Sub

Private Sub Form_Load()
xpos = 100
ypos = 75
angle = 270
pentype = "pd"
drawarrow

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
Form1.KeyPreview = True
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
