VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "equations graphs"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CommandButton Command5 
      Caption         =   "minor dividers"
      Height          =   615
      Left            =   9600
      TabIndex        =   37
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "major dividers"
      Height          =   615
      Left            =   9600
      TabIndex        =   36
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   11040
      TabIndex        =   35
      Text            =   "i"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   10320
      TabIndex        =   34
      Text            =   "h"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   9600
      TabIndex        =   33
      Text            =   "g"
      Top             =   3960
      Width           =   735
   End
   Begin VB.OptionButton Option7 
      Caption         =   "y = a^x"
      Height          =   195
      Left            =   9480
      TabIndex        =   32
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton Option6 
      Caption         =   "y=a[bx+c]+d"
      Height          =   195
      Left            =   9480
      TabIndex        =   31
      Top             =   480
      Width           =   1815
   End
   Begin VB.OptionButton Option5 
      Caption         =   "y=a(Abs(bx+c))+d"
      Height          =   195
      Left            =   9480
      TabIndex        =   30
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      Caption         =   "y=a(Sin(bx+c))+d"
      Height          =   195
      Left            =   9480
      TabIndex        =   29
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "y=a(Cos(bx+c))+d"
      Height          =   195
      Left            =   9480
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "y=a(Tan(bx+c))+d"
      Height          =   195
      Left            =   9480
      TabIndex        =   27
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "y= ax^5 + bx^4 + cx^3 + dx^2 +ex + f+gx^-1 +hx^-2 + ix^-3"
      Height          =   615
      Left            =   9480
      TabIndex        =   26
      Top             =   2280
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8880
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "functions.frx":0000
      Left            =   9840
      List            =   "functions.frx":0037
      TabIndex        =   10
      Text            =   "100"
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      Height          =   615
      Left            =   9600
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   11040
      TabIndex        =   7
      Text            =   "f"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Text            =   "e"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Text            =   "d"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Text            =   "c"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Text            =   "b"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Text            =   "a"
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "draw"
      Height          =   615
      Left            =   9600
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16711680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "pick color"
      Height          =   615
      Left            =   9600
      TabIndex        =   0
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   195
      Left            =   8760
      TabIndex        =   25
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "-50"
      Height          =   195
      Left            =   4680
      TabIndex        =   24
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "-75"
      Height          =   195
      Left            =   4680
      TabIndex        =   23
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      Height          =   195
      Left            =   7680
      TabIndex        =   22
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   195
      Left            =   4560
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      Height          =   195
      Left            =   4560
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "-50"
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "-75"
      Height          =   195
      Left            =   960
      TabIndex        =   17
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "-100"
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   195
      Left            =   6600
      TabIndex        =   15
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "-25"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "-25"
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      Height          =   195
      Left            =   4560
      TabIndex        =   12
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      Height          =   200
      Left            =   5520
      TabIndex        =   11
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "zoom                          %"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   1122
      Y2              =   1125
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      X1              =   4400
      X2              =   4600
      Y1              =   6750
      Y2              =   6750
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   7875
      Y2              =   7875
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   4520
      X2              =   4320
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   4600
      X2              =   4400
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   0
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   3375
      X2              =   3375
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   2250
      X2              =   2250
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   1125
      X2              =   1125
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   9000
      X2              =   9000
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   6750
      X2              =   6750
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   7875
      X2              =   7875
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   5625
      X2              =   5625
      Y1              =   4600
      Y2              =   4400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4500
      X2              =   4500
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   9000
      Y1              =   4500
      Y2              =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CommonDialog1.Flags = 1
CommonDialog1.ShowColor
Form1.ForeColor = CommonDialog1.Color
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Option1.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 1000)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 1000)
r = Val(Text1.Text) * a ^ 5 + Val(Text2.Text) * a ^ 4 + Val(Text3.Text) * a ^ 3 + Val(Text4.Text) * a ^ 2 + Val(Text5.Text) * a + Val(Text6.Text) + Val(Text7.Text) * a ^ (-1) + Val(Text8.Text) * a ^ (-2) + Val(Text9.Text) * a ^ (-3)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) * t ^ 5 + Val(Text2.Text) * t ^ 4 + Val(Text3.Text) * t ^ 3 + Val(Text4.Text) * t ^ 2 + Val(Text5.Text) * t + Val(Text6.Text) + Val(Text7.Text) * t ^ (-1) + Val(Text8.Text) * t ^ (-2) + Val(Text9.Text) * t ^ (-3)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If



If Option2.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) * 3.14159265358979 / 30000)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) * 3.14159265358979 / 30000)
r = Val(Text1.Text) * (Tan(Val(Text2.Text) * a + Val(Text3.Text))) + Val(Text4.Text)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) * (Tan(Val(Text2.Text) * t + Val(Text3.Text))) + Val(Text4.Text)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If

If Option3.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 400)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 400)
r = Val(Text1.Text) * (Cos(Val(Text2.Text) * a + Val(Text3.Text))) + Val(Text4.Text)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) * (Cos(Val(Text2.Text) * t + Val(Text3.Text))) + Val(Text4.Text)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If

If Option4.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 400)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 400)
r = Val(Text1.Text) * (Sin(Val(Text2.Text) * a + Val(Text3.Text))) + Val(Text4.Text)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) * (Sin(Val(Text2.Text) * t + Val(Text3.Text))) + Val(Text4.Text)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If


If Option5.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 100)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 100)
r = Val(Text2.Text) * a + Val(Text3.Text)
If r < 0 Then
r = r * -1
End If
r = Val(Text1.Text) * r + Val(Text4.Text)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text2.Text) * t + Val(Text3.Text)
If b < 0 Then
b = b * -1
End If
b = Val(Text1.Text) * b + Val(Text4.Text)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If

If Option6.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 100)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 100)
r = Val(Text1.Text) * Int(Val(Text2.Text) * a + Val(Text3.Text)) + Val(Text4.Text)
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) * Int(Val(Text2.Text) * t + Val(Text3.Text)) + Val(Text4.Text)
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - r
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If

On Error Resume Next
If Option7.Value = True Then
For i = 0 To 2 * Val(Label17.Caption) Step (Val(Label17.Caption) / 2000)
a = i - Val(Label17.Caption)
t = a - (Val(Label17.Caption) / 2000)
r = Val(Text1.Text) ^ a
r = r + Val(Label17.Caption)
r = r * 4500 / Val(Label17.Caption)
y = Form1.Height - r
b = Val(Text1.Text) ^ t
b = b + Val(Label17.Caption)
b = b * 4500 / Val(Label17.Caption)
c = Form1.Height - b
t = t * 4500 / Val(Label17.Caption)
a = a * 4500 / Val(Label17.Caption)
t = t + 4500
a = a + 4500
If t < 9000 And t > 0 And c < 9000 And c > 0 And a < 9000 And a > 0 And y < 9000 And y > 0 Then
Form1.Line (t, c)-(a, y)
End If
Next i
End If
End Sub

Private Sub Command3_Click()

Form1.Width = 0
Form1.Height = 0
Form1.Width = 12000
Form1.Height = 9000
End Sub

Private Sub Command4_Click()
For x = 0 To 9000 Step 1125
If x <> 4500 Then
Form1.Line (0, x)-(9000, x)
Form1.Line (x, 0)-(x, 9000)
End If
Next x
End Sub

Private Sub Command5_Click()
For x = 0 To 9000 Step 225
If x <> 4500 Then
Form1.ForeColor = &HFFFFFF
Form1.Line (0, x)-(9000, x)
Form1.Line (x, 0)-(x, 9000)
End If
Next x
Form1.ForeColor = CommonDialog1.Color
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label3.Caption = 2500 / Val(Combo1.Text)
Label4.Caption = 2500 / Val(Combo1.Text)
Label5.Caption = -2500 / Val(Combo1.Text)
Label6.Caption = -2500 / Val(Combo1.Text)
Label7.Caption = 5000 / Val(Combo1.Text)
Label13.Caption = 5000 / Val(Combo1.Text)
Label10.Caption = -5000 / Val(Combo1.Text)
Label16.Caption = -5000 / Val(Combo1.Text)
Label14.Caption = 7500 / Val(Combo1.Text)
Label12.Caption = 7500 / Val(Combo1.Text)
Label9.Caption = -7500 / Val(Combo1.Text)
Label15.Caption = -7500 / Val(Combo1.Text)
Label17.Caption = 10000 / Val(Combo1.Text)
Label11.Caption = 10000 / Val(Combo1.Text)
Label8.Caption = -10000 / Val(Combo1.Text)
End Sub
