VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ÇáÞÝÒ"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Tag             =   "5"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   8040
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   8160
      Width           =   735
   End
   Begin VB.Image Image48 
      Height          =   300
      Left            =   11040
      Picture         =   "ÝÞÒ.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   300
   End
   Begin VB.Image Image47 
      Height          =   300
      Left            =   2880
      Picture         =   "ÝÞÒ.frx":0182
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   300
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   6600
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   0
      Top             =   7560
      Width           =   5490
   End
   Begin VB.Image Image46 
      Height          =   255
      Left            =   4080
      Picture         =   "ÝÞÒ.frx":0304
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Image45 
      Height          =   255
      Left            =   7440
      Picture         =   "ÝÞÒ.frx":041E
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Image44 
      Height          =   255
      Left            =   840
      Picture         =   "ÝÞÒ.frx":0538
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   1005
      Left            =   0
      Tag             =   "2"
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image43 
      Height          =   255
      Left            =   3240
      Picture         =   "ÝÞÒ.frx":0652
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   6600
      Width           =   495
   End
   Begin VB.Image Image42 
      Height          =   255
      Left            =   6600
      Picture         =   "ÝÞÒ.frx":076C
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   6600
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   1005
      Left            =   10920
      Tag             =   "1"
      Top             =   7560
      Width           =   855
   End
   Begin VB.Image Image41 
      Height          =   255
      Left            =   10440
      Picture         =   "ÝÞÒ.frx":0886
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   6600
      Width           =   495
   End
   Begin VB.Image Image40 
      Height          =   150
      Left            =   7560
      Picture         =   "ÝÞÒ.frx":09A0
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   150
   End
   Begin VB.Image Image39 
      Height          =   150
      Left            =   10560
      Picture         =   "ÝÞÒ.frx":0B22
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   150
   End
   Begin VB.Image Image38 
      Height          =   150
      Left            =   9960
      Picture         =   "ÝÞÒ.frx":0CA4
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image Image37 
      Height          =   150
      Left            =   7080
      Picture         =   "ÝÞÒ.frx":0E26
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image Image36 
      Height          =   150
      Left            =   10800
      Picture         =   "ÝÞÒ.frx":0FA8
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   150
   End
   Begin VB.Image Image35 
      Height          =   150
      Left            =   11040
      Picture         =   "ÝÞÒ.frx":112A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   150
   End
   Begin VB.Image Image34 
      Height          =   150
      Left            =   10200
      Picture         =   "ÝÞÒ.frx":12AC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   150
   End
   Begin VB.Image Image33 
      Height          =   150
      Left            =   10920
      Picture         =   "ÝÞÒ.frx":142E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   150
   End
   Begin VB.Image Image32 
      Height          =   150
      Left            =   10920
      Picture         =   "ÝÞÒ.frx":15B0
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image Image31 
      Height          =   150
      Left            =   7560
      Picture         =   "ÝÞÒ.frx":1732
      Stretch         =   -1  'True
      Top             =   960
      Width           =   150
   End
   Begin VB.Image Image30 
      Height          =   150
      Left            =   240
      Picture         =   "ÝÞÒ.frx":18B4
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   150
   End
   Begin VB.Image Image29 
      Height          =   150
      Left            =   600
      Picture         =   "ÝÞÒ.frx":1A36
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   150
   End
   Begin VB.Image Image28 
      Height          =   150
      Left            =   1440
      Picture         =   "ÝÞÒ.frx":1BB8
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   150
   End
   Begin VB.Image Image27 
      Height          =   150
      Left            =   1920
      Picture         =   "ÝÞÒ.frx":1D3A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   150
   End
   Begin VB.Image Image26 
      Height          =   150
      Left            =   2520
      Picture         =   "ÝÞÒ.frx":1EBC
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   150
   End
   Begin VB.Image Image25 
      Height          =   150
      Left            =   720
      Picture         =   "ÝÞÒ.frx":203E
      Stretch         =   -1  'True
      Top             =   960
      Width           =   150
   End
   Begin VB.Image Image24 
      Height          =   150
      Left            =   3600
      Picture         =   "ÝÞÒ.frx":21C0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   150
   End
   Begin VB.Image Image23 
      Height          =   150
      Left            =   2160
      Picture         =   "ÝÞÒ.frx":2342
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   150
   End
   Begin VB.Image Image22 
      Height          =   150
      Left            =   360
      Picture         =   "ÝÞÒ.frx":24C4
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   150
   End
   Begin VB.Image Image21 
      Height          =   150
      Left            =   2640
      Picture         =   "ÝÞÒ.frx":2646
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   150
   End
   Begin VB.Image Image20 
      Height          =   150
      Left            =   3960
      Picture         =   "ÝÞÒ.frx":27C8
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   150
   End
   Begin VB.Image Image19 
      Height          =   150
      Left            =   5520
      Picture         =   "ÝÞÒ.frx":294A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   150
   End
   Begin VB.Image Image18 
      Height          =   150
      Left            =   3120
      Picture         =   "ÝÞÒ.frx":2ACC
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   150
   End
   Begin VB.Image Image17 
      Height          =   150
      Left            =   1560
      Picture         =   "ÝÞÒ.frx":2C4E
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image Image16 
      Height          =   150
      Left            =   3600
      Picture         =   "ÝÞÒ.frx":2DD0
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image Image15 
      Height          =   150
      Left            =   9120
      Picture         =   "ÝÞÒ.frx":2F52
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   150
   End
   Begin VB.Image Image14 
      Height          =   150
      Left            =   5880
      Picture         =   "ÝÞÒ.frx":30D4
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   150
   End
   Begin VB.Image Image13 
      Height          =   150
      Left            =   6360
      Picture         =   "ÝÞÒ.frx":3256
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   150
   End
   Begin VB.Image Image12 
      Height          =   150
      Left            =   9600
      Picture         =   "ÝÞÒ.frx":33D8
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   150
   End
   Begin VB.Image Image11 
      Height          =   150
      Left            =   6600
      Picture         =   "ÝÞÒ.frx":355A
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   150
   End
   Begin VB.Image Image10 
      Height          =   150
      Left            =   6960
      Picture         =   "ÝÞÒ.frx":36DC
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   150
   End
   Begin VB.Image Image9 
      Height          =   150
      Left            =   5400
      Picture         =   "ÝÞÒ.frx":385E
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   150
   End
   Begin VB.Image Image8 
      Height          =   150
      Left            =   8280
      Picture         =   "ÝÞÒ.frx":39E0
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   150
   End
   Begin VB.Image Image7 
      Height          =   150
      Left            =   8400
      Picture         =   "ÝÞÒ.frx":3B62
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   150
   End
   Begin VB.Image Image6 
      Height          =   150
      Left            =   4680
      Picture         =   "ÝÞÒ.frx":3CE4
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   150
      Left            =   8280
      Picture         =   "ÝÞÒ.frx":3E66
      Stretch         =   -1  'True
      Top             =   480
      Width           =   150
   End
   Begin VB.Image Image4 
      Height          =   150
      Left            =   4200
      Picture         =   "ÝÞÒ.frx":3FE8
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   150
   End
   Begin VB.Image Image3 
      Height          =   150
      Left            =   6000
      Picture         =   "ÝÞÒ.frx":416A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   8880
      Picture         =   "ÝÞÒ.frx":42EC
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   150
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Tag             =   "0"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   5160
      Picture         =   "ÝÞÒ.frx":446E
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   150
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   12000
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12000
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   500
      Left            =   1800
      Tag             =   "1"
      Top             =   6700
      Width           =   350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()
Select Case Text1.Text
Case "8"
Timer1.Tag = Shape1.Top
Text1.Tag = Text1.Text
Text1.Text = ""
Shape1.Tag = 1
Case "2"
If Shape1.Top < Form1.Height - 500 Then
Shape1.Top = Shape1.Top + 90
Text1.Tag = Text1.Text
Text1.Text = ""
Shape1.Tag = 1
End If
Case "4"
If Shape1.Left > 0 Then
Shape1.Left = Shape1.Left - 180
Text1.Tag = Text1.Text
Text1.Text = ""
Shape1.Tag = 1
End If

Case "6"
If Shape1.Left < Form1.Width - 400 Then
Shape1.Left = Shape1.Left + 180
Text1.Tag = Text1.Text
Text1.Text = ""
Shape1.Tag = 1
End If

Case "5"
Shape5.Visible = True
Text1.Text = ""
Shape5.Left = Shape1.Left + 125
Shape5.Top = Shape1.Top

End Select
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
If Shape2.Tag = 1 Then
Shape2.Top = Shape2.Top - 250
End If
If Shape2.Tag = 2 Then
Shape2.Top = Shape2.Top + 250
End If
If Shape2.Top > 8000 Then
Shape2.Tag = 1
End If
If Shape2.Top < 0 Then
Shape2.Tag = 2
End If
If Shape3.Tag = 1 Then
Shape3.Top = Shape3.Top - 250
End If
If Shape3.Tag = 2 Then
Shape3.Top = Shape3.Top + 250
End If
If Shape3.Top > 8000 Then
Shape3.Tag = 1
End If
If Shape3.Top < 0 Then
Shape3.Tag = 2
End If

Text1.SetFocus
Select Case Text1.Tag
Case "8"
If Shape1.Tag = "1" Then
Shape1.Top = Shape1.Top - 90
End If

If Shape1.Tag = "2" And ((Shape1.Top + 500) / 1800 <> Int((Shape1.Top + 500) / 1800)) Then
Shape1.Top = Shape1.Top + 90
End If
If Shape1.Top <= Val(Timer1.Tag) - 1800 Then
Shape1.Tag = "2"
End If
End Select

If (Shape1.Top + 500) / 1800 <> Int((Shape1.Top + 500) / 1800) Then
Shape1.Top = Shape1.Top + 30
End If
If Shape1.Top < Image1.Top + 150 And Shape1.Top > Image1.Top - 500 And Shape1.Left > Image1.Left - 500 And Shape1.Left < Image1.Left + 150 And Image1.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image1.Visible = False
End If
If Shape1.Top < Image2.Top + 150 And Shape1.Top > Image2.Top - 500 And Shape1.Left > Image2.Left - 500 And Shape1.Left < Image2.Left + 150 And Image2.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image2.Visible = False
End If
If Shape1.Top < Image3.Top + 150 And Shape1.Top > Image3.Top - 500 And Shape1.Left > Image3.Left - 500 And Shape1.Left < Image3.Left + 150 And Image3.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image3.Visible = False
End If
If Shape1.Top < Image4.Top + 150 And Shape1.Top > Image4.Top - 500 And Shape1.Left > Image4.Left - 500 And Shape1.Left < Image4.Left + 150 And Image4.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image4.Visible = False
End If
If Shape1.Top < Image5.Top + 150 And Shape1.Top > Image5.Top - 500 And Shape1.Left > Image5.Left - 500 And Shape1.Left < Image5.Left + 150 And Image5.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image5.Visible = False
End If

If Shape1.Top < Image6.Top + 150 And Shape1.Top > Image6.Top - 500 And Shape1.Left > Image6.Left - 500 And Shape1.Left < Image6.Left + 150 And Image6.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image6.Visible = False
End If

If Shape1.Top < Image7.Top + 150 And Shape1.Top > Image7.Top - 500 And Shape1.Left > Image7.Left - 500 And Shape1.Left < Image7.Left + 150 And Image7.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image7.Visible = False
End If

If Shape1.Top < Image8.Top + 150 And Shape1.Top > Image8.Top - 500 And Shape1.Left > Image8.Left - 500 And Shape1.Left < Image8.Left + 150 And Image8.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image8.Visible = False
End If

If Shape1.Top < Image9.Top + 150 And Shape1.Top > Image9.Top - 500 And Shape1.Left > Image9.Left - 500 And Shape1.Left < Image9.Left + 150 And Image9.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image9.Visible = False
End If

If Shape1.Top < Image10.Top + 150 And Shape1.Top > Image10.Top - 500 And Shape1.Left > Image10.Left - 500 And Shape1.Left < Image10.Left + 150 And Image10.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image10.Visible = False
End If

If Shape1.Top < Image11.Top + 150 And Shape1.Top > Image11.Top - 500 And Shape1.Left > Image11.Left - 500 And Shape1.Left < Image11.Left + 150 And Image11.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image11.Visible = False
End If

If Shape1.Top < Image12.Top + 150 And Shape1.Top > Image12.Top - 500 And Shape1.Left > Image12.Left - 500 And Shape1.Left < Image12.Left + 150 And Image12.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image12.Visible = False
End If
If Shape1.Top < Image13.Top + 150 And Shape1.Top > Image13.Top - 500 And Shape1.Left > Image13.Left - 500 And Shape1.Left < Image13.Left + 150 And Image13.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image13.Visible = False
End If

If Shape1.Top < Image14.Top + 150 And Shape1.Top > Image14.Top - 500 And Shape1.Left > Image14.Left - 500 And Shape1.Left < Image14.Left + 150 And Image14.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image14.Visible = False
Label1.Tag = 1
End If

If Shape1.Top < Image15.Top + 150 And Shape1.Top > Image15.Top - 500 And Shape1.Left > Image15.Left - 500 And Shape1.Left < Image15.Left + 150 And Image15.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image15.Visible = False
End If

If Shape1.Top < Image16.Top + 150 And Shape1.Top > Image16.Top - 500 And Shape1.Left > Image16.Left - 500 And Shape1.Left < Image16.Left + 150 And Image16.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image16.Visible = False
End If

If Shape1.Top < Image17.Top + 150 And Shape1.Top > Image17.Top - 500 And Shape1.Left > Image17.Left - 500 And Shape1.Left < Image17.Left + 150 And Image17.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image17.Visible = False
End If

If Shape1.Top < Image18.Top + 150 And Shape1.Top > Image18.Top - 500 And Shape1.Left > Image18.Left - 500 And Shape1.Left < Image18.Left + 150 And Image18.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image18.Visible = False
End If

If Shape1.Top < Image19.Top + 150 And Shape1.Top > Image19.Top - 500 And Shape1.Left > Image19.Left - 500 And Shape1.Left < Image19.Left + 150 And Image19.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image19.Visible = False
End If

If Shape1.Top < Image20.Top + 150 And Shape1.Top > Image20.Top - 500 And Shape1.Left > Image20.Left - 500 And Shape1.Left < Image20.Left + 150 And Image20.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image20.Visible = False
End If

If Shape1.Top < Image21.Top + 150 And Shape1.Top > Image21.Top - 500 And Shape1.Left > Image21.Left - 500 And Shape1.Left < Image21.Left + 150 And Image21.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image21.Visible = False
End If

If Shape1.Top < Image22.Top + 150 And Shape1.Top > Image22.Top - 500 And Shape1.Left > Image22.Left - 500 And Shape1.Left < Image22.Left + 150 And Image22.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image22.Visible = False
End If

If Shape1.Top < Image23.Top + 150 And Shape1.Top > Image23.Top - 500 And Shape1.Left > Image23.Left - 500 And Shape1.Left < Image23.Left + 150 And Image23.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image23.Visible = False
End If

If Shape1.Top < Image24.Top + 150 And Shape1.Top > Image24.Top - 500 And Shape1.Left > Image24.Left - 500 And Shape1.Left < Image24.Left + 150 And Image24.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image24.Visible = False
End If

If Shape1.Top < Image25.Top + 150 And Shape1.Top > Image25.Top - 500 And Shape1.Left > Image25.Left - 500 And Shape1.Left < Image25.Left + 150 And Image25.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image25.Visible = False
End If

If Shape1.Top < Image26.Top + 150 And Shape1.Top > Image26.Top - 500 And Shape1.Left > Image26.Left - 500 And Shape1.Left < Image26.Left + 150 And Image26.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image26.Visible = False
End If

If Shape1.Top < Image27.Top + 150 And Shape1.Top > Image27.Top - 500 And Shape1.Left > Image27.Left - 500 And Shape1.Left < Image27.Left + 150 And Image27.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image27.Visible = False
End If

If Shape1.Top < Image28.Top + 150 And Shape1.Top > Image28.Top - 500 And Shape1.Left > Image28.Left - 500 And Shape1.Left < Image28.Left + 150 And Image28.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image28.Visible = False
End If

If Shape1.Top < Image29.Top + 150 And Shape1.Top > Image29.Top - 500 And Shape1.Left > Image29.Left - 500 And Shape1.Left < Image29.Left + 150 And Image29.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image29.Visible = False
End If

If Shape1.Top < Image30.Top + 150 And Shape1.Top > Image30.Top - 500 And Shape1.Left > Image30.Left - 500 And Shape1.Left < Image30.Left + 150 And Image30.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image30.Visible = False
End If

If Shape1.Top < Image31.Top + 150 And Shape1.Top > Image31.Top - 500 And Shape1.Left > Image31.Left - 500 And Shape1.Left < Image31.Left + 150 And Image31.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image31.Visible = False
End If

If Shape1.Top < Image32.Top + 150 And Shape1.Top > Image32.Top - 500 And Shape1.Left > Image32.Left - 500 And Shape1.Left < Image32.Left + 150 And Image32.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image32.Visible = False
End If

If Shape1.Top < Image33.Top + 150 And Shape1.Top > Image33.Top - 500 And Shape1.Left > Image33.Left - 500 And Shape1.Left < Image33.Left + 150 And Image33.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image33.Visible = False
End If

If Shape1.Top < Image34.Top + 150 And Shape1.Top > Image34.Top - 500 And Shape1.Left > Image34.Left - 500 And Shape1.Left < Image34.Left + 150 And Image34.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image34.Visible = False
End If

If Shape1.Top < Image35.Top + 150 And Shape1.Top > Image35.Top - 500 And Shape1.Left > Image35.Left - 500 And Shape1.Left < Image35.Left + 150 And Image35.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image35.Visible = False
End If

If Shape1.Top < Image36.Top + 150 And Shape1.Top > Image36.Top - 500 And Shape1.Left > Image36.Left - 500 And Shape1.Left < Image36.Left + 150 And Image36.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image36.Visible = False
End If

If Shape1.Top < Image37.Top + 150 And Shape1.Top > Image37.Top - 500 And Shape1.Left > Image37.Left - 500 And Shape1.Left < Image37.Left + 150 And Image37.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image37.Visible = False
End If

If Shape1.Top < Image38.Top + 150 And Shape1.Top > Image38.Top - 500 And Shape1.Left > Image38.Left - 500 And Shape1.Left < Image38.Left + 150 And Image38.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image38.Visible = False
End If

If Shape1.Top < Image39.Top + 150 And Shape1.Top > Image39.Top - 500 And Shape1.Left > Image39.Left - 500 And Shape1.Left < Image39.Left + 150 And Image39.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image39.Visible = False
End If

If Shape1.Top < Image40.Top + 150 And Shape1.Top > Image40.Top - 500 And Shape1.Left > Image40.Left - 500 And Shape1.Left < Image40.Left + 150 And Image40.Visible = True Then
Label1.Caption = Val(Label1.Caption) + 10
Image40.Visible = False
End If
Image41.Left = Image41.Left - 200
If Image41.Left <= 0 Then
Image41.Left = 10500
Image41.Top = Shape2.Top + 500
End If
Image42.Left = Image42.Left - 200
If Image42.Left <= 0 Then
Image42.Left = 10500
Image42.Top = Shape2.Top + 500
End If
Image43.Left = Image43.Left - 200
If Image43.Left <= 0 Then
Image43.Left = 10500
Image43.Top = Shape2.Top + 500
End If
Image44.Left = Image44.Left + 200
If Image44.Left >= 11000 Then
Image44.Left = 700
Image44.Top = Shape3.Top + 500
End If
Image45.Left = Image45.Left + 200
If Image45.Left >= 11000 Then
Image45.Left = 700
Image45.Top = Shape3.Top + 500
End If
Image46.Left = Image46.Left + 200
If Image46.Left >= 11000 Then
Image46.Left = 700
Image46.Top = Shape3.Top + 500
End If

On Error Resume Next
If Shape1.Top < Image41.Top + 250 And Shape1.Top > Image41.Top - 500 And Shape1.Left > Image41.Left - 500 And Shape1.Left < Image41.Left + 495 And Image41.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If
If Shape1.Top < Image42.Top + 250 And Shape1.Top > Image42.Top - 500 And Shape1.Left > Image42.Left - 500 And Shape1.Left < Image42.Left + 495 And Image42.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If
If Shape1.Top < Image43.Top + 250 And Shape1.Top > Image43.Top - 500 And Shape1.Left > Image43.Left - 500 And Shape1.Left < Image43.Left + 495 And Image43.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If
If Shape1.Top < Image44.Top + 250 And Shape1.Top > Image44.Top - 500 And Shape1.Left > Image44.Left - 500 And Shape1.Left < Image44.Left + 495 And Image44.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If
If Shape1.Top < Image45.Top + 250 And Shape1.Top > Image45.Top - 500 And Shape1.Left > Image45.Left - 500 And Shape1.Left < Image45.Left + 495 And Image45.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If
If Shape1.Top < Image46.Top + 250 And Shape1.Top > Image46.Top - 500 And Shape1.Left > Image46.Left - 500 And Shape1.Left < Image46.Left + 495 And Image46.Visible = True Then
Shape4.Width = Shape4.Width - 90
End If

If Shape5.Visible = True Then
Shape5.Top = Shape5.Top - 160
End If
If Shape5.Top < 0 Then
Shape5.Top = Shape1.Top
Shape5.Left = Shape1.Left + 125
Shape5.Visible = False
End If


If Shape5.Top < Image41.Top + 250 And Shape1.Top > Image41.Top - 100 And Shape5.Left > Image41.Left - 100 And Shape5.Left < Image41.Left + 495 And Image41.Visible = True And Shape5.Visible = True Then
Image41.Tag = Val(Image41.Tag - 1)
Shape5.Visible = False
End If
If Image41.Tag = 0 Then
Image41.Visible = False
End If
If Shape5.Top < Image42.Top + 250 And Shape1.Top > Image42.Top - 100 And Shape5.Left > Image42.Left - 100 And Shape5.Left < Image42.Left + 495 And Image42.Visible = True And Shape5.Visible = True Then
Image42.Tag = Val(Image42.Tag - 1)
Shape5.Visible = False
End If
If Image42.Tag = 0 Then
Image42.Visible = False
End If
If Shape5.Top < Image43.Top + 250 And Shape1.Top > Image43.Top - 100 And Shape5.Left > Image43.Left - 100 And Shape5.Left < Image43.Left + 495 And Image43.Visible = True And Shape5.Visible = True Then
Image43.Tag = Val(Image43.Tag - 1)
Shape5.Visible = False
End If
If Image43.Tag = 0 Then
Image43.Visible = False
End If
If Shape5.Top < Image44.Top + 250 And Shape1.Top > Image44.Top - 100 And Shape5.Left > Image44.Left - 100 And Shape5.Left < Image44.Left + 495 And Image44.Visible = True And Shape5.Visible = True Then
Image44.Tag = Val(Image44.Tag - 1)
Shape5.Visible = False
End If
If Image44.Tag = 0 Then
Image44.Visible = False
End If
If Shape5.Top < Image45.Top + 250 And Shape1.Top > Image45.Top - 100 And Shape5.Left > Image45.Left - 100 And Shape5.Left < Image45.Left + 495 And Image45.Visible = True And Shape5.Visible = True Then
Image45.Tag = Val(Image45.Tag - 1)
Shape5.Visible = False
End If
If Image45.Tag = 0 Then
Image45.Visible = False
End If
If Shape5.Top < Image46.Top + 250 And Shape1.Top > Image46.Top - 100 And Shape5.Left > Image46.Left - 100 And Shape5.Left < Image46.Left + 495 And Image46.Visible = True And Shape5.Visible = True Then
Image46.Tag = Val(Image46.Tag - 1)
Shape5.Visible = False
End If
If Image46.Tag = 0 Then
Image46.Visible = False
End If


If Shape1.Top < Image47.Top + 300 And Shape1.Top > Image47.Top - 500 And Shape1.Left > Image47.Left - 500 And Shape1.Left < Image47.Left + 300 And Image47.Visible = True Then
Shape4.Width = Shape4.Width + 810
Image47.Visible = False
End If
If Shape1.Top < Image48.Top + 300 And Shape1.Top > Image48.Top - 500 And Shape1.Left > Image48.Left - 500 And Shape1.Left < Image48.Left + 300 And Image48.Visible = True Then
Shape4.Width = Shape4.Width + 810
Image48.Visible = False
End If
If Shape4.Width <= 20 Then
End
End If

If Val(Label1.Caption) / 400 = Int(Val(Label1.Caption) / 400) And Label1.Tag = 1 Then
d = Label1.Caption
Unload Me
Load Form1
Form1.Show
Label1.Caption = d
Label1.Tag = 0
End If
End Sub
