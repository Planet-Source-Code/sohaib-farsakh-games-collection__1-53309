VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   DrawWidth       =   7
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub newgame()
Form1.Cls
Form1.ForeColor = vbBlue
For i = 1 To 7
For j = 1 To 7
a = i * 900 + 1000
b = j * 900 + 500
Form1.PSet (a, b)
Next j
Next i

End Sub


Private Sub Form_Load()
newgame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X - 1000) Mod 900 < 150 Or (X - 1000) Mod 900 > 750 Then
If (X - 1000) Mod 900 < 150 Then
form1.line
End If

End Sub
