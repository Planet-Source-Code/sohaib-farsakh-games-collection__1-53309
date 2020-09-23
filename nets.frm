VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Form1.Cls
Form1.Scale (-10, 10)-(10, -10)
For i = -10 To 10 Step 0.5
Form1.Line (-10, i)-(10, i)
Form1.Line (i - 1, -10)-(i + 1, 10)
Next i

End Sub

