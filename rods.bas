Attribute VB_Name = "Module1"

Sub newgame()
Form1.Label3.Tag = Val(Form2.Text2.Text)
Form1.Tag = True
Form1.Label1.Tag = 0
For i = 0 To 35
Form1.Label2(i).BackColor = vbYellow
Form1.Label2(i).Enabled = True
Next i
Form1.Command1.Tag = Val(Form2.Text1.Text)
Form1.Label3.Caption = "Blue player turn"
If Val(Form2.Text1.Text) > 18 Then
Form1.Label1.Height = 6800
Else
Form1.Label1.Height = 3400
End If

End Sub

