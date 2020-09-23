Attribute VB_Name = "Module1"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public mazenum As Integer
Public lengthincrease As Integer
Public remainingapples As Integer
Public direction As String, lastdir As String
Public snakelength As Integer, targetlength As Integer
Public snakeparts(1000, 1) As Integer
Public foodx As Integer, foody As Integer
Public mazeblocks(150, 1) As Integer
Public blocksnum As Integer
Public foodcolor As String
Public score As Integer
Public level As Integer
Public password(2 To 10) As String * 4
Public livescore As Integer
Public multiplicator As Integer
Public scorescomplete(0 To 9, 1 To 2)
Public scorescrash(0 To 9, 1 To 2)
Public gamerun As Boolean
Public s_interval As Integer
Public pictureindex As Integer
Public framecolor As Variant



Sub gotolevel(levelnum As Integer)
Select Case levelnum
Case 1
s_interval = 250
mazenum = 0
lengthincrease = 2
remainingapples = 5
multiplicator = 2
Case 2
s_interval = 150
mazenum = 0
lengthincrease = 3
remainingapples = 7
multiplicator = 3
Case 3
s_interval = 100
mazenum = 1
lengthincrease = 3
remainingapples = 10
multiplicator = 4
Case 4
s_interval = 75
mazenum = 1
lengthincrease = 5
remainingapples = 10
multiplicator = 6
Case 5
s_interval = 60
mazenum = 2
lengthincrease = 5
remainingapples = 15
multiplicator = 8
Case 6
s_interval = 45
mazenum = 2
lengthincrease = 7
remainingapples = 15
multiplicator = 11
Case 7
s_interval = 55
mazenum = 3
lengthincrease = 10
remainingapples = 15
multiplicator = 15
Case 8
s_interval = 40
mazenum = 3
lengthincrease = 12
remainingapples = 15
multiplicator = 20
Case 9
s_interval = 70
mazenum = 4
lengthincrease = 12
remainingapples = 12
multiplicator = 30
Case 10
s_interval = 40
mazenum = 4
lengthincrease = 18
remainingapples = 15
multiplicator = 40
End Select
newgame
frmMain.Label5.Caption = Str(levelnum)
frmMain.Label8.Caption = Str(remainingapples)
End Sub
Sub entername()
Dim search As Integer
d = InputBox("Your score is in the highscores list,enter your name", "Enter your name")
If frmMain.crash.Checked = True Then
search = 8
rank = 9
Do
If score > Val(scorescrash(search, 2)) Then
rank = search
Else
Exit Do
End If
search = search - 1
If search < 0 Then Exit Do
Loop
For i = 9 To rank + 1 Step -1
scorescrash(i, 1) = scorescrash(i - 1, 1)
scorescrash(i, 2) = scorescrash(i - 1, 2)
Next
scorescrash(rank, 1) = d
scorescrash(rank, 2) = score
Else
search = 8
rank = 9
Do
If score > Val(scorescomplete(search, 2)) Then
rank = search
Else
Exit Do
End If
search = search - 1
If search < 0 Then Exit Do
Loop
For i = 9 To rank + 1 Step -1
scorescomplete(i, 1) = scorescomplete(i - 1, 1)
scorescomplete(i, 2) = scorescomplete(i - 1, 2)
Next
scorescomplete(rank, 1) = d
scorescomplete(rank, 2) = score

End If

frmscores.Show , frmMain
If frmMain.complete.Checked = True Then
frmscores.Option1.Value = True
displayscores ("complete")
Else
frmscores.Option2.Value = True
displayscores ("crash")
End If

End Sub
Sub newgame()
frmMain.Picture1.Cls
drawframe
For i = 0 To 5
snakeparts(i, 0) = 7
snakeparts(i, 1) = 15 - i
Next
direction = "right"
lastdir = "right"
snakelength = 6
targetlength = 6
drawsnake
blocksnum = -1
applymaze (mazenum)
gamerun = False
frmMain.Label1.Caption = "Start"
frmMain.Label1.Enabled = "true"
putfood
If frmMain.crash.Checked = True Then
score = 0
frmMain.Label3.Caption = "0"
End If
End Sub
Public Sub drawframe()
frmMain.Picture1.ForeColor = framecolor
frmMain.Picture1.DrawWidth = 5
frmMain.Picture1.Line (0.7, 0.7)-(0.7, 26.3)
frmMain.Picture1.Line (0.7, 0.7)-(26.3, 0.7)
frmMain.Picture1.Line (26.3, 26.2)-(26.3, 0.7)
frmMain.Picture1.Line (0.7, 26.3)-(26.3, 26.3)
End Sub
Sub drawsnake()
drawframe
For i = 1 To snakelength - 1
frmMain.Picture1.PaintPicture frmPic.Picture1(4).Picture, snakeparts(i, 1), snakeparts(i, 0), 1, 1
Next i
Select Case direction
Case "up"
frmMain.Picture1.PaintPicture frmPic.Picture1(0).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
Case "down"
frmMain.Picture1.PaintPicture frmPic.Picture1(1).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
Case "right"
frmMain.Picture1.PaintPicture frmPic.Picture1(2).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
Case "left"
frmMain.Picture1.PaintPicture frmPic.Picture1(3).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
End Select
End Sub
Public Sub applymaze(mazenum As Integer)
Select Case mazenum
Case 0
blocksnum = -1
Case 1
For i = 5 To 21
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 17, 1, 1
mazeblocks(i - 5, 0) = 17
mazeblocks(i - 5, 1) = i
Next
blocksnum = 16 'from 0 to 16
Case 2
blocksnum = 0
For i = 7 To 19
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 13, 1, 1
mazeblocks(blocksnum, 0) = 13
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 7 To 19
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, 6, i, 1, 1
mazeblocks(blocksnum, 0) = i
mazeblocks(blocksnum, 1) = 6
blocksnum = blocksnum + 1
Next
For i = 7 To 19
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, 20, i, 1, 1
mazeblocks(blocksnum, 0) = i
mazeblocks(blocksnum, 1) = 20
blocksnum = blocksnum + 1
Next
blocksnum = blocksnum - 1
Case 3
blocksnum = 0
For i = 5 To 21
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, 5, i, 1, 1
mazeblocks(blocksnum, 0) = i
mazeblocks(blocksnum, 1) = 5
blocksnum = blocksnum + 1
Next
For i = 5 To 21
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, 21, i, 1, 1
mazeblocks(blocksnum, 0) = i
mazeblocks(blocksnum, 1) = 21
blocksnum = blocksnum + 1
Next
For i = 6 To 12
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 5, 1, 1
mazeblocks(blocksnum, 0) = 5
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 15 To 20
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 5, 1, 1
mazeblocks(blocksnum, 0) = 5
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 6 To 11
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 21, 1, 1
mazeblocks(blocksnum, 0) = 21
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 14 To 20
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 21, 1, 1
mazeblocks(blocksnum, 0) = 21
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 9 To 17
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 10, 1, 1
mazeblocks(blocksnum, 0) = 10
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
For i = 9 To 17
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, i, 16, 1, 1
mazeblocks(blocksnum, 0) = 16
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
blocksnum = blocksnum - 1
Case 4
blocksnum = 0
For i = 2 To 24 Step 2
For j = 2 To 24 Step 2
frmMain.Picture1.PaintPicture frmPic.Picture1(8).Picture, j, i, 1, 1
mazeblocks(blocksnum, 0) = j
mazeblocks(blocksnum, 1) = i
blocksnum = blocksnum + 1
Next
Next
blocksnum = blocksnum - 1
End Select
End Sub
Sub putfood()
Dim overthing As Boolean
Randomize
a = Int(Rnd * 25) + 1
If a = 1 Then
pictureindex = 9
foodcolor = "bomb"
Else
d = Int(Rnd * 7) + 1
Select Case d
Case 1 To 4
pictureindex = 5
foodcolor = "red"
Case 5 To 6
pictureindex = 6
foodcolor = "blue"
Case 7
pictureindex = 7
foodcolor = "yellow"
End Select
End If
again:
a = Int(Rnd * 25)
b = Int(Rnd * 25)
foodx = a + 1
foody = b + 1
overthing = False
For i = 0 To snakelength - 1
If Abs(snakeparts(i, 0) - foody) < 0.1 And Abs(snakeparts(i, 1) - foodx) < 0.1 Then
overthing = True
Exit For
End If
Next
For i = 0 To blocksnum
If Abs(mazeblocks(i, 0) - foody) < 0.1 And Abs(mazeblocks(i, 1) - foodx) < 0.1 Then
overthing = True
Exit For
End If
Next

If overthing = True Then GoTo again
frmMain.Picture1.PaintPicture frmPic.Picture1(pictureindex).Picture, foodx, foody, 1, 1
End Sub

Sub displayscores(gametype As String)
Select Case gametype
Case "complete"
For i = 0 To 9
frmscores.Label1(i).Caption = scorescomplete(i, 1)
frmscores.Label2(i).Caption = scorescomplete(i, 2)
Next
Case "crash"
For i = 0 To 9
frmscores.Label1(i).Caption = scorescrash(i, 1)
frmscores.Label2(i).Caption = scorescrash(i, 2)
Next
End Select


End Sub
Sub mini_delay(lngDelayAmt As Long)

  Dim lngNewTime   As Long
  Dim lngCurrent   As Long
  
  lngNewTime = GetTickCount + lngDelayAmt
  
  Do
      lngCurrent = GetTickCount      ' get the current millisecond count
      DoEvents
      
      If lngCurrent >= lngNewTime Then
          Exit Do
      End If
  Loop

End Sub
Public Sub movesnake()
Dim increase As Boolean, crashed As Boolean
If targetlength > snakelength Then
snakelength = snakelength + 1
increase = True
Else
increase = False
End If
If increase = False Then
a = snakeparts(snakelength - 1, 1)
b = snakeparts(snakelength - 1, 0)
'Call clearcell(a, b)
frmMain.Picture1.ForeColor = vbBlack
frmMain.Picture1.FillColor = vbBlack
frmMain.Picture1.DrawWidth = 1
frmMain.Picture1.Line (a, b)-(a + 0.92, b + 0.92), cr, BF

End If




For i = (snakelength - 1) To 1 Step -1
snakeparts(i, 0) = snakeparts(i - 1, 0)
snakeparts(i, 1) = snakeparts(i - 1, 1)
Next i


Select Case direction
Case "up"
snakeparts(0, 0) = snakeparts(0, 0) - 1
If frmMain.gowall.Checked = True And snakeparts(0, 0) = 0 Then
snakeparts(0, 0) = 25
End If
frmMain.Picture1.PaintPicture frmPic.Picture1(0).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
lastdir = "up"
Case "down"
snakeparts(0, 0) = snakeparts(0, 0) + 1
If frmMain.gowall.Checked = True And snakeparts(0, 0) = 26 Then
snakeparts(0, 0) = 1
End If
frmMain.Picture1.PaintPicture frmPic.Picture1(1).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
lastdir = "down"
Case "right"
snakeparts(0, 1) = snakeparts(0, 1) + 1
If frmMain.gowall.Checked = True And snakeparts(0, 1) = 26 Then
snakeparts(0, 1) = 1
End If

frmMain.Picture1.PaintPicture frmPic.Picture1(2).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
lastdir = "right"
Case "left"
snakeparts(0, 1) = snakeparts(0, 1) - 1
If frmMain.gowall.Checked = True And snakeparts(0, 1) = 0 Then
snakeparts(0, 1) = 25
End If

frmMain.Picture1.PaintPicture frmPic.Picture1(3).Picture, snakeparts(0, 1), snakeparts(0, 0), 1, 1
lastdir = "left"
End Select


frmMain.Picture1.PaintPicture frmPic.Picture1(4).Picture, snakeparts(1, 1), snakeparts(1, 0), 1, 1
checkeat
crashed = False
If snakeparts(0, 0) < 0.9 Or snakeparts(0, 0) > 25.1 Or snakeparts(0, 1) < 0.9 Or snakeparts(0, 1) > 25.1 Then
crashed = True
End If
For i = 1 To snakelength - 1
If Abs(snakeparts(0, 0) - snakeparts(i, 0)) < 0.1 And Abs(snakeparts(0, 1) - snakeparts(i, 1)) < 0.1 Then
crashed = True
End If
Next
For i = 0 To blocksnum
If Abs(snakeparts(0, 0) - mazeblocks(i, 0)) < 0.1 And Abs(snakeparts(0, 1) - mazeblocks(i, 1)) < 0.1 Then
crashed = True
End If
Next

If crashed = True Then
If frmMain.crash.Checked = True Then
s = "You lost the game, your score was " + Str(score)
MsgBox (s)
gamerun = False
If highscore(score, "crash") = True Then entername
frmMain.Label1.Caption = "Restart"
Else
If Val(frmMain.Label6.Caption) = 0 Then
s = "You lost the game, your score was " + Str(score)
MsgBox (s)
gamerun = False
If highscore(score, "complete") = True Then entername
frmMain.Label1.Caption = "Restart"
Exit Sub
Else
newgame
End If
frmMain.Label6.Caption = Val(frmMain.Label6.Caption) - 1
End If
End If
End Sub
Function highscore(nscore As Integer, gametype As String) As Boolean
Select Case gametype
Case "complete"
If nscore > Val(scorescomplete(9, 2)) Then
highscore = True
Else
highscore = False
End If
Case "crash"
If nscore > Val(scorescrash(9, 2)) Then
highscore = True
Else
highscore = False
End If
End Select
End Function

Public Sub checkeat()
If Abs(snakeparts(0, 0) - foody) < 0.1 And Abs(snakeparts(0, 1) - foodx) < 0.1 Then
If foodcolor = "bomb" Then
a = snakelength
snakelength = Int(snakelength * 0.75)
b = a - snakelength
targetlength = targetlength - b
For i = snakelength To a
x = snakeparts(i - 1, 1)
y = snakeparts(i - 1, 0)
frmMain.Picture1.ForeColor = vbBlack
frmMain.Picture1.FillColor = vbBlack
frmMain.Picture1.DrawWidth = 1
frmMain.Picture1.Line (x, y)-(x + 0.92, y + 0.92), cr, BF
Next
Else
Select Case foodcolor
Case "red"
isworth = 1
Case "blue"
isworth = 2
Case "yellow"
isworth = 4
End Select
If frmMain.complete.Checked = True Then
isworth = isworth * multiplicator
Else
If frmMain.slow.Checked = True Then isworth = isworth * 2
If frmMain.normal.Checked = True Then isworth = isworth * 4
If frmMain.fast.Checked = True Then isworth = isworth * 7
If frmMain.insane.Checked = True Then isworth = isworth * 12

If mazenum = 1 Then isworth = isworth * 1.2
If mazenum = 2 Then isworth = isworth * 2
If mazenum = 3 Then isworth = isworth * 4
If mazenum = 4 Then isworth = isworth * 10

If lengthincrease = 3 Then isworth = isworth * 1.2
If lengthincrease = 5 Then isworth = isworth * 1.7
If lengthincrease = 8 Then isworth = isworth * 2.2
If lengthincrease = 12 Then isworth = isworth * 3
If lengthincrease = 17 Then isworth = isworth * 3.8
End If

isworth = isworth / 2
If frmMain.gowall.Checked = True Then isworth = isworth * 0.88
score = score + isworth
frmMain.Label3.Caption = Str(score)
targetlength = targetlength + lengthincrease
If frmMain.complete.Checked = True Then
If score > livescore Then
frmMain.Label6.Caption = Val(frmMain.Label6.Caption) + 1
livescore = livescore + 200
End If
remainingapples = remainingapples - 1
frmMain.Label8.Caption = Str(remainingapples)
If remainingapples = 0 Then
If level = 10 Then
s = "You've completed the game, your score was " + Str(score)
MsgBox (s)
gamerun = False
If highscore(score, "complete") = True Then entername
frmMain.Label1.Caption = "Restart"
Else
level = level + 1
s = "You completed level" + Str(level - 1) + ", the password for level" + Str(level) + " is  (" + password(level) + ")"
MsgBox (s)
gotolevel (level)
frmMain.Label5.Caption = Str(level)
Exit Sub
End If
End If
End If
End If
putfood
End If
End Sub


Public Sub mainloop()
    Dim t As Long 't is the time taken to do everything

Do
        t = GetTickCount
If gamerun = True Then
movesnake
frmMain.action.Caption = frmMain.Label1.Caption

End If
        t = GetTickCount - t    'This compensates for the time taken
        If t > 200 Then t = 0   '    to move and draw everything, keeps the game

Wait (s_interval - t) 's_interval) '(s_interval + 1)


DoEvents

Loop

End Sub
Sub Delay(secs)
Dim start
start = Timer
While (Timer < (start + secs))
DoEvents
Wend
End Sub

Sub main()
Load frmMain
frmMain.Show
mainloop
End Sub
Public Sub Wait(w As Integer, Optional DEvents As Boolean = True)
    Dim s As Double
    s = GetTickCount + w
    Do While GetTickCount < s
        If DEvents Then DoEvents  'Keeps every thing else running smoothy
    Loop
End Sub

