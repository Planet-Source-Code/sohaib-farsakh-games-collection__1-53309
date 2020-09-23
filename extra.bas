Attribute VB_Name = "Module1"
Private Function safelines() As Boolean
For i = 1 To boardsize - 1
For j = 1 To boardsize - 1
If grid(i, j) < 2 Then
If grid(i + 1, j) < 2 Or grid(i, j + 1) < 2 Then
safelines = True
Exit Function
End If
End If
Next
Next
safelines = False
End Function

'Dim stopwin As Boolean
'Dim takelinee As Boolean
'Dim cells(145, 1)
'Dim wins(0 To 20) As Integer
'Dim lines As Integer

If safelines = False And diff = 4 Then


For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i


ended = False
numberr = 0

Do Until ended = True
found = False
For i = 1 To boardsize Step 1
If found = True Then Exit For
For j = 1 To boardsize Step 1
If found = True Then Exit For
If grid2(i, j) = 3 Then
found = True
numberr = numberr + 1
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
grid2(i, j) = 4
If Label2(a).Enabled = True And i > 1 Then
grid2(i - 1, j) = grid2(i - 1, j) + 1
cells(numberr, 0) = i
cells(numberr, 1) = j
End If
If Label2(B).Enabled = True And i < boardsize Then
grid2(i + 1, j) = grid2(i + 1, j) + 1
cells(numberr, 0) = i
cells(numberr, 1) = j
End If
If Label3(c).Enabled = True And j > 1 Then
grid2(i, j - 1) = grid2(i, j - 1) + 1
cells(numberr, 0) = i
cells(numberr, 1) = j
End If
If Label3(d).Enabled = True And j < boardsize Then
grid2(i, j + 1) = grid2(i, j + 1) + 1
cells(numberr, 0) = i
cells(numberr, 1) = j
End If
End If
Next j
Next i
If found = False Then
ended = True
End If

Loop
'Form1.Caption = Str(numberr)


If numberr <= 5 And numberr >= 2 Then
lines = 0
For i = 1 To boardsize
For j = 1 To boardsize
If grid(i, j) = 3 Then
a = (i - 1) * boardsize + j - 1
B = a + boardsize
c = (i - 1) * (boardsize + 1) + j - 1
d = c + 1
If Label2(a).Enabled = True Then
lines = lines + 1
wins(lines) = a
End If
If Label2(B).Enabled = True Then
lines = lines + 1
wins(lines) = B
End If
If Label3(c).Enabled = True Then
lines = lines + 1
wins(lines) = c + (horlinenum + 1)
End If
If Label3(d).Enabled = True Then
lines = lines + 1
wins(lines) = d + (horlinenum + 1)
End If
End If
Next j
Next i


stopwin = True
For i = 1 To lines

If stopwin = False Then Exit For

For y = 1 To boardsize
For l = 1 To boardsize
grid2(y, l) = grid(y, l)
Next l
Next y

If Val(wins(i)) < horlinenum Then
u = Int(wins(i) / boardsize) + 1
v = (wins(i) Mod boardsize) + 1
If u > 1 Then
grid2(u - 1, v) = grid2(u - 1, v) + 1
End If
If u < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
Else
u = Int((wins(i) - (horlinenum + 1)) / (boardsize + 1)) + 1
v = (wins(i) - (horlinenum + 1) Mod (boardsize + 1)) + 1
If v > 1 Then
grid2(u, v - 1) = grid2(u, v - 1) + 1
End If
If v < (boardsize + 1) Then
grid2(u, v) = grid2(u, v) + 1
End If
End If
takelinee = False

For j = 1 To numberr

a = cells(j, 0)
B = cells(j, 1)
If grid2(a, B) < 3 Then

If a > 1 Then
If Label2((a - 1) * boardsize + B - 1).Enabled = True And grid2(a - 1, B) < 3 Then
takelinee = True
Exit For
End If
End If

If a < (boardsize) Then
If Label2((a) * boardsize + B - 1).Enabled = True And grid2(a + 1, B) < 3 Then
takelinee = True
Exit For
End If
End If

If B > 1 Then
If Label3(a * (boardsize + 1) + B - 1).Enabled = True And grid2(a, B - 1) < 3 Then
takelinee = True
Exit For
End If
End If

If B < (boardsize) Then
If Label3(a * (boardsize + 1) + B).Enabled = True And grid2(a, B + 1) < 3 Then
takelinee = True
Exit For
End If
End If

End If

Next j

If takelinee = True Then
element = 1
plays(1) = wins(i)
element = 1
stopwin = False
End If
Next i


For i = 1 To boardsize
For j = 1 To boardsize
grid2(i, j) = grid(i, j)
Next j
Next i

If stopwin = True Then

For j = 1 To numberr
a = cells(j, 0)
B = cells(j, 1)
If grid2(a, B) < 3 Then

If a > 1 Then
If Label2((a - 1) * boardsize + B - 1).Enabled = True And grid2(a - 1, B) < 3 Then
plays(1) = a * B - 1
element = 1
Exit For
End If
End If

If a < (boardsize + 1) Then
If Label2((a) * boardsize + B - 1).Enabled = True And grid2(a - 1, B) < 3 Then
plays(1) = (a + 1) * B - 1
element = 1
Exit For
End If
End If

If B > 1 Then
If Label3(a * (boardsize + 1) + B - 1).Enabled = True And grid2(a, B - 1) < 3 Then
plays(1) = (a * B - 1) + (horlinenum + 1)
element = 1
Exit For
End If
End If

If B < (boardsize + 1) Then
If Label3(a * (boardsize + 1) + B).Enabled = True And grid2(a, B + 1) < 3 Then
plays(1) = (a * (B + 1) - 1) + (horlinenum + 1)
element = 1
Exit For
End If
End If

End If
Next j


End If



End If
End If

