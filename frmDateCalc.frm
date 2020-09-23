VERSION 5.00
Begin VB.Form frmDateCalc 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txted 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtetm 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txteth 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtsd 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtstm 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtsth 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdCalcDate 
      Caption         =   "Calulate Date"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "(mm/dd/yyyy)"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Minute"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Hour (24)"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmDateCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

End Sub
'BTCC(rlt) Calculate Elapsed Time
Private Sub cmdCalcDate_Click()

Dim sd, ed As Date
Dim sth, stm, eth, etm, smin, emin, tdiff, ww, hh, dd, mm, tday

sd = DateValue(txtsd.Text)   ' Starting Date
ed = DateValue(txted.Text)   ' Ending Date

sth = Int(txtsth.Text)  ' Starting Hour in 24hr
stm = Int(txtstm.Text)  ' Starting Minutes
eth = Int(txteth.Text)  ' Ending Hour in 24 hr
etm = Int(txtetm.Text)  ' Ending Minutes


' COMPUTE ACTUAL TIME SPENT ON STARTING DAY
smin = ((23 - sth) * 60) + (60 - stm)  ' TIME BETWEEN START TIME & MIDNIGHT

' COMPUTE ACTUAL TIME SPENT ON ENDING DAY
emin = (eth * 60) + etm                ' TIME SINCE MIDNIGHT

tdiff = smin + emin                    ' TOTAL TIME SPENT ON START & END DAY

If tdiff <= 1440 Then                  ' ADJUST IF START / END DATE WITHIN 24 HR PRD
    tday = ed - sd - 1
ElseIf tdiff > 1440 Then             ' ADJUST FOR NUMBER OF DAYS INBETWEEN
    tdiff = tdiff - 1440
    tday = ed - sd
End If
      
tdiff = tdiff + (tday * 1440) ' ADD 24 HRS FOR EVERYDAY IN BETWEEN

' CALCULATE DAYS
If tdiff >= 1440 Then
  dd = 1
  tdiff = tdiff - 1440
  Do While tdiff > 1440
     tdiff = tdiff - 1440
     dd = dd + 1
  Loop
End If

' CALCULATE HOURS
If tdiff >= 60 Then
  hh = 1
  tdiff = tdiff - 60
  Do While tdiff >= 60
      tdiff = tdiff - 60
    hh = hh + 1
  Loop
End If

' IF EXACTLY 24 HOURS BUMP UP THE DAY COUNTER
If hh = 24 Then
   hh = 0
   dd = dd + 1
End If

' MINUTES IS WHAT IS LEFT
mm = tdiff

' CONVERT DAYS TO WEEKS IF REQUIRED
If dd >= 7 Then
   ww = 1
   dd = dd - 7
   Do While dd >= 7
      dd = dd - 7
      ww = ww + 1
   Loop
End If

Text1.Text = "Weeks: " & ww & vbNewLine & _
             "Days: " & dd & vbNewLine & _
             "Hours: " & hh & vbNewLine & _
             "Min: " & mm & vbNewLine
 
Text1.FontBold = True

End Sub

Private Sub Form_Load()

' SAMPLE STARTING DATE AND TIMES FOR TESTING

txtsd.Text = "08/1/2001"
txted.Text = "08/3/2001"
txtsth.Text = 22
txtstm.Text = 50
txteth.Text = 10
txtetm.Text = 45

End Sub
