VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Equations"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu zoom 
         Caption         =   "Zoom..."
      End
      Begin VB.Menu clrall 
         Caption         =   "Clear All"
      End
      Begin VB.Menu cls 
         Caption         =   "Clear Drawing"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu sd 
         Caption         =   "Show Gridlines"
      End
      Begin VB.Menu spl 
         Caption         =   "Show Pointer Location"
         Checked         =   -1  'True
      End
      Begin VB.Menu draw 
         Caption         =   "Drawing color"
      End
      Begin VB.Menu clrchg 
         Caption         =   "Auto Color Change"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
      Begin VB.Menu sci 
         Caption         =   "Scientific Calculator"
      End
      Begin VB.Menu equ 
         Caption         =   "Equation Solver"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub about_Click()
    ShellAbout Me.hWnd, "Equations Program", "Author: Sohaib Abu Farsakh" & vbCrLf, Me.Icon
End Sub

Private Sub clrall_Click()
Unload Form11
Load Form11
Form11.Show
End Sub

Private Sub clrchg_Click()
clrchg.Checked = Not (clrchg.Checked)
End Sub

Private Sub cls_Click()
Form11.cls
Form11.Combo1.Tag = 0
Form11.Text11.Text = Form11.Text11.Text + "1"
End Sub

Private Sub deg_Click()
Form11.Combo1.Text = "Degrees"
deg.Checked = True
rad.Checked = False
grad.Checked = False

End Sub

Private Sub draw_Click()
Form11.CommonDialog1.ShowColor
Form11.ForeColor = Form11.CommonDialog1.Color
End Sub

Private Sub equ_Click()
If Form11.Label1.Caption = "5" Or Form11.Label1.Caption = "7" Or Form11.Label1.Caption = "8" Or Form11.Label1.Caption = "11" Then
d = MsgBox("Can't solve fifth degree,sin,cos or complex equations", vbOKOnly + vbExclamation, "Equations")
Else
Form11.Frame1.Visible = True
Form11.Option1.Visible = True
Form11.Option2.Visible = True
Form11.Option3.Visible = True
Form11.Option4.Visible = True
Form11.Option5.Visible = True
Form11.Option6.Visible = True
Form11.Option7.Visible = True
Form11.Combo1.Visible = True
Form11.Text3.Text = "0"
Form11.Label12.Caption = ""
End If
Form11.Frame1.Top = 720
Form11.Frame1.Left = 4560
End Sub


Private Sub grad_Click()
Form11.Combo1.Text = "Gradians"
grad.Checked = True
rad.Checked = False
deg.Checked = False

End Sub


Private Sub MDIForm_Load()
Load Form11
Form11.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Form1
End Sub

Private Sub rad_Click()
Form11.Combo1.Text = "Radians"
rad.Checked = True
deg.Checked = False
grad.Checked = False

End Sub

Private Sub right_Click()

End Sub

Private Sub sci_Click()
Load Form1
Form1.Show
End Sub

Private Sub sd_Click()
Select Case sd.Checked
Case True
sd.Checked = False
Form11.Check2.Value = 0
Case False
sd.Checked = True
Form11.Check2.Value = 1
End Select
End Sub

Private Sub spl_Click()
Select Case spl.Checked
Case True
spl.Checked = False
Form11.Label6.Visible = False
Form11.Label7.Visible = False
Case False
spl.Checked = True
Form11.Label6.Visible = True
Form11.Label7.Visible = True
End Select
End Sub

Private Sub Timer1_Timer()
Load Form11
Form11.Show
Timer1.Enabled = False
End Sub

Private Sub zoom_Click()
Form11.Frame3.Visible = True
Form11.Frame3.Top = 720
Form11.Frame3.Left = 4560
End Sub
