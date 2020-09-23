VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Window"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Text1_Change()
Form1.Winsock1.SendData "C" + Text1.Text
End Sub
