VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "áÚÈÉ ÇáËÚÈÇä"
   ClientHeight    =   6420
   ClientLeft      =   2310
   ClientTop       =   555
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   7110
   Begin VB.CommandButton Command6 
      Caption         =   "ÅíÞÇÝ"
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ÅÚÇÏÉ"
      Height          =   495
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   5160
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   3000
      Max             =   10
      Min             =   1
      TabIndex        =   7
      Top             =   6120
      Value           =   1
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "íÓÇÑ"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÊÍÊ"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "íãíä"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÝæÞ"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1200
      Top             =   4680
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   416
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   415
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   414
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   413
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   412
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   411
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   410
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   409
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   408
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   407
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   406
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   405
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   404
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   403
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   402
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   401
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   400
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   399
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   398
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   397
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   396
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   395
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   394
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   393
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   392
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   391
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   390
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   389
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   388
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   387
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   386
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   385
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   384
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   383
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   382
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   381
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   380
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   379
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   378
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   377
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   376
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   374
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   373
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   372
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   371
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   370
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   369
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   368
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   367
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   366
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   365
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   364
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   363
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   362
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   361
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   360
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   359
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   358
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   357
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   356
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   355
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   354
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   353
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   352
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   351
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   350
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   349
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   348
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   347
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   346
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   345
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   344
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   343
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   342
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   341
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   340
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   339
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   338
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   337
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   336
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   335
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   334
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   333
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   332
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   331
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   329
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   328
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   327
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   326
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   325
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   324
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   323
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   322
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   321
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   320
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   319
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   318
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   317
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   316
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   315
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   314
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   313
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   312
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   311
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   310
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   309
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   308
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   307
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   306
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   305
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   304
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   303
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   302
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   301
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   300
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   299
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   298
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   297
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   296
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   295
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   294
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   293
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   292
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   291
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   290
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   289
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   288
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   287
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   286
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   285
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   284
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   283
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   282
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   281
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   280
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   279
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   278
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   277
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   276
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   275
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   274
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   273
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   272
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   271
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   270
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   269
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   268
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   267
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   266
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   265
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   264
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   263
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   262
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   261
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   260
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   259
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   258
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   257
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   256
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   255
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   254
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   253
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   252
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   251
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   250
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   249
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   248
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   247
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   246
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   245
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   244
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   243
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   242
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   241
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   239
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   238
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   237
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   236
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   235
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   234
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   233
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   232
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   231
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   230
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   229
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   228
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   227
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   226
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   225
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   224
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   223
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   222
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   221
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   220
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   219
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   218
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   217
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   216
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   215
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   214
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   213
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   212
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   211
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   210
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   209
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   208
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   207
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   206
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   205
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   204
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   203
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   202
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   201
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   200
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   199
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   198
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   197
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   196
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   195
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   194
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   193
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   192
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   191
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   190
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   189
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   188
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   187
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   186
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   185
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   184
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   183
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   182
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   181
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   180
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   179
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   178
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   177
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   176
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   175
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   174
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   173
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   172
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   171
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   170
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   169
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   168
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   167
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   166
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   165
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   164
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   163
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   162
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   161
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   160
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   159
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   158
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   157
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   156
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   155
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   154
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   153
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   152
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   151
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   150
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   149
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   148
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   147
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   146
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   145
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   144
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   143
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   142
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   141
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   140
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   139
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   138
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   137
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   136
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   134
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   133
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   132
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   131
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   130
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   129
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   128
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   127
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   126
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   125
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   124
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   123
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   122
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   121
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   120
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   119
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   118
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   117
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   116
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   115
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   114
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   113
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   112
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   111
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   110
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   109
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   108
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   107
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   106
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   105
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   104
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   103
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   102
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   101
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   100
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   99
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   98
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   97
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   96
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   95
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   94
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   93
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   92
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   91
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   90
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   89
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   88
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   87
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   86
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   85
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   84
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   83
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   82
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   81
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   80
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   79
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   78
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   77
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   76
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   75
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   74
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   73
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   72
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   71
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   70
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   69
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   68
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   67
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   66
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   65
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   64
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   63
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   62
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   61
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   60
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   59
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   58
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   57
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   56
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   55
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   54
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   53
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   52
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   51
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   50
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   49
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   48
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   47
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   46
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   45
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   44
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   43
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   42
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   41
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   40
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   39
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   38
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   37
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   36
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   35
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   34
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   33
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   32
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   31
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   30
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   29
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   28
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   27
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   26
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   25
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   24
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   23
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   22
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   21
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   20
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   19
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   18
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   17
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   16
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   15
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   14
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   13
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   12
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   11
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   10
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   9
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   8
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   7
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "ÃÚáì ÏÑÌÉ"
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "ÇáÓÑÚÉ"
      Height          =   255
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÏÑÌÉ"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape25 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   200
      Left            =   4000
      Top             =   3400
      Width           =   200
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6840
      Y1              =   4400
      Y2              =   4400
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   2160
      Top             =   800
      Width           =   195
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   200
      Left            =   2000
      Top             =   800
      Width           =   200
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   200
      Left            =   1800
      Shape           =   1  'Square
      Top             =   800
      Width           =   200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Label1.Caption = "8"
End Sub

Private Sub Command2_Click()
Label1.Caption = "6"
End Sub

Private Sub Command3_Click()
Label1.Caption = "2"
End Sub

Private Sub Command4_Click()
Label1.Caption = "4"
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
Command5.Visible = False
Shape1.Top = 800
Shape1.Left = 1800
Shape2.Top = 800
Shape2.Left = 2000
Shape3(0).Top = 800
Shape3(0).Left = 2160
Command6.Caption = "ÇÈÏÃ"
End Sub

Private Sub Command6_Click()
Select Case Command6.Caption
Case "ÅíÞÇÝ"
Command6.Caption = "ÅßãÇá"
Case "ÅßãÇá"
Command6.Caption = "ÅíÞÇÝ"
Case "ÇÈÏÃ"
For i = 1 To 416 Step 1
Shape3(i).Visible = False
Next i
Label5.Caption = "0"
Label2.Caption = "0"


Command6.Caption = "ÅíÞÇÝ"
End Select
End Sub

Private Sub HScroll1_Change()
Label3.Caption = Str$(HScroll1.Value)
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Label1.Caption = Text1.Text
End If
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
If Command6.Caption = "ÅíÞÇÝ" Then
If Label1.Caption = "8" Or Label1.Caption = "4" Or Label1.Caption = "6" Or Label1.Caption = "2" Then
Shape3(Label2.Caption).Visible = True
If Label1.Caption <> "" Then
For i = 416 To 1 Step -1
Shape3(i).Top = Shape3(i - 1).Top
Shape3(i).Left = Shape3(i - 1).Left
Next i
Shape3(0).Top = Shape2.Top
Shape3(0).Left = Shape2.Left
Shape2.Top = Shape1.Top
Shape2.Left = Shape1.Left

Select Case Label1.Caption
Case "8"
Shape1.Top = Shape1.Top - 200
Case "2"
Shape1.Top = Shape1.Top + 200
Case "6"
Shape1.Left = Shape1.Left + 200
Case "4"
Shape1.Left = Shape1.Left - 200
End Select
If Shape1.Top = Shape25.Top And Shape1.Left = Shape25.Left Then
Label2.Caption = Val(Label2.Caption) + 1
Select Case Shape25.FillColor
Case &HFF&
Label5.Caption = Val(Label5.Caption) + (1 * Val(Label3.Caption))
Case &HC00000
Label5.Caption = Val(Label5.Caption) + (2 * Val(Label3.Caption))
Case &HFFFF&
Label5.Caption = Val(Label5.Caption) + (4 * Val(Label3.Caption))
End Select
Randomize
a = Int(Rnd * 22)
Shape25.Top = a * 200
b = Int(Rnd * 35)
Shape25.Left = b * 200
c = Int(Rnd * 6) + 1
If c = 6 Then
Shape25.FillColor = &HFFFF&
Else
If c = 4 Or c = 5 Then
Shape25.FillColor = &HC00000
Else
Shape25.FillColor = &HFF&
End If
End If
End If
End If
End If
If Val(Label5.Caption) > Val(Label8.Caption) Then
Label8.Caption = Label5.Caption
End If
End If
End Sub

Private Sub Timer2_Timer()
Select Case HScroll1.Value
Case 1
Timer1.Interval = 500
Case 2
Timer1.Interval = 350
Case 3
Timer1.Interval = 250
Case 4
Timer1.Interval = 200
Case 5
Timer1.Interval = 160
Case 6
Timer1.Interval = 120
Case 7
Timer1.Interval = 80
Case 8
Timer1.Interval = 50
Case 9
Timer1.Interval = 30
Case 10
Timer1.Interval = 18
End Select
If Shape1.Left > 7000 Or Shape1.Left < 0 Or Shape1.Top < 0 Or Shape1.Top > 4200 Then
Command5.Visible = True
Timer1.Enabled = False
End If
For i = 0 To 416
If Shape1.Top = Shape3(i).Top And Shape1.Left = Shape3(i).Left And Shape3(i).Visible = True Then
Command5.Visible = True
Timer1.Enabled = False
End If
Next i
End Sub
