VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ÞÇãæÓ"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ÊÑÌã"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Text1.Text
Case "one"
Text2.Text = "æÇÍÏ"
Case "two"
Text2.Text = "ÇËäÇä"
Case "three"
Text2.Text = "ËáÇËÉ"
Case "four"
Text2.Text = "ÃÑÈÚÉ"
Case "five"
Text2.Text = "ÎãÓÉ"
Case "six"
Text2.Text = "ÓÊÉ"
Case "seven"
Text2.Text = "ÓÈÚÉ"
Case "eight"
Text2.Text = "ËãÇäíÉ"
Case "nine"
Text2.Text = "ÊÓÚÉ"
Case "ten"
Text2.Text = "ÚÔÑÉ"
Case "name"
Text2.Text = "ÇÓã"
Case "hello"
Text2.Text = "ÃåáÇð"
Case "plane"
Text2.Text = "ØÇÆÑÉ"
Case "kite"
Text2.Text = "ØÇÆÑÉ ÒÑÞíÉ"
Case "car"
Text2.Text = "ÓíÇÑÉ"
Case "bus"
Text2.Text = "ÍÇÝáÉ"
Case "doll"
Text2.Text = "ÏãíÉ"
Case "house"
Text2.Text = "ÈíÊ"
Case "boy"
Text2.Text = "æáÏ"
Case "girl"
Text2.Text = "ÈäÊ"
Case "robot"
Text2.Text = "ÑÌá Âáí"
Case "car"
Text2.Text = "ÓíÇÑÉ"
Case "bag"
Text2.Text = "ÍÞíÈÉ"
Case "shoe"
Text2.Text = "ÍÐÇÁ"
Case "box"
Text2.Text = "ÕäÏæÞ"
Case "pencil"
Text2.Text = "Þáã ÑÕÇÕ"
Case "rubber"
Text2.Text = "ããÍÇÉ"
Case "cat"
Text2.Text = "ÞØÉ"
Case "look"
Text2.Text = "ÇäÙÑ"
Case "apple"
Text2.Text = "ÊÝÇÍÉ"
Case "owl"
Text2.Text = "ÈæãÉ"
Case "insect"
Text2.Text = "ÍÔÑÉ"
Case "ball"
Text2.Text = "ßÑÉ"
Case "dog"
Text2.Text = "ßáÈ"
Case "umbrella"
Text2.Text = "ãÙáÉ"
Case "parrot"
Text2.Text = "ÈÈÛÇÁ"
Case "elephant"
Text2.Text = "Ýíá"
Case "purple"
Text2.Text = "ÈäÝÓÌí"
Case "blue"
Text2.Text = "ÃÒÑÞ"
Case "green"
Text2.Text = "ÃÎÖÑ"
Case "yellow"
Text2.Text = "ÃÕÝÑ"
Case "red"
Text2.Text = "ÃÍãÑ"
Case "orange"
Text2.Text = "ÈÑÊÞÇáí"
Case "color"
Text2.Text = "áæä"
Case "balloon"
Text2.Text = "ÈÇáæä"
Case "bear"
Text2.Text = "ÏÈ"
Case "lion"
Text2.Text = "ÃÓÏ"
Case "monkey"
Text2.Text = "ÞÑÏ"
Case "zebra"
Text2.Text = "ÍãÇÑ æÍÔí"
Case "sunny"
Text2.Text = "ãÔãÓ"
Case "dinosaur"
Text2.Text = "ÏíäÇÕæÑ"
Case "little"
Text2.Text = "ÕÛíÑ"
Case "big"
Text2.Text = "ßÈíÑ"
Case "candle"
Text2.Text = "ÔãÚÉ"
Case "cake"
Text2.Text = "ßÚßÉ"
Case "hat"
Text2.Text = "ØÇÞíÉ"
Case "present"
Text2.Text = "åÏíÉ"
Case "party"
Text2.Text = "ÍÝáÉ"
Case "birthday"
Text2.Text = "ÚíÏ ãíáÇÏ"
Case "monkey"
Text2.Text = "ÞÑÏ"
Case "clown"
Text2.Text = "ãåÑÌ"
Case "happy"
Text2.Text = "ÓÚíÏ"
Case "sad"
Text2.Text = "ÍÒíä"
Case "friend"
Text2.Text = "ÕÏíÞ"
Case "hand"
Text2.Text = "íÏ"
Case "arm"
Text2.Text = "ÐÑÇÚ"
Case "leg"
Text2.Text = "ÑÌá"
Case "foot"
Text2.Text = "ÞÏã"
Case "frog"
Text2.Text = "ÖÝÏÚ"
Case "pen"
Text2.Text = "Þáã"
Case "hen"
Text2.Text = "ÏÌÇÌÉ"
Case "fish"
Text2.Text = "ÓãßÉ"
Case "bird"
Text2.Text = "ØÇÆÑ"
Case "rabbit"
Text2.Text = "ÃÑäÈ"
Case "boat"
Text2.Text = "ÞÇÑÈ"
Case "train"
Text2.Text = "ÞØÇÑ"
Case "hair"
Text2.Text = "ÔÚÑ"
Case "eye"
Text2.Text = "Úíä"
Case "kitchen"
Text2.Text = "ãØÈÎ"
Case "table"
Text2.Text = "ØÇæáÉ"
Case "bed"
Text2.Text = "ÓÑíÑ"
Case "sofa"
Text2.Text = "ßäÈÉ"
Case "head"
Text2.Text = "ÑÃÓ"
Case "phone"
Text2.Text = "åÇÊÝ"
Case "fridge"
Text2.Text = "ËáÇÌÉ"
Case "chair"
Text2.Text = "ßÑÓí"
Case "duck"
Text2.Text = "ÈØÉ"
Case "desk"
Text2.Text = "ãßÊÈ"
Case "over"
Text2.Text = "ÝæÞ"
Case "under"
Text2.Text = "ÊÍÊ"
Case "between"
Text2.Text = "Èíä"
Case "book"
Text2.Text = "ßÊÇÈ"
Case "mouse"
Text2.Text = "ÝÃÑ"
Case "circle"
Text2.Text = "ÏÇÆÑÉ"
Case "triangle"
Text2.Text = "ãËáË"
Case "square"
Text2.Text = "ãÑÈÚ"
Case "yes"
Text2.Text = "äÚã"
Case "no"
Text2.Text = "áÇ"
Case "grandfather"
Text2.Text = "ÌÏ"
Case "grandmother"
Text2.Text = "ÌÏÉ"
Case "brother"
Text2.Text = "ÃÎæ"
Case "sister"
Text2.Text = "ÃÎÊ"
Case "pet"
Text2.Text = "ÍíæÇä ÃáíÝ"
Case "zoo"
Text2.Text = "ÍÏíÞÉ ÍíæÇäÇÊ"
Case "kangaroo"
Text2.Text = "ßäÛÑ"
Case "cow"
Text2.Text = "ÈÞÑÉ"
Case "horse"
Text2.Text = "ÍÕÇä"
Case "hill"
Text2.Text = "ÊáÉ"
Case "tree"
Text2.Text = "ÔÌÑÉ"
Case "man"
Text2.Text = "ÑÌá"
Case "picture"
Text2.Text = "ÕæÑÉ"
Case "teddy"
Text2.Text = "ÏÈÏæÈ"
Case "train"
Text2.Text = "ÞØÇÑ"
Case "bike"
Text2.Text = "ÏÑÇÌÉ"
Case "say"
Text2.Text = "íÞæá"
Case "fall"
Text2.Text = "íÓÞØ"
Case "count"
Text2.Text = "íÚÏ"
Case "goat"
Text2.Text = "ÚäÒÉ"
Case "prince"
Text2.Text = "ÃãíÑ"
Case "princess"
Text2.Text = "ÃãíÑÉ"
Case "castle"
Text2.Text = "ÞáÚÉ"
Case "clock"
Text2.Text = "ÓÇÚÉ"
Case "room"
Text2.Text = "ÛÑÝÉ"
Case "kitchen"
Text2.Text = "ãØÈÎ"
Case "watch"
Text2.Text = "ÓÇÚÉ íÏ"
Case "bedroom"
Text2.Text = "ÛÑÝÉ äæã"
Case "blonde"
Text2.Text = "ÃÔÞÑ"
Case "mouth"
Text2.Text = "Ýã"
Case "brown"
Text2.Text = "Èäí"
Case "pink"
Text2.Text = "æÑÏí"
Case "long"
Text2.Text = "Øæíá"
Case "tall"
Text2.Text = "Øæíá"
Case "short"
Text2.Text = "ÞÕíÑ"
Case "song"
Text2.Text = "ÃÛäíÉ"
Case "sing"
Text2.Text = "íÛäí"
Case "egg"
Text2.Text = "ÈíÖÉ"
Case "tea"
Text2.Text = "ÔÇí"
Case "grey"
Text2.Text = "ÑãÇÏí"
Case "coffee"
Text2.Text = "ÞåæÉ"
Case "night"
Text2.Text = "áíá"
Case "ear"
Text2.Text = "ÃÐä"
Case "crocodile"
Text2.Text = "ÊãÓÇÍ"
Case "nose"
Text2.Text = "ÃäÝ"
Case "neck"
Text2.Text = "ÑÞÈÉ"
Case "tail"
Text2.Text = "Ðíá"
Case "giraff"
Text2.Text = "ÒÑÇÝÉ"
Case "animal"
Text2.Text = "ÍíæÇä"
Case "inside"
Text2.Text = "ÏÇÎá"
Case "trunk"
Text2.Text = "ÎÑØæã"
Case "morning"
Text2.Text = "ÕÈÇÍ"
Case "sit"
Text2.Text = "íÌáÓ"
Case "write"
Text2.Text = "íßÊÈ"
Case "read"
Text2.Text = "íÞÑÃ"
Case "draw"
Text2.Text = "íÑÓã"
Case "stand"
Text2.Text = "íÞÝ"
Case "open"
Text2.Text = "íÝÊÍ"
Case "close"
Text2.Text = "íÛáÞ"
Case "copy"
Text2.Text = "íäÓÎ"
Case "banana"
Text2.Text = "ãæÒÉ"
Case "wood"
Text2.Text = "ÎÔÈ"
Case "slow"
Text2.Text = "ÈØíÁ"
Case "fast"
Text2.Text = "ÓÑíÚ"
Case "many"
Text2.Text = "ßËíÑ"
Case "much"
Text2.Text = "ßËíÑ"
Case "river"
Text2.Text = "äåÑ"
Case "street"
Text2.Text = "ÔÇÑÚ"
Case "swim"
Text2.Text = "íÓÈÍ"
Case "run"
Text2.Text = "íÌÑí"
Case "walk"
Text2.Text = "íãÔí"
Case "jump"
Text2.Text = "íÞÝÒ"
Case "can"
Text2.Text = "íÓÊØíÚ"
Case "ride"
Text2.Text = "íÑßÈ"
Case "drive"
Text2.Text = "íÞæÏ"
Case "hop"
Text2.Text = "íÞÝÒ"
Case "fly"
Text2.Text = "íØíÑ"
Case "good"
Text2.Text = "ÌíÏ"
Case "bad"
Text2.Text = "ÓíÁ"
Case "sleep"
Text2.Text = "íäÇã"
Case "eat"
Text2.Text = "íÃßá"
Case "drink"
Text2.Text = "íÔÑÈ"
Case "see"
Text2.Text = "íÑì"
Case "leaf"
Text2.Text = "æÑÞÉ ÔÌÑ"
Case "paper"
Text2.Text = "æÑÞÉ"
Case "meat"
Text2.Text = "áÍãÉ"
Case "desert"
Text2.Text = "ÕÍÑÇÁ"
Case "farm"
Text2.Text = "ãÒÑÚÉ"
Case "forest"
Text2.Text = "ÛÇÈÉ"
Case "snake"
Text2.Text = "ÃÝÚì"
Case "camel"
Text2.Text = "Ìãá"
Case "flower"
Text2.Text = "ÒåÑÉ"
Case "zebra"
Text2.Text = "ÍãÇÑ æÍÔí"
Case "sandwich"
Text2.Text = "ÓäÏæíÔÉ"
Case "look"
Text2.Text = "íäÙÑ"
Case "hear"
Text2.Text = "íÓãÚ"
Case "cry"
Text2.Text = "íÈßí"
Case "king"
Text2.Text = "ãáß"
Case "kitten"
Text2.Text = "ÞØÉ ÕÛíÑÉ"
Case "time"
Text2.Text = "æÞÊ"
Case "basket"
Text2.Text = "ÓáÉ"
Case "garden"
Text2.Text = "ÍÏíÞÉ ÇáãäÒá"
Case "home"
Text2.Text = "ãäÒá"
Case "play"
Text2.Text = "íáÚÈ"
Case "woman"
Text2.Text = "ÇãÑÃÉ"
Case "sun"
Text2.Text = "ÔãÓ"
Case "lemon"
Text2.Text = "áíãæä"
Case "fruit"
Text2.Text = "ÝÇßåÉ"
Case "bee"
Text2.Text = "äÍáÉ"
Case "black"
Text2.Text = "ÃÓæÏ"
Case "pink"
Text2.Text = "æÑÏí"
Case "stand"
Text2.Text = "íÞÝ"
Case "eleven"
Text2.Text = "ÃÍÏ ÚÔÑ"
Case "twelve"
Text2.Text = "ÇËäÇ ÚÔÑ"
Case "thirteen"
Text2.Text = "ËáÇË ÚÔÑ"
Case "fourteen"
Text2.Text = "ÃÑÈÚ ÚÔÑ"
Case "fifteen"
Text2.Text = "ÎãÓ ÚÔÑ"
Case "sixteen"
Text2.Text = "ÓÊ ÚÔÑ"
Case "seventeen"
Text2.Text = "ÓÈÚ ÚÔÑ"
Case "eighteen"
Text2.Text = "ËãÇäí ÚÔÑ"
Case "nineteen"
Text2.Text = "ÊÓÚ ÚÔÑ"
Case "twenty"
Text2.Text = "ÚÔÑæä"
Case "door"
Text2.Text = "ÈÇÈ"
Case "window"
Text2.Text = "äÇÝÐÉ"

End Select
End Sub
