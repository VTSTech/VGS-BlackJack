VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "VGS-BlackJack v0.01"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "New Game"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   20760
      Top             =   6360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stand"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hit"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   6240
      Width           =   5895
   End
   Begin VB.Image Image65 
      Height          =   1545
      Left            =   6480
      Picture         =   "Form1.frx":0016
      Top             =   6240
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image64 
      Height          =   1485
      Left            =   4800
      Picture         =   "Form1.frx":0568
      Top             =   240
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image63 
      Height          =   1485
      Left            =   3720
      Picture         =   "Form1.frx":0BB0
      Top             =   240
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image62 
      Height          =   1485
      Left            =   2640
      Picture         =   "Form1.frx":11F8
      Top             =   240
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image61 
      Height          =   1485
      Left            =   4800
      Picture         =   "Form1.frx":1840
      Top             =   3720
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image60 
      Height          =   1485
      Left            =   3720
      Picture         =   "Form1.frx":1E88
      Top             =   3720
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image59 
      Height          =   1485
      Left            =   2640
      Picture         =   "Form1.frx":24D0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "VGS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2767
      TabIndex        =   3
      Top             =   8160
      Width           =   390
   End
   Begin VB.Image Image58 
      Height          =   1485
      Left            =   1320
      Picture         =   "Form1.frx":2B18
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Image Image57 
      Height          =   1485
      Left            =   1320
      Picture         =   "Form1.frx":3160
      Top             =   240
      Width           =   1080
   End
   Begin VB.Image Image56 
      Height          =   1485
      Left            =   240
      Picture         =   "Form1.frx":37A8
      Top             =   240
      Width           =   1080
   End
   Begin VB.Image Image55 
      Height          =   1485
      Left            =   240
      Picture         =   "Form1.frx":3DF0
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Image Image54 
      Height          =   1485
      Left            =   480
      Picture         =   "Form1.frx":4438
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Image Image53 
      Height          =   1485
      Left            =   6480
      Picture         =   "Form1.frx":4A80
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image52 
      Height          =   1545
      Left            =   20520
      Picture         =   "Form1.frx":50C8
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image51 
      Height          =   1560
      Left            =   19440
      Picture         =   "Form1.frx":5606
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image50 
      Height          =   1545
      Left            =   18360
      Picture         =   "Form1.frx":5EE5
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image49 
      Height          =   1545
      Left            =   17280
      Picture         =   "Form1.frx":6805
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image48 
      Height          =   1545
      Left            =   16200
      Picture         =   "Form1.frx":7130
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image47 
      Height          =   1545
      Left            =   15120
      Picture         =   "Form1.frx":7788
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image46 
      Height          =   1545
      Left            =   14040
      Picture         =   "Form1.frx":7DC6
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image45 
      Height          =   1545
      Left            =   12960
      Picture         =   "Form1.frx":83BF
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image44 
      Height          =   1545
      Left            =   11880
      Picture         =   "Form1.frx":8982
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image43 
      Height          =   1560
      Left            =   10680
      Picture         =   "Form1.frx":8F1B
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image42 
      Height          =   1545
      Left            =   9600
      Picture         =   "Form1.frx":94D6
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image41 
      Height          =   1545
      Left            =   8520
      Picture         =   "Form1.frx":9A6B
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image40 
      Height          =   1545
      Left            =   7440
      Picture         =   "Form1.frx":9FB5
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image39 
      Height          =   1545
      Left            =   20520
      Picture         =   "Form1.frx":A4EB
      Top             =   1560
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image38 
      Height          =   1560
      Left            =   19440
      Picture         =   "Form1.frx":AA3D
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image37 
      Height          =   1560
      Left            =   18360
      Picture         =   "Form1.frx":B34F
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image36 
      Height          =   1545
      Left            =   17280
      Picture         =   "Form1.frx":BCC5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image35 
      Height          =   1545
      Left            =   16200
      Picture         =   "Form1.frx":C60C
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image34 
      Height          =   1545
      Left            =   15120
      Picture         =   "Form1.frx":CC37
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image33 
      Height          =   1545
      Left            =   14040
      Picture         =   "Form1.frx":D227
      Top             =   1560
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image32 
      Height          =   1545
      Left            =   12960
      Picture         =   "Form1.frx":D7FD
      Top             =   1560
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image31 
      Height          =   1545
      Left            =   11880
      Picture         =   "Form1.frx":DDAB
      Top             =   1560
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image30 
      Height          =   1560
      Left            =   10800
      Picture         =   "Form1.frx":E32A
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image29 
      Height          =   1545
      Left            =   9720
      Picture         =   "Form1.frx":E8C8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image28 
      Height          =   1545
      Left            =   8640
      Picture         =   "Form1.frx":EE42
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image27 
      Height          =   1545
      Left            =   7560
      Picture         =   "Form1.frx":F380
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image26 
      Height          =   1515
      Left            =   20520
      Picture         =   "Form1.frx":F8D0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image25 
      Height          =   1530
      Left            =   19440
      Picture         =   "Form1.frx":FE0B
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image24 
      Height          =   1515
      Left            =   18360
      Picture         =   "Form1.frx":106F2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image23 
      Height          =   1545
      Left            =   17280
      Picture         =   "Form1.frx":1101D
      Top             =   3120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image22 
      Height          =   1545
      Left            =   16200
      Picture         =   "Form1.frx":11994
      Top             =   3120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image21 
      Height          =   1545
      Left            =   15120
      Picture         =   "Form1.frx":11FF8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image20 
      Height          =   1545
      Left            =   14040
      Picture         =   "Form1.frx":12629
      Top             =   3120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image19 
      Height          =   1530
      Left            =   12960
      Picture         =   "Form1.frx":12C4F
      Top             =   3120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image18 
      Height          =   1560
      Left            =   11880
      Picture         =   "Form1.frx":1322F
      Top             =   3120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image17 
      Height          =   1545
      Left            =   10800
      Picture         =   "Form1.frx":13824
      Top             =   3120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image16 
      Height          =   1545
      Left            =   9720
      Picture         =   "Form1.frx":13DEA
      Top             =   3120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image15 
      Height          =   1515
      Left            =   8640
      Picture         =   "Form1.frx":14391
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   1545
      Left            =   7560
      Picture         =   "Form1.frx":148F2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image13 
      Height          =   1515
      Left            =   20520
      Picture         =   "Form1.frx":14E4E
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image12 
      Height          =   1560
      Left            =   19440
      Picture         =   "Form1.frx":15377
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image11 
      Height          =   1545
      Left            =   18360
      Picture         =   "Form1.frx":15CE0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image10 
      Height          =   1545
      Left            =   17280
      Picture         =   "Form1.frx":16610
      Top             =   4680
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image9 
      Height          =   1545
      Left            =   16200
      Picture         =   "Form1.frx":16FD1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image8 
      Height          =   1545
      Left            =   15120
      Picture         =   "Form1.frx":17638
      Top             =   4680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image7 
      Height          =   1545
      Left            =   14040
      Picture         =   "Form1.frx":17C64
      Top             =   4680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image6 
      Height          =   1545
      Left            =   12960
      Picture         =   "Form1.frx":1827A
      Top             =   4680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image5 
      Height          =   1515
      Left            =   11880
      Picture         =   "Form1.frx":18853
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image4 
      Height          =   1560
      Left            =   10800
      Picture         =   "Form1.frx":18DF8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   1545
      Left            =   9720
      Picture         =   "Form1.frx":193B9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   1545
      Left            =   8640
      Picture         =   "Form1.frx":1994F
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   7560
      Picture         =   "Form1.frx":19EAD
      Top             =   4680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Deck(52)
Dim HandCards1, HandCards2, HandCards3, HandCards4, HandCards5, ValueCards1, ValueCards2, NewHand, DealerHand
Dim CardImage, CardIndex, PlayerorNPC, HandCard, ValueCard
Dim GameStarted, GameEnded As Boolean
Dim Cash As Integer
Public Function NewCard(PlayerorNPC, CardIndex, HandCard)

For x = 1 To 52 Step 13

If CardIndex = x Then
ValueCard = 2
End If
If CardIndex = x + 1 Then
ValueCard = 3
End If
If CardIndex = x + 2 Then
ValueCard = 4
End If
If CardIndex = x + 3 Then
ValueCard = 5
End If
If CardIndex = x + 4 Then
ValueCard = 6
End If
If CardIndex = x + 5 Then
ValueCard = 7
End If
If CardIndex = x + 6 Then
ValueCard = 8
End If
If CardIndex = x + 7 Then
ValueCard = 9
End If
If CardIndex = x + 8 Then
ValueCard = 10
End If
If CardIndex = x + 9 Then
ValueCard = 10
End If
If CardIndex = x + 10 Then
ValueCard = 10
End If
If CardIndex = x + 11 Then
ValueCard = 10
End If
If CardIndex = x + 12 Then
ValueCard = 11
End If

Next x

If CardIndex <= 13 Then
CardSuit = 2
CardsSuit = "Diamonds"
ElseIf CardIndex >= 14 And CardIndex <= 26 Then
CardSuit = 1
CardsSuit = "Hearts"
ElseIf CardIndex >= 27 And CardIndex <= 39 Then
CardSuit = 3
CardsSuit = "Spades"
ElseIf CardIndex >= 40 And CardIndex <= 52 Then
CardSuit = 4
CardsSuit = "Clubs"
End If


If ValueCard & CardSuit = "111" Then
Image54.Picture = Image26.Picture
ElseIf ValueCard & CardSuit = "21" Then
Image54.Picture = Image14.Picture
ElseIf ValueCard & CardSuit = "31" Then
Image54.Picture = Image15.Picture
ElseIf ValueCard & CardSuit = "41" Then
Image54.Picture = Image16.Picture
ElseIf ValueCard & CardSuit = "51" Then
Image54.Picture = Image17.Picture
ElseIf ValueCard & CardSuit = "61" Then
Image54.Picture = Image18.Picture
ElseIf ValueCard & CardSuit = "71" Then
Image54.Picture = Image19.Picture
ElseIf ValueCard & CardSuit = "81" Then
Image54.Picture = Image20.Picture
ElseIf ValueCard & CardSuit = "91" Then
Image54.Picture = Image21.Picture
ElseIf ValueCard & CardSuit = "101" Then
Image54.Picture = Image22.Picture
ElseIf ValueCard & CardSuit = "112" Then
Image54.Picture = Image13.Picture
ElseIf ValueCard & CardSuit = "22" Then
Image54.Picture = Image1.Picture
ElseIf ValueCard & CardSuit = "32" Then
Image54.Picture = Image2.Picture
ElseIf ValueCard & CardSuit = "42" Then
Image54.Picture = Image3.Picture
ElseIf ValueCard & CardSuit = "52" Then
Image54.Picture = Image4.Picture
ElseIf ValueCard & CardSuit = "62" Then
Image54.Picture = Image5.Picture
ElseIf ValueCard & CardSuit = "72" Then
Image54.Picture = Image6.Picture
ElseIf ValueCard & CardSuit = "82" Then
Image54.Picture = Image7.Picture
ElseIf ValueCard & CardSuit = "92" Then
Image54.Picture = Image8.Picture
ElseIf ValueCard & CardSuit = "102" Then
Image54.Picture = Image9.Picture
ElseIf ValueCard & CardSuit = "113" Then
Image54.Picture = Image39.Picture
ElseIf ValueCard & CardSuit = "23" Then
Image54.Picture = Image27.Picture
ElseIf ValueCard & CardSuit = "33" Then
Image54.Picture = Image28.Picture
ElseIf ValueCard & CardSuit = "43" Then
Image54.Picture = Image29.Picture
ElseIf ValueCard & CardSuit = "53" Then
Image54.Picture = Image30.Picture
ElseIf ValueCard & CardSuit = "63" Then
Image54.Picture = Image31.Picture
ElseIf ValueCard & CardSuit = "73" Then
Image54.Picture = Image32.Picture
ElseIf ValueCard & CardSuit = "83" Then
Image54.Picture = Image33.Picture
ElseIf ValueCard & CardSuit = "93" Then
Image54.Picture = Image34.Picture
ElseIf ValueCard & CardSuit = "103" Then
Image54.Picture = Image35.Picture
ElseIf ValueCard & CardSuit = "114" Then
Image54.Picture = Image52.Picture
ElseIf ValueCard & CardSuit = "24" Then
Image54.Picture = Image40.Picture
ElseIf ValueCard & CardSuit = "34" Then
Image54.Picture = Image41.Picture
ElseIf ValueCard & CardSuit = "44" Then
Image54.Picture = Image42.Picture
ElseIf ValueCard & CardSuit = "54" Then
Image54.Picture = Image43.Picture
ElseIf ValueCard & CardSuit = "64" Then
Image54.Picture = Image44.Picture
ElseIf ValueCard & CardSuit = "74" Then
Image54.Picture = Image45.Picture
ElseIf ValueCard & CardSuit = "84" Then
Image54.Picture = Image46.Picture
ElseIf ValueCard & CardSuit = "94" Then
Image54.Picture = Image47.Picture
ElseIf ValueCard & CardSuit = "104" Then
Image54.Picture = Image48.Picture
End If

If HandCard = "HandCards1" Then
Image55.Picture = Image54.Picture
ElseIf HandCard = "HandCards2" Then
Image58.Picture = Image54.Picture
ElseIf HandCard = "HandCards3" Then
Image59.Picture = Image54.Picture
ElseIf HandCard = "HandCards4" Then
Image60.Picture = Image54.Picture
ElseIf HandCard = "HandCards5" Then
Image61.Picture = Image54.Picture
ElseIf HandCard = "DealerCards1" Then
'Image56.Picture = Image54.Picture
'Actual Card in invisible off screen card
Image65.Picture = Image54.Picture
'Face down
Image56.Picture = Image53.Picture
ElseIf HandCard = "DealerCards2" Then
Image57.Picture = Image54.Picture
ElseIf HandCard = "DealerCards3" Then
Image62.Picture = Image54.Picture
Image62.Visible = True
ElseIf HandCard = "DealerCards4" Then
Image63.Picture = Image54.Picture
Image63.Visible = True
ElseIf HandCard = "DealerCards5" Then
Image64.Picture = Image54.Picture
Image64.Visible = True
End If

NewCard = HandCard & "," & ValueCard & "," & CardsSuit

Image54.Picture = Image53.Picture

End Function
Public Function UpdateText(HandCard)
HandCard = Split(HandCard, ",")

If HandCard(0) = "HandCards1" Or HandCard(0) = "HandCards2" Or HandCard(0) = "HandCards3" Or HandCard(0) = "HandCards4" Or HandCard(0) = "HandCards5" Then
Text1.Text = Text1.Text & vbCrLf & "Player draws " & HandCard(1) & " of " & HandCard(2)
NewHand = NewHand + HandCard(1)
'Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
End If

If HandCard(0) = "DealerCards1" Or HandCard(0) = "DealerCards2" Or HandCard(0) = "DealerCards3" Or HandCard(0) = "DealerCards4" Or HandCard(0) = "DealerCards5" Then
'Text1.Text = Text1.Text & vbCrLf & "Dealer draws " & HandCard(1) & " of " & HandCard(2)
DealerHand = DealerHand + HandCard(1)
'Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
End If

If HandCard(0) = "DealerCards2" Or HandCard(0) = "DealerCards3" Or HandCard(0) = "DealerCards4" Or HandCard(0) = "DealerCards5" Then
Text1.Text = Text1.Text & vbCrLf & "Dealer draws " & HandCard(1) & " of " & HandCard(2)
End If
End Function

Private Sub Command1_Click()
If GameEnded = True Then
MsgBox "Start a New Game first!"
Else
If HandCards3 = Empty Then
    Image59.Visible = True
    HandCards3 = Int(Rnd * 52 + 1)
    If Deck(HandCards3) = False Then
    Deck(HandCards3) = True
    End If
    PlayerCard3 = NewCard(Player, HandCards3, "HandCards3")
    UpdateText (PlayerCard3)
ElseIf HandCards4 = Empty Then
    Image60.Visible = True
    HandCards4 = Int(Rnd * 52 + 1)
    If Deck(HandCards4) = False Then
    Deck(HandCards4) = True
    End If
    PlayerCard4 = NewCard(Player, HandCards4, "HandCards4")
    UpdateText (PlayerCard4)
ElseIf HandCards5 = Empty Then
    Image61.Visible = True
    HandCards5 = Int(Rnd * 52 + 1)
    If Deck(HandCards5) = False Then
    Deck(HandCards5) = True
    End If
    PlayerCard5 = NewCard(Player, HandCards5, "HandCards5")
    UpdateText (PlayerCard5)
End If
End If
End Sub

Private Sub Command2_Click()
If GameEnded = True Then
MsgBox "Start a New Game first!"
Else
If NewHand = 21 And HandCards3 = Empty Then
GameStarted = False
GameEnded = True
Text1.Text = Text1.Text & vbCrLf & "Player has BlackJack! Player Wins!"
MsgBox "Player Wins"
ElseIf NewHand = 21 Then
GameStarted = False
GameEnded = True
Text1.Text = Text1.Text & vbCrLf & "Player has 21! Player Wins!"
MsgBox "Player Wins"
End If

While DealerHand < 17 And GameStarted = True:
If DealerHand <= 17 Then

If DealerCards3 = Empty Then
    Image62.Visible = True
    DealerCards3 = Int(Rnd * 52 + 1)
    If Deck(DealerCards3) = False Then
    Deck(DealerCards3) = True
    ElseIf Deck(DealerCards3) = True Then
        While Deck(DealerCards3) = True:
            DealerCards3 = Int(Rnd * 52 + 1)
        Wend
        Deck(DealerCards3) = True
    End If
    DealerCard3 = NewCard(Player, DealerCards3, "DealerCards3")
    UpdateText (DealerCard3)
ElseIf DealerCards4 = Empty Then
    Image63.Visible = True
    DealerCards4 = Int(Rnd * 52 + 1)
    If Deck(DealerCards4) = False Then
    Deck(DealerCards4) = True
    ElseIf Deck(DealerCards4) = True Then
        While Deck(DealerCards4) = True:
            DealerCards4 = Int(Rnd * 52 + 1)
        Wend
        Deck(DealerCards4) = True
    End If
    DealerCard4 = NewCard(Player, DealerCards4, "DealerCards4")
    UpdateText (DealerCard4)
ElseIf DealerCards5 = Empty Then
    Image64.Visible = True
    DealerCards5 = Int(Rnd * 52 + 1)
    If Deck(DealerCards5) = False Then
    Deck(DealerCards5) = True
    ElseIf Deck(DealerCards5) = True Then
        While Deck(DealerCards5) = True:
            DealerCards5 = Int(Rnd * 52 + 1)
        Wend
        Deck(DealerCards5) = True
    End If
    DealerCard5 = NewCard(Player, DealerCards5, "DealerCards5")
    UpdateText (DealerCard5)
End If

End If

Wend

If NewHand >= 22 And GameStarted = True Then
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
Text1.Text = Text1.Text & vbCrLf & "Player has bust! Dealer Wins!"
GameStarted = False
GameEnded = True
MsgBox "Dealer Wins"
Image56.Picture = Image65.Picture
ElseIf DealerHand >= 22 And GameStarted = True Then
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has bust! Player Wins!"
GameStarted = False
GameEnded = True
MsgBox "Player Wins"
Image56.Picture = Image65.Picture
ElseIf NewHand > DealerHand And DealerHand <= 20 Then
    If GameStarted = True Then
        Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
        Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
        Text1.Text = Text1.Text & vbCrLf & "Player wins!"
        GameStarted = False
        GameEnded = True
        MsgBox "Player Wins"
        Image56.Picture = Image65.Picture
    End If
ElseIf DealerHand > NewHand And NewHand <= 20 Then
    If GameStarted = True Then
        Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
        Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
        Text1.Text = Text1.Text & vbCrLf & "Player has lost! Dealer Wins!"
        GameStarted = False
        GameEnded = True
        MsgBox "Dealer Wins"
        Image56.Picture = Image65.Picture
    End If
End If

If NewHand = DealerHand And GameStarted = True Then
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
Text1.Text = Text1.Text & vbCrLf & "Draw"
Image56.Picture = Image65.Picture
GameStarted = False
GameEnded = True
End If

End If
End Sub

Private Sub Command3_Click()
Randomize Timer
For x = 1 To 52
Deck(x) = False
Next x
NewHand = 0
DealerHand = 0
GameEnded = False
GameStarted = True
Image62.Visible = False
Image63.Visible = False
Image64.Visible = False

DealerCards3 = Empty
DealerCards4 = Empty
DealerCards5 = Empty

Image59.Visible = False
Image60.Visible = False
Image61.Visible = False

HandCards3 = Empty
HandCards4 = Empty
HandCards5 = Empty

Text1.Text = "VGS-BlackJack v" & Build
'1-13 D
'14-26 H
'27-39 S
'40-52 C

HandCards1 = Int(Rnd * 52 + 1)
Deck(HandCards1) = True
HandCards2 = Int(Rnd * 52 + 1)

If HandCards1 = HandCards2 Then
HandCards2 = Int(Rnd * 52 + 1)
Deck(HandCards2) = True
Else
Deck(HandCards2) = True
End If

DealerCards1 = Int(Rnd * 52 + 1)
If Deck(DealerCards1) = False Then
Deck(DealerCards1) = True
ElseIf Deck(DealerCards1) = True Then
        While Deck(DealerCards1) = True:
            DealerCards1 = Int(Rnd * 52 + 1)
        Wend
        Deck(DealerCards1) = True
End If

DealerCards2 = Int(Rnd * 52 + 1)
If Deck(DealerCards2) = False Then
Deck(DealerCards2) = True
ElseIf Deck(DealerCards2) = True Then
        While Deck(DealerCards2) = True:
            DealerCards2 = Int(Rnd * 52 + 1)
        Wend
        Deck(DealerCards2) = True
End If

PlayerCard1 = NewCard(Player, HandCards1, "HandCards1")
UpdateText (PlayerCard1)
PlayerCard2 = NewCard(Player, HandCards2, "HandCards2")
UpdateText (PlayerCard2)
DealerCard1 = NewCard(Dealer, DealerCards1, "DealerCards1")
UpdateText (DealerCard1)
DealerCard2 = NewCard(Dealer, DealerCards2, "DealerCards2")
UpdateText (DealerCard2)

End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Private Sub Form_Load()
Build = "0.03"
Form1.Caption = "VGS-BlackJack v" & Build
Text1.Text = "VGS-BlackJack v" & Build
GameStarted = False
GameEnded = True
Cash = 1000
End Sub

Private Sub Image13_Click()
'Ace of Diamonds
End Sub

Private Sub Image26_Click()
'Ace of Hearts
End Sub

Private Sub Image39_Click()
'Ace of Spades
End Sub

Private Sub Image52_Click()
'Ace of Clubs
End Sub

Private Sub Image53_Click()
'Face Down
End Sub

Private Sub Image55_Click()
'Player Hand 1
End Sub

Private Sub Image56_Click()
'Dealer Hand 1
End Sub

Private Sub Image57_Click()
'Dealer Hand 2
End Sub

Private Sub Image58_Click()
'Player Hand 2
End Sub

