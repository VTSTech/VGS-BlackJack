VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "VGS-BlackJack v"
   ClientHeight    =   9465
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2280
      Max             =   5000
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   10
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0000C000&
      Caption         =   "Dealer Muck"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Caption         =   "Shuffle Each Hand"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Hand"
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
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   6240
      Width           =   6015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Bet: $100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Cards Left: 52"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Cash: $1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   5
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image Image65 
      Height          =   1500
      Left            =   6480
      Picture         =   "Form1.frx":0016
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image64 
      Height          =   1500
      Left            =   4800
      Picture         =   "Form1.frx":0EAD
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image63 
      Height          =   1500
      Left            =   3600
      Picture         =   "Form1.frx":14F5
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image62 
      Height          =   1500
      Left            =   2400
      Picture         =   "Form1.frx":1B3D
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image61 
      Height          =   1500
      Left            =   4800
      Picture         =   "Form1.frx":2185
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image60 
      Height          =   1500
      Left            =   3600
      Picture         =   "Form1.frx":27CD
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image59 
      Height          =   1500
      Left            =   2400
      Picture         =   "Form1.frx":2E15
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
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
      Left            =   2760
      TabIndex        =   3
      Top             =   9240
      Width           =   390
   End
   Begin VB.Image Image58 
      Height          =   1500
      Left            =   1200
      Picture         =   "Form1.frx":345D
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Image Image57 
      Height          =   1500
      Left            =   1200
      Picture         =   "Form1.frx":3AA5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1200
   End
   Begin VB.Image Image56 
      Height          =   1500
      Left            =   0
      Picture         =   "Form1.frx":40ED
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1200
   End
   Begin VB.Image Image55 
      Height          =   1500
      Left            =   0
      Picture         =   "Form1.frx":4735
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Image Image54 
      Height          =   1500
      Left            =   480
      Picture         =   "Form1.frx":4D7D
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Image Image53 
      Height          =   1500
      Left            =   6480
      Picture         =   "Form1.frx":53C5
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image52 
      Height          =   1500
      Left            =   20520
      Picture         =   "Form1.frx":5A0D
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image51 
      Height          =   1500
      Left            =   19440
      Picture         =   "Form1.frx":68A4
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image50 
      Height          =   1500
      Left            =   18360
      Picture         =   "Form1.frx":853A
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image49 
      Height          =   1500
      Left            =   17280
      Picture         =   "Form1.frx":A224
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image48 
      Height          =   1500
      Left            =   16200
      Picture         =   "Form1.frx":BB9B
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image47 
      Height          =   1500
      Left            =   15120
      Picture         =   "Form1.frx":D84C
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image46 
      Height          =   1500
      Left            =   14040
      Picture         =   "Form1.frx":F21C
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image45 
      Height          =   1500
      Left            =   12960
      Picture         =   "Form1.frx":10B5D
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image44 
      Height          =   1500
      Left            =   11880
      Picture         =   "Form1.frx":1216F
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image43 
      Height          =   1500
      Left            =   10680
      Picture         =   "Form1.frx":1364E
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image42 
      Height          =   1500
      Left            =   9600
      Picture         =   "Form1.frx":14999
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image41 
      Height          =   1500
      Left            =   8520
      Picture         =   "Form1.frx":15A72
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image40 
      Height          =   1500
      Left            =   7440
      Picture         =   "Form1.frx":16AD3
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image39 
      Height          =   1500
      Left            =   20520
      Picture         =   "Form1.frx":17912
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image38 
      Height          =   1500
      Left            =   19440
      Picture         =   "Form1.frx":1875E
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image37 
      Height          =   1500
      Left            =   18360
      Picture         =   "Form1.frx":1A2DC
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image36 
      Height          =   1500
      Left            =   17280
      Picture         =   "Form1.frx":1BF84
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image35 
      Height          =   1500
      Left            =   16200
      Picture         =   "Form1.frx":1D914
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image34 
      Height          =   1500
      Left            =   15120
      Picture         =   "Form1.frx":1F6C4
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image33 
      Height          =   1500
      Left            =   14040
      Picture         =   "Form1.frx":211A6
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image32 
      Height          =   1500
      Left            =   12960
      Picture         =   "Form1.frx":22AEE
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image31 
      Height          =   1500
      Left            =   11880
      Picture         =   "Form1.frx":24184
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image30 
      Height          =   1500
      Left            =   10800
      Picture         =   "Form1.frx":25729
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image29 
      Height          =   1500
      Left            =   9720
      Picture         =   "Form1.frx":26AA1
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image28 
      Height          =   1500
      Left            =   8640
      Picture         =   "Form1.frx":27BF0
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1500
      Left            =   7560
      Picture         =   "Form1.frx":28B8A
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image26 
      Height          =   1500
      Left            =   20520
      Picture         =   "Form1.frx":29956
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image25 
      Height          =   1500
      Left            =   19440
      Picture         =   "Form1.frx":2A71A
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image24 
      Height          =   1500
      Left            =   18360
      Picture         =   "Form1.frx":2C2D7
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1500
      Left            =   17280
      Picture         =   "Form1.frx":2DF30
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image22 
      Height          =   1500
      Left            =   16200
      Picture         =   "Form1.frx":2F8E2
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image21 
      Height          =   1500
      Left            =   15120
      Picture         =   "Form1.frx":31353
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image20 
      Height          =   1500
      Left            =   14040
      Picture         =   "Form1.frx":32AB9
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image19 
      Height          =   1500
      Left            =   12960
      Picture         =   "Form1.frx":340F0
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image18 
      Height          =   1500
      Left            =   11880
      Picture         =   "Form1.frx":354F5
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image17 
      Height          =   1500
      Left            =   10800
      Picture         =   "Form1.frx":36835
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image16 
      Height          =   1500
      Left            =   9720
      Picture         =   "Form1.frx":379DB
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image15 
      Height          =   1500
      Left            =   8640
      Picture         =   "Form1.frx":3896C
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image14 
      Height          =   1500
      Left            =   7560
      Picture         =   "Form1.frx":3986E
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image13 
      Height          =   1500
      Left            =   20520
      Picture         =   "Form1.frx":3A5B4
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image12 
      Height          =   1500
      Left            =   19440
      Picture         =   "Form1.frx":3B214
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image11 
      Height          =   1500
      Left            =   18360
      Picture         =   "Form1.frx":3CD55
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image10 
      Height          =   1500
      Left            =   17280
      Picture         =   "Form1.frx":3E9AE
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image9 
      Height          =   1500
      Left            =   16200
      Picture         =   "Form1.frx":4022A
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   1500
      Left            =   15120
      Picture         =   "Form1.frx":417C3
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image7 
      Height          =   1500
      Left            =   14040
      Picture         =   "Form1.frx":42B47
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image6 
      Height          =   1500
      Left            =   12960
      Picture         =   "Form1.frx":43E64
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   1500
      Left            =   11880
      Picture         =   "Form1.frx":44F97
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image4 
      Height          =   1500
      Left            =   10800
      Picture         =   "Form1.frx":46042
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image3 
      Height          =   1500
      Left            =   9720
      Picture         =   "Form1.frx":46FA3
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   8760
      Picture         =   "Form1.frx":47D4C
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   7560
      Picture         =   "Form1.frx":48AB9
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
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
Dim CardImage, CardIndex, PlayerOrNPC, HandCard, ValueCard, Build
Dim GameStarted, GameEnded, PlayerHasAce, DealerHasAce As Boolean
Dim Cash, CardsLeft, Bet, SliderVal As Integer
Public Function NewCard(PlayerOrNPC, CardIndex, HandCard)

If Deck(CardIndex) = False Then
Deck(CardIndex) = True
ElseIf Deck(CardIndex) = True Then
    While Deck(CardIndex) = True:
        CardIndex = Int(Rnd * 52 + 1)
        DoEvents
    Wend
    Deck(CardIndex) = True
End If

CardsLeft = CardsLeft - 1

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
'MsgBox "Player Total: " & NewHand & " Dealer Total: " & DealerHand
    If PlayerOrNPC = "Player" Then
        PlayerHasAce = True
         If NewHand >= 11 Then ValueCard = 1
    End If
    If PlayerOrNPC = "Dealer" Then
        DealerHasAce = True
         If DealerHand >= 11 Then ValueCard = 1
    End If

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
ElseIf ValueCard & CardSuit = "14" Then
Image54.Picture = Image52.Picture
ElseIf ValueCard & CardSuit = "13" Then
Image54.Picture = Image39.Picture
ElseIf ValueCard & CardSuit = "12" Then
Image54.Picture = Image13.Picture
ElseIf ValueCard & CardSuit = "11" Then
Image54.Picture = Image26.Picture
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
DoEvents
Image54.Picture = Image53.Picture

End Function
Public Function UpdateText(HandCard)
HandCard = Split(HandCard, ",")

If HandCard(0) = "HandCards1" Or HandCard(0) = "HandCards2" Or HandCard(0) = "HandCards3" Or HandCard(0) = "HandCards4" Or HandCard(0) = "HandCards5" Then
Text1.Text = Text1.Text & vbCrLf & "Player draws " & HandCard(1) & " of " & HandCard(2)

If (NewHand + HandCard(1) >= 22) And PlayerHasAce = True Then
NewHand = NewHand + HandCard(1) - 10
PlayerHasAce = False
Else
NewHand = NewHand + HandCard(1)
End If
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
    If NewHand >= 22 And PlayerHasAce = False Then
        'Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
        Text1.Text = Text1.Text & vbCrLf & "Player has bust! Dealer Wins!"
        GameStarted = False
        GameEnded = True
        MsgBox "Dealer Wins"
        If Check2.Value = 0 Then Image56.Picture = Image65.Picture
    End If
End If

If HandCard(0) = "DealerCards1" Or HandCard(0) = "DealerCards2" Or HandCard(0) = "DealerCards3" Or HandCard(0) = "DealerCards4" Or HandCard(0) = "DealerCards5" Then
'Text1.Text = Text1.Text & vbCrLf & "Dealer draws " & HandCard(1) & " of " & HandCard(2)
If (DealerHand + HandCard(1) >= 22) And DealerHasAce = True Then
DealerHand = DealerHand + HandCard(1) - 10
DealerHasAce = False
Else
DealerHand = DealerHand + HandCard(1)
End If
'Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
End If

If HandCard(0) = "DealerCards2" Or HandCard(0) = "DealerCards3" Or HandCard(0) = "DealerCards4" Or HandCard(0) = "DealerCards5" Then
Text1.Text = Text1.Text & vbCrLf & "Dealer draws " & HandCard(1) & " of " & HandCard(2)
End If
End Function

Private Sub About_Click()
MsgBox "VGS-BlackJack v" & Build & vbCrLf & "Written by Veritas (Nigel Todman)" & vbCrLf & "Open Source: https://github.com/Veritas83/VGS-BlackJack"
End Sub

Private Sub Command1_Click()
If GameEnded = True Or Cash < 0 Then
MsgBox "Start a New Game first!"
Else
If HandCards3 = Empty Then
    Image59.Visible = True
    HandCards3 = Int(Rnd * 52 + 1)
    PlayerCard3 = NewCard("Player", HandCards3, "HandCards3")
    UpdateText (PlayerCard3)
ElseIf HandCards4 = Empty Then
    Image60.Visible = True
    HandCards4 = Int(Rnd * 52 + 1)
    PlayerCard4 = NewCard("Player", HandCards4, "HandCards4")
    UpdateText (PlayerCard4)
ElseIf HandCards5 = Empty Then
    Image61.Visible = True
    HandCards5 = Int(Rnd * 52 + 1)
    PlayerCard5 = NewCard("Player", HandCards5, "HandCards5")
    UpdateText (PlayerCard5)
Else
Command2.Value = True
End If
End If
End Sub

Private Sub Command2_Click()
If GameEnded = True Or Cash < 0 Then
MsgBox "Start a New Game first!"
Else
If NewHand = 21 And HandCards3 = Empty Then
GameStarted = False
GameEnded = True
Text1.Text = Text1.Text & vbCrLf & "Player has BlackJack! Player Wins!"
MsgBox "Player Wins (+$25 Bonus)"
Cash = Cash + (Bet * 2) + 25
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet * 2) + 25 & " won!"
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
ElseIf NewHand = 21 Then
GameStarted = False
GameEnded = True
Text1.Text = Text1.Text & vbCrLf & "Player has 21! Player Wins!"
MsgBox "Player Wins"
Cash = Cash + (Bet * 2)
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet * 2) & " won!"
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
End If

While DealerHand < 18 And GameStarted = True:
If DealerHand <= 17 Then

If DealerCards3 = Empty Then
    Image62.Visible = True
    DealerCards3 = Int(Rnd * 52 + 1)
    DealerCard3 = NewCard("Dealer", DealerCards3, "DealerCards3")
    UpdateText (DealerCard3)
ElseIf DealerCards4 = Empty Then
    Image63.Visible = True
    DealerCards4 = Int(Rnd * 52 + 1)
    DealerCard4 = NewCard("Dealer", DealerCards4, "DealerCards4")
    UpdateText (DealerCard4)
ElseIf DealerCards5 = Empty Then
    Image64.Visible = True
    DealerCards5 = Int(Rnd * 52 + 1)
    DealerCard5 = NewCard("Dealer", DealerCards5, "DealerCards5")
    UpdateText (DealerCard5)
Else
GameStarted = False
GameEnded = True
MsgBox "Player Wins"
Cash = Cash + (Bet * 2)
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet * 2) & " won!"
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
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
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet) & " lost."
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
ElseIf DealerHand >= 22 And GameStarted = True Then
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has bust! Player Wins!"
GameStarted = False
GameEnded = True
MsgBox "Player Wins"
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet * 2) & " won!"
Cash = Cash + (Bet * 2)
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
ElseIf NewHand > DealerHand And DealerHand <= 20 Then
    If GameStarted = True Then
        Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
        Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
        Text1.Text = Text1.Text & vbCrLf & "Player wins!"
        GameStarted = False
        GameEnded = True
        MsgBox "Player Wins"
        Text1.Text = Text1.Text & vbCrLf & "$" & (Bet * 2) & " won!"
        Cash = Cash + (Bet * 2)
        If Check2.Value = 0 Then Image56.Picture = Image65.Picture
    End If
ElseIf DealerHand > NewHand And NewHand <= 20 Then
    If GameStarted = True Then
        Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
        Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
        Text1.Text = Text1.Text & vbCrLf & "Player has lost! Dealer Wins!"
        GameStarted = False
        GameEnded = True
        MsgBox "Dealer Wins"
        Text1.Text = Text1.Text & vbCrLf & "$" & (Bet) & " lost."
        If Check2.Value = 0 Then Image56.Picture = Image65.Picture
    End If
End If

If NewHand = DealerHand And GameStarted = True Then
Text1.Text = Text1.Text & vbCrLf & "Player has: " & NewHand
Text1.Text = Text1.Text & vbCrLf & "Dealer has: " & DealerHand
Text1.Text = Text1.Text & vbCrLf & "Push"
MsgBox "Push"
Text1.Text = Text1.Text & vbCrLf & "$" & (Bet) & " returned."
If Check2.Value = 0 Then Image56.Picture = Image65.Picture
Cash = Cash + Bet
GameStarted = False
GameEnded = True
End If

End If
End Sub

Private Sub Command3_Click()
Randomize Timer

If Cash <= 0 Then
Cash = 1000
MsgBox "$1000 added"
End If

If Check1.Value = 1 Or CardsLeft <= 10 Then
    If CardsLeft <= 10 Then MsgBox ("Not enough cards! Shuffling deck.")
        CardsLeft = 52
        For x = 1 To 52
        Deck(x) = False
        Next x
End If

If Bet <= 0 Then
MsgBox "Bet must be positive value ;)"
Bet = 1
End If
Cash = Cash - Bet
Label2.Caption = "Cash $" & Cash
NewHand = 0
DealerHand = 0
GameEnded = False
GameStarted = True
PlayerHasAce = False
DealerHasAce = False
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
PlayerCard1 = NewCard("Player", HandCards1, "HandCards1")
UpdateText (PlayerCard1)
HandCards2 = Int(Rnd * 52 + 1)
PlayerCard2 = NewCard("Player", HandCards2, "HandCards2")
UpdateText (PlayerCard2)
DealerCards1 = Int(Rnd * 52 + 1)
DealerCard1 = NewCard("Dealer", DealerCards1, "DealerCards1")
UpdateText (DealerCard1)
DealerCards2 = Int(Rnd * 52 + 1)
DealerCard2 = NewCard("Dealer", DealerCards2, "DealerCards2")
UpdateText (DealerCard2)

End Sub

Private Sub Command4_Click()
If GameStarted = True Then
MsgBox "Cannot change bet when hand is in progress!"
Else
Bet = Bet + 2
End If
End Sub

Private Sub Command5_Click()
If GameStarted = True Then
MsgBox "Cannot change bet when hand is in progress!"
Else
Bet = Bet - 1
End If
End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Private Sub Form_Load()
Build = "0.4.3"
Form1.Caption = "VGS-BlackJack v" & Build
Text1.Text = "VGS-BlackJack v" & Build
GameStarted = False
GameEnded = True
Cash = 1000
Bet = 100
CardsLeft = 52

Label2.Caption = "Cash $" & Cash
Timer1.Interval = 1000
Timer1.Enabled = True
HScroll1.Value = 2500
Bet = Bet - 2500
End Sub

Private Sub HScroll1_Change()

If HScroll1.Value > SliderVal Then
Bet = Bet + ((HScroll1.Value) - (SliderVal))
ElseIf HScroll1.Value < SliderVal Then
Bet = Bet - ((SliderVal) - (HScroll1.Value))
End If
SliderVal = HScroll1.Value

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

Private Sub Timer1_Timer()
Label4.Caption = "Bet: $" & Bet
Label3.Caption = "Cards Left: " & CardsLeft
Label2.Caption = "Cash $" & Cash
End Sub
