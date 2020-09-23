VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdIntoOrder 
      Caption         =   "Put Cards In Order"
      Enabled         =   0   'False
      Height          =   405
      Left            =   9135
      TabIndex        =   2
      Top             =   7470
      Width           =   1965
   End
   Begin VB.CommandButton cmdShuffle 
      Caption         =   "Shuffle The Pack"
      Height          =   405
      Left            =   7080
      TabIndex        =   1
      Top             =   7470
      Width           =   1965
   End
   Begin VB.CommandButton cmdFaceChange 
      Caption         =   "Turn Cards Face Down"
      Height          =   405
      Left            =   5010
      TabIndex        =   0
      Top             =   7470
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "label1"
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
      Left            =   10590
      TabIndex        =   3
      Top             =   7125
      Width           =   525
   End
   Begin VB.Image ICard 
      Height          =   1425
      Index           =   0
      Left            =   105
      Top             =   60
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CardInfo
CardName As String
CardValue As Integer
CardIndex As Integer
FaceUp As Boolean
End Type

'A formwide array to hold info for each card
Dim TheCardInfo(52) As CardInfo

Private Sub cmdIntoOrder_Click()
cmdIntoOrder.Enabled = False
PutCardsInOrder
CardsToFaceUp
End Sub

Private Sub cmdShuffle_Click()
ShuffleThePack
cmdIntoOrder.Enabled = True
End Sub

Private Sub Form_Load()
Load52Cards
PutCardsInOrder
CardsToFaceUp
Label1.Caption = "It's not a card game. Just a simple demo to show that card games can be written without using any OCX or DLL files"
End Sub

Private Sub cmdFaceChange_Click()
With cmdFaceChange
    If Len(.Tag) Then
    .Tag = vbNullString
    .Caption = "Turn Cards Face Down"
    CardsToFaceUp
    Else:
    .Tag = "0"
    CardsToFaceDown
    .Caption = "Turn Cards Face Up"
    End If
End With
End Sub



Private Sub Load52Cards()
'Load and position 52 images to hold pictures of the cards
Dim I As Integer
Dim TheLeft As Integer
Dim TheTop As Integer
TheTop = 400
TheLeft = 600
'Do The First Row
For I = 1 To 13
Load ICard(I)
ICard(I).Move TheLeft, TheTop
ICard(I).Visible = True
ICard(I).ZOrder
ICard(I).Tag = I
TheLeft = TheLeft + 800
Next
'Do the second row
TheLeft = 600
TheTop = 2100
For I = 14 To 26
Load ICard(I)
ICard(I).Move TheLeft, TheTop
ICard(I).Visible = True
ICard(I).ZOrder
TheLeft = TheLeft + 800
Next
'Do the third row
TheLeft = 600
TheTop = 3800
For I = 27 To 39
Load ICard(I)
ICard(I).Move TheLeft, TheTop
ICard(I).Visible = True
ICard(I).ZOrder
TheLeft = TheLeft + 800
Next
'Do the forth row
TheLeft = 600
TheTop = 5500
For I = 40 To 52
Load ICard(I)
ICard(I).Move TheLeft, TheTop
ICard(I).Visible = True
ICard(I).ZOrder
TheLeft = TheLeft + 800
Next
End Sub

Private Sub CardsToFaceDown()
'Put the face down picture in to all 52 image controls
Dim I As Integer
For I = 1 To 52
ICard(I).Picture = LoadResPicture(55, 0) 'Back of cards
TheCardInfo(I).FaceUp = False
Next
End Sub

Public Sub CardsToFaceUp()
Dim I As Integer
For I = 1 To 52
ICard(I).Picture = LoadResPicture(TheCardInfo(I).CardIndex, 0)
TheCardInfo(I).FaceUp = True
Next
End Sub

Private Sub ICard_Click(Index As Integer)
Dim S As String
Dim MsgReturn As Integer
Dim FaceUp As Boolean
FaceUp = IIf(TheCardInfo(Index).FaceUp, True, False)
S = "The Card you clicked on is the " & TheCardInfo(Index).CardName & vbCrLf
S = S & "It's Value is " & TheCardInfo(Index).CardValue & vbCrLf
If FaceUp Then
S = S & "It is Face up, Do you want to turn it face down"
Else:
S = S & "It is Face down, Do you want to turn it face up"
End If

MsgReturn = MsgBox(S, vbQuestion + vbYesNo, "Hello World.")

If MsgReturn = vbNo Then Exit Sub
If FaceUp Then
ICard(Index).Picture = LoadResPicture(55, 0)
TheCardInfo(Index).FaceUp = False
Else:
ICard(Index).Picture = LoadResPicture(TheCardInfo(Index).CardIndex, 0)
TheCardInfo(Index).FaceUp = True
End If
End Sub

Private Function GetCardValue(CardIndex As Integer) As Integer
'Ace = 1, 2 = 2, Jack = 11, Queen = 12 etc
Dim I As Integer
I = CardIndex Mod 13
I = IIf(I = 0, 13, I)
GetCardValue = I
End Function

Private Function GetCardSuit(CardIndex As Integer) As String
Select Case CardIndex
Case Is < 14: GetCardSuit = "Spades"
Case 14 To 26: GetCardSuit = "Diamonds"
Case 27 To 39: GetCardSuit = "Clubs"
Case Else: GetCardSuit = "Hearts"
End Select

End Function

Private Function GetCardType(CardIndex As Integer) As String
Dim I As Integer
I = GetCardValue(CardIndex)
Select Case I
Case 1: GetCardType = "Ace"
Case 2: GetCardType = "Two"
Case 3: GetCardType = "Three"
Case 4: GetCardType = "Four"
Case 5: GetCardType = "Five"
Case 6: GetCardType = "Six"
Case 7: GetCardType = "Seven"
Case 8: GetCardType = "Eight"
Case 9: GetCardType = "Nine"
Case 10: GetCardType = "Ten"
Case 11: GetCardType = "Jack"
Case 12: GetCardType = "Queen"
Case 13: GetCardType = "King"
End Select
End Function

Private Sub ShuffleThePack()
Dim I As Integer, N As Integer, T As Integer, G As Integer
Dim TempIndex As Integer
Dim LowNum As Integer
Randomize
'Create a 2 dimentional array
Dim TempArray(52, 1) As Integer

'Load the array with 52 random numbers between 0 and 1000
For I = 1 To 52
TempArray(I, 0) = I
TempArray(I, 1) = 1000 * Rnd
Next

'A very simple but slow sorting algoritham, but OK for 52 items to sort
For I = 1 To 52
LowNum = 1001
    For N = 1 To 52
        If TempArray(N, 1) < LowNum Then
        LowNum = TempArray(N, 1)
        T = TempArray(N, 0)
        End If
    Next
    TempArray(T, 1) = 1001
    TheCardInfo(I).CardIndex = T
    TheCardInfo(I).CardName = GetCardType(T) & " of " & GetCardSuit(T)
    TheCardInfo(I).CardValue = GetCardValue(T)
    TheCardInfo(I).FaceUp = TheCardInfo(T).FaceUp
Next

For I = 1 To 52
If TheCardInfo(I).FaceUp = True Then
ICard(I).Picture = LoadResPicture(TheCardInfo(I).CardIndex, 0)
Else:
ICard(I).Picture = LoadResPicture(55, 0)
End If
Next

End Sub

Private Sub PutCardsInOrder()
'Return the pack to the default order
Dim I As Integer
For I = 1 To 52
TheCardInfo(I).CardIndex = I
TheCardInfo(I).CardName = GetCardType(I) & " of " & GetCardSuit(I)
TheCardInfo(I).CardValue = GetCardValue(I)
Next
End Sub
