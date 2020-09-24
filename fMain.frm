VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "The Goat Problem"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btReset 
      Caption         =   "Reset"
      Height          =   525
      Left            =   240
      TabIndex        =   42
      Top             =   2880
      Width           =   630
   End
   Begin VB.CommandButton btStop 
      Caption         =   "Stop"
      Height          =   525
      Left            =   210
      TabIndex        =   1
      Top             =   1995
      Width           =   630
   End
   Begin VB.CommandButton btStart 
      Caption         =   "Start"
      Height          =   525
      Left            =   210
      TabIndex        =   0
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lbCarBehind 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   0
      Left            =   5760
      TabIndex        =   43
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Shape shp 
      Height          =   375
      Index           =   3
      Left            =   1080
      Top             =   2760
      Width           =   9255
   End
   Begin VB.Shape shp 
      Height          =   1695
      Index           =   2
      Left            =   1080
      Top             =   4800
      Width           =   9255
   End
   Begin VB.Shape shp 
      Height          =   1695
      Index           =   1
      Left            =   1080
      Top             =   3120
      Width           =   9255
   End
   Begin VB.Shape shp 
      Height          =   1695
      Index           =   0
      Left            =   1080
      Top             =   1080
      Width           =   9255
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   $"fMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Index           =   5
      Left            =   240
      TabIndex        =   41
      Top             =   120
      Width           =   10170
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   37
      Left            =   9840
      TabIndex        =   40
      Top             =   6120
      Width           =   120
   End
   Begin VB.Label lbPercent2Win 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8520
      TabIndex        =   39
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   36
      Left            =   9840
      TabIndex        =   38
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label lbPercent1Win 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8520
      TabIndex        =   37
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label lbWins2Total 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   36
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label lbChoices2Total 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   35
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   33
      Left            =   3240
      TabIndex        =   34
      Top             =   6120
      Width           =   450
   End
   Begin VB.Label lbWins1Total 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   33
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label lbChoices1Total 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   32
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   30
      Left            =   3240
      TabIndex        =   31
      Top             =   4440
      Width           =   450
   End
   Begin VB.Label lbCarBehindTotal 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   30
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   28
      Left            =   3720
      TabIndex        =   29
      Top             =   2415
      Width           =   450
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Win"
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
      Index           =   27
      Left            =   7560
      TabIndex        =   28
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Choice"
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
      Index           =   26
      Left            =   4920
      TabIndex        =   27
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lbPlayer2Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   2
      Left            =   6600
      TabIndex        =   26
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label lbPlayer2Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   1
      Left            =   6600
      TabIndex        =   25
      Top             =   5280
      Width           =   1290
   End
   Begin VB.Label lbPlayer2Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   0
      Left            =   6600
      TabIndex        =   24
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   2
      Left            =   6600
      TabIndex        =   23
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   1
      Left            =   6600
      TabIndex        =   22
      Top             =   3600
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Win 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   0
      Left            =   6600
      TabIndex        =   21
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label lbPlayer2Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   2
      Left            =   4200
      TabIndex        =   20
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label lbPlayer2Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   1
      Left            =   4200
      TabIndex        =   19
      Top             =   5280
      Width           =   1290
   End
   Begin VB.Label lbPlayer2Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   0
      Left            =   4200
      TabIndex        =   18
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 3"
      Height          =   315
      Index           =   16
      Left            =   3240
      TabIndex        =   17
      Top             =   5640
      Width           =   480
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 2"
      Height          =   315
      Index           =   15
      Left            =   3240
      TabIndex        =   16
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 1"
      Height          =   315
      Index           =   14
      Left            =   3240
      TabIndex        =   15
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 3"
      Height          =   315
      Index           =   13
      Left            =   3240
      TabIndex        =   14
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 2"
      Height          =   315
      Index           =   12
      Left            =   3240
      TabIndex        =   13
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Door 1"
      Height          =   315
      Index           =   11
      Left            =   3240
      TabIndex        =   12
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lbCarBehind 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   2
      Left            =   5760
      TabIndex        =   11
      Top             =   1935
      Width           =   1290
   End
   Begin VB.Label lbCarBehind 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   1
      Left            =   5760
      TabIndex        =   10
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   2
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   1290
   End
   Begin VB.Label lbPlayer1Door 
      Alignment       =   1  'Rechts
      Height          =   330
      Index           =   0
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Car behind Door 3"
      Height          =   315
      Index           =   4
      Left            =   3720
      TabIndex        =   6
      Top             =   1935
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Car behind Door 2"
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Car behind Door 1"
      Height          =   315
      Index           =   2
      Left            =   3720
      TabIndex        =   4
      Top             =   1215
      Width           =   1290
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   705
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'There are three doors in this game; CarBehind one door is a car, CarBehind the other two doors are goats.
'The player decides for one of the doors, whereupon the game host opens another door which is
'not the player's choice and which does not contain the car, and asks the player to decide again if he wants to.
'Player 1 sticks to his decision but Player 2 changes his decision now. Whose chances of winning the car are better?

'See http://en.wikipedia.org/wiki/Monty_Hall_problem

Option Explicit

Private CarBehind               As Long
Private CarBehinds(0 To 2)      As Long
Private CarBehindTotal          As Long
Private Player1Choice           As Long
Private Player1Choices(0 To 2)  As Long
Private Player2Choice           As Long
Private Player2Choices(0 To 2)  As Long
Private Player1Wins(0 To 2)     As Long
Private Wins1Total              As Long
Private Player2Wins(0 To 2)     As Long
Private Wins2Total              As Long
Private Done                    As Boolean
Private Unloading               As Boolean

Private Sub btReset_Click()

  Dim i   As Long

    Randomize -Time

    Done = False
    btStart.Enabled = True

    CarBehindTotal = 0
    Wins1Total = 0
    Wins2Total = 0
    lbPercent1Win = Format$(0, "#0.00")
    lbPercent2Win = Format$(0, "#0.00")
    lbCarBehindTotal = "0"
    lbWins1Total = "0"
    lbWins2Total = "0"
    lbChoices1Total = "0"
    lbChoices2Total = "0"

    For i = 0 To 2
        CarBehinds(i) = 0
        Player1Choices(i) = 0
        Player2Choices(i) = 0
        Player1Wins(i) = 0
        Player2Wins(i) = 0
        lbCarBehind(i) = "0"
        lbPlayer1Door(i) = "0"
        lbPlayer2Door(i) = "0"
        lbPlayer1Win(i) = 0
        lbPlayer2Win(i) = 0
    Next i

End Sub

Private Sub btStart_Click()

  Dim i   As Long

    Done = False
    btStart.Enabled = False
    Do

        CarBehind = Int(Rnd * 3)

        'choosing
        Player1Choice = Int(Rnd * 3)
        Player2Choice = Player1Choice

        'the host will now open a door which does not contain the car and is not the players
        'choice, and then asks the player to decide again.
        'player 1 sticks to his decision, but player 2 will thereupon change his mind
        'and select a door which is still closed and which is not his original choice.

        Select Case CarBehind
          Case 0                                        'Car  Goat  Goat
            If Player2Choice = 0 Then                   'player has chosen the car door
                Player2Choice = IIf(Rnd < 0.5, 1, 2)    'so host opens one goat door randomly and player chooses the other (not his original car door and not the opened goat door)
              Else 'NOT PLAYER2CHOICE...                'player has chosen one of the goat doors so host opens the other goat door
                Player2Choice = 0                       'and player changes his mind (not his original goat door and not the opened goat door)
            End If

          Case 1                                        'Goat  Car  Goat
            If Player2Choice = 1 Then                   'same as above
                Player2Choice = IIf(Rnd < 0.5, 0, 2)
              Else 'NOT PLAYER2CHOICE...
                Player2Choice = 1
            End If

          Case 2                                        'Goat  Goat  Car
            If Player2Choice = 2 Then                   'same as above
                Player2Choice = IIf(Rnd < 0.5, 0, 1)
              Else 'NOT PLAYER2CHOICE...
                Player2Choice = 2
            End If
        End Select

        'counting
        CarBehindTotal = CarBehindTotal + 1
        CarBehinds(CarBehind) = CarBehinds(CarBehind) + 1
        Player1Choices(Player1Choice) = Player1Choices(Player1Choice) + 1
        Player2Choices(Player2Choice) = Player2Choices(Player2Choice) + 1
        Player1Wins(Player1Choice) = Player1Wins(Player1Choice) - (Player1Choice = CarBehind)
        Player2Wins(Player2Choice) = Player2Wins(Player2Choice) - (Player2Choice = CarBehind)
        Wins1Total = Wins1Total - (Player1Choice = CarBehind)
        Wins2Total = Wins2Total - (Player2Choice = CarBehind)

        'update display
        If CarBehindTotal Mod 37 = 0 Then
            lbCarBehind(CarBehind) = CarBehinds(CarBehind)
            lbPlayer1Door(Player1Choice) = Player1Choices(Player1Choice)
            lbPlayer2Door(Player2Choice) = Player2Choices(Player2Choice)
            lbPlayer1Win(Player1Choice) = Player1Wins(Player1Choice)
            lbPlayer2Win(Player2Choice) = Player2Wins(Player2Choice)
            lbCarBehindTotal = CarBehindTotal
            lbChoices1Total = CarBehindTotal
            lbChoices2Total = CarBehindTotal
            lbWins1Total = Wins1Total
            lbWins2Total = Wins2Total
            lbPercent1Win = Format$(Wins1Total * 100 / CarBehindTotal, "#0.00")
            lbPercent2Win = Format$(Wins2Total * 100 / CarBehindTotal, "#0.00")
            DoEvents
        End If
    Loop Until Done Or CarBehindTotal >= 999999

    If Not Unloading Then
        For i = 0 To 2
            lbCarBehind(i) = CarBehinds(i)
            lbPlayer1Door(i) = Player1Choices(i)
            lbPlayer2Door(i) = Player2Choices(i)
            lbPlayer1Win(i) = Player1Wins(i)
            lbPlayer2Win(i) = Player2Wins(i)
        Next i
    End If

End Sub

Private Sub btStop_Click()

    Done = True
    btStart.Enabled = (CarBehindTotal < 999999)

End Sub

Private Sub Form_Load()

    btReset_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unloading = True 'prevents accessing the form again during unload
    Done = True

End Sub

':) Ulli's VB Code Formatter V2.24.21 (2009-Mrz-04 11:25)  Decl: 22  Code: 140  Total: 162 Lines
':) CommentOnly: 14 (8,6%)  Commented: 12 (7,4%)  Filled: 129 (79,6%)  Empty: 33 (20,4%)  Max Logic Depth: 4
