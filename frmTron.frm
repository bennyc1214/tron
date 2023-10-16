VERSION 5.00
Begin VB.Form frmTron 
   BackColor       =   &H00000000&
   Caption         =   "Tron!"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLead 
      Interval        =   1
      Left            =   8880
      Top             =   7800
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   7800
      Width           =   3615
   End
   Begin VB.Timer tmrPowerUp 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   10680
      Top             =   7800
   End
   Begin VB.Timer tmrCounter2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10080
      Top             =   7800
   End
   Begin VB.Timer tmrCounter1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9480
      Top             =   7800
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Stop/Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   4815
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00000000&
      Height          =   7500
      Left            =   120
      ScaleHeight     =   496
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   0
      Top             =   120
      Width           =   7500
   End
   Begin VB.Label lblLead 
      Caption         =   "Who's in the lead?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label lblMessage 
      Caption         =   $"frmTron.frx":0000
      Height          =   2775
      Left            =   7680
      TabIndex        =   4
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Image imgTron 
      Height          =   3135
      Left            =   7800
      Picture         =   "frmTron.frx":02F6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblPlayer2Score 
      Caption         =   "Score(Player2): "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Label lblPlayer1Score 
      Caption         =   "Score(Player1):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   6720
      Width           =   4095
   End
End
Attribute VB_Name = "frmTron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Benny Chen
'Purpose: Tron Game
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Option Explicit
Dim intX As Integer
Dim intY As Integer
Dim intX1 As Integer
Dim intY1 As Integer
Dim intCounter As Integer
Dim intCounter2 As Integer
Dim strDirection As String
Dim strDirection2 As String
Dim intRandomX As Integer
Dim intRandomY As Integer
Dim intRandomX1 As Integer
Dim intRandomY1 As Integer


Private Sub cmdGo_Click()
If tmrCounter1.Enabled = True Then
    tmrCounter1.Enabled = False
Else
    tmrCounter1.Enabled = True
End If
If tmrCounter2.Enabled = True Then
    tmrCounter2.Enabled = False
Else
    tmrCounter2.Enabled = True
End If
If tmrPowerUp.Enabled = True Then
    tmrPowerUp.Enabled = False
Else
    tmrPowerUp.Enabled = True
End If
End Sub

Private Sub cmdReset_Click()
picOutput.Cls
intCounter = 0
intCounter2 = 0
strDirection = "D1"
strDirection2 = "U2"
intX = 100
intY = 100
intX1 = 400
intY1 = 400
intRandomX = 0
intRandomY = 0
intRandomX1 = 0
intRandomY1 = 0
lblLead.Caption = "Who's in the lead?"
picOutput.SetFocus
End Sub

Private Sub Form_Activate()
intCounter = 0
intCounter2 = 0
strDirection = "D1"
strDirection2 = "U2"
intX = 100
intY = 100
intX1 = 400
intY1 = 400
intRandomX = 0
intRandomY = 0
intRandomX1 = 0
intRandomY1 = 0
picOutput.SetFocus
End Sub
Private Sub picOutput_KeyPress(KeyAscii As Integer)
If KeyAscii = 119 Then ' W was pressed
    strDirection = "U1"
ElseIf KeyAscii = 100 Then ' D was pressed
    strDirection = "D1"
ElseIf KeyAscii = 97 Then ' A was pressed
    strDirection = "L1"
ElseIf KeyAscii = 115 Then ' S was pressed
    strDirection = "R1"
ElseIf KeyAscii = 114 Then ' R was pressed
    picOutput.Cls
    intCounter = 0
    intCounter2 = 0
    strDirection = "D1"
    strDirection2 = "U2"
    intX = 100
    intY = 100
    intX1 = 400
    intY1 = 400
    intRandomX = 0
    intRandomY = 0
    intRandomX1 = 0
    intRandomY1 = 0
    picOutput.SetFocus
End If
If KeyAscii = 111 Then ' O was pressed
    strDirection2 = "U2"
ElseIf KeyAscii = 107 Then ' K was pressed
    strDirection2 = "L2"
ElseIf KeyAscii = 108 Then ' L was pressed
    strDirection2 = "R2"
ElseIf KeyAscii = 59 Then ' ; was pressed
    strDirection2 = "D2"
ElseIf KeyAscii = 112 Then ' P was pressed
    If tmrCounter1.Enabled = True Then
        tmrCounter1.Enabled = False
    Else
        tmrCounter1.Enabled = True
    End If
    If tmrCounter2.Enabled = True Then
        tmrCounter2.Enabled = False
    Else
        tmrCounter2.Enabled = True
    End If
    If tmrPowerUp.Enabled = True Then
        tmrPowerUp.Enabled = False
    Else
        tmrPowerUp.Enabled = True
    End If
End If
End Sub



Private Sub tmrCounter1_Timer()
intCounter = intCounter + 1

If strDirection = "D1" Then
    intX = intX + 1
ElseIf strDirection = "L1" Then
    intX = intX - 1
ElseIf strDirection = "U1" Then
    intY = intY - 1
ElseIf strDirection = "R1" Then
    intY = intY + 1
End If
'===============
If CStr(picOutput.Point(intX, intY)) = vbRed Or CStr(picOutput.Point(intX, intY)) = vbBlue Or intX > 500 Or intX < 0 Or intY > 500 Or intY < 0 Then
    tmrCounter1.Enabled = False
    tmrCounter2.Enabled = False
    tmrPowerUp.Enabled = False
    picOutput.Circle (intX, intY), 5, vbBlue
    picOutput.Circle (intX, intY), 10, vbBlue
    picOutput.Circle (intX, intY), 15, vbBlue
    picOutput.Circle (intX, intY), 20, vbBlue
    lblLead.Caption = "Player 2 wins!"
    MsgBox ("Player 1 crashed, Player 2 wins!")
End If
'===============
If CStr(picOutput.Point(intX, intY)) = vbWhite Then
    picOutput.FillColor = vbBlack
    picOutput.FillStyle = vbSolid
    picOutput.Circle (intX, intY), 10, vbBlack
    intCounter = intCounter + 100
ElseIf CStr(picOutput.Point(intX, intY)) = vbGreen Then
    picOutput.FillColor = vbBlack
    picOutput.FillStyle = vbSolid
    picOutput.Circle (intX, intY), 100, vbBlack
End If
If intCounter2 > 500 And intCounter2 < 1000 Then
    tmrCounter2.Interval = 7
ElseIf intCounter2 > 1000 And intCounter2 < 1500 Then
    tmrCounter2.Interval = 5
ElseIf intCounter2 > 1500 And intCounter < 2000 Then
    tmrCounter2.Interval = 3
ElseIf intCounter2 > 2000 Then
    tmrCounter2.Interval = 1
End If
lblPlayer1Score.Caption = "Score(Player 1):" & intCounter
picOutput.PSet (intX, intY), vbRed
End Sub

Private Sub tmrCounter2_Timer()
intCounter2 = intCounter2 + 1
If strDirection2 = "D2" Then
    intX1 = intX1 + 1
ElseIf strDirection2 = "L2" Then
    intX1 = intX1 - 1
ElseIf strDirection2 = "U2" Then
    intY1 = intY1 - 1
ElseIf strDirection2 = "R2" Then
    intY1 = intY1 + 1
End If
If CStr(picOutput.Point(intX1, intY1)) = vbRed Or CStr(picOutput.Point(intX1, intY1)) = vbBlue Or intX1 > 500 Or intX1 < 0 Or intY1 > 500 Or intY1 < 0 Then
    tmrCounter2.Enabled = False
    tmrCounter1.Enabled = False
    tmrPowerUp.Enabled = False
    picOutput.Circle (intX1, intY1), 5, vbBlue
    picOutput.Circle (intX1, intY1), 10, vbBlue
    picOutput.Circle (intX1, intY1), 15, vbBlue
    picOutput.Circle (intX1, intY1), 20, vbBlue
    lblLead.Caption = "Player 1 Wins!"
    MsgBox ("Player 2 crashed, Player 1 wins")
End If
If CStr(picOutput.Point(intX1, intY1)) = vbWhite Then
    picOutput.FillColor = vbBlack
    picOutput.FillStyle = vbSolid
    picOutput.Circle (intX1, intY1), 10, vbBlack
    intCounter2 = intCounter2 + 100
ElseIf CStr(picOutput.Point(intX1, intY1)) = vbGreen Then
    picOutput.FillColor = vbBlack
    picOutput.FillStyle = vbSolid
    picOutput.Circle (intX1, intY1), 100, vbBlack
End If
If intCounter2 > 500 And intCounter2 < 1000 Then
    tmrCounter2.Interval = 7
ElseIf intCounter2 > 1000 And intCounter2 < 1500 Then
    tmrCounter2.Interval = 5
ElseIf intCounter2 > 1500 And intCounter < 2000 Then
    tmrCounter2.Interval = 3
ElseIf intCounter2 > 2000 Then
    tmrCounter2.Interval = 1
End If
lblPlayer2Score.Caption = "Score(Player 2):" & intCounter2
picOutput.PSet (intX1, intY1), vbBlue
End Sub

Private Sub tmrLead_Timer()
If intCounter > intCounter2 Then
    lblLead.Caption = "Player1 is in the lead!"
ElseIf intCounter = intCounter2 Then
    lblLead.Caption = "The two players are tied"
Else
    lblLead.Caption = "Player2 is in the lead!"
End If
End Sub

Private Sub tmrPowerUp_Timer()
intRandomX = Int(Rnd * 500) + 1
intRandomY = Int(Rnd * 500) + 1
intRandomX1 = Int(Rnd * 500) + 1
intRandomY1 = Int(Rnd * 500) + 1
picOutput.Circle (intRandomX, intRandomY), 3, vbWhite
picOutput.Circle (intRandomX1, intRandomY), 3, vbGreen
End Sub
