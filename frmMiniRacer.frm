VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMiniRacer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MiniRacer"
   ClientHeight    =   9345
   ClientLeft      =   1020
   ClientTop       =   1800
   ClientWidth     =   11715
   DrawStyle       =   2  'Dot
   DrawWidth       =   2
   Icon            =   "frmMiniRacer.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMiniRacer.frx":08CA
   ScaleHeight     =   623
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   781
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   781
      ImageHeight     =   623
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":165164
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2C9A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CA120
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CA832
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CAF44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CB656
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CBD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CC47A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CCB8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CD29E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CD9B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CE0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CE7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CEEE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CF5F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2CFD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D0B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D1240
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D1952
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D2064
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D2776
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D2E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D359A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D3CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D43BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D4AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiniRacer.frx":2D51E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7305
      TabIndex        =   6
      Top             =   180
      Width           =   4035
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HiScores:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5535
      TabIndex        =   5
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblHiScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5520
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblHiScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5520
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblHiScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblHiScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblHiScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   5430
      TabIndex        =   7
      Top             =   90
      Width           =   1425
   End
End
Attribute VB_Name = "frmMiniRacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*             MiniRacer by Lars Holm Jensen
'*      Use arrows for stearing and Enter for new game
'*              larsholmjensen@hotmail.com

Dim Speed As Single
Dim Direction As Single
Dim TurnSpeed As Single
Dim numDirections As Integer
Dim Acceleration As Single
Dim posX As Single, posY As Single
Dim carWidth As Single
Dim tireWidth As Single
Dim rightPressed As Boolean
Dim upPressed As Boolean
Dim leftPressed As Boolean
Dim downPressed As Boolean
Dim conRight As Single
Dim conLeft As Single
Dim tireMarkColor As Long
Dim Crash As Boolean
Dim CrashSideRight As Boolean
Dim picCrash1 As Boolean
Dim JustStarted As Boolean
Dim StartTime As Single, BestTime As Single, GoalTime As Single
Dim passedHalf As Boolean
Dim myScores(4) As Single
Const twoPI = 6.28318530717959
Const halfPI = 1.5707963267949
Const sqrTwo = 1.4142135623731
Const myFile = "HiScore.dat"
Const MaskColor = &HFF00FF
Dim myPath As String
Dim oldLeftX As Single
Dim oldLeftY As Single
Dim oldRightX As Single
Dim oldRightY As Single

Private Sub Form_Activate()

'carWidth is used for collision detection
carWidth = 8
'Indirectly how far tiremarks are apart
tireWidth = 10
Speed = 0
'The car can drive in 24 different directions
numDirections = 24
TurnSpeed = twoPI / numDirections
Acceleration = 1
Direction = 0
'Start position
posX = ScaleWidth \ 2 - 14
posY = ScaleHeight \ 2
rightPressed = False
upPressed = False
leftPressed = False
downPressed = False
'These are used to make the car slow down after many consecutive turnkeys
conRight = 1
conLeft = 1
tireMarkColor = vbBlack
Crash = False
'First start timer when up-key pressed
JustStarted = True
'Check if we're cheating
passedHalf = False
BestTime = 100000
tmrRun.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 37
    rightPressed = True
Case 38
    upPressed = True
    'Time first count when up-key is pressed
    'So you can just turn 180 degrees if you want to drive the track the other way around
    If JustStarted Then
        StartTime = Timer
        JustStarted = False
    End If
Case 39
    leftPressed = True
Case 40
    downPressed = True
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 37
    rightPressed = False
Case 38
    upPressed = False
Case 39
    leftPressed = False
Case 40
    downPressed = False
Case 13 'Enter: New Game !
    newGame
End Select

End Sub


Private Sub newGame()

'Just initialize the vars for a new game
Refresh 'Remove tiremarks
Direction = 0
Speed = 0
posX = ScaleWidth \ 2 - 14
posY = ScaleHeight \ 2
rightPressed = False
upPressed = False
leftPressed = False
downPressed = False
conRight = 1
conLeft = 1
tireMarkColor = vbBlack
Crash = False
JustStarted = True
passedHalf = False
tmrRun.Enabled = True

End Sub

Private Sub Form_Load()
Dim t As Integer

'Load the HiScores
myPath = App.Path
If Right(myPath, 1) <> "\" Then myPath = myPath & "\"
If Dir(myPath & myFile) <> "" Then 'Check if file exists..
    Open myPath & myFile For Input As #1
    For t = 0 To 4
        Input #1, myScores(t)
    Next t
    Close #1
Else    'if not make one with 6, 9, 12, 15 and 18 seconds
    Open myPath & myFile For Output As #1
    For t = 0 To 4
        myScores(t) = 6 + 3 * t
        ' The replace thing is for countries with comma for decimalpoint
        Print #1, Replace(Str(myScores(t)), ",", ".")
    Next t
    Close #1
End If

'Put the HiScores in the tables
For t = 0 To 4
    lblHiScore(t).Caption = Str(t + 1) & ": " & myScores(t)
Next

End Sub

' Sub to save HiScores
Private Sub SaveScores()
Dim t As Integer

Open myPath & myFile For Output As #1
For t = 0 To 4
    ' The replace thing is for countries with comma for decimalpoint
    Print #1, Replace(Str(myScores(t)), ",", ".")
Next t
Close #1

End Sub

Private Sub tmrRun_Timer()
' I know it's ugly. Only reason is if car get out of form, takes time to check
On Error Resume Next

Dim myRightX As Single, myRightY As Single, myLeftX As Single, myLeftY As Single
Dim rightSide As Long, leftSide As Long
Dim t As Integer, c As Integer
Dim numCarPic As Integer

If Crash Then
    'If we have crashed then slow down
    Speed = Speed - Acceleration
    'And when speed is zero, stop timer
    If Speed < 0 Then Speed = 0: tmrRun.Enabled = False
    '(It wouldn't do any good if you could drive on..)
    upPressed = False
    ' Make some spinning motions
    If CrashSideRight Then
        Direction = Direction - TurnSpeed
        rightPressed = True
        'Make sure tiremarks show
        If conRight < 1.728 Then conRight = 1.728
    Else
        Direction = Direction + TurnSpeed
        leftPressed = True
        'Make sure tiremarks show
        If conLeft < 1.728 Then conLeft = 1.728
    End If
End If

If rightPressed Then
    'Make the turn
    Direction = Direction - TurnSpeed
    'High speed and turning makes tiremarks and slow you down
    If Speed > 8 Then
        Speed = Speed - (conRight - 1)
        If Speed < 0 Then Speed = 0
    End If
    conRight = conRight * 1.2
    'No matter how much you turn, don't stop the car completely
    If conRight > 6 Then conRight = 6
    If conRight > 2 And (Speed > 8 Or Crash) Then
        'Draw the tiremarks
        myLeftX = posX + Cos(Direction + 1.5 * halfPI) * tireWidth
        myLeftY = posY + Sin(Direction + 1.5 * halfPI) * tireWidth
        If conRight > 2.4 Then Line (myLeftX, myLeftY)-(oldLeftX, oldLeftY), tireMarkColor
        myRightX = posX + Cos(Direction - 1.5 * halfPI) * tireWidth
        myRightY = posY + Sin(Direction - 1.5 * halfPI) * tireWidth
        If conRight > 2.4 Then Line (myRightX, myRightY)-(oldRightX, oldRightY), tireMarkColor
        oldLeftX = myLeftX
        oldLeftY = myLeftY
        oldRightX = myRightX
        oldRightY = myRightY
    End If
Else
    'conRight is actually dualfunctioned, both consecutive turn counter and
    'slow down factor
    conRight = 1
End If

If upPressed Then
    'Just accelerate
    Speed = Speed + Acceleration
End If

If leftPressed Then
    'Make the turn
    Direction = Direction + TurnSpeed
    'High speed and turning makes tiremarks and slow you down
    If Speed > 8 Then
        Speed = Speed - (conLeft - 1)
        If Speed < 0 Then Speed = 0
    End If
    conLeft = conLeft * 1.2
    'No matter how much you turn, don't stop the car completely
    If conLeft > 6 Then conLeft = 6
    If conLeft > 2 And (Speed > 8 Or Crash) Then
        'Draw the tiremarks
        myLeftX = posX + Cos(Direction + 1.5 * halfPI) * tireWidth
        myLeftY = posY + Sin(Direction + 1.5 * halfPI) * tireWidth
        If conLeft > 2.4 Then Line (myLeftX, myLeftY)-(oldLeftX, oldLeftY), tireMarkColor
        myRightX = posX + Cos(Direction - 1.5 * halfPI) * tireWidth
        myRightY = posY + Sin(Direction - 1.5 * halfPI) * tireWidth
        If conLeft > 2.4 Then Line (myRightX, myRightY)-(oldRightX, oldRightY), tireMarkColor
        oldLeftX = myLeftX
        oldLeftY = myLeftY
        oldRightX = myRightX
        oldRightY = myRightY
    End If
Else
    'conLeft too is dualfunctioned, both consecutive turn counter and
    'slow down factor
    conLeft = 1
End If

If downPressed Then
    Speed = Speed - Acceleration
    'No reverse
    If Speed < 0 Then Speed = 0
End If

'If within the form then remove last pic of car..
If posX > 12 And posY > 12 Then Me.PaintPicture ImageList2.ListImages(1).Picture, posX - 12, posY - 12, 24, 24, posX - 12, posY - 12, 24, 24

'Add the new keypresses to the motion of the car
posX = posX + Cos(Direction) * Speed
posY = posY + Sin(Direction) * Speed

'Collision detection (Pretty simple)
rightSide = Point(posX + Cos(Direction + halfPI) * carWidth, posY + Sin(Direction + halfPI) * carWidth)
leftSide = Point(posX + Cos(Direction - halfPI) * carWidth, posY + Sin(Direction - halfPI) * carWidth)

'If pixels next to the car are green then we have crashed
If (rightSide = vbGreen) Then
    Crash = True
    CrashSideRight = False
End If
If (leftSide = vbGreen) Then
    Crash = True
    CrashSideRight = True
End If

'Make sure Direction stays positive
'Eventhough Cos() and Sin() don't care, I use Direction for getting the right pic of car
If Direction < 0 Then Direction = Direction + twoPI

numCarPic = 1 + Round((Direction * 24 / twoPI), 0) Mod 24
If Crash Then
    'Draw car as usual
    ImageList1.ListImages(numCarPic).Draw Me.hDC, posX - 12, posY - 12, imlTransparent
    'Put one of the crash pics on top
    If picCrash1 Then
        ImageList1.ListImages(26).Draw Me.hDC, posX - 12, posY - 12, imlTransparent
    Else
        ImageList1.ListImages(27).Draw Me.hDC, posX - 12, posY - 12, imlTransparent
    End If
    'Switch crash pic
    picCrash1 = Not picCrash1
Else
    'Just draw car
    ImageList1.ListImages(numCarPic).Draw Me.hDC, posX - 12, posY - 12, imlTransparent
End If

'Check if we have driven half way around
If posY > 0.75 * ScaleHeight And Abs(posX - ScaleWidth \ 2) < 100 Then
    passedHalf = True
End If

'GoalTime is both the timer we see when driving and the time of the race when finished
GoalTime = Round(Timer - StartTime, 6)

If JustStarted = False Then lblTime.Caption = "Time: " & GoalTime

'Check if we have reached the goal and we haven't cheated
If passedHalf And (Abs(posX - (ScaleWidth \ 2)) < Speed) And Abs(ScaleHeight \ 2 - posY - 5) < 55 Then
    tmrRun.Enabled = False
    'Check if any record has been broken, and update
    For t = 0 To 4
        If GoalTime < myScores(t) Then
            For c = 4 To t + 1 Step -1
                myScores(c) = myScores(c - 1)
            Next c
            myScores(t) = GoalTime
            For c = 0 To 4
                lblHiScore(c).Caption = Str(c + 1) & ": " & myScores(c)
            Next
            lblTime.Caption = "New HiScore: " & GoalTime & " seconds!"
            SaveScores
            Exit For
        Else
            lblTime.Caption = "Completed in " & GoalTime & " seconds!"
        End If
    Next t
End If

End Sub












