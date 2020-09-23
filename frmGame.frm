VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":7E94
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":FA1E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":10050
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":10682
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":1820C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":1FD96
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":21548
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":22CFA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":2738C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox eSubm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":2BA1E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eSub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":2DD88
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShipm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":300F2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":3245C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bombm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9960
      Picture         =   "frmGame.frx":347C6
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox bomb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9840
      Picture         =   "frmGame.frx":34948
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox eplanem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10320
      Picture         =   "frmGame.frx":34ACA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eplane 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10200
      Picture         =   "frmGame.frx":36E34
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9960
      Picture         =   "frmGame.frx":3919E
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9840
      Picture         =   "frmGame.frx":39230
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox playm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":392C2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox play 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":3B62C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "frmGame.frx":3D996
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000080FF&
         Caption         =   "About"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdTScore 
         BackColor       =   &H000080FF&
         Caption         =   "Top Score"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H000080FF&
         Caption         =   "New Game"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub cmdAbout_Click()
MsgBox ("Bomber v1.0 developed by Kevin Fleet" & vbCrLf & "Copyright(R) 2002 KevCom"), vbOKOnly, "About"
End Sub

Sub cmdEnd_Click()
End
End Sub

Sub cmdNew_Click()
Randomize
NewGame
End Sub

Function MainLoop()
Dim C As Long, AddS, AddSh, AddPl, I

MakeMenu
 
Do
If C > 1000 And GetTickCount > 500 Then
C = 0

If Pause > 0 Then
Pause = Pause - 1
Running = False
Message = "Paused"
If Pause = 1 Then Running = True
End If

If Running = True And Pause <= 0 Then

If AddS > 15 Then
AddS = 0
If Int(Rnd * 5) = 3 Then MakeSub
Else
AddS = AddS + 1
End If

If AddPl > 10 Then
AddPl = 0
If Int(Rnd * 5) = 3 Then MakePlane
Else
AddPl = AddPl + 1
End If

If AddSh > 15 Then
AddSh = 0
If Int(Rnd * 5) = 3 Then MakeShip
Else
AddSh = AddSh + 1
End If

CheckPlayerInput
MoveShots
MoveEnemys
MovePlayer
RunExplo

board.Cls

For I = 1 To 20
If PShot(I).Act = True Then
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If PlShot(I).Act = True Then
E.DrawObj board.hdc, PlShot(I).X, PlShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, PlShot(I).X, PlShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If ShShot(I).Act = True Then
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If I < 11 Then
If B(I).Act = True Then
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bombm.hdc, 0, 0, Mask
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bomb.hdc, 0, 0, sprite
End If
End If
Next I

For I = 1 To 10
If S(I).Act = True Then
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSubm.hdc, S(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSub.hdc, S(I).Dir * 50, 0, sprite
End If

If P(I).Act = True Then
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplanem.hdc, P(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplane.hdc, P(I).Dir * 50, 0, sprite
End If

If Sh(I).Act = True Then
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShipm.hdc, Sh(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShip.hdc, Sh(I).Dir * 50, 0, sprite
End If
Next I

E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, playm.hdc, Player.Dir * 50, 0, Mask
E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, play.hdc, Player.Dir * 50, 0, sprite

For I = 1 To 20
If Ex(I).Act = True Then
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explom(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, Mask
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explo(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, sprite
End If
Next I

board.CurrentX = 10
board.CurrentY = 10
board.Print "Score: " & Player.Score
board.CurrentX = 10
board.CurrentY = 20 + board.TextHeight("|")
board.Print "HP: "
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + 50, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbRed, BF
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + Player.HP, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbGreen, BF

Else
board.Cls

board.ForeColor = vbBlack
board.FontSize = 46
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(Message) \ 2
board.CurrentY = board.ScaleHeight \ 2 - (board.TextHeight("|") * 2)
board.Print Message

board.FontSize = 10

If E.IsPressed(vbKeyDown) Then Selected = Selected + 1: If Selected > 4 Then Selected = 4
If E.IsPressed(vbKeyUp) Then Selected = Selected - 1: If Selected < 1 Then Selected = 1
If E.IsPressed(vbKeyControl) Then DoItem

DrawMenu board
End If
Else
C = C + 1
End If

If E.IsPressed(vbKeyF3) Then
Select Case Pause
Case Is <= 0: Pause = 9999999
Case Is > 0: Pause = 0: Running = True
End Select
End If

DoEvents
Loop
End Function

Sub cmdTScore_Click()
Load frmTopScore
frmTopScore.Show
End Sub

Private Sub Form_Load()
Message = "Bomber"
Me.Visible = True
MainLoop
End Sub

Private Sub Form_Resize()
board.Left = 0
board.Top = 0
End Sub

