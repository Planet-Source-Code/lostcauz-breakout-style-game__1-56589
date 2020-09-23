VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "KES Break<>Out"
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4080
      Top             =   5640
   End
   Begin MSComctlLib.ImageList imWarp 
      Left            =   6960
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   200
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":995E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12FB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imDing 
      Left            =   6840
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C606
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imspBlock 
      Left            =   5520
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   24
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CF5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EB02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F8D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":206AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imvwall 
      Left            =   5640
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2147E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22CD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imhwall 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24522
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25D76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imGolden 
      Left            =   2400
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   32
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":275CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28E1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imGameOver 
      Left            =   2280
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   200
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":47B86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imPause 
      Left            =   2280
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   220
      ImageHeight     =   50
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6509A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imLife 
      Left            =   1560
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D1D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D52A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imBlock 
      Left            =   1560
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   24
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D87E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E652
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F426
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":701FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":70FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":71DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":72B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7394A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7471E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imPaddle 
      Left            =   1560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   24
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":754F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77946
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":79D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7E642
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imBall 
      Left            =   1560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80A96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   1920
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   44
      ImageHeight     =   53
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":82B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":846EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":86292
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":899E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B58A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8D132
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'BreakOut style game by lostcauz


Private Sub Form_Load()
  Dim i As Long
  Randomize
  'Generate the stars
  For i = LBound(Stars) To UBound(Stars)
    With Stars(i)
      .x = Me.ScaleWidth * Rnd + 1
      .y = Me.ScaleHeight * Rnd + 1
      .size = MaxSize * Rnd + 1
      .speed = MaxSpeed * Rnd + 1
      .Color = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
    End With
  Next i
  newGame 1
  Brk.currentlevel = 1
  Open App.Path & "\hiscores.txt" For Input As #5
    For i = 0 To 9
      Input #5, Brk.hiscores(i)
    Next i
  Close #5
  Timer1.Enabled = True
  'gcom = False
  'Call ttt
  ShowCursor 0
End Sub

Private Sub newGame(level As Integer)
  Dim i As Long, d As Long, f As Long
  
  Randomize
  zPaddle.y = Me.ScaleHeight - 50
  zPaddle.image = 1 + Int(Rnd * 5)
  zBall.x = 1 + Int(Rnd * Me.ScaleWidth)
  zBall.y = 1 + Int(Rnd * 60)
  brake = 5
  With Brk
    .lives = 3
    .calcDone = False
    .Pause = True
    If .gameOver Then .currentlevel = 1: .gameOver = False
  End With
  blocksgone = 0
  intWarpImage = 1
  
  For i = 1 To UBound(goldBlock)
    goldBlock(i).dead = True
    goldBlock(i).hits = 0
  Next i
  
  For i = 1 To UBound(Wall)
    Wall(i).dead = True
  Next i
  
  For i = 1 To UBound(vWall)
    vWall(i).dead = True
  Next i
  
  Select Case level
    Case 0
    Case 1                      'Level 1
      Brk.score = 0
      zBall.vMom = 1 + Int(Rnd * 5)
      zBall.hMom = 1 + Int(Rnd * 5)
      d = 100
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          .x = d: .y = f
          If .x > 600 Then
            d = 100
            f = f + 24: d = d + 48
            .x = d: .y = f
          End If
        End With
      Next i
      
    Case 2                      'Level 2
      zBall.vMom = 1 + Int(Rnd * 6)
      zBall.hMom = 1 + Int(Rnd * 6)
      d = 52
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If i = 3 Or i = 6 Or i = 10 Or i = 15 Or i = 21 _
            Or i = 28 Or i = 36 Or i = 45 Or i = 55 Or i = 66 Then
            d = 52
            f = f + 24: d = d + 48
          End If
          .x = d: .y = f
        End With
      Next i
      
    Case 3                      'Level 3
      With goldBlock(1)
        .dead = False: .x = 340: .y = 40
      End With
      zBall.vMom = 1 + Int(Rnd * 7)
      zBall.hMom = 1 + Int(Rnd * 7)
      d = 148
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If (i - 1) Mod 7 = 0 Then
            d = 148
            f = f + 24: d = d + 48
          End If
          .x = d: .y = f
        End With
      Next i
      
    Case 4                      'Level 4
      d = -47
      f = 80
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If i Mod 9 = 0 Then
            d = 1: f = f + 24
          End If
          .x = d: .y = f
        End With
      Next i
      
      d = 120
      f = 300
      For i = 1 To UBound(goldBlock)
        With goldBlock(i)
          .dead = False: .hits = 0
          d = d + 64
          .x = d: .y = f
        End With
      Next i
  
      d = -127
      f = 400
      For i = 1 To 3
        With Wall(i)
          .dead = False
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      zBall.vMom = 1 + Int(Rnd * 8)
      zBall.hMom = 1 + Int(Rnd * 8)
      
    Case 5                      'Level 5
      Dim k As Long
      k = Me.ScaleWidth / 2 - 72
      d = k
      f = 80
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If i = 2 Or i = 4 Or i = 7 Or i = 11 Or i = 16 _
            Or i = 22 Or i = 29 Or i = 37 Or i = 46 _
            Or i = 56 Or i = 67 Then
            k = k - 24
            d = k + 48: f = f + 24
          End If
          .x = d: .y = f
        End With
      Next i
      
      d = 50
      f = 50
      For i = 1 To UBound(goldBlock)
        With goldBlock(i)
          .dead = False
          .hits = 0
          d = d + 96
          .x = d: .y = f
        End With
      Next i
  
      d = 30
      f = 400
      For i = 1 To 3
        With Wall(i)
          .dead = False
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      zBall.vMom = 1 + Int(Rnd * 9)
      zBall.hMom = 1 + Int(Rnd * 9)
      
    Case 6                      'Level 6
      d = 240
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          .x = d: .y = f
          If .x > Me.ScaleWidth - 180 Then
            d = 240
            f = f + 24: d = d + 48
            .x = d: .y = f
          End If
        End With
      Next i
      
      d = 230
      f = -60
      For i = 1 To 4
        With vWall(i)
          .dead = False
          f = f + 128
          If i > 2 Then
            d = 660
          End If
          .x = d: .y = f
          If i = 2 Then f = -60
        End With
      Next i
      
      d = 130
      f = 420
      For i = 1 To 3
        With Wall(i)
          .dead = False
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      zBall.vMom = 1 + Int(Rnd * 10)
      zBall.hMom = 1 + Int(Rnd * 10)
      
    Case 7                      'Level 7
      Dim strlev(255) As String
      Dim curPos As Integer
      d = -47
      f = 100: z = 0
      Open App.Path & "\level7.txt" For Input As #7
        Do Until EOF(7) = True
          z = z + 1
          Input #7, strlev(z)
        Loop
      Close #7
      
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
tagain:
          curPos = curPos + 1
          If curPos Mod 16 = 0 Then f = f + 24: d = -47
          d = d + 48
          If strlev(curPos) = "B" Then
            .x = d: .y = f
          Else
            GoTo tagain
          End If
        End With
      Next i
      
      d = -48
      f = 236
      For i = 1 To UBound(goldBlock)
        With goldBlock(i)
          .dead = False
          .hits = 0
          d = d + 128
          .x = d: .y = f
        End With
      Next i
  
      d = 50
      f = 220
      For i = 1 To 7
        With Wall(i)
          .dead = False
          If i = 4 Then f = 400: d = 10
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      zBall.vMom = 1 + Int(Rnd * 10)
      zBall.hMom = 1 + Int(Rnd * 10)
      
    Case 8                      'Level 8
      With goldBlock(1)
        .dead = False
        .x = 340: .y = 40
      End With
      zBall.vMom = 1 + Int(Rnd * 7)
      zBall.hMom = 1 + Int(Rnd * 7)
      d = 148
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If (i - 1) Mod 7 = 0 Then
            d = 148
            f = f + 24: d = d + 48
          End If
          .x = d: .y = f
        End With
      Next i
      
      d = 180
      f = 0
      For i = 1 To 4
        With vWall(i)
          .dead = False
          f = f + 128
          If i > 2 Then d = 130
          .x = d: .y = f
          If i = 2 Then f = -28
        End With
      Next i
      
      d = 20
      f = 100
      For i = 1 To 3
        With Wall(i)
          .dead = False
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      zBall.vMom = 1 + Int(Rnd * 10)
      zBall.hMom = 1 + Int(Rnd * 10)
      
    Case 9                      'Level 9
      zBall.vMom = 1 + Int(Rnd * 10)
      zBall.hMom = 1 + Int(Rnd * 10)
      d = 52
      f = 100
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 48
          If i = 3 Or i = 6 Or i = 10 Or i = 15 Or i = 21 _
            Or i = 28 Or i = 36 Or i = 45 Or i = 55 _
            Or i = 66 Then
            d = 52
            f = f + 24: d = d + 48
          End If
          .x = d: .y = f
        End With
      Next i
    
      d = -127
      f = 400
      For i = 1 To 5
        With Wall(i)
          .dead = False
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
      d = 0
      f = 364
      For i = 1 To UBound(goldBlock)
        With goldBlock(i)
          .dead = False
          .hits = 0
          d = d + 128
          .x = d: .y = f
        End With
      Next i
      
    Case 10                     'Level 10
      zBall.vMom = 1 + Int(Rnd * 10)
      zBall.hMom = 1 + Int(Rnd * 10)
      d = 20
      f = 84
      For i = 1 To UBound(Block)
        With Block(i)
          .dead = False
          .image = 1 + Int(Rnd * 9)
          d = d + 100
          If i Mod 6 = 0 Then
            d = 20: f = f + 24
          End If
          .x = d: .y = f
        End With
      Next i
      
      d = 84
      f = 0
      For i = 1 To 12
        With vWall(i)
          .dead = False
          f = f + 128
          .x = d: .y = f
          If i Mod 2 = 0 Then f = 0: d = d + 100
        End With
      Next i
    Case 11                     'Game Completed
      finishGame
    Case Default
  End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
  zPaddle.x = x - 64
  If zPaddle.x < 0 Then zPaddle.x = 0
  If zPaddle.x + 128 > Form1.ScaleWidth Then zPaddle.x = Form1.ScaleWidth - 128
End Sub

Private Sub Timer1_Timer() 'Main Game loop
  Dim i As Long, d As Long
  'While gcom = False
  
  'DoEvents
  BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, 0, vbBlackness
  
  For i = 0 To UBound(Stars)
    'Move the star
    With Stars(i)
      .y = (.y Mod Me.ScaleHeight) + .speed
      .x = (.x Mod Me.ScaleWidth) + .speed
      'Relocate the X position
      If .y > Me.ScaleHeight Then
        .x = Me.ScaleWidth * Rnd + 1
        .speed = MaxSpeed * Rnd + 1
      End If
      'Set the color
      Me.FillColor = .Color
      Me.ForeColor = .Color
      'Draw the star
      Ellipse Me.hdc, .x, .y, .x + .size, .y + .size
    End With
  Next i
  
  'draw paddle
  imPaddle.ListImages.Item(zPaddle.image).Draw Me.hdc, zPaddle.x, zPaddle.y, imlTransparent
  
  With zBall
    'set new ball coordinates based on vertical and
    'horizontal momentum
    .x = .x + .hMom
    .y = .y + .vMom
  
    'ball/screen collision detection
    If (.x + 20) > Form1.ScaleWidth Then
      .x = Form1.ScaleWidth - 20
      .hMom = -.hMom            'reverse ball's direction
      SoundPlay 2
    ElseIf .y + 20 > Form1.ScaleHeight Then
      .x = 1 + Int(Rnd * Me.ScaleWidth)
      .y = 1 + Int(Rnd * 60)
      If Not Brk.Pause Then
        If Not Brk.gameOver Then
          Brk.lives = Brk.lives - 1
        End If
      End If
    End If
  End With
  
  'draw blocks
  For i = 1 To UBound(Block)
    'if ball hit block
    'eliminate block with explosion drawing and 'ding'
    'change momentum accordingly
    With Block(i)
      If .dead Then GoTo dblock
      'ball/block collision detection
      If zBall.y + 20 >= .y Then
        If zBall.y <= .y + 24 Then
          If zBall.x + 20 >= .x Then
            If zBall.x <= .x + 48 Then
              If Not Brk.Pause Then
                If Not Brk.gameOver Then
                  SoundPlay 1
                  imDing.ListImages.Item(1).Draw Me.hdc, _
                    .x, .y - 16, imlTransparent
                  .dead = True
                  blocksgone = blocksgone + 1
                  Brk.score = Brk.score + (1 + Rnd * 500)
                  zBall.hMom = 1 + Int(Rnd * 4) + Brk.currentlevel
                  zBall.vMom = -zBall.vMom
                  'special block
                  If i = 37 Then
                    dropSpecial = True
                    SoundPlay 6
                    specialBlock.y = .y
                    specialBlock.x = .x
                  End If
                  'check for clear screen
                  If blocksgone >= UBound(Block) Then
                    Dim counter As Integer, deadG As Integer
                    For counter = 1 To UBound(goldBlock)
                      If goldBlock(counter).dead Then deadG = deadG + 1
                    Next counter
                    If deadG >= UBound(goldBlock) Then
                      Brk.currentlevel = Brk.currentlevel + 1
                      'show a level announcement
                      newGame Brk.currentlevel: Exit Sub
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      imBlock.ListImages.Item(.image).Draw Me.hdc, .x, .y, imlTransparent
dblock:
    End With
  Next i
  
  Dim deadGb As Integer
  For i = 1 To UBound(goldBlock)
    With goldBlock(i)
      If .dead Then GoTo gtest
      'ball/goldblock collision detection
      If zBall.y + 20 >= .y Then
        If zBall.y <= .y + 32 Then
          If zBall.x + 20 >= .x Then
            If zBall.x <= .x + 64 Then
              If Not Brk.Pause Then
                If Not Brk.gameOver Then
                  SoundPlay 7
                  imDing.ListImages.Item(1).Draw Me.hdc, _
                    .x, .y - 16, imlTransparent
                  .hits = .hits + 1
                  Brk.score = Brk.score + (1 + Rnd * 1500)
                  zBall.hMom = 1 + Int(Rnd * 4) + Brk.currentlevel
                  zBall.vMom = -zBall.vMom
                  If .hits >= 5 Then
                    .dead = True
gtest:
                    deadGb = deadGb + 1
                    If deadGb >= UBound(goldBlock) Then
                      If blocksgone >= UBound(Block) Then
                        Brk.currentlevel = Brk.currentlevel + 1
                        'show a level announcement
                        newGame Brk.currentlevel: Exit Sub
                      End If
                    End If
                    GoTo gblock
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      imGolden.ListImages.Item(2).Draw Me.hdc, .x, .y, imlTransparent
gblock:
    End With
  Next i
  
  For i = 1 To UBound(Wall)
    With Wall(i)
      If .dead Then GoTo hwall
      'ball/horizontal wall collision detection
      If zBall.y + 20 >= .y Then
        If zBall.y <= .y + 16 Then
          If zBall.x + 20 >= .x Then
            If zBall.x <= .x + 128 Then
              If Not Brk.Pause Then
                If Not Brk.gameOver Then
                  SoundPlay 2
                  imhwall.ListImages.Item(2).Draw Me.hdc, _
                    .x, .y - 16, imlTransparent
                  imhwall.ListImages.Item(2).Draw Me.hdc, _
                    .x, .y + 16, imlTransparent
                  zBall.vMom = -zBall.vMom
                End If
              End If
            End If
          End If
        End If
      End If
      imhwall.ListImages.Item(1).Draw Me.hdc, .x, .y, imlTransparent
hwall:
    End With
  Next i
  
  For i = 1 To UBound(vWall)
    With vWall(i)
      If .dead Then GoTo vrwall
      'ball/vertical wall collision detection
      If zBall.y + 20 >= .y Then
        If zBall.y <= .y + 128 Then
          If zBall.x + 20 >= .x Then
            If zBall.x <= .x + 16 Then
              If Not Brk.Pause Then
                If Not Brk.gameOver Then
                  SoundPlay 2
                  imvwall.ListImages.Item(2).Draw Me.hdc, _
                    .x - 16, .y, imlTransparent
                  imvwall.ListImages.Item(2).Draw Me.hdc, _
                    .x + 16, .y, imlTransparent
                  zBall.hMom = -zBall.hMom
                End If
              End If
            End If
          End If
        End If
      End If
      imvwall.ListImages.Item(1).Draw Me.hdc, .x, .y, imlTransparent
vrwall:
    End With
  Next i
  
  'special block is in play
  If dropSpecial Then
    If Brk.currentlevel = 4 Or Brk.currentlevel = 7 Then blnWarp = True
    With specialBlock
      .y = .y + 1
      If .y + 24 > zPaddle.y Then
        If .x + 48 > zPaddle.x Then
          If .x <= zPaddle.x + 128 Then
            'caught it
            SoundPlay 4
            zPaddle.image = 1 + Int(Rnd * 5)
            If Brk.lives < 3 Then Brk.lives = Brk.lives + 1
            Brk.score = Brk.score + Int(Rnd * (1000 * Brk.currentlevel))
            dropSpecial = False: blnWarp = False
          End If
        End If
        If .y >= Me.ScaleHeight Then: dropSpecial = False
      End If
      advanceSpecial = advanceSpecial + 1
      If advanceSpecial > 4 Then advanceSpecial = 1
      imspBlock.ListImages.Item(advanceSpecial).Draw Me.hdc, _
        .x, .y, imlTransparent
    End With
  End If
    
  'ball/paddle collision detection
  With zBall
    If (.y + 20) > zPaddle.y Then
      If (.x + 20) >= zPaddle.x Then
        If .x <= zPaddle.x + 128 Then
          If .x <= zPaddle.x + 32 Then
            .hMom = -.hMom
          End If
          SoundPlay 2
          imhwall.ListImages.Item(2).Draw Me.hdc, _
          zPaddle.x, zPaddle.y - 16, imlTransparent
          .vMom = -.vMom
        End If
      End If
    End If
  End With
  
  'draw lives
  With imLife.ListImages.Item(1)
    Select Case Brk.lives
      Case 0
        Brk.gameOver = True
      Case 1
        .Draw Me.hdc, 16, 16, imlTransparent
      Case 2
        .Draw Me.hdc, 16, 16, imlTransparent
        .Draw Me.hdc, 32, 16, imlTransparent
      Case 3
        .Draw Me.hdc, 16, 16, imlTransparent
        .Draw Me.hdc, 32, 16, imlTransparent
        .Draw Me.hdc, 48, 16, imlTransparent
      Case Default
    End Select
  End With
  
  'draw brakes -- brake slows ball
  With imLife.ListImages.Item(2)
    Select Case brake
      Case 0
      Case 1
        .Draw Me.hdc, 16, 32, imlTransparent
      Case 2
        .Draw Me.hdc, 16, 32, imlTransparent
        .Draw Me.hdc, 32, 32, imlTransparent
      Case 3
        For i = 1 To 3
          .Draw Me.hdc, (16 * i), 32, imlTransparent
        Next i
      Case 4
        For i = 1 To 4
          .Draw Me.hdc, (16 * i), 32, imlTransparent
        Next i
      Case 5
        For i = 1 To 5
          .Draw Me.hdc, (16 * i), 32, imlTransparent
        Next i
      Case Default
    End Select
  End With
  
  If Brk.gameOver Then
    imGameOver.ListImages.Item(1).Draw Me.hdc, Me.ScaleWidth / 2 - 100, Me.ScaleHeight / 2 - 100, imlTransparent
    If Not Brk.calcDone Then calcHiscore
    
    d = 320
    setupTextDraw 14, vbRed, 610, 760, d - 10, d + 20, "High scores"
    For i = 0 To 9
      d = d + 20
      setupTextDraw 12, vbGreen, 640, 740, d, d + 20, i + 1 & ": " & Str$(Brk.hiscores(i))
    Next i
  End If
  
  If Brk.Pause Then
    If Brk.gameOver Then
    Else
      imPause.ListImages.Item(1).Draw Me.hdc, 290, 220, imlTransparent
      
      'draw instruction text
      setupTextDraw 14, vbRed, 280, 390, 300, 330, " Left-Click -"
      setupTextDraw 14, vbRed, 270, 390, 330, 360, "Right-Click -"
      setupTextDraw 14, vbRed, 330, 390, 360, 390, "Esc -"
      setupTextDraw 14, vbRed, 330, 390, 390, 420, "N -"
      setupTextDraw 14, vbGreen, 400, 450, 300, 330, "Brake"
      setupTextDraw 14, vbGreen, 390, 550, 330, 360, "Pause/UnPause"
      setupTextDraw 14, vbGreen, 400, 440, 360, 390, "Exit"
      setupTextDraw 14, vbGreen, 400, 500, 390, 420, "New Game"
      
      'load the scores
      d = 320
      setupTextDraw 14, vbRed, 610, 760, d - 10, d + 20, "High scores"
      For i = 0 To 9
        d = d + 20
        setupTextDraw 12, vbGreen, 640, 740, d, d + 20, i + 1 & ": " & Str$(Brk.hiscores(i))
      Next i
    End If
  End If
  
  'Ensure ball isn't caught in corners
  With zBall
    If .x < 0 Then .x = 0: .hMom = -.hMom: SoundPlay 2
    If .y < 0 Then .y = 0: .vMom = -.vMom: SoundPlay 2
  End With
  
  'draw warp
  If blnWarp Then
    If zPaddle.x >= Me.ScaleWidth - 128 Then
      SoundPlay 5
      'show warp animation
      Brk.currentlevel = Brk.currentlevel + 2
      dropSpecial = False
      blnWarp = False
      'show a level announcement
      newGame Brk.currentlevel: Exit Sub
    End If
    If intWarpImage = 3 Then
      intWarpImage = 1
    Else
      intWarpImage = intWarpImage + 1
    End If
    imWarp.ListImages.Item(intWarpImage).Draw Me.hdc, Me.ScaleWidth - 64, Me.ScaleHeight - 150, imlTransparent
  End If
  
  'draw ball
  imBall.ListImages.Item(1).Draw Me.hdc, zBall.x, zBall.y, imlTransparent
  
  'score
  setupTextDraw 14, vbGreen, 100, 200, 16, 36, Str$(Brk.score)
  
  'level
  setupTextDraw 14, vbRed, 220, 280, 16, 36, "Level"
  setupTextDraw 14, vbGreen, 280, 310, 16, 36, Str$(Brk.currentlevel)
  
  'Update
  Me.Refresh
  'Wend
End Sub

Private Sub Form_Unload(cancel As Integer)
  ShowCursor 1
End Sub

Private Sub SoundPlay(PlayIt As Integer) ' Play sound
  'If Not soundOn Then Exit Sub
  If Brk.gameOver Or Brk.Pause Then Exit Sub
  Dim wFlags As Long, x As Long
  wFlags = SND_ASYNC Or SND_NODEFAULT
  DoEvents
  Select Case PlayIt
    Case 1
      x = sndPlaySound(App.Path & "\sound\ding.wav", wFlags)
    Case 2
      x = sndPlaySound(App.Path & "\sound\bat.wav", wFlags)
    Case 3
      x = sndPlaySound(App.Path & "\sound\ringout.wav", wFlags)
    Case 4
      x = sndPlaySound(App.Path & "\sound\mysound.wav", wFlags)
    Case 5
      x = sndPlaySound(App.Path & "\sound\warp.wav", wFlags)
    Case 6
      x = sndPlaySound(App.Path & "\sound\sound7.wav", wFlags)
    Case 7
      x = sndPlaySound(App.Path & "\sound\sound3.wav", wFlags)
    Case Default
  End Select
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
  If KeyCode = vbKeyN Then '110
    newGame 1
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
  With Brk
    If Button = 2 Then 'RightButton
      If .Pause Then
        .Pause = False
      Else
        .Pause = True
      End If
    ElseIf Button = 1 Then 'LeftButton
      If .Pause Or .gameOver Then Exit Sub
      If brake <= 0 Then
        brake = 0
      Else
        brake = brake - 1
        zBall.hMom = 3: zBall.vMom = 3
      End If
    End If
  End With
End Sub

Private Sub setupTextDraw(fs As Integer, c As ColorConstants, _
  recL As Long, recR As Long, recTop As Long, recB As Long, sPrintText As String)
  Dim lSuccess As Long
  Dim MyRect As RECT
  
  Me.Font.size = fs
  Me.ForeColor = c
  With MyRect
    .left = recL: .Right = recR
    .top = recTop: .Bottom = recB
  End With
  lSuccess = DrawText(Me.hdc, sPrintText, Len(sPrintText), _
    MyRect, DT_CENTER Or DT_WORDBREAK)
End Sub

Private Sub calcHiscore()
  Dim i As Long, j As Long, k As Long

  With Brk
    Open App.Path & "\hiscores.txt" For Input As #2
      For i = 0 To 9
        Input #2, .hiscores(i)
        If Not i = 0 Then
          If .hiscores(i) > .hiscores(i - 1) Then
            j = .hiscores(i - 1)
            .hiscores(i - 1) = .hiscores(i)
            .hiscores(i) = j
          End If
        End If
      Next i
    Close #2

    If .score > .hiscores(9) Then .hiscores(9) = .score
again:
    For i = 0 To 8
      j = .hiscores(i)
      k = .hiscores(i + 1)
      If k > j Then
        .hiscores(i) = k
        .hiscores(i + 1) = j
        GoTo again
      End If
    Next i
  
    Open App.Path & "\hiscores.txt" For Output As #3
      For i = 0 To 9
        Write #3, .hiscores(i)
      Next i
    Close #3
    Close
    .calcDone = True
  End With
End Sub

Private Sub finishGame()
  Dim i As Long
  Randomize
  'Generate the stars
  For i = LBound(rstar) To UBound(rstar)
    With rstar(i)
      .x = Me.ScaleWidth / 2
      .y = Me.ScaleHeight * Rnd + 1
      .size = MaxSize * Rnd + 1
      .speed = MaxSpeed * Rnd + 1
      .Color = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
    End With
  Next i
  Timer1.Enabled = False
  'gcom = True
  Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer() 'game completed
  Dim i As Long, d As Long
  Dim scwover2 As Long, schover2 As Long
  
  scwover2 = Me.ScaleWidth / 2
  schover2 = Me.ScaleHeight / 2
  
  BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, 0, vbBlackness
  
  For i = 0 To UBound(rstar)
    'Move the star
    With rstar(i)
      .y = (.y Mod Me.ScaleHeight) - .speed
      If i Mod 2 = 0 Then
        .drift = 1
      Else
        .drift = -1
      End If
      .x = .x + (.speed * .drift)
   
      If .y < 0 Then
        .y = Me.ScaleHeight * Rnd + 1
        .x = scwover2
        .speed = MaxSpeed * Rnd + 1
      End If
      'Set the color
      Me.FillColor = .Color
      Me.ForeColor = .Color
      'Draw the star
      Ellipse Me.hdc, .x, .y, .x + .size, .y + .size
    End With
  Next i
  
  z = z + 1
  Select Case z
    Case 1 To 50
      letters(1).x = z
      letters(1).y = z
      iml.ListImages.Item(1).Draw Me.hdc, letters(1).x, letters(1).y, imlTransparent
    
    Case 51 To 100
      letters(2).x = z
      letters(2).y = 1 + Int(Rnd * Me.ScaleHeight / 3)
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      iml.ListImages.Item(1).Draw Me.hdc, letters(1).x, letters(1).y, imlTransparent
      iml.ListImages.Item(2).Draw Me.hdc, letters(2).x, letters(2).y, imlTransparent
    
    Case 101 To 150
      letters(3).x = Me.ScaleWidth - (z / 2)
      letters(3).y = 1 + Int(Rnd * (schover2))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      For i = 1 To 3
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 151 To 200
      letters(4).x = z
      letters(4).y = 1 + Int(Rnd * (Me.ScaleHeight / 3))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      For i = 1 To 4
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 201 To 250
      letters(5).x = Me.ScaleWidth - (z / 3)
      letters(5).y = 1 + Int(Rnd * (Me.ScaleHeight / 3))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      letters(4).x = (scwover2 - 44) + Int(Rnd * 5)
      letters(4).y = schover2 + 2 + Int(Rnd * 4)
      For i = 1 To 5
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 251 To 300
      letters(6).x = Me.ScaleWidth - (z / 5)
      letters(6).y = 1 + Int(Rnd * (Me.ScaleHeight / 5))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      letters(4).x = (scwover2 - 44) + Int(Rnd * 5)
      letters(4).y = schover2 + 2 + Int(Rnd * 4)
      letters(5).x = scwover2 + Int(Rnd * 5)
      letters(5).y = schover2 + 1 + Int(Rnd * 4)
      For i = 1 To 6
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 301 To 350
      letters(7).x = z / 2
      letters(7).y = 1 + Int(Rnd * (Me.ScaleHeight / 5))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      letters(4).x = (scwover2 - 44) + Int(Rnd * 5)
      letters(4).y = schover2 + 2 + Int(Rnd * 4)
      letters(5).x = scwover2 + Int(Rnd * 5)
      letters(5).y = schover2 + 1 + Int(Rnd * 4)
      letters(6).x = scwover2 + 44 + Int(Rnd * 3)
      letters(6).y = schover2 + 1 + Int(Rnd * 3)
      For i = 1 To 7
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 351 To 400
      letters(8).x = Me.ScaleWidth - (z / 5)
      letters(8).y = 1 + Int(Rnd * (Me.ScaleHeight / 5))
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      letters(4).x = (scwover2 - 44) + Int(Rnd * 5)
      letters(4).y = schover2 + 2 + Int(Rnd * 4)
      letters(5).x = scwover2 + Int(Rnd * 5)
      letters(5).y = schover2 + 1 + Int(Rnd * 4)
      letters(6).x = scwover2 + 44 + Int(Rnd * 3)
      letters(6).y = schover2 + 1 + Int(Rnd * 3)
      letters(7).x = scwover2 + 88 + Int(Rnd * 3)
      letters(7).y = schover2 + 1 + Int(Rnd * 3)
      For i = 1 To 8
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
    
    Case 401 To 1000
      letters(1).x = (scwover2 - 176) + Int(Rnd * 3)
      letters(1).y = schover2 + Int(Rnd * 3)
      letters(2).x = (scwover2 - 132) + Int(Rnd * 3)
      letters(2).y = schover2 + 1 + Int(Rnd * 4)
      letters(3).x = (scwover2 - 88) + Int(Rnd * 3)
      letters(3).y = schover2 + Int(Rnd * 4) - 1
      letters(4).x = (scwover2 - 44) + Int(Rnd * 5)
      letters(4).y = schover2 + 2 + Int(Rnd * 4)
      letters(5).x = scwover2 + Int(Rnd * 5)
      letters(5).y = schover2 + 1 + Int(Rnd * 4)
      letters(6).x = scwover2 + 44 + Int(Rnd * 3)
      letters(6).y = schover2 + 1 + Int(Rnd * 3)
      letters(7).x = scwover2 + 88 + Int(Rnd * 3)
      letters(7).y = schover2 + 1 + Int(Rnd * 3)
      letters(8).x = scwover2 + 132 + Int(Rnd * 3)
      letters(8).y = schover2 + 1 + Int(Rnd * 3)
      For i = 1 To 8
        iml.ListImages.Item(i).Draw Me.hdc, letters(i).x, letters(i).y, imlTransparent
      Next i
      Me.FillColor = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
      Me.ForeColor = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
      Ellipse Me.hdc, scwover2 - 300, Me.ScaleHeight / 5, scwover2 + 300, schover2 + 220
      Me.FillColor = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
      Me.ForeColor = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
      Ellipse Me.hdc, scwover2 - 200, schover2 - 150, scwover2 + 200, schover2 + 200
    
    Case 1001 To 1500
      imGameOver.ListImages.Item(2).Draw Me.hdc, scwover2 - 100, schover2 - 100, imlTransparent
  
    Case Else
      z = 0
  End Select
  Me.Refresh
End Sub
