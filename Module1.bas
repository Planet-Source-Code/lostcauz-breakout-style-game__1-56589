Attribute VB_Name = "Module1"
Option Explicit

' API's
Public Declare Function ShowCursor Lib "user32" _
  (ByVal bShow As Long) As Long
  
Public Declare Function DrawText Lib "user32" Alias _
  "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
  ByVal nCount As Long, lpRect As RECT, _
  ByVal wFormat As Long) As Long

Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long, _
  ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long
  
Public Declare Function Rectangle Lib "gdi32" _
  (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
  ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function Ellipse Lib "gdi32" _
  (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
  ByVal X2 As Long, ByVal Y2 As Long) As Long
  
Public Declare Function sndPlaySound Lib "winmm.dll" _
  Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
  ByVal uFlags As Long) As Long
  
' Types
Private Type Game
  score             As Long
  hiscores(0 To 9)  As Long
  lives             As Integer
  currentlevel      As Integer
  gameOver          As Boolean
  Pause             As Boolean
  calcDone          As Boolean
  soundOn           As Boolean
End Type
  
Private Type Star
  x       As Long
  y       As Long
  speed   As Long
  size    As Long
  Color   As Long
  drift   As Long
End Type

Public Type RECT
  left    As Long
  top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type Blocks
  x       As Long
  y       As Long
  image   As Long
  hits    As Long
  dead    As Boolean
End Type

Private Type Letter
  x       As Long
  y       As Long
End Type

Private Type Ball
  x       As Long
  y       As Long
  vMom    As Integer         ' vertical momentum
  hMom    As Integer         ' horizontal momentum
End Type

Private Type Paddle
  x       As Long
  y       As Long
  image   As Integer
End Type

' Create Game
Public Brk                As Game

' Create Ball
Public zBall              As Ball

' Create Paddle
Public zPaddle            As Paddle

' Star field arrays
Public Stars(12)          As Star
Public rstar(300)         As Star

' Block arrays
Public Block(1 To 70)     As Blocks
Public goldBlock(1 To 5)  As Blocks

' Special block
Public specialBlock       As Blocks

' Horizontal Wall array
Public Wall(1 To 20)      As Blocks

' Vertical Wall array
Public vWall(1 To 20)     As Blocks

' Letter array
Public letters(1 To 8)    As Letter

' Text Draw Constants
Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

' Sound Constants
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

' Mouse Constants
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

' Misc.
Public Const MaxSize    As Long = 5
Public Const MaxSpeed   As Long = 4
Public z                As Long
Public blocksgone       As Integer
Public advanceSpecial   As Integer
Public brake            As Integer
Public intWarpImage     As Integer
Public dropSpecial      As Boolean
Public blnWarp          As Boolean
Public gcom             As Boolean
