VERSION 5.00
Begin VB.Form hand 
   BorderStyle     =   0  'None
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   40
      Left            =   1080
      Top             =   2040
   End
End
Attribute VB_Name = "hand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this example is a one picture animation example
'by eran shahar (eranisme@hotmail.com)
'if you want any other example so send me
'you may use this example for your games
'if you wont use this you will eat a big 18 meters pie!
'please write me on your credit if you will use this. you dont need to but please do
'I hope this will help you
Option Explicit
'lets dim those stuff!
'dim directX stuff
Dim Dx As New DirectX7
Dim DDraw As DirectDraw7
Dim ddsPrimary As DirectDrawSurface7
Dim ddsBackBuffer As DirectDrawSurface7
Dim s640x480rect As RECT
'dim the hand
Dim hand As DirectDrawSurface7
Dim handR As RECT
'mor directX stuff
Dim ddsdPrimary As DDSURFACEDESC2
Dim ddsdBackbuffer As DDSURFACEDESC2
'is this on?
Dim running, work As Boolean
Private Sub DDCreateSurface(surface As DirectDrawSurface7, BmpPath As String, RECTvar As RECT, Optional TransCol As Integer = 0, Optional UseSystemMemory As Boolean = True)
'setting up the surface
Dim tempddsd As DDSURFACEDESC2
Set surface = Nothing
tempddsd.lFlags = DDSD_CAPS
If UseSystemMemory = True Then
tempddsd.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY Or DDSCAPS_OFFSCREENPLAIN
Else
tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
End If
Set surface = DDraw.CreateSurfaceFromFile(BmpPath, tempddsd)
RECTvar.Right = tempddsd.lWidth
RECTvar.Bottom = tempddsd.lHeight
Dim ddckColourKey As DDCOLORKEY
ddckColourKey.low = TransCol
ddckColourKey.high = TransCol
surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
End Sub
Private Sub Form_Click()
work = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'if you want to leave
If KeyCode = vbKeyEscape Then running = False
End Sub
Private Sub Form_Load()
running = True
On Error Resume Next
Set DDraw = Dx.DirectDrawCreate("")
Me.Show
DDraw.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
DDraw.SetDisplayMode 640, 480, 32, 0, DDSDM_DEFAULT
ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ddsdPrimary.lBackBufferCount = 1
Set ddsPrimary = DDraw.CreateSurface(ddsdPrimary)
Dim Caps As DDSCAPS2
Caps.lCaps = DDSCAPS_BACKBUFFER
Set ddsBackBuffer = ddsPrimary.GetAttachedSurface(Caps)
ddsBackBuffer.GetSurfaceDesc ddsdBackbuffer
'the hand
DDCreateSurface hand, App.Path & "\hand.bmp", handR
'the first picture of the hand (setting up the size)
handR.Left = 0
handR.Right = 20
'lets loop those stuff!
Do
DoEvents
'making the screen black and 640 on 480
ddsBackBuffer.BltColorFill s640x480rect, RGB(0, 0, 0)
If work = True Then ddsBackBuffer.BltFast 300, 300, hand, handR, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'its font time
ddsBackBuffer.SetFont Me.Font
ddsBackBuffer.SetForeColor vbWhite
'the font needed to bee without those () and with false in the end to work
ddsBackBuffer.DrawText 0, 100, "press on the mouse to start", False
ddsBackBuffer.DrawText 0, 200, "press esc to exit", False
'if you wont flip it , the backround of the form will be the same
ddsPrimary.Flip Nothing, DDFLIP_WAIT
'oh , we must say goodbay to our old loop
Loop Until running = False
Set ddsPrimary = Nothing
Set ddsBackBuffer = Nothing
DDraw.RestoreDisplayMode
DDraw.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
Unload Me
End Sub
Private Sub Timer_Timer()
'change the pictures
If work = True Then
'every picture size is 20 pixels
handR.Left = handR.Left + 20
handR.Right = handR.Right + 20
'dont let it out of the picture (120 pixels is the hole picture)
If handR.Right > 120 Then handR.Right = 20
If handR.Left > 100 Then handR.Left = 0
End If
End Sub
