VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cyril's SpaceQuest"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    
Private Sub Form_Load()
INITVars
'First Hide the Mouse
ShowCursor 0

'Initialise the DirectX Components
DXMain_Init

End Sub
'This randomly chooses music to play
Sub Main_PlayMusic()
Dim i As Integer

    Call DM_CreateLoaderPerformance(frmMain.hWnd)

    Randomize
    i = Int((3 * Rnd) + 1)

    If i = 1 Then
        Call DM_LoadPlayMidi("music.Mid")
    ElseIf i = 2 Then
        Call DM_LoadPlayMidi("music2.Mid")
    Else
        Call DM_LoadPlayMidi("Electric.Mid")
    End If

End Sub
Private Sub DXMain_Init()
'On Error GoTo errorout:

Set ddMain = DXMain.DirectDrawCreate("") 'Create an instance of DirectDraw
Set dsMain = DXMain.DirectSoundCreate("")
Me.Show 'Show the form

'Set the co-operative level of DirectX
ddMain.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
ddMain.SetDisplayMode 320, 240, 16, 0, DDSDM_DEFAULT 'Set The Screen Size

'Set up the primary screen surface to show on the screen
sdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
sdMain.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
sdMain.lBackBufferCount = 1

Set dsPrim = ddMain.CreateSurface(sdMain) 'This sets the primary surface

'Create a Backbuffer, to draw on in the background.
'This is used to increase animation and reduce flicker
Dim ddsCaps As DDSCAPS2
ddsCaps.lCaps = DDSCAPS_BACKBUFFER
Set dsBbuf = dsPrim.GetAttachedSurface(ddsCaps)

Do_SetStars 'This makes the Stars array and finalises attributes
DxMain_InitSurfaces 'Load The Surfaces

sAngle = 0 'The turning angle
sSpd = 1 'the speed multiplier

'The DirectInput Handler to control the star movements
Set diMain = DXMain.DirectInputCreate
Set diDev = diMain.CreateDevice("GUID_SysKeyboard")
diDev.SetCommonDataFormat DIFORMAT_KEYBOARD
diDev.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
diDev.Acquire

dsBbuf.SetForeColor vbRed
dsBbuf.SetFont Me.Font
Call DM_UnloadStopMidi
'Plays music
Call Main_PlayMusic

'This is the game loop, it's infinite and fast, however it's slow on older PCs
Do
    Do_Keys 'Check for keyboard input
    DXMain_Blit  'Draw The Screen
    DoEvents
Loop

'errorout:
'    DXMain_EndIT
End Sub

Private Sub DXMain_Blit()
'This is the main drawing routine
'All the stars are written onto a backbuffer and then the backbuffer is drawn
'onto the primary surface

'On Error GoTo errorout

Dim rback As RECT 'A rect is used to set the picture size
Dim Xas As Integer

dsBbuf.SetFillColor 0
dsBbuf.DrawBox 0, 0, 320, 240

'Set the Star Height to 6 pixels
rback.Top = 0: rback.Bottom = 3

'Draw and move the 150 stars
For Xas = 0 To 149
    'Define the picture used using the picture number set
    'In the array, this is cleaner faster and easier then
    'Using a single bitmap for each star
    rback.Left = sStar(Xas, 3) * 3
    rback.Right = rback.Left + 3
    
    'Draw the star onto the backbuffer
    dsBbuf.BltFast sStar(Xas, 0), sStar(Xas, 1), dsStar, rback, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    'Move the stars using sine and cosine
    sStar(Xas, 1) = sStar(Xas, 1) + (sStar(Xas, 2) * Cos(sAngle) * sSpd)
    sStar(Xas, 0) = sStar(Xas, 0) + (sStar(Xas, 2) * Sin(sAngle) * sSpd)
    
    'Check if the star is off the screen and if so, put them back
    If sStar(Xas, 1) < 0 Then sStar(Xas, 1) = 240 + sStar(Xas, 1)
    If sStar(Xas, 1) > 240 Then sStar(Xas, 1) = sStar(Xas, 1) - 240
    If sStar(Xas, 0) < 0 Then sStar(Xas, 0) = 320 + sStar(Xas, 0)
    If sStar(Xas, 0) > 320 Then sStar(Xas, 0) = sStar(Xas, 0) - 320
Next Xas

'Blit The Ship
ShipX = ShipX + ShipXShift
ShipY = ShipY + ShipYShift
If ShipX <= 0 Then ShipX = 0
If ShipX > 280 Then ShipX = 280
If ShipY < 0 Then ShipY = 0
If ShipY > 200 Then ShipY = 200
dsBbuf.BltFast ShipX, ShipY, dsShip, rShip, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
ShipXShift = 0
ShipYShift = 0

If (ShipX > StoneX And ShipX < StoneX + rStone.Right And ShipY >= StoneY And ShipY < StoneY + rStone.Bottom) Then
    DXMain_EndIT
End If
'Check if Missile Hits Stone
If Not (MissX > StoneX And MissX < StoneX + rStone.Right And MissY >= StoneY And MissY < StoneY + rStone.Bottom) Then
    'Blit The Missile
    If MissileVisi = True Then
        If MissX = 0 Then
            MissX = ShipX + 16
            MissY = ShipY - 12
        End If
        dsBbuf.BltFast MissX, MissY, dsMissile, rMissile, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        MissY = MissY - 6
        If MissY < 0 Then MissileVisi = False
    End If
    
    'STONE BLIT
    dsBbuf.BltFast StoneX, StoneY, dsStone, rStone, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    StoneY = StoneY + StoneSpeed
    If StoneY > 300 Then StoneY = 0
Else
    DoEvents
    MissileVisi = False
    StoneY = 0
    MissX = 0
    MissY = 0
    Randomize
    StoneX = Int((200 - 1 + 1) * Rnd + 1)
    Score = Score + 10
    StoneSpeed = StoneSpeed + 0.03
End If


'Blit The Text
dsBbuf.DrawText 3, 3, "Score: " & Score, False
dsBbuf.DrawText 225, 224, "Cyril's SpaceQuest", False

dsPrim.Flip Nothing, DDFLIP_WAIT

'errorout:
End Sub

Private Sub Do_Keys()
'This sub processes the DirectInput Commands
Dim Xas As Integer

diDev.GetDeviceStateKeyboard diState

If diState.Key(DIK_ESCAPE) <> 0 Then DXMain_EndIT

'The left key reduces the angle, if it gets below zero it's set to 2pi
If diState.Key(DIK_LEFT) <> 0 Then
    sAngle = sAngle + 0.025
    If sAngle > 6.28 Then sAngle = 0
    ShipXShift = -ShipSpeed
End If

'The right key increases the angle if its above 2pi its reduced to zero
If diState.Key(DIK_RIGHT) <> 0 Then
    sAngle = sAngle - 0.025
    If sAngle < 0 Then sAngle = 6.28
    ShipXShift = ShipSpeed
End If

'The down key reduces speed
If diState.Key(DIK_DOWN) Then
'    sSpd = sSpd * 0.99
    ShipYShift = ShipSpeed
End If

'The up key increases speed
If diState.Key(DIK_UP) <> 0 Then
'    sSpd = sSpd * 1.01
'    If sSpd > 25 Then sSpd = 25
    ShipYShift = -ShipSpeed
End If

If diState.Key(DIK_ADD) <> 0 Then
    ShipSpeed = ShipSpeed + 0.1
End If

If diState.Key(DIK_SUBTRACT) <> 0 Then
    ShipSpeed = ShipSpeed - 0.1
    If ShipSpeed <= 0 Then ShipSpeed = 1
End If

If Not MissileVisi Then
    If diState.Key(DIK_SPACE) <> 0 Then
        MissileVisi = True
        MissX = 0
        MissY = 0
    End If
    PlaySound App.Path & "\BLIP.wav", 0, SND_FILENAME Or SND_ASYNC
End If
End Sub




Private Sub DXMain_EndIT()
'This sub unloads DirectX and puts control back to the computer
ShowCursor 1
ddMain.RestoreDisplayMode 'Restores the old resolution
ddMain.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
diDev.Unacquire 'Disable directinput

DoEvents
DoEvents
MsgBox "That's the end." & vbCrLf & "If you liked this game vote for it."
End
End Sub

Private Sub DxMain_InitSurfaces()
'This sub loads the surface containing the five star pictures

'THE STARS INIT
Dim Clrkey As DDCOLORKEY 'Creates a color key
Dim Key2 As DDCOLORKEY 'Here's a Color Key
Dim Key3 As DDCOLORKEY
Clrkey.low = 0
Clrkey.high = 0 'Makes the stars transparent
sdStar.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
sdStar.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
sdStar.lWidth = 15
sdStar.lHeight = 3 'The size of the star
Set dsStar = ddMain.CreateSurfaceFromFile(App.Path & "\Stars.bmp", sdStar) 'Load the bitmap
dsStar.SetColorKey DDCKEY_SRCBLT, Clrkey 'Set The ColorKey

'THE SHIP INIT
sdShip.lFlags = DDSD_CAPS
sdShip.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsShip = ddMain.CreateSurfaceFromFile(App.Path & "\SHIP.BMP", sdShip)
rShip.Bottom = sdShip.lHeight
rShip.Right = sdShip.lWidth
Key2.low = 0
Key2.high = 0
dsShip.SetColorKey DDCKEY_SRCBLT, Key2

'THE missile INIT
sdMissile.lFlags = DDSD_CAPS
sdMissile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsMissile = ddMain.CreateSurfaceFromFile(App.Path & "\missil.BMP", sdMissile)
rMissile.Bottom = sdMissile.lHeight
rMissile.Right = sdMissile.lWidth
Key2.low = 0
Key2.high = 0
dsMissile.SetColorKey DDCKEY_SRCBLT, Key2


'THE Stone INIT
sdStone.lFlags = DDSD_CAPS
sdStone.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set dsStone = ddMain.CreateSurfaceFromFile(App.Path & "\stone.BMP", sdStone)
rStone.Bottom = sdStone.lHeight
rStone.Right = sdStone.lWidth
Key2.low = 0
Key2.high = 0
dsStone.SetColorKey DDCKEY_SRCBLT, Key2
End Sub

Private Sub Do_SetStars()
'In this routine the stars are loaded into an array
'all values are random

Dim Xas As Integer

For Xas = 0 To 149 'There are 150 Stars
    sStar(Xas, 0) = Int(Rnd * 320) 'The Verticals Pos
    sStar(Xas, 1) = Int(Rnd * 240) 'The Horizontal Pos
    sStar(Xas, 2) = Int(Rnd * 5) 'The Speed
    sStar(Xas, 3) = Fix(4 - Int(sStar(Xas, 2))) 'Calculate what star it is
Next Xas
' Note: The bitmap used with this sample uses 5 stars
' The first star is the brightest and the last one the
' Darkest. To add realism the fastest moving stars will
' Be the ones using the first picture because it is the
' Closest to calculate this we round the star speed and
' Invert the number...
End Sub

Private Sub INITVars()
ShipX = 200
ShipY = 200
ShipSpeed = 4
StoneSpeed = 1
End Sub
