VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tile Scrolling Demo"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
' Smooth Scrolling Tiles demo. Try messing around
' with the values of the SCREEN_WIDTH,
' SCREEN_HEIGHT, and SCREEN_BITDEPTH constants to
' see what effect it has on your frame rates.
'
' As always, feel free to modify/steal/distribute
' this code as you see fit. Lucky don't care!
'
' - Lucky
' Lucky's VB Gaming Site
' http://members.home.net/theluckyleper
'*************************************************

Option Explicit
                                                
'DirectX variables
Dim mdx As New DirectX7                     'Grandpa!
Dim mdd As DirectDraw7                      'Daddy!
Dim msurfFront As DirectDrawSurface7        'Front surface (the screen)
Dim msurfBack As DirectDrawSurface7         'Flipping Surface
Dim msurfTiles As DirectDrawSurface7        'Our tileset surface

'Some constants
Const SCREEN_WIDTH = 640
Const SCREEN_HEIGHT = 480
Const SCREEN_BITDEPTH = 8
Const TILE_WIDTH = 32
Const TILE_HEIGHT = 32
Const SCROLL_SPEED = 1

'Our permanent rectangles
Dim mrectScreen As RECT                     'Rectangle the size of the screen

'Program flow variables
Dim mblnRunning As Boolean                  'Is the main loop still running?
Dim mlngTimer As Long                       'Our timer variable
Dim mintFPSCounter As Integer               'Our FPS counter
Dim mintFPS As Integer                      'Our FPS storage variable

'Tile-Scrolling variables
Dim mintX As Integer                        '"Player" X coordinate
Dim mintY As Integer                        '"Player" Y coordinate
Dim mbytMap(100, 100) As Byte               'Our map array

'Keyboard stuffs
Dim mblnLeftKey As Boolean
Dim mblnRightKey As Boolean
Dim mblnUpKey As Boolean
Dim mblnDownKey As Boolean
          
Private Sub Form_Load()

Dim ddsdMain As DDSURFACEDESC2
Dim ddsdFlip As DDSURFACEDESC2
Dim i As Integer
Dim j As Integer
    
    'Show the main form
    Me.Show
      
    'Initialize DirectDraw
    Set mdd = mdx.DirectDrawCreate("")
    
    'Set the cooperative level (Fullscreen exclusive)
    mdd.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    'Set the resolution
    mdd.SetDisplayMode SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_BITDEPTH, 0, DDSDM_DEFAULT

    'Describe the flipping chain architecture we'd like to use
    ddsdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdMain.lBackBufferCount = 1
    ddsdMain.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    
    'Create the primary surface
    Set msurfFront = mdd.CreateSurface(ddsdMain)
    
    'Create the backbuffer
    ddsdFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set msurfBack = msurfFront.GetAttachedSurface(ddsdFlip.ddsCaps)
    
    'Set the text colour for the backbuffer
    msurfBack.SetForeColor vbWhite
    msurfBack.SetFontTransparency True

    'Create our screen-sized rectangle
    mrectScreen.Bottom = SCREEN_HEIGHT
    mrectScreen.Right = SCREEN_WIDTH
    
    'Load our surfaces
    LoadSurfaces
    
    'Load the map array with random tiles
    Randomize
    For i = 0 To UBound(mbytMap, 1)
        For j = 0 To UBound(mbytMap, 2)
            mbytMap(i, j) = CByte(Rnd() * 3)
        Next j
    Next i
    
    'Set the initial player X,Y coords to the center
    mintX = (UBound(mbytMap, 1) * TILE_WIDTH) \ 2
    mintY = (UBound(mbytMap, 2) * TILE_HEIGHT) \ 2
    
    'Start the main loop!
    MainLoop
    
End Sub

Private Sub MainLoop()

    'Start the loop running
    mblnRunning = True

    Do While mblnRunning
        If LostSurfaces Then LoadSurfaces       'Check for and restore lost surfaces
        msurfBack.BltColorFill mrectScreen, 0   'Clear the backbuffer
        MoveScreen                              'Move the screen
        DrawTiles                               'Display our tiles
        FPS                                     'Count/display the FPS
        msurfFront.Flip Nothing, 0              'Flip!!!
        DoEvents                                'Let other events occur
    Loop
    
    'Unload everything
    Terminate

End Sub

Private Sub DrawTiles()

Dim i As Integer
Dim j As Integer
Dim rectTile As RECT    'Ahahahahha! You said...
Dim bytTileNum As Byte
Dim intX As Integer
Dim intY As Integer

    'Draw the tiles according to the map array
    For i = 0 To CInt(SCREEN_WIDTH / TILE_WIDTH)
        For j = 0 To CInt(SCREEN_HEIGHT / TILE_HEIGHT)
            'Calc X,Y coords for this tile's placement
            intX = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            intY = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            'Which tile do we display?
            bytTileNum = GetTile(intX, intY)
            'Get the rectangle
            GetRect bytTileNum, intX, intY, rectTile
            'Blit the tile
            msurfBack.BltFast intX, intY, msurfTiles, rectTile, DDBLTFAST_WAIT
        Next j
    Next i

End Sub

Private Function GetTile(intTileX As Integer, intTileY As Integer) As Integer

    'Return the value returned by the map array for the given tile
    GetTile = mbytMap((intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH, (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT)

End Function

Private Sub GetRect(bytTileNumber As Byte, ByRef intTileX As Integer, ByRef intTileY As Integer, ByRef rectTile As RECT)

    'Calc rect
    With rectTile
        .Left = 0
        .Right = TILE_WIDTH
        .Top = bytTileNumber * TILE_HEIGHT
        .Bottom = .Top + TILE_HEIGHT
    
    'Clip rect
        
        'If this tile is off the left side of the screen...
        If intTileX < 0 Then
            .Left = .Left - intTileX
            intTileX = 0
        End If
        'If this tile is off the top of the screen...
        If intTileY < 0 Then
            .Top = .Top - intTileY
            intTileY = 0
        End If
        'If this tile is off the right side of the screen...
        If intTileX + TILE_WIDTH > SCREEN_WIDTH Then .Right = .Right + (SCREEN_WIDTH - (intTileX + TILE_WIDTH))
        'If this tile is off the bottom of the screen...
        If intTileY + TILE_HEIGHT > SCREEN_HEIGHT Then .Bottom = .Bottom + (SCREEN_HEIGHT - (intTileY + TILE_HEIGHT))
    End With

End Sub

Private Sub MoveScreen()

    'Move screen
    If mblnDownKey = True Then mintY = mintY + SCROLL_SPEED
    If mblnUpKey = True Then mintY = mintY - SCROLL_SPEED
    If mblnLeftKey = True Then mintX = mintX - SCROLL_SPEED
    If mblnRightKey = True Then mintX = mintX + SCROLL_SPEED
    
    'Ensure we don't go off the edge, that'd cause an error!
    If mintX < SCREEN_WIDTH \ 2 Then mintX = SCREEN_WIDTH \ 2
    If mintX > UBound(mbytMap, 1) * TILE_WIDTH - SCREEN_WIDTH \ 2 Then mintX = UBound(mbytMap, 1) * TILE_WIDTH - SCREEN_WIDTH \ 2
    If mintY < SCREEN_HEIGHT \ 2 Then mintY = SCREEN_HEIGHT \ 2
    If mintY > UBound(mbytMap, 2) * TILE_HEIGHT - SCREEN_HEIGHT \ 2 Then mintY = UBound(mbytMap, 2) * TILE_HEIGHT - SCREEN_HEIGHT \ 2

End Sub

Private Sub FPS()

    'Count FPS
    If mlngTimer + 1000 <= mdx.TickCount Then
        mlngTimer = mdx.TickCount
        mintFPS = mintFPSCounter + 1
        mintFPSCounter = 0
    Else
        mintFPSCounter = mintFPSCounter + 1
    End If
    
    'Display FPS and text
    msurfBack.DrawText 0, 0, "Press ESC to exit, arrow keys move. Current FPS: " & mintFPS, False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Exit program on escape key
    If KeyCode = vbKeyEscape Then mblnRunning = False
    
    'Move screen
    If KeyCode = vbKeyUp Then mblnUpKey = True
    If KeyCode = vbKeyDown Then mblnDownKey = True
    If KeyCode = vbKeyLeft Then mblnLeftKey = True
    If KeyCode = vbKeyRight Then mblnRightKey = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    'Stop moving screen
    If KeyCode = vbKeyUp Then mblnUpKey = False
    If KeyCode = vbKeyDown Then mblnDownKey = False
    If KeyCode = vbKeyLeft Then mblnLeftKey = False
    If KeyCode = vbKeyRight Then mblnRightKey = False

End Sub

Private Sub LoadSurfaces()

Dim ddsdGeneric As DDSURFACEDESC2
    
    'Set up generic surface description
    ddsdGeneric.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdGeneric.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

    'Load our tileset
    ddsdGeneric.lHeight = 128
    ddsdGeneric.lWidth = 32
    Set msurfTiles = mdd.CreateSurfaceFromFile(App.Path & "\tileset.bmp", ddsdGeneric)

End Sub

Private Function ExclusiveMode() As Boolean

Dim lngTestExMode As Long
    
    'This function tests if we're still in exclusive mode
    lngTestExMode = mdd.TestCooperativeLevel
    
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
    
End Function

Public Function LostSurfaces() As Boolean

    'This function will tell if we should reload our bitmaps or not
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    
    'If we did lose our bitmaps, restore the surfaces and return 'true'
    DoEvents
    If LostSurfaces Then
        mdd.RestoreAllSurfaces
    End If
    
End Function

Private Sub Terminate()

    'Terminate the render loop
    mblnRunning = False

    'Restore resolution
    mdd.RestoreDisplayMode
    mdd.SetCooperativeLevel 0, DDSCL_NORMAL

    'Kill the surfaces
    Set msurfTiles = Nothing
    Set msurfBack = Nothing
    Set msurfFront = Nothing
    
    'Kill directdraw
    Set mdd = Nothing

    'Unload the form
    Unload Me

End Sub
