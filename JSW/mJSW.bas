Attribute VB_Name = "mJSW"
'==============================================================================
'
' JET SET WILLY
'
' Based on the first release for the ZX Spectrum
' by Matthew Smith - Software Projects Ltd ©1984
'
' Author:        Carles P.V.
' Version:       1.2.0
' Date:          13-July-2006
'
' Last modified: 16-Jan-2012
'
'==============================================================================



Option Explicit

'-- API

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm" () As Long

'-- Data files (original mem. offsets)
                                        
Private Const DATA_ROPELUT              As Long = 33536 ' 256  bytes
Private Const DATA_TRIANGLES            As Long = 33841 ' 32     "
Private Const DATA_SCROLLY              As Long = 33876 ' 256    "
Private Const DATA_TTUNE                As Long = 34299 ' 100    "
Private Const DATA_GTUNE                As Long = 34399 ' 64     "
Private Const DATA_TITLECA              As Long = 38912 ' 512    "
Private Const DATA_PANELCA              As Long = 39424 ' 256    "
Private Const DATA_MAIN                 As Long = 39680 ' 25855  "

'-- Data offsets

Private Const OFFSET_SPRITES            As Long = 0
Private Const OFFSET_GUARDIANTABLE      As Long = 1280
Private Const OFFSET_OBJECTTABLE        As Long = 2303
Private Const OFFSET_ROOMDATA           As Long = 9472

'-- Room data offsets

Private Const OFFSET_SCREENDEF          As Long = &H0
Private Const OFFSET_ROOMNAME           As Long = &H80
Private Const OFFSET_BLOCKSDEF          As Long = &HA0
Private Const OFFSET_CONVEYORDEF        As Long = &HD6
Private Const OFFSET_SLOPEDEF           As Long = &HDA
Private Const OFFSET_BORDERCOLOR        As Long = &HDE
Private Const OFFSET_ITEMPATTERN        As Long = &HE1
Private Const OFFSET_ROOMEXITS          As Long = &HE9
Private Const OFFSET_GUARDIANSDEF       As Long = &HF0

'-- Color index constants

Private Const IDX_BLACK                 As Byte = 0
Private Const IDX_BLUE                  As Byte = 1
Private Const IDX_RED                   As Byte = 2
Private Const IDX_MAGENTA               As Byte = 3
Private Const IDX_GREEN                 As Byte = 4
Private Const IDX_CYAN                  As Byte = 5
Private Const IDX_YELLOW                As Byte = 6
Private Const IDX_WHITE                 As Byte = 7

Private Const IDX_BRBLACK               As Byte = 8
Private Const IDX_BRBLUE                As Byte = 9
Private Const IDX_BRRED                 As Byte = 10
Private Const IDX_BRMAGENTA             As Byte = 11
Private Const IDX_BRGREEN               As Byte = 12
Private Const IDX_BRCYAN                As Byte = 13
Private Const IDX_BRYELLOW              As Byte = 14
Private Const IDX_BRWHITE               As Byte = 15

Private Const IDX_NULL                  As Byte = 0
Private Const IDX_MASK                  As Byte = 255

'-- Quite important constants :-))

Private Const JSW_CHEATCODE             As String = "WRITETYPER"
Private Const JSW_INFO                  As String = "4361726C657320502E562E207F20323030362D32303132"

'-- Exit locations & max. height flags

Private Const XMIN                      As Byte = 0
Private Const XMAX                      As Byte = 31
Private Const YMIN                      As Byte = 0
Private Const YMAX                      As Byte = 14
Private Const DYDEATH                   As Byte = 36
Private Const YFDEATH                   As Byte = &HFF
 
'-- Block type *constants* (CAs)

Private BLOCK_AIR                       As Byte
Private BLOCK_FLOOR                     As Byte
Private BLOCK_WALL                      As Byte
Private BLOCK_NASTY                     As Byte
Private BLOCK_SLOPE                     As Byte
Private BLOCK_CONVEYOR                  As Byte

'-- Private enums.

Private Enum eMode
    [eLoading] = 0                      ' loading screen
    [eTitle] = 1                        ' title tune + scrolly
    [eGamePlay] = 2                     ' playing
    [eGameDie] = 3                      ' ouch!
    [eGameOver] = 4                     ' boot in action
End Enum

Private Enum eWillyMode
    [eWalk] = 0                         ' walking/standing (also on rope)
    [eJump] = 1                         ' jumping
    [eFall] = 2                         ' falling
End Enum

Private Enum eExit
    [eLeft] = 0
    [eRight] = 1
    [eUp] = 2
    [eDown] = 3
End Enum

'-- Private types

Private Type tWilly
    DIB(7)      As New cDIB08           ' sprites
    Ink         As Byte                 ' fore color
    x           As Byte                 ' current x pos [chrs]
    y           As Integer              ' current y pos [pxs]
    Dir         As Byte                 ' direction 0: right / 1: left
    Frame       As Byte                 ' current frame
    mode        As eWillyMode           ' 0: standing-walking / 1: jumping / 2: falling
    Flag        As Byte                 ' 0: standing / 1: right / 2: left
    f           As Integer              ' fall counter (-> max height)
    c           As Byte                 ' internal counter (-> jump)
    JustDied    As Boolean              ' Willy just died (skips first frame)
    OnConveyor  As Boolean              ' Willy is on a conveyor
    OnSlope     As Boolean              ' Willy is on a slope
    OnRope      As Boolean              ' Willy is on a rope
    RopeID      As Byte                 ' rope #
    RopeAnchor  As Byte                 ' rope anchor dot
    RopeExit    As Boolean              ' does rope have exit?
End Type

Private Type tSpecial
    DIB(3)      As New cDIB08           ' sprites
    x           As Byte                 ' current x pos [chrs]
    y           As Byte                 ' current x pos [chrs]
    Frame       As Byte                 ' current frame
    c           As Byte                 ' internal counter
End Type

Private Type tItem
    Room        As Byte
    Ink         As Byte                 ' fore color
    x           As Byte                 ' x pos [chrs]
    y           As Byte                 ' y pos [chrs]
    Flag        As Byte                 ' 0/1: collected
End Type

Private Type tNasty
    x           As Byte                 ' x pos [chrs]
    y           As Byte                 ' y pos [chrs]
End Type

Private Type tConveyor                  ' (animation)
    x           As Byte                 ' x pos [chrs] (left extreme)
    y           As Byte                 ' y pos [chrs]
    Dir         As Byte                 ' direction 1: right / 0: left
    Len         As Byte                 ' length [chrs]
End Type

Private Type tSlope
    x           As Byte                 ' x pos [chrs] (left extreme)
    y           As Byte                 ' y pos [chrs]
    Dir         As Byte                 ' direction 1: right / 0: left
    Len         As Byte                 ' length [chrs]
End Type

Private Type tGuardian
    DIB(7)      As New cDIB08           ' sprites
    Type        As Byte                 ' 0: horz. / 1: vert.
    Ink         As Byte                 ' fore color
    Min         As Byte                 ' top-left extreme [pxs]
    Max         As Byte                 ' bottom-right extreme [pxs]
    x           As Byte                 ' current x pos [chrs]
    y           As Integer              ' current y pos [pxs]
    Dir         As Byte                 ' direction (H)
    Speed       As Integer              ' speed (V)
    Fast        As Byte                 ' fast animation (V)
    Frame       As Byte                 ' current frame
    FrameI      As Byte                 ' initial frame
    FrameF      As Byte                 ' final frame
    c           As Byte                 ' internal counter
End Type

Private Type tPoint
    x           As Integer              ' x pos [pxs]
    y           As Integer              ' y pos [pxs]
End Type

Private Type tRope
    Dot()       As tPoint               ' rope dots
    Ink         As Byte                 ' fore color
    x           As Integer              ' anchor x pos [pxs]
    y           As Integer              ' anchor y pos [pxs]
    Dir         As Integer              ' current direction
    Len         As Byte                 ' lenght (dots)
    Swing       As Byte                 ' = range
    c           As Integer              ' internal counter
End Type

Private Type tArrow
    DIB         As New cDIB08           ' sprite
    x           As Byte                 ' current x pos [chrs]
    y           As Byte                 ' current y pos [chrs]
    Dir         As Byte                 ' direction 1: right / 0: left
End Type

Private Type tPanel
    Lives       As Byte                 ' Willy's lives
    t           As Long                 ' time counter
    AMPM        As String * 2           ' am/pm
    c           As Byte                 ' internal counter
End Type

'-- Private variables

Private m_oForm         As Form         ' destination Form
Private m_aDefaultPal() As Byte         ' default palette
Private m_aGreenPal()   As Byte         ' green palette
Private m_aBWPal()      As Byte         ' greyscale palette

Private m_oDIBBack      As New cDIB08   ' back buffer 1
Private m_oDIBFore      As New cDIB08   ' back buffer 2
Private m_oDIBMask      As New cDIB08   ' mask DIB (FX purposes)
Private m_oDIBFore2x    As New cDIB08   ' 2x screen
Private m_oDIBChar(95)  As New cDIB08   ' font char sprites

Private m_aData()       As Byte         ' main data block
Private m_eMode         As eMode        ' current mode
Private m_aRoomID       As Byte         ' current room #
Private m_aRoomExit(3)  As Byte         ' room exits

Private m_aBlockCA(5)   As Byte         ' CA for each block type
Private m_aRoomCA(511)  As Byte         ' room CA layout
Private m_aRoomFA(511)  As Byte         ' room FA layout
Private m_tPanel        As tPanel       ' panel (scores/lives)
Private m_aPanelCA()    As Byte         ' panel CA (data)

Private m_tSlope        As tSlope       ' slope
Private m_tConveyor     As tConveyor    ' conveyor (animation)
Private m_tNasty()      As tNasty       ' nasties

Private m_tItem()       As tItem        ' items
Private m_oDIBItem      As New cDIB08   ' item sprite
Private m_aItems        As Byte         ' items found
Private m_aItemsLeft    As Byte         ' items left

Private m_tGuardian()   As tGuardian    ' guardians
Private m_tRope()       As tRope        ' ropes
Private m_aRopeTable()  As Byte         ' rope offsets table
Private m_tArrow()      As tArrow       ' arrows

Private m_tWilly        As tWilly       ' Willy
Private m_tWillySafe    As tWilly       ' Willy *safe* state
Private m_bFlee         As Boolean      ' flag
Private m_bVomit        As Boolean      ' flag

Private m_oDIBFoot      As New cDIB08   ' foot sprite
Private m_oDIBBarrel    As New cDIB08   ' barrel sprite
Private m_oDIBPig(7)    As New cDIB08   ' pig sprites
Private m_tMaria        As tSpecial     ' Maria special sprites
Private m_tToilet       As tSpecial     ' toilet special sprites

Private m_aScrolly()    As Byte         ' title message
Private m_aTitleTune()  As Byte         ' title tune
Private m_aGameTune()   As Byte         ' in-game tune

'-- Misc. variables

Private m_aPow(7)       As Byte         ' quick 2^x
Private m_aNoteINV()    As Byte         ' LUT JSW-note
Private m_bKey(255)     As Boolean      ' LUT keys state

Private m_lFrameDt      As Long         ' current frame time interval (loop)
Private m_snFrameFactor As Single       ' time interval speed-factor
Private m_bPause        As Boolean      ' flag (loop)
Private m_bExit         As Boolean      ' flag (loop)

Private m_sCheatCode    As String * 10  ' cheat code string
Private m_bCheated      As Boolean      ' flag
Private m_bTune         As Boolean      ' flag
Private m_bFXTV         As Boolean      ' flag (FX TV scanlines)
Private m_bFXTVColor    As Byte         ' flag (FX green / black & white monitor)

Private m_hFXJump       As Long         ' sound FX handles
Private m_hFXDead       As Long
Private m_hFXPick       As Long
Private m_hFXArrow      As Long
Private m_hFXTone       As Long
Private m_hFXFX         As Long

Private m_hChannelN1    As Long         ' channel handles
Private m_hChannelN2    As Long
Private m_hChannelTune  As Long

Private m_nNoteLen      As Integer      ' notes length and value
Private m_nNote1        As Integer
Private m_nNote2        As Integer

Private m_lc0           As Long         ' counters
Private m_lc1           As Long
Private m_lc2           As Long

Private m_ltFPS         As Long         ' fps
Private m_lcFPS         As Long
Private m_lnFPS         As Long
Private m_bShowFPS      As Boolean

Private m_bJSWInfo      As Boolean      ' info
Private m_sJSWInfo      As String



'========================================================================================
' Main initialization
'========================================================================================

Public Sub Initialize(oForm As Form)
    
  Dim a() As Byte
  Dim c   As Integer
  Dim g   As Byte
    
    '-- Target Form ---------------------------------------------------------------------
    
    Set m_oForm = oForm
    
    '-- Fast 2^x (0-7)
    
    m_aPow(0) = 1
    For c = 1 To 7
        m_aPow(c) = 2 * m_aPow(c - 1)
    Next c
  
    '-- Palettes ------------------------------------------------------------------------
 
    '-- Load default palette
    
    m_aDefaultPal() = VB.LoadResData("DPAL", "ZX")
    
    '-- Greyscale / green palettes
    
    ReDim m_aGreenPal(1023)
    ReDim m_aBWPal(1023)
    
    For c = 0 To 4 * 31 Step 4
        
        g = (114& * m_aDefaultPal(c + 0) + _
             587& * m_aDefaultPal(c + 1) + _
             299& * m_aDefaultPal(c + 2) _
             ) \ 1000
        
        '-- Enlighten (fix blue tones)
        If (g > 0) Then
            If (g < 208) Then
                g = g + 48
              Else
                g = 255
            End If
          Else
            g = 16
        End If
            
        '-- Green tone
        m_aGreenPal(c + 1) = g
        
        '-- Grey *tone*
        m_aBWPal(c + 0) = g
        m_aBWPal(c + 1) = g
        m_aBWPal(c + 2) = g
    Next c
 
    '-- DIBs ----------------------------------------------------------------------------

    '-- Mask DIB (first two thirds of screen)
    Call m_oDIBMask.Create(256, 128)
    
    '-- Bback and fore DIBs (first and second buffers)
    Call m_oDIBBack.Create(256, 192)
    Call m_oDIBFore.Create(256, 192)
    
    '-- Screen DIB (2x zoomed + 20 pixels border)
    Call m_oDIBFore2x.Create(592, 464)
    Call m_oDIBFore2x.SetPalette(m_aDefaultPal())
    
    '-- Font chars ----------------------------------------------------------------------
 
    '-- Load packed chars
    a() = VB.LoadResData("FONT", "ZX")
    
    '-- Unpack
    For c = 0 To 95
        Call Unpack08(a(), 8 * c, m_oDIBChar(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Data ----------------------------------------------------------------------------

    '-- Title scrolly
    Call LoadData(DataFile(DATA_SCROLLY), m_aScrolly(), 0, 256)
    
    '-- Title and in-game tunes
    Call LoadData(DataFile(DATA_TTUNE), m_aTitleTune(), 0, 100)
    Call LoadData(DataFile(DATA_GTUNE), m_aGameTune(), 0, 64)
    
    '-- Title screen CA
    Call LoadData(DataFile(DATA_PANELCA), m_aPanelCA(), 0, 256)
    
    '-- Rope offsets table
    Call LoadData(DataFile(DATA_ROPELUT), m_aRopeTable(), 0, 256)
    
    '-- Main data
    Call LoadData(DataFile(DATA_MAIN), m_aData(), 0, 25855)
    
    '-- Sprites -------------------------------------------------------------------------
    
    '-- Willy
    With m_tWilly
        For c = 0 To 7
            Call Unpack16(m_aData(), 512 + 32 * c, .DIB(c), IDX_MASK, IDX_NULL)
        Next c
    End With
    
    '-- Maria (special)
    With m_tMaria
        For c = 0 To 3
            Call Unpack16(m_aData(), 384 + 32 * c, .DIB(c), IDX_MASK, IDX_NULL)
        Next c
        .x = 14 ' hardcoded
        .y = 11 ' hardcoded
    End With
    
    '-- Toilet (special)
    With m_tToilet
        For c = 0 To 3
            Call Unpack16(m_aData(), 2816 + 32 * c, .DIB(c), IDX_MASK, IDX_NULL)
        Next c
        .x = 28 ' hardcoded
        .y = 13 ' hardcoded
    End With
    
    '-- Foot
    Call Unpack16(m_aData(), 256 + 64, m_oDIBFoot, IDX_MASK, IDX_NULL)
    
    '-- Barrel
    Call Unpack16(m_aData(), 256 + 96, m_oDIBBarrel, IDX_MASK, IDX_NULL)
    
    '-- Pig
    For c = 0 To 3
        Call Unpack16(m_aData(), 7040 + 32 * c, m_oDIBPig(c + 0), IDX_MASK, IDX_NULL)
    Next c
    For c = 0 To 3
        Call Unpack16(m_aData(), 6912 + 32 * c, m_oDIBPig(c + 4), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Sound -------------------------------------------------------------------------
    
    '-- JSW LUT note
    m_aNoteINV() = LoadResData("NINV", "ZX")
    
    '-- Don't play tune
    m_bTune = False
    
    '-- FMOD and sound FXs
    Call FSOUND_Init(44100, 8, 0)
    m_hFXJump = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\jump.fx", FSOUND_NORMAL, 0, 0)
    m_hFXDead = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\dead.fx", FSOUND_NORMAL, 0, 0)
    m_hFXPick = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\pick.fx", FSOUND_NORMAL, 0, 0)
    m_hFXArrow = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\arrow.fx", FSOUND_NORMAL, 0, 0)
    m_hFXTone = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\tone.fx", FSOUND_NORMAL, 0, 0)
    m_hFXFX = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\fx.fx", FSOUND_NORMAL, 0, 0)
    
    '-- Info ----------------------------------------------------------------------------
    
    For c = 1 To Len(JSW_INFO) Step 2
        m_sJSWInfo = m_sJSWInfo & Chr$("&H" & Mid$(JSW_INFO, c, 2))
    Next c
End Sub

Public Sub Terminate()
    
    '-- Reset
    Call FXResetAll
    
    '-- Free all
    Call FSOUND_Sample_Free(m_hFXJump)
    Call FSOUND_Sample_Free(m_hFXDead)
    Call FSOUND_Sample_Free(m_hFXPick)
    Call FSOUND_Sample_Free(m_hFXArrow)
    Call FSOUND_Sample_Free(m_hFXTone)
    Call FSOUND_Sample_Free(m_hFXFX)
    
    '-- Close FMOD
    Call FSOUND_Close
End Sub

'========================================================================================
' Main loop
'========================================================================================

Public Sub StartGame()
  
  Dim t As Long
    
    '-- Intro mode
    Call SetMode([eLoading])
    
    '-- Default speed-factor
    m_snFrameFactor = 1
    
    m_ltFPS = 0
    m_lnFPS = 0
    m_lcFPS = 0
    
    '-- Game loop
    Do
        '-- FPS
        If (timeGetTime() - m_ltFPS >= 1000) Then
            m_ltFPS = timeGetTime()
            m_lnFPS = m_lcFPS
            m_lcFPS = 0
        End If
        
        '-- Frame timing
        If (timeGetTime() - t >= m_lFrameDt * m_snFrameFactor) Then
            t = timeGetTime()
            m_lcFPS = m_lcFPS + 1
            Call DoFrame
        End If
        
        '-- Keep a low CPU usage;
        '   filter allows to test maximum speed
        If (m_snFrameFactor <> 0) Then
            Call Sleep(1)
        End If
        
        '   Allow 'multitasking'
        Call VBA.DoEvents
        
    Loop Until m_bExit
End Sub

Public Sub StopGame()

    '-- Make sure to restore screen mode!
    If (mFullScreen.IsFullScreen) Then
        Call mFullScreen.ToggleFullScreen
    End If

    '-- Stop loop
    m_bExit = True
End Sub

'========================================================================================
' Setting mode and game mode
'========================================================================================

Private Sub SetMode( _
            ByVal mode As eMode _
            )
    
    '-- Mode
    m_eMode = mode
    
    '-- Frame dt
    Select Case mode
        Case [eLoading]
            m_lFrameDt = 80
        Case [eTitle]
            m_lFrameDt = 80
        Case [eGamePlay]
            m_lFrameDt = 70
        Case [eGameDie]
            m_lFrameDt = 15
        Case [eGameOver]
            m_lFrameDt = 80
    End Select
    
    '-- Reset all counters
    m_lc0 = 0
    m_lc1 = 0
    m_lc2 = 0
    
    '-- Stop all channels
    Call FXResetAll
End Sub

'========================================================================================
' Loop main routines
'========================================================================================

Private Sub DoFrame()
    
    Select Case m_eMode
    
        Case [eLoading]
            
            Call DoIntro
            
            If (KeysCheckAnyKey()) Then
                Call SetMode([eTitle])
                m_bKey(vbKeyReturn) = False
            End If
            
        Case [eTitle]
            
            Call DoTitle
            
            If (m_bKey(vbKeyReturn)) Then
                 
                Call SetMode([eGamePlay])
                Call InitializePlay
            End If
        
        Case [eGamePlay]
        
            If (m_bPause = False) Then
                Call DoGamePlay
            End If
                
            If (m_bKey(vbKeyEscape)) Then
                Call SetMode([eTitle])
            End If
            
        Case [eGameDie]
            
            Call DoGameDie
        
        Case [eGameOver]
            
            Call DoGameOver
    End Select
End Sub

Private Sub DoIntro()
    
    '-- First frame
    If (m_lc0 = 0) Then
        
        '-- Blue screen
        Call FXBorder(m_oDIBFore2x, IDX_BLUE, , m_bFXTV)
        Call m_oDIBFore.Cls(IDX_BLUE)
        
        '-- Yellow rectangle
        Call FXRect(m_oDIBFore, 40, 80, 176, 24, IDX_YELLOW)
    End If
    
    '-- Every 5 frames
    If (m_lc0 Mod 5 = 0) Then
        If (m_lc0 Mod 2 = 0) Then
            Call FXText(m_oDIBFore, 48, 88, "JetSet Willy loading", m_oDIBChar(), IDX_WHITE, IDX_RED) ' Red on white
          Else
            Call FXText(m_oDIBFore, 48, 88, "JetSet Willy loading", m_oDIBChar(), IDX_RED, IDX_WHITE) ' White on red
        End If
    End If
    
    '-- Update
    Call ScreenUpdate
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoTitle()
  
    Select Case m_lc0
        
        Case 0
            
            '-- Reset border / render screen
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, , m_bFXTV)
            Call RenderTitleScreen
            m_nNoteLen = 3
            
        Case 1 To 300
            
            If (m_lc1 = 0) Then
                '-- Set length and get JSW note
                m_nNote1 = GetJSWNote(m_aTitleTune(m_lc2)) + 48
                '-- Second note
                m_nNote2 = m_nNote1 + 0
                '-- Play 1st
                m_hChannelN1 = FXPlay(m_hFXTone, , 25 * GetNoteFreq(m_nNote1), True)
            End If
            
            If (m_lc1 = 1) Then
                '-- Play 2nd (loop)
                If (m_hChannelN2 = 0) Then
                    m_hChannelN2 = FXPlay(m_hFXTone, , , True)
                End If
                Call FXChangeFreq(m_hChannelN2, 25 * GetNoteFreq(m_nNote2))
            End If
             
            '-- Update screen
            Call ScreenUpdate
            Call ScreenFlash(DIB:=m_oDIBFore)
            
            '-- Counter (# frames = length)
            m_lc1 = m_lc1 + 1
            If (m_lc1 = m_nNoteLen) Then
                m_lc1 = 0
                m_lc2 = m_lc2 + 1
                '-- Stop 1st
                Call FXStop(m_hChannelN1)
           End If
            
        Case 301
        
            '-- Stop 2nd
            Call FXStop(m_hChannelN2)
            
            '-- Reset char counter
            m_lc2 = 1
            m_lFrameDt = 48
            
        Case 302 To 525
            
            '-- Rotate border color
            Call FXBorder(m_oDIBFore2x, 7 - m_lc0 Mod 7, , m_bFXTV)
            
            '-- Rotate screen color
            Call FXShift(m_oDIBFore, Inc:=3)
            
            '-- Scroll message...
            For m_lc1 = m_lc2 To m_lc2 + 31
                Call BltMask(m_oDIBFore, (m_lc1 - m_lc2) * 8, 152, 8, 8, IDX_BRBLUE, IDX_BRWHITE, m_oDIBChar(m_aScrolly(m_lc1) - 32))
            Next m_lc1
            m_lc2 = m_lc2 + 1
            
            '-- Update
            Call ScreenUpdate
        
            '-- Sound FX
            m_hChannelN1 = FXPlay(m_hFXFX, , 20000 + 500 * (80 - m_lFrameDt))
            
            '-- Speed up
            m_lFrameDt = m_lFrameDt + 1
            If (m_lFrameDt > 80) Then
                m_lFrameDt = 49
            End If
            
        Case 526
            
            '-- Reset all counters
            m_lc0 = 0
            m_lc1 = 0
            m_lc2 = 0
            Exit Sub
    End Select
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoGamePlay()
    
    '-- Play in-game tune and check if
    '   teleporter mask-code has been typed
    Call PlayInGameTune
    Call KeysCheckTeleporterMask
    
    '-- Flip buffer
    Call ScreenFlipBuffer
    
    '-- Do all animations
    Call DoPanel
    Call DoGuardians
    Call DoArrows
    Call DoRopes
    Call DoItems
    Call DoSpecial
    
    '-- Do Willy
    Call DoWilly
    Call DoWillyRope
    
    '-- Refresh screen
    Call ScreenUpdate
    
    '-- Post-processing
    Call ScreenAnimateConveyor
    Call ScreenFlash(DIB:=m_oDIBBack)
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoGameDie()
    
    Select Case m_lc2
    
        Case 0
        
            '-- Activate flag (skip first Willy frame)
            With m_tWillySafe
                .JustDied = True
            End With
            
            '-- Black border
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, , m_bFXTV)
        
            '-- Sound FX
            Call FXPlay(m_hFXDead)
        
        Case 1 To 8
        
            '-- FX (from white to black)
            Call BltMask(m_oDIBFore, 0, 0, 256, 128, IDX_BLACK, 16 - m_lc2, m_oDIBMask)
            Call ScreenUpdate
       
        Case 9
            
            '-- Little pause
            m_lFrameDt = 100
       
        Case 10
            
            '-- Check if last live
            With m_tPanel
            
                '-- Sorry...
                .Lives = .Lives - 1
    
                '-- Game over or try again
                If (.Lives = 0) Then
                    Call SetMode([eGameOver])
                  Else
                    Call SetMode([eGamePlay])
                    Call InitializeRoom(m_aRoomID)
                End If
            End With
    End Select
    
    '-- Counter (m_lC0 & m_lC1 used by in-game tune)
    m_lc2 = m_lc2 + 1
End Sub

Private Sub DoGameOver()
    
    Select Case m_lc0
    
        Case 0
            
            '-- Prepare background
            Call FXRect(m_oDIBBack, 0, 0, 256, 128, IDX_BLACK)
            Call BltFast(m_oDIBBack, 120, 0, 16, 16, m_oDIBFoot)
            Call BltFast(m_oDIBBack, 124, 96, 16, 16, m_tWilly.DIB(0))
            
        Case 1 To 48
            
            '-- Foot
            Call BltFast(m_oDIBBack, 120, m_lc0 * 2, 16, 16, m_oDIBFoot)
            
            '-- Masked screen
            Call BltMask(m_oDIBFore, 0, 0, 256, 128, m_lc0 Mod 4 + 8, IDX_BRWHITE, m_oDIBBack)
            
            '-- Red barrel
            Call MaskBltMask(m_oDIBFore, 120, 112, 16, 16, IDX_BRRED, m_oDIBBarrel)
            
            '-- Update
            Call ScreenUpdate
            
            '-- Sound FX
            Call FXPlay(m_hFXJump, , 5000 + 500 * m_lc0)
            
            '-- Speed up boot
            If (m_lc0 Mod 4 = 0) Then
                m_lFrameDt = m_lFrameDt - 5
            End If
                                
        Case 49 To 126
        
            '-- Red barrel
            Call MaskBltMask(m_oDIBFore, 120, 112, 16, 16, IDX_BRRED, m_oDIBBarrel)
            
            '-- "Game"
            Call BltMask(m_oDIBFore, 80, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 0, m_oDIBChar(39))
            Call BltMask(m_oDIBFore, 88, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 1, m_oDIBChar(65))
            Call BltMask(m_oDIBFore, 96, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 2, m_oDIBChar(77))
            Call BltMask(m_oDIBFore, 104, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 3, m_oDIBChar(69))
            
            '-- "Over"
            Call BltMask(m_oDIBFore, 144, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 4, m_oDIBChar(47))
            Call BltMask(m_oDIBFore, 152, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 5, m_oDIBChar(86))
            Call BltMask(m_oDIBFore, 160, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 6, m_oDIBChar(69))
            Call BltMask(m_oDIBFore, 168, 48, 8, 8, IDX_BLACK, m_lc0 Mod 8 + 7, m_oDIBChar(82))
            Call ScreenUpdate
            
        Case 127
        
            '-- Title mode
            Call SetMode([eTitle])
            Exit Sub
    End Select
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

'========================================================================================
' Game routines
'========================================================================================

'----------------------------------------------------------------------------------------
' Initializations
'----------------------------------------------------------------------------------------

Private Sub InitializePlay()

    '-- 3 lives / reset score and extra-live counter
    With m_tPanel
        .Lives = 8
        .t = 100800
        .AMPM = "am"
        .c = 0
    End With
    
    '-- Reset Willy
    With m_tWilly
        .x = 20  ' hardcoded
        .y = 104 ' hardcoded
        .c = 0
        .mode = [eWalk]
        .Flag = 0
        .f = 0
        .Dir = 0
        .Frame = 0
        .OnSlope = False
        .OnRope = False
        .RopeID = 0
        .RopeAnchor = 0
        .JustDied = False
    End With
    m_bFlee = False
    m_bVomit = False
    
    '-- Store safe copy
    LSet m_tWillySafe = m_tWilly
    
    '-- Initialize all items
    Call InitializePlayItems
    
    '-- 'The Bathroom'
    Call InitializeRoom(33) ' hardcoded
End Sub

Private Sub InitializeRoom( _
            ByVal RoomID As Byte, _
            Optional ByVal InitWilly As Boolean = True _
            )
    
    '-- Store room number
    m_aRoomID = RoomID
    
    '-- Initialize all
    Call InitializeLayout
    Call InitializeItems
    Call InitializeNasties
    Call InitializeGuardians
    
    '-- Initialize Willy
    If (InitWilly) Then
        Call InitializeWilly
    End If
    
    '-- Special (auto-collectable items)
    Call CheckAutoCollectableItems
End Sub

Private Sub InitializePlayItems()
    
  Dim a As Byte
  Dim b As Byte
  Dim c As Integer
    
    ReDim m_tItem(256 - m_aData(OFFSET_OBJECTTABLE))
    
    For c = 1 To UBound(m_tItem())
        
        With m_tItem(c)
            
            '-- Packed room and position
            a = m_aData(OFFSET_OBJECTTABLE + c + m_aData(OFFSET_OBJECTTABLE))
            b = m_aData(OFFSET_OBJECTTABLE + c + 256 + m_aData(OFFSET_OBJECTTABLE))
            
            '-- Unpack
            .Room = a And &H3F
            .x = (b And &H1F)
            .y = (b And &HE0) \ 32 + (a And &H80) \ 16
            
            '-- Not-collected flag
            .Flag = 0
        End With
    Next c
    
    '-- Initialize both counters: left and collected
    m_aItemsLeft = 256 - m_aData(OFFSET_OBJECTTABLE)
    m_aItems = 0
End Sub

Private Sub InitializeLayout()

  Dim DIB(5) As New cDIB08
  
  Dim ro     As Integer
  Dim o      As Integer
  Dim c      As Integer
  Dim z      As Integer
  
  Dim a      As Byte
  Dim b      As Byte
  Dim d      As Byte
  
  Dim s      As Byte
  Dim m      As Byte
  Dim n      As Long
    
    '-- Current room data offset
    
    ro = OFFSET_ROOMDATA + 256 * m_aRoomID
      
    '-- Unpack blocks
    '   4 bytes: 4 directions
    '   0=left, 1=right, 2=up, 3=down
    
    m_aRoomExit(0) = m_aData(ro + OFFSET_ROOMEXITS + 0) ' left
    m_aRoomExit(1) = m_aData(ro + OFFSET_ROOMEXITS + 1) ' right
    m_aRoomExit(2) = m_aData(ro + OFFSET_ROOMEXITS + 2) ' up
    m_aRoomExit(3) = m_aData(ro + OFFSET_ROOMEXITS + 3) ' down
    
    '-- Unpack blocks
    '   1-byte color def. + 8-byte pattern
    
    For c = 0 To 5
        
        '-- 1st byte offset
        o = ro + OFFSET_BLOCKSDEF + 9 * c
        
        '-- CA
        m_aBlockCA(c) = m_aData(o)
        
        '-- Unpack sprite
        Call Unpack08(m_aData(), o + 1, DIB(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Store CAs *constants*
    
    BLOCK_AIR = m_aBlockCA(0)
    BLOCK_FLOOR = m_aBlockCA(1)
    BLOCK_WALL = m_aBlockCA(2)
    BLOCK_NASTY = m_aBlockCA(3)
    BLOCK_SLOPE = m_aBlockCA(4)
    BLOCK_CONVEYOR = m_aBlockCA(5)
    
    '-- Unpack screen definition
    '   128 4-block bytes
    
    For c = 0 To &H7F
        
        '-- Offset
        o = ro + OFFSET_SCREENDEF + c
        
        '-- 4-block byte
        a = m_aData(ro + OFFSET_SCREENDEF + c)
        
        '-- Unpack
        For d = 0 To 3
            s = m_aPow(2 * d)   ' shift
            m = 3 * s           ' mask
            b = (a And m) \ s   ' block ID (0-3)
            n = 4 * c + (3 - d) ' index (0-511)
            
            '-- Render
            Call BltMask(m_oDIBBack, 8 * (n Mod 32), 8 * (n \ 32), 8, 8, CAPaper(m_aBlockCA(b)), CAInk(m_aBlockCA(b)), DIB(b))
            Call BltFast(m_oDIBMask, 8 * (n Mod 32), 8 * (n \ 32), 8, 8, DIB(b))
            
            '-- Store ID's CA
            m_aRoomCA(n) = (m_aBlockCA(b))
        Next d
    Next c
    
    '-- Conveyor definition
    '   byte    1 - direction: 0=left, 1=right [2=off, 3=sticky]
    '   bytes 2+3 - packed start position
    '   byte    4 - length
    
    '-- 1st byte offset
    
    o = ro + OFFSET_CONVEYORDEF
     
    '-- Store conveyor def.
    
    With m_tConveyor
    
        '-- Unpack position
        n = m_aData(o + 1) + 256& * m_aData(o + 2)
        .x = n And &H1F
        .y = (n \ 32) And &HF
        
        '-- Direction and blocks-length
        .Dir = m_aData(o + 0)
        .Len = m_aData(o + 3)
        
        '-- Render conveyor blocks and store them
        For c = 0 To .Len - 1
            Call BltMask(m_oDIBBack, 8 * (.x + c), 8 * .y, 8, 8, CAPaper(BLOCK_CONVEYOR), CAInk(BLOCK_CONVEYOR), DIB(5))
            Call BltFast(m_oDIBMask, 8 * (.x + c), 8 * .y, 8, 8, DIB(5))
            m_aRoomCA(32 * .y + .x + c) = BLOCK_CONVEYOR
        Next
    End With
    
    '-- Slope definition
    '   byte    1 - direction: 0=left, 1=right
    '   bytes 2+3 - packed start position
    '   byte    4 - length
    
    '-- 1st byte offset
    
    o = ro + OFFSET_SLOPEDEF
       
    '-- Store slope def.
    
    With m_tSlope
    
        '-- Unpack position
        n = m_aData(o + 1) + 256& * m_aData(o + 2)
        .x = n And &H1F
        .y = ((n \ 2) And &HF0&) \ 16
        
        '-- Direction and blocks-length
        .Dir = m_aData(o + 0)
        .Len = m_aData(o + 3)
        
        '-- Render conveyor blocks and store them
        z = IIf(.Dir, 1, -1)
        For c = 0 To .Len - 1
            Call BltMask(m_oDIBBack, 8 * (.x + z * c), 8 * (.y - c), 8, 8, CAPaper(BLOCK_SLOPE), CAInk(BLOCK_SLOPE), DIB(4))
            Call BltFast(m_oDIBMask, 8 * (.x + z * c), 8 * (.y - c), 8, 8, DIB(4))
            m_aRoomCA(32 * (.y - c) + .x + z * c) = BLOCK_SLOPE
        Next
    End With
    
    '-- Store 'flash' attribute independently (copy)
    
    For c = 0 To 511
        m_aRoomFA(c) = m_aRoomCA(c)
    Next c
    
    '-- Render room name
    '   32-length string
    
    For c = 0 To 31
        '-- Char offset
        o = ro + OFFSET_ROOMNAME + c
        '-- Char ascii
        a = m_aData(o) And &H7F
        '-- Render
        Call BltMask(m_oDIBBack, c * 8, 128, 8, 8, IDX_BLACK, IDX_BRYELLOW, m_oDIBChar(a - 32))
    Next c
    
    '-- Border color
    
    Call FXBorder(m_oDIBFore2x, m_aData(ro + OFFSET_BORDERCOLOR), , m_bFXTV)
End Sub

Private Sub InitializeWilly()
    
    LSet m_tWilly = m_tWillySafe
    
    With m_tWilly
        .OnConveyor = False
        .OnRope = False
        .RopeID = 0
        .RopeAnchor = 0
        .RopeExit = (m_aRoomExit(2) <> m_aRoomID)
    End With
End Sub

Private Sub InitializeItems()

  Dim ro As Integer
  Dim c  As Integer
  Dim i  As Byte
 
    '-- Room data offset
    ro = OFFSET_ROOMDATA + 256 * m_aRoomID

    '-- Unpack item (object) sprite
    Call Unpack08(m_aData(), ro + OFFSET_ITEMPATTERN, m_oDIBItem, IDX_MASK, IDX_NULL)
    
    '-- Initialize inks
    i = IDX_MAGENTA
    For c = 1 To UBound(m_tItem())
        With m_tItem(c)
            If (.Room) = m_aRoomID Then
                .Ink = i
                i = i + 1
                If (i > 6) Then
                    i = 3
                End If
            End If
        End With
    Next c
End Sub

Private Sub InitializeGuardians()
    
    Dim ro  As Integer
 
    Dim c   As Integer
    Dim o   As Integer
    Dim f   As Integer
    
    Dim rd1 As Byte ' room data bytes 1 and 2
    Dim rd2 As Byte
    
    Dim gt1 As Byte ' guardian table bytes 1 to 8
    Dim gt2 As Byte
    Dim gt3 As Byte
    Dim gt4 As Byte
    Dim gt5 As Byte
    Dim gt6 As Byte
    Dim gt7 As Byte
    Dim gt8 As Byte
    
    Dim ng  As Byte ' counters
    Dim na  As Byte
    Dim nr  As Byte
     
    '-- Room data offset
    ro = OFFSET_ROOMDATA + 256 * m_aRoomID
    
    '-- Reset all
    ReDim m_tGuardian(0)
    ReDim m_tRope(0)
    ReDim m_tArrow(0)
    
    '-- 8 guardian class instances (2 bytes each)
    For c = 240 To 255 Step 2
        
        '-- Get bytes
        rd1 = m_aData(ro + c + 0)
        rd2 = m_aData(ro + c + 1)
        
        '-- &HFF = end of sequence
        If (rd1 = &HFF) Then
            
            '-- Skip
            Exit For
            
          Else
            
            '-- Guardian instance offset
            o = 8 * rd1
            
            '-- Get guardian bytes
            gt1 = m_aData(OFFSET_GUARDIANTABLE + o + 0)
            gt2 = m_aData(OFFSET_GUARDIANTABLE + o + 1)
            gt3 = m_aData(OFFSET_GUARDIANTABLE + o + 2)
            gt4 = m_aData(OFFSET_GUARDIANTABLE + o + 3)
            gt5 = m_aData(OFFSET_GUARDIANTABLE + o + 4)
            gt6 = m_aData(OFFSET_GUARDIANTABLE + o + 5)
            gt7 = m_aData(OFFSET_GUARDIANTABLE + o + 6)
            gt8 = m_aData(OFFSET_GUARDIANTABLE + o + 7)
            
            Select Case gt1 And &HF
                
                '-- Horizontal ----------------------------------------------------------
                
                Case 1
                    
                    ng = ng + 1
                    ReDim Preserve m_tGuardian(ng)
                    
                    With m_tGuardian(ng)
                        
                        '-- Store type
                        .Type = 1
                        
                        '-- Ink
                        .Ink = (gt2 And &H7) + (gt2 And &H8)
                        
                        '-- Starting position and direction
                        .x = rd2 And &H1F
                        .y = gt4 \ 2
                        .Dir = (gt1 And &H80) \ 128
                        
                        '-- Range
                        .Min = gt7
                        .Max = gt8
                        
                        '-- Frame range
                        .FrameI = (rd2 And &HE0) \ 32
                        .FrameF = (gt2 And &HE0) \ 32 + .FrameI
                        
                        '-- Fix final frame
                        If (.FrameF > 7) Then
                            .FrameF = 7
                        End If
                        
                        '-- Starting frame
                        If (.FrameF - .FrameI = 7) Then
                            .Frame = (gt1 And &H60) \ 32 + .FrameI + 4 * .Dir
                          Else
                            .Frame = (gt1 And &H60) \ 32 + .FrameI
                        End If
                        
                        '-- Sprites page offset
                        o = (gt6 - &H9B) * 256 + (rd2 And &HE0)
                        
                        '-- Unpack sprites
                        For f = 0 To .FrameF - .FrameI
                            Call Unpack16(m_aData(), o + 32 * f, .DIB(f), IDX_MASK, IDX_NULL)
                        Next f
                    End With
                
                '-- Vertical ------------------------------------------------------------
                
                Case 2
                    
                    ng = ng + 1
                    ReDim Preserve m_tGuardian(ng)
                    
                    With m_tGuardian(ng)
                    
                        '-- Store type
                        .Type = 2
                        
                        '-- Ink
                        .Ink = (gt2 And &H7) + (gt2 And &H8)
                        
                        '-- Starting position
                        .x = rd2 And &H1F
                        .y = gt4 \ 2
                        
                        '-- Range
                        .Min = gt7 \ 2
                        .Max = gt8 \ 2
                        
                        '-- Speed (and implicitly: starting direction)
                        .Speed = gt5 \ 2
                        If (.Speed > 64) Then
                            .Speed = .Speed - 128
                        End If
                        
                        '-- Fast animation
                        .Fast = 1 - (gt1 And &H10) \ 16 ' 1/0: fast/normal
                        .c = .Fast
                        
                        '-- Starting frame
                        .FrameI = (rd2 And &HE0) \ 32
                        .Frame = 0
                        
                        '-- Ending frame
                        .FrameF = .FrameI Or (gt2 And &HE0) \ 32
                         
                        '-- Fixes
                        If (.y < .Min) Then
                            .y = .Min - .Speed
                        End If
                        If (.y > .Max) Then
                            .y = .Max + .Speed
                        End If
                        
                        '-- Sprites page offset
                        o = (gt6 - &H9B) * 256
                        
                        '-- Unpack sprites
                        Select Case .FrameF - .FrameI
                        
                            Case 3
                                
                                '-- a-b-c-d sequence
                                Call Unpack16(m_aData(), o + 32 * (.FrameI + 0), .DIB(0), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * (.FrameI + 1), .DIB(1), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * (.FrameI + 2), .DIB(2), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * (.FrameI + 3), .DIB(3), IDX_MASK, IDX_NULL)
                            
                            Case 2
                            
                                '-- a-a-b-b sequence
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(0), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(1), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameF, .DIB(2), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameF, .DIB(3), IDX_MASK, IDX_NULL)
                            
                            Case 1
                                
                                '-- a-b-a-b sequence
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(0), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameF, .DIB(1), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(2), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameF, .DIB(3), IDX_MASK, IDX_NULL)
                            
                            Case 0
                                
                                '-- a-a-a-a sequence
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(0), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(1), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(2), IDX_MASK, IDX_NULL)
                                Call Unpack16(m_aData(), o + 32 * .FrameI, .DIB(3), IDX_MASK, IDX_NULL)
                        End Select
                        
                    End With
                
                '-- Rope ----------------------------------------------------------------
                
                Case 3
                    
                    nr = nr + 1
                    ReDim Preserve m_tRope(nr)
                    
                    With m_tRope(nr)
                    
                        '-- Rope color (air ink)
                        .Ink = CAInk(BLOCK_AIR)
                        
                        '-- Rope anchor
                        .x = 8 * rd2
                        .y = 0
                        
                        '-- Starting direction (fix)
                        .Dir = (gt1 And &H80) \ 128
                        .Dir = -2 + 4 * -(.Dir = 0)
                        .c = 22 '!?
                       
                        '-- Rope lenght (dots) and swing
                        .Len = gt5
                        .Swing = gt8
                        
                        '-- Initialize rope dots array
                        ReDim .Dot(.Len - 1)
                    End With
                    
                '-- Arrow ---------------------------------------------------------------
                
                Case 4
                    
                    na = na + 1
                    ReDim Preserve m_tArrow(na)
                    
                    With m_tArrow(na)
                        
                        '-- Direction
                        .Dir = (gt1 And &H80) \ 128
                        
                        '-- Start at
                        .x = gt5
                        .y = rd2 \ 2
                        
                        '-- Arrow *sprite*
                        Call UnpackArrow(gt7, .DIB, IDX_MASK, IDX_NULL)
                    End With
            End Select
        End If
    Next c
End Sub

Private Sub InitializeNasties()
  
  Dim c As Integer
  Dim n As Byte
    
    ReDim m_tNasty(0)
    n = 0
    
    '-- Get positions
    For c = 0 To 511
        If (m_aRoomCA(c) = BLOCK_NASTY) Then
            n = n + 1
            ReDim Preserve m_tNasty(n)
            With m_tNasty(n)
                .x = c Mod 32
                .y = c \ 32
            End With
        End If
    Next c
End Sub

Private Sub CheckAutoCollectableItems()
    
  Dim c As Integer
  
    '-- White background (air) ink? (Willy ink)
    If (CAInk(BLOCK_AIR) = IDX_WHITE) Then
        
        For c = 1 To UBound(m_tItem())
    
            With m_tItem(c)
                
                '-- Correct room?
                If (.Room = m_aRoomID) Then
                
                    '-- Not collected?
                    If (.Flag = 0) Then
                            
                        '-- Auto-collect!
                        .Flag = 1
                            
                        '-- Update
                        m_aItems = m_aItems + 1
                        m_aItemsLeft = m_aItemsLeft - 1
                    End If
                End If
            End With
        Next c
    End If
End Sub

'----------------------------------------------------------------------------------------
' Willy rountines
'----------------------------------------------------------------------------------------

Private Sub DoWilly()
  
  Dim Keys As Byte
    
    With m_tWilly
        
        '-- Skip first *frame* if just died
        If (.JustDied) Then
            .JustDied = False
            GoTo check
        End If
                 
        '-- Check keys (right/left/jump)
        If Not (m_bFlee) Then
            Call KeysCheckWillyKeys(Keys)
          Else
            Keys = 1 ' force right
        End If
        
        '-- Previous checks
        If Not (.OnRope) Then
        
            If (.mode = [eWalk]) Then
                
                '-- At last row?
                If Not (WillyCheckExit([eDown])) Then
                    
                    '-- Fall?
                    If (.y Mod 8 = 0) Then
                        If Not (WillyCheckFeet()) Then
                            .mode = [eFall]
                            .c = 18
                            .f = .y
                            GoTo skip
                        End If
                    End If
                    
                    '-- Conveyor?
                    Call WillyCheckConveyor(Keys)
                    
                    '-- At last!
                    If (m_aItemsLeft = 0) Then

                        Select Case True
                            
                            '-- 'Master Beadroom': by the bed
                            Case (m_aRoomID = 35) And (.x = 6) And Not (m_bFlee)
                                
                                m_bFlee = True
                                
                            '-- 'The Bathroom': by the toilet
                            Case (m_aRoomID = 33) And (.x = 28) And (m_bFlee) And Not (m_bVomit)
                                
                                m_bVomit = True
                                Call DoSpecial
                        End Select
                    End If
                  
                  Else
                    GoTo skip
                End If
            End If
        End If
        
        '-- Process depending on mode...
        Select Case .mode

            Case [eWalk] ' walking/standing
                
                '-- Reset counter
                .c = 0
                
                '-- Pre-set jump mode
                If (Keys And 4) Then
                    .mode = [eJump]
                End If
         
                '-- Right and/or left keys pressed
                Select Case Keys And 3

                    Case 0                          ' nor right nor left
                        
                        .Flag = 0                   ' nullify direction flag
                    
                    Case 1                          ' right

                        If (.Dir = 1) Then          ' facing left
                            .Dir = 0                ' change dir
                            .Flag = 0               ' nullify direction flag
                            .Frame = .Frame - 4     ' toggle frame counter
                          Else                      ' already going right
                            .Flag = 1               ' change flag
                            Call WillyRight         ' go right
                        End If

                    Case 2                          ' facing right

                        If (.Dir = 0) Then          ' turned to left
                            .Dir = 1                ' change dir
                            .Flag = 0               ' nullify direction flag
                            .Frame = .Frame + 4     ' toggle frame counter
                          Else                      ' already going left
                            .Flag = 2               ' change flag
                            Call WillyLeft          ' go left
                        End If
                End Select
             
                '-- Jump mode?
                If (.mode = [eJump]) Then
                    .RopeID = 0                     ' reset (.onrope is processed)
                    .y = .y And &HF8                ' fix y
                    .f = .y
                    If (WillyCheckHead()) Then      ' something over head
                        .mode = [eFall]             ' sorry: fall
                        .c = 18
                      Else
                        If (.OnRope) Then           ' jump has to face right/left
                            .Flag = .Dir + 1
                        End If
                    End If
                End If
                
            Case [eJump] ' jumping
           
                Call WillyJump                      ' jump!
                
                Select Case .mode                   ' something has changed?
                    
                    Case [eWalk]                    ' now walking/standing
                        
                        If Not (.OnConveyor) Then   ' reset direction flag?
                            .Flag = 0
                        End If
                        If (.c < 18) Then           ' during jump stage?
                            Call DoWilly            ' repeat
                            Exit Sub
                        End If
                    
                    Case [eJump]                    ' still jumping
                        
                        Select Case .Flag
                            Case 1                  ' jumping right
                                Call WillyRight
                            Case 2                  ' jumping left
                                Call WillyLeft
                        End Select
                End Select
                
            Case [eFall] ' falling

                Call WillyFall                      ' fall!
                    
                If (.f = YFDEATH) Then              ' too much height?
                    Call SetMode([eGameDie])
                    Call DoFrame
                    Exit Sub
                End If
                    
                If (.mode = [eWalk]) Then           ' soft landing
                    If (.f <> .y) Then              ' not ouch!
                        Call DoWilly                ' repeat
                        Exit Sub
                    End If
                End If
        End Select
        
        '-- Check item/nasty/guardian
check:  If Not (m_bFlee Or m_bVomit) Then
                            
            Select Case True
    
                Case WillyCheckItem()
                    '-- Do nothing: continue
    
                Case WillyCheckNasty()
                    '-- Change mode
                    Call SetMode([eGameDie])
    
                Case WillyCheckGuardian()
                    '-- Change mode
                    Call SetMode([eGameDie])
                    
                Case WillyCheckArrow()
                    '-- Change mode
                    Call SetMode([eGameDie])
            End Select
        End If
        
skip:   '-- Render Willy
        If Not (m_bVomit) Then
            If (m_aRoomID = 29) Then
                Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, IDX_WHITE, m_oDIBPig(.Frame))
              Else
                Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, IDX_WHITE, .DIB(.Frame))
            End If
        End If
    End With
End Sub

Private Sub DoWillyRope()

  Dim c   As Integer
  Dim d   As Integer
  Dim x   As Byte
  Dim y   As Byte
    
    With m_tWilly
    
        If (Not .OnRope) Then
        
            For c = 1 To UBound(m_tRope())
        
                For d = 0 To m_tRope(c).Len - 1
                    
                    '-- Collide?
                    If (.DIB(.Frame).GetPixel(m_tRope(c).Dot(d).x - 8 * .x, _
                                              m_tRope(c).Dot(d).y - .y) _
                                              ) Then
                                              
                        .OnRope = True  ' on rope, now
                        .RopeID = c     ' rope #
                        .RopeAnchor = d ' anchor dot
                        .mode = [eWalk] ' turn to 'walk' mode
                        
                        GoTo check
                    End If
                Next d
            Next c
        End If
    End With
       
check:
    If (UBound(m_tRope())) Then
        
        With m_tWilly
            
            If (.RopeID) Then
                
                '-- Center Willy on dot
                x = m_tRope(.RopeID).Dot(.RopeAnchor).x - 4
                y = m_tRope(.RopeID).Dot(.RopeAnchor).y - 8
                
                '-- Correct units
                .x = x \ 8
                .y = y
                
                '-- Adjust frame
                .Frame = 4 * .Dir + (x Mod 8) \ 2
            End If
        End With
    End If
End Sub

Private Function WillyCheckExit( _
                 ByVal ExitID As eExit _
                 ) As Boolean
    
    With m_tWilly

        Select Case ExitID
        
            Case [eLeft]
            
                If (.x = XMIN) Then
                    .x = XMAX - 1
                    
                    '-- Fix y
                    .y = IIf(.mode = [eWalk], .y And &HF8, .y)
                    
                    '-- Left exit
                    WillyCheckExit = True
                End If
            
            Case [eRight]
            
                If (.x = XMAX - 1) Then
                    .x = XMIN
                    
                    '-- Fix y
                    .y = IIf(.mode = [eWalk], .y And &HF8, .y)
                    
                    '-- Right exit
                    WillyCheckExit = True
                End If
                
            Case [eUp]
            
                If (.y = 8 * YMIN Or .RopeID > 0) Then
                    .y = 8 * YMAX - 8
                    
                    '-- Reset to 'walk' mode
                    .mode = [eWalk]
                    .Flag = 0
                    .c = 0
                    
                    '-- Up exit
                    WillyCheckExit = True
                End If
                
            Case [eDown]
            
                If (.y = 8 * YMAX) Then
                    .y = 8 * YMIN
                    
                    '-- Fix starting fall position
                    .f = IIf(.mode = [eFall], .f - (128 - 32), .y)
                    .Flag = 0
                    
                    '-- Down exit
                    WillyCheckExit = True
                End If
        End Select
        
        '-- Room left
        If (WillyCheckExit) Then
            
            '-- Initialize new room
            Call InitializeRoom(m_aRoomExit(ExitID), InitWilly:=False)
            Call ScreenFlipBuffer(SkipPanel:=True)
            
            '-- Force a *first frame*
            Call DoGuardians
            Call DoArrows
            Call DoRopes
            Call DoItems
            Call DoSpecial
            
            '-- Willy adjustments
            Select Case .mode
                Case [eWalk]
                    Call WillyCheckFeet
                Case [eFall]
                    Call WillyFall
            End Select
            
            '-- Store safe copy
            LSet m_tWillySafe = m_tWilly
            
            '-- Initialize Willy
            Call InitializeWilly
        End If
    End With
End Function

Private Sub WillyLeft()
    
    With m_tWilly
    
        If (.Frame > 4) Then
            .Frame = .Frame - 1                    ' previous frame
            If (.mode = [eWalk]) Then
                Call WillyCheckLeftSlope
            End If
          Else
            If Not (WillyCheckExit([eLeft])) Then
                If Not (WillyCheckLeft()) Then
                    .Frame = 7                     ' last frame
                    .x = .x - 1                    ' one block left
                    If (.mode = [eWalk]) Then
                        Call WillyCheckLeftSlope
                    End If
                End If
            End If
        End If
          
        If (.OnRope) Then
            If (m_tRope(.RopeID).c < 0) Then
                If (.RopeAnchor > IIf(.RopeExit, 2, 11)) Then
                    .RopeAnchor = .RopeAnchor - 1
                  Else
                    If (.RopeExit) Then
                        Call WillyCheckExit([eUp])
                    End If
                End If
              Else
                If (.RopeAnchor < m_tRope(.RopeID).Len - 1) Then
                    .RopeAnchor = .RopeAnchor + 1
                  Else
                    .RopeID = 0
                    If (.mode = [eWalk]) Then
                        .mode = [eFall]
                        .c = 18
                        .y = .y And &HF8
                        .f = .y
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub WillyRight()

    With m_tWilly
    
        If (.Frame < 3) Then
            .Frame = .Frame + 1                    ' next frame
            If (.mode = [eWalk]) Then
                Call WillyCheckRightSlope
            End If
          Else
            If Not (WillyCheckExit([eRight])) Then
                If Not (WillyCheckRight()) Then
                    .Frame = 0                     ' first frame
                    .x = .x + 1                    ' one block right
                    If (.mode = [eWalk]) Then
                        Call WillyCheckRightSlope
                    End If
                End If
            End If
        End If

        If (.OnRope) Then
            If (m_tRope(.RopeID).c > 0) Then
                If (.RopeAnchor > IIf(.RopeExit, 2, 11)) Then
                    .RopeAnchor = .RopeAnchor - 1
                  Else
                    If (.RopeExit) Then
                        Call WillyCheckExit([eUp])
                    End If
                End If
              Else
                If (.RopeAnchor < m_tRope(.RopeID).Len - 1) Then
                    .RopeAnchor = .RopeAnchor + 1
                  Else
                    .RopeID = 0
                    If (.mode = [eWalk]) Then
                        .mode = [eFall]
                        .c = 18
                        .y = .y And &HF8
                        .f = .y
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub WillyJump()
    
    With m_tWilly
        If Not (WillyCheckExit([eUp])) Then
            If (WillyCheckHead() And .c < 9) Then ' going up
                .mode = [eWalk]
                .c = 18
              Else
                .y = .y + ((.c And &HFE) - 8) \ 2
                .c = .c + 1
                Select Case .c
                    Case Is = 9                   ' top
                        .f = .y
                    Case Is > 9                   ' going down
                        If (WillyCheckFeet()) Then
                            .y = .y And &HF8
                            .mode = [eWalk]
                            .OnRope = False
                          Else
                            If (.c = 18) Then
                                .mode = [eFall]
                                .OnRope = False
                            End If
                        End If
                End Select
            End If
        End If
        
        '-- Sound FX
        If (.mode <> [eWalk]) Then
            Call FXPlay(m_hFXJump, , 22050 + 1500 * (9 - Sqr((.c - 9) ^ 2)))
        End If
    End With
End Sub

Private Sub WillyFall()
    
    With m_tWilly
        If Not (WillyCheckExit([eDown])) Then
            If (WillyCheckFeet()) Then
                If (.y - .f > DYDEATH) Then   ' max. height reached
                    .f = YFDEATH              ' death flag
                  Else                        ' saved!
                    .mode = [eWalk]           ' OK
                    .Flag = 0
                    .OnRope = False           ' release now
                End If
              Else
                .y = .y + 4                   ' go down
                If (.c < 26) Then             ' counter (sound FX)
                    .c = .c + 1
                  Else
                    .c = 22
                End If
            End If
        End If
        
        '-- Sound FX
        If (.mode = [eFall] And .c > 18) Then
            Call FXPlay(m_hFXJump, , 22050 + 1500 * (9 - Sqr((.c - 9) ^ 2)))
        End If
    End With
End Sub

Private Function WillyCheckRight( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Boolean
  Dim b2 As Boolean
  Dim b3 As Boolean
    
    With m_tWilly
    
        If (.x < XMAX) Then
        
            o = .x + 2 + (.y \ 8 + 0) * 32
            
            b1 = (m_aRoomCA(o + 0 * 32) = BLOCK_WALL) And Not (.c = 12 Or .c = 15 Or .c = 17 Or .OnSlope)
            b2 = (m_aRoomCA(o + 1 * 32) = BLOCK_WALL)
            b3 = (m_aRoomCA(o + 2 * 32) = BLOCK_WALL) And (.c > 10 Or .y Mod 8 > 0)
            
            WillyCheckRight = (b1 Or b2 Or b3)
        End If
    End With
End Function

Private Function WillyCheckLeft( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Boolean
  Dim b2 As Boolean
  Dim b3 As Boolean
    
    With m_tWilly
        
        If (.x > XMIN) Then
        
            o = .x - 1 + (.y \ 8 + 0) * 32
            
            b1 = (m_aRoomCA(o + 0 * 32) = BLOCK_WALL) And Not (.c = 0 Or .c = 12 Or .c = 15 Or .c = 17 Or .OnSlope)
            b2 = (m_aRoomCA(o + 1 * 32) = BLOCK_WALL)
            b3 = (m_aRoomCA(o + 2 * 32) = BLOCK_WALL) And (.c > 10 Or .y Mod 8 > 0)
            
            WillyCheckLeft = (b1 Or b2 Or b3)
        End If
    End With
End Function

Private Sub WillyCheckLeftSlope()
    
  Dim b As Byte
  
    With m_tWilly
    
        If (m_tSlope.Dir = 1) Then
            If (.Frame = 6) Then
                b = m_aRoomCA(.x + (.y \ 8 + 2) * 32 + 1)
                .OnSlope = (b = BLOCK_SLOPE)
                .OnConveyor = (b = BLOCK_CONVEYOR)
            End If
          Else
            If (.Frame = 7) Then
                b = m_aRoomCA(.x + (.y \ 8 + 1) * 32 - 0)
                .OnSlope = (b = BLOCK_SLOPE)
                .OnConveyor = (b = BLOCK_CONVEYOR)
            End If
        End If
        If (.OnSlope) Then
            If (m_tSlope.Dir = 1) Then
                If Not (WillyCheckExit([eDown])) Then
                    .y = .y + 2
                End If
              Else
                If Not (WillyCheckExit([eUp])) Then
                    .y = .y - 2
                End If
            End If
        End If
    End With
End Sub

Private Sub WillyCheckRightSlope()
    
  Dim b As Byte
    
    With m_tWilly
        If (m_tSlope.Dir = 1) Then
            If (.Frame = 0) Then
                b = m_aRoomCA(.x + (.y \ 8 + 1) * 32 + 1)
                .OnSlope = (b = BLOCK_SLOPE)
                .OnConveyor = (b = BLOCK_CONVEYOR)
            End If
          Else
            If (.Frame = 1) Then
                b = m_aRoomCA(.x + (.y \ 8 + 2) * 32 - 0)
                .OnSlope = (b = BLOCK_SLOPE)
                .OnConveyor = (b = BLOCK_CONVEYOR)
            End If
        End If
        If (.OnSlope) Then
            If (m_tSlope.Dir = 1) Then
                If Not (WillyCheckExit([eUp])) Then
                    .y = .y - 2
                End If
              Else
                If Not (WillyCheckExit([eDown])) Then
                    .y = .y + 2
                End If
            End If
        End If
    End With
End Sub
  
Private Sub WillyAdjustSlope()

    With m_tWilly
        If (.Dir = 0) Then
            If (m_tSlope.Dir = 1) Then
                .y = (.y And &HF8) - 2 * (.Frame - 2) + 2
              Else
                .y = (.y And &HF8) + 2 * (.Frame + 1) - 2
            End If
          Else
            If (m_tSlope.Dir = 1) Then
                .y = (.y And &HF8) - 2 * (.Frame - 8) - 2
              Else
                .y = (.y And &HF8) + 2 * (.Frame - 5) + 2
            End If
        End If
    End With
End Sub

Private Function WillyCheckHead( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Boolean
  Dim b2 As Boolean
    
    With m_tWilly

        If (.y Mod 8 = 0 And .y \ 8 > YMIN) Then
            
            o = .x + (.y \ 8 - 1) * 32
            
            b1 = (m_aRoomCA(o + 0) = BLOCK_WALL)
            b2 = (m_aRoomCA(o + 1) = BLOCK_WALL)
            
            WillyCheckHead = (b1 Or b2)
        End If
    End With
End Function

Private Function WillyCheckFeet( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Byte
  Dim b2 As Byte
                                        
    With m_tWilly
        
        If (.y Mod 8 = 0 And .y \ 8 < YMAX) Then
            
            '-- Check feet for walkable block
            o = .x + (.y \ 8 + 2) * 32
            b1 = m_aRoomCA(o + 0)
            b2 = m_aRoomCA(o + 1)
           
            WillyCheckFeet = (b1 <> BLOCK_AIR And b1 <> BLOCK_NASTY) Or (b2 <> BLOCK_AIR And b2 <> BLOCK_NASTY)
            
            '-- Check for conveyor
            .OnConveyor = (b1 = BLOCK_CONVEYOR Or b2 = BLOCK_CONVEYOR)

            '-- Also, check for slope
            If (.Dir = 0) Then
                If (m_tSlope.Dir = 1) Then
                    .OnSlope = (b2 = BLOCK_SLOPE)
                  Else
                    .OnSlope = (b1 = BLOCK_SLOPE)
                End If
              Else
                If (m_tSlope.Dir = 1) Then
                    .OnSlope = (b2 = BLOCK_SLOPE)
                  Else
                    .OnSlope = (b1 = BLOCK_SLOPE)
                End If
            End If
            
            '-- ... If any, adjust y position (feet)
            If (.OnSlope) Then
                Call WillyAdjustSlope
            End If
        End If
    End With
End Function

Private Sub WillyCheckConveyor( _
            Keys As Byte _
            )
            
    With m_tWilly
        
        If (.OnConveyor) Then
    
             Select Case m_tConveyor.Dir
                 
                 Case 0 ' Right
                    
                    If (Keys And 1) Then                 ' <right> key pressed
                        Select Case .Flag                ' stored direction
                            Case 2                       ' left
                                Keys = (Keys And 6) Or 2
                            Case 0                       ' none
                                Keys = (Keys And 6)
                        End Select
                      Else
                        Keys = (Keys And 6) Or 2         ' turn to left
                    End If
                
                Case 1 ' Left
                    
                    If (Keys And 2) Then                 ' <left> key pressed
                        Select Case .Flag                ' stored direction
                            Case 1                       ' right
                                Keys = (Keys And 5) Or 1
                            Case 0                       ' none
                                Keys = (Keys And 5)
                        End Select
                      Else
                        Keys = (Keys And 5) Or 1         ' turn to right
                    End If
            End Select
        End If
    End With
End Sub

Private Function WillyCheckItem( _
                 ) As Boolean

  Dim c As Integer

    For c = 1 To UBound(m_tItem())

        With m_tItem(c)
            
            '-- Correct room?
            If (.Room = m_aRoomID) Then
            
                '-- Not collected?
                If (.Flag = 0) Then
    
                    '-- Collide?
                    If (.x = m_tWilly.x \ 1 Or .x = m_tWilly.x \ 1 + 1) And _
                       (.y = m_tWilly.y \ 8 Or .y = m_tWilly.y \ 8 + 1 Or ((.y = m_tWilly.y \ 8 + 2) And (m_tWilly.y Mod 8 > 0))) Then
    
                        '-- Collect it!
                        .Flag = 1
                        
                        '-- Update
                        m_aItems = m_aItems + 1
                        m_aItemsLeft = m_aItemsLeft - 1
    
                        '-- Sound FX
                        Call FXPlay(m_hFXPick)
                        WillyCheckItem = True
                    End If
                End If
            End If
        End With
    Next c
End Function

Private Function WillyCheckNasty( _
                 ) As Boolean

  Dim c   As Integer
  Dim x   As Integer
  Dim y   As Integer

    '-- Willy's
    With m_tWilly
        x = .x
        y = .y
    End With
    
    For c = 1 To UBound(m_tNasty())

        With m_tNasty(c)

            '-- Collide?
            If (.x = x \ 1 Or .x = x \ 1 + 1) And _
               (.y = y \ 8 Or .y = y \ 8 + 1 Or .y = y \ 8 + 2) Then
                GoTo collide
            End If
        End With
    Next c
    Exit Function

collide:
    WillyCheckNasty = True
End Function

Private Function WillyCheckGuardian( _
                 ) As Boolean

  Dim c   As Integer
  Dim x   As Integer
  Dim y   As Integer
  Dim DIB As New cDIB08

    '-- Willy's
    With m_tWilly
        x = .x
        y = .y
        If (m_aRoomID = 29) Then
            Set DIB = m_oDIBPig(.Frame)
          Else
            Set DIB = .DIB(.Frame)
        End If
    End With
    
    If (m_aRoomID = 35) Then ' Maria
        
        If (m_aItemsLeft > 0) Then
            
            With m_tMaria
                
                '-- Collide?
                If (FXImageCollide(DIB, 8 * (.x - x), 8 * .y - y, .DIB(.Frame))) Then
                    GoTo collide
                End If
            End With
        End If
      
      Else
      
        For c = 1 To UBound(m_tGuardian())
            
            With m_tGuardian(c)
                
                '-- Collide?
                If (.Type = 1) Then
                    If (FXImageCollide(DIB, 8 * (.x - x), .y - y, .DIB(.Frame - .FrameI))) Then
                        GoTo collide
                    End If
                  Else
                    If (FXImageCollide(DIB, 8 * (.x - x), .y - y, .DIB(.Frame))) Then
                        GoTo collide
                    End If
                End If
            End With
        Next c
    End If
    Exit Function

collide:
    WillyCheckGuardian = True
End Function

Private Function WillyCheckArrow( _
                 ) As Boolean

  Dim c   As Integer
  Dim x   As Integer
  Dim y   As Integer
  Dim DIB As cDIB08

    '-- Willy's
    With m_tWilly
        x = .x
        y = .y
        Set DIB = .DIB(.Frame)
    End With

    For c = 1 To UBound(m_tArrow())
        
        With m_tArrow(c)
            
            '-- Collide?
            If (FXImageCollide(DIB, 8 * (.x - x), .y - y, .DIB)) Then
                GoTo collide
            End If
        End With
    Next c
    Exit Function

collide:
    WillyCheckArrow = True
End Function

'----------------------------------------------------------------------------------------
' Items, guardians, panel...
'----------------------------------------------------------------------------------------

Private Sub DoItems()
    
  Dim c As Integer

    For c = 1 To UBound(m_tItem())
        
        With m_tItem(c)
            
            '-- Correct room?
            If (.Room = m_aRoomID) Then
            
                '-- Not collected?
                If (.Flag = 0) Then
                    
                    '-- Rotate ink (yellow-cyan-green-magenta)
                    .Ink = .Ink + 1
                    If (.Ink > 6) Then
                        .Ink = 3
                    End If
                    
                    '-- Render
                    Call MaskBltMask(m_oDIBFore, 8 * .x, 8 * .y, 8, 8, .Ink, m_oDIBItem)
                End If
            End If
        End With
    Next c
End Sub

Private Sub DoGuardians()
  
  Dim c As Integer
  
    For c = 1 To UBound(m_tGuardian())
    
        With m_tGuardian(c)
                    
            Select Case .Type
            
                Case 1 ' Horizontal
                     
                    '-- Depending on direction...
                    Select Case .FrameF - .FrameI
                        
                        Case 3 ' Single sequence
                        
                            If (.Dir = 1) Then                           ' right
                                If (.x = .Max And .Frame = .FrameF) Then ' right extreme reached
                                    .Dir = 0                             ' change dir
                                  Else
                                    If (.Frame < .FrameF) Then
                                        .Frame = .Frame + 1
                                      Else
                                        .Frame = .FrameI
                                        .x = .x + 1
                                    End If
                                End If
                              Else                                       ' left
                                If (.x = .Min And .Frame = .FrameI) Then ' left extreme reached
                                    .Dir = 1                             ' change dir
                                  Else
                                    If (.Frame > .FrameI) Then
                                        .Frame = .Frame - 1
                                     Else
                                        .Frame = .FrameF
                                         .x = .x - 1
                                    End If
                                End If
                            End If
                          
                        Case 7 ' Double sequence
                          
                            If (.Dir = 1) Then                           ' right
                                If (.x = .Max And .Frame = 7) Then       ' right extreme reached
                                    .Dir = 0                             ' change dir
                                    .Frame = 3
                                  Else
                                    If (.Frame < 7) Then
                                        .Frame = .Frame + 1
                                      Else
                                        .Frame = 4
                                        .x = .x + 1
                                    End If
                                End If
                              Else                                       ' left
                                If (.x = .Min And .Frame = 0) Then       ' left extreme reached
                                    .Dir = 1                             ' change dir
                                    .Frame = 4
                                  Else
                                    If (.Frame > 0) Then
                                        .Frame = .Frame - 1
                                     Else
                                        .Frame = 3
                                         .x = .x - 1
                                    End If
                                End If
                            End If
                    End Select
                   
                    '-- Render
                    Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, .Ink, .DIB(.Frame - .FrameI))
                
                Case 2 ' Vertical
                    
                    '-- Advance
                    If (.y + .Speed > .Max) Or _
                       (.y + .Speed < .Min) Then
                        .Speed = -.Speed
                    End If
                    .y = .y + .Speed
                        
                    '-- Normal/slow animation
                    .c = .c + .Fast
                    If (.c Mod 2 = 0) Then
                        .c = 0
                        .Frame = .Frame + 1
                        If (.Frame = 4) Then
                            .Frame = 0
                        End If
                    End If
                    
                    '-- Render
                    Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, .Ink, .DIB(.Frame))
            End Select
        End With
    Next c
End Sub

Private Sub DoArrows()

  Dim c As Integer
  
    For c = 1 To UBound(m_tArrow())
    
        With m_tArrow(c)
            
            '-- Move
            Select Case .Dir
           
                Case 1 ' Right
                    
                    '-- Loop
                    If (.x < 255) Then
                        .x = .x + 1
                      Else
                        .x = 0
                    End If
                    
                    '-- Sound FX 10 chars before
                    If (.x = 245) Then
                        Call FXPlay(m_hFXArrow)
                    End If
                    
                Case 0 ' Left
                    
                    '-- Loop
                    If (.x > 0) Then
                        .x = .x - 1
                      Else
                        .x = 255
                    End If
                    
                    '-- Sound FX 10 chars before
                    If (.x = 41) Then
                        Call FXPlay(m_hFXArrow)
                    End If
            End Select
                                
            '-- Render
            Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 8, 8, 7, .DIB)
        End With
    Next c
End Sub

Private Sub DoRopes()

  Dim c  As Integer
  Dim d  As Integer
  Dim o  As Byte
  Dim dx As Byte
  Dim dy As Byte
  Dim x  As Byte
  Dim y  As Byte
 
    For c = 1 To UBound(m_tRope())
    
        With m_tRope(c)
        
            '-- 1st dot
            x = .x
            y = .y
            Call FXRect(m_oDIBFore, x, y, 1, 1, .Ink)
            
            '-- Next dots
            For d = 0 To .Len - 1
                
                '-- Get dot dx & dy
                o = d + .Swing - Abs(.c)        ' LUT offset
                dx = m_aRopeTable(o)            ' dx value
                dy = m_aRopeTable(o + &H80) \ 2 ' dy value
                
                '-- Get dot position
                If (.Dir > 0) Then
                    x = x + dx
                  Else
                    x = x - dx
                End If
                y = y + dy
                
                '-- Draw dot
                Call FXRect(m_oDIBFore, x, y, 1, 1, .Ink)
                
                '-- Store
                With .Dot(d)
                    .x = x
                    .y = y
                End With
            Next d
            
            '-- Next step
            .c = .c + .Dir
            
            '-- Also adjust *speed* on range...
            If (.Dir > 0 And (.c > 34 Or .c < -38) Or _
                .Dir < 0 And (.c > 38 Or .c < -34) _
                ) Then
                .c = .c + .Dir
            End If
              
            '-- Swing!
            If (Abs(.c) = .Swing) Then
                .Dir = -.Dir
            End If
        End With
    Next c
End Sub

Private Sub DoSpecial()
    
    Select Case m_aRoomID
    
        Case 33 ' 'The Bathroom'

            '-- Toilet
            With m_tToilet
                
                '-- Choose sprites
                If (m_bVomit) Then
                    .Frame = .c + 2
                  Else
                    .Frame = .c
                End If
                
                '-- Render toilet
                Call BltMask(m_oDIBFore, 8 * .x, 8 * .y, 16, 16, CAPaper(m_aBlockCA(0)), IDX_WHITE, .DIB(.Frame))
                
                '-- Animation counter
                .c = .c + 1
                If (.c = 2) Then
                    .c = 0
                End If
            End With
            
        Case 35 ' 'Master Bedroom'
            
            '-- Maria
            With m_tMaria
                
                '-- At least, one item left
                If (m_aItemsLeft > 0) Then
                    
                    '-- Depending on Willy *y*
                    If (m_tWilly.y \ 8 < 13) Then
                        .Frame = 1 + (13 - m_tWilly.y \ 8)
                        If (.Frame > 3) Then
                            .Frame = 3
                        End If
                      Else
                        .Frame = .c \ 2
                    End If
                    
                    '-- Render Maria
                    Call BltMask(m_oDIBFore, 8 * .x, 8 * (.y + 0), 16, 8, CAPaper(m_aBlockCA(0)), IDX_BRCYAN, .DIB(.Frame), 0, 0)
                    Call BltMask(m_oDIBFore, 8 * .x, 8 * (.y + 1), 16, 8, CAPaper(m_aBlockCA(0)), IDX_WHITE, .DIB(.Frame), 0, 8)
                    
                    '-- Animation counter
                    .c = .c + 1
                    If (.c = 4) Then
                        .c = 0
                    End If
                End If
            End With
    End Select
End Sub

Private Sub DoPanel()
    
  Dim c As Integer
  Dim s As String
  
    With m_tPanel
        
        '-- Time string
        If (.t = 187200) Then
            .t = 14400
            If (.AMPM = "am") Then
                .AMPM = "pm"
              Else
                .AMPM = "am"
            End If
        End If
        
        s = Format$(.t \ 14400, "@@") & _
            ":" & _
            Format$((.t \ 240) Mod 60, "00") & _
            .AMPM
        
        '-- Render masked text
        Call FXText(m_oDIBFore, 0, 152, "Items collected", m_oDIBChar(), IDX_NULL, IDX_MASK)
        Call FXText(m_oDIBFore, 128, 152, Format$(m_aItems, "000"), m_oDIBChar(), IDX_NULL, IDX_MASK)
        Call FXText(m_oDIBFore, 160, 152, "Time", m_oDIBChar(), IDX_NULL, IDX_MASK)
        Call FXText(m_oDIBFore, 200, 152, s, m_oDIBChar(), IDX_NULL, IDX_MASK)
        
        '-- Render masked Willys?
        If (.Lives) Then
            For c = 1 To .Lives - 1
                Call BltFast(m_oDIBFore, c * 16 - 16, 168, 16, 16, m_tWilly.DIB(.c \ 4))
            Next c
        End If
        
        '-- Apply CA
        For c = 0 To 31
            Call FXMaskRect(m_oDIBFore, 8 * c, 152, 8, 8, IDX_NULL, CAPaper(m_aPanelCA(c + 96)), CAInk(m_aPanelCA(c + 96)))
        Next
        If (.Lives) Then
            For c = 0 To 2 * (.Lives - 1) - 1
                Call FXMaskRect(m_oDIBFore, 8 * c, 168, 8, 8, IDX_NULL, CAPaper(m_aPanelCA(c + 160)), CAInk(m_aPanelCA(c + 160)))
                Call FXMaskRect(m_oDIBFore, 8 * c, 176, 8, 8, IDX_NULL, CAPaper(m_aPanelCA(c + 192)), CAInk(m_aPanelCA(c + 192)))
            Next
        End If
        
        '-- Animate Willys sprites
        If (m_bTune) Then
            .c = .c + 1
            If (.c = 16) Then
                .c = 0
            End If
        End If
        
        '-- Time never stops
        .t = .t + 1
    End With
End Sub

'========================================================================================
' Misc
'========================================================================================

'----------------------------------------------------------------------------------------
' Screen
'----------------------------------------------------------------------------------------

Private Sub ScreenFlipBuffer( _
            Optional ByVal SkipPanel As Boolean = False _
            )
    
    '-- Get buffer bits
    If (SkipPanel) Then
        Call CopyMemory(ByVal m_oDIBFore.lpBits, ByVal m_oDIBBack.lpBits, 32768 + 2048)
      Else
        Call CopyMemory(ByVal m_oDIBFore.lpBits, ByVal m_oDIBBack.lpBits, 49152)
    End If
End Sub

Private Sub ScreenFlash(DIB As cDIB08)
    
  Dim c As Integer
  Dim i As Byte
  Dim p As Byte
  Dim b As Byte
    
    If (m_lc0 Mod 4 = 1) Then

        For c = 0 To 511
            
            '-- Has flash bit?
            If (m_aRoomFA(c) And &H80) Then
                
                '-- get base ink and paper
                i = (m_aRoomFA(c) And &H7)
                p = (m_aRoomFA(c) And &H38) \ 8
                
                '-- Bright bit?
                b = (m_aRoomFA(c) And &H40) \ 64
                
                '-- Swap base i-p
                m_aRoomFA(c) = m_aRoomFA(c) And Not &H3F
                m_aRoomFA(c) = m_aRoomFA(c) Or p Or i * 8
                
                '-- Add bright offset
                i = i + 8 * b
                p = p + 8 * b
                
                '-- *Inverse* colors
                Call FXMaskRect(DIB, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, p, i, p)
            End If
        Next c
    End If
End Sub
  
Private Sub ScreenAnimateConveyor()
    
  Dim c As Integer
  
    With m_tConveyor
        If (.Len > 0) Then
            For c = 0 To .Len - 1
                '-- FX
                Call FXConveyorBlock(m_oDIBBack, 8 * (.x + c), 8 * .y, .Dir)
                Call FXConveyorBlock(m_oDIBMask, 8 * (.x + c), 8 * .y, .Dir)
            Next c
        End If
    End With
End Sub
    
Public Sub ScreenUpdate()
    
    If (Not m_oForm Is Nothing) Then
        
        '-- Stretch 2x with 20-pixels border
        Call FXStretch2x(m_oDIBFore2x, m_oDIBFore, m_bFXTV)

        '-- Show info
        If (m_bJSWInfo) Then
            Call FXText(m_oDIBFore2x, 1, 1, m_sJSWInfo, m_oDIBChar(), IDX_WHITE, IDX_BLACK)
        End If
        
        '-- Show FPS
        If (m_bShowFPS) Then
            Call FXText(m_oDIBFore2x, 1, 1 + 9 * -m_bJSWInfo, Format$(m_lnFPS, "0000 FPS"), m_oDIBChar(), IDX_WHITE, IDX_BLACK)
        End If
        
        '-- Paint on given DC
        Call m_oDIBFore2x.Paint(m_oForm.hDC, (m_oForm.ScaleWidth - 592) \ 2, (m_oForm.ScaleHeight - 464) \ 2)
    End If
End Sub

'----------------------------------------------------------------------------------------
' Special renderings and FXs
'----------------------------------------------------------------------------------------

Private Sub RenderTitleScreen()
    
  Dim a()    As Byte
  
  Dim DIB(3) As New cDIB08
  Dim b      As Byte
  Dim m      As Integer
  Dim c      As Integer
  Dim i      As Byte
  Dim p      As Byte

    '-- Cls
    Call m_oDIBFore.Reset
    
    '-- Load triangle' packed bitmaps
    Call LoadData(DataFile(DATA_TRIANGLES), a(), 0, 32)
    
    '-- and unpack them
    For c = 0 To 3
        Call Unpack08(a(), 8 * c, DIB(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Load screen color-attributes
    Call LoadData(DataFile(DATA_TITLECA), a(), 0, 512)
    
    '-- Render screen
    For c = 0 To 511
        
        i = CAInk(a(c))   ' ink
        p = CAPaper(a(c)) ' paper
        
        '-- Store (flash FX)
        m_aRoomFA(c) = a(c)
        
        '-- Render masked triangles
        If (i < IDX_BRRED) Then           ' not JET SET WILLY blocks
            If (p <> i) Then              ' not *solid* block
                m = 100 * i + p           ' color *mask*
                Select Case m             ' choose triangle bitmap
                    Case 1, 105, 405, 500
                        b = 0 + c Mod 2
                    Case 5, 150, 504, 400
                        b = 2 + c Mod 2
                End Select
                '-- Render triangle mask
                Call BltFast(m_oDIBFore, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, DIB(b))
            End If
        End If
        
        '-- Apply CA
        If (m = 405) Then
            '-- Special case
            Call FXMaskRect(m_oDIBFore, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, IDX_MASK, p, i)
          Else
            '-- Any other
            Call FXMaskRect(m_oDIBFore, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, IDX_MASK, i, p)
       End If
    Next c
    
    '-- Render message's first 32 chars
    For c = 0 To 31
        Call BltMask(m_oDIBFore, c * 8, 152, 8, 8, IDX_BLACK, IDX_BRYELLOW, m_oDIBChar(m_aScrolly(c) - 32))
    Next c
End Sub

'----------------------------------------------------------------------------------------
' Decoding packed sprites
'----------------------------------------------------------------------------------------

Private Sub UnpackArrow( _
            ByRef Data As Byte, _
            ByRef DIB As cDIB08, _
            ByVal Ink As Byte, _
            ByVal Paper As Byte _
            )
  
  Dim t(63) As Byte
  Dim p     As Integer
    
    For p = 0 To 7
        t(8 + p) = Ink
    Next p
    For p = 0 To 7
        If (Data And m_aPow(7 - p)) Then
            t(0 + p) = Ink
            t(16 + p) = Ink
          Else
            t(0 + p) = Paper
            t(16 + p) = Paper
        End If
    Next p
    
    Call DIB.Create(8, 8)
    Call CopyMemory(ByVal DIB.lpBits, t(0), 64)
End Sub

Private Sub Unpack08( _
            ByRef Data() As Byte, _
            ByVal offset As Integer, _
            ByRef DIB As cDIB08, _
            ByVal Ink As Byte, _
            ByVal Paper As Byte _
            )
  
  Dim t(63) As Byte
  Dim p     As Integer
  Dim q     As Integer
    
    For p = 0 To 7
        For q = 0 To 7
            If (Data(offset + p) And m_aPow(7 - q)) Then
                t(p * 8 + q) = Ink
              Else
                t(p * 8 + q) = Paper
            End If
        Next q
    Next p
    Call DIB.Create(8, 8)
    Call CopyMemory(ByVal DIB.lpBits, t(0), 64)
End Sub

Private Sub Unpack16( _
            ByRef Data() As Byte, _
            ByVal offset As Integer, _
            ByRef DIB As cDIB08, _
            ByVal Ink As Byte, _
            ByVal Paper As Byte _
            )
  
  Dim t(255) As Byte
  Dim p      As Integer
  Dim q      As Integer
    
    For p = 0 To 31
        For q = 0 To 7
            If (Data(offset + p) And m_aPow(7 - q)) Then
                t(p * 8 + q) = Ink
              Else
                t(p * 8 + q) = Paper
            End If
        Next q
    Next p
    Call DIB.Create(16, 16)
    Call CopyMemory(ByVal DIB.lpBits, t(0), 256)
End Sub

'----------------------------------------------------------------------------------------
' Color attributes
'----------------------------------------------------------------------------------------

Private Function CAInk( _
                 ByVal CA As Byte _
                 ) As Byte
                 
    CAInk = (CA And &H7) + 8 * -((CA And &H40) <> 0)
End Function

Private Function CAPaper( _
                 ByVal CA As Byte _
                 ) As Byte
                 
    CAPaper = (CA And &H38) \ 8 + 8 * -((CA And &H40) <> 0)
End Function

Private Function CAFlash( _
                 ByVal CA As Byte _
                 ) As Byte

    CAFlash = CBool(CA And &H80)
End Function

'----------------------------------------------------------------------------------------
' Keys
'----------------------------------------------------------------------------------------

Public Sub KeyDown( _
           ByVal KeyCode As Integer _
           )
    
    '-- Disable pause
    m_bPause = False
    
    Select Case KeyCode
    
        Case vbKeyF1
        
            '-- Toggle FPS on/off
            m_bShowFPS = Not m_bShowFPS
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, True, m_bFXTV)
            
        Case vbKeyF5
        
            '-- Toggle info on/off
            m_bJSWInfo = Not m_bJSWInfo
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, True, m_bFXTV)
    
        Case vbKeyF8
        
            '-- Toggle full-screen/windowed
            Call mFullScreen.ToggleFullScreen
        
        Case vbKeyF11
        
            '-- FX TV scanlines
            m_bFXTV = Not m_bFXTV
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, True, m_bFXTV)
        
        Case vbKeyF12
        
            '-- FX color / green / black & white monitor
            m_bFXTVColor = m_bFXTVColor + 1
            If (m_bFXTVColor > 2) Then
                m_bFXTVColor = 0
            End If
            Select Case m_bFXTVColor
                Case 0 ' color
                    Call m_oDIBFore2x.SetPalette(m_aDefaultPal())
                Case 1 ' green
                    Call m_oDIBFore2x.SetPalette(m_aGreenPal())
                Case 2 ' black & white
                    Call m_oDIBFore2x.SetPalette(m_aBWPal())
            End Select
            
        Case vbKeyA, vbKeyS, vbKeyD, vbKeyF, vbKeyG
            
            '-- Store
            m_bKey(KeyCode) = True
    
            '-- Pause
            If (m_eMode = [eGamePlay]) Then
                m_bPause = True
                Call FXStop(m_hChannelTune)
            End If
        
        Case vbKeyH, vbKeyJ, vbKeyK, vbKeyL, vbKeyReturn
            
            '-- Store
            m_bKey(KeyCode) = True
            
            '-- Tune on/off
            If (m_eMode = [eGamePlay]) Then
                m_bTune = Not m_bTune
                If Not (m_bTune) Then
                    Call FXStop(m_hChannelTune)
                End If
            End If
            
        Case vbKeySubtract
            
            '-- Speed down (min. 1:1)
            If (m_snFrameFactor <> 0 And m_snFrameFactor < 1) Then
                m_snFrameFactor = m_snFrameFactor + 0.025
            End If
            
        Case vbKeyAdd
            
            '-- Speed up (max. 2.0:1)
            If (m_snFrameFactor > 0.5) Then
                m_snFrameFactor = m_snFrameFactor - 0.025
            End If
            
        Case vbKeyMultiply
            
            '-- Toggles 1:1 speed - maximum speed
            If (m_snFrameFactor = 0) Then
                m_snFrameFactor = 1
              Else
                m_snFrameFactor = 0
            End If
            
        Case Else
            
            '-- Store
            m_bKey(KeyCode) = True
            
            '-- Check cheat stream
            Call KeysCheckCheatCode(KeyCode)
    End Select
End Sub

Public Sub KeyUp( _
           ByVal KeyCode As Integer _
           )
           
    '-- Reset key
    m_bKey(KeyCode) = False
    
    '-- Reset Willy direction flag?
    With m_tWilly
        If (.mode <> [eJump]) And Not (.OnConveyor) Then
            .Flag = 0
        End If
    End With
End Sub

Private Function KeysCheckAnyKey( _
                 ) As Boolean
  
  Dim c As Integer
    
    '-- Check for any pressed key
    For c = 1 To 255
        If (m_bKey(c)) Then
            KeysCheckAnyKey = True
            Exit For
        End If
    Next c
End Function

Private Sub KeysCheckWillyKeys( _
            KeyCode As Byte _
            )
    
    '-- Right/left
    With m_tWilly
        If ((.Flag And 2) = 0) Then
            KeyCode = KeyCode Or 1 * -(m_bKey(vbKeyW) Or m_bKey(vbKeyR) Or m_bKey(vbKeyY) Or m_bKey(vbKeyI) Or m_bKey(vbKeyP) Or m_bKey(vbKeyRight))
        End If
        If ((.Flag And 1) = 0) Then
            KeyCode = KeyCode Or 2 * -(m_bKey(vbKeyQ) Or m_bKey(vbKeyE) Or m_bKey(vbKeyT) Or m_bKey(vbKeyU) Or m_bKey(vbKeyO) Or m_bKey(vbKeyLeft))
        End If
    End With
    
    '-- Jump
    KeyCode = KeyCode Or 4 * -(m_bKey(vbKeyZ) Or m_bKey(vbKeyX) Or m_bKey(vbKeyC) Or m_bKey(vbKeyV) Or m_bKey(vbKeyB) Or m_bKey(vbKeyUp) Or m_bKey(vbKeyN) Or m_bKey(vbKeyM) Or m_bKey(vbKeySpace) Or m_bKey(vbKeyShift) Or m_bKey(226) Or m_bKey(188) Or m_bKey(190) Or m_bKey(189))
End Sub

Private Sub KeysCheckCheatCode( _
            ByVal KeyCode As Integer _
            )
    
    '-- Already cheated?
    If (m_bCheated = False) Then
        
        '-- Playing?
        If (m_eMode = [eGamePlay]) Then
            
            '-- Correct room?
            If (m_aRoomID = 28) Then
        
                '-- Add char to stream
                m_sCheatCode = Right$(m_sCheatCode, Len(m_sCheatCode) - 1) & UCase$(Chr$(KeyCode))
        
                '-- Well?
                m_bCheated = (m_sCheatCode = JSW_CHEATCODE)
            End If
        End If
    End If
End Sub

Private Sub KeysCheckTeleporterMask()

  Dim k As Integer
  Dim m As Byte
  
    If (m_bCheated) Then
        
        '-- Key '9' flag?
        If (m_bKey(vbKey9)) Then
            
            '-- Keys '1' to '6': build room # mask
            For k = vbKey1 To vbKey6
                If (m_bKey(k)) Then
                    m = m + m_aPow(k - vbKey1)
                End If
            Next k
            
            '-- Valid?
            If (m < 61) Then
                
                '-- Turn to 'fall' mode
                With m_tWilly
                    .mode = [eFall]
                    .Flag = 0
                    .c = 18
                    .y = .y And &HF8
                    .f = .y
                End With
                
                '-- Safe Willy
                LSet m_tWillySafe = m_tWilly
                
                '-- Force room initialization
                Call InitializeRoom(m)
            End If
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------
' Notes and music
'----------------------------------------------------------------------------------------

Private Sub PlayInGameTune()

    If (m_bTune) Then
        If (m_hChannelTune = 0) Then
            m_hChannelTune = FXPlay(m_hFXTone, 75, , True)
        End If
        Call FXChangeFreq(m_hChannelTune, 100 * GetNoteFreq(GetJSWNote(m_aGameTune(m_lc2)) + 48))
        m_lc1 = m_lc1 + 1
        If (m_lc1 = 2) Then
            m_lc1 = 0
            m_lc2 = m_lc2 + 1
            If (m_lc2 = 64) Then
                m_lc2 = 0
            End If
        End If
    End If
End Sub

Private Function GetJSWNote( _
                 ByVal Note As Byte _
                 ) As Byte
  
  Dim c As Integer
  Dim b As Integer
    
    If (Note < 16) Then
        GetJSWNote = &HFF
      Else
        b = UBound(m_aNoteINV())
        For c = 0 To b
            If (Note <= m_aNoteINV(c)) Then
                GetJSWNote = b - c
                Exit For
            End If
        Next c
    End If
End Function

Private Function GetNoteFreq( _
                 ByVal Note As Integer, _
                 Optional ByVal Tone As Integer = 440 _
                 ) As Single
    
    GetNoteFreq = (Tone / 32) * (2 ^ ((Note - 9) / 12))
End Function

Private Function FXPlay( _
                 ByVal hFX As Long, _
                 Optional ByVal Vol As Long = -1, _
                 Optional ByVal freq As Long = -1, _
                 Optional ByVal Looped As Boolean = False _
                 ) As Long
    
  Dim h As Long
    
    h = FSOUND_PlaySound(FSOUND_FREE, hFX)
    If (h <> 0) Then
        If (Vol <> -1) Then
            Call FSOUND_SetVolume(h, Vol)
        End If
        If (freq <> -1) Then
            Call FSOUND_SetFrequency(h, freq)
        End If
        If (Looped) Then
            Call FSOUND_SetLoopMode(h, FSOUND_LOOP_NORMAL)
        End If
        FXPlay = h
    End If
End Function

Private Sub FXStop( _
            hChannel As Long _
            )
            
    Call FSOUND_StopSound(hChannel)
    hChannel = 0
End Sub

Private Sub FXChangeFreq( _
            ByVal hChannel As Long, _
            ByVal freq As Long _
            )
            
    Call FSOUND_SetFrequency(hChannel, freq)
End Sub

Private Sub FXResetAll()
    
    '-- In case...
    Call FSOUND_StopSound(m_hChannelN1)
    m_hChannelN1 = 0
    Call FSOUND_StopSound(m_hChannelN2)
    m_hChannelN2 = 0
    Call FSOUND_StopSound(m_hChannelTune)
    m_hChannelTune = 0
End Sub

'----------------------------------------------------------------------------------------
' File I/O
'----------------------------------------------------------------------------------------

Private Sub LoadData( _
            ByVal File As String, _
            ByRef Data() As Byte, _
            ByVal offset As Long, _
            ByVal length As Long _
            )
    
  Dim hFile As Integer
    
    hFile = VBA.FreeFile()
    Open File For Binary Access Read As #hFile
        ReDim Data(length - 1)  ' ensure space
        Seek #hFile, offset + 1 ' offset
        Get #hFile, , Data()    ' get data
    Close #hFile
End Sub

Private Sub SaveData( _
            ByVal File As String, _
            ByRef Data() As Byte _
            )
    
  Dim hFile As Integer
    
    hFile = VBA.FreeFile()
    Open File For Binary Access Write As #hFile
        Put #hFile, , Data()    ' write data
    Close #hFile
End Sub

Private Function AppPath( _
                 ) As String
    
    '-- Fix path
    If (Right$(App.Path, 1) = "\") Then
        AppPath = App.Path
      Else
        AppPath = App.Path & "\"
    End If
End Function

Private Function DataFile( _
                 ByVal FileID As Long _
                 ) As String
    
    '-- Build data file path
    DataFile = AppPath & "Data\" & CStr(FileID) & ".bin"
End Function
