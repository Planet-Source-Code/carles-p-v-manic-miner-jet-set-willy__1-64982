Attribute VB_Name = "mMM"
'==============================================================================
'
' MANIC MINER
'
' Based on the first release for the ZX Spectrum
' by Matthew Smith - Bug-Byte Ltd ©1983
'
' Author:        Carles P.V.
' Version:       1.2.0
' Date:          15-Apr-2006
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

Private Const DATA_WILLY                    As Long = 33280
Private Const DATA_TTUNE                    As Long = 33902
Private Const DATA_GTUNE                    As Long = 34188
Private Const DATA_SCROLLY                  As Long = 40192
Private Const DATA_TITLECA                  As Long = 40448
Private Const DATA_TITLEPX                  As Long = 40960
Private Const DATA_ROOMS                    As Long = 45056

'-- Room data offsets

Private Const OFFSET_SCREENLAYOUT           As Long = 0
Private Const OFFSET_ROOMNAME               As Long = 512
Private Const OFFSET_BLOCKGRAPHICS          As Long = 544
Private Const OFFSET_WILLYSSTART            As Long = 616
Private Const OFFSET_CONVEYOR               As Long = 623
Private Const OFFSET_BORDERCOLOR            As Long = 627 ' 627
Private Const OFFSET_ITEMS                  As Long = 629 ' 653
Private Const OFFSET_PORTAL                 As Long = 655
Private Const OFFSET_ITEMGRAPHIC            As Long = 692
Private Const OFFSET_AIR                    As Long = 700
Private Const OFFSET_HORZGUARDIANS          As Long = 702
Private Const OFFSET_VERTGUARDIANS          As Long = 733
Private Const OFFSET_SPECIAL                As Long = 736
Private Const OFFSET_GUARDIANGRAPHICS       As Long = 768

'-- Block type constants

Private Const BLOCK_AIR                     As Byte = 0
Private Const BLOCK_FLOOR                   As Byte = 1
Private Const BLOCK_CRUMBLINGFLOOR          As Byte = 2
Private Const BLOCK_WALL                    As Byte = 3
Private Const BLOCK_CONVEYOR                As Byte = 4
Private Const BLOCK_NASTY1                  As Byte = 5
Private Const BLOCK_NASTY2                  As Byte = 6
Private Const BLOCK_SPARE                   As Byte = 7

'-- Color index constants

Private Const IDX_BLACK                     As Byte = 0
Private Const IDX_BLUE                      As Byte = 1
Private Const IDX_RED                       As Byte = 2
Private Const IDX_MAGENTA                   As Byte = 3
Private Const IDX_GREEN                     As Byte = 4
Private Const IDX_CYAN                      As Byte = 5
Private Const IDX_YELLOW                    As Byte = 6
Private Const IDX_WHITE                     As Byte = 7

Private Const IDX_BRBLACK                   As Byte = 8
Private Const IDX_BRBLUE                    As Byte = 9
Private Const IDX_BRRED                     As Byte = 10
Private Const IDX_BRMAGENTA                 As Byte = 11
Private Const IDX_BRGREEN                   As Byte = 12
Private Const IDX_BRCYAN                    As Byte = 13
Private Const IDX_BRYELLOW                  As Byte = 14
Private Const IDX_BRWHITE                   As Byte = 15

Private Const IDX_NULL                      As Byte = 0
Private Const IDX_MASK                      As Byte = 255

'-- Quite important constants

Private Const DYDEATH                       As Byte = 36
Private Const YFDEATH                       As Byte = &HFF
Private Const MM_NULLNOTE                   As Byte = &HFF
Private Const MM_EXTRALIVE                  As Long = 10000
Private Const MM_CHEATCODE                  As String = "6031769"
Private Const MM_INFO                       As String = "4361726C657320502E562E207F20323030362D32303132"

'-- Private types

Private Type tWilly
    Ink         As Byte                     ' fore color
    x           As Byte                     ' current x pos [chrs]
    y           As Integer                  ' current y pos [pxs]
    Dir         As Byte                     ' direction 0: right / 1: left
    Frame       As Byte                     ' current frame
    OnConveyor  As Boolean                  ' Willy's on conveyor
    mode        As eWillyMode               ' 0: standing-walking / 1: jumping / 2: falling
    Flag        As Byte                     ' 0: standing / 1: right / 2: left
    f           As Byte                     ' fall counter (-> max height)
    c           As Byte                     ' internal counter (-> jump)
End Type

Private Type tItem
    Ink         As Byte                     ' fore color
    Paper       As Byte                     ' back color
    x           As Byte                     ' x pos [chrs]
    y           As Byte                     ' y pos [chrs]
    Flag        As Byte                     ' 0/1: eaten
End Type

Private Type tConveyor                      ' (animation)
    x           As Byte                     ' x pos [chrs] (left extreme)
    y           As Byte                     ' y pos [chrs]
    Dir         As Byte                     ' direction 1: right / 0: left
    Len         As Byte                     ' animation length [chrs]
End Type

Private Type tGuardianV
    Ink         As Byte                     ' fore color
    Paper       As Byte                     ' back color
    y1          As Byte                     ' top extreme [pxs]
    y2          As Byte                     ' bottom extreme [pxs]
    x           As Byte                     ' x pos [chrs]
    y           As Integer                  ' current y pos [pxs]
    Speed       As Integer                  ' dy [pxs]
    Frame       As Byte                     ' current frame
End Type

Private Type tGuardianH
    Ink         As Byte                     ' fore color
    Paper       As Byte                     ' back color
    x1          As Byte                     ' left extreme [chrs]
    x2          As Byte                     ' right extreme [chrs]
    x           As Byte                     ' current x pos [chrs]
    y           As Byte                     ' y pos [chrs]
    Dir         As Byte                     ' direction 0: right / 1: left
    Speed       As Byte                     ' speed mode  0: normal / 1: slow (half)
    Frame       As Byte                     ' current frame
    c           As Byte                     ' internal counter (-> speed)
End Type

Private Type tEugene
    DIB         As New cDIB08               ' Eugene's graphic (special)
    Ink         As Byte                     ' fore color (fixed)
    Paper       As Byte                     ' back color (fixed)
    x           As Byte                     ' x pos (fixed)
    y           As Integer                  ' current y pos [pxs]
    y1          As Byte                     ' top extreme [pxs] (fixed)
    y2          As Byte                     ' bottom extreme [pxs] (fixed)
    Speed       As Integer                  ' dy [pxs] (fixed)
    Flag        As Byte                     ' 0/1: activated
End Type

Private Type tKong
    Ink         As Byte                     ' fore color (fixed)
    Paper       As Byte                     ' back color (fixed)
    x           As Byte                     ' x pos (fixed)
    y           As Byte                     ' current y pos [pxs]
    y1          As Byte                     ' top extreme [pxs] (fixed->fall)
    y2          As Byte                     ' bottom extreme [pxs] (fixed->fall)
    Frame       As Byte                     ' current frame
    Flag        As Byte                     ' 0/1/2: falling and dead
    c           As Byte                     ' internal counter (-> fall speed)
End Type

Private Type tPortal
    Ink         As Byte                     ' fore color
    Paper       As Byte                     ' back color
    x           As Byte                     ' x pos [chrs]
    y           As Byte                     ' y pos [chrs]
    Flag        As Byte                     ' 0/1: open
    c           As Byte                     ' internal counter (-> flash speed)
End Type

Private Type tAir
    Cur         As Byte                     ' current value
    c           As Byte                     ' internal counter (-> speed)
End Type

Private Type tPanel
    Lives       As Byte                     ' Willy's lives
    Extra       As Byte                     ' extra lives counter
    SC          As Long                     ' current score
    HI          As Long                     ' current high-score
    c1          As Byte                     ' internal counter (-> animation speed)
    c2          As Byte                     ' internal counter (-> Willy's live frame)
End Type

Private Type tNasty
    x           As Byte                     ' x pos [chrs]
    y           As Byte                     ' y pos [chrs]
End Type

Private Type tSwitch
    x           As Byte                     ' x pos [chrs]
    y           As Byte                     ' y pos [chrs]
    Flag        As Byte                     ' 0/1/2: activated and already activated
    c           As Byte                     ' internal counter (-> animation)
End Type

'-- Private enums.

Private Enum eMode
    [eIntro] = 0                            ' *dancing* "MANIC-MINER" intro
    [eTitle] = 1                            ' title tune + scrolly
    [eDemo] = 2                             ' demo
    [eGamePlay] = 3                         ' playing
    [eGameSuccess] = 4                      ' room successfully finished
    [eGameDie] = 5                          ' ouch!
    [eGameOver] = 6                         ' boot in action
    [eGameEnd] = 7                          ' last room reached (and finished) without cheating
End Enum

Private Enum eWillyMode
    [eWalk] = 0                             ' walking/standing (also on rope)
    [eJump] = 1                             ' jumping
    [eFall] = 2                             ' falling
End Enum

'-- Private variables

Private m_oForm             As Form         ' destination Form
Private m_aDefaultPal()     As Byte         ' default palette
Private m_aGreenPal()       As Byte         ' green palette
Private m_aBWPal()          As Byte         ' greyscale palette

Private m_oDIBBack          As New cDIB08   ' back buffer 1
Private m_oDIBFore          As New cDIB08   ' back buffer 2
Private m_oDIBMask          As New cDIB08   ' mask DIB (FX purposes)
Private m_oDIBFore2x        As New cDIB08   ' 2x screen
Private m_oDIBChar(95)      As New cDIB08   ' font char graphics

Private m_eMode             As eMode        ' current mode
Private m_aRoomID           As Byte         ' current room #

Private m_aRoomData()       As Byte         ' room data (1Kb)
Private m_aRoomBlock(511)   As Byte         ' room block layout
Private m_aBlockCA(7)       As Byte         ' blocks' CA

Private m_tWilly            As tWilly       ' Willy
Private m_oDIBWilly(7)      As New cDIB08   ' Willy graphics

Private m_tGuardianH()      As tGuardianH   ' horizontal guardians
Private m_tGuardianV()      As tGuardianV   ' vertical guardians
Private m_oDIBGuardian(7)   As New cDIB08   ' guardian graphics
Private m_bHasGuardianVs    As Boolean      ' flag
Private m_tEugene           As tEugene      ' Eugene (special case)
Private m_tKong             As tKong        ' Kong   (special case)

Private m_tConveyor         As tConveyor    ' conveyor (animation)
Private m_tItem()           As tItem        ' items
Private m_oDIBItem          As New cDIB08   ' item graphic
Private m_tNasty()          As tNasty       ' nasties
Private m_tSwitch()         As tSwitch      ' switches

Private m_tPortal           As tPortal      ' portal
Private m_oDIBPortal        As New cDIB08   ' portal graphic
Private m_tAir              As tAir         ' air supply
Private m_tPanel            As tPanel       ' panel (scores/lives)

Private m_aMANIC()          As Byte         ' MANIC blocks (intro)
Private m_aMINER()          As Byte         ' MINER blocks (intro)
Private m_aScrolly()        As Byte         ' title message
Private m_aTitleTune()      As Byte         ' title tune
Private m_aGameTune()       As Byte         ' in-game tune
Private m_oDIBSpecial(2)    As New cDIB08   ' special graphics

'-- Misc. variables

Private m_aPow(7)           As Byte         ' quick 2^x
Private m_aNoteINV()        As Byte         ' LUT MM-note
Private m_bKey(255)         As Boolean      ' LUT keys state

Private m_lFrameDt          As Long         ' current frame time interval (loop)
Private m_snFrameFactor     As Single       ' time interval speed-factor
Private m_bPause            As Boolean      ' flag (loop)
Private m_bExit             As Boolean      ' flag (loop)

Private m_sCheatCode        As String * 7   ' cheat code string
Private m_bCheated          As Boolean      ' flag
Private m_bTune             As Boolean      ' flag
Private m_bFXTV             As Boolean      ' flag (FX TV scanlines)
Private m_bFXTVColor        As Byte         ' flag (FX green / black & white monitor)

Private m_hFXJump           As Long         ' sound FX handles
Private m_hFXDead           As Long
Private m_hFXTone           As Long

Private m_hChannelN1        As Long         ' channel handles
Private m_hChannelN2        As Long
Private m_hChannelTune      As Long
Private m_hChannelFX        As Long

Private m_nNoteLen          As Integer      ' notes length and value
Private m_nNote1            As Integer
Private m_nNote2            As Integer

Private m_lc0               As Long         ' counters
Private m_lc1               As Long
Private m_lc2               As Long
Private m_lc3               As Long
Private m_lc4               As Long

Private m_ltFPS             As Long         ' fps
Private m_lcFPS             As Long
Private m_lnFPS             As Long
Private m_bShowFPS          As Boolean

Private m_bMMInfo           As Boolean      ' info
Private m_sMMInfo           As String


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
    
    '-- Back and fore DIBs (first and second buffers)
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
    
    '-- Load intro screen 'title':
    '   arrays storing [x,y,W,H,color] for 8x8 blocks
    m_aMANIC() = VB.LoadResData("MANIC", "MM")
    m_aMINER() = VB.LoadResData("MINER", "MM")
       
    '-- Load title scrolly
    Call LoadData(DataFile(DATA_SCROLLY), m_aScrolly(), 0, 256)
    
    '-- Load title and in-game tunes
    Call LoadData(DataFile(DATA_TTUNE), m_aTitleTune(), 0, 286)
    Call LoadData(DataFile(DATA_GTUNE), m_aGameTune(), 0, 64)
        
    '-- load high-score value
    Call LoadHiSc
    
    '-- Sprites -------------------------------------------------------------------------
    
    '-- Load Willy graphics
    Call LoadData(DataFile(DATA_WILLY), a(), 0, 256)
    For c = 0 To 7
        Call Unpack16(a(), 32 * c, m_oDIBWilly(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Load special graphics
    For c = 0 To 2
        '-- Get room data
        Call LoadData(DataFile(DATA_ROOMS), m_aRoomData(), 1024 * c, 1024)
        '-- Unpack graphic
        Call Unpack16(m_aRoomData(), OFFSET_SPECIAL, m_oDIBSpecial(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Sound -------------------------------------------------------------------------
    
    '-- MM LUT note
    m_aNoteINV() = LoadResData("NINV", "ZX")
    
    '-- Play tune
    m_bTune = True
    
    '-- FMOD and sound FXs
    Call FSOUND_Init(44100, 8, 0)
    m_hFXTone = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\tone.fx", FSOUND_NORMAL, 0, 0)
    m_hFXJump = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\jump.fx", FSOUND_NORMAL, 0, 0)
    m_hFXDead = FSOUND_Sample_Load(FSOUND_FREE, AppPath & "SFX\dead.fx", FSOUND_NORMAL, 0, 0)
    
    '-- Info ----------------------------------------------------------------------------
    
    For c = 1 To Len(MM_INFO) Step 2
        m_sMMInfo = m_sMMInfo & Chr$("&H" & Mid$(MM_INFO, c, 2))
    Next c
End Sub

Public Sub Terminate()
    
    '-- Reset
    Call FXStopAll
    
    '-- Free all
    Call FSOUND_Sample_Free(m_hFXJump)
    Call FSOUND_Sample_Free(m_hFXDead)
    Call FSOUND_Sample_Free(m_hFXTone)
    
    '-- Close FMOD
    Call FSOUND_Close
    
    '-- Save high-score value
    Call SaveHiSc
End Sub

'========================================================================================
' Main loop
'========================================================================================

Public Sub StartGame()
  
  Dim t As Long
    
    '-- Intro mode
    Call SetMode([eIntro])
    
    '-- Default speed-factor
    m_snFrameFactor = 1
    
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
        Case [eIntro]
            m_lFrameDt = 80
        Case [eTitle]
            m_lFrameDt = 60
        Case [eDemo]
            m_lFrameDt = 70
        Case [eGamePlay]
            m_lFrameDt = 70
        Case [eGameSuccess]
            m_lFrameDt = 20
        Case [eGameDie]
            m_lFrameDt = 15
        Case [eGameOver]
            m_lFrameDt = 70
        Case [eGameEnd]
            m_lFrameDt = 60
    End Select
    
    '-- Reset all counters
    m_lc0 = 0
    m_lc1 = 0
    m_lc2 = 0
    m_lc3 = 0
    m_lc4 = 0
    
    '-- Stop all channels
    Call FXStopAll
End Sub

'========================================================================================
' Loop main routines
'========================================================================================

Private Sub DoFrame()
    
    Select Case m_eMode
    
        Case [eIntro]
            
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
            
        Case [eDemo]
        
            Call DoDemo
            
            If (KeysCheckAnyKey()) Then
                Call SetMode([eTitle])
                m_bKey(vbKeyReturn) = False
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
            
        Case [eGameSuccess]
        
            Call DoGameSuccess
            
        Case [eGameOver]
            
            Call DoGameOver
        
        Case [eGameEnd]
            
            Call DoGameEnd
    End Select
End Sub

Private Sub DoIntro()
    
    '-- Toggle MANIC-MINER
    If (m_lc0 Mod 5 = 0) Then
        Call m_oDIBFore.Reset
        Call RenderIntroPart(m_lc0 Mod 2)
    End If
    Call ScreenUpdate
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoTitle()
    
    Select Case m_lc0
        
        Case 0
            
            '-- Render background
            Call RenderTitlePart(ID:=0)
            Call RenderTitlePart(ID:=1)
            Call RenderTitlePart(ID:=2)
            Call BltMask(m_oDIBBack, 236, 72, 16, 16, IDX_RED, IDX_WHITE, m_oDIBWilly(0))
            
        Case 1 To 300
            
            '-- Play notes
            If (m_lc1 = 0) Then
            
                '-- Get notes and length
                m_nNoteLen = (m_aTitleTune(m_lc2) + 1) / 10 - 2
                m_nNote2 = GetMMNote(m_aTitleTune(m_lc2 + 2)) + 48
                m_nNote1 = GetMMNote(m_aTitleTune(m_lc2 + 1)) + 48
                
                '-- Overlapped? Nullify N2
                If (m_nNote2 = m_nNote1 - 1) Then
                    m_nNote2 = MM_NULLNOTE
                End If
                
                '-- Flip buffer
                Call ScreenFlipBuffer
                
                '-- Play and draw red note
                If (m_nNote2 < MM_NULLNOTE) Then
                    '-- Start FX
                    m_hChannelN2 = FXPlay(m_hFXTone, , 75 * GetNoteFreq(m_nNote2), True)
                    '-- Draw red note
                    Call FXRect(m_oDIBFore, Int(m_nNote2 * 0.8 - 40) * 8, 120, 7, 8, IDX_BRRED)
                End If
                
                '-- Play and draw cyan note
                If (m_nNote1 < MM_NULLNOTE) Then
                    '-- Start FX
                    m_hChannelN1 = FXPlay(m_hFXTone, , 75 * GetNoteFreq(m_nNote1), True)
                    '-- Draw cyan note
                    Call FXRect(m_oDIBFore, Int(m_nNote1 * 0.8 - 40) * 8, 120, 7, 8, IDX_CYAN)
                End If
            End If
            
            '-- Update border + screen
            Call FXBorder(m_oDIBFore2x, m_nNote1 Mod 8, , m_bFXTV)
            Call ScreenUpdate
            
            '-- Counter (# frames = length)
            m_lc1 = m_lc1 + 1
            If (m_lc1 = m_nNoteLen) Then
                m_lc1 = 0
                m_lc2 = m_lc2 + 3
                Call FXStop(m_hChannelN1)
                Call FXStop(m_hChannelN2)
            End If
            
        Case 301
            
            '-- Erase both last notes
            Call FXRect(m_oDIBFore, Int(m_nNote2 * 0.8 - 40) * 8, 120, 7, 8, IDX_WHITE)
            Call FXRect(m_oDIBFore, Int(m_nNote1 * 0.8 - 40) * 8, 120, 7, 8, IDX_WHITE)
            
            '-- Reset counters and change framing dt
            m_lc1 = 0
            m_lc2 = 0
            m_lc3 = 0
            m_lc4 = 0
            m_lFrameDt = 100
            
        Case Else
            
            '-- Print message...
            For m_lc1 = m_lc2 To m_lc2 + 31
                Call BltMask(m_oDIBFore, (m_lc1 - m_lc2) * 8, 152, 8, 8, IDX_BLACK, IDX_BRYELLOW, m_oDIBChar(m_aScrolly(m_lc1) - 32))
            Next m_lc1
            
            '-- Scroll 1-char left / check end
            m_lc2 = m_lc2 + 1
            If (m_lc2 = 256 - 31) Then
                Call SetMode([eDemo])
                Exit Sub
            End If
            
            '-- Dancing Willy (/2)
            Call BltMask(m_oDIBFore, 236, 72, 16, 16, IDX_RED, IDX_WHITE, m_oDIBWilly(m_lc4))
            m_lc3 = m_lc3 + 1
            If (m_lc3 = 2) Then
                m_lc3 = 0
                m_lc4 = m_lc4 + 2
                If (m_lc4 = 8) Then
                    m_lc4 = 0
                End If
            End If
            
            '-- Update screen
            Call ScreenUpdate
    End Select
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoDemo()

    Select Case m_lc0
        
        Case 0
            
            '-- Initialize
            Call InitializePlay
            
        Case 1
        
            '-- Load room
            Call LoadRoom(m_aRoomID)
            Call InitializeRoom

        Case 2 To 65
           
            '-- Play tune if enabled
            Call PlayInGameTune
           
            '-- Do it all except Willy
            Call ScreenFlipBuffer
            Call DoGuardians
            Call DoItems
            Call DoSolarRay
            Call DoPortal
            Call DoAir
            Call DoPanel
            Call DoConveyor
            Call ScreenUpdate
            
        Case 66
        
            '-- Stop and reset tune channel
            Call FXStop(m_hChannelTune)
            
            '-- Speed up
            m_lFrameDt = 20
            
        Case 67 To 81
            
            '-- FX
            Call BltMask(m_oDIBFore, 0, 0, 256, 128, 81 - m_lc0, 81 - m_lc0 + 1, m_oDIBMask)
            Call ScreenUpdate
        
        Case 82
            
            '-- Next/first room
            m_aRoomID = m_aRoomID + 1
            If (m_aRoomID = 20) Then
                m_aRoomID = 0
            End If
            
            '-- Speed down and reset main counter (loop -> m_lc0=1)
            m_lFrameDt = 70
            m_lc0 = 0
    End Select
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoGamePlay()
    
    '-- Play in-game tune and check if
    '   room mask-code has been typed
    Call PlayInGameTune
    Call KeysCheckCheatMask
    
    '-- Flip buffer
    Call ScreenFlipBuffer
    
    '-- Do all animations
    Call DoGuardians
    Call DoItems
    Call DoWilly
    Call DoSolarRay
    Call DoPortal
    Call DoAir
    Call DoPanel
    Call DoConveyor
    
    '-- Refresh screen
    Call ScreenUpdate
End Sub

Private Sub DoGameDie()
    
    Select Case m_lc2
    
        Case 0
            
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
            
            With m_tPanel
            
                '-- Sorry...
                .Lives = .Lives - 1
    
                '-- Game over or try again
                If (.Lives = 0) Then
                    Call SetMode([eGameOver])
                  Else
                    Call SetMode([eGamePlay])
                    Call InitializeRoom
                End If
            End With
            Exit Sub
    End Select
    
    '-- Counter (m_lc0 & m_lc1 used by in-game tune)
    m_lc2 = m_lc2 + 1
End Sub

Private Sub DoGameSuccess()

    '-- Last room?
    If (m_aRoomID = 19 And m_bCheated = False) Then
        Call SetMode([eGameEnd])
        Exit Sub
    End If
    
    '-- Not last...
    Select Case m_lc2
    
        Case 0 To 15
            
            '-- FX
            Call BltMask(m_oDIBFore, 0, 0, 256, 128, 15 - m_lc2, 16 - m_lc2, m_oDIBMask)
            Call ScreenUpdate
        
        Case 16
            
            '-- Black border
            Call FXBorder(m_oDIBFore2x, IDX_BLACK, , m_bFXTV)
            
            '-- Start FX
            m_hChannelFX = FXPlay(m_hFXTone, , 1000 + 150 * m_tAir.Cur, True)
            
        Case Else
        
            With m_tAir
                
                '-- Air points
                If (.Cur > 0) Then
                    If (.Cur > 1) Then
                        .Cur = .Cur - 2
                        Call ScoreAdd(18)
                      Else
                        .Cur = .Cur - 1
                        Call ScoreAdd(9)
                    End If
                    
                    '-- Update bar and score
                    Call FXRect(m_oDIBFore, 32, 136, 80, 8, IDX_BRRED)
                    Call FXRect(m_oDIBFore, 80, 136, 176, 8, IDX_BRGREEN)
                    Call FXRect(m_oDIBFore, 32, 138, .Cur, 4, IDX_BRWHITE)
                    Call FXText(m_oDIBFore, 208, 152, Format$(m_tPanel.SC, "000000"), m_oDIBChar(), IDX_BLACK, IDX_BRYELLOW)
                    Call ScreenUpdate
                    
                    '-- Sound FX
                    Call FXChangeFreq(m_hChannelFX, 1500 + 150 * .Cur)
                        
                  Else
                  
                    '-- Stop FX
                    Call FXStop(m_hChannelFX)
                  
                    '-- Next/first room
                    m_aRoomID = m_aRoomID + 1
                    If (m_aRoomID > 19) Then
                        m_aRoomID = 0
                    End If
                    
                    '-- Initialize room
                    Call LoadRoom(m_aRoomID)
                    Call InitializeRoom
                    Call SetMode([eGamePlay])
                    Exit Sub
                End If
            End With
    End Select
    
    '-- Counter
    m_lc2 = m_lc2 + 1
End Sub

Private Sub DoGameOver()
    
    Select Case m_lc0
    
        Case 0
            
            '-- Prepare background
            Call FXRect(m_oDIBBack, 0, 0, 256, 128, IDX_BLACK)
            Call BltFast(m_oDIBBack, 120, 0, 16, 16, m_oDIBSpecial(2))
            Call BltFast(m_oDIBBack, 124, 96, 16, 16, m_oDIBWilly(0))
            Call BltFast(m_oDIBBack, 120, 112, 16, 16, m_oDIBSpecial(1))
            
        Case 1 To 48
            
            '-- Boot
            Call BltFast(m_oDIBBack, 120, m_lc0 * 2, 16, 16, m_oDIBSpecial(2))
            
            '-- Masked screen
            Call BltMask(m_oDIBFore, 0, 0, 256, 128, m_lc0 Mod 4 + 8, IDX_BRWHITE, m_oDIBBack)
            Call ScreenUpdate
            
            '-- Sound FX
            Call FXPlay(m_hFXJump, , 5000 + 500 * m_lc0)
            
            '-- Speed up boot
            If (m_lc0 Mod 4 = 0) Then
                m_lFrameDt = m_lFrameDt - 5
            End If
                                
        Case 49 To 149
            
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
            
        Case 150
        
            '-- HI-SC?
            With m_tPanel
                If (.SC > .HI) Then
                    .HI = .SC
                End If
            End With
            
            '-- Title mode
            Call SetMode([eTitle])
            Exit Sub
    End Select
    
    '-- Counter
    m_lc0 = m_lc0 + 1
End Sub

Private Sub DoGameEnd()

    Select Case m_lc0
    
        Case 0
            
            '-- Background
            Call ScreenFlipBuffer
            
            '-- Willy out!
            Call MaskBltMask(m_oDIBFore, 152, 16, 16, 16, IDX_WHITE, m_oDIBWilly(m_tWilly.Frame))
            
            '-- Sword-fish!
            Call BltFast(m_oDIBFore, 152, 40, 16, 16, m_oDIBSpecial(0))
            Call FXMaskRect(m_oDIBFore, 152, 40, 8, 8, IDX_BLACK, IDX_BLACK, IDX_CYAN)
            Call FXMaskRect(m_oDIBFore, 160, 40, 8, 8, IDX_BLACK, IDX_BLACK, IDX_CYAN)
            Call FXMaskRect(m_oDIBFore, 152, 48, 8, 8, IDX_BLACK, IDX_BLACK, IDX_YELLOW)
            Call FXMaskRect(m_oDIBFore, 160, 48, 8, 8, IDX_BLACK, IDX_BLACK, IDX_WHITE)
            
            '-- Refresh
            Call ScreenUpdate
            
        Case 1
            
            '-- Start FX
            m_hChannelFX = FXPlay(m_hFXTone, , 75000 + 25000 * (m_lc0 Mod 2), True)
            
        Case 2 To 97
            
            '-- Sound FX
            Call FXChangeFreq(m_hChannelFX, 75000 + 25000 * (m_lc0 Mod 2))
            
        Case 98
            
            '-- Stop FX
            Call FXStop(m_hChannelFX)
            
        Case 99
            
            '-- Start again!
            Call SetMode([eGameSuccess])
            
            '-- Last room plus one (avoid repeating DoGameSuccess())
            m_aRoomID = 20
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
        .Lives = 3
        .Extra = 0
        .SC = 0
    End With
    
    '-- First room
    m_aRoomID = 0
    Call LoadRoom(m_aRoomID)
    Call InitializeRoom
End Sub

Private Sub InitializeRoom()
    
    '-- Initialize all
    Call InitializeBackground
    Call InitializeConveyor
    Call InitializeSwitches
    Call InitializeItems
    Call InitializeNasties
    Call InitializeGuardians
    Call InitializePortal
    Call InitializeAir
    Call InitializeWilly
    Call InitializeScores
End Sub

Private Sub InitializeBackground()
  
  Dim c           As Integer
  Dim o           As Integer
  Dim a           As Byte
  Dim ii(255)     As Byte
  Dim DIBBlock(7) As New cDIB08
    
    '-- Border color
    o = OFFSET_BORDERCOLOR                  ' offset
    a = m_aRoomData(o) And &H7              ' packed index
    Call FXBorder(m_oDIBFore2x, a, , m_bFXTV)
    
    '-- Unpack 8 8x8 blocks
    For c = 0 To 7
        
        o = OFFSET_BLOCKGRAPHICS + 9 * c    ' offset
        a = m_aRoomData(o)                  ' packed color-attributes
        
        If (a <> 0) Then                    ' store as inverse index
            ii(a) = c
        End If
        
        '-- Unpack
        Call Unpack08(m_aRoomData(), o + 1, DIBBlock(c), IDX_MASK, IDX_NULL)
    Next c
    
    '-- Get block IDs and CAs: render blocks
    For c = 0 To 511
        
        o = OFFSET_SCREENLAYOUT + c         ' offset
        a = m_aRoomData(o)                  ' packed color-attributes
        
        m_aRoomBlock(c) = ii(a)             ' block ID
        m_aBlockCA(ii(a)) = a               ' block color-attributes
        
        '-- Render
        Call BltMask(m_oDIBBack, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, CAPaper(a), CAInk(a), DIBBlock(ii(a)))
        Call BltFast(m_oDIBMask, 8 * (c Mod 32), 8 * (c \ 32), 8, 8, DIBBlock(ii(a)))
    Next c
    
    '-- Special case: last room
    If (m_aRoomID = 19) Then
        Call RenderTitlePart(ID:=0)
    End If
    
    '-- Render panel background
    Call RenderTitlePart(ID:=2)
    
    '-- Render room name
    o = OFFSET_ROOMNAME                     ' offset
    For c = 0 To 31
        Call MaskBltMask(m_oDIBBack, c * 8, 128, 8, 8, IDX_BLACK, m_oDIBChar(m_aRoomData(o + c) - 32))
    Next c
End Sub

Private Sub InitializeWilly()
    
  Dim o  As Integer
  Dim a2 As Byte
  Dim a3 As Byte
  Dim a5 As Byte
  Dim a6 As Byte
  
    o = OFFSET_WILLYSSTART  ' offset
    a2 = m_aRoomData(o + 1) ' starting frame
    a3 = m_aRoomData(o + 2) ' starting dir
    a5 = m_aRoomData(o + 4) ' packed starting position
    a6 = m_aRoomData(o + 5)
    
    With m_tWilly
        
        '-- Starting xy
        .x = (a5 And &H1F)
        .y = (a5 And &HE0) \ &H20 + (a6 And &H1) * &H8
        .y = 8 * .y ' to pixels
        
        '-- Starting dir
        .Dir = a3
        .Flag = .Dir
        
        '-- Starting frame
        .Frame = a2 + 4 * .Dir
        
        '-- Mode walking/standing
        .mode = [eWalk]
        .OnConveyor = False
        
        '-- Reset counters
        .c = 0
        .f = 0
    End With
End Sub

Private Sub InitializeItems()
    
  Dim c  As Integer
  Dim n  As Byte
  Dim o  As Integer
  Dim a1 As Byte
  Dim a2 As Byte
  Dim a3 As Byte

    ReDim m_tItem(0)
    n = 0
    
    For c = 0 To 4                  ' 5 items max
        
        o = OFFSET_ITEMS + 5 * c    ' offset
        a1 = m_aRoomData(o + 0)     ' packed color-attributes
        a2 = m_aRoomData(o + 1)     ' packed position
        a3 = m_aRoomData(o + 2)
        
        Select Case a1
            
            Case &H0  ' no item
            
            Case &HFF ' end of sequence
                
                Exit For
            
            Case Else ' valid item
                
                n = n + 1
                ReDim Preserve m_tItem(n)
                
                With m_tItem(n)
                    
                    '-- Color
                    .Ink = CAInk(a1)
                    .Paper = CAPaper(a1)
                    
                    '-- x, y
                    .x = (a2 And &H1F)
                    .y = (a2 And &HE0) \ &H20 + (a3 And &H1) * &H8
                End With
        End Select
    Next c
    
    '-- Graphic
    Call Unpack08(m_aRoomData(), OFFSET_ITEMGRAPHIC, m_oDIBItem, IDX_MASK, IDX_NULL)
End Sub

Private Sub InitializeNasties()
  
  Dim c As Integer
  Dim n As Byte
    
    ReDim m_tNasty(0)
    n = 0
    
    '-- Get positions
    For c = 0 To 511
        If (m_aRoomBlock(c) = BLOCK_NASTY1 Or m_aRoomBlock(c) = BLOCK_NASTY2) Then
            n = n + 1
            ReDim Preserve m_tNasty(n)
            With m_tNasty(n)
                .x = c Mod 32
                .y = c \ 32
            End With
        End If
    Next c
End Sub

Private Sub InitializeSwitches()
  
  Dim c As Integer
  Dim n As Byte
    
    ReDim m_tSwitch(0)
    n = 0
    
    '-- Get positions
    If (m_aRoomID = 7 Or m_aRoomID = 11) Then
        For c = 0 To 511
            If (m_aRoomBlock(c) = BLOCK_SPARE) Then
                n = n + 1
                ReDim Preserve m_tSwitch(n)
                With m_tSwitch(n)
                    .x = c Mod 32
                    .y = c \ 32
                    .Flag = 0
                End With
            End If
        Next c
    End If
End Sub

Private Sub InitializeConveyor()
    
  Dim o  As Integer
  Dim a1 As Byte
  Dim a2 As Byte
  Dim a3 As Byte
  Dim a4 As Byte
  
    o = OFFSET_CONVEYOR     ' offset
    a1 = m_aRoomData(o + 0) ' mode (direction)
    a2 = m_aRoomData(o + 1) ' packed position
    a3 = m_aRoomData(o + 2)
    a4 = m_aRoomData(o + 3) ' length
    
    With m_tConveyor
        
        '-- Dir (0/1: left/right)
        .Dir = a1
        
        '-- Left anchor
        .x = (a2 And &H1F)
        .y = (a2 And &HE0) \ &H20 + (a3 And &H8)
        
        '-- Length (chrs)
        .Len = a4
    End With
End Sub

Private Sub InitializeGuardians()
    
  Dim c  As Integer
  Dim n  As Byte
  Dim o  As Integer
  Dim a1 As Byte
  Dim a2 As Byte
  Dim a3 As Byte
  Dim a4 As Byte
  Dim a5 As Byte
  Dim a6 As Byte
  Dim a7 As Byte
  
    '-- Verticals -----------------------------------------------------------------------
    
    ReDim m_tGuardianV(0)
    
    Select Case m_aRoomID
        
        Case 4     ' Eugene
            
            Call InitializeEugene
            m_bHasGuardianVs = False
            
        Case 7, 11 ' Kong
            
            Call InitializeKong
            m_bHasGuardianVs = True
                
        Case Else
            
            n = 0
            For c = 0 To 3                          ' 4 vert. guardians max
                
                o = OFFSET_VERTGUARDIANS + 7 * c    ' offset
                a1 = m_aRoomData(o + 0)             ' packed color-attributes
                a2 = m_aRoomData(o + 1)             ' starting frame
                a3 = m_aRoomData(o + 2)             ' packed starting y position
                a4 = m_aRoomData(o + 3)             ' starting x position
                a5 = m_aRoomData(o + 4)             ' packed speed
                a6 = m_aRoomData(o + 5)             ' packed extreme top position
                a7 = m_aRoomData(o + 6)             ' packed extreme bottom position
                
                Select Case a1
                    
                    Case &HFF ' end of sequence
                        
                        Exit For
                    
                    Case Else ' valid guardian
                        
                        n = n + 1
                        ReDim Preserve m_tGuardianV(n)
                        
                        With m_tGuardianV(n)
                            
                            '-- Color
                            .Ink = CAInk(a1)
                            .Paper = CAPaper(a1)
                            
                            '-- Starting frame
                            .Frame = a2
                            
                            '-- Starting x, y
                            .x = a4                            ' [chrs]
                            .y = (a3 And &H78) + (a3 And &H7)  ' [pxs]
                            
                            '-- Range
                            .y1 = (a6 And &H78) + (a6 And &H7) ' [pxs]
                            .y2 = (a7 And &H78) + (a7 And &H7) ' [pxs]
                            
                            '-- Speed (pxs/frame)
                            .Speed = IIf(a5 > &H7F, a5 - &H100, a5)
                        End With
                End Select
            Next c
            m_bHasGuardianVs = (UBound(m_tGuardianV()) > 0)
    End Select
    
    '-- Horizontals ---------------------------------------------------------------------
    
    ReDim m_tGuardianH(0)
    
    n = 0
    For c = 0 To 4                          ' 5 horz. guardians max
        
        o = OFFSET_HORZGUARDIANS + 7 * c    ' offset
        a1 = m_aRoomData(o + 0)             ' packed color-attributes
        a2 = m_aRoomData(o + 1)             ' packed starting position
        a3 = m_aRoomData(o + 2)
        a5 = m_aRoomData(o + 4)             ' starting frame
        a6 = m_aRoomData(o + 5)             ' packed extreme left position
        a7 = m_aRoomData(o + 6)             ' packed extreme right position
        
        a4 = m_aRoomData(o + 3)
        Select Case a1
            
            Case &H0  ' no guardian
            
            Case &HFF ' end of sequence
                
                Exit For
            
            Case Else ' valid guardian
                
                n = n + 1
                ReDim Preserve m_tGuardianH(n)
                
                With m_tGuardianH(n)
                    
                    '-- Color
                    .Ink = CAInk(a1)
                    .Paper = CAPaper(a1)
                    
                    '-- Speed (normal/slow)
                    .Speed = (a1 And &H80) \ &H80
                    
                    '-- Starting x, y
                    .x = (a2 And &H1F)                             ' chrs
                    .y = (a2 And &HE0) \ &H20 + (a3 And &H1) * &H8 ' chrs
                    
                    '-- Range
                    .x1 = (a6 And &H1F)                            ' chrs
                    .x2 = (a7 And &H1F)                            ' chrs
                    
                    '-- Starting frame and dir (*)
                    If (m_bHasGuardianVs) Then
                        .Frame = IIf(a5 = 0, 4, 7)
                        .Dir = IIf(a5 = 0, 0, 1)
                      Else
                        .Frame = a5
                        .Dir = IIf(a5 < 4, 0, 1)
                    End If
                    
                    '-- Counter ?! (*)
                    .c = .Speed * -(m_aRoomID <> 12)
                End With
        End Select
    Next c
    
    '-- Graphics
    For c = 0 To 7
        Call Unpack16(m_aRoomData(), OFFSET_GUARDIANGRAPHICS + 32 * c, m_oDIBGuardian(c), IDX_MASK, IDX_NULL)
    Next c

' *: Can't get correct initial frame... (see room 12).
'    In fact, problem comes with 'slow' guardians.
'    Is something hardcoded?
End Sub

Private Sub InitializeEugene()
    
    With m_tEugene
        
        '-- Color
        .Ink = IDX_WHITE
        .Paper = IDX_RED
        
        '-- Starting x, y
        .x = 15
        .y = 0
        
        '-- Range
        .y1 = 0
        .y2 = 88
        
        '-- Speed (and dir)
        .Speed = 1
        
        '-- Flag (activated?)
        .Flag = 0
    
        '-- Graphic (as special)
        Call Unpack16(m_aRoomData(), OFFSET_SPECIAL, .DIB, IDX_MASK, IDX_NULL)
    End With
End Sub

Private Sub InitializeKong()
    
    With m_tKong
    
        '-- Color
        .Ink = IDX_BRGREEN
        .Paper = IDX_BLACK
        
        '-- Starting x, y
        .x = 15
        .y = 0
        
        '-- Range
        .y1 = 0
        .y2 = 104
        
        '-- Starting frame
        .Frame = 0
        
        '-- Flag (falling?)
        .Flag = 0
        
        '-- Counter
        .c = 6
    End With
End Sub

Private Sub InitializePortal()
    
  Dim o   As Integer
  Dim a1  As Byte
  Dim a34 As Byte
  Dim a35 As Byte
  
    o = OFFSET_PORTAL           ' offset
    a1 = m_aRoomData(o + 0)     ' packed color-attributes
    a34 = m_aRoomData(o + 33)   ' packed position
    a35 = m_aRoomData(o + 34)
    
    With m_tPortal
        
        '-- Color
        .Ink = CAInk(a1)
        .Paper = CAPaper(a1)
        
        '-- x, y
        .x = (a34 And &H1F)
        .y = (a34 And &HE0) \ &H20 + (a35 And &H1) * &H8
        
        '-- Flag (open?)
        .Flag = 0
        
        '-- Counter
        .c = 0
    End With
    
    '-- Graphic
    Call m_oDIBPortal.Create(16, 16)
    Call Unpack16(m_aRoomData(), o + 1, m_oDIBPortal, IDX_MASK, IDX_NULL)
End Sub

Private Sub InitializeAir()
        
  Dim o  As Integer
  Dim a1 As Byte
  Dim a2 As Byte
  Dim c  As Integer
   
    o = OFFSET_AIR          ' offset
    a1 = m_aRoomData(o + 0) ' 4px-chars
    a2 = m_aRoomData(o + 1) ' extra pixels
    
    With m_tAir
        
        '-- Supply (# 4px-chars + # extra pixels)
        Do: c = c + 1: Loop Until m_aPow(c) = &H100 - a2
        
        '-- Current
        .Cur = 4 * a1 - c - 26
        
        '-- Counter
        .c = 6
    End With
End Sub

Private Sub InitializeScores()
    
    With m_tPanel
    
        '-- "High Score 000000"
        Call FXText(m_oDIBBack, 0, 152, "High Score " & Format$(.HI, "000000"), m_oDIBChar(), IDX_BLACK, IDX_BRYELLOW)
        
        '-- "Score 000000"
        Call FXText(m_oDIBBack, 160, 152, "Score " & Format$(.SC, "000000"), m_oDIBChar(), IDX_BLACK, IDX_BRYELLOW)
        
        '-- Reset counters
        .c1 = 0
        .c2 = 0
    End With
End Sub

'//

Private Sub LoadRoom( _
            ByVal n As Byte _
            )
    
    '-- Get room data (1024 bytes)
    Call LoadData(DataFile(DATA_ROOMS), m_aRoomData(), 1024 * n, 1024)
    
    '-- Store number
    m_aRoomID = n
End Sub

'----------------------------------------------------------------------------------------
' Willy rountines
'----------------------------------------------------------------------------------------

Private Sub DoWilly()
  
  Dim Keys As Byte
    
    With m_tWilly
                 
        '-- Check keys (right/left/jump)
        Call KeysCheckWillyKeys(Keys)               ' check right/left/jump keys
        
        '-- Previous checks
        If (.mode = [eWalk]) Then                   ' only if walking
                
            '-- Fall?
            If Not (WillyCheckFeet()) Then          ' nothing under feet?
                .mode = [eFall]                     ' fall mode
                .c = 18                             ' fall start counter
                .f = .y                             ' max. height start counter
                GoTo skip
            End If
                
            '-- Conveyor?
            Call WillyCheckConveyor(Keys)           ' on conveyor?

            '-- Crumbling block?
            Call WillyCheckCrumblingFloor           ' on crumbling floor?
        End If
        
        '-- Process depending on mode...
        Select Case .mode

            Case 0 ' walking/standing
                
                '-- Reset counter
                .c = 0
                    
                '-- Right and/or left keys pressed
                Select Case Keys And 3

                    Case 1                          ' right

                        If (.Dir = 1) Then          ' facing left
                            .Dir = 0                ' change dir
                            .Flag = 0               ' nullify direction flag
                            .Frame = .Frame - 4     ' toggle frame counter
                          Else                      ' already going right
                            .Flag = 1               ' right flag
                            Call WillyRight         ' go right
                        End If

                    Case 2                          ' left

                        If (.Dir = 0) Then          ' facing right
                            .Dir = 1                ' change dir
                            .Flag = 0               ' nullify direction flag
                            .Frame = .Frame + 4     ' toggle frame counter
                          Else                      ' already going left
                            .Flag = 2               ' left flag
                            Call WillyLeft          ' go left
                        End If

                    Case Else                       ' nor right nor left
                        .Flag = 0                   ' nullify direction flag
                End Select
                
                '-- Jump key pressed
                If (Keys And 4) Then
                    If Not (WillyCheckHead()) Then  ' nothing on head
                        .mode = [eJump]             ' ok, jump
                      Else                          ' ouch!
                        .mode = [eFall]             ' fall
                        .c = 18
                        .f = .y
                    End If
                End If
         
            Case 1 ' jumping
           
                Call WillyJump                      ' jump!
                
                Select Case .mode
                    
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
                
            Case 2 ' falling

                Call WillyFall                      ' fall!
                
                If (.f = YFDEATH) Then              ' death flag?
                    Call SetMode([eGameDie])        ' yes: die
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
                            
        '-- Check item/switch/portal/nasty/guardian
        Select Case True
            
            Case WillyCheckItem()
                '-- Do nothing: continue
                
            Case WillyCheckSwitch()
                '-- Do nothing: continue
            
            Case WillyCheckPortal()
                '-- Change mode
                Call SetMode([eGameSuccess])
            
            Case WillyCheckNasty()
                '-- Change mode
                Call SetMode([eGameDie])
            
            Case WillyCheckGuardian()
                '-- Change mode
                Call SetMode([eGameDie])
        End Select
        
skip:   '-- Render Willy
        Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, IDX_WHITE, m_oDIBWilly(.Frame))
    End With
End Sub

Private Sub WillyLeft()

    With m_tWilly
        If (.Frame > 4) Then
            .Frame = .Frame - 1               ' previous frame
          Else
            If Not (WillyCheckLeft()) Then
                .Frame = 7                    ' last frame
                .x = .x - 1                   ' one block left
            End If
        End If
    End With
End Sub

Private Sub WillyRight()
    
    With m_tWilly
        If (.Frame < 3) Then
            .Frame = .Frame + 1               ' next frame
          Else
            If Not (WillyCheckRight()) Then
                .Frame = 0                    ' first frame
                .x = .x + 1                   ' one block right
            End If
        End If
    End With
End Sub

Private Sub WillyJump()
    
    With m_tWilly
        If (WillyCheckHead() And .c < 9) Then ' going up
            .mode = [eWalk]
          Else
            .y = .y + ((.c And &HFE) - 8) \ 2
            .c = .c + 1
            Select Case .c
                Case Is = 9                   ' top
                    .f = .y
                Case Is > 9                   ' going down
                    If (WillyCheckFeet()) Then
                        .mode = [eWalk]
                      Else
                        If (.c = 18) Then
                            .mode = [eFall]
                        End If
                    End If
            End Select
        End If
        
        '-- Sound FX
        If (.mode <> [eWalk]) Then
            Call FXPlay(m_hFXJump, , 22050 + 1500 * (9 - Sqr((.c - 9) ^ 2)))
        End If
    End With
End Sub

Private Sub WillyFall()

    With m_tWilly
        If (WillyCheckFeet()) Then
            If (.y - .f > DYDEATH) Then       ' max. height reached
                .f = YFDEATH                  ' death flag
              Else                            ' saved!
                .mode = [eWalk]
                .Flag = 0
            End If
          Else
            .y = .y + 4                       ' go down
            If (.c < 26) Then                 ' counter (sound FX)
                .c = .c + 1
              Else
                .c = 22
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
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
    
    With m_tWilly
        o = .x + 2 + (.y \ 8 + 0) * 32
        b1 = m_aRoomBlock(o + 0)
        b2 = m_aRoomBlock(o + 32)
        b3 = m_aRoomBlock(o + 64)
        Let WillyCheckRight = (b1 = BLOCK_WALL) Or (b2 = BLOCK_WALL) Or ((b3 = BLOCK_WALL) And ((.c > 10) Or (.y Mod 8 > 0)))
    End With
End Function

Private Function WillyCheckLeft( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
    
    With m_tWilly
        o = .x - 1 + (.y \ 8 + 0) * 32
        b1 = m_aRoomBlock(o + 0)
        b2 = m_aRoomBlock(o + 32)
        b3 = m_aRoomBlock(o + 64)
        Let WillyCheckLeft = (b1 = BLOCK_WALL) Or (b2 = BLOCK_WALL) Or ((b3 = BLOCK_WALL) And ((.c > 10) Or (.y Mod 8 > 0)))
    End With
End Function

Private Function WillyCheckHead( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Byte
  Dim b2 As Byte
    
    With m_tWilly
        If (.y Mod 8 = 0) Then
            o = .x + (.y \ 8 - 1) * 32
            b1 = m_aRoomBlock(o + 0)
            b2 = m_aRoomBlock(o + 1)
            Let WillyCheckHead = (b1 = BLOCK_WALL) Or (b2 = BLOCK_WALL)
        End If
    End With
End Function

Private Function WillyCheckFeet( _
                 ) As Boolean
    
  Dim o  As Integer
  Dim b1 As Byte
  Dim b2 As Byte
  
    With m_tWilly
        If (.y Mod 8 = 0) Then
            o = .x + (.y \ 8 + 2) * 32
            b1 = m_aRoomBlock(o + 0)
            b2 = m_aRoomBlock(o + 1)
            Let WillyCheckFeet = (b1 > BLOCK_AIR And b1 < BLOCK_NASTY1) Or (b1 = BLOCK_SPARE) Or (b2 > BLOCK_AIR And b2 < BLOCK_NASTY1) Or (b2 = BLOCK_SPARE)
                                 
            '-- Check if on conveyor
            .OnConveyor = (b1 = BLOCK_CONVEYOR Or b2 = BLOCK_CONVEYOR)
        End If
    End With
End Function
               
Private Sub WillyCheckConveyor( _
            Keys As Byte _
            )
        
             
    With m_tWilly
        
        If (.OnConveyor) Then
    
             Select Case m_tConveyor.Dir
                 
                 Case 0 ' right
                    
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
                
                Case 1 ' left
                    
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

Private Sub WillyCheckCrumblingFloor()

  Dim o As Integer
  Dim c As Integer
  Dim f As Boolean
    
    With m_tWilly
        
        '-- Offset
        o = .x + (.y \ 8 + 2) * 32
        
        '-- Two blocks
        For c = 0 To 1
            
            If (m_aRoomBlock(o + c) = BLOCK_CRUMBLINGFLOOR) Then
                
                '-- 1 pixel down
                Call FXCrumblingBlock(m_oDIBBack, 8 * (.x + c), .y + 16, CAPaper(m_aRoomData(o + c)), f)
                Call FXCrumblingBlock(m_oDIBMask, 8 * (.x + c), .y + 16, IDX_BLACK, False)
                
                '-- Dropped?
                If (f) Then
                    
                    '-- Turn it background
                    m_aRoomBlock(o + c) = BLOCK_AIR
                    
                    '-- Paint with background color
                    Call FXRect(m_oDIBBack, 8 * (.x + c), .y + 16, 8, 8, CAPaper(m_aBlockCA(0)))
                    Call FXRect(m_oDIBMask, 8 * (.x + c), .y + 16, 8, 8, IDX_BLACK)
                End If
            End If
        Next c
    End With
End Sub
   
Private Function WillyCheckItem( _
                 ) As Boolean

  Dim c As Integer
  Dim b As Boolean
    
    '-- Portal closed?
    If (m_tPortal.Flag = 0) Then
    
        For c = 1 To UBound(m_tItem())
            
            With m_tItem(c)
                
                '-- Not eaten?
                If (.Flag = 0) Then
                    
                    '-- At least, one found
                    b = True
                   
                    '-- Collide?
                    If (.x = m_tWilly.x \ 1 Or .x = m_tWilly.x \ 1 + 1) And _
                       (.y = m_tWilly.y \ 8 Or .y = m_tWilly.y \ 8 + 1 Or ((.y = m_tWilly.y \ 8 + 2) And (m_tWilly.y Mod 8 > 0))) Then
                        
                        '-- Eat it!
                        .Flag = 1
                        
                        '-- 100 points!
                        Call ScoreAdd(100)
                        Let WillyCheckItem = True
                    End If
                End If
            End With
        Next c
        
        If Not (b) Then         ' no items found
            m_tPortal.Flag = 1  ' open portal
            With m_tEugene      ' activate Eugene, if any
                .Flag = 1
                .Ink = IDX_BRMAGENTA
                .Speed = 1
            End With
        End If
    End If
End Function

Private Function WillyCheckSwitch( _
                 ) As Boolean

  Dim c As Integer
    
    '-- Need to check?
    If (m_aRoomID = 7 Or m_aRoomID = 11) Then
    
        For c = 1 To UBound(m_tSwitch())
            
            With m_tSwitch(c)
                
                '-- Not activated?
                If (.Flag = 0) Then
                    
                    '-- Collide?
                    If (.x = m_tWilly.x \ 1 Or .x = m_tWilly.x \ 1 + 1) And _
                       (.y = m_tWilly.y \ 8) Then
                        
                        '-- Activate it!
                        .Flag = 1
                        
                        '-- FX
                        Call FXSwitchBlock(m_oDIBBack, 8 * .x, 8 * .y)
                        Call FXSwitchBlock(m_oDIBMask, 8 * .x, 8 * .y)
                        Let WillyCheckSwitch = True
                        Exit For
                    End If
                End If
            End With
        Next c
         
        '-- Depending on switch...
        Select Case True
            
            Case m_tSwitch(1).Flag = 1 ' wall
                
                With m_tSwitch(1)
                    
                    '-- Small animation...
                    .c = .c + 1
                    If (.c < 9) Then
                        
                        '-- FX
                        Call FXRect(m_oDIBBack, 17 * 8, 12 * 8 - .c, 8, 2 * .c, CAPaper(m_aRoomData(11 * 32 + 17)))
                      
                      Else
                        '-- Already activated flag
                        m_tSwitch(1).Flag = 2
                        
                        '-- Destroy blocks
                        m_aRoomBlock(11 * 32 + 17) = BLOCK_AIR
                        m_aRoomBlock(12 * 32 + 17) = BLOCK_AIR
                        Call FXRect(m_oDIBBack, 17 * 8, 11 * 8, 8, 8, IDX_BLACK)
                        Call FXRect(m_oDIBBack, 17 * 8, 12 * 8, 8, 8, IDX_BLACK)
                        Call FXRect(m_oDIBMask, 17 * 8, 11 * 8, 8, 8, IDX_NULL)
                        Call FXRect(m_oDIBMask, 17 * 8, 12 * 8, 8, 8, IDX_NULL)
                        
                        '-- Change guardian right margin
                        m_tGuardianH(2).x2 = m_tGuardianH(2).x2 + 3
                    End If
                End With

            Case m_tSwitch(2).Flag = 1 ' Kong
                
                '-- <already activated> flag
                m_tSwitch(2).Flag = 2
                
                '-- Destroy blocks
                m_aRoomBlock(2 * 32 + 15) = BLOCK_AIR
                m_aRoomBlock(2 * 32 + 16) = BLOCK_AIR
                Call FXRect(m_oDIBBack, 15 * 8, 2 * 8, 8, 8, IDX_BLACK)
                Call FXRect(m_oDIBBack, 16 * 8, 2 * 8, 8, 8, IDX_BLACK)
                Call FXRect(m_oDIBMask, 15 * 8, 2 * 8, 8, 8, IDX_NULL)
                Call FXRect(m_oDIBMask, 16 * 8, 2 * 8, 8, 8, IDX_NULL)
                
                '-- Activate Kong!
                With m_tKong
                    .Flag = 1
                    .Ink = IDX_YELLOW
                End With
        End Select
    End If
End Function

Private Function WillyCheckNasty( _
                 ) As Boolean

  Dim c As Integer
  
    For c = 1 To UBound(m_tNasty())
        
        With m_tNasty(c)
            
            '-- Collide?
            If (.x = m_tWilly.x \ 1 Or .x = m_tWilly.x \ 1 + 1) And _
               (.y = m_tWilly.y \ 8 Or .y = m_tWilly.y \ 8 + 1 Or .y = m_tWilly.y \ 8 + 2) Then
                GoTo collide
            End If
        End With
    Next c
    Exit Function

collide:
    Let WillyCheckNasty = True
End Function

Private Function WillyCheckGuardian( _
                 ) As Boolean
  
  Dim c    As Integer
    
  Dim wx   As Integer
  Dim wy   As Integer
  Dim wDIB As cDIB08
    
    '-- Store Willy's position in pixels
    With m_tWilly
        wx = 8 * .x
        wy = .y
        Set wDIB = m_oDIBWilly(.Frame)
    End With
    
    '-- Verticals
    Select Case m_aRoomID
        
        Case 4     ' Eugene
            
            With m_tEugene
                '-- Collide?
                If (FXImageCollide(wDIB, 8 * .x - wx, .y - wy, .DIB)) Then
                    GoTo collide
                End If
            End With
        
        Case 7, 11 ' Kong
            
            With m_tKong
                '-- Active?
                If (.Flag = 0) Then
                    ' collide?
                    If (FXImageCollide(wDIB, 8 * .x - wx, .y - wy, m_oDIBGuardian(.Frame))) Then
                        GoTo collide
                    End If
                End If
            End With
        
        Case Else  ' normal guardians
            
            For c = 1 To UBound(m_tGuardianV())
                With m_tGuardianV(c)
                    '-- Collide?
                    If (FXImageCollide(wDIB, 8 * .x - wx, .y - wy, m_oDIBGuardian(.Frame))) Then
                        GoTo collide
                    End If
                End With
            Next c
    End Select
    
    '-- Horizontals
    For c = 1 To UBound(m_tGuardianH())
        With m_tGuardianH(c)
            If (FXImageCollide(wDIB, 8 * .x - wx, 8 * .y - wy, m_oDIBGuardian(.Frame))) Then
                GoTo collide
            End If
        End With
    Next c
    
    Exit Function

collide:
    Let WillyCheckGuardian = True
End Function

Private Function WillyCheckPortal( _
                 ) As Boolean

    With m_tWilly
        
        '-- Open?
        If (m_tPortal.Flag = 1) Then
            
            '-- Inside?
            If ((m_tPortal.x = .x) And Abs(8 * m_tPortal.y - .y) < 8) Then
                Let WillyCheckPortal = True
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Items, conveyor, switches, guardians, portal...
'----------------------------------------------------------------------------------------

Private Sub DoItems()
    
  Dim c As Integer

    For c = 1 To UBound(m_tItem())
        
        With m_tItem(c)
            
            '-- Not eaten?
            If (.Flag = 0) Then
                
                '-- Rotate ink (yellow-cyan-green-magenta)
                .Ink = .Ink - 1
                If (.Ink < 3) Then
                    .Ink = 6
                End If
                
                '-- Render
                Call BltMask(m_oDIBFore, 8 * .x, 8 * .y, 8, 8, .Paper, .Ink, m_oDIBItem)
            End If
        End With
    Next c
End Sub

Private Sub DoConveyor()
    
  Dim c As Integer
  
    With m_tConveyor
        For c = 0 To .Len - 1
            '-- FX
            Call FXConveyorBlock(m_oDIBBack, 8 * (.x + c), 8 * .y, .Dir)
            Call FXConveyorBlock(m_oDIBMask, 8 * (.x + c), 8 * .y, .Dir)
        Next c
    End With
End Sub

Private Sub DoGuardians()
    
    '-- Horizontals
    Call DoGuardianHs
    
    '-- Verticals
    Select Case m_aRoomID
        Case 4     ' Eugene
            Call DoEugene
        Case 7, 11 ' Kong
            Call DoKong
        Case Else  ' normal guardians
            Call DoGuardianVs
    End Select
End Sub

Private Sub DoGuardianVs()
    
  Dim c As Integer
  
    For c = 1 To UBound(m_tGuardianV())
        
        With m_tGuardianV(c)
            
            Select Case m_aRoomID
            
                Case 13 ' Skylabs
            
                    If (.y < .y2) Then                   ' falling
                        .y = .y + .Speed
                      Else                               ' crashing
                        .Frame = .Frame + 1
                        If (.Frame > 7) Then             ' crashed
                            .Frame = 0                   ' reset frame
                            .y = .y1                     ' reset y
                            .x = .x + 8                  ' 8 blocks right
                            If (.x > 31) Then            ' adjust 8-modulo
                                .x = .x Mod 8
                            End If
                        End If
                    End If
                    
                Case Else ' any other room

                    Select Case True
                        Case .Speed > 0                  ' going down
                            If (.y + .Speed >= .y2) Then ' bottom reached
                                .Speed = -.Speed
                                .y = .y - .Speed
                            End If
                        Case .Speed < 0                  ' going up
                            If (.y + .Speed < .y1) Then  ' top reached
                                .Speed = -.Speed
                                .y = .y - .Speed
                            End If
                    End Select
                    
                    .y = .y + .Speed                     ' advance
                    .Frame = .Frame + 1                  ' frame loop
                    If (.Frame > 3) Then
                        .Frame = 0
                    End If
            End Select
            
            '-- Render
            Call BltMask(m_oDIBFore, 8 * .x, .y, 16, 16, .Paper, .Ink, m_oDIBGuardian(.Frame))
        End With
    Next c
End Sub

Private Sub DoGuardianHs()
    
  Dim c   As Integer
  Dim fr1 As Byte
  Dim fr2 As Byte
  Dim fl1 As Byte
  Dim fl2 As Byte
    
    '-- Define frames' range
    If (m_bHasGuardianVs And m_aRoomID <> 13) Then
        fr1 = 4: fr2 = 7
        fl1 = 4: fl2 = 7
      Else
        fr1 = 0: fr2 = 3
        fl1 = 4: fl2 = 7
    End If
  
    For c = 1 To UBound(m_tGuardianH())
        
        With m_tGuardianH(c)
            
            '-- Normal/slow speed
            .c = .c + .Speed                            ' speed counter
            If (.c Mod 2 = 0) Then                      ' 0.5x or 1.0x speed
                .c = 0
                
                '-- Depending on direction...
                If (.Dir = 0) Then                      ' right
                    If (.x = .x2 And .Frame = fr2) Then ' right extreme reached
                        .Dir = 1                        ' change dir
                        .Frame = fl2                    ' reverse frame counter
                      Else
                        .Frame = .Frame + 1             ' frame loop
                        If (.Frame > fr2) Then
                            .Frame = fr1
                            .x = .x + 1
                        End If
                    End If
                  Else                                  ' left
                    If (.x = .x1 And .Frame = fl1) Then ' left extreme reached
                        .Dir = 0                        ' change dir
                        .Frame = fr1                    ' reverse frame counter
                      Else
                        .Frame = .Frame - 1             ' frame loop
                        If (.Frame < fl1) Then
                            .Frame = fl2
                             .x = .x - 1
                        End If
                    End If
                End If
            End If
            
            '-- Render
            Call BltMask(m_oDIBFore, 8 * .x, 8 * .y, 16, 16, .Paper, .Ink, m_oDIBGuardian(.Frame))
        End With
    Next c
End Sub

Private Sub DoEugene()
    
    With m_tEugene
        
        '-- Move
        If (.Flag = 0) Then                 ' not acticated?
            .y = .y + .Speed                ' advance
            If (.y = .y1 Or .y = .y2) Then  ' reverse direction
                .Speed = -.Speed
                .y = .y + .Speed
            End If
          Else                              ' oh no!
            .y = .y + .Speed                ' go down...
            If (.y = .y2) Then
                .Speed = 0
            End If
            .Ink = .Ink - 1                 ' ink loop
            If (.Ink < 8) Then
                .Ink = 15
            End If
        End If
        
        '-- Render
        Call BltMask(m_oDIBFore, 8 * .x, .y, 16, 16, .Paper, .Ink, .DIB)
    End With
End Sub

Private Sub DoKong()

    With m_tKong
        
        If (.Flag <> 2) Then                ' dead?
            
            .c = .c + 1                     ' frame loop
            If (.c = 8) Then
                .c = 0
                .Frame = 1 - .Frame
            End If
            
            If (.Flag = 1) Then             ' falling?
            
                '-- Sound FX
                Call FXPlay(m_hFXJump, , 22050 - 200 * .y)
                
                .y = .y + 4                 ' go down
                If (.y = .y2) Then          ' bottom reached?
                    .Flag = 2               ' yes: dead
                    Exit Sub
                  Else
                    Call ScoreAdd(100)      ' give me some points
                End If
            End If
            
            '-- Render
            Call MaskBltMask(m_oDIBFore, 8 * .x, .y, 16, 16, .Ink, m_oDIBGuardian(.Frame + 2 * .Flag))
        End If
    End With
End Sub

Private Sub DoSolarRay()
    
  Dim c As Integer
  Dim x As Byte
  Dim y As Byte
  Dim d As Byte
    
    If (m_aRoomID = 18) Then
        
        '-- Starting block and dir
        x = 23
        y = 0
        d = 0
        
        '-- GO!
        Do While m_aRoomBlock(32 * y + x) = BLOCK_AIR
        
            '-- Paint ray block
            Call FXMaskRect(m_oDIBFore, 8 * x, 8 * y, 8, 8, IDX_GREEN, IDX_BRYELLOW, IDX_BRWHITE)
            
            '-- Willy there?
            With m_tWilly
                If ((x = .x Or x = .x + 1) And (y = .y \ 8 Or y = .y \ 8 + 1)) Then
                    Call DoAir(SolarRay:=True)
                End If
            End With
            
            '-- Check for vertical guardians
            For c = 1 To UBound(m_tGuardianV())
                With m_tGuardianV(c)
                    If (.x = x Or .x = x - 1) And _
                       (.y \ 8 = y Or .y \ 8 = y - 1 Or .y \ 8 = y - 2) Then
                        d = 1 - d
                    End If
                End With
            Next c
            
            '-- Check for horizontal guardians
            For c = 1 To UBound(m_tGuardianH())
                With m_tGuardianH(c)
                    If (.x = x Or .x = x - 1) And _
                       (.y = y Or .y = y - 1) Then
                        d = 1 - d
                    End If
                End With
            Next c
            
            '-- Step forward
            If (d) Then
                x = x - 1
              Else
                y = y + 1
            End If
        Loop
    End If
End Sub

Private Sub DoPortal()
    
  Dim i As Byte
    
    With m_tPortal
    
        '-- Flash?
        If (.Flag = 1) Then  ' open
            .c = .c + 1
            If (.c = 4) Then ' swap ink/paper
                .c = 0
                i = .Ink: .Ink = .Paper: .Paper = i
            End If
        End If
        
        '-- Render
        Call BltMask(m_oDIBFore, 8 * .x, 8 * .y, 16, 16, .Paper, .Ink, m_oDIBPortal)
    End With
End Sub

'//

Private Sub DoAir( _
            Optional ByVal SolarRay As Boolean = False _
            )
    
    With m_tAir
        
        If (SolarRay) Then               ' solar ray flag?
            .c = 6                       ' yes: speed up things
        End If
        
        .c = .c + 1                      ' counter
        If (.c = 8) Then                 ' 1 point less every 8 frames
            .c = 0
            If (.Cur) Then
                .Cur = .Cur - 1
              Else
                Call SetMode([eGameDie]) ' sorry: no air!
                Exit Sub
            End If
        End If
        
        '-- Render
        Call FXText(m_oDIBFore, 0, 136, "AIR", m_oDIBChar(), IDX_BRRED, IDX_BRWHITE)
        Call FXRect(m_oDIBFore, 32, 138, .Cur, 4, IDX_BRWHITE)
    End With
End Sub

Private Sub DoPanel()
    
  Dim c As Integer
  
    With m_tPanel
        
        '-- A last chance?
        If (.Lives) Then
            For c = 1 To .Lives - 1
                Call MaskBltMask(m_oDIBFore, c * 16 - 16, 168, 16, 16, IDX_BRCYAN, m_oDIBWilly(.c2))
            Next c
        End If
        
        '-- Cheating!
        If (m_bCheated) Then
            Call MaskBltMask(m_oDIBFore, c * 16 - 16, 168, 16, 16, IDX_BRCYAN, m_oDIBSpecial(2))
        End If
        
        '-- My score is...
        Call FXText(m_oDIBFore, 208, 152, Format$(.SC, "000000"), m_oDIBChar(), IDX_BLACK, IDX_BRYELLOW)
        
        '-- Animate Willys
        .c1 = .c1 + 1
        If (.c1 = 4) Then
            .c1 = 0
            If (m_bTune) Then
                .c2 = .c2 + 1
                If (.c2 > 3) Then
                    .c2 = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub ScoreAdd( _
            ByVal Value As Long _
            )
    
    '-- Add points and check for extra live
    With m_tPanel
        .SC = .SC + Value
        If (.SC \ MM_EXTRALIVE > .Extra) Then
            .Extra = .Extra + 1
            .Lives = .Lives + 1
        End If
    End With
End Sub

'========================================================================================
' Misc
'========================================================================================

'----------------------------------------------------------------------------------------
' Screen
'----------------------------------------------------------------------------------------

Private Sub ScreenFlipBuffer()
    
    '-- Get buffer bits
    Call CopyMemory(ByVal m_oDIBFore.lpBits, ByVal m_oDIBBack.lpBits, m_oDIBBack.Size)
End Sub

Public Sub ScreenUpdate()
    
    If (Not m_oForm Is Nothing) Then
        
        '-- Stretch 2x with 20-pixels border
        Call FXStretch2x(m_oDIBFore2x, m_oDIBFore, m_bFXTV)
        
        '-- Show info
        If (m_bMMInfo) Then
            Call FXText(m_oDIBFore2x, 1, 1, m_sMMInfo, m_oDIBChar(), IDX_WHITE, IDX_BLACK)
        End If
        
        '-- Show FPS
        If (m_bShowFPS) Then
            Call FXText(m_oDIBFore2x, 1, 1 + 9 * -m_bMMInfo, Format$(m_lnFPS, "0000 FPS"), m_oDIBChar(), IDX_WHITE, IDX_BLACK)
        End If
        
        '-- Paint on given DC
        Call m_oDIBFore2x.Paint(m_oForm.hDC, (m_oForm.ScaleWidth - 592) \ 2, (m_oForm.ScaleHeight - 464) \ 2)
    End If
End Sub

'----------------------------------------------------------------------------------------
' Special renderings and FXs
'----------------------------------------------------------------------------------------

Private Sub RenderIntroPart( _
            ByVal ID As Byte _
            )
    
  Dim c As Integer
  
    Select Case ID
        
        Case 0 ' MANIC
            
            For c = 0 To UBound(m_aMANIC()) - 4 Step 5
                Call FXRect(m_oDIBFore, m_aMANIC(c), m_aMANIC(c + 1), m_aMANIC(c + 2), m_aMANIC(c + 3), m_aMANIC(c + 4))
            Next c
        
        Case 1 ' MINER
            
            For c = 0 To UBound(m_aMINER()) - 4 Step 5
                Call FXRect(m_oDIBFore, m_aMINER(c), m_aMINER(c + 1), m_aMINER(c + 2), m_aMINER(c + 3), m_aMINER(c + 4))
            Next c
    End Select
End Sub

Private Sub RenderTitlePart( _
            ByVal ID As Byte _
            )
    
  Dim a1() As Byte
  Dim a2() As Byte
  Dim c    As Integer
  Dim d    As Integer
  Dim e    As Byte
  Dim i    As Byte
  Dim p    As Byte
  Dim s1   As Byte
  Dim s2   As Byte
    
    Select Case ID
    
        Case 0, 1
        
            '-- Load screen pixels (video RAM format)
            Call LoadData(DataFile(DATA_TITLEPX), a1(), 2048 * ID, 2048)
            
            '-- Unpack pixels ('scanlined' and '8-row' interlaced)
            ReDim a2(255)
            s1 = 64 * ID
            s2 = 64 * ID
            For c = 0 To 2047
                For e = 0 To 7
                    If (a1(c) And m_aPow(7 - e)) Then
                        a2(d) = &HFF
                      Else
                        a2(d) = &H0
                    End If
                    d = d + 1
                Next e
                If (d = 256) Then
                    Call CopyMemory(ByVal m_oDIBBack.lpBits + 256 * s1, a2(0), 256)
                    Call CopyMemory(ByVal m_oDIBMask.lpBits + 256 * s1, a2(0), 256)
                    d = 0
                    s1 = s1 + 8
                    If (s1 > 63 + 64 * ID) Then
                        s2 = s2 + 1
                        s1 = s2
                    End If
                End If
            Next c
            
            '-- Get color attributes
            If (ID = 0) Then
                Call LoadData(DataFile(DATA_ROOMS), a2(), 19 * 1024, 1024)
              Else
                Call LoadData(DataFile(DATA_TITLECA), a2(), 0, 256)
            End If
            
            '-- Apply color-attributes
            For c = 0 To 255
                i = CAInk(a2(c))
                p = CAPaper(a2(c))
                Call FXMaskRect(m_oDIBBack, 8 * (c Mod 32), 8 * (c \ 32) + 64 * ID, 8, 8, IDX_MASK, i, p)
            Next c
        
        Case 2
            
            '-- Room name background
            Call FXRect(m_oDIBBack, 0, 128, 256, 8, IDX_YELLOW)
            
            '-- Air supply background
            Call FXRect(m_oDIBBack, 0, 136, 80, 8, IDX_BRRED)
            Call FXRect(m_oDIBBack, 80, 136, 176, 8, IDX_BRGREEN)
            
            '-- Blacken last 6 rows
            Call FXRect(m_oDIBBack, 0, 144, 256, 48, IDX_BLACK)
    End Select
End Sub

'----------------------------------------------------------------------------------------
' Decoding packed graphics
'----------------------------------------------------------------------------------------

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
                 
    Let CAInk = (CA And &H7) + 8 * -((CA And &H40) <> 0)
End Function

Private Function CAPaper( _
                 ByVal CA As Byte _
                 ) As Byte
                 
    Let CAPaper = (CA And &H38) \ &H8 + 8 * -((CA And &H40) <> 0)
End Function

Private Function CAFlash( _
                 ByVal CA As Byte _
                 ) As Byte

    Let CAFlash = CBool(CA And &H80)
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
            m_bMMInfo = Not m_bMMInfo
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
            
            '-- Toggles 1:1 speed and maximum speed
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
            Let KeysCheckAnyKey = True
            Exit For
        End If
    Next c
End Function

Private Sub KeysCheckWillyKeys( _
            KeyCode As Byte _
            )
    
    '-- Right/left
    With m_tWilly
        If ((.Flag And 2) <> 2) Then
            KeyCode = KeyCode Or 1 * -(m_bKey(vbKeyW) Or m_bKey(vbKeyR) Or m_bKey(vbKeyY) Or m_bKey(vbKeyI) Or m_bKey(vbKeyP) Or m_bKey(vbKeyRight))
        End If
        If ((.Flag And 1) <> 1) Then
            KeyCode = KeyCode Or 2 * -(m_bKey(vbKeyQ) Or m_bKey(vbKeyE) Or m_bKey(vbKeyT) Or m_bKey(vbKeyU) Or m_bKey(vbKeyO) Or m_bKey(vbKeyLeft))
        End If
    End With
    
    '-- Jump
    KeyCode = KeyCode Or 4 * -(m_bKey(vbKeyZ) Or m_bKey(vbKeyX) Or m_bKey(vbKeyC) Or m_bKey(vbKeyV) Or m_bKey(vbKeyB) Or m_bKey(vbKeyUp) Or m_bKey(vbKeyN) Or m_bKey(vbKeyM) Or m_bKey(vbKeySpace) Or m_bKey(vbKeyShift) Or m_bKey(226) Or m_bKey(188) Or m_bKey(190) Or m_bKey(189))
End Sub

Private Sub KeysCheckCheatCode( _
            ByVal KeyCode As Integer _
            )

    If (m_bCheated = False) Then
        
        '-- Add char to stream
        m_sCheatCode = Right$(m_sCheatCode, 6) & Chr$(KeyCode)
        
        '-- Correct?
        m_bCheated = (m_sCheatCode = MM_CHEATCODE And m_eMode = [eGamePlay])
    End If
End Sub

Private Sub KeysCheckCheatMask()

  Dim k As Integer
  Dim m As Byte
  
    If (m_bCheated) Then
        
        '-- Key '6' flag?
        If (m_bKey(vbKey6)) Then
            
            '-- Keys '1' to '5': build room # mask
            For k = vbKey1 To vbKey5
                If (m_bKey(k)) Then
                    m = m + m_aPow(k - vbKey1)
                End If
            Next k
            
            '-- Valid?
            If (m < 20) Then
                
                '-- Store new room and load
                If (m_aRoomID <> m) Then
                    m_aRoomID = m
                    Call LoadRoom(m)
                End If
                
                '-- Force room initialization
                Call InitializeRoom
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
        Call FXChangeFreq(m_hChannelTune, 100 * GetNoteFreq(GetMMNote(m_aGameTune(m_lc2)) + 48))
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

Private Function GetMMNote( _
                 ByVal Note As Byte _
                 ) As Byte
  
  Dim c As Integer
  Dim b As Integer
    
    If (Note < 16) Then
        Let GetMMNote = MM_NULLNOTE
      Else
        b = UBound(m_aNoteINV())
        For c = 0 To b
            If (Note <= m_aNoteINV(c)) Then
                Let GetMMNote = b - c
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

Private Sub FXChangeFreq( _
            ByVal hChannel As Long, _
            ByVal freq As Long _
            )
            
    Call FSOUND_SetFrequency(hChannel, freq)
End Sub

Private Sub FXStop( _
            hChannel As Long _
            )
            
    Call FSOUND_StopSound(hChannel)
    hChannel = 0
End Sub

Private Sub FXStopAll()
    
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

Private Sub LoadHiSc()
  
  Dim a() As Byte
    
    '-- Load and unpack
    Call LoadData(AppPath & "MM.dat", a(), 0, 3)
    With m_tPanel
        .HI = a(0) + &H100 * a(1) + &H10000 * a(2)
    End With
End Sub

Private Sub SaveHiSc()

  Dim a(2) As Byte
    
    '-- Pack and save
    With m_tPanel
        a(0) = (.HI And &HFF&)
        a(1) = (.HI And &HFF00&) \ &H100
        a(2) = (.HI And &HFF0000) \ &H10000
    End With
    Call SaveData(AppPath & "MM.dat", a())
End Sub

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
