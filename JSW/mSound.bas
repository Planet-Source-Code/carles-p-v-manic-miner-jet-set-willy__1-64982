Attribute VB_Name = "mSound"
Option Explicit

Public Enum FSOUND_INITMODES
    FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1       ' causes MIDI playback to force software decoding.
    FSOUND_INIT_GLOBALFOCUS = &H2               ' for DirectSound output - sound is not muted when window is out of focus.
    FSOUND_INIT_ENABLESYSTEMCHANNELFX = &H4     ' for DirectSound output - Allows FSOUND_FX api to be used on global software mixer output!
    FSOUND_INIT_ACCURATEVULEVELS = &H8          ' this latency adjusts FSOUND_GetCurrentLevels, but incurs a small cpu and memory hit
    FSOUND_INIT_PS2_DISABLECORE0REVERB = &H10   ' PS2 only - Disable reverb on CORE 0 to regain SRAM
    FSOUND_INIT_PS2_DISABLECORE1REVERB = &H20   ' PS2 only - Disable reverb on CORE 1 to regain SRAM
    FSOUND_INIT_PS2_SWAPDMACORES = &H40         ' PS2 only - By default FMOD uses DMA CH0 for mixing, CH1 for uploads, this flag swaps them around
    FSOUND_INIT_DONTLATENCYADJUST = &H80        ' callbacks are not latency adjusted, and are called at mix time.  Also information functions are immediate
    FSOUND_INIT_GC_INITLIBS = &H100             ' Gamecube only - Initializes GC audio libraries
    FSOUND_INIT_STREAM_FROM_MAIN_THREAD = &H200 ' turns off fmod streamer thread, and makes streaming update from FSOUND_Update called by the user
    FSOUND_INIT_PS2_USEVOLUMERAMPING = &H400    ' PS2 only   - Turns on volume ramping system to remove hardware clicks.
    FSOUND_INIT_DSOUND_DEFERRED = &H800         ' Win32 only - For DirectSound output.  3D commands are batched together and executed at FSOUND_Update.
    FSOUND_INIT_DSOUND_HRTF_LIGHT = &H1000      ' Win32 only - For DirectSound output.  FSOUND_HW3D buffers use a slightly higher quality algorithm when 3d hardware acceleration is not present.
    FSOUND_INIT_DSOUND_HRTF_FULL = &H2000       ' Win32 only - For DirectSound output.  FSOUND_HW3D buffers use full quality 3d playback when 3d hardware acceleration is not present.
    FSOUND_INIT_XBOX_REMOVEHEADROOM = &H4000    ' XBox only - By default directsound attenuates all sound by 6db to avoid clipping/distortion.  CAUTION.  If you use this flag you are responsible for the final mix to make sure clipping / distortion doesn't happen.
    FSOUND_INIT_PSP_SILENCEONUNDERRUN = &H8000  ' PSP only - If streams skip / stutter when device is powered on, either increase stream buffersize, or use this flag instead to play silence while the UMD is recovering.
End Enum

Public Enum FSOUND_MODES
    FSOUND_LOOP_OFF = &H1                       ' for non looping samples.
    FSOUND_LOOP_NORMAL = &H2                    ' for forward looping samples.
    FSOUND_LOOP_BIDI = &H4                      ' for bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = &H8                          ' for 8 bit samples.
    FSOUND_16BITS = &H10                        ' for 16 bit samples.
    FSOUND_MONO = &H20                          ' for mono samples.
    FSOUND_STEREO = &H40                        ' for stereo samples.
    FSOUND_UNSIGNED = &H80                      ' for source data containing unsigned samples.
    FSOUND_SIGNED = &H100                       ' for source data containing signed data.
    FSOUND_DELTA = &H200                        ' for source data stored as delta values.
    FSOUND_IT214 = &H400                        ' for source data stored using IT214 compression.
    FSOUND_IT215 = &H800                        ' for source data stored using IT215 compression.
    FSOUND_HW3D = &H1000                        ' attempts to make samples use 3d hardware acceleration. (if the card supports it)
    FSOUND_2D = &H2000                          ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = &H4000                  ' for realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = &H8000                  ' for FSOUND_Sample_Load - name will be interpreted as a pointer to data
    FSOUND_LOADRAW = &H10000                    ' for FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = &H20000               ' for FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = &H40000                  ' for forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = &H80000                       ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = &H100000                  ' allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency
    FSOUND_MPEGHALFRATE = &H200000              ' for FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = &H400000                    ' for XBOX only - Describes a user sample that its contents are compressed as XADPCM
    FSOUND_VAG = &H800000                       ' for PS2 only - Describes a user sample that its contents are compressed as Sony VAG format.
    FSOUND_NONBLOCKING = &H1000000              ' for FSOUND_Stream_Open - Causes stream to open in the background and not block the foreground app - stream plays only when ready.
    FSOUND_GCADPCM = &H2000000                  ' for Gamecube only - Contents are compressed as Gamecube DSP-ADPCM format
    FSOUND_MULTICHANNEL = &H4000000             ' for PS2 only - Contents are interleaved into a multi-channel (more than stereo) format
    FSOUND_USECORE0 = &H8000000                 ' for PS2 only - Sample/Stream is forced to use hardware voices 00-23
    FSOUND_USECORE1 = &H10000000                ' for PS2 only - Sample/Stream is forced to use hardware voices 24-47
    FSOUND_LOADMEMORYIOP = &H20000000           ' for PS2 only - "name" will be interpreted as a pointer to data for streaming and samples.  The address provided will be an IOP address
    FSOUND_IGNORETAGS = &H40000000              ' dkips id3v2 etc tag checks when opening a stream, to reduce seek/read overhead when opening files (helps with CD performance)
    FSOUND_STREAM_NET = &H80000000              ' specifies an internet stream

    FSOUND_NORMAL = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum

Public Enum FSOUND_CHANNELSAMPLEMODE
    FSOUND_FREE = -1                            ' definition for dynamically allocated channel or sample
    FSOUND_UNMANAGED = -2                       ' definition for allocating a sample that is NOT managed by fsound
    FSOUND_ALL = -3                             ' for a channel index or sample index, this flag affects ALL channels or samples available!  Not supported by all functions.
    FSOUND_STEREOPAN = -1                       ' definition for full middle stereo volume on both channels
    FSOUND_SYSTEMCHANNEL = -1000                ' special channel ID for channel based functions that want to alter the global FSOUND software mixing output channel
    FSOUND_SYSTEMSAMPLE = -1000                 ' special sample ID for all sample based functions that want to alter the global FSOUND software mixing output sample
End Enum

Public Declare Function FSOUND_Init Lib "fmod" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal flags As FSOUND_INITMODES) As Byte
Public Declare Function FSOUND_Close Lib "fmod" Alias "_FSOUND_Close@0" () As Long

Public Declare Function FSOUND_Sample_Load Lib "fmod" Alias "_FSOUND_Sample_Load@20" (ByVal index As Long, ByVal name As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Public Declare Function FSOUND_Sample_Free Lib "fmod" Alias "_FSOUND_Sample_Free@4" (ByVal sptr As Long) As Long

Public Declare Function FSOUND_PlaySound Lib "fmod" Alias "_FSOUND_PlaySound@8" (ByVal channel As Long, ByVal sptr As Long) As Long
Public Declare Function FSOUND_StopSound Lib "fmod" Alias "_FSOUND_StopSound@4" (ByVal channel As Long) As Byte

Public Declare Function FSOUND_SetFrequency Lib "fmod" Alias "_FSOUND_SetFrequency@8" (ByVal channel As Long, ByVal freq As Long) As Byte
Public Declare Function FSOUND_SetVolume Lib "fmod" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Public Declare Function FSOUND_SetLoopMode Lib "fmod" Alias "_FSOUND_SetLoopMode@8" (ByVal channel As Long, ByVal loopmode As Long) As Byte
