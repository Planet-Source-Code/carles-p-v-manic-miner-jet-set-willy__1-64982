VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB08.cls (cDIB simplified)
' Author:        Carles P.V.
' Dependencies:  -
' Last revision: 14-Mar-2006
'================================================

Option Explicit

'-- API

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const COLORONCOLOR   As Long = 3

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal length As Long)

'-- Private variables

Private m_tBI     As BITMAPINFO
Private m_aBits() As Byte
Private m_lScan   As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()

    '-- Destroy current DIB, if any
    Call Me.Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create( _
                ByVal Width As Long, _
                ByVal Height As Long _
                ) As Boolean
    
    
    On Error GoTo errH
    
    If (Width > 0 And Height > 0) Then
    
        '-- Scan width
        m_lScan = 4 * ((Width * 8 + 31) \ 32)
        
        With m_tBI.bmiHeader
        
            '-- Define DIB header
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 8
            .biWidth = Width
            .biHeight = -Height
            .biSizeImage = m_lScan * -.biHeight
        
            '-- Initialize DIB data array
            ReDim m_aBits(0 To .biSizeImage - 1)
        End With
        
        '-- Success
        Create = True
    End If
    
errH:
    On Error GoTo 0
End Function

Public Sub Reset()
    
    With m_tBI.bmiHeader
        If (.biSizeImage) Then
            Call ZeroMemory(m_aBits(0), .biSizeImage)
        End If
    End With
End Sub

Public Sub Destroy()

    With m_tBI
    
        If (.bmiHeader.biSizeImage) Then
        
            '-- Reset BI structure
            Call ZeroMemory(.bmiHeader, 40)
            Call ZeroMemory(.bmiColors(0), 1024)
            
            '-- Delete DIB data
            Erase m_aBits()
        End If
    End With
End Sub

Public Sub Cls(ByVal Idx As Byte)
    
  Dim i As Long, i2 As Long
  Dim j As Long, j2 As Long
  Dim o As Long
  
    With m_tBI.bmiHeader
    
        If (.biSizeImage) Then
        
            i2 = .biWidth - 1
            j2 = -.biHeight - 1
            
            For j = 0 To j2
                For i = o To o + i2
                    m_aBits(i) = Idx
                Next i
                o = o + m_lScan
            Next j
        End If
    End With
End Sub

Public Function Paint( _
                ByVal hDC As Long, _
                Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, _
                Optional ByVal ROP As RasterOpConstants = vbSrcCopy _
                ) As Long
    
    With m_tBI.bmiHeader
        Paint = Me.Stretch(hDC, x, y, .biWidth, -.biHeight, , , , , ROP)
    End With
End Function

Public Function Stretch( _
                ByVal hDC As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal Width As Long, ByVal Height As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal WidthSrc As Long, Optional ByVal HeightSrc As Long, _
                Optional ByVal ROP As RasterOpConstants = vbSrcCopy _
                ) As Long
  
  Dim lOldMode As Long
    
    With m_tBI.bmiHeader
    
        If (.biSizeImage) Then
            
            If (WidthSrc = 0) Then
                WidthSrc = .biWidth
            End If
            If (HeightSrc = 0) Then
                HeightSrc = -.biHeight
            End If
            
            lOldMode = SetStretchBltMode(hDC, COLORONCOLOR)
            Stretch = StretchDIBits(hDC, x, y, Width, Height, xSrc, ySrc, WidthSrc, HeightSrc, m_aBits(0), m_tBI, DIB_RGB_COLORS, ROP)
            Call SetStretchBltMode(hDC, lOldMode)
        End If
    End With
End Function

'//

Public Sub SetPalette(Palette() As Byte)

    With m_tBI
        If (.bmiHeader.biSizeImage) Then
            Call CopyMemory(.bmiColors(0), Palette(0), 1024)
        End If
    End With
    
'-- Note: palette array must be initialized before
End Sub

Public Sub GetPalette(Palette() As Byte)
    
    With m_tBI
        If (.bmiHeader.biSizeImage) Then
            Call CopyMemory(Palette(0), .bmiColors(0), 1024)
        End If
    End With
    
'-- Note: palette array must be initialized before
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get HasDIB() As Boolean
    HasDIB = (m_tBI.bmiHeader.biSizeImage > 0)
End Property

Public Property Get lpBits() As Long
    lpBits = VarPtr(m_aBits(0))
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_tBI.bmiHeader.biHeight
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = m_lScan
End Property

Public Property Get Size() As Long
    Size = m_tBI.bmiHeader.biSizeImage
End Property

'========================================================================================
' Helpers
'========================================================================================

Public Function GetPixel(ByVal x As Long, ByVal y As Long) As Byte
    
    With m_tBI.bmiHeader
        If (.biSizeImage) Then
            If (x >= 0 And y >= 0 And x < .biWidth And y < -.biHeight) Then
                GetPixel = m_aBits(y * m_lScan + x)
            End If
        End If
    End With
End Function
