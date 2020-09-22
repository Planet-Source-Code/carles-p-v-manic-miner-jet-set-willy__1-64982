Attribute VB_Name = "mRender"
Option Explicit

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds     As SAFEARRAYBOUND
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long



'========================================================================================
' Methods
'========================================================================================

Public Sub BltFast( _
           DIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           DIBSrc As cDIB08, _
           Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0 _
           )

  Dim b()  As Byte
  Dim a()  As Byte
  Dim SA1b As SAFEARRAY1D
  Dim SA1a As SAFEARRAY1D
  Dim rb   As RECT2
  Dim ra   As RECT2
  Dim sb   As Long
  Dim sa   As Long
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long
    
    Call SetRect(rb, 0, 0, DIB.Width, DIB.Height)
    Call SetRect(ra, x, y, x + Width, y + Height)
    Call IntersectRect(rb, rb, ra)
    
    If (IsRectEmpty(rb) = 0) Then
        
        Call OffsetRect(rb, -x, -y)
        
        With rb
            i1 = .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        sb = DIB.BytesPerScanline
        k1 = (i1 + x) + (rb.y1 - 0 + y) * sb
        sa = DIBSrc.BytesPerScanline
        i1 = (i1 + xSrc) + (j1 + ySrc) * sa
        
        Call MapDIBitsByte(SA1b, b(), DIB)
        Call MapDIBitsByte(SA1a, a(), DIBSrc)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                b(i + k2) = a(i)
            Next i
            i1 = i1 + sa
            k1 = k1 + sb
        Next j
        
        Call UnmapDIBitsByte(b())
        Call UnmapDIBitsByte(a())
    End If
End Sub

Public Sub BltMask( _
           DIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           ByVal IdxBack As Byte, _
           ByVal IdxFore As Byte, _
           DIBSrc As cDIB08, _
           Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0 _
           )

  Dim b()  As Byte
  Dim a()  As Byte
  Dim SA1b As SAFEARRAY1D
  Dim SA1a As SAFEARRAY1D
  Dim rb   As RECT2
  Dim ra   As RECT2
  Dim sb   As Long
  Dim sa   As Long
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long
    
    Call SetRect(rb, 0, 0, DIB.Width, DIB.Height)
    Call SetRect(ra, x, y, x + Width, y + Height)
    Call IntersectRect(rb, rb, ra)
    
    If (IsRectEmpty(rb) = 0) Then
        
        Call OffsetRect(rb, -x, -y)
        
        With rb
            i1 = .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        sb = DIB.BytesPerScanline
        k1 = (i1 + x) + (rb.y1 - 0 + y) * sb
        sa = DIBSrc.BytesPerScanline
        i1 = (i1 + xSrc) + (j1 + ySrc) * sa
        
        Call MapDIBitsByte(SA1b, b(), DIB)
        Call MapDIBitsByte(SA1a, a(), DIBSrc)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (a(i)) Then
                    b(i + k2) = IdxFore
                  Else
                    b(i + k2) = IdxBack
                End If
            Next i
            i1 = i1 + sa
            k1 = k1 + sb
        Next j
        
        Call UnmapDIBitsByte(b())
        Call UnmapDIBitsByte(a())
    End If
End Sub

Public Sub MaskBltMask( _
           DIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           ByVal Idx As Byte, _
           DIBSrc As cDIB08, _
           Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
           Optional ByVal IdxSrcMask As Byte = 0 _
           )

  Dim b()  As Byte
  Dim a()  As Byte
  Dim SA1b As SAFEARRAY1D
  Dim SA1a As SAFEARRAY1D
  Dim rb   As RECT2
  Dim ra   As RECT2
  Dim sb   As Long
  Dim sa   As Long
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long
    
    Call SetRect(rb, 0, 0, DIB.Width, DIB.Height)
    Call SetRect(ra, x, y, x + Width, y + Height)
    Call IntersectRect(rb, rb, ra)
    
    If (IsRectEmpty(rb) = 0) Then
        
        Call OffsetRect(rb, -x, -y)
        
        With rb
            i1 = .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        sb = DIB.BytesPerScanline
        k1 = (i1 + x) + (rb.y1 - 0 + y) * sb
        sa = DIBSrc.BytesPerScanline
        i1 = (i1 + xSrc) + (j1 + ySrc) * sa
        
        Call MapDIBitsByte(SA1b, b(), DIB)
        Call MapDIBitsByte(SA1a, a(), DIBSrc)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (a(i) <> IdxSrcMask) Then
                    b(i + k2) = Idx
                End If
            Next i
            i1 = i1 + sa
            k1 = k1 + sb
        Next j
        
        Call UnmapDIBitsByte(b())
        Call UnmapDIBitsByte(a())
    End If
End Sub

'//

Public Sub FXConveyorBlock( _
           DIB As cDIB08, _
           ByVal x As Long, _
           ByVal y As Long, _
           Optional ByVal Dir As Byte = 0 _
           )

  Dim a() As Byte
  Dim SA1 As SAFEARRAY1D
  Dim s   As Long
  
  Dim i  As Long
  Dim i1 As Long
  Dim i2 As Long
  Dim t1 As Byte
  Dim t2 As Byte
  
    s = DIB.BytesPerScanline
    
    If (y * s + x >= 0 And (y + 2) * s + x + 7 < DIB.Size) Then

        Call MapDIBitsByte(SA1, a(), DIB)
        
        i1 = (y + 0) * s + x
        i2 = i1 + 7
        
        If (Dir = 0) Then
            t1 = a(i1 + 0)
            t2 = a(i1 + 1)
            For i = i1 To i2 - 2 Step 1
                a(i) = a(i + 2)
            Next i
            a(i2 - 1) = t1
            a(i2 - 0) = t2
          Else
            t1 = a(i2 - 1)
            t2 = a(i2 - 0)
            For i = i2 To i1 + 2 Step -1
                a(i) = a(i - 2)
            Next i
            a(i1 + 0) = t1
            a(i1 + 1) = t2
        End If
        
        i1 = (y + 2) * s + x
        i2 = i1 + 7

        If (Dir = 0) Then
            t1 = a(i2 - 1)
            t2 = a(i2 - 0)
            For i = i2 To i1 + 2 Step -1
                a(i) = a(i - 2)
            Next i
            a(i1 + 0) = t1
            a(i1 + 1) = t2
          Else
            t1 = a(i1 + 0)
            t2 = a(i1 + 1)
            For i = i1 To i2 - 2 Step 1
                a(i) = a(i + 2)
            Next i
            a(i2 - 1) = t1
            a(i2 - 0) = t2
        End If
        
        Call UnmapDIBitsByte(a())
    End If
End Sub

'//

Public Sub FXRect( _
           DIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           ByVal Idx As Byte _
           )

  Dim a() As Byte
  Dim SA1 As SAFEARRAY1D
  Dim s   As Long
  
  Dim r1 As RECT2
  Dim r2 As RECT2
  
  Dim i As Long, i1 As Long, j1 As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    Call SetRect(r1, 0, 0, DIB.Width, DIB.Height)
    Call SetRect(r2, x, y, x + Width, y + Height)
    Call IntersectRect(r1, r1, r2)
    
    If (IsRectEmpty(r1) = 0) Then
        
        s = DIB.BytesPerScanline
        
        With r1
            i1 = .y1 * s + .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        Call MapDIBitsByte(SA1, a(), DIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                a(i) = Idx
            Next i
            i1 = i1 + s
        Next j
        
        Call UnmapDIBitsByte(a())
    End If
End Sub

Public Sub FXMaskRect( _
           DIB As cDIB08, _
           ByVal x As Long, ByVal y As Long, _
           ByVal Width As Long, ByVal Height As Long, _
           ByVal IdxMask As Byte, _
           ByVal IdxBack As Byte, _
           ByVal IdxFore As Byte _
           )

  Dim a() As Byte
  Dim SA1 As SAFEARRAY1D
  Dim s   As Long
  
  Dim r1 As RECT2
  Dim r2 As RECT2
  
  Dim i As Long, i1 As Long, j1 As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    Call SetRect(r1, 0, 0, DIB.Width, DIB.Height)
    Call SetRect(r2, x, y, x + Width, y + Height)
    Call IntersectRect(r1, r1, r2)
    
    If (IsRectEmpty(r1) = 0) Then
        
        s = DIB.BytesPerScanline
        
        With r1
            i1 = .y1 * s + .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        Call MapDIBitsByte(SA1, a(), DIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                If (a(i) = IdxMask) Then
                    a(i) = IdxBack
                   Else
                    a(i) = IdxFore
                End If
            Next i
            i1 = i1 + s
        Next j
        
        Call UnmapDIBitsByte(a())
    End If
End Sub

Public Sub FXText( _
           DIB As cDIB08, _
           ByVal x As Long, _
           ByVal y As Long, _
           ByVal Text As String, _
           DIBchar() As cDIB08, _
           ByVal IdxBack As Byte, _
           ByVal IdxFore As Byte _
           )

  Dim c As Long
  Dim a As Byte

    For c = 1 To Len(Text)
        a = Asc(Mid$(Text, c, 1))
        Call BltMask(DIB, x + 8 * (c - 1), y, 8, 8, IdxBack, IdxFore, DIBchar(a - 32))
    Next c
End Sub

Public Function FXImageCollide( _
                ByVal DIB1 As cDIB08, ByVal x As Long, ByVal y As Long, _
                ByVal DIB2 As cDIB08 _
                ) As Boolean
                            
  Dim b()  As Byte
  Dim a()  As Byte
  Dim SA1b As SAFEARRAY1D
  Dim SA1a As SAFEARRAY1D
  Dim rb   As RECT2
  Dim ra   As RECT2
  Dim sb   As Long
  Dim sa   As Long
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long
    
    Call SetRect(rb, 0, 0, DIB1.Width, DIB1.Height)
    Call SetRect(ra, x, y, x + DIB2.Width, y + DIB2.Height)
    Call IntersectRect(rb, rb, ra)
    
    If (IsRectEmpty(rb) = 0) Then
        
        Call OffsetRect(rb, -x, -y)
        
        With rb
            i1 = .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        sb = DIB1.BytesPerScanline
        k1 = (i1 + x) + (rb.y1 - 0 + y) * sb
        sa = DIB2.BytesPerScanline
        i1 = (i1 + 0) + (j1 + 0) * sa
        
        Call MapDIBitsByte(SA1b, b(), DIB1)
        Call MapDIBitsByte(SA1a, a(), DIB2)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If (b(i + k2) And a(i)) Then
                    FXImageCollide = True
                    GoTo skip
                End If
            Next i
            i1 = i1 + sa
            k1 = k1 + sb
        Next j
        
skip:   Call UnmapDIBitsByte(b())
        Call UnmapDIBitsByte(a())
    End If
End Function

'//

Public Sub FXShift( _
           DIB As cDIB08, _
           Optional ByVal Inc As Byte = 1 _
           )

  Dim a() As Byte
  Dim SA1 As SAFEARRAY1D
  
  Dim i As Long
    
    Call MapDIBitsByte(SA1, a(), DIB)
    
    For i = 0 To DIB.Size - 1
        a(i) = (a(i) + Inc) And &H7
    Next i
    
    Call UnmapDIBitsByte(a())
End Sub

Public Sub FXBorder( _
           DIB As cDIB08, _
           ByVal Idx As Byte, _
           Optional ByVal KeepIdx As Boolean = False, _
           Optional ByVal FXTV As Boolean = False _
           ) ' be careful: no checks!

           
  Dim a() As Byte
  Dim SA1 As SAFEARRAY1D
  Dim s   As Long
  
  Dim i As Long, i1 As Long, j1 As Long
  Dim j As Long, i2 As Long, j2 As Long
        
    s = DIB.BytesPerScanline
    
    Call MapDIBitsByte(SA1, a(), DIB)
    
    If (KeepIdx) Then
        Idx = a(0)
    End If
    
    j1 = 0: j2 = 39
    i1 = j1 * s: i2 = s - 1
    For j = j1 To j2
        If (FXTV And j Mod 2) Then
            For i = i1 To i1 + i2
                a(i) = Idx + 16
            Next i
          Else
            For i = i1 To i1 + i2
                a(i) = Idx
            Next i
        End If
        i1 = i1 + s
    Next j

    j1 = 424: j2 = 463
    i1 = j1 * s: i2 = s - 1
    For j = j1 To j2
        If (FXTV And j Mod 2) Then
            For i = i1 To i1 + i2
                a(i) = Idx + 16
            Next i
          Else
            For i = i1 To i1 + i2
                a(i) = Idx
            Next i
        End If
        i1 = i1 + s
    Next j
    
    j1 = 40: j2 = 443
    i1 = j1 * s + 0: i2 = 39
    For j = j1 To j2
        If (FXTV And j Mod 2) Then
            For i = i1 To i1 + i2
                a(i) = Idx + 16
            Next i
          Else
            For i = i1 To i1 + i2
                a(i) = Idx
            Next i
        End If
        i1 = i1 + s
    Next j
    
    j1 = 40: j2 = 443
    i1 = j1 * s + 552: i2 = 39
    For j = j1 To j2
        If (FXTV And j Mod 2) Then
            For i = i1 To i1 + i2
                a(i) = Idx + 16
            Next i
          Else
            For i = i1 To i1 + i2
                a(i) = Idx
            Next i
        End If
        i1 = i1 + s
    Next j
    
    Call UnmapDIBitsByte(a())
End Sub

Public Sub FXStretch2x( _
           DIB As cDIB08, _
           DIBSrc As cDIB08, _
           Optional ByVal FXTV As Boolean = False _
           ) ' be careful: no checks!

  Dim b()  As Byte
  Dim a()  As Byte
  Dim SA1b As SAFEARRAY1D
  Dim SA1a As SAFEARRAY1D
  Dim sb   As Long
  Dim sa   As Long

  Dim xLU() As Long
  Dim yLU() As Long

  Dim i  As Long, w  As Long
  Dim j  As Long, h  As Long
  Dim po As Long, pn As Long, qn As Long

    sb = DIB.BytesPerScanline
    sa = DIBSrc.BytesPerScanline

    w = DIB.Width - 80 - 1
    h = DIB.Height - 80 - 1

    ReDim xLU(w)
    For i = 0 To w
        xLU(i) = (i \ 2)
    Next i
    ReDim yLU(h)
    For i = 0 To h
        yLU(i) = (i \ 2) * sa
    Next i

    Call MapDIBitsByte(SA1b, b(), DIB)
    Call MapDIBitsByte(SA1a, a(), DIBSrc)
    
    pn = 40 * sb + 40
    For j = 0 To h
        po = yLU(j)
        qn = pn
        If (FXTV And j Mod 2) Then
            For i = 0 To w
                b(qn) = a(po + xLU(i)) + 16
                qn = qn + 1
            Next i
          Else
            For i = 0 To w
                b(qn) = a(po + xLU(i))
                qn = qn + 1
            Next i
        End If
        pn = pn + sb
    Next j

    Call UnmapDIBitsByte(b())
    Call UnmapDIBitsByte(a())
End Sub

'//

Private Sub MapDIBitsByte( _
            SA1 As SAFEARRAY1D, _
            a() As Byte, _
            DIB As cDIB08 _
            )

    With SA1
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = DIB.Size
        .pvData = DIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(a()), VarPtr(SA1), 4)
End Sub

Private Sub UnmapDIBitsByte( _
            a() As Byte _
            )
    
    Call CopyMemory(ByVal VarPtrArray(a()), 0&, 4)
End Sub
