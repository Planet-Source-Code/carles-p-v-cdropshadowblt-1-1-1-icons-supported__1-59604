VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDropShadowBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' Class:         cDropShadowBlt.cls
' Author:        Carles P.V. - 2005
' Dependencies:
' Last revision: 2005.03.26
' Version:       1.1.1
'----------------------------------------------------------------------------------------
'
' History:
'
'     1.0.0: First release.
'     1.0.1: Added 'MaxAlpha' param. to Paint() sub.
'     1.0.2: - Improved pvRenderDropShadow() routine: DIB bits maped (no buffers).
'            - Added 'ShadowColor' option.
'     1.0.3: - Restricted minimum offsets (1) and added 1 pixel edge for blurring cell:
'              avoided expensive PtInRect(). Now it's time to use 1D byte arrays instead
'              of RGBQUAD type for more speed.
'     1.1.0: - Icon source supported: CreateFromStdPicture() -> CreateFromHandle().
'     1.1.1: - Fixed shadow for left and top edges.
'========================================================================================

Option Explicit

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

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

Private Type ICONINFO
    fIcon    As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask  As Long
    hbmColor As Long
End Type

Private Type BLENDFUNCTION
    BlendOp             As Byte
    BlendFlags          As Byte
    SourceConstantAlpha As Byte
    AlphaFormat         As Byte
End Type

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Const AC_SRC_OVER    As Long = 0
Private Const AC_SRC_ALPHA   As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hDcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long



'========================================================================================
' cDropShadowBlt
'========================================================================================

Private Const BLUR_EXTENT As Long = 1
Private Const OFFSET_MIN  As Long = 1

Private m_uBIH            As BITMAPINFOHEADER
Private m_hDC             As Long
Private m_hDIB            As Long
Private m_hDIBOld         As Long
Private m_lpBits          As Long

Private Sub Class_Terminate()
    '-- Destroy current DIB
    Call pvDestroy
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function CreateFromHandle( _
                ByVal hImage As Long, _
                Optional ByVal MaskColor As OLE_COLOR = vbMagenta, _
                Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
                Optional ByVal Opacity As Long = 75, _
                Optional ByVal xOffset As Long = 1, _
                Optional ByVal yOffset As Long = 1 _
                ) As Boolean

  Dim bSuccess As Boolean
  Dim lhDC     As Long
  Dim lhBmpOld As Long
  Dim uBI      As BITMAP
  Dim uII      As ICONINFO
  Dim bIsIcon  As Boolean
 
    '-- Check minimums
    If (xOffset < OFFSET_MIN) Then xOffset = OFFSET_MIN
    If (yOffset < OFFSET_MIN) Then yOffset = OFFSET_MIN
    
    If (hImage) Then
        
        '-- Check object type
        If (GetObjectType(hImage) = OBJ_BITMAP) Then
            '-- Get bitmap info
            Call GetObject(hImage, LenB(uBI), uBI)
            '-- Success
            bSuccess = True
          Else
            '-- Icon?
            If (GetIconInfo(hImage, uII)) Then
                '-- Get icon color-bitmap info
                Call GetObject(uII.hbmColor, LenB(uBI), uBI)
                hImage = uII.hbmColor
                '-- Success
                bSuccess = True
                bIsIcon = True
            End If
        End If
        
        If (bSuccess) Then
            
            '-- Create 32-bit DIB section
            Call pvCreateDIBSection32(uBI.bmWidth + xOffset + BLUR_EXTENT, uBI.bmHeight + yOffset + BLUR_EXTENT)
            
            '-- Fill background using MaskColor (extra shadow area: masking purposes)
            If (bIsIcon = False) Then
                Call pvCls(MaskColor)
            End If
            
            '-- BitBlt image to our DIB DC
            lhDC = CreateCompatibleDC(m_hDC)
            lhBmpOld = SelectObject(lhDC, hImage)
            Call BitBlt(m_hDC, 0, 0, uBI.bmWidth, uBI.bmHeight, lhDC, 0, 0, vbSrcCopy)
            Call SelectObject(lhDC, lhBmpOld)
            Call DeleteDC(lhDC)
            
            '-- Mask DIB
            If (bIsIcon) Then
                Call pvMaskDIBFromMask(uII.hbmMask)
                Call DeleteObject(uII.hbmColor)
                Call DeleteObject(uII.hbmMask)
              Else
                Call pvMaskDIBFromColor(MaskColor)
            End If
            
            '-- Render drop shadow
            Call pvRenderDropShadow(ShadowColor, Opacity, xOffset, yOffset)
        End If
    End If
    
    CreateFromHandle = bSuccess
End Function

Public Function CreateFromDC( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                Optional ByVal MaskColor As OLE_COLOR = vbMagenta, _
                Optional ByVal ShadowColor As OLE_COLOR = vbBlack, _
                Optional ByVal Opacity As Long = 75, _
                Optional ByVal xOffset As Long = 1, _
                Optional ByVal yOffset As Long = 1 _
                ) As Boolean

    '-- Check minimums
    If (Width < 1) Then Width = 1
    If (Height < 1) Then Height = 1
    If (xOffset < OFFSET_MIN) Then xOffset = OFFSET_MIN
    If (yOffset < OFFSET_MIN) Then yOffset = OFFSET_MIN
    
    '-- Create DIB
    If (pvCreateDIBSection32(Width + xOffset + BLUR_EXTENT, Height + yOffset + BLUR_EXTENT)) Then
            
        '-- Clear background using MaskColor (extra shadow area: masking purposes)
        Call pvCls(MaskColor)
        '-- BitBlt image to our DIB DC
        Call BitBlt(m_hDC, 0, 0, Width, Height, hDC, x, y, vbSrcCopy)
            
        '-- Success: mask bitmap and render 'drop shadow'
        Call pvMaskDIBFromColor(MaskColor)
        Call pvRenderDropShadow(ShadowColor, Opacity, xOffset, yOffset)
        CreateFromDC = True
    End If
End Function

Public Sub Paint( _
           ByVal hDC As Long, _
           ByVal x As Long, _
           ByVal y As Long, _
           Optional ByVal MaxAlpha As Byte = &HFF _
           )

  Dim uBFN  As BLENDFUNCTION
  Dim lpBFN As Long
    
    With uBFN
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .AlphaFormat = AC_SRC_ALPHA
        .SourceConstantAlpha = MaxAlpha
    End With
    Call CopyMemory(lpBFN, uBFN, LenB(uBFN))
    
    '-- Paint AlphaBlend
    With m_uBIH
        Call AlphaBlend(hDC, x, y, .biWidth, .biHeight, m_hDC, 0, 0, .biWidth, .biHeight, lpBFN)
    End With
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvDestroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hDIBOld)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, LenB(m_uBIH))
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hDIBOld = 0
    m_lpBits = 0
End Sub

Private Sub pvCls( _
            ByVal lColor As OLE_COLOR _
            )
    
  Dim hBrush As Long
  Dim uRect  As RECT2
  
    Call OleTranslateColor(lColor, 0, lColor)
    hBrush = CreateSolidBrush(lColor)
    Call SetRect(uRect, 0, 0, m_uBIH.biWidth, m_uBIH.biHeight)
    Call FillRect(m_hDC, uRect, hBrush)
    Call DeleteObject(hBrush)
End Sub

Private Function pvCreateDIBSection32( _
                 ByVal lWidth As Long, _
                 ByVal lHeight As Long _
                 ) As Boolean
    
    '-- Destroy previous
    Call pvDestroy
    
    '-- Define DIB header
    With m_uBIH
        .biSize = LenB(m_uBIH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = lHeight
        .biSizeImage = (4 * .biWidth) * .biHeight
    End With
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        m_hDIB = CreateDIBSection(m_hDC, m_uBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (m_hDIB <> 0) Then
            m_hDIBOld = SelectObject(m_hDC, m_hDIB)
          Else
            Call pvDestroy
        End If
    End If
    
    '-- Success
    pvCreateDIBSection32 = (m_hDIB <> 0)
End Function

Private Sub pvMaskDIBFromColor( _
            ByVal MaskColor As OLE_COLOR _
            )

  Dim lColor   As Long
  Dim R        As Long
  Dim G        As Long
  Dim B        As Long
  Dim aBits()  As RGBQUAD
  Dim uSA      As SAFEARRAY2D
  Dim x        As Long, xEnd As Long
  Dim y        As Long, yEnd As Long

    '-- Map DIB bits

    Call pvMapDIBits(uSA, aBits())

    '-- 'Mask' transparent color ('null' pixel)

    Call OleTranslateColor(MaskColor, 0, lColor)
    R = (lColor And &HFF&)
    G = (lColor And &HFF00&) \ &H100
    B = (lColor And &HFF0000) \ &H10000

    With m_uBIH
        yEnd = .biHeight - 1
        xEnd = .biWidth - 1
    End With

    For y = 0 To yEnd
        For x = 0 To xEnd
            With aBits(x, y)
                If (.R <> R Or .G <> G Or .B <> B) Then
                    .A = 255
                  Else
                    .R = 0
                    .G = 0
                    .B = 0
                End If
            End With
        Next x
    Next y

    '-- Unmap DIB bits

    Call pvUnmapDIBits(aBits())
End Sub

Private Sub pvMaskDIBFromMask( _
            ByVal hBitmapMask As Long _
            )
            
  Dim lhDC1     As Long
  Dim lhBmp1Old As Long
  Dim lhDC2     As Long
  Dim lhBmp2    As Long
  Dim lhBmp2Old As Long
  Dim aBitsM()  As Long
  Dim aBitsC()  As RGBQUAD
  Dim uSAC      As SAFEARRAY2D
  Dim x         As Long, xEnd As Long
  Dim y         As Long, yEnd As Long
    
    With m_uBIH
        
        '-- Select image into a DC
        lhDC1 = CreateCompatibleDC(0)
        lhBmp1Old = SelectObject(lhDC1, hBitmapMask)
        
        '-- Create temporary bitmap
        '   This is done for:
        '   - Take in account DIB final size (shadow offset and blur extent)
        '   - Invert (vbNotSrcCopy) mask for processing added area as 'not-mask')
        lhDC2 = CreateCompatibleDC(0)
        lhBmp2 = CreateCompatibleBitmap(lhDC2, .biWidth, .biHeight)
        lhBmp2Old = SelectObject(lhDC2, lhBmp2)
        
        '-- BitBlt 'inverted' mask to the temporary DC
        Call BitBlt(lhDC2, 0, 0, .biWidth, .biHeight, lhDC1, 0, 0, vbNotSrcCopy)
        
        '-- Initialize color array and extract color info
        ReDim aBitsM(.biWidth - 1, .biHeight - 1)
        Call GetDIBits(lhDC2, lhBmp2, 0, .biHeight, aBitsM(0, 0), m_uBIH, DIB_RGB_COLORS)
    End With
    
    '-- Map DIB bits

    Call pvMapDIBits(uSAC, aBitsC())

    '-- 'Mask' transparent color ('null' pixel)

    With m_uBIH
        yEnd = .biHeight - 1
        xEnd = .biWidth - 1
    End With
    
    For y = 0 To yEnd
        For x = 0 To xEnd
            With aBitsC(x, y)
                If (aBitsM(x, y) <> 0) Then
                    .A = 255
                  Else
                    .R = 0
                    .G = 0
                    .B = 0
                End If
            End With
        Next x
    Next y

    '-- Unmap DIB bits / clean up

    Call pvUnmapDIBits(aBitsC())
    Call SelectObject(lhDC1, lhBmp1Old)
    Call DeleteDC(lhDC1)
    Call SelectObject(lhDC2, lhBmp2Old)
    Call DeleteObject(lhBmp2)
    Call DeleteDC(lhDC2)
End Sub

Private Sub pvRenderDropShadow( _
            ByVal ShadowColor As OLE_COLOR, _
            ByVal Opacity As Long, _
            ByVal xOffset As Long, _
            ByVal yOffset As Long _
            )

  Dim lColor   As Long
  Dim R        As Long
  Dim G        As Long
  Dim B        As Long
  Dim aBits()  As RGBQUAD
  Dim uSA      As SAFEARRAY2D
  Dim lAcm     As Long
  Dim x        As Long, xStart As Long, xEnd As Long, xIn As Long
  Dim y        As Long, yStart As Long, yEnd As Long, yIn As Long
  Dim i        As Long, iIn As Long
  Dim j        As Long, jIn As Long
  
    '-- Check opacity [0-100]

    If (Opacity < 0) Then Opacity = 0 Else If (Opacity > 100) Then Opacity = 100

    '-- Map DIB bits

    Call pvMapDIBits(uSA, aBits())

    '-- 'Apply' shadow (3x3 blur + offset)
    
    With m_uBIH
        yStart = yOffset:      yEnd = (.biHeight - 1) + BLUR_EXTENT
        xStart = -BLUR_EXTENT: xEnd = (.biWidth - 1) - xOffset
    End With

    Call OleTranslateColor(ShadowColor, 0, lColor)
    R = (lColor And &HFF&)
    G = (lColor And &HFF00&) \ &H100
    B = (lColor And &HFF0000) \ &H10000
    
    For y = yStart To yEnd
        yIn = y - yOffset
        For x = xStart To xEnd
            xIn = x + xOffset
            
            If (aBits(xIn, yIn).A = 0) Then

                lAcm = 0
                
                For j = -1 To 1
                    jIn = y + j
                    For i = -1 To 1
                        iIn = x + i
                        If (iIn > xStart And jIn < yEnd) Then
                            lAcm = lAcm + aBits(iIn, jIn).A
                        End If
                    Next i
                Next j

                If (lAcm) Then
                    With aBits(xIn, yIn)
                        .A = (lAcm * Opacity) \ 900
                        .R = (R * .A) \ 255
                        .G = (G * .A) \ 255
                        .B = (B * .A) \ 255
                    End With
                End If
            End If
        Next x
    Next y

    '-- Unmap DIB bits
    
    Call pvUnmapDIBits(aBits())
End Sub

Private Sub pvMapDIBits(uSA As SAFEARRAY2D, aBits() As RGBQUAD)

    With uSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_uBIH.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = m_uBIH.biWidth
        .pvData = m_lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIBits(aBits() As RGBQUAD)

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
