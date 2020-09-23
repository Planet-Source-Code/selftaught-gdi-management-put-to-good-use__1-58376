Attribute VB_Name = "mGDI"
'==================================================================================================
'mGDI.bas              1/17/04
'
'           GENERAL PURPOSE:
'               Create, Destroy and manipulate gdi objects.
'               Identical brushes, pens and fonts are pooled using pcGDIObjectStore.cls.
'               Provide debugging and statistic functions that are enabled using compiler switches.
'               Numerous enums that make api calls much more entertaining.
'
'               This module provides 3 kinds of functions:
'                   Public API functions
'                       various.  Tried to keep it minimal or else I'd be tempted to put them in a tlb.
'
'                   Functions that are public API decs or regular public functions depending on compiler switches
'                       i.e. CreateDC is exposed as an api declaration unless bDebugDCs is true, when
'                            a regular public function is exposed that calls the api and tracks the handle for debugging.
'
'                   Functions that are always regular public functions but have a signature identical to api functions
'                       i.e. CreateFont is always exposed as a public function because even when it is not in
'                            debug mode the handles must be tracked for pooling.
'
'           LINEAGE:
'               TileArea from www.vbaccelerator.com
'               DrawGradient from www.pscode.com submission "Let's talk about speed" by Light Templer
'
'           COMPILER SWITCHES:
'
'               bDebugStatistics    -   Exposes the Statistics sub - if this is true, then the same constant must be true in pcGDIObjectStore.cls in order to compile.
'               bDebugBitmaps       -   Displays a warning if bitmaps are leaked
'               bDebugDCs           -   Displays a warning if DCs are leaked
'               bDebugSelects       -   Displays Selects in the debug window
'
'               Additional GDI compiler switches available in pcGDIObjectStore.cls
'
'==================================================================================================

Option Explicit

#Const bDebug = False

#Const bDebugStatistics = True
#Const bDebugBitmaps = bDebug
#Const bDebugDCs = bDebug
#Const bDebugSelects = bDebug

#If bDebugBitmaps Then
    Private moDebugBitmaps As pcDebug
#End If

#If bDebugDCs Then
    Private moDebugDCs As pcDebug
#End If

Public Enum eBrushStyle
    gdiBSDibPattern = 5
    gdiBSDibPatternPt = 6
    gdiBSHatched = 2
    gdiBSNull = 1
    gdiBSPattern = 3
    gdiBSSolid = 0
End Enum

Public Enum eHatchStyle
    gdiHSBDiagonal = 3
    gdiHSCross = 4
    gdiHSDiagCross = 5
    gdiHSFDiagonal = 2
    gdiHSHorizontal = 0
    gdiHSVertical = 1
End Enum

Public Enum eTextAlignment
    gdiTALeft = 0
    gdiTANoUpdateCurPos = 0
    gdiTATop = 0
    gdiTARight = 2
    gdiTAUpdateCurPos = 1
    gdiTACenter = 6
    gdiTABottom = 8
    gdiTABaseLine = 24
End Enum

Public Enum eGdiObjectType
    gdiBitmap = 7
    gdiBrush = 2
    gdiColorSpace = 14
    gdiDC = 3
    gdiEnhMetaDc = 12
    gdiEnhMetaFile = 13
    gdiExtPen = 11
    gdiFont = 6
    gdiMemDc = 10
    gdiMetaFile = 9
    gdiMetaDc = 4
    gdiPal = 5
    gdiPen = 1
    gdiRegion = 8
End Enum

Public Enum eMapMode
     gdiMMAnisotripic = 8
     gdiMMHiEnglish = 5
     gdiMMHiMetric = 3
     gdiMMIsotropic = 7
     gdiMMLoMetric = 2
     gdiMMText = 1
     gdiMMTwips = 6
End Enum

Public Enum eDeviceCapability
    gdiCapDriverVersion = 0
    gdiCapTechnology = 2
    gdiCapHorzSize = 4
    gdiCapVertSize = 6
    gdiCapHorzRes = 8
    gdiCapVertRes = 10
    gdiCapLogPixelsX = 88
    gdiCapLogPixelsY = 90
    gdiCapBitsPixel = 12
    gdiCapPlanes = 14
    gdiCapNumBrushes = 16
    gdiCapNumPens = 18
    gdiCapNumFonts = 22
    gdiCapNumColors = 24
    gdiCapAspectX = 40
    gdiCapAspectY = 42
    gdiCapPDeviceSize = 26
    gdiCapClipCaps = 36
    gdiCapSizePalette = 104
    gdiCapNumReserved = 106
    gdiCapColorRes = 108
    gdiCapPhysicalWidth = 110
    gdiCapPhysicalHeight = 111
    gdiCapPhysicalOffsetX = 112
    gdiCapPhysicalOffsetY = 113
    gdiCapVRefresh = 116
    gdiCapDesktopHorzRes = 118
    gdiCapDesktopVertRes = 117
    gdiCapBltAlignment = 119
    gdiCapRasterCaps = 38
    gdiCapCurveCaps = 28
    gdiCapLineCaps = 30
End Enum

Public Enum ePenStyle
    gdiPSDash = 1
    gdiPSDashDot = 3
    gdiPSDashDotDot = 4
    gdiPSDot = 2
    gdiPSNull = 5
    gdiPSSolid = 0
    gdiPSInsideFrame = 6
End Enum

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Const LF_FACESIZE As Long = 32
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type

Public Type LOGBRUSH
    lbStyle As eBrushStyle
    lbColor As OLE_COLOR
    lbHatch As eHatchStyle
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Public Const DisplayDriver As String = "DISPLAY"

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As eDeviceCapability) As Long
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As eGdiObjectType
Public Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetMapMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nMapMode As eMapMode) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As eTextAlignment) As Long
Public Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

#If bDebugBitmaps Then
    Private Declare Function CreateBitmapIndirectApi Lib "gdi32.dll" Alias "CreateBitmapIndirect" (lpBitmap As BITMAP) As Long
    Private Declare Function CreateCompatibleBitmapApi Lib "gdi32.dll" Alias "CreateCompatibleBitmap" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateDIBSectionApi Lib "gdi32.dll" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
#Else
    Public Declare Function CreateBitmapIndirect Lib "gdi32.dll" (lpBitmap As BITMAP) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
#End If
    
#If bDebugDCs Then
    Private Declare Function CreateCompatibleDCApi Lib "gdi32.dll" Alias "CreateCompatibleDC" (ByVal hdc As Long) As Long
    Private Declare Function CreateDCApi Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
    Private Declare Function DeleteDCApi Lib "gdi32.dll" Alias "DeleteDC" (ByVal hdc As Long) As Long
#Else
    Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
    Public Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
#End If

#If bDebugSelects Then
    Private Declare Function SelectObjectApi Lib "gdi32.dll" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
#Else
    Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
#End If

Private Declare Function DeleteObjectApi Lib "gdi32.dll" Alias "DeleteObject" (ByVal hObject As Long) As Long

Private moBrushes As pcGDIObjectStore
Private moPens As pcGDIObjectStore
Private moFonts As pcGDIObjectStore

#If bDebugBitmaps Then
    
    Public Function CreateDIBSection(ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
        pInitDebugBitmap
        CreateDIBSection = CreateDIBSectionApi(hdc, pBitmapInfo, un, lplpVoid, handle, dw)
        If CreateDIBSection Then moDebugBitmaps.Add CreateDIBSection Else Debug.Assert False
    End Function

    Public Function CreateCompatibleBitmap(ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        pInitDebugBitmap
        CreateCompatibleBitmap = CreateCompatibleBitmapApi(hdc, nWidth, nHeight)
        If CreateCompatibleBitmap Then moDebugBitmaps.Add CreateCompatibleBitmap Else Debug.Assert False
    End Function
    
    Public Function CreateBitmapIndirect(lpBitmap As BITMAP) As Long
        pInitDebugBitmap
        CreateBitmapIndirect = CreateBitmapIndirectApi(lpBitmap)
        If CreateBitmapIndirect Then moDebugBitmaps.Add CreateBitmapIndirect Else Debug.Assert False
    End Function
    
    Private Sub pInitDebugBitmap()
        If moDebugBitmaps Is Nothing Then
            Set moDebugBitmaps = New pcDebug
            moDebugBitmaps.Module = "mGDI"
            moDebugBitmaps.Name = "Bitmap"
        End If
    End Sub
    
#End If

#If bDebugDCs Then
    
    Public Function CreateCompatibleDC(ByVal hdc As Long) As Long
        pInitDebugDCs
        CreateCompatibleDC = CreateCompatibleDCApi(hdc)
        If CreateCompatibleDC Then moDebugDCs.Add CreateCompatibleDC
    End Function
    
    Public Function CreateDC(ByRef lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
        pInitDebugDCs
        CreateDC = CreateDCApi(lpDriverName, lpDeviceName, lpOutput, lpInitData)
        If CreateDC Then moDebugDCs.Add CreateDC
    End Function
    
    Private Sub pInitDebugDCs()
        If moDebugDCs Is Nothing Then
            Set moDebugDCs = New pcDebug
            moDebugDCs.Module = "mGDI"
            moDebugDCs.Name = "DC"
        End If
    End Sub
    
    Public Function DeleteDC(ByVal hdc As Long) As Long
        DeleteDC = DeleteObject(hdc)
    End Function
    
#End If

#If bDebugSelects Then
    
    Public Function SelectObject(ByVal hdc As Long, ByVal hObject As Long) As Long
        SelectObject = SelectObjectApi(hdc, hObject)
        Debug.Print "Selected Object: " & hObject, "Into: " & hdc, "Old: " & SelectObject
        Debug.Assert SelectObject
    End Function
    
#End If


Public Function DeleteObject(ByVal hObject As Long) As Long
    Dim liType As eGdiObjectType
    
    liType = GetObjectType(hObject)
    If liType = gdiBrush Then
        DeleteObject = pDeleteBrush(hObject)
    ElseIf liType = gdiPen Then
        DeleteObject = pDeletePen(hObject)
    ElseIf liType = gdiFont Then
        DeleteObject = pDeleteFont(hObject)
    
    ElseIf liType = gdiBitmap Then
        DeleteObject = DeleteObjectApi(hObject)
        
        #If bDebugBitmaps Then
            If CBool(DeleteObject) And Not moDebugBitmaps Is Nothing Then moDebugBitmaps.Remove hObject Else Debug.Assert False
        #End If

    ElseIf liType = gdiMemDc Or liType = gdiDC Then
        
        #If bDebugDCs Then
            DeleteObject = DeleteDCApi(hObject)
            If CBool(DeleteObject) And Not moDebugDCs Is Nothing Then moDebugDCs.Remove hObject Else Debug.Assert False
            
        #Else
            DeleteObject = DeleteDC(hObject)
            
        #End If
    
    Else
        
        Debug.Assert False
        
        If hObject Then
            
            If liType <> gdiEnhMetaDc And liType <> gdiMetaDc Then
                DeleteObject = DeleteObjectApi(hObject)
            Else
                #If bDebugDCs Then
                    DeleteObject = DeleteDCApi(hObject)
                #Else
                    DeleteObject = DeleteDC(hObject)
                #End If
            End If
            
            Debug.Print "Deleted Misc. (" & liType & ") Object: " & hObject, "Return Value: " & DeleteObject
            
        End If
        
    End If
    
    Debug.Assert DeleteObject
End Function

Public Function CreateSolidBrush(ByVal iColor As OLE_COLOR) As Long
    
    Dim ltBrush As LOGBRUSH
    With ltBrush
        .lbColor = iColor
        .lbStyle = gdiBSSolid
    End With
    
    CreateSolidBrush = CreateBrushIndirect(ltBrush)
    
End Function

Public Function CreateBrushIndirect(ByRef tLogBrush As LOGBRUSH) As Long
        
    If moBrushes Is Nothing Then
        Set moBrushes = New pcGDIObjectStore
        moBrushes.Init gdiBrush
    End If
    
    CreateBrushIndirect = moBrushes.AddRef(VarPtr(tLogBrush))
    Debug.Assert CreateBrushIndirect

End Function

Private Function pDeleteBrush(ByVal hBrush As Long) As Long
    If Not moBrushes Is Nothing Then
        pDeleteBrush = moBrushes.Release(hBrush)
    End If
    
    Debug.Assert pDeleteBrush
    If pDeleteBrush = ZeroL Then
        pDeleteBrush = DeleteObjectApi(hBrush)
        Debug.Print "Deleted Unknown/Pattern Brush " & hBrush, "Return Value: " & pDeleteBrush
    End If
    
End Function

Public Function CreateFont( _
            Optional ByVal nHeight As Long, _
            Optional ByVal nWidth As Long, _
            Optional ByVal nOrientation As Long, _
            Optional ByVal nWeight As Long, _
            Optional ByVal bItalic As Boolean, _
            Optional ByVal bUnderline As Boolean, _
            Optional ByVal bStrikeOut As Boolean, _
            Optional ByVal nCharSet As Byte, _
            Optional ByVal nOutputPrecision As Byte, _
            Optional ByVal nClipPrecision As Byte, _
            Optional ByVal nQuality As Byte, _
            Optional ByVal nPitchAndFamily As Byte, _
            Optional ByRef sFaceName As String) _
                As Long

    Dim ltLF As LOGFONT
    With ltLF
        .lfHeight = nHeight
        .lfWeight = nWidth
        .lfOrientation = nOrientation
        .lfWeight = nWeight
        .lfItalic = Abs(bItalic)
        .lfUnderline = Abs(bUnderline)
        .lfStrikeOut = Abs(bStrikeOut)
        .lfCharSet = nCharSet
        .lfOutPrecision = nOutputPrecision
        .lfClipPrecision = nClipPrecision
        .lfQuality = nQuality
        .lfPitchAndFamily = nPitchAndFamily
        
        Dim lsTemp As String
        Dim liLen As Long
        
        lsTemp = StrConv(sFaceName, vbFromUnicode)
        liLen = LenB(sFaceName)
        If liLen > LF_FACESIZE Then liLen = LF_FACESIZE
        
        If liLen Then CopyMemory .lfFaceName(0), ByVal StrPtr(lsTemp), liLen
        
        CreateFont = CreateFontIndirect(ltLF)
        
    End With
End Function

Public Function CreateFontIndirect(ByRef tLogFont As LOGFONT) As Long
    If moFonts Is Nothing Then
        Set moFonts = New pcGDIObjectStore
        moFonts.Init gdiFont
    End If
    
    CreateFontIndirect = moFonts.AddRef(VarPtr(tLogFont))
    Debug.Assert CreateFontIndirect
    
End Function

Private Function pDeleteFont(ByVal hFont As Long) As Long
    If Not moFonts Is Nothing Then pDeleteFont = moFonts.Release(hFont)

    Debug.Assert pDeleteFont
    If pDeleteFont = ZeroL Then
        pDeleteFont = DeleteObjectApi(hFont)
        Debug.Print "Deleted Unknown Font: " & hFont, "Return Value: " & pDeleteFont
    End If
    
End Function


Public Function CreatePen(ByVal nPenStyle As ePenStyle, ByVal nWidth As Long, ByVal crColor As OLE_COLOR) As Long
    Dim ltPen As LOGPEN
    
    ltPen.lopnColor = crColor
    ltPen.lopnStyle = nPenStyle
    ltPen.lopnWidth.x = nWidth
    
    CreatePen = CreatePenIndirect(ltPen)
    
End Function

Public Function CreatePenIndirect(ByRef tLogPen As LOGPEN) As Long
    If moPens Is Nothing Then
        Set moPens = New pcGDIObjectStore
        moPens.Init gdiPen
    End If

    CreatePenIndirect = moPens.AddRef(VarPtr(tLogPen))
    Debug.Assert CreatePenIndirect
    
End Function

Private Function pDeletePen(ByVal hPen As Long) As Long
    If Not moPens Is Nothing Then pDeletePen = moPens.Release(hPen)

    Debug.Assert pDeletePen
    If pDeletePen = ZeroL Then
        pDeletePen = DeleteObjectApi(hPen)
        Debug.Print "Deleted Unknown Pen: " & hPen, "Return Value: " & pDeletePen
    End If

End Function

Public Sub DrawGradient(ByVal hdc As Long, _
                        ByVal iLeft As Long, ByVal iTop As Long, _
                        ByVal iWidth As Long, ByVal iHeight As Long, _
                        ByVal iColorFrom As OLE_COLOR, ByVal iColorTo As OLE_COLOR, _
                        Optional ByVal bVertical As Boolean)
    
    Dim ltBits() As RGBQUAD, ltBIH As BITMAPINFOHEADER
    
    Dim R  As Long, G  As Long, b  As Long
    Dim dR As Long, dG As Long, dB As Long
    Dim d  As Long, dEnd As Long
    
    If bVertical Then
        dEnd = iHeight
        'swap to/from colors
        iColorTo = iColorTo Xor iColorFrom
        iColorFrom = iColorTo Xor iColorFrom
        iColorTo = iColorFrom Xor iColorTo
    Else
        dEnd = iWidth
    End If
    
    If dEnd > 0& Then
        
        'ensure RGB format
        OleTranslateColor iColorTo, 0&, iColorTo
        OleTranslateColor iColorFrom, 0&, iColorFrom
        
        'split from color to R, G, B
        R = iColorFrom And &HFF&
        iColorFrom = iColorFrom \ &H100&
        G = iColorFrom And &HFF&
        iColorFrom = iColorFrom \ &H100&
        b = iColorFrom And &HFF&
        
        'get the relative changes in R, G, B
        dR = (iColorTo And &HFF&) - R
        iColorTo = iColorTo \ &H100&
        dG = (iColorTo And &HFF&) - G
        iColorTo = iColorTo \ &H100&
        dB = (iColorTo And &HFF&) - b
    
        'allocate the bitmap bits
        ReDim ltBits(0 To dEnd - 1&)
        
        For d = 0& To dEnd - 1&
            With ltBits(d)
                .rgbRed = (R + dR * d \ dEnd)
                .rgbGreen = (G + dG * d \ dEnd)
                .rgbBlue = (b + dB * d \ dEnd)
            End With
        Next
        
        With ltBIH
            'initialize the bitmap structure
            .biSize = Len(ltBIH)
            .biBitCount = 32&
            .biPlanes = 1&
            
            If bVertical Then
                .biWidth = 1&
                .biHeight = dEnd
            Else
                .biWidth = dEnd
                .biHeight = 1&
            End If
            
            'draw the gradient!
            StretchDIBits hdc, _
                iLeft, iTop, iLeft + iWidth, iTop + iHeight, _
                0&, 0&, .biWidth, .biHeight, _
                ltBits(0), ltBIH, 0&, vbSrcCopy
            
        End With
    End If
End Sub

Public Sub TileArea( _
            ByVal hDcDst As Long, _
            ByVal xDst As Long, _
            ByVal yDst As Long, _
            ByVal cxDst As Long, _
            ByVal cyDst As Long, _
            ByVal hDcSrc As Long, _
            ByVal cxSrc As Long, _
            ByVal cySrc As Long, _
   Optional ByVal cyOffset As Long, _
   Optional ByVal cxOffset As Long)

Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((xDst + cxOffset) Mod cxSrc)
    lSrcStartY = ((yDst + cyOffset) Mod cySrc)
    lSrcStartWidth = (cxSrc - lSrcStartX)
    lSrcStartHeight = (cySrc - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = yDst
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (yDst + cyDst)
        If (lDstY + lDstHeight) > (yDst + cyDst) Then
            lDstHeight = yDst + cyDst - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = xDst
        lSrcX = lSrcStartX
        Do While lDstX < (xDst + cxDst)
            If (lDstX + lDstWidth) > (xDst + cxDst) Then
                lDstWidth = xDst + cxDst - lDstX
                If (lDstWidth = ZeroL) Then lDstWidth = 4&
            End If
            BitBlt hDcDst, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = cxSrc
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = cySrc
    Loop
End Sub

#If bDebugStatistics Then
    
    Public Sub Statistics(ByRef iFontsRequested As Long, ByRef iFontsCreated As Long, ByRef iBrushesRequested As Long, ByRef iBrushesCreated As Long, ByRef iPensRequested As Long, ByRef iPensCreated As Long)
        
        If Not moFonts Is Nothing Then moFonts.Statistics iFontsRequested, iFontsCreated
        If Not moBrushes Is Nothing Then moBrushes.Statistics iBrushesRequested, iBrushesCreated
        If Not moPens Is Nothing Then moPens.Statistics iPensRequested, iPensCreated
        
    End Sub

#End If
