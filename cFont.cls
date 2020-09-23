VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cFont.cls                  1/17/04
'
'           GENERAL PURPOSE:
'               Wrap a LOGFONT structure and expose its members.
'               Convert LOGFONT <--> StdFont
'
'               Browse For Font.
'
'               Font Source can be any system, ambient or custom font.
'               When set to a system or ambient font, you will need to call
'               OnSettingChange or OnAmbientFontChanged in the WM_SETTINGCHANGE
'               handler or the usercontrol's AmbientChanged event, respectively,
'               for the font to update itself.
'
'               Persistable.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit


Public Enum eFontWeight
    fntWeightDefault = 0
    fntWeightThin = 100
    fntWeightExtraLight = 200
    fntWeightLight = 300
    fntWeightNormal = 400
    fntWeightMedium = 500
    fntWeightSemibold = 600
    fntWeightBold = 700
    fntWeightExtraBold = 800
    fntWeightHeavy = 900
End Enum

Public Enum eFontCharset
    fntCharsetANSI = 0
    fntCharsetDefault = 1
    fntCharsetSymbol = 2
    fntCharsetShiftJIS = 128
    fntCharsetOEM = 255
End Enum

Public Enum eFontOutputPrecision
    fntOutPrecCharacter = 2
    fntOutPrecDefault = 0
    fntOutPrecDevice = 5
    fntOutPrecOutline = 8
    fntOutPrecRaster = 6
    fntOutPrecString = 1
    fntOutPrecStroke = 3
    fntOutPrecTTOnly = 7
    fntOutPrecTT = 4
End Enum

Public Enum eFontClipPrecision
    fntClipPrecDefault = 0
    fntClipPrecCharacter = 1
    fntClipPrecStroke = 2
    fntClipPrecMask = &HF&
    fntClipPrecEmbedded = 128
    fntClipPrecLHAngles = 16
End Enum

Public Enum eFontQuality
    fntQualityDefault = 0
    fntQualityDraft = 1
    fntQuallityProof = 2
End Enum

Public Enum eFontPitchAndFamily
    fntPitchDefault = 0
    fntPitchFixed = 1
    fntPitchVariable = 2
    fntFamilyDecorative = 80
    fntFamilyDefault = 0
    fntFamilyModern = 48
    fntFamilyRoman = 16
    fntFamilyScript = 64
    fntFamilySwiss = 32
End Enum

Public Enum eFontSource
    fntSourceCustom = 0
    fntSourceAmbient
    fntSourceSysMenu
    fntSourceSysMessage
    fntSourceSysStatus
    fntSourceSysCaption
    fntSourceSysSmallCaption
End Enum

Public Enum eFontDialog
    dlgFontScreenFonts = &H1
    dlgFontPrinterFonts = &H2
    dlgFontScreenAndPrinterFonts = &H3
    dlgFontUseStyle = &H80
    dlgFontEffects = &H100
    dlgFontAnsiOnly = &H400
    dlgFontNoVectorFonts = &H800
    dlgFontNoOemFonts = dlgFontNoVectorFonts
    dlgFontNoSimulations = &H1000
    dlgFontFixedPitchOnly = &H4000
    dlgFontWysiwyg = &H8000&  ' Must also have ScreenFonts And PrinterFonts
    dlgFontForceExist = &H10000
    dlgFontScalableOnly = &H20000
    dlgFontTTOnly = &H40000
    dlgFontNoFaceSel = &H80000
    dlgFontNoStyleSel = &H100000
    dlgFontNoSizeSel = &H200000
    ' Win95 only
    dlgFontSelectScript = &H400000
    dlgFontNoScriptSel = &H800000
    dlgFontNoVertFonts = &H1000000

    
    dlgFontApply = &H200
    dlgFontEnableHook = &H8
    'dlgFontEnableTemplate = &H10
    'dlgFontEnableTemplateHandle = &H20
    'dlgFontNotSupported = &H238
    dlgFontRaiseError = &H10000000
End Enum

Public Event Changed()

Private Type TCHOOSEFONT
    lStructSize                     As Long      ' Filled with UDT size
    hWndOwner                       As Long      ' Caller's window handle
    hdc                             As Long      ' Printer DC/IC or NULL
    lpLogFont                       As Long      ' Pointer to LOGFONT
    iPointSize                      As Long      ' 10 * size in points of font
    Flags                           As Long      ' Type flags
    rgbColors                       As Long      ' Returned text color
    lCustData                       As Long      ' Data passed to hook function
    lpfnHook                        As Long      ' Pointer to hook function
    lpTemplateName                  As Long      ' Custom template name
    hInstance                       As Long      ' Instance handle for template
    lpszStyle                       As String    ' Return style field
    nFontType                       As Integer   ' Font type bits
    iAlign                          As Integer   ' Filler
    nSizeMin                        As Long      ' Minimum point size allowed
    nSizeMax                        As Long      ' Maximum point size allowed
End Type

Private Const DEF_FaceName          As String = "MS Sans Serif"
Private Const DEF_Height            As Long = -11&
Private Const DEF_Width             As Long = ZeroL
Private Const DEF_Escapement        As Long = ZeroL
Private Const DEF_Orientation       As Long = ZeroL
Private Const DEF_Weight            As Long = fntWeightNormal
Private Const DEF_Italic            As Byte = ZeroY
Private Const DEF_Underline         As Byte = ZeroY
Private Const DEF_Strikeout         As Byte = ZeroY
Private Const DEF_Charset           As Byte = ZeroY
Private Const DEF_OutPrecision      As Byte = ZeroY
Private Const DEF_ClipPrecision     As Byte = ZeroY
Private Const DEF_Quality           As Byte = ZeroY
Private Const DEF_PitchAndFamily    As Byte = ZeroY
Private Const DEF_Source            As Long = fntSourceAmbient

Private Const PROP_FaceName         As String = "Name"
Private Const PROP_Height           As String = "Height"
Private Const PROP_Width            As String = "Width"
Private Const PROP_Escapement       As String = "Esc"
Private Const PROP_Orientation      As String = "Orient"
Private Const PROP_Weight           As String = "Weight"
Private Const PROP_Italic           As String = "Italic"
Private Const PROP_Underline        As String = "Underline"
Private Const PROP_Strikeout        As String = "Strikeout"
Private Const PROP_Charset          As String = "Charset"
Private Const PROP_OutPrecision     As String = "OutPrec"
Private Const PROP_ClipPrecision    As String = "ClipPrec"
Private Const PROP_Quality          As String = "Quality"
Private Const PROP_PitchAndFamily   As String = "PitchFam"
Private Const PROP_Source           As String = "Source"

Private Const SPI_SETNONCLIENTMETRICS As Long = 42

Private Const ClassName             As String = "cFont"

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByRef yStart1 As Byte, ByRef yStart2 As Byte) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private mtLogFont                   As LOGFONT
Private miFontSource                As eFontSource

Private Property Let pFaceName(ByRef sName As String)
    
    Dim ls As String
    Dim iLen As Long
    
    ls = StrConv(sName, vbFromUnicode)
    
    iLen = LenB(ls)
    If iLen > LF_FACESIZE Then iLen = LF_FACESIZE
    
    If iLen > 0& Then CopyMemory mtLogFont.lfFaceName(0), ByVal StrPtr(ls), iLen
    
    If iLen < LF_FACESIZE _
        Then ZeroMemory mtLogFont.lfFaceName(iLen), (LF_FACESIZE - iLen) _
        Else mtLogFont.lfFaceName(LF_FACESIZE - 1&) = ZeroY
    
End Property

Private Property Get pFaceName() As String
    pFaceName = StrConv(mtLogFont.lfFaceName, vbUnicode)
    Dim i As Long
    i = InStr(1&, pFaceName, vbNullChar)
    If i Then pFaceName = Left$(pFaceName, i - 1&)
End Property

Private Sub pPutStdFont(ByVal oFont As StdFont)
    On Error Resume Next
    pFaceName = oFont.Name
    With mtLogFont
        .lfHeight = -MulDiv(oFont.Size, 1440& / Screen.TwipsPerPixelY, 72&)
        .lfWeight = IIf(oFont.Bold, fntWeightBold, fntWeightNormal)
        .lfItalic = Abs(oFont.Italic)
        .lfUnderline = Abs(oFont.Underline)
        .lfStrikeOut = Abs(oFont.Strikethrough)
        .lfCharSet = oFont.Charset And &HFF
        .lfEscapement = 0&
        .lfOrientation = 0&
        .lfWidth = 0&
        .lfOutPrecision = 0
        .lfClipPrecision = 0
        .lfQuality = 0
        .lfPitchAndFamily = 0
    End With
    On Error GoTo 0
End Sub

Private Sub pGetStdFont(ByVal oFont As StdFont)
    On Error Resume Next
    With oFont
        .Name = pFaceName()
        If mtLogFont.lfHeight Then
            .Size = MulDiv(72&, Abs(mtLogFont.lfHeight), (1440& / Screen.TwipsPerPixelY))
        End If
        .Charset = mtLogFont.lfCharSet
        .Italic = CBool(mtLogFont.lfItalic)
        .Underline = CBool(mtLogFont.lfUnderline)
        .Strikethrough = CBool(mtLogFont.lfStrikeOut)
        .Bold = CBool(mtLogFont.lfWeight > fntWeightNormal)
    End With
    On Error GoTo 0
End Sub

Private Sub pPutFont(ByVal oFont As cFont)
    oFont.fGetLogFont mtLogFont, miFontSource
End Sub

Private Sub pGetFont(ByVal oFont As cFont)
    oFont.fPutLogFont mtLogFont, miFontSource
End Sub

Friend Sub fPutLogFont(ByRef tLogFont As LOGFONT, ByVal iFontSource As eFontSource)
    LSet mtLogFont = tLogFont
    miFontSource = iFontSource
    RaiseEvent Changed
End Sub

Friend Sub fGetLogFont(ByRef tLogFont As LOGFONT, ByRef iFontSource As eFontSource)
    LSet tLogFont = mtLogFont
    iFontSource = miFontSource
End Sub

Friend Sub fPutLogFontLong(ByVal lptr As Long)
    CopyMemory mtLogFont, ByVal lptr, Len(mtLogFont)
    miFontSource = fntSourceCustom
    RaiseEvent Changed
End Sub

Friend Sub fGetLogFontLong(ByVal lptr As Long)
    CopyMemory ByVal lptr, mtLogFont, Len(mtLogFont)
End Sub


Public Property Get Height() As Long
    Height = mtLogFont.lfHeight
End Property
Public Property Let Height(ByVal iNew As Long)
    If CBool(iNew Xor mtLogFont.lfHeight) Then
        mtLogFont.lfHeight = iNew
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Width() As Long
    Width = mtLogFont.lfWidth
End Property
Public Property Let Width(ByVal iNew As Long)
    If CBool(iNew Xor mtLogFont.lfWidth) Then
        mtLogFont.lfWidth = iNew
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Escapement() As Long
    Escapement = mtLogFont.lfEscapement
End Property
Public Property Let Escapement(ByVal iNew As Long)
    If CBool(iNew Xor mtLogFont.lfEscapement) Then
        mtLogFont.lfEscapement = iNew
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Orientation() As Long
    Orientation = mtLogFont.lfOrientation
End Property
Public Property Let Orientation(ByVal iNew As Long)
    If CBool(iNew Xor mtLogFont.lfOrientation) Then
        mtLogFont.lfOrientation = iNew
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Weight() As eFontWeight
    Weight = mtLogFont.lfWeight
End Property
Public Property Let Weight(ByVal iNew As eFontWeight)
    If CBool(iNew Xor mtLogFont.lfWeight) Then
        mtLogFont.lfWeight = iNew
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Italic() As Boolean
    Italic = CBool(mtLogFont.lfItalic)
End Property
Public Property Let Italic(ByVal bNew As Boolean)
    If CBool(bNew Xor Italic) Then
        mtLogFont.lfItalic = Abs(bNew)
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Underline() As Boolean
    Underline = CBool(mtLogFont.lfUnderline)
End Property
Public Property Let Underline(ByVal bNew As Boolean)
    If CBool(bNew Xor Underline) Then
        mtLogFont.lfUnderline = Abs(bNew)
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Strikeout() As Boolean
    Strikeout = mtLogFont.lfStrikeOut
End Property
Public Property Let Strikeout(ByVal bNew As Boolean)
    If CBool(bNew Xor Me.Strikeout) Then
        mtLogFont.lfStrikeOut = Abs(bNew)
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Charset() As eFontCharset
    Charset = mtLogFont.lfCharSet
End Property
Public Property Let Charset(ByVal iNew As eFontCharset)
    If CBool(iNew Xor CLng(mtLogFont.lfCharSet)) Then
        mtLogFont.lfCharSet = iNew And &HFF
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get OutPrecision() As eFontOutputPrecision
    OutPrecision = mtLogFont.lfOutPrecision
End Property
Public Property Let OutPrecision(ByVal iNew As eFontOutputPrecision)
    If CBool(iNew Xor CLng(mtLogFont.lfOutPrecision)) Then
        mtLogFont.lfOutPrecision = iNew And &HFF
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get ClipPrecision() As eFontClipPrecision
    ClipPrecision = mtLogFont.lfClipPrecision
End Property
Public Property Let ClipPrecision(ByVal iNew As eFontClipPrecision)
    If CBool(iNew Xor CLng(mtLogFont.lfClipPrecision)) Then
        mtLogFont.lfClipPrecision = iNew And &HFF
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get Quality() As eFontQuality
    Quality = mtLogFont.lfQuality
End Property
Public Property Let Quality(ByVal iNew As eFontQuality)
    If CBool(iNew Xor CLng(mtLogFont.lfQuality)) Then
        mtLogFont.lfQuality = iNew And &HFF
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get PitchAndFamily() As eFontPitchAndFamily
    PitchAndFamily = mtLogFont.lfPitchAndFamily
End Property
Public Property Let PitchAndFamily(ByVal iNew As eFontPitchAndFamily)
    If CBool(iNew Xor CLng(mtLogFont.lfPitchAndFamily)) Then
        mtLogFont.lfPitchAndFamily = iNew And &HFF
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
End Property

Public Property Get FaceName() As String
    FaceName = pFaceName()
End Property

Public Property Let FaceName(ByRef sNew As String)
    pFaceName = sNew
    miFontSource = fntSourceCustom
    RaiseEvent Changed
End Property

Public Property Get Source() As eFontSource
    Source = miFontSource
End Property

Public Property Let Source(ByVal iNew As eFontSource)
    If iNew >= fntSourceCustom And iNew <= fntSourceSysSmallCaption Then
        If iNew > fntSourceAmbient Then
            mSysParamInfo.GetSystemFont iNew, mtLogFont
            miFontSource = iNew
            RaiseEvent Changed
        Else
            If iNew Xor miFontSource Then
                miFontSource = iNew
                RaiseEvent Changed
            End If
        End If
    Else
        miFontSource = fntSourceCustom
    End If
End Property

Public Sub OnAmbientFontChanged(ByVal oFont As StdFont)
    If miFontSource = fntSourceAmbient Then
        If Not oFont Is Nothing Then
            pPutStdFont oFont
            miFontSource = fntSourceAmbient
            RaiseEvent Changed
        End If
    End If
End Sub

Public Sub OnSettingChange(Optional ByVal wParam As Long = SPI_SETNONCLIENTMETRICS)
    If wParam = SPI_SETNONCLIENTMETRICS Then
        If miFontSource >= fntSourceSysMenu And miFontSource <= fntSourceSysSmallCaption Then
            mSysParamInfo.GetSystemFont miFontSource, mtLogFont
            RaiseEvent Changed
        End If
    End If
End Sub

Public Function GetHandle() As Long
    GetHandle = CreateFontIndirect(mtLogFont)
End Function
Public Function ReleaseHandle(ByVal hFont As Long) As Boolean
    ReleaseHandle = DeleteObject(hFont) <> 0&
End Function

Public Sub GetFontInfo(ByVal oObject As Object)
    If Not oObject Is Nothing Then
        If TypeOf oObject Is StdFont Then
            pGetStdFont oObject
        ElseIf TypeOf oObject Is cFont Then
            pGetFont oObject
        End If
    End If
End Sub

Public Sub PutFontInfo(ByVal oFont As Object)
    If Not oFont Is Nothing Then
        If TypeOf oFont Is StdFont Then
            pPutStdFont oFont
        ElseIf TypeOf oFont Is cFont Then
            pPutFont oFont
        End If
    End If
    miFontSource = fntSourceCustom
    RaiseEvent Changed
End Sub

Public Function Browse( _
                Optional ByVal iFlags As eFontDialog = dlgFontScreenFonts, _
                Optional ByVal hdc As Long, _
                Optional ByVal hWndOwner As Long, _
                Optional ByVal iMinSize As Long = 6, _
                Optional ByVal iMaxSize As Long = 72) _
                    As Boolean
    
    'init to log font and limit size
    Const FontFlags = &H2000& Or &H40&
    
    Dim ltLogFont As LOGFONT
   
    LSet ltLogFont = mtLogFont
   
    Dim ltChooseFont As TCHOOSEFONT
    With ltChooseFont
        .lStructSize = LenB(ltChooseFont)
        .Flags = iFlags Or FontFlags
        .hdc = hdc
        .hInstance = ZeroL
        .hWndOwner = hWndOwner
        .lpLogFont = VarPtr(ltLogFont)
        .nSizeMax = iMaxSize
        .nSizeMin = iMinSize
    End With
    
    If ChooseFont(ltChooseFont) = OneL Then
        Browse = True
        LSet mtLogFont = ltLogFont
        miFontSource = fntSourceCustom
        RaiseEvent Changed
    End If
    
End Function

Friend Function fComp(ByRef tLF As LOGFONT) As Boolean
    fComp = CBool(MemCmp(VarPtr(tLF), VarPtr(mtLogFont), LenB(mtLogFont)))
End Function

Public Function Comp(ByVal oFont As cFont) As Boolean
    Comp = oFont.fComp(mtLogFont)
End Function

Private Sub Class_InitProperties()
    miFontSource = DEF_Source
    pFaceName = DEF_FaceName
    With mtLogFont
        .lfCharSet = DEF_Charset
        .lfClipPrecision = DEF_ClipPrecision
        .lfEscapement = DEF_Escapement
        .lfHeight = DEF_Height
        .lfItalic = DEF_Italic
        .lfOrientation = DEF_Orientation
        .lfOutPrecision = DEF_OutPrecision
        .lfPitchAndFamily = DEF_PitchAndFamily
        .lfQuality = DEF_Quality
        .lfStrikeOut = DEF_Strikeout
        .lfUnderline = DEF_Underline
        .lfWeight = DEF_Weight
        .lfWidth = DEF_Width
    End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    miFontSource = PropBag.ReadProperty(PROP_Source, DEF_Source)
    If miFontSource >= fntSourceSysMenu And miFontSource <= fntSourceSysSmallCaption Then
        mSysParamInfo.GetSystemFont miFontSource, mtLogFont
    ElseIf miFontSource <> fntSourceAmbient Then
        miFontSource = fntSourceCustom
        pFaceName = PropBag.ReadProperty(PROP_FaceName, DEF_FaceName)
        With mtLogFont
            .lfHeight = PropBag.ReadProperty(PROP_Height, DEF_Height)
            .lfWidth = PropBag.ReadProperty(PROP_Width, DEF_Width)
            .lfEscapement = PropBag.ReadProperty(PROP_Escapement, DEF_Escapement)
            .lfOrientation = PropBag.ReadProperty(PROP_Orientation, DEF_Orientation)
            .lfWeight = PropBag.ReadProperty(PROP_Weight, DEF_Weight)
            .lfItalic = PropBag.ReadProperty(PROP_Italic, DEF_Italic)
            .lfUnderline = PropBag.ReadProperty(PROP_Underline, DEF_Underline)
            .lfStrikeOut = PropBag.ReadProperty(PROP_Strikeout, DEF_Strikeout)
            .lfCharSet = PropBag.ReadProperty(PROP_Charset, DEF_Charset)
            .lfOutPrecision = PropBag.ReadProperty(PROP_OutPrecision, DEF_OutPrecision)
            .lfClipPrecision = PropBag.ReadProperty(PROP_ClipPrecision, DEF_ClipPrecision)
            .lfQuality = PropBag.ReadProperty(PROP_Quality, DEF_Quality)
            .lfPitchAndFamily = PropBag.ReadProperty(PROP_PitchAndFamily, DEF_PitchAndFamily)
        End With
    End If
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_Source, miFontSource, DEF_Source
    If miFontSource = fntSourceCustom Then
        PropBag.WriteProperty PROP_FaceName, pFaceName, DEF_FaceName
        With mtLogFont
            PropBag.WriteProperty PROP_Height, .lfHeight, DEF_Height
            PropBag.WriteProperty PROP_Width, .lfWidth, DEF_Width
            PropBag.WriteProperty PROP_Escapement, .lfEscapement, DEF_Escapement
            PropBag.WriteProperty PROP_Orientation, .lfOrientation, DEF_Orientation
            PropBag.WriteProperty PROP_Weight, .lfWeight, DEF_Weight
            PropBag.WriteProperty PROP_Italic, .lfItalic, DEF_Italic
            PropBag.WriteProperty PROP_Underline, .lfUnderline, DEF_Underline
            PropBag.WriteProperty PROP_Strikeout, .lfStrikeOut, DEF_Strikeout
            PropBag.WriteProperty PROP_Charset, .lfCharSet, DEF_Charset
            PropBag.WriteProperty PROP_OutPrecision, .lfOutPrecision, DEF_OutPrecision
            PropBag.WriteProperty PROP_ClipPrecision, .lfClipPrecision, DEF_ClipPrecision
            PropBag.WriteProperty PROP_Quality, .lfQuality, DEF_Quality
            PropBag.WriteProperty PROP_PitchAndFamily, .lfPitchAndFamily, DEF_PitchAndFamily
        End With
    End If
End Sub
