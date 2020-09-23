Attribute VB_Name = "mSysParamInfo"
'==================================================================================================
'mSysParamInfo.bas                      1/17/04
'
'           GENERAL PURPOSE:
'               Get a LOGFONT from System Fonts.
'
'           LINEAGE:
'               Padding size for NONCLIENTMETRICS from an article on www.vbaccelerator.com
'
'==================================================================================================

Option Explicit

Public Enum eSystemFonts
    sysFontMenu = 2
    sysFontMessage
    sysFontStatus
    sysFontCaption
    sysFontSmallCaption
End Enum

Private Const SPI_GETNONCLIENTMETRICS = 41&
Private Const LF_FACESIZE = 32&

Private Type NMLOGFONT
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
    lfFaceName(0 To LF_FACESIZE - 4) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As NMLOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As NMLOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As NMLOGFONT
    lfStatusFont As NMLOGFONT
    lfMessageFont As NMLOGFONT
    Padding(0 To 14) As Byte
End Type

Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private mtNCM As NONCLIENTMETRICS

Public Sub GetSystemFont(ByVal iFontType As eSystemFonts, ByRef tLogFont As LOGFONT)
    Dim lR As Long
    Dim liLen As Long
    
    liLen = Len(mtNCM)
    
    mtNCM.cbSize = liLen
    lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, liLen, mtNCM, ZeroL)
    
    If iFontType = sysFontMessage Then
        pLF mtNCM.lfMessageFont, tLogFont
    ElseIf iFontType = sysFontCaption Then
        pLF mtNCM.lfCaptionFont, tLogFont
    ElseIf iFontType = sysFontMenu Then
        pLF mtNCM.lfMenuFont, tLogFont
    ElseIf iFontType = sysFontSmallCaption Then
        pLF mtNCM.lfSMCaptionFont, tLogFont
    ElseIf iFontType = sysFontStatus Then
        pLF mtNCM.lfStatusFont, tLogFont
    End If
    
End Sub

Private Sub pLF(ByRef tNCLF As NMLOGFONT, ByRef tLF As LOGFONT)
    CopyMemory tLF, tNCLF, Len(tNCLF)
    ZeroMemory tLF.lfFaceName(LF_FACESIZE - 5), 4&
End Sub
