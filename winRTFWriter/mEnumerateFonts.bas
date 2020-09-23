Attribute VB_Name = "mEnumerateFonts"
Option Explicit

' Win32 SDK recommends the use of EnumFontFamiliesEx rather than the other versions:
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
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
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Type ENUMLOGFONTEX
    elfLogFont As LOGFONT
    elfFullName(LF_FULLFACESIZE - 1) As Byte
    elfStyle(LF_FACESIZE - 1) As Byte
    elfScript(LF_FACESIZE - 1) As Byte
End Type

Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ' Additional to TEXTMETRIC
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type

Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Type NEWTEXTMETRICEX
    ntmTm As NEWTEXTMETRIC
    ntmFontSig As FONTSIGNATURE
End Type

' Declares:
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long

Private m_lID As Long

Private mbShowStyle As Long

'/* EnumFonts Masks */
Public Const RASTER_FONTTYPE = 1&
Public Const DEVICE_FONTTYPE = 2&
Public Const TRUETYPE_FONTTYPE = 4&

Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
Public Const JOHAB_CHARSET = 130
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const THAI_CHARSET = 222
Public Const EASTEUROPE_CHARSET = 238
Public Const RUSSIAN_CHARSET = 204

Public Const MAC_CHARSET = 77
Public Const BALTIC_CHARSET = 186

Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_MODERN = 48
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

' Object to add items to:
Private m_cSink As IEnumFontSink
Private m_bPrinterFont As Boolean

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Function GetFonts( _
        ByVal lhDc As Long, _
        cSink As IEnumFontSink, _
        ByVal bPrinter As Boolean, _
        Optional ByVal sFaceName As String = "", _
        Optional ByVal lCharset As Long = ANSI_CHARSET _
    ) As Long
Dim tLF As LOGFONT
Dim I As Integer
    ' No re-entrancy, please:
    If Not (m_cSink Is Nothing) Then Exit Function
    ' Get the fonts:
    m_bPrinterFont = bPrinter
    Set m_cSink = cSink
    m_lID = m_lID + 1
    ' Convert the face name into a byte array:
    If Len(sFaceName) > 0 Then
        For I = 1 To Len(sFaceName)
            tLF.lfFaceName(I - 1) = Asc(Mid$(sFaceName, I, 1))
        Next I
        mbShowStyle = True
    Else
      mbShowStyle = False
    End If
    tLF.lfCharSet = lCharset
    ' Start the enumeration:
    GetFonts = EnumFontFamiliesEx(lhDc, tLF, AddressOf EnumFontFamExProc, mbShowStyle, 0)
    ' Clear up reference to the caller:
    Set m_cSink = Nothing
End Function

Public Function EnumFontFamExProc(ByVal lpelfe As Long, ByVal lpntme As Long, ByVal iFontType As Long, ByVal lParam As Long) As Long
' The callback function for EnumFontFamiliesEx.

' lpelf points to an ENUMLOGFONTEX structure, lpntm points to either
' a NEWTEXTMETRICEX (if true type) or a TEXTMETRIC (non-true type)
' structure.

Dim tLFEx As ENUMLOGFONTEX
Dim sFace As String, sScript As String
Dim sStyle As String, sFullName As String
Dim sFontFamily As String
Dim lTMP As Long
Dim lPos As Long
Dim sItem As String
Dim lPitch As Long

    CopyMemory tLFEx, ByVal lpelfe, LenB(tLFEx) ' Get the ENUMLOGFONTEX info
    ' Face Name
    sFace = StrConv(tLFEx.elfLogFont.lfFaceName, vbUnicode)
    lPos = InStr(sFace, Chr$(0))
    If (lPos > 0) Then sFace = Left$(sFace, (lPos - 1))
    
    lTMP = tLFEx.elfLogFont.lfPitchAndFamily
    
    If (lTMP And FF_DECORATIVE) = FF_DECORATIVE Then
        sFontFamily = "Decorative"
    End If
    
    If (lTMP And FF_DONTCARE) = FF_DONTCARE Then
        sFontFamily = "Dont Care"
    End If
    
    If (lTMP And FF_MODERN) = FF_MODERN Then
        sFontFamily = "Modern"
    End If
    
    If (lTMP And FF_ROMAN) = FF_ROMAN Then
        sFontFamily = "Roman"
    End If
    
    If (lTMP And FF_SWISS) = FF_SWISS Then
        sFontFamily = "Swiss"
    End If
    
    If (lTMP And DEFAULT_PITCH) = DEFAULT_PITCH Then
        lPitch = DEFAULT_PITCH
    End If
    
    If (lTMP And FIXED_PITCH) = FIXED_PITCH Then
        lPitch = FIXED_PITCH
    End If
    
    If (lTMP And VARIABLE_PITCH) = VARIABLE_PITCH Then
        lPitch = VARIABLE_PITCH
    End If
    
    ' Script
    sScript = StrConv(tLFEx.elfScript, vbUnicode)
    lPos = InStr(sScript, Chr$(0))
    If (lPos > 0) Then sScript = Left$(sScript, (lPos - 1))
    
    ' mbShowStyle
    If lParam = True Then
      ' Style
      sStyle = StrConv(tLFEx.elfStyle, vbUnicode)
      lPos = InStr(sStyle, Chr$(0))
      If (lPos > 0) Then sStyle = Left$(sStyle, (lPos - 1))
    Else
      sStyle = ""
    End If
    
    ' Full Name
    sFullName = StrConv(tLFEx.elfFullName, vbUnicode)
    lPos = InStr(sFullName, Chr$(0))
    If (lPos > 0) Then sFullName = Left$(sFullName, (lPos - 1))
    
    ' Only display printer and true type fonts:
    If Not (m_bPrinterFont) Then
        If (iFontType And TRUETYPE_FONTTYPE) <> TRUETYPE_FONTTYPE Then
            EnumFontFamExProc = 1
            Exit Function
        End If
    End If
    ' Only display a given font once:
    If Not (m_cSink.HasFont(sFace & " " & sScript)) Then
        m_cSink.AddFont sFace, sFontFamily, lPitch, sStyle, sScript, tLFEx.elfLogFont.lfCharSet, m_bPrinterFont
    End If
    ' Ask for more fonts:
    EnumFontFamExProc = 1
    
End Function


