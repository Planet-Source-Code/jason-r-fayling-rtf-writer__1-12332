Attribute VB_Name = "modRTFWriter"
Option Explicit

Public Const RTF_START = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033"
Public Const RTF_END = "}"

Public Const FONT_TABLE_START = "{\fonttbl"
Public Const FONT_TABLE_END = "}"

Public Const COLOR_TABLE_START = "{\colortbl "
Public Const COLOR_TABLE_END = ";}"

Public Const LINE_BREAK = "\par" & vbCrLf
Public Const RTF_TAB = "\tab "
Public Const NEW_PARAGRAPH = "\pard"
Public Const LEFT_INDENT = "\li"
Public Const RIGHT_INDENT = "\ri"
Public Const BULLET_BREAK = "\line "

Public Const ALIGN_LEFT = "\nowidctlpar\ql"
Public Const ALIGN_CENTER = "\nowidctlpar\qc"
Public Const ALIGN_RIGHT = "\nowidctlpar\qr"

Public Const BOLD_START = "\b"
Public Const BOLD_END = "\b0"

Public Const UNDERLINE_START = "\ul"
Public Const UNDERLINE_END = "\ulnone"

Public Const ITIALIC_START = "\i"
Public Const ITIALIC_END = "\i0"

Public Const DOCUMENT_START = "\viewkind4\uc1"



Public gobjFonts As winRTFFonts
Public gobjColors As winRTFColors
Public gobjSystemFonts As winRTFSystemFonts
Public gobjInstalledFonts As winRTFInstalledFonts

Public Function GetRedValue(ByVal lngColor As Long) As Integer

On Error GoTo ErrorCode

  GetRedValue = lngColor And &HFF

ClearVariables:
  Exit Function

ErrorCode:
    GoTo ClearVariables

End Function

Public Function GetBlueValue(ByVal lngColor As Long) As Integer

On Error GoTo ErrorCode

  GetBlueValue = (lngColor \ &H10000) And &HFF

ClearVariables:
  Exit Function

ErrorCode:
    GoTo ClearVariables
End Function
Public Function GetGreenValue(ByVal lngColor As Long) As Integer

On Error GoTo ErrorCode

  GetGreenValue = (lngColor \ &H100) And &HFF

ClearVariables:
  Exit Function

ErrorCode:
    GoTo ClearVariables
    
End Function
Public Function RoundFontSize(FontSize As Double) As Double

On Error GoTo ErrorCode

Dim dNewSize As Double
Dim strFontSize As String
Dim lLeftSize As Long
Dim lRightSize As Long
Dim I As Integer

    If CInt(FontSize) = FontSize Then
        RoundFontSize = FontSize
        GoTo ClearVariables
    End If
    
    strFontSize = CStr(FontSize)
    I = InStr(1, strFontSize, ".", vbBinaryCompare)
    If I <> 0 Then
        Select Case Mid(strFontSize, I + 1, 1)
            Case 0 To 4
                RoundFontSize = CDbl(Mid(strFontSize, 1, I - 1))
            Case 5
                RoundFontSize = CDbl(Mid(strFontSize, 1, I + 1))
            Case 6 To 9
                RoundFontSize = CDbl(Mid(strFontSize, 1, I - 1)) + 1
        End Select
    Else
        RoundFontSize = FontSize
        GoTo ClearVariables
    End If
    
ClearVariables:
    Exit Function
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Function
