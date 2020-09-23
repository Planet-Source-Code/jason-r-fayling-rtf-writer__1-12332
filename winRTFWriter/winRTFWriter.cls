VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum winRTFAlignment
    winRTFAlignLeft = 0
    winRTFAlignCenter = 1
    winRTFAlignRight = 2
End Enum

Dim mstrTempFile As String
Dim miFileHandel As Integer
Dim mbolFileIsOpen As Boolean
Dim mlCurrentAlignment As winRTFAlignment
Dim mlCurrentLeftIndent As Long
Dim mlCurrentRightIndent As Long
Dim mbolDocumentHasContent As Boolean

' Objects
Private mobjCurrentFont As winRTFFont
Private mobjCurrentColor As winRTFColor
Private mobjFonts As winRTFFonts
Private mobjColors As winRTFColors

Public Sub AddBullet(ByVal Text As String, Optional ByVal Font As winRTFFont, Optional ByVal ForeColor As winRTFColor)

On Error GoTo ErrorCode

Dim strText As String
Dim strData As String
Dim objFont As winRTFFont
Dim objColor As winRTFColor
Dim objSymbol As winRTFFont
Dim lIndent As Long

Dim B1 As String
Dim B2 As String
Dim U1 As String
Dim U2 As String
Dim I1 As String
Dim I2 As String

    If mbolFileIsOpen = False Then GoTo ClearVariables
    
    If Font Is Nothing Then
        Set objFont = Me.CurrentFont
    Else
        Set objFont = Font
    End If
    
    If ForeColor Is Nothing Then
        Set objColor = Me.CurrentColor
    Else
        Set objColor = ForeColor
    End If
    
    With objFont
        If .Bold = True Then
            B1 = BOLD_START
            B2 = BOLD_END
        End If
        
        If .Itialic = True Then
            I1 = ITIALIC_START
            I2 = ITIALIC_END
        End If
        
        If .Underline = True Then
            U1 = UNDERLINE_START
            U2 = UNDERLINE_END
        End If
    End With

    If mobjFonts.DoesExist("Symbol") = False Then
        mobjFonts.Add "Symbol"
    End If
    
    Set objSymbol = mobjFonts("Symbol")
    
    lIndent = 4
    lIndent = lIndent + mlCurrentLeftIndent
    
    strData = CStr("{\pntext" & "\" & objSymbol.FontCode & "\'B7\tab}{\*\pn\pnlvlblt\pn" & objSymbol.FontCode & "\pnindent0{\pntxtb\'B7}}\fi-720\li" & (lIndent * 180) & " ")
    strText = Replace(Replace(Text, vbCrLf, BULLET_BREAK), vbTab, RTF_TAB)
    strText = CStr("\cf" & objColor.ColorCode & "\" & objFont.FontCode & "\fs" & (RoundFontSize(objFont.Size) * 2) & B1 & U1 & I1 & " " & strText & I2 & U2 & B2)
    
    WriteToFile strData & strText
    
ClearVariables:
    Set objSymbol = Nothing
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub
Public Sub AddLineBreak()
    WriteToFile LINE_BREAK
End Sub
Public Sub AddText(ByVal Text As String, Optional ByVal Font As winRTFFont, Optional ByVal ForeColor As winRTFColor)

On Error GoTo ErrorCode

Dim strText As String
Dim strData As String
Dim objFont As winRTFFont
Dim objColor As winRTFColor

Dim B1 As String
Dim B2 As String
Dim U1 As String
Dim U2 As String
Dim I1 As String
Dim I2 As String

    If mbolFileIsOpen = False Then GoTo ClearVariables
    
    If Font Is Nothing Then
        Set objFont = Me.CurrentFont
    Else
        Set objFont = Font
    End If
    
    If ForeColor Is Nothing Then
        Set objColor = Me.CurrentColor
    Else
        Set objColor = ForeColor
    End If
    
    With objFont
        If .Bold = True Then
            B1 = BOLD_START
            B2 = BOLD_END
        End If
        
        If .Itialic = True Then
            I1 = ITIALIC_START
            I2 = ITIALIC_END
        End If
        
        If .Underline = True Then
            U1 = UNDERLINE_START
            U2 = UNDERLINE_END
        End If
    End With
    
    strText = Replace(Replace(Text, vbCrLf, LINE_BREAK), vbTab, RTF_TAB)
    strText = CStr("\cf" & objColor.ColorCode & "\" & objFont.FontCode & "\fs" & (RoundFontSize(objFont.Size) * 2) & B1 & U1 & I1 & " " & strText & I2 & U2 & B2)

    WriteToFile strText
   

ClearVariables:
    Set objFont = Nothing
    Set objColor = Nothing
    
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Sub
Public Sub BeginDocument()

On Error GoTo ErrorCode

    If mobjColors.DoesExist("black") = True Then
        Set Me.CurrentColor = mobjColors("black")
    End If
    
    Set Me.CurrentFont = mobjFonts(1)

    KillFile
    OpenFile
    mbolDocumentHasContent = False
    
ClearVariables:
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub



Public Sub Clear()
    KillFile
    OpenFile
End Sub
Private Sub CloseFile()
    
On Error GoTo ErrorCode

    If mbolFileIsOpen = True Then Close miFileHandel
    mbolFileIsOpen = False
    
ClearVariables:
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub

Public Property Get Colors() As winRTFColors
    Set Colors = mobjColors
End Property
Private Function CreateColorTable() As String

On Error GoTo ErrorCode

Dim strData As String
Dim objColor As winRTFColor
    
    If mobjColors.Count < 1 Then GoTo ClearVariables
    
    strData = ""
    strData = strData & COLOR_TABLE_START
        For Each objColor In mobjColors
            If Not objColor Is Nothing Then
                With objColor
                    strData = strData & CStr(";\red" & .Red & "\green" & .Green & "\blue" & .Blue)
                End With
            End If
        Next
    strData = strData & COLOR_TABLE_END

    CreateColorTable = strData
    
ClearVariables:
    Set objColor = Nothing
    Exit Function
    
ErrorCode:
    strData = ""
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Function

Private Function CreateFontTable() As String

On Error GoTo ErrorCode

Dim strData As String
Dim objFont As winRTFFont
    
    If mobjFonts.Count < 1 Then GoTo ClearVariables
    
    strData = ""
    strData = strData & FONT_TABLE_START
        For Each objFont In mobjFonts
            If objFont.FontCode = "f0" Then
                strData = strData & "{\" & objFont.FontCode & "\fnil\fcharset" & objFont.Charset & " " & objFont.FontName & ";}"
            Else
                strData = strData & "{\" & objFont.FontCode & "\" & objFont.FontFamilyCode & "\fprq" & objFont.Pitch & "\fcharset" & objFont.Charset & " " & objFont.FontName & ";}"
            End If
        Next
    strData = strData & FONT_TABLE_END

    CreateFontTable = strData
    
ClearVariables:
    Set objFont = Nothing
    Exit Function
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Function

Public Property Let CurrentAlignment(ByVal NewValue As winRTFAlignment)
    
    If NewValue < winRTFAlignLeft Or NewValue > winRTFAlignRight Then
        NewValue = winRTFAlignLeft
    End If
    
    mlCurrentAlignment = NewValue
    
End Property
Public Property Get CurrentAlignment() As winRTFAlignment
    CurrentAlignment = mlCurrentAlignment
End Property

Public Property Set CurrentColor(ByVal NewValue As winRTFColor)
    Set mobjCurrentColor = NewValue
End Property
Public Property Get CurrentColor() As winRTFColor
    Set CurrentColor = mobjCurrentColor
End Property

Public Property Set CurrentFont(ByVal NewValue As winRTFFont)
    Set mobjCurrentFont = NewValue
End Property
Public Property Get CurrentFont() As winRTFFont
    Set CurrentFont = mobjCurrentFont
End Property

Public Sub EndDocument()

On Error GoTo ErrorCode

    CloseFile
    
ClearVariables:
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub
Public Property Get Fonts() As winRTFFonts
    Set Fonts = mobjFonts
End Property
Public Property Get InstalledFonts() As winRTFInstalledFonts
    Set InstalledFonts = gobjInstalledFonts
End Property
Private Sub KillFile()

On Error GoTo ErrorCode

    CloseFile
    If DoesFileExist(mstrTempFile) = True Then Kill mstrTempFile

    
ClearVariables:
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub
Public Sub NewParagraph(ByVal Alignment As winRTFAlignment, Optional ByVal LeftIndent = 0, Optional ByVal RightIndent = 0)

Dim strAlign As String
Dim lLeftIndent As Long
Dim lRightIndent As Long
Dim strTMP As String
Dim strIndent As String

    Select Case Alignment
        Case winRTFAlignment.winRTFAlignLeft
            strAlign = ALIGN_LEFT
        Case winRTFAlignment.winRTFAlignCenter
            strAlign = ALIGN_CENTER
        Case winRTFAlignment.winRTFAlignRight
            strAlign = ALIGN_RIGHT
        Case Else
            strAlign = ""
    End Select
    
    If LeftIndent < 0 Then LeftIndent = 0
    lLeftIndent = LeftIndent * 180
    
    If RightIndent < 0 Then RightIndent = 0
    lRightIndent = RightIndent * 180
    
    mlCurrentLeftIndent = LeftIndent
    mlCurrentRightIndent = RightIndent
    
    If mbolDocumentHasContent = True Then
        strTMP = LINE_BREAK
    Else
        strTMP = ""
    End If
    
    If lLeftIndent <> 0 Then
        strIndent = CStr(LEFT_INDENT & lLeftIndent)
    End If
    
    If lRightIndent <> 0 Then
        strIndent = strIndent & CStr(RIGHT_INDENT & lRightIndent)
    End If
    
    WriteToFile strTMP & NEW_PARAGRAPH & strIndent & strAlign
    
End Sub
Private Sub OpenFile()

On Error GoTo ErrorCode

    CloseFile
    
    miFileHandel = FreeFile
    Open mstrTempFile For Output As miFileHandel
    mbolFileIsOpen = True
    
ClearVariables:
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub

Public Function RTFCode() As String


Dim strRTF As String
Dim strRTFDocument As String
Dim bolNotDoneYet As String

    bolNotDoneYet = mbolFileIsOpen
    CloseFile

    strRTF = RTF_START
        strRTF = strRTF & CreateFontTable & vbCrLf
        strRTF = strRTF & CreateColorTable
        strRTFDocument = OpenFileAsString(mstrTempFile)
        strRTF = strRTF & strRTFDocument & LINE_BREAK
    strRTF = strRTF & RTF_END
    
    If bolNotDoneYet = True Then
        OpenFile
        WriteToFile strRTFDocument
    End If
    
    RTFCode = strRTF

End Function
Public Function SaveRTF(FileName As String) As Long

On Error GoTo ErrorCode
        
    SaveRTF = WriteStringAsFile(RTFCode, FileName)
    
ClearVariables:
    Exit Function
    
ErrorCode:
    SaveRTF = Err.Number
    GoTo ClearVariables
    
End Function
Private Sub WriteToFile(Text As String)
    If mbolFileIsOpen = False Then Exit Sub
    Print #miFileHandel, Text;
    mbolDocumentHasContent = True
End Sub
Private Sub Class_Initialize()

On Error GoTo ErrorCode

Dim objFont As winRTFFont
    
    mstrTempFile = TempFile("~tmp")
    
    Set gobjSystemFonts = New winRTFSystemFonts
        gobjSystemFonts.GetSystemFonts
        
    Set gobjFonts = New winRTFFonts
        Set mobjFonts = gobjFonts
        Set mobjCurrentFont = mobjFonts(1)
        mobjCurrentFont.Size = 10
    
    Set gobjInstalledFonts = New winRTFInstalledFonts
        For Each objFont In gobjSystemFonts
            gobjInstalledFonts.Add objFont
        Next
    
    Set gobjColors = New winRTFColors
        Set mobjColors = gobjColors
        Set mobjCurrentColor = mobjColors("black")
        
ClearVariables:
    Set objFont = Nothing
    Exit Sub
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables
    
End Sub
Private Sub Class_Terminate()

On Error Resume Next
    
    KillFile
    
    Set gobjInstalledFonts = Nothing
    
    Set mobjCurrentColor = Nothing
    Set mobjCurrentFont = Nothing
    
    Set mobjColors = Nothing
    Set gobjColors = Nothing
    
    Set gobjSystemFonts = Nothing
    
    Set mobjFonts = Nothing
    Set gobjFonts = Nothing
    
End Sub
