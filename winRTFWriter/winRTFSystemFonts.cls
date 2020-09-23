VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFSystemFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IEnumFontSink

Dim mCol As Collection
Public Function DoesExist(FontName As String) As Boolean

On Error GoTo ErrorCode

Dim objTMP As Object

    DoesExist = False
        Set objTMP = mCol(FontName)
    DoesExist = True

ClearVariables:
    Set objTMP = Nothing
    Exit Function
    
ErrorCode:
    DoesExist = False
    GoTo ClearVariables

End Function

Public Sub GetSystemFonts()

Dim lhDc As Long

    lhDc = GetDC(GetDesktopWindow())
        GetFonts lhDc, Me, True, , DEFAULT_CHARSET
    ReleaseDC GetDesktopWindow, lhDc


End Sub


Public Property Get Item(vntIndexKey As Variant) As winRTFFont
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Friend Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Private Sub IEnumFontSink_AddFont(ByVal sFaceName As String, ByVal sFontFamily As String, ByVal lPitch As Long, ByVal sStyle As String, ByVal sCharset As String, ByVal lCharset As Long, ByVal bPrinterFont As Boolean)
    DoEvents
    Me.Add sFaceName, sFontFamily, lPitch
End Sub
Friend Function Add(FontName As String, FontFamily As String, Pitch As Long, Optional Size As Double = 10, Optional Bold As Boolean = False, Optional Underline As Boolean = False, Optional Itialic As Boolean = False) As winRTFFont

On Error GoTo ErrorCode

Dim objNewMember As winRTFFont
Dim objFont As StdFont

    Set objNewMember = New winRTFFont
    
    If Me.DoesExist(FontName) = True Then
        Set Add = Me(FontName)
        GoTo ClearVariables
    End If
    
    Set objFont = New StdFont
    objFont.Name = FontName
    If LCase(objFont.Name) <> LCase(FontName) Then GoTo ClearVariables

    objNewMember.FontName = FontName
    objNewMember.FontFamily = FontFamily
    objNewMember.Pitch = Pitch
    objNewMember.Bold = Bold
    objNewMember.Underline = Underline
    objNewMember.Size = Size
    objNewMember.Itialic = Itialic
    objNewMember.FontCode = "f" & mCol.Count
    
    mCol.Add objNewMember, FontName
    
    'return the object created
    Set Add = objNewMember

ClearVariables:
    Set objFont = Nothing
    Set objNewMember = Nothing
    
    Exit Function
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Function

Private Property Get IEnumFontSink_HasFont(ByVal sFaceName As String) As Boolean
    IEnumFontSink_HasFont = DoesExist(sFaceName)
End Property
