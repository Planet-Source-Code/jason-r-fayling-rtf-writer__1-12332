VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mobjFont As StdFont
Private mstrFontCode As String
Private mstrFontFamily As String
Private mlPitch As Long
Friend Property Get FontFamilyCode() As String
    Select Case LCase(mstrFontFamily)
        Case "decorative"
            FontFamilyCode = "fdecor"
        Case "dont care"
            FontFamilyCode = "fnil"
        Case "modern"
            FontFamilyCode = "fmodern"
        Case "roman"
            FontFamilyCode = "froman"
        Case "script"
            FontFamilyCode = "fscript"
        Case "swiss"
            FontFamilyCode = "fswiss"
    End Select
End Property
Public Property Let Itialic(ByVal vData As Boolean)
    mobjFont.Italic = vData
End Property
Public Property Get Itialic() As Boolean
    Itialic = mobjFont.Italic
End Property
Public Sub Reset()
    Me.Bold = False
    Me.Itialic = False
    Me.Underline = False
End Sub
Public Property Let Size(ByVal vData As Double)
    mobjFont.Size = vData
End Property
Public Property Get Size() As Double
    Size = mobjFont.Size
End Property
Friend Property Let FontName(ByVal vData As String)
    mobjFont.Name = vData
End Property
Public Property Get FontName() As String
    FontName = mobjFont.Name
End Property
Public Property Let Underline(ByVal vData As Boolean)
    mobjFont.Underline = vData
End Property
Public Property Get Underline() As Boolean
    Underline = mobjFont.Underline
End Property
Public Property Let Bold(ByVal vData As Boolean)
    mobjFont.Bold = vData
End Property
Public Property Get Bold() As Boolean
    Bold = mobjFont.Bold
End Property
Private Sub Class_Initialize()
    Set mobjFont = New StdFont
End Sub
Private Sub Class_Terminate()
    Set mobjFont = Nothing
End Sub
Public Property Get Strikethrough() As Boolean
    Strikethrough = mobjFont.Strikethrough
End Property
Public Property Let Strikethrough(ByVal vNewValue As Boolean)
    mobjFont.Strikethrough = vNewValue
End Property
Public Property Get Weight() As Integer
    Weight = mobjFont.Weight
End Property
Public Property Let Weight(NewValue As Integer)
    mobjFont.Weight = NewValue
End Property
Public Property Get Charset() As Integer
    Charset = mobjFont.Charset
End Property
Public Property Let Charset(ByVal vNewValue As Integer)
    mobjFont.Charset = vNewValue
End Property
Friend Property Get FontCode() As String
    FontCode = mstrFontCode
End Property
Friend Property Let FontCode(ByVal vNewValue As String)
    mstrFontCode = vNewValue
End Property
Public Property Get FontFamily() As String
    FontFamily = mstrFontFamily
End Property
Friend Property Let FontFamily(ByVal vNewValue As String)
    mstrFontFamily = vNewValue
End Property

Public Property Get Pitch() As Long
    Pitch = mlPitch
End Property
Friend Property Let Pitch(ByVal vNewValue As Long)
    mlPitch = vNewValue
End Property
