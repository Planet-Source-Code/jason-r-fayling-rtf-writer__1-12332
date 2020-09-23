VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"winRTFFont"
Attribute VB_Ext_KEY = "Member0" ,"winRTFFont"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(FontName As String, Optional Size As Double = 10, Optional Bold As Boolean = False, Optional Underline As Boolean = False, Optional Itialic As Boolean = False) As winRTFFont

On Error GoTo ErrorCode

Dim objNewMember As winRTFFont
Dim objFont As winRTFFont

    Set objNewMember = New winRTFFont

    If DoesExist(FontName) = True Then
        Set Add = mCol(FontName)
        GoTo ClearVariables
    End If

    If gobjSystemFonts.DoesExist(FontName) = False Then GoTo ClearVariables
    Set objFont = gobjSystemFonts(FontName)

    objNewMember.FontName = FontName
    objNewMember.FontFamily = objFont.FontFamily
    objNewMember.Pitch = objFont.Pitch
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
Public Sub Clear()

    Set mCol = Nothing
    Set mCol = New Collection
    
    If gobjSystemFonts.DoesExist("Times New Roman") = True Then
        Me.Add "Times New Roman"
    Else
        Me.Add gobjSystemFonts(1).FontName
    End If
    
End Sub
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
Public Property Get Item(vntIndexKey As Variant) As winRTFFont
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
    If gobjSystemFonts.DoesExist("Times New Roman") = True Then
        Me.Add "Times New Roman"
    Else
        Me.Add gobjSystemFonts(1).FontName
    End If
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

