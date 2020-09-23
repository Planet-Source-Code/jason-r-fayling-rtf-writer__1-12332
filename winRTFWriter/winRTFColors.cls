VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"winRTFColor"
Attribute VB_Ext_KEY = "Member0" ,"winRTFColor"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
    Me.Add 0, "black"
End Sub


Public Function DoesExist(ColorName As String) As Boolean

On Error GoTo ErrorCode

Dim objTMP As Object

    DoesExist = False
        Set objTMP = mCol(ColorName)
    DoesExist = True

ClearVariables:
    Set objTMP = Nothing
    Exit Function
    
ErrorCode:
    DoesExist = False
    GoTo ClearVariables

End Function
Public Function Add(Value As Long, Optional ColorName As String) As winRTFColor

On Error GoTo ErrorCode

Dim objNewMember As winRTFColor

    Set objNewMember = New winRTFColor
    
    If Len(ColorName) <> 0 Then
        If DoesExist(ColorName) = True Then
            Set Add = Me(ColorName)
            GoTo ClearVariables
        End If
    End If

    objNewMember.Value = Value
    objNewMember.ColorName = ColorName
    objNewMember.ColorCode = CStr(mCol.Count + 1)
    
    If Len(ColorName) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, ColorName
    End If

    'return the object created
    Set Add = objNewMember
    
ClearVariables:
    Set objNewMember = Nothing
    Exit Function
    
ErrorCode:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Function
Public Property Get Item(vntIndexKey As Variant) As winRTFColor
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

    If TypeName(vntIndexKey) = "String" And vntIndexKey = "Automatic" Then Exit Sub
    If TypeName(vntIndexKey) = 0 Then Exit Sub
    
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
    'creates the collection when this class is created
    Set mCol = New Collection
    Me.Add 0, "black"
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

