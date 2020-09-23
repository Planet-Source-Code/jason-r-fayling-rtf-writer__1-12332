VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winRTFColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarValue As Long 'local copy
Private mvarColorName As String 'local copy
Private mvarColorCode As String 'local copy
Friend Property Let ColorCode(ByVal vData As String)
    mvarColorCode = vData
End Property
Friend Property Get ColorCode() As String
    ColorCode = mvarColorCode
End Property
Public Property Get Blue() As Integer
    Blue = GetBlueValue(mvarValue)
End Property
Public Property Get Green() As Integer
    Green = GetGreenValue(mvarValue)
End Property
Public Property Get Red() As Integer
    Red = GetRedValue(mvarValue)
End Property
Friend Property Let ColorName(ByVal vData As String)
    mvarColorName = vData
End Property
Public Property Get ColorName() As String
    ColorName = mvarColorName
End Property
Friend Property Let Value(ByVal vData As Long)
    mvarValue = vData
End Property
Public Property Get Value() As Long
    Value = mvarValue
End Property
