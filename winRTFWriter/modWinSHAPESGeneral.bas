Attribute VB_Name = "modWinSHAPESGeneral"
Option Explicit

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Const LVM_FIRST = &H1000
    Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
    Public Const LVM_SETITEMSTATE = LVM_FIRST + 43
    Public Const LVM_GETITEMSTATE = LVM_FIRST + 44
    Public Const LVIS_STATEIMAGEMASK = &HF000
    Public Const LVM_GETITEM = LVM_FIRST + 5 '75 for unicode?
    Public Const LVIF_STATE = &H8
    
    Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
    Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
    Public Const LVS_EX_FULLROWSELECT = &H20
    Public Const WM_SETREDRAW = &HB
    Public Const LVS_EX_GRIDLINES = &H1
    Public Const LVS_EX_SUBITEMIMAGES = &H2
    Public Const LVS_EX_CHECKBOXES = &H4
    Public Const LVS_EX_TRACKSELECT = &H8
    Public Const LVS_EX_HEADERDRAGDROP = &H10
    
    Public Const LVSCW_AUTOSIZE = -1
    Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCAND = &H8800C6
    Public Const SRCCOPY = &HCC0020
    Public Const SRCPAINT = &HEE0086

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_SETREADONLY = &HCF

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Const SW_SHOWNORMAL = 1
    Private Const SW_MAXIMIZE = 3
    Private Const SW_MINIMIZE = 6


' Module      : modForms
' Description : Routines for working with VB forms
' Source      : Total VB SourceBook 6
'
Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) _
  As Long

Private Declare Function SetWindowPos _
  Lib "user32" _
  (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) _
  As Long

Private Declare Function GetSystemMenu _
  Lib "user32" _
  (ByVal hwnd As Long, _
    ByVal bRevert As Long) _
  As Long

Private Declare Function ModifyMenu _
  Lib "user32" _
  Alias "ModifyMenuA" _
  (ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpString As Any) _
  As Long

Private Declare Function GetMenuItemID _
  Lib "user32" _
  (ByVal hMenu As Long, _
    ByVal nPos As Long) _
  As Long

Private Const WM_SYSCOMMAND = &H112
Private Const MOUSE_MOVE = &HF012
Private Const WM_LBUTTONUP = &H202

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type POINTS
  X  As Integer
  Y  As Integer
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = 1
Private Const SWP_NOMOVE = 2
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_GRAYED = &H1&
Private Const SC_CLOSE = &HF060
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

Private Declare Function GetTempPath _
  Lib "kernel32" _
  Alias "GetTempPathA" _
  (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) _
  As Long
  
Private Declare Function GetTempFileName _
  Lib "kernel32" _
  Alias "GetTempFileNameA" _
  (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) _
  As Long

Public Type GUID ' a structure For Global Uniq. ID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0, 7) As Byte
End Type
    Declare Function CoCreateGuid Lib "ole32" (ByRef lpGUID As GUID) As Long
    Declare Function StringFromGUID2 Lib "ole32" (ByRef lpGUID As GUID, ByVal lpStr As String, ByVal lSize As Long) As Long

            
Public Function GetNewGUIDStr() As String
    Dim pGuid As GUID
    Dim lResult As Long
    Dim s As String
    'this is a buffer string to be passed in
    '     API function
    '100 chars will be enough
    s = String(100, " ")
    'creating new ID and obtaining result in
    '     pointer to GUID
    lResult = CoCreateGuid(pGuid)
    'converting GUID structure to string
    lResult = StringFromGUID2(pGuid, s, 100)
    'removing all trailing blanks
    s = Trim(s)
    'converting a sting from unicode
    GetNewGUIDStr = StrConv(s, vbFromUnicode)
End Function
Public Function SaveResItemToDisk( _
            ByVal iResourceNum As Integer, _
            ByVal sResourceType As String, _
            ByVal sDestFileName As String _
            ) As Long
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    
    On Error GoTo SaveResItemToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        
        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return 0 for success
    SaveResItemToDisk = 0
    
    Exit Function
SaveResItemToDisk_err:
    Close #iFileNumOut
    'Return error number
    SaveResItemToDisk = Err.Number
End Function

Private Function TrimNulls(ByVal strIn As String) As String
  ' Comments  : Returns the passed string terminated at the first null
  ' Parameters: strIn - Value to parse
  ' Returns   : Parsed string
  ' Source    : Total VB SourceBook 6
  '
  Dim intPos As Integer
  
  On Error GoTo PROC_ERR
    
  intPos = InStr(strIn, vbNullChar)
  
  If intPos = 0 Then
    ' No nulls in the string, just return it as is
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      TrimNulls = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      TrimNulls = Left$(strIn, intPos - 1)
    End If
  End If
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Resume PROC_EXIT
    
End Function
Public Sub CenterForm(frmIn As Form)
  ' Comments  : Centers the form on the screen
  ' Parameters: frmIn - form to center on the screen
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  frmIn.Move (Screen.Width - frmIn.Width) / 2, _
    (Screen.Height - frmIn.Height) / 2

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CenterForm"
  Resume PROC_EXIT

End Sub

Public Sub DisableCloseMenu(frmIn As Form)
  ' Comments  : Grays out the Close item on the form's
  '             system menu
  ' Parameters: frmIn - form to modify
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim lnghMenu As Long
  Dim lnghItem As Long

  On Error GoTo PROC_ERR

  ' get handle to form's system menu
  lnghMenu = GetSystemMenu(frmIn.hwnd, 0)
  
  ' get handle to the 6th item (Close)
  lnghItem = GetMenuItemID(lnghMenu, 6)
  
  ' gray out this item
  lngResult = ModifyMenu( _
    lnghMenu, _
    lnghItem, _
    MF_BYCOMMAND Or MF_GRAYED, _
    -10, _
    "Close")

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "DisableCloseMenu"
  Resume PROC_EXIT

End Sub

Public Function DoesFileExist(FileName As String) As Boolean

On Error GoTo ErrorCode
    
    DoesFileExist = False
        If Len(FileName) = 0 Then GoTo ClearVariables
        If Len(Dir(FileName, vbNormal)) = 0 Then GoTo ClearVariables
    DoesFileExist = True

ClearVariables:
    Exit Function
ErrorCode:
    DoesFileExist = False
    GoTo ClearVariables

End Function
Public Sub FormOnTop( _
  frmIn As Form, _
  ByVal fOnTop As Boolean)
  ' Comments  : Sets the form's style to be always on
  '             top, or to remove the always on top style
  ' Parameters: frmIn - the form to modify
  '             fOnTop - true to set the form to be always
  '             on top of other windows. Set to False to
  '             remove this attribute
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngState As Long
  
  On Error GoTo PROC_ERR

  If fOnTop Then
    lngState = HWND_TOPMOST
  Else
    lngState = HWND_NOTOPMOST
  End If
  
  SetWindowPos frmIn.hwnd, lngState, 0&, 0&, 0&, 0&, _
    SWP_NOSIZE Or SWP_NOMOVE

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FormOnTop"
  Resume PROC_EXIT

End Sub

Public Function GetFormNamed(ByVal strForm As String) As Form
  ' Comments  : Retrieves a reference to form based on a string name
  ' Parameters: strForm - the name of the form to locate
  ' Returns   : A reference to the form if found, otherwise Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  Dim frm As Form
  Dim fFound As Boolean
  
  strForm = UCase(strForm)
  
  ' Search in the forms collection to see if a form with the selected name
  ' is found
  For Each frm In Forms
    If UCase(frm.Name) = strForm Then
      fFound = True
      Set GetFormNamed = frm
      Exit For
    End If
  Next frm
  
  If Not fFound Then
    Set GetFormNamed = Nothing
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetFormNamed"
  Resume PROC_EXIT

End Function

Public Function TempFile(strPrefix As String) As String

  ' Returns    : A temporary file name based on the value of strPrefix.
  ' Source: Total VB SourceBook 6
  '
  Dim strTemp As String
  Dim lngRet As Long
  Dim strTempPath As String
    
  On Error GoTo ErrorCode
  
  strTempPath = Space$(255)
  lngRet = GetTempPath(Len(strTempPath), strTempPath)
  
  strTemp = Space$(255)
  lngRet = GetTempFileName(strTempPath, strPrefix, 0, ByVal strTemp)
  
  TempFile = TrimNulls(strTemp)
  
  If DoesFileExist(TempFile) = True Then Kill TempFile
  
ClearVariables:
  Exit Function
  
ErrorCode:
  GoTo ClearVariables

End Function
Public Function LoadFormByName(strForm As String) As Form
  ' Comments  : Loads a form by using a string variable containing
  '             the name of the form to avoid hard-coding form references
  ' Parameters: strForm - The name of the form to load
  ' Returns   : A pointer to the form that was loaded
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  On Error Resume Next
  Forms.Add strForm
  If Err.Number = 0 Then
    Set LoadFormByName = Forms(Forms.Count - 1)
  Else
    Set LoadFormByName = Nothing
  End If

  On Error GoTo PROC_ERR
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoadFormByName"
  Resume PROC_EXIT

End Function

Public Sub MakeTransparent(frmIn As Form)
  ' Comments  : Sets the form's style to be transparent. This call should
  '             be made before the form is shown, for example in the Load
  '             event
  ' Parameters: frmIn - form to modify
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  lngResult = SetWindowLong(frmIn.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MakeTransparent"
  Resume PROC_EXIT

End Sub


Public Function OpenFileAsString(FileName As String) As String

On Error GoTo ErrorCode

Dim iNextFree As Integer
Dim bolDidOpen As Boolean

    If DoesFileExist(FileName) = False Then GoTo ClearVariables
    iNextFree = FreeFile
    
    Open FileName For Binary As iNextFree
        bolDidOpen = True
        OpenFileAsString = Input(FileLen(FileName), #iNextFree)

ClearVariables:
    If bolDidOpen = True Then Close iNextFree
    Exit Function
    
ErrorCode:
    OpenFileAsString = ""
    GoTo ClearVariables

End Function
Public Sub OpenWebPage(FormName As Form, URL As String, Optional WindowSize As Integer = SW_SHOWNORMAL)
'
' API's
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal 'lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, 'ByVal nShowCmd As Long) As Long
'
'Private Const SW_SHOWNORMAL = 1

Dim iret As Long
    
    If Len(URL) = 0 Then Exit Sub
    DoEvents
    iret = ShellExecute(FormName.hwnd, vbNullString, URL, vbNullString, "c:\", WindowSize)

End Sub
Public Function Quote(strData As String) As String
    Quote = "'" & strData & "'"
End Function
Public Sub FloodUpdate(objPercent As PictureBox, Progress As Long, Msg As String, upperLimit As Long, Optional FillDirection As Long = 1, Optional lBackColor As Long = &HFFFFFF, Optional lForeColor As Long = &H0)
    
On Error Resume Next

Dim r As Long

    If objPercent Is Nothing = True Then Exit Sub
    
    objPercent.BackColor = lBackColor
    objPercent.ForeColor = lForeColor
    objPercent.AutoRedraw = True
    objPercent.DrawMode = 10
       
   'solid fill
    objPercent.FillStyle = 0
    

    
   'make sure that the flood display hasn't already hit 100%
    If Progress <= upperLimit Then
     'error trap in case the code below attempts to set
     'the scalewidth greater than the max allowable
     
     Select Case FillDirection
     
        Case 1
            If Progress > objPercent.ScaleWidth Then
                Progress = objPercent.ScaleWidth
                
            End If
            
            'erase the flood
            objPercent.Cls
    
            objPercent.ScaleWidth = upperLimit
                        
            'calculate the string's X & Y coordinates
            'in the PictureBox ... here, centered
            objPercent.CurrentX = (objPercent.ScaleWidth - objPercent.TextWidth(Msg)) \ 2
            objPercent.CurrentY = (objPercent.ScaleHeight - objPercent.TextHeight(Msg)) \ 2
            
            'print the percentage string in the text colour
            objPercent.Print Msg
            
            'print the flood bar to the new progress length in the line colour
            objPercent.Line (0, 0)-(Progress, objPercent.ScaleHeight), objPercent.ForeColor, BF
            
              
            'call BitBlit to invert that portion of the text over the bar
            r = BitBlt(objPercent.hdc, 0, 0, _
                 objPercent.ScaleWidth, objPercent.ScaleHeight, _
                 objPercent.hdc, 0, 0, SRCCOPY)
     
    
        Case Else
            
            If Progress > objPercent.ScaleHeight Then
                Progress = objPercent.ScaleHeight
                
            End If
            
            'erase the flood
            objPercent.Cls
            
            objPercent.ScaleHeight = upperLimit
            
            'calculate the string's X & Y coordinates
            'in the PictureBox ... here, centered
            objPercent.CurrentX = (objPercent.ScaleWidth - objPercent.TextWidth(Msg)) \ 2
            objPercent.CurrentY = (objPercent.ScaleHeight - objPercent.TextHeight(Msg)) \ 2
            
            'print the percentage string in the text colour
            objPercent.Print Msg
            
            'print the flood bar to the new progress length in the line colour
            objPercent.Line (0, upperLimit)-(objPercent.ScaleWidth, upperLimit - Progress), objPercent.ForeColor, BF
            
              
            'call BitBlit to invert that portion of the text over the bar
            r = BitBlt(objPercent.hdc, 0, 0, _
                 objPercent.ScaleWidth, objPercent.ScaleHeight, _
                 objPercent.hdc, 0, 0, SRCCOPY)
            
            
     End Select

    
    'allow the flood to complete drawing
      DoEvents
        
    
    End If
        
End Sub

Public Function isFormLoaded(FormName As String) As Boolean

Dim objForm As Form

    isFormLoaded = False
    For Each objForm In Forms
        If LCase(objForm.Name) = LCase(FormName) Then
            isFormLoaded = True
            Exit Function
        End If
    Next

End Function
Public Sub AutoSizeColumns(m_ListView As Object)
  ' Comments  : Sizes each column in the listview control to fit
  '             the widest data in each column
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim intColumn As Integer
  
  On Error GoTo PROC_ERR
  
  For intColumn = 0 To m_ListView.ColumnHeaders.Count - 1
    SendMessageLong _
      m_ListView.hwnd, _
      LVM_SETCOLUMNWIDTH, _
      intColumn, _
      LVSCW_AUTOSIZE_USEHEADER
  Next intColumn

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AutoSizeColumns"
  Resume PROC_EXIT

End Sub

Public Function SQLSafe(strData As String) As String
    SQLSafe = Replace(strData, "'", "''")
End Function


Public Function WriteStringAsFile(Buffer As String, FileName As String) As Boolean

On Error GoTo ErrorCode

Dim iNextFree As Integer
Dim bolDidOpen As Boolean

    WriteStringAsFile = False
    iNextFree = FreeFile
    
    Open FileName For Binary As iNextFree
        bolDidOpen = True
        Put #iNextFree, , Buffer
    WriteStringAsFile = True
    
ClearVariables:
    If bolDidOpen = True Then Close iNextFree
    Exit Function
    
ErrorCode:
    WriteStringAsFile = False
    GoTo ClearVariables

End Function
