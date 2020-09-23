VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Click Me!"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   10815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mRTFWrite As winRTFWrite.winRTFWriter
Private Sub Command2_Click()

Dim strText As String

     With Me.mRTFWrite
        ' make sure the object is cleared
        .Clear
        
        ' Use this to tell the object you wish to begin a new document
        .BeginDocument
        
            ' You should tell the object what fonts you will be using in
            ' the document
            If .InstalledFonts.IsInstalled("Times New Roman") = True Then
                .Fonts.Add "Times New Roman"
            Else
                .Fonts.Add .InstalledFonts(1).FontName
            End If
            
            ' You should also define any colors that you will be using
            ' black is already installed
            .Colors.Add vbRed, "red"
            .Colors.Add RGB(0, 128, 0), "green"
            .Colors.Add vbBlue, "blue"
                    
            ' Always begin each aligned section as a new paragraph
            ' this paragraph will be left aligned with no indentation
            .NewParagraph winRTFAlignLeft
                .AddText "This is an example of a left aligned paragraph"
                
                ' you can add a line break two ways
                
                ' 1st method
                .AddLineBreak
                
                ' 2nd method is when you have a block of text to add, just use the vbCRLF
                ' constant
                strText = "This is an example of a block of text" & vbCrLf & "added at one time!"
                .AddText strText
                
                ' to add a space between lines you need to call AddLineBreak twice
                .AddLineBreak
                .AddLineBreak
                
                ' Examples on how to change the line color
                .AddText "This is a RED line.", , .Colors("red")
                
                ' the above example uses a color you defined eariler in the code
                ' the below example shows how to use an undefined color
                .AddLineBreak
                .AddText "This is a Brown line.", , .Colors.Add(RGB(129, 128, 0), "olive")
                
                ' Example on how to switch fonts
                ' first check to see if the font is installed on the system
                ' and add it to the font collection
                If .InstalledFonts.IsInstalled("Verdana") = True Then
                    .Fonts.Add "Verdana"
                    
                    ' Set the current font to Verdana
                    Set .CurrentFont = .Fonts("Verdana")
                End If
                
                .AddLineBreak
                .AddLineBreak
                strText = "This line should be the " & .CurrentFont.FontName & " font"
                .AddText strText
                
                ' reset the font back to the default
                Set .CurrentFont = .Fonts(1)
                
                .AddLineBreak
                .AddLineBreak
                
                ' If you wish to use a font for only a shot period of time you
                ' can supply in in the AddText Method

                If .InstalledFonts.IsInstalled("System") = True Then
                    ' check to see if the System font has been added to our font collection
                    If .Fonts.DoesExist("System") = False Then .Fonts.Add "System"
                    .AddText "This line should be the System font.", .Fonts("System")
                Else
                    .AddText "Could not display this line using the System font", .Colors("red")
                End If
                
                ' notice that that next line does not do anything with the font
                ' object, and the display should be Time New Roman
                .AddLineBreak
                .AddLineBreak
                .AddText "This line is " & .CurrentFont.FontName, , .Colors("green")
                
                ' Example on how to change the font properties
                With .CurrentFont
                    .Bold = True
                    .Itialic = True
                    .Underline = True
                    .Size = 30
                End With
                
                .AddLineBreak
                .AddLineBreak
                .AddText "Example on how to change the font properties"
                
                ' this reset all the properties but the size
                .CurrentFont.Reset
                .CurrentFont.Size = 12
                
                ' Center alignment example
                Set .CurrentColor = .Colors("green")
                .AddLineBreak
                .NewParagraph winRTFAlignCenter
                    strText = "Now is the time for all good men to" & vbCrLf & "come to the aid" & vbCrLf & "of thier contry."
                    .AddText strText
                
                ' Right alignment example
                Set .CurrentColor = .Colors("red")
                .AddLineBreak
                .NewParagraph winRTFAlignRight
                    .AddText strText
                    
                
                ' Indentation example
                ' Setting the Left Indent to 4 is equal to .5 inches
                Set .CurrentColor = .Colors(1)
                .AddLineBreak
                .NewParagraph winRTFAlignLeft
                    .AddText "This is a 0 inch indent"
                    .NewParagraph winRTFAlignLeft, 4
                        .AddText "This is a .5 inch indent"
                        .AddLineBreak
                        .AddText "blah, blah, blah"
                        .NewParagraph winRTFAlignLeft, 8
                            .AddText "This is a 1 inch indent" & vbCrLf
                            .AddText "Blah, Blah , Blah"
                
                .AddLineBreak
                
                'Bullet Example
                .NewParagraph winRTFAlignLeft
                    .AddBullet "This is bullet 1", , .Colors("red")
                    .NewParagraph winRTFAlignLeft, 8
                        .AddBullet "This is bullet 2", , .Colors("blue")
                        .NewParagraph winRTFAlignLeft, 16
                            .AddBullet "This is bullet 3", , .Colors("green")
                            
                
                .NewParagraph winRTFAlignLeft
                .AddLineBreak
                .AddLineBreak
                
                
                ' This is an advanced example of what this object can do!
                Dim objFont As winRTFFont
                Dim i As Integer
                Dim j As Integer
                
                    j = 20
                    If .InstalledFonts.Count < j Then j = .InstalledFonts.Count
                    
                    For i = 1 To j
                        Set objFont = .InstalledFonts(i)
                        
                            If .Fonts.DoesExist(objFont.FontName) = False Then
                                .Fonts.Add objFont.FontName
                            End If
                            
                            objFont.Size = 10
                            
                            strText = "This is the "
                            .AddText strText, objFont
                            
                            objFont.Reset
                            objFont.Bold = True
                            objFont.Underline = True
                            objFont.Size = 14
                            
                            .AddText objFont.FontName & " ", , .Colors("red")
                            
                            objFont.Reset
                            objFont.Size = 10
                            
                            .AddText "font.", objFont
                            .AddLineBreak
                                                   
                    Next
                    
        ' You are done with the document
        .EndDocument
        
        ' this is the raw RTF code you just created!
        MsgBox .RTFCode
        
        ' Save the new document
        If .SaveRTF(App.Path & "\RTF Test1.rtf") = True Then
            ' load the document into the viewer
            Me.RichTextBox1.LoadFile App.Path & "\RTF Test1.rtf"
        End If
        
     End With

End Sub

Private Sub Form_Load()
    Set mRTFWrite = New winRTFWriter
End Sub


