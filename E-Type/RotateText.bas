Attribute VB_Name = "modRotateText"
Option Explicit

' Just tested this in vb6.
' Rotates the text applied to a picturebox.
'
' Add a picture box to a form. For best looks make
' it flat borderless and align it to left or right.
' Remember to choose a truetype font in the picture box.
' The default MS Sans Serif will not work!
' Sett the font type, font size and collors to the picturebox
' In form load call it like this:
' Call RotateText("Text goes here", picture1)
'
' Trond Sørensen, trond.sorensen@bi.no


Public Declare Function CreateFontIndirect Lib "gdi32" Alias _
     "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc _
     As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal _
     hObject As Long) As Long
Public Const LF_FACESIZE = 32

Public Type LOGFONT
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
     lfFaceName As String * LF_FACESIZE
   End Type




Public Function RotateText(text As String, picturebox As picturebox)
    Dim font As LOGFONT
    Dim prevFont As Long, hFont As Long, ret As Long
    
    Const FONTSIZE = 10 ' Desired point size of font
    font.lfEscapement = 900    ' 180-degree rotation
    font.lfFaceName = picturebox.font & Chr$(0)  'Null character at end

    ' Windows expects the font size to be in pixels and to
    ' be negative if you are specifying the character height
    ' you want.
    font.lfHeight = (picturebox.FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(font)
    prevFont = SelectObject(picturebox.hdc, hFont)
    picturebox.CurrentX = picturebox.Left '+ picturebox.Width / 2
    picturebox.CurrentY = picturebox.ScaleHeight - 100 '/ 2
    picturebox.Print text
'    ' Clean up by restoring original font.
'    ret = SelectObject(Picture1.hdc, prevFont)
'    ret = DeleteObject(hFont)
'    Picture1.CurrentY = Picture1.ScaleHeight / 2
'    Picture1.Print "Normal Text"

End Function
