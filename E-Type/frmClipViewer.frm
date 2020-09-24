VERSION 5.00
Begin VB.Form frmClipViewer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3795
   ClientLeft      =   3285
   ClientTop       =   3780
   ClientWidth     =   4155
   Icon            =   "frmClipViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBorder 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   330
      TabIndex        =   8
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picBitmap 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   1200
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrCheckClip 
      Interval        =   1000
      Left            =   3360
      Top             =   2760
   End
   Begin VB.TextBox txtClipboardText 
      Height          =   1935
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "Palette"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "DIB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "Metafile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitmap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "Link"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmClipViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FormatValue() As Integer

' Initialize the format values.
Private Sub Form_Load()
    Call rotateText("E-Type ClipViewer", picBorder)
    ReDim FormatValue(0 To lblFormat.UBound)
    FormatValue(0) = vbCFLink
    FormatValue(1) = vbCFText
    FormatValue(2) = vbCFBitmap
    FormatValue(3) = vbCFMetafile
    FormatValue(4) = vbCFDIB
    FormatValue(5) = vbCFPalette
    GetClipData
End Sub


Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    wid = ScaleWidth - txtClipboardText.Left
    If wid < 120 Then wid = 120
    hgt = ScaleHeight
    
    txtClipboardText.Move _
        txtClipboardText.Left, 0, wid, hgt
End Sub

' See which formats are available.
Private Sub tmrCheckClip_Timer()
    GetClipData
End Sub


Public Sub GetClipData()
Dim i As Integer

    For i = 0 To lblFormat.UBound
        If Clipboard.GetFormat(FormatValue(i)) Then
            lblFormat(i).ForeColor = vbBlack
        Else
            lblFormat(i).ForeColor = &H808080
        End If
    Next i

    ' Check for text.
    If Clipboard.GetFormat(vbCFText) Then
        txtClipboardText.text = Clipboard.GetText
        txtClipboardText.Visible = True
    Else
        txtClipboardText.Visible = False
    End If

    ' Check for a bitmap.
    If Clipboard.GetFormat(vbCFBitmap) Then
        picBitmap.Picture = Clipboard.GetData(vbCFBitmap)
        picBitmap.Visible = True
    Else
        picBitmap.Visible = False
    End If

End Sub
