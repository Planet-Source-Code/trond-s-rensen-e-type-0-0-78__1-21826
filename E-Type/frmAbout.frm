VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   3525
   ClientTop       =   3600
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "frmMDI"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Tag             =   "About Project1"
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
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   600
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1020
      ScaleWidth      =   3030
      TabIndex        =   1
      Top             =   360
      Width           =   3060
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "OK"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":A1D0
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   2880
      Width           =   4770
   End
   Begin VB.Label lblMailTo 
      AutoSize        =   -1  'True
      Caption         =   "trond.sorensen@bi.no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      MouseIcon       =   "frmAbout.frx":A264
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3600
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "All Rights Reserved."
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   840
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 2000 Trond Sørensen."
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   2475
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "This program is Freeware"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   480
      X2              =   5800
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   5800
      Y1              =   2805
      Y2              =   2805
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      Height          =   1215
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      Height          =   375
      Left            =   720
      Shape           =   1  'Square
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call rotateText("About E-Type", picBorder)
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    lblDisplay.Caption = "The author accepts no responsibility for any damage "
    lblDisplay.Caption = lblDisplay.Caption & Chr(13) & "or data loss caused by this program."
End Sub

Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub lblMailTo_Click()
    'To start default email program
    Dim ret
    ret = Shell("start mailto:" & "trond.sorensen@bi.no" & "?subject=" & "E-Type" & "", 0)
End Sub
