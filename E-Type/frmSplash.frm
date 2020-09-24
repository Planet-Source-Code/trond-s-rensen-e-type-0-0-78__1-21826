VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1950
   ClientLeft      =   3330
   ClientTop       =   3120
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer trUnloadSplash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   720
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   1320
      Picture         =   "frmSplash.frx":0E42
      ScaleHeight     =   1020
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   240
      Width           =   3030
   End
   Begin VB.Timer trload 
      Interval        =   500
      Left            =   4920
      Top             =   240
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   2160
      TabIndex        =   1
      Tag             =   "Version"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1560
      Left            =   -2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Show
    MakeTopMost hwnd
End Sub

Private Sub trload_Timer()
    ' load the main form
    Load frmMDI
End Sub

Private Sub trUnloadSplash_Timer()
    Unload Me
    ' The RebuildWinList sub will count the splash screen as a form and
    ' enable the select window toolbarbutton. Therefore we must rebuild
    ' the windowlist after the splash screen unloads.
    
    ' Call RebuildWinList
    frmMDI.ChildStatusChanged

End Sub
