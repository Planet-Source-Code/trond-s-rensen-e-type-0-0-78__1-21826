VERSION 5.00
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2640
   ClientTop       =   2760
   ClientWidth     =   5895
   Icon            =   "Tip of the Day.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Tip of the Day.frx":0E42
         Top             =   120
         Width           =   480
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2800
         Left            =   -120
         Top             =   -120
         Width           =   855
      End
      Begin VB.Label lblDidYouKnow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         BorderWidth     =   3
         X1              =   720
         X2              =   7080
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sfile As String) As Boolean

    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sfile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sfile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sfile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()

    ' save whether or not this form should be displayed at startup
    SaveSetting App.Title, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value

End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    ' save  settings to registry
    SaveSetting App.Title, "StartUp", "ShowTip", chkLoadTipsAtStartup.Value
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim ShowAtStartup As Long
    
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.Title, "StartUp", "ShowTip", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
