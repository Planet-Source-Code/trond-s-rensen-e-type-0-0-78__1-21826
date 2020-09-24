VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFileInfo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   2160
   ClientTop       =   2175
   ClientWidth     =   6255
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   6255
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
      Height          =   5520
      Left            =   0
      ScaleHeight     =   5520
      ScaleWidth      =   330
      TabIndex        =   27
      Top             =   0
      Width           =   330
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Document"
      TabPicture(0)   =   "frmFileInfo.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFilename"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fradates"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraattrib"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraStatistics"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fraStatistics 
         Caption         =   "Statistics:"
         Height          =   1455
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   5175
         Begin VB.Label lblLines1 
            Caption         =   "Lines:"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblLines 
            Caption         =   "Lines"
            Height          =   255
            Left            =   1680
            TabIndex        =   25
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblWords1 
            Caption         =   "Words:"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblWords 
            Caption         =   "Characters"
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblCharacters1 
            Caption         =   "Characters:"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblCharacters 
            Caption         =   "Characters"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblsize1 
            Caption         =   "Size:"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblSize 
            Caption         =   "FileSize"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraattrib 
         Caption         =   "Attributes:"
         Height          =   2175
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
         Begin VB.CheckBox attributes 
            Caption         =   "Hidden"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "System"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "Read Only"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "Archive"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "Temporary"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "Compressed"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   11
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox attributes 
            Caption         =   "Normal"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   10
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.Frame fradates 
         Caption         =   "Dates:"
         Height          =   2175
         Left            =   2040
         TabIndex        =   2
         Top             =   2400
         Width           =   3375
         Begin VB.Label lbldate 
            Caption         =   "00.00.0000 00:00:00"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   8
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbldatetxt 
            Caption         =   "Last Accessed:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbldate 
            Caption         =   "00.00.0000 00:00:00"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lbldatetxt 
            Caption         =   "Last Modified:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbldatetxt 
            Caption         =   "Created:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lbldate 
            Caption         =   "00.00.0000 00:00:00"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Label lblFilename 
         Caption         =   "Filename"
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   4965
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub updatestats()

    'Dim ftime As SYSTEMTIME               ' Initialise variables
    'Dim tfilename As String
    'tfilename = dialog.Filename

    tfilename = frmMDI.ActiveForm.Caption

    Dim filedata As WIN32_FIND_DATA

    filedata = Findfile(tfilename)      ' Get information

    lblFilename.Caption = tfilename     ' Put name in text box
    lblFilename.ToolTipText = "Full Filename: " & tfilename ' Put filename in tooltip

    ' updating the linenumbers info
    lblLines.Caption = getLineInfo(frmMDI.ActiveForm.Text1, 0)


    If filedata.nFileSizeHigh = 0 Then  ' Put size into text box
      lblSize.Caption = filedata.nFileSizeLow & " Bytes"
    Else
      lblSize.Caption = filedata.nFileSizeHigh & "Bytes"
    End If

    ' Do not change the order on the next 6 lines!!
    Call FileTimeToSystemTime(filedata.ftCreationTime, ftime)   ' Determine Creation date and time, then format it
    If ftime.wDay = "1" And ftime.wMonth = "1" And ftime.wYear = "1601" Then
        lbldate(0) = Format(Now, "d/m/yyyy h:mm:ss ")
    Else
        lbldate(0) = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
    End If
    Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime)  ' Determine Last Modified date and time
    If ftime.wDay = "1" And ftime.wMonth = "1" And ftime.wYear = "1601" Then
        ' nothing
    Else
        lbldate(1) = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
    End If
    Call FileTimeToSystemTime(filedata.ftLastAccessTime, ftime) ' Determine Last accessed date (note no time is recorded)
    If ftime.wDay = "1" And ftime.wMonth = "1" And ftime.wYear = "1601" Then
        ' nothing
    Else
        lbldate(2) = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear
    End If
    
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
      attributes(1).Value = 1 ' Determine if file has hidden attribute set
    Else
      attributes(1).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM Then
      attributes(2).Value = 1 ' Determine if file has system attribute set
    Else
      attributes(2).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
      attributes(3).Value = 1 ' Determine if file has readonly attribute set
    Else
      attributes(3).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE Then
      attributes(4).Value = 1 ' Determine if file has archive attribute set
    Else
      attributes(4).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY Then
      attributes(5).Value = 1 ' Determine if file has temporary attribute set
    Else
      attributes(5).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL Then
      attributes(6).Value = 1 ' Determine if file has normal attribute set
    Else
      attributes(6).Value = 0
    End If
    If (filedata.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED Then
      attributes(7).Value = 1 ' Determine if file has compressed attribute set
    Else
      attributes(7).Value = 0
    End If
End Sub

Private Sub attributes_GotFocus(index As Integer)
    cmdOK.SetFocus     ' Make it impossible to change check boxes
End Sub

Private Sub cmdOK_Click()
     Unload Me
End Sub

Private Sub Form_Load()
    Call rotateText("E-Type File Info", picBorder)

    MakeTopMost hwnd
    ' first check if there is any documents open
    ' if not, unload frmInfo
    If (Forms.Count > 2) Then
        On Error GoTo errhand
        updatestats     ' Update infomation on form
        Me.Show         ' Show the form
        cmdOK.SetFocus  ' Set focus to Done button
    Else
        MsgBox "No Open Documents", vbInformation
        Unload Me
    End If
    Exit Sub
    
errhand:
    If Err.Number = cdlCancel Then
        Call MsgBox("You Pressed Cancel!", vbExclamation)   ' Cancel was pressed in common dialog box
    Else
        Call MsgBox("An error has occured!", vbExclamation) ' An unexpected eror has occured
    End If
    End
    
End Sub

