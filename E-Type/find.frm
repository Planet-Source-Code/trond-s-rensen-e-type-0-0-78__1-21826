VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   5730
   ClientTop       =   2775
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "find.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
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
      Height          =   3090
      Left            =   0
      ScaleHeight     =   3090
      ScaleWidth      =   330
      TabIndex        =   30
      Top             =   0
      Width           =   330
   End
   Begin VB.ComboBox cboFindInFiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "find.frx":0E42
      Left            =   2040
      List            =   "find.frx":0E44
      TabIndex        =   17
      Text            =   "Find what"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindFiles 
      Caption         =   "Find Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      ToolTipText     =   "Start Search"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      ToolTipText     =   "Start Search"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cboReplace 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CheckBox chkSearchInSubDirs 
      Caption         =   "Search Sub Directories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      ToolTipText     =   "Case Sensitivity"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CheckBox chkListDirs 
      Caption         =   "List &Dirs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   26
      ToolTipText     =   "Case Sensitivity"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Whole &Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Case Sensitivity"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cboInFolder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   18
      Text            =   "C:\"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraDirections 
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
      Begin VB.OptionButton optDirection 
         Caption         =   "&Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Search to Beginning of Document"
         Top             =   240
         Width           =   612
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "&Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   9
         ToolTipText     =   "Search to End of Document"
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdFindInFiles 
      Caption         =   "Find &in Files"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub Directories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   22
      ToolTipText     =   "Case Sensitivity"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &Case"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   13
      ToolTipText     =   "Case Sensitivity"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnMark 
      Caption         =   "&Unmark All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdMark 
      Caption         =   "&Mark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame fraCont1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   720
      TabIndex        =   14
      Top             =   720
      Width           =   5535
      Begin VB.Label lblFileTypes 
         Caption         =   "File Types:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   510
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblFind 
         Caption         =   "Fi&nd What:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblProgress 
         Caption         =   "Found: Nothing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label lblInFolder 
         Caption         =   "In Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblInFiles 
         Caption         =   "In Files/Types:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   510
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblReplace 
         Caption         =   "Replace &With:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblFoundNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbsFind 
      Height          =   2415
      Left            =   435
      TabIndex        =   0
      Top             =   120
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   4260
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabFixedWidth   =   1958
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "Find"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Replace"
            Key             =   "Replace"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find and Mark"
            Key             =   "Mark"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find Files"
            Key             =   "FindFiles"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Find dialog box for searching text.        ***
'*** Uses: public variables gFindCase (toggles  ***
'*** case sensitivity); gFindString (text to    ***
'*** find); gFindDirection (toggles search      ***
'*** direction); gFirstTime (toggles start from ***
'*** beginning of text)                         ***
'**************************************************

Option Explicit
Dim Position As Integer

Private Sub cboFind_Change()
    
    ' Set the public variable.
    gFirstTime = True
    ' If the cbobox is empty, disable the find button.
    If cboFind.text = "" Then
        cmdFind.Enabled = False
        cmdFindInFiles.Enabled = False
        cmdFindFiles.Enabled = False
    Else
        cmdFind.Enabled = True
        cmdFindInFiles.Enabled = True
        cmdFindFiles.Enabled = True
    End If

End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    
    Dim boxwhnd As Integer
    Dim r As Long
    Dim duplicate As Boolean
        
    duplicate = CheckDup(cboFind.text, cboFind)
    ' if enter is pressed
    If KeyAscii = 13 Then
        ' check to se if the text allready is in combo
        If duplicate = False Then
            cboFind.AddItem cboFind.text
            cboFind.SetFocus
            cmdFind_Click
            boxwhnd = GetFocus()
            r = SendMessage(boxwhnd, CB_SHOWDROPDOWN, 0, 0)
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub cboFind_LostFocus()
    
    'cboFind.AddItem cboFind.Text

End Sub

Private Sub cboFindInFiles_Change()
    
'    Dim boxwhnd As Integer
'    Dim r As Long
'
'    If KeyAscii = 13 Then
'
'        cboFindInFiles.AddItem cboFindInFiles.Text
'        cboFindInFiles.SetFocus
'        cmdFind_Click
'        boxwhnd = GetFocus()
'        r = SendMessage(boxwhnd, CB_SHOWDROPDOWN, 0, 0)
'        KeyAscii = 0
'    End If
'If CancelDirSearch Then Exit Sub
End Sub

Private Sub cboInFolder_KeyPress(KeyAscii As Integer)
    
    Dim boxwhnd As Integer
    Dim r As Long
    
    If KeyAscii = 13 Then
        
        cboInFolder.AddItem cboInFolder.text
        cboInFolder.SetFocus
        'cmdFind_Click
        boxwhnd = GetFocus()
        r = SendMessage(boxwhnd, CB_SHOWDROPDOWN, 0, 0)
        KeyAscii = 0
    End If

End Sub

Private Sub cboReplace_KeyPress(KeyAscii As Integer)
    
    Dim boxwhnd As Integer
    Dim r As Long
    Dim duplicate As Boolean
        
    duplicate = CheckDup(cboReplace.text, cboReplace)

    
    If KeyAscii = 13 Then
        If duplicate = False Then
            cboReplace.AddItem cboReplace.text
            cboReplace.SetFocus
            cmdReplace_Click
            boxwhnd = GetFocus()
            r = SendMessage(boxwhnd, CB_SHOWDROPDOWN, 0, 0)
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub cboReplace_LostFocus()
    
    'cboFind.AddItem cboFind.Text

End Sub

Private Sub chkCase_Click()
    ' Assign a value to the public variable.
    gFindCase = chkCase.Value
End Sub


Private Sub cmdBrowse_Click()
    Dim tmp As String
    tmp = BrowseForFolder(hwnd, "Select search folder...")
    cboInFolder.AddItem (tmp), 0
    cboInFolder.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    ' Save the values to the public variables.
    gFindString = cboFind.text
    gFindCase = chkCase.Value
    ' Unload the find dialog.
    Unload frmFind
    ' cancel any ongoing searches
    CancelDirSearch = True
End Sub

Private Sub cmdFind_Click()
    ' Assigns the text string to a public variable.
    gFindString = cboFind.text
    FindIt
End Sub


Private Sub cmdFindFiles_Click()

Dim fs1 As New FileSearch      ' Create a New Instans of Class
Dim iSize As Long              ' Filesize
Dim i As Long
Dim itmX As ListItem

    On Error GoTo searcherror:
    With fs1
        .Reset                                  ' Reset, only needed if class is global
        .ProgressLabel = lblProgress            ' if tou want to show progress info on Label1
        .StartSearchHere = cboInFolder.text     ' required, start here
        .Filters = cboFindInFiles.text          ' Filters ex. "*.bmp;*.exe;explorer.exe"
        .ListDirs = chkListDirs                 ' a numeric value, default fsAllFiles
        .SearchInSubDirs = chkSearchInSubDirs   ' = check1.value works fine !!!
        .Start                                  ' GO !!!!

        frmMDI.lvFindResults.ListItems.Clear    'Clear Out Old Items
        
        ' select container 3 on sstDriveFileList
        frmMDI.sstDriveFilelist.Tab = 3
        
       ' Loop through found items
        For i = 0 To .FileCount - 1
            .ListIndex = i
            ' Do What you want with search result !!!
                ' If .IsDirectory Then         ' if directory...
                ' X = .FileName                ' get filename
                 iSize = (.FileLenght / 1024) + 0.5   ' calculate filesize in Kb
                ' X = iSize & " Kb"
                ' X = .FileAttrib              ' get file attributes
                ' X = fs1.Location             ' get file locations , full path
                Set itmX = frmMDI.lvFindResults.ListItems.Add(, , .FileName)
'                Set clmX = ListView1.ColumnHeaders.Add(, , "Size", ListView1.Width / 3, lvwColumnRight)

                itmX.SubItems(1) = (iSize)
                itmX.Icon = 4          ' Set an icon from imlDriveFileList.
                itmX.SmallIcon = 4      ' Set an icon from ImageList2.
                itmX.SubItems(2) = (.Location)
        Next i
        
        lblProgress.Caption = "Found: " & .FileCount & " Files"

    End With

Exit Sub
searcherror:
    MsgBox Err.Description, vbExclamation, "Runtime Error: " & Err.Number
    fs1.Reset

End Sub

Private Sub cmdMark_Click()
    ' This routine searches the rtf for a string and marks all found
    ' words red.
    
    Dim old_pos As Integer
    Dim old_len As Integer
    old_pos = frmMDI.ActiveForm.Text1.SelStart
    old_len = frmMDI.ActiveForm.Text1.SelLength
    Dim x As Integer    ' counter variable
    ' Enabling the Unmark All button
    cmdUnMark.Enabled = True

    ' Enabling the Unmark All button
    cmdUnMark.Enabled = True

    ' This routine searches the rtf for a string and marks all found
    ' words red.

    'declare the temp variables
    Dim lWhere, lPos As Long
    Dim sTmp, sSearch As String

    'set lPos to 1 for mid()
    lPos = 1
    'this is the searched string
    sSearch = cboFind.text

    ' copy all text from main editor to temp editor. Works much faster
    frmMDI.ActiveForm.rtfTmp.text = frmMDI.ActiveForm.Text1.text

    'loop the whole text
    Do While lPos < Len(frmMDI.ActiveForm.rtfTmp.text)

        'get sub string from the text
        'this is because the InStr() returns the
        'position of first occurence of the string...
        sTmp = Mid(frmMDI.ActiveForm.rtfTmp.text, lPos, Len(frmMDI.ActiveForm.rtfTmp.text))

        'find the string in sub string
        lWhere = InStr(sTmp, sSearch)
        'accumulate the lPos to be relative to the actual text
        lPos = lPos + lWhere

        If lWhere Then   ' If found,
            x = x + 1
            frmMDI.ActiveForm.rtfTmp.SelStart = lPos - 2   ' set selection start and
            frmMDI.ActiveForm.rtfTmp.SelLength = Len(sSearch)   ' set selection length.   Else
            frmMDI.ActiveForm.rtfTmp.SelColor = RGB(255, 0, 0) 'change color to red
        Else
            lblFoundNumber.Caption = "Found: " & Str(x)
            Exit Do 'we are ready
        End If
    Loop

    'unselect the last found
    frmMDI.ActiveForm.rtfTmp.SelLength = 0
    
    ' now copy the text from rtfTmp to text1
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.text)
    frmMDI.ActiveForm.rtfTmp.SelStart = 0
    frmMDI.ActiveForm.rtfTmp.SelLength = Len(frmMDI.ActiveForm.rtfTmp.text)
    frmMDI.ActiveForm.Text1.SelRTF = frmMDI.ActiveForm.rtfTmp.SelRTF
    ' always clear the rtfTmp after use
    frmMDI.ActiveForm.rtfTmp.text = ""
    ' setting the cursor back to origianl possition
    frmMDI.ActiveForm.Text1.SelStart = old_pos
    frmMDI.ActiveForm.Text1.SelLength = old_len

End Sub

Private Sub cmdReplace_Click()

    Dim FindFlags As Integer

    frmMDI.ActiveForm.Text1.SelText = cboReplace.text
    FindFlags = chkCase.Value * 4 + chkWholeWord.Value * 2
    Position = frmMDI.ActiveForm.Text1.Find(cboFind.text, Position + 1, , FindFlags)
    If Position > 0 Then
        'frmMDI.ActiveForm.Text1.SetFocus
    Else
        MsgBox "String not found", vbOKOnly, "Search Help"
        'cmdReplace.Enabled = False
        'cmdReplaceAll.Enabled = False
    End If

End Sub

Private Sub cmdReplaceAll_Click()
    
    Dim FindFlags As Long

    FindFlags = chkCase.Value * 4 + chkWholeWord.Value * 2
    frmMDI.ActiveForm.Text1.SelText = cboReplace.text
    Position = frmMDI.ActiveForm.Text1.Find(cboFind.text, Position + 1, , FindFlags)
    While Position > 0
        frmMDI.ActiveForm.Text1.SelText = cboReplace.text
        Position = frmMDI.ActiveForm.Text1.Find(cboFind.text, Position + 1, , FindFlags)
    Wend
        'ReplaceButton.Enabled = False
        'ReplaceAllButton.Enabled = False
        MsgBox "Done replacing", vbOKOnly, "Search Help"

End Sub

Private Sub cmdUnMark_Click()
    
    ' This routine marks the whole text and colors it
    ' words red.
    
    Dim old_pos As Integer
    Dim old_len As Integer
    old_pos = frmMDI.ActiveForm.Text1.SelStart
    old_len = frmMDI.ActiveForm.Text1.SelLength

    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.text)
    ' Reading and setting collor settings for the editor
    frmMDI.ActiveForm.Text1.BackColor = GetSetting(App.Title, "Colors", "TextBack", &HFFFFFF)   'white
    frmMDI.ActiveForm.Text1.SelColor = GetSetting(App.Title, "Colors", "TextColor", &H0&)       'black
    'unselect the last found
    frmMDI.ActiveForm.rtfTmp.SelLength = 0
    ' setting the cursor back to origianl possition
    frmMDI.ActiveForm.Text1.SelStart = old_pos
    frmMDI.ActiveForm.Text1.SelLength = old_len

End Sub

Private Sub Form_Load()

    Call rotateText("E-Type Find", picBorder)

    ' click the first tab to hide the unwanted items
    tbsFind_Click
    ' Test to se if text is marked in editor
    If frmMDI.ActiveForm.Text1.SelText = "" Then    ' nothing
        ' Disable the find and findinfiles buttons - no text to search for yet.
        cmdFind.Enabled = False
        cmdFindInFiles.Enabled = False

        ' removing possible text from the searchtext textbox
        cboFind.text = ""
    Else
        ' putting the marked text from editor in to searchtext textbox
        cboFind.text = frmMDI.ActiveForm.Text1.SelText
    End If
    
    ' Read the public variable & set the option button.
    optDirection(gFindDirection).Value = 1
End Sub

Private Sub optDirection_Click(index As Integer)
    ' Assign a value to the public variable.
    gFindDirection = index
End Sub

Private Sub tbsFind_Click()
    
    ' all items
    'lblfind.Visible =
    'lblReplace.Visible =
    'cboFind.Visible =
    'cboReplace.Visible =
    'chkCase.Visible =
    'chkWholeWord.Visible =
    'fraDirections.Visible =
    'cmdFind.Visible =
    'cmdReplace.Visible =
    'cmdReplaceAll.Visible =
    'cmdMark.Visible =
    'cmdUnMark.Visible =
    'cmdBrowse
    'cboFindInFiles
    'cboInFolder
    'lblInFiles
    'lblInFolder
    'chkSubFolders
    'cmdFindInFiles
    'cmdFindFiles
    'chkListDirs
    'chkSearchInSubDirs
    'lblProgress
    'lblFileTypes
    'lblFoundNumber
    
    If tbsFind.SelectedItem.key = "Find" Then
        lblFind.Visible = True
        lblReplace.Visible = False
        cboFind.Visible = True
        cboReplace.Visible = False
        chkCase.Visible = True
        chkWholeWord.Visible = True
        fraDirections.Visible = True
        cmdFind.Visible = True
        cmdReplace.Visible = False
        cmdReplaceAll.Visible = False
        cmdMark.Visible = False
        cmdUnMark.Visible = False
        cmdBrowse.Visible = False
        cboFindInFiles.Visible = False
        cboInFolder.Visible = False
        lblInFiles.Visible = False
        lblInFolder.Visible = False
        chkSubFolders.Visible = False
        cmdFindInFiles.Visible = False
        cmdFindFiles.Visible = False
        chkListDirs.Visible = False
        chkSearchInSubDirs.Visible = False
        lblProgress.Visible = False
        lblFileTypes.Visible = False
        lblFoundNumber.Visible = False
           
    ElseIf tbsFind.SelectedItem.key = "Replace" Then
        lblFind.Visible = True
        lblReplace.Visible = True
        cboFind.Visible = True
        cboReplace.Visible = True
        chkCase.Visible = True
        chkWholeWord.Visible = True
        fraDirections.Visible = True
        cmdFind.Visible = True
        cmdReplace.Visible = True
        cmdReplaceAll.Visible = True
        cmdMark.Visible = False
        cmdUnMark.Visible = False
        cmdBrowse.Visible = False
        cboFindInFiles.Visible = False
        cboInFolder.Visible = False
        lblInFiles.Visible = False
        lblInFolder.Visible = False
        chkSubFolders.Visible = False
        cmdFindInFiles.Visible = False
        cmdFindFiles.Visible = False
        chkListDirs.Visible = False
        chkSearchInSubDirs.Visible = False
        lblProgress.Visible = False
        lblFileTypes.Visible = False
        lblFoundNumber.Visible = False
        
    ElseIf tbsFind.SelectedItem.key = "Mark" Then
        lblFind.Visible = True
        lblReplace.Visible = False
        cboFind.Visible = True
        cboReplace.Visible = False
        chkCase.Visible = True
        chkWholeWord.Visible = False
        fraDirections.Visible = False
        cmdFind.Visible = False
        cmdReplace.Visible = False
        cmdReplaceAll.Visible = False
        cmdMark.Visible = True
        cmdUnMark.Visible = True
        cmdBrowse.Visible = False
        cboFindInFiles.Visible = False
        cboInFolder.Visible = False
        lblInFiles.Visible = False
        lblInFolder.Visible = False
        chkSubFolders.Visible = False
        cmdFindInFiles.Visible = False
        cmdFindFiles.Visible = False
        chkListDirs.Visible = False
        chkSearchInSubDirs.Visible = False
        lblProgress.Visible = False
        lblFileTypes.Visible = False
        lblFoundNumber.Visible = True
    
    ElseIf tbsFind.SelectedItem.key = "FindFiles" Then
        lblFind.Visible = False
        lblReplace.Visible = False
        cboFind.Visible = False
        cboReplace.Visible = False
        chkCase.Visible = False
        chkWholeWord.Visible = False
        fraDirections.Visible = False
        cmdFind.Visible = False
        cmdReplace.Visible = False
        cmdReplaceAll.Visible = False
        cmdMark.Visible = False
        cmdUnMark.Visible = False
        cmdBrowse.Visible = True
        cboFindInFiles.Visible = True
        cboInFolder.Visible = True
        lblInFiles.Visible = False
        lblInFolder.Visible = True
        chkSubFolders.Visible = False
        cmdFindInFiles.Visible = False
        cmdFindFiles.Visible = True
        chkListDirs.Visible = True
        chkSearchInSubDirs.Visible = True
        lblProgress.Visible = True
        lblFileTypes.Visible = True
        lblFoundNumber.Visible = False
        
    ElseIf tbsFind.SelectedItem.key = "FindInFiles" Then
        lblFind.Visible = True
        lblReplace.Visible = False
        cboFind.Visible = True
        cboReplace.Visible = False
        chkCase.Visible = False
        chkWholeWord.Visible = False
        fraDirections.Visible = False
        cmdFind.Visible = True
        cmdReplace.Visible = False
        cmdReplaceAll.Visible = False
        cmdMark.Visible = False
        cmdUnMark.Visible = False
        cmdBrowse.Visible = True
        cboFindInFiles.Visible = True
        cboInFolder.Visible = True
        lblInFiles.Visible = True
        lblInFolder.Visible = True
        chkSubFolders.Visible = True
        cmdFindInFiles.Visible = True
        cmdFindFiles.Visible = False
        chkListDirs.Visible = False
        chkSearchInSubDirs.Visible = False
        lblProgress.Visible = False
        lblFileTypes.Visible = False
        lblFoundNumber.Visible = False

    End If

End Sub
 Function CheckDup(MyValue As Variant, MyCombo As ComboBox) As Boolean
    ' usage:
    'Dim duplicate As Boolean
    'duplicate = CheckDup(cboFind.Text, cboFind)
    
    Dim i ' Declare variable.

    For i = 0 To MyCombo.ListCount - 1
        If MyCombo.List(i) = MyValue Then
            CheckDup = True
            Exit Function
        End If
    Next i
    CheckDup = False
End Function

