VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNotePad 
   Caption         =   "Untitled:"
   ClientHeight    =   2010
   ClientLeft      =   5970
   ClientTop       =   2340
   ClientWidth     =   5775
   Icon            =   "notepad.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Tag             =   "1"
   Begin RichTextLib.RichTextBox Text1 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Tag             =   "docUnChanged"
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"notepad.frx":0E42
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
   Begin RichTextLib.RichTextBox rtfTmp 
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"notepad.frx":0F0B
   End
   Begin VB.PictureBox picLines 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   0
      ScaleHeight     =   134
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      Begin VB.Image imgMarkers 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "notepad.frx":0FDC
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.ListBox lstSortTempx 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "notepad.frx":1126
      Left            =   2760
      List            =   "notepad.frx":112D
      Sorted          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "C&lose All"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRevertToSaved 
         Caption         =   "Revert To Save&d"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertFile 
         Caption         =   "&Insert File"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Pr&int Preview"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page &Setup"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuESep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnuEditTime 
            Caption         =   "Time/&Date"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuLongDate 
            Caption         =   "Long Date"
         End
         Begin VB.Menu mnuShortDate 
            Caption         =   "Short Date"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuLongTime 
            Caption         =   "Long Time"
         End
         Begin VB.Menu mnuShortTime 
            Caption         =   "Short Time"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuESep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPathAndFilename 
            Caption         =   "&Path and Filename"
         End
         Begin VB.Menu mnuFilename 
            Caption         =   "&Filename"
         End
         Begin VB.Menu mnuInsSymbol 
            Caption         =   "&Symbol"
         End
      End
      Begin VB.Menu mnuWordWrap 
         Caption         =   "&Wordwrap"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchFindPrev 
         Caption         =   "Find &Previous"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSearchFindandMark 
         Caption         =   "Find and Mark"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchFindFiles 
         Caption         =   "Find File(s)"
      End
      Begin VB.Menu mnuCompareFiles 
         Caption         =   "Compare Files"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "&Statusbar"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuLineNumbers 
         Caption         =   "&Line Numbers"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuViewSidebar 
         Caption         =   "Side&bar"
         Checked         =   -1  'True
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuViewDocTabs 
         Caption         =   "&Document tabs"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "File &Info"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuClipBoard 
         Caption         =   "&ClipBoard"
      End
      Begin VB.Menu mnuMultiClip 
         Caption         =   "&MultiClip"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "For&mat"
      Begin VB.Menu mnuSortAsc 
         Caption         =   "Sort &Ascending"
      End
      Begin VB.Menu mnuSortDec 
         Caption         =   "Sort &Descending"
      End
      Begin VB.Menu mnuToUpperCase 
         Caption         =   "To &Upper Case"
      End
      Begin VB.Menu mnuToLowerCase 
         Caption         =   "To &Lower Case"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrypto 
         Caption         =   "&Crypto"
      End
      Begin VB.Menu mnuCustomizeToolbar 
         Caption         =   "Customize &Toolbar"
      End
      Begin VB.Menu mnuOptionsDialog 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuDuplicate 
         Caption         =   "&Duplicate Window"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade Windows"
      End
      Begin VB.Menu mnuWindowTileHor 
         Caption         =   "&Tile Horizontal"
      End
      Begin VB.Menu mnuTileWindowsVert 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNextWindow 
         Caption         =   "&Next Window"
      End
      Begin VB.Menu mnuPrevWindow 
         Caption         =   "&Previous Window"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize All Windows"
      End
      Begin VB.Menu mnuMaximizeWindows 
         Caption         =   "M&aximize All Windows"
      End
      Begin VB.Menu mnuRestoreWindows 
         Caption         =   "&Restore All Windows"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTipOfTheDay 
         Caption         =   "Tip Of The &Day"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Child form for the MDI Notepad sample application  ***
'**********************************************************
Option Explicit

Public fLineHeight As Single        ' holds the character height in picLines
Public fCharWidth As Single         ' holds the character width in picLines
Private WordWrapState As Boolean    ' Wordwrap on off
'*
Private trapUndo As Boolean         ' flag to indicate whether actions should be trapped
Private UndoStack As New Collection ' collection of undo elements
Private RedoStack As New Collection ' collection of redo elements
'*

Private Const WM_USER = &H400
Private Const WM_PASTE = &H302
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300

'Private Enum eTextMode
'    TM_PLAINTEXT = 1
'    TM_RICHTEXT = 2                ' /* default behavior */
'    TM_SINGLELEVELUNDO = 4
'    TM_MULTILEVELUNDO = 8          ' /* default behavior */
'    TM_SINGLECODEPAGE = 16
'    TM_MULTICODEPAGE = 32          ' /* default behavior */
'End Enum

Private Const EM_LINEINDEX = &HBB&
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_LINEFROMCHAR = &HC9&
Private Const EM_CANPASTE = (WM_USER + 50)
Private Const EM_HIDESELECTION = (WM_USER + 63)
Private Const EM_REQUESTRESIZE = (WM_USER + 65)
Private Const EM_SETUNDOLIMIT = (WM_USER + 82)
Private Const EM_REDO = (WM_USER + 84)
Private Const EM_CANREDO = (WM_USER + 85)
Private Const EM_GETUNDONAME = (WM_USER + 86)
Private Const EM_GETREDONAME = (WM_USER + 87)
Private Const EM_STOPGROUPTYPING = (WM_USER + 88)
Private Const EM_SETTEXTMODE = (WM_USER + 89)
Private Const EM_GETTEXTMODE = (WM_USER + 90)
Private Const EM_AUTOURLDETECT = (WM_USER + 91)

Private Sub Form_Activate()
    
    Call SetFonts
        
    ' PicLines
    picLines.Visible = GetSetting(App.Title, "Apperance", "ShowLineNumbers", 0)
    picLines.BackColor = GetSetting(App.Title, "colors", "LineNumBack", &H8000000A)
    picLines.ForeColor = GetSetting(App.Title, "colors", "LineNumText", &H80000010)
    picLines.FONTSIZE = Text1.font.Size 'Settings.FontSize
    picLines.FontName = Text1.font.Name  'Settings.Font
    ' draw linenumbers
    If picLines.Visible = True Then
        DrawLines
    End If
    
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If
    
    ' If statusbar is visible, update the linenumber and cursorpos in statusbar
    If frmMDI.SBarMain.Visible = True Then
        frmMDI.DocumentSelChange Text1
    End If

    ' updating the view statusbar,view toolbar, view Sidebar and view linenumbers menu
    mnuStatusBar.Checked = frmMDI.SBarMain.Visible
    mnuToolbar.Checked = frmMDI.tbToolBar.Visible
    mnuLineNumbers.Checked = picLines.Visible
    mnuViewSidebar.Checked = frmMDI.picSideBar.Visible
    mnuViewDocTabs.Checked = frmMDI.tbaDocuments.Visible
    
    ' updating the recentfiles menu
    Call GetRecentFiles
    
    ' set fucus to the tab coresponding to the active document
    frmMDI.tabDocuments.Tabs("key" & Me.Tag).Caption = StripPath(Me.Caption)
    frmMDI.tabDocuments.Tabs("key" & Me.Tag).Selected = True
    frmMDI.tabDocuments.Tabs("key" & Me.Tag).ToolTipText = Me.Caption

    ' setting the properties of rtfTmp the same as text1
    rtfTmp.font = Text1.font
    rtfTmp.font.Bold = Text1.font.Bold
    rtfTmp.font.Size = Text1.font.Size
    rtfTmp.font.Italic = Text1.font.Italic
    rtfTmp.font.Weight = Text1.font.Weight
    
    ' add the new window to the toolbar windowlist button
    frmMDI.ChildStatusChanged

    ' Enabling Revert To Saved menu if document first 9 letters in
    ' filename is something other than "untitled:"
    If Not Left((Me.Caption), 9) = "Untitled:" Then
        If Me.Text1.Tag = "docChanged" Then
            mnuRevertToSaved.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    '*
    Dim textChangedStatus As String
    textChangedStatus = Me.Text1.Tag
    trapUndo = True         'Enable Undo Trapping
'    Call Text1_Change       'Initialize First Undo
'    Call Text1_SelChange    'Initialize Menus
    ' After this the promgram will think the text has changed
    ' so we must set the changed tag back to what it was
    Me.Text1.Tag = textChangedStatus
    '*
    
    ' add doctab for the new document
    frmMDI.tabDocuments.Tabs.Add fIndex, ("key" & Format(fIndex)), "me.tag: " & Me.Tag    'Me.Caption
    
    ' setting wordwrap from registry
    Me.mnuWordWrap.Checked = GetSetting(App.Title, "Doc Defaults", "LineWrap", 0)
    WordWrapState = mnuWordWrap.Checked
    If WordWrapState = False Then
        Text1.RightMargin = 200000
    End If
    Call fixMenues
    ' Disabling Revert To Saved menu.
    ' The file has not changed yet"
    mnuRevertToSaved.Enabled = False

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    Dim strFilename As String
    Dim intResponse As Integer

    ' Check to see if the text has been changed.
    If Me.Text1.Tag = "docChanged" Then
        strFilename = Me.Caption
        strMsg = "The text in [" & strFilename & "] has changed."
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "Do you want to save the changes?"
        intResponse = MsgBox(strMsg, 51, frmMDI.Caption)
        Select Case intResponse
            Case 6      ' User chose Yes.
                If Left(Me.Caption, 8) = "Untitled" Then
                    ' The file hasn't been saved yet.
                    strFilename = "untitled.txt"
                    ' Get the strFilename, and then call the save procedure, GetstrFilename.
                    strFilename = GetFileName(strFilename)
                Else
                    ' The form's Caption contains the name of the open file.
                    strFilename = Me.Caption
                End If
                ' Call the save procedure. If strFilename = Empty, then
                ' the user chose Cancel in the Save As dialog box; otherwise,
                ' save the file.
                If strFilename <> "" Then
                    SaveFileAs strFilename
                End If
            Case 7      ' User chose No. Unload the file.
                Cancel = False
            Case 2      ' User chose Cancel. Cancel the unload.
                Cancel = True
        End Select
    End If
    ' remove the active doctab
    frmMDI.tabDocuments.Tabs.Remove ("key" & Format(Me.Tag))  'frmMDI.tabDocuments.SelectedItem.index

End Sub

Private Sub Form_Resize()
    ' check now if piclines is supposed to be visible or not
    picLines.Visible = GetSetting(App.Title, "Apperance", "ShowLineNumbers", 0)

    ' Resizing the RTF control with the form
    If Me.Width > 105 And Me.Height > 755 Then
        ' Expand text box to fill the current child form's internal area.
        Text1.Height = ScaleHeight
    
        If picLines.Visible = True Then
            Text1.Left = picLines.Width + 2
        Else
            Text1.Left = 2
        End If
        If picLines.Visible = True Then
            Text1.Width = ScaleWidth - picLines.Width - 2
        Else
            Text1.Width = ScaleWidth - 2
        End If
        ' draw linenumbers
        If picLines.Visible = True Then
            DrawLines
        End If
    Else
        Text1.Width = 100
        Text1.Height = 100
    End If
    ' setting the margins according to the wordwrap state of the form
    If WordWrapState = True Then
        Text1.RightMargin = Text1.Width
    ElseIf WordWrapState = False Then
        Text1.RightMargin = 200000
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' Show the current form instance as deleted
    FState(Me.Tag).Deleted = True

    ' Hide the toolbar edit buttons if no notepad windows exist.
    If Not AnyPadsLeft() Then
'        frmMDI.imgCutButton.Visible = False
'        frmMDI.imgCopyButton.Visible = False
'        frmMDI.imgPasteButton.Visible = False
        ' Toggle the public tool state variable
'        gToolsHidden = True
        ' Call the recent file list procedure
        GetRecentFiles
    End If
'    ' add/remove the new window to the toolbar windowlist button
'    Call RebuildWinList
    frmMDI.ChildStatusChanged
'    ' remove the active doctab
'    frmMDI.tabDocuments.Tabs.Remove frmMDI.tabDocuments.SelectedItem.index

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModeless, frmMDI ' Me makes the window minimized with the main form
End Sub

Private Sub mnuClipBoard_Click()
    frmClipViewer.Show vbModeless, frmMDI ' Me makes the window minimized with the main form

End Sub

Private Sub mnuCloseAll_Click()
    
'    Dim x
'    For x = 1 To UBound(Document)     ' number of open files
'        Unload Document(x)  ' unload all open files
'    Next

    ' next is from Visual Basic 101 tech tips 10th. edition page 20
    Dim vForm As Variant
    For Each vForm In Forms
        If Not TypeOf vForm Is MDIForm Then
            If vForm.MDIChild Then
                Unload vForm
            End If
        End If
    Next    ' vform

End Sub

Private Sub mnuCompareFiles_Click()
    frmCompare.Show vbModeless, frmMDI
End Sub

Private Sub mnuCrypto_Click()
    On Error Resume Next
    frmCrypto.Show vbModeless, frmMDI ' Me makes the window minimized with the main form
End Sub

Private Sub mnuCustomizeToolbar_Click()
    ' load the toolbar cutomitazion wizard
    frmMDI.tbToolBar.Customize
End Sub

Private Sub mnuDuplicate_Click()
    Call frmMDI.dupeWindow
End Sub

Private Sub mnuEditCopy_Click()
    ' Call the copy procedure
    EditCopy
'    EditCopyProc
End Sub

Private Sub mnuEditCut_Click()
    ' Call the cut procedure
    EditCut
'    EditCutProc
End Sub

Private Sub mnuEditDelete_Click()
    ' If the mouse pointer is not at the end of the notepad...
    If Screen.ActiveControl.SelStart <> Len(Screen.ActiveControl.text) Then
        ' If nothing is selected, extend the selection by one.
        If Screen.ActiveControl.SelLength = 0 Then
            Screen.ActiveControl.SelLength = 1
            ' If the mouse pointer is on a blank line, extend the selection by two.
            If Asc(Screen.ActiveControl.SelText) = 13 Then
                Screen.ActiveControl.SelLength = 2
            End If
        End If
        ' Delete the selected text.
        Screen.ActiveControl.SelText = ""
    End If
End Sub

Private Sub mnuEditPaste_Click()
    ' Call the paste procedure.
    EditPaste
'    EditPasteProc
End Sub

Private Sub mnuEditSelectAll_Click()
    ' Use SelStart & SelLength to select the text.
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.text)
End Sub

Private Sub mnuEditTime_Click()
    ' Insert the current time and date.
    Text1.SelText = Now
End Sub

Private Sub mnuFileClose_Click()
    ' Unload this form.
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    ' Unloading the MDI form invokes the QueryUnload event
    ' for each child form, and then the MDI form.
    ' Setting the Cancel argument to True in any of the
    ' QueryUnload events cancels the unload.
    Unload frmMDI
End Sub

Private Sub mnuFileInfo_Click()
    
    frmFileInfo.Show vbModeless, frmMDI ' Me makes the window minimized with the main form

End Sub

Private Sub mnuFilename_Click()
    Text1.SelText = StripPath(Me.Caption)
End Sub

Private Sub mnuFileNew_Click()
    ' Call the new form procedure
    FileNew
End Sub


Private Sub mnuFileOpen_Click()
    ' Call the file open procedure.
    FileOpenProc
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If
End Sub

Private Sub mnuFileSave_Click()
    Call SaveFile
End Sub

Private Sub mnuFileSaveAll_Click()
'*************************************************
' Purpose:  Saves all RTF controls on the forms.
'*************************************************

'    Dim intI As Integer 'counter
'    Dim frm As Form 'current form variable
'    For intI = 0 To Forms.Count - 1 'count forms
'        Set frm = Forms(intI) 'set it into variable
'        If frm.MDIChild = True Then 'if child
'            Call SaveFile
''            Menu_FileSave Forms(intI) 'save it
'        End If
'    Next intI
    
    Dim vForm As Variant
    For Each vForm In Forms
        If Not TypeOf vForm Is MDIForm Then
            If vForm.MDIChild Then
                Call SaveFile
            End If
        End If
    Next    ' vform


End Sub

Private Sub mnuFileSaveAs_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Assign the form caption to the variable.
    strDefaultName = Me.Caption
    If Left(Me.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, strSaveFileName.
        
        strSaveFileName = GetFileName("Untitled.txt")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    Else
        ' The form's Caption contains the name of the open file.
        
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    End If

End Sub

Private Sub mnuFont_Click()
    On Error GoTo ErrorHandler
    
    frmMDI.CMDialog1.Flags = cdlCFScreenFonts
    frmMDI.CMDialog1.ShowFont
    
    'changing the selected fontattributes on the active form
    With Text1.font
        .Name = frmMDI.CMDialog1.FontName
        .Bold = frmMDI.CMDialog1.FontBold
        .Italic = frmMDI.CMDialog1.FontItalic
        .Size = frmMDI.CMDialog1.FONTSIZE
    End With
    'changing the selected fontattributes on the active forms linenumbers
    With frmMDI.ActiveForm.picLines.font
        .Name = frmMDI.CMDialog1.FontName
        .Bold = frmMDI.CMDialog1.FontBold
        .Italic = frmMDI.CMDialog1.FontItalic
        .Size = frmMDI.CMDialog1.FONTSIZE
    End With
    If picLines.Visible = True Then
        ' repainting the linenumbers
        DrawLines
    End If
    ' Saving the new font settings to registry.
    ' The fontdialog defalts to no fontname selected, therefore dont change
    ' fontname if user selects cansel.
    If frmMDI.CMDialog1.FontName = "" Then
        ' nothing
    Else
        SaveSetting App.Title, "Doc Defaults", "FontName", frmMDI.CMDialog1.FontName
    End If
    SaveSetting App.Title, "Doc Defaults", "FontSize", frmMDI.CMDialog1.FONTSIZE
    If frmMDI.CMDialog1.FontBold = True Then
        SaveSetting App.Title, "Doc Defaults", "FontBold", 1
    Else
        SaveSetting App.Title, "Doc Defaults", "FontBold", 0
    End If
    If frmMDI.CMDialog1.FontItalic = True Then
        SaveSetting App.Title, "Doc Defaults", "FontItalic", 1
    Else
        SaveSetting App.Title, "Doc Defaults", "FontItalic", 0
    End If

    Exit Sub

ErrorHandler:
    Exit Sub

End Sub



Private Sub mnuInsertFile_Click()
    InsertFile
End Sub

Private Sub mnuInsSymbol_Click()
    frmSymbols.Show vbModeless, frmMDI ' Me makes the window minimized with the main form
End Sub

Private Sub mnuLineNumbers_Click()

    mnuLineNumbers.Checked = Not mnuLineNumbers.Checked
    picLines.Visible = mnuLineNumbers.Checked
    ' save all settings to registry
    SaveSetting App.Title, "Apperance", "ShowLineNumbers", mnuLineNumbers.Checked

    ' resize the form to update the text1 size
    Form_Resize


End Sub

Private Sub mnuLongDate_Click()
    Text1.SelText = Format(Date$, "Long Date")
End Sub

Private Sub mnuLongTime_Click()
    Text1.SelText = Format(Time$, "Long Time")
End Sub

Private Sub mnuMaximizeWindows_Click()
    Dim x
    
    For x = 1 To fIndex
        Document(x).WindowState = 2
    Next
End Sub

Private Sub mnuMinimize_Click()
    Dim x
    
    For x = 1 To fIndex
        Document(x).WindowState = 1
    Next
End Sub

Private Sub mnuMultiClip_Click()
    frmMultiClip.Show vbModeless, frmMDI ' Me makes the window minimized with the main form

End Sub

Private Sub mnuNextWindow_Click()
    Call frmMDI.SelNextWin
End Sub

Private Sub mnuOptionsDialog_Click()
    frmOptions.Show vbModeless, frmMDI ' Me makes the window minimized with the main form
End Sub

Private Sub mnuPageSetup_Click()
'Purpose: To display frmPageSetup to setup the Print header and footer

    Load frmPageSetup
    frmPageSetup.txtHeader.text = sPrintHeader 'set defaults
    frmPageSetup.txtFooter.text = sPrintFooter
    frmPageSetup.Show vbModal

End Sub

Private Sub mnuPathAndFilename_Click()
    Text1.SelText = Me.Caption
End Sub

Private Sub mnuPrevWindow_Click()
    Call frmMDI.SelPrevWin
End Sub

Private Sub mnuPrint_Click()
    ' Check to se if a printer is installed
    Dim pTest
    
    pTest = Printer.papersize
    
    If pTest = 0 Then   ' no printer installed
        GoTo errHandler
    End If

    ' call printText
    Call printText
    
    Exit Sub

errHandler:
    MsgBox "No printer found! Printing not possible.", vbCritical
    Exit Sub

End Sub

Private Sub mnuPrintPreview_Click()
    
    ' Check to se if a printer is installed
    Dim pTest
    
    pTest = Printer.papersize
    
    If pTest = 0 Then   ' no printer installed
        GoTo errHandler
    End If

    'setup header and footer
    sPrintText = Text1.text
    sHeader = SetPrintLine(sPrintHeader)
    sFooter = SetPrintLine(sPrintFooter)
    sPrintText = sHeader & vbCrLf & vbCrLf & sPrintText & vbCrLf & vbCrLf & sFooter
    Me.rtfTmp.text = sPrintText

    frmDocPreview.Show vbModal
    
    Exit Sub

errHandler:
    MsgBox "No printer found! Preview not possible.", vbCritical
    Exit Sub

End Sub

Private Sub mnuRecentFile_Click(index As Integer)
        
        Dim strOpenFileName As String
        
        strOpenFileName = (mnuRecentFile(index).Caption)
        
        ' check if the file is already open
        Dim a As Integer
        Dim ArrayCount As Integer
        
        ArrayCount = UBound(Document)
        
        ' Cycle through the document array. If one of the
        ' documents filenames matches the file droped, inform
        ' the user that the file already is open.
        For a = 1 To ArrayCount
            If Document(a).Caption = strOpenFileName Then
                ' now check if the already open file has been edited
                If Document(a).Text1.Tag = "docChanged" Then
                    Dim strMsg As String
                    Dim strFilename As String
                    Dim intResponse As Integer
    
                    strFilename = Document(a).Caption
                    strMsg = "The text in [" & strFilename & "] has changed."
                    strMsg = strMsg & vbCrLf
                    strMsg = strMsg & "Reload file and loose changes?"
                    intResponse = MsgBox(strMsg, vbYesNo + vbExclamation + vbDefaultButton2, "File alredy open!")
                    Select Case intResponse
                    Case 6      ' User chose Yes.
                        ' first close the already open document
                        ' Unload this form.
                        Document(a).Text1.Tag = "docUnChanged"
                        Unload Document(a)
                        ' Call the file open procedure, passing a
                        ' reference to the selected file name
                        OpenFile strOpenFileName
                        ' Update the list of recently opened files in the File menu control array.
                        UpdateFileMenu strOpenFileName
                    Case 7      ' User chose No.
                        ' nothing
                        frmMDI.ActiveForm.Text1.SetFocus
                    End Select
                Else
                    MsgBox "Document already open and in the same state.", vbInformation
                    frmMDI.ActiveForm.Text1.SetFocus
                End If
                ' the document is already open an in the same state.
                ' No need to open it
                Exit Sub
            End If
        Next
        ' The file is not already open.
        ' Call the file open procedure, passing a
        ' reference to the selected file name
        OpenFile (strOpenFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strOpenFileName)
    
        ' Update the list of recently opened files in the File menu control array.
        GetRecentFiles
    
    ' Enabling Revert To Saved menu if document first 9 letters in
    ' filename is something other than "untitled:"
    If Not Left((Me.Caption), 9) = "Untitled:" Then
        If Not Me.Text1.Tag = "docChanged" Then
            mnuRevertToSaved.Enabled = False
        End If
    End If

End Sub

Private Sub mnuRedo_Click()
    Call Redo
'    EditRedo
End Sub

Private Sub mnuReplace_Click()
    SearchReplace
'    ' If there is text in the textbox, assign it to
'    ' the textbox on the Find form, otherwise assign
'    ' the last findtext value.
'    If Me.Text1.SelText <> "" Then
'        frmFind.cboFind.Text = Me.Text1.SelText
'    Else
'        frmFind.cboFind.Text = gFindString
'    End If
'    ' Set the public variable to start at the beginning.
'    gFirstTime = True
'    ' Set the case checkbox to match the public variable
'    If (gFindCase) Then
'        frmFind.chkCase = 1
'    End If
'    ' Display the Find form.
'    frmFind.Show 1
'    ' click the Replace tab
'    frmFind.tbsFind.Tabs(2).Selected = True

End Sub

Private Sub mnuRestoreWindows_Click()
    Dim x
    
    For x = 1 To fIndex
        Document(x).WindowState = 0
    Next
End Sub

Private Sub mnuRevertToSaved_Click()
    Dim strMsg As String
    Dim strFilename As String
    Dim intResponse As Integer

    strFilename = Me.Caption
    strMsg = "The text in [" & strFilename & "] has changed."
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "Do you want to reload file and loose changes?"
    intResponse = MsgBox(strMsg, 52, frmMDI.Caption)
    Select Case intResponse
        Case 6      ' User chose Yes.
            Text1.LoadFile (Text1.FileName)
            mnuRevertToSaved.Enabled = False
        Case 7      ' User chose No. Do not revert.
            ' nothing
    End Select
     Me.Text1.SetFocus
End Sub

Private Sub mnuSearchFind_Click()
    Call SearchFind

End Sub

Private Sub mnuSearchFindandMark_Click()
    Call SearchFindAndMark
End Sub

Private Sub mnuSearchFindFiles_Click()
    Call SearchFindFiles
End Sub

Private Sub mnuSearchFindNext_Click()
    ' Assign a value to the public variable.
    gFindDirection = 1
    Call SearchFindNext
End Sub

Private Sub mnuSearchFindPrev_Click()

    ' Assign a value to the public variable.
    gFindDirection = 0
    Call SearchFindPrev

End Sub

Private Sub mnuShortDate_Click()
    Text1.SelText = Format(Date$, "Short Date")
End Sub

Private Sub mnuShortTime_Click()
    Text1.SelText = Format(Time$, "Short Time")
End Sub

Private Sub mnuSortAsc_Click()

    Dim lines() As String
    Dim new_line As String
    Dim num_lines As Integer
    Dim fnum As Integer
    Dim i As Integer

    MousePointer = vbHourglass
    DoEvents
    
    ' saving active text to tempfile in app path
    Call Text1.SaveFile(GetTmpPath & "temp.txt", 1)
    
    ' Read the lines from the tempfile.
    fnum = FreeFile
    Open GetTmpPath & "temp.txt" For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, new_line
        new_line = Trim$(new_line)

        num_lines = num_lines + 1
        ReDim Preserve lines(1 To num_lines)
        lines(num_lines) = new_line
    Loop
    Close fnum

    SelectionSort lines, 1, num_lines

    ' Save the results to the output file.
    Open GetTmpPath & "temp.txt" For Output As fnum
    ' Save in ascending order.
    For i = 1 To num_lines
        Print #fnum, lines(i)
    Next i
    
    Close fnum
    Call Text1.LoadFile(GetTmpPath & "temp.txt")
    Kill GetTmpPath & "temp.txt"

    MousePointer = vbDefault
'    MsgBox "Sorted " & Format$(num_lines) & " lines"

End Sub

Private Sub mnuSortDec_Click()
    Dim lines() As String
    Dim new_line As String
    Dim num_lines As Integer
    Dim fnum As Integer
    Dim i As Integer

    MousePointer = vbHourglass
    DoEvents
    
    ' saving active text to tempfile in app path
    Call Text1.SaveFile(GetTmpPath & "temp.txt", 1)
    
    ' Read the lines from the tempfile.
    fnum = FreeFile
    Open GetTmpPath & "temp.txt" For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, new_line
        new_line = Trim$(new_line)

        num_lines = num_lines + 1
        ReDim Preserve lines(1 To num_lines)
        lines(num_lines) = new_line
    Loop
    Close fnum

    SelectionSort lines, 1, num_lines

    ' Save the results to the output file.
    Open GetTmpPath & "temp.txt" For Output As fnum
    ' Save in descending order.
    For i = num_lines To 1 Step -1
        Print #fnum, lines(i)
    Next i
    Close fnum
    Call Text1.LoadFile(GetTmpPath & "temp.txt")
    Kill GetTmpPath & "temp.txt"

    MousePointer = vbDefault
'    MsgBox "Sorted " & Format$(num_lines) & " lines"

End Sub

Private Sub mnuStatusBar_Click()
    mnuStatusBar.Checked = Not mnuStatusBar.Checked
    frmMDI.SBarMain.Visible = mnuStatusBar.Checked
    ' resize the controls on frmMDI.sidebar
    frmMDI.resizeME
End Sub

Private Sub mnuTileWindowsVert_Click()
    ' Tile the child forms.
    frmMDI.Arrange vbTileVertical
End Sub

Private Sub mnuTipOfTheDay_Click()
    ' sett the Show tip at startup setting in registry
    SaveSetting App.Title, "StartUp", "ShowTip", 1
    Load frmTip
    frmTip.Show vbModal ' wait here
End Sub

Private Sub mnuToLowerCase_Click()
    Clipboard.SetText Text1.SelText
    Text1.SelText = LCase(Clipboard.GetText)
End Sub

Private Sub mnuToolbar_Click()
    mnuToolbar.Checked = Not mnuToolbar.Checked
    frmMDI.tbToolBar.Visible = mnuToolbar.Checked
    ' make shure the toolbar is placed on top of doctabs
    frmMDI.tbToolBar.Top = 0
    ' resize the controls on frmMDI.sidebar
    frmMDI.resizeME
End Sub

Private Sub mnuToUpperCase_Click()
    Clipboard.SetText Text1.SelText
    Text1.SelText = UCase(Clipboard.GetText)
End Sub

Private Sub mnuUndo_Click()
    Call Undo
'    EditUndo
End Sub

Private Sub mnuViewDocTabs_Click()

    mnuViewDocTabs.Checked = Not mnuViewDocTabs.Checked
    frmMDI.tbaDocuments.Visible = mnuViewDocTabs.Checked
    ' make shure the toolbar is placed on top of doctabs
    frmMDI.tbToolBar.Top = 0
    ' resize the controls on frmMDI.sidebar
    frmMDI.resizeME

End Sub

Private Sub mnuViewSidebar_Click()
    mnuViewSidebar.Checked = Not mnuViewSidebar.Checked
    frmMDI.picSideBar.Visible = mnuViewSidebar.Checked
    Call frmMDI.StartSideBar
    
    ' cascade all documents for good looks
    frmMDI.Arrange vbCascade

End Sub

Private Sub mnuWindowArrange_Click()
    ' Arrange the icons for any minimzied child forms.
    frmMDI.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    ' Cascade the child forms.
    frmMDI.Arrange vbCascade
End Sub


Private Sub mnuWindowTileHor_Click()
    ' Tile the child forms.
    frmMDI.Arrange vbTileHorizontal
End Sub

Private Sub mnuWordWrap_Click()

    If WordWrapState = True Then
        WordWrapState = False
        mnuWordWrap.Checked = WordWrapState
        Text1.RightMargin = 200000
    ElseIf WordWrapState = False Then
        WordWrapState = True
        mnuWordWrap.Checked = WordWrapState
        Text1.RightMargin = Text1.Width
    End If

End Sub

Private Sub picLines_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static MarksCount As Integer
    If Button = vbLeftButton Then
        Load imgMarkers(MarksCount + 1)
        imgMarkers(MarksCount + 1).Left = picLines.Width - 16
        imgMarkers(MarksCount + 1).Top = y
        imgMarkers(MarksCount + 1).Visible = True
        MarksCount = MarksCount + 1
    End If

End Sub

Private Sub Text1_Change()
    ' Set the public variable to show that text has changed.
    Text1.Tag = "docChanged"
    
    If picLines.Visible = True Then
        ' repainting the linenumbers
        DrawLines
    End If
    
    '*
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim c%, l&

    'remove all redo items because of the change
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = Text1.SelStart
    newElement.TextLen = Len(Text1.text)
    newElement.text = Text1.text

    'add it to the undo stack
    UndoStack.Add Item:=newElement
    'enable controls accordingly
    EnableControls
    
    ' Enabling Revert To Saved menu if document first 9 letters in
    ' filename is something other than "untitled:"
    If Not Left((Me.Caption), 9) = "Untitled:" Then
        mnuRevertToSaved.Enabled = True
    End If

End Sub

Private Sub Text1_GotFocus()
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Here the we captures the keys pressed and cancels them
    ' before new actions is performed.
    
    If (Shift = vbCtrlMask) Then
        Select Case KeyCode
            Case vbKeyZ
                KeyCode = 0
                mnuUndo_Click
            Case vbKeyY
                KeyCode = 0
                mnuRedo_Click
            Case vbKeyV
                KeyCode = 0
                mnuEditPaste_Click
        End Select
    End If

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show popupmenu
    If Button = vbRightButton Then 'do the popup menu
        PopupMenu mnuEdit
    End If

End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picLines.Visible = True Then
        ' repainting the linenumbers
        DrawLines
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'File or directory?
'    On Error Resume Next
    If (GetAttr(Data.Files(Data)) And vbDirectory) = vbDirectory Then
        Me.OLEDropMode = 0
    Else
        ' Open the selected file.
        Text1.LoadFile Data.Files(Data)
        If Err Then
            MsgBox "Can't open file: " + Data, vbExclamation
            Exit Sub
        End If
        ' Change the mouse pointer to an hourglass.
        Screen.MousePointer = 11
        
        ' Change the form's caption and display the new text.
        Me.Caption = UCase(Data)
        Me.Text1.text = StrConv(InputB(LOF(1), 1), vbUnicode)
        Me.Text1.Tag = "docChanged"
        Close #1
        ' Reset the mouse pointer.
        Screen.MousePointer = 0
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu Command$
        ' If statusbar is visible, update infomation on form
        If frmMDI.SBarMain.Visible = True Then
            GetFileStats    ' filedata, date,size ect.
        End If
    End If
    
End Sub

Private Sub Text1_SelChange()
    ' If statusbar is visible, update the linenumber and cursorpos in statusbar
    If frmMDI.SBarMain.Visible = True Then
        frmMDI.DocumentSelChange Text1
    End If
    If picLines.Visible = True Then
        ' repainting the linenumbers
        DrawLines
    End If
'    ' undo redo stuff
'    UpdateToolbar
'*
    Dim ln&
    If Not trapUndo Then Exit Sub
    ln& = Text1.SelLength
    mnuEditCut.Enabled = ln&    'disabled if length of selected text is 0
    mnuEditCopy.Enabled = ln&   'disabled if length of selected text is 0
    mnuEditPaste.Enabled = Len(Clipboard.GetText(1)) 'disabled if length of clipboard text is 0
    mnuEditDelete.Enabled = ln&  'disabled if length of selected text is 0
    mnuEditSelectAll.Enabled = CBool(Len(Text1.text)) 'disabled if length of textbox's text is 0
    With frmMDI.tbToolBar
        .Buttons("Cut").Enabled = mnuEditCut.Enabled
        .Buttons("Copy").Enabled = mnuEditCopy.Enabled
        .Buttons("Paste").Enabled = (SendMessage(Text1.hwnd, EM_CANPASTE, 0, 0) = 1)
        .Buttons("Undo").Enabled = (SendMessage(Text1.hwnd, EM_CANUNDO, 0, 0) = 1)
        .Buttons("Redo").Enabled = (SendMessage(Text1.hwnd, EM_CANREDO, 0, 0) = 1)
    End With
    
    With Me
        .mnuEditCut.Enabled = frmMDI.tbToolBar.Buttons("Cut").Enabled
        .mnuEditCopy.Enabled = frmMDI.tbToolBar.Buttons("Copy").Enabled
        .mnuEditPaste.Enabled = frmMDI.tbToolBar.Buttons("Paste").Enabled
        .mnuUndo.Enabled = frmMDI.tbToolBar.Buttons("Undo").Enabled
        .mnuRedo.Enabled = frmMDI.tbToolBar.Buttons("Redo").Enabled
    End With
    
End Sub

Public Sub DrawLines()
    Dim lngFirstLine As Long
    Dim fTop As Single
    Dim lngLinesOnScreen As Long
    Dim lngFirstLineIndex As Long
    Dim lngLine As Long
    Dim yPixels As Long
    Dim xPixels As Long
    Dim i As Long
    Dim GetFirstLineVisible
    
    ' getting the height of a character from picLines
    fLineHeight = Me.picLines.TextHeight("TESTTEXT")
    fCharWidth = Me.picLines.TextWidth("0")
    
    With picLines
        If .Visible = False Then Exit Sub
        GetFirstLineVisible = SendMessage(Text1.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)

        lngFirstLine = GetFirstLineVisible
        lngFirstLineIndex = SendMessage(Text1.hwnd, EM_LINEINDEX, lngFirstLine, 0&)
        lngLinesOnScreen = (Text1.Height) / fLineHeight
        .Cls
        fTop = 0
        For i = 0 To lngLinesOnScreen + 1
            .CurrentY = fTop
            .CurrentX = 0
            fTop = fTop + fLineHeight
            picLines.Print i + lngFirstLine + 1 ' adding one extra for good looks
        Next
        ' scaling the with of piclines according to the actual with of the numbers displayed
        ' first test the with in pixels the last printed number is.
        ' then add the number of pixels one char with is.
        ' then add 3 more pixels for good looks
        picLines.Width = Me.picLines.TextWidth(lngLinesOnScreen + lngFirstLine) + fCharWidth + 3
        ' move the text1 to the edge of pcLines and scale it to right witdh
        Text1.Left = 0 + picLines.Width
        Text1.Width = ScaleWidth - picLines.Width

    End With
End Sub


Public Sub ColorLine()
'
'      Text1.SelStart = Where - 1   ' set selection start and
'      Text1.SelLength = Len(Search)   ' set selection length.
   
End Sub

Public Sub SearchFind()
    ' If there is text in the textbox, assign it to
    ' the textbox on the Find form, otherwise assign
    ' the last findtext value.
    If Me.Text1.SelText <> "" Then
        frmFind.cboFind.text = Me.Text1.SelText
    Else
        frmFind.cboFind.text = gFindString
    End If
    ' Set the public variable to start at the beginning.
    gFirstTime = True
    ' Set the case checkbox to match the public variable
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Display the Find form.
    frmFind.Show vbModeless, frmMDI  ' Me makes the window minimized with the main form
    frmFind.cboFind.SetFocus
End Sub
Public Sub SearchFindAndMark()
    ' If there is text in the textbox, assign it to
    ' the textbox on the Find form, otherwise assign
    ' the last findtext value.
    If Me.Text1.SelText <> "" Then
        frmFind.cboFind.text = Me.Text1.SelText
    Else
        frmFind.cboFind.text = gFindString
    End If
    ' Set the public variable to start at the beginning.
    gFirstTime = True
    ' Set the case checkbox to match the public variable
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Display the Find form.
    frmFind.Show vbModeless, frmMDI  ' Me makes the window minimized with the main form
    ' click the Replace tab
    frmFind.tbsFind.Tabs(3).Selected = True
    frmFind.cboFind.SetFocus

End Sub
Public Sub SearchFindFiles()
    ' Display the Find form.
    frmFind.Show vbModeless, frmMDI  ' Me makes the window minimized with the main form
    ' click the Replace tab
    frmFind.tbsFind.Tabs(4).Selected = True
    frmFind.cboFindInFiles.SetFocus

End Sub

Public Sub SearchFindNext()

    ' If the public variable isn't empty, call the
    ' find procedure, otherwise call the find menu
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If

End Sub

Public Sub SearchFindPrev()
    ' If the public variable isn't empty, call the
    ' find procedure, otherwise call the find menu
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If


End Sub

Public Sub SearchReplace()

    ' If there is text in the textbox, assign it to
    ' the textbox on the Find form, otherwise assign
    ' the last findtext value.
    If Me.Text1.SelText <> "" Then
        frmFind.cboFind.text = Me.Text1.SelText
    Else
        frmFind.cboFind.text = gFindString
    End If
    ' Set the public variable to start at the beginning.
    gFirstTime = True
    ' Set the case checkbox to match the public variable
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Display the Find form.
    frmFind.Show vbModeless, frmMDI  ' Me makes the window minimized with the main form
    frmFind.cboFind.SetFocus
    ' click the Replace tab
    frmFind.tbsFind.Tabs(2).Selected = True
 
End Sub

Public Sub SaveFile()
    Dim strFilename As String

    If Left(Me.Caption, 9) = "Untitled:" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = Me.Caption
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
    End If
    
    ' disabling the reverttosaved menu
    mnuRevertToSaved.Enabled = False

End Sub


Public Sub printText()
    
    On Error GoTo ErrorHandler
    
    'setup header and footer
    sPrintText = Text1.text
    sHeader = SetPrintLine(sPrintHeader)
    sFooter = SetPrintLine(sPrintFooter)
    sPrintText = sHeader & vbCrLf & vbCrLf & sPrintText & vbCrLf & vbCrLf & sFooter
    Me.rtfTmp.text = sPrintText
    
    ' This is where the printing is called
    On Error GoTo ErrorHandler
    frmMDI.CMDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
        
    If Text1.SelLength = 0 Then
        frmMDI.CMDialog1.Flags = frmMDI.CMDialog1.Flags + cdlPDAllPages
    Else
        frmMDI.CMDialog1.Flags = frmMDI.CMDialog1.Flags + cdlPDSelection
    End If
    frmMDI.CMDialog1.ShowPrinter
    ' Printing with margin at all four sides.
    ' To use the PrintRTF function we must send it margins in TWIPS. Since the
    ' pagesetup form uses millimeters we must convert them to twips first.
    ' There is aproximatly 57 TWIPS in 1 millimeter.
    PrintRTF Text1, (gLeftMargin * 57), (gTopMargin * 57), (gRightMargin * 57), (gBottomMargin * 57)

    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub UpdateToolbar()
    
    With frmMDI.tbToolBar
        .Buttons("Cut").Enabled = (Text1.SelLength <> 0)
        .Buttons("Copy").Enabled = (Text1.SelLength <> 0)
        .Buttons("Paste").Enabled = (SendMessage(Text1.hwnd, EM_CANPASTE, 0, 0) = 1)
        .Buttons("Undo").Enabled = (SendMessage(Text1.hwnd, EM_CANUNDO, 0, 0) = 1)
        .Buttons("Redo").Enabled = (SendMessage(Text1.hwnd, EM_CANREDO, 0, 0) = 1)
    End With
    
    With Me
        .mnuEditCut.Enabled = frmMDI.tbToolBar.Buttons("Cut").Enabled
        .mnuEditCopy.Enabled = frmMDI.tbToolBar.Buttons("Copy").Enabled
        .mnuEditPaste.Enabled = frmMDI.tbToolBar.Buttons("Paste").Enabled
        .mnuUndo.Enabled = frmMDI.tbToolBar.Buttons("Undo").Enabled
        .mnuRedo.Enabled = frmMDI.tbToolBar.Buttons("Redo").Enabled
    End With
    
'    With m_ab
'
'        .Tools("miECut").Enabled = (RTF.SelLength <> 0)
'        .Tools("miECopy").Enabled = (RTF.SelLength <> 0)
'        .Tools("miEPaste").Enabled = (SendMessage(RTF.hwnd, EM_CANPASTE, 0, 0) = 1)
'        .Tools("miEUndo").Enabled = (SendMessage(RTF.hwnd, EM_CANUNDO, 0, 0) = 1)
'        .Tools("miERedo").Enabled = (SendMessage(RTF.hwnd, EM_CANREDO, 0, 0) = 1)
'
'        .Refresh
'    End With
    
'    tmr.Enabled = False
End Sub

Public Sub EditCopy()
    SendMessage Text1.hwnd, WM_COPY, 0, 0
    ' Copy the selected text to the list on frmMultiClip
    frmMultiClip.lstClip.AddItem (Clipboard.GetText)
    Call frmMultiClip.KillDupes

End Sub

Public Sub EditPaste()
    SendMessage Text1.hwnd, WM_PASTE, 0, 0
End Sub

Public Sub EditCut()
    SendMessage Text1.hwnd, WM_CUT, 0, 0
'    text1.SetFocus
End Sub

Public Sub EditUndo()
Dim hr As Long
    hr = SendMessage(Text1.hwnd, EM_GETUNDONAME, 0&, 0&)
    ' Debug.Print hr, Choose(hr + 1, "Unknown", "Typing", "Delete", "Drag Drop", "Cut", "Paste")
    SendMessage Text1.hwnd, EM_UNDO, 0, 0
End Sub

Public Sub EditRedo()
    SendMessage Text1.hwnd, EM_REDO, 0, 0
End Sub

Public Sub Undo()
'*
Dim chg$, x&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object

    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
'            cmdDummy.SetFocus   'change focus of form
            x& = SendMessage(Text1.hwnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            Text1.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            Text1.SelLength = objElement.TextLen - objElement2.TextLen
            Text1.SelText = ""
            x& = SendMessage(Text1.hwnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.text, objElement2.text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.text) - Len(objElement2.text)))
            Text1.SelStart = objElement2.SelStart
            Text1.SelLength = 0
            Text1.SelText = chg$
            Text1.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                Text1.SelLength = Len(chg$)
            Else
                Text1.SelStart = Text1.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    EnableControls
    trapUndo = True
    Text1.SetFocus
End Sub

Public Sub Redo()
'*
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(Text1.text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            Text1.SelStart = objElement.SelStart
            Text1.SelLength = Len(Text1.text) - objElement.TextLen
            Text1.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(Text1.text, objElement.text, objElement.SelStart + 1)
            Text1.SelStart = objElement.SelStart - Len(chg$)
            Text1.SelLength = 0
            Text1.SelText = chg$
            Text1.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                Text1.SelLength = Len(chg$)
            Else
                Text1.SelStart = Text1.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    EnableControls
    trapUndo = True
    Text1.SetFocus
End Sub

Private Sub EnableControls()
'*
    mnuUndo.Enabled = UndoStack.Count > 1
    mnuRedo.Enabled = RedoStack.Count > 0

    Text1_SelChange
    
    With frmMDI.tbToolBar
        .Buttons("Cut").Enabled = (Text1.SelLength <> 0)
        .Buttons("Copy").Enabled = (Text1.SelLength <> 0)
        .Buttons("Paste").Enabled = (SendMessage(Text1.hwnd, EM_CANPASTE, 0, 0) = 1)
        .Buttons("Undo").Enabled = mnuUndo.Enabled
        .Buttons("Redo").Enabled = mnuRedo.Enabled
    End With

End Sub

Private Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
'*
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub fixMenues()
    ' This sub sets some menue items to captions the
    ' Menu Editor doesn't allow or sets key shortcuts
    ' without setting them in the editor.
    
    mnuEditPaste.Caption = "&Paste" & vbTab & "Ctrl+V"
    mnuUndo.Caption = "&Undo" & vbTab & "Ctrl+Z"
    mnuRedo.Caption = "&Redo" & vbTab & "Ctrl+Y"
    mnuNextWindow.Caption = mnuNextWindow.Caption & vbTab & "Ctrl+Tab"
    mnuPrevWindow.Caption = mnuPrevWindow.Caption & vbTab & "Ctrl+Shift+Tab"
    
End Sub

Public Sub SetFonts()
    ' Everytime a form is activated the colors will be sett from the setting in registry.
    ' Therefore the promgram will think the text has changed so we must test to se if
    ' it allready have. If it has not we must sett it to unchanged after the colors
    ' has been set.
    Dim textChangedStatus As String
    textChangedStatus = Me.Text1.Tag
    ' Reading and setting collor settings for the editor
    Text1.BackColor = GetSetting(App.Title, "Colors", "TextBack", &HFFFFFF)   'white
    Text1.SelColor = GetSetting(App.Title, "Colors", "TextColor", &H0&)       'black
    
    'changing the selected fontattributes on the active form
    With Text1.font
        .Name = GetSetting(App.Title, "Doc Defaults", "FontName", "Arial")
        .Bold = GetSetting(App.Title, "Doc Defaults", "FontBold", 0)
        .Italic = GetSetting(App.Title, "Doc Defaults", "FontItalic", 0)
        .Size = GetSetting(App.Title, "Doc Defaults", "FontSize", "8")
    End With
    
    'changing the selected fontattributes on the active forms linenumbers
    With frmMDI.ActiveForm.picLines.font
        .Name = Text1.font
        .Bold = Text1.font.Bold
        .Italic = Text1.font.Italic
        .Size = Text1.font.Size
    End With
    
    ' set the form backcolor the same as text1 backcolor
    Me.BackColor = Text1.BackColor
    ' After the colorchange the promgram will think the text has changed
    ' so we must set the changed tag back to what it was
    Me.Text1.Tag = textChangedStatus

End Sub
