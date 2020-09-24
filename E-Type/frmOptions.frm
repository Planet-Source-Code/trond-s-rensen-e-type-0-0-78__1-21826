VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4905
   ClientLeft      =   3450
   ClientTop       =   2790
   ClientWidth     =   6465
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Tag             =   "Options"
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
      Height          =   4905
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   330
      TabIndex        =   51
      Top             =   0
      Width           =   330
   End
   Begin MSComDlg.CommonDialog cdOptions 
      Left            =   480
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4050
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -19635
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   7
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -19635
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   6
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -19635
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   4
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin TabDlg.SSTab sstTabs 
      Height          =   4095
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Apperance"
      TabPicture(0)   =   "frmOptions.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkLineNr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDocumentTabs"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkStatusBar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkToolBar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkSideBar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Doc Defaults"
      TabPicture(1)   =   "frmOptions.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFontSize"
      Tab(1).Control(1)=   "txtFontName"
      Tab(1).Control(2)=   "chkFontItalic"
      Tab(1).Control(3)=   "chkFontBold"
      Tab(1).Control(4)=   "picFontNamrContainer"
      Tab(1).Control(5)=   "txtTabSize"
      Tab(1).Control(6)=   "hssTabSize"
      Tab(1).Control(7)=   "cmdBrowseFont"
      Tab(1).Control(8)=   "chkLineWrap"
      Tab(1).Control(9)=   "lblTabSize"
      Tab(1).Control(10)=   "lblDefaultFont"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Colors"
      TabPicture(2)   =   "frmOptions.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picBackColor"
      Tab(2).Control(1)=   "picText"
      Tab(2).Control(2)=   "picLineNumBack"
      Tab(2).Control(3)=   "picLineNumText"
      Tab(2).Control(4)=   "picActiveLineText"
      Tab(2).Control(5)=   "picActiveLineBack"
      Tab(2).Control(6)=   "Picture1"
      Tab(2).Control(7)=   "cmdUseDefault"
      Tab(2).Control(8)=   "chkEnableLineColoring"
      Tab(2).Control(9)=   "Label6"
      Tab(2).Control(10)=   "Label14"
      Tab(2).Control(11)=   "Label12"
      Tab(2).Control(12)=   "Label11"
      Tab(2).Control(13)=   "Label10"
      Tab(2).Control(14)=   "Label8"
      Tab(2).Control(15)=   "Label7"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Accosiations"
      TabPicture(3)   =   "frmOptions.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblList"
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(2)=   "lstFileTypes"
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtFontSize 
         Height          =   285
         Left            =   -70440
         TabIndex        =   50
         Text            =   "FontSize"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtFontName 
         Height          =   285
         Left            =   -70440
         TabIndex        =   49
         Text            =   "FontName"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkFontItalic 
         Caption         =   "FontItalic"
         Height          =   375
         Left            =   -70440
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkFontBold 
         Caption         =   "FontBold"
         Height          =   255
         Left            =   -70440
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picFontNamrContainer 
         ForeColor       =   &H00FFFFFF&
         Height          =   290
         Left            =   -74760
         ScaleHeight     =   225
         ScaleWidth      =   3915
         TabIndex        =   45
         Top             =   1320
         Width           =   3975
         Begin VB.Label lblFontAttributes 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.TextBox txtTabSize 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73920
         TabIndex        =   44
         Text            =   "4"
         Top             =   1800
         Width           =   375
      End
      Begin VB.HScrollBar hssTabSize 
         Enabled         =   0   'False
         Height          =   280
         Left            =   -73560
         TabIndex        =   41
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseFont 
         Caption         =   "..."
         Height          =   300
         Left            =   -70680
         TabIndex        =   40
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkLineWrap 
         Caption         =   "Word Wrap"
         Height          =   195
         Left            =   -74760
         TabIndex        =   39
         Top             =   600
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkSideBar 
         Caption         =   "Show Sidebar"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox chkToolBar 
         Caption         =   "Show Toolbar"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkStatusBar 
         Caption         =   "Show Statusbar"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkDocumentTabs 
         Caption         =   "Show Document Tabs"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox chkLineNr 
         Caption         =   "Show Line Numbers"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.PictureBox picBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.PictureBox picLineNumBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox picLineNumText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   21
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picActiveLineText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74745
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox picActiveLineBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74745
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   2535
         Left            =   -72330
         Picture         =   "frmOptions.frx":0EB2
         ScaleHeight     =   2475
         ScaleWidth      =   3045
         TabIndex        =   13
         Top             =   720
         Width           =   3105
         Begin VB.PictureBox picPWBackColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   810
            ScaleHeight     =   2055
            ScaleWidth      =   2190
            TabIndex        =   16
            Top             =   345
            Width           =   2190
            Begin VB.Label lblPWText 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Some text"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   705
            End
            Begin VB.Label lblPWHLLine 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF0000&
               Caption         =   "mode con codepage select=850"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   0
               TabIndex        =   17
               Top             =   195
               Width           =   2295
            End
         End
         Begin VB.PictureBox picPWLinenum 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   75
            ScaleHeight     =   2055
            ScaleWidth      =   735
            TabIndex        =   14
            Top             =   345
            Width           =   735
            Begin VB.Label lblPWLineNum 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "123456789"
               ForeColor       =   &H80000008&
               Height          =   2010
               Left            =   60
               TabIndex        =   15
               Top             =   15
               Width           =   420
            End
         End
      End
      Begin VB.CommandButton cmdUseDefault 
         Caption         =   "Use Defaults"
         Height          =   375
         Left            =   -70305
         TabIndex        =   12
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox chkEnableLineColoring 
         Caption         =   "Enable active line coloring"
         Height          =   255
         Left            =   -74745
         TabIndex        =   11
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ListBox lstFileTypes 
         Height          =   1185
         ItemData        =   "frmOptions.frx":19968
         Left            =   -74760
         List            =   "frmOptions.frx":19978
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Assosiate"
         Height          =   375
         Left            =   -72120
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblTabSize 
         Caption         =   "Tab Size:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   1815
         Width           =   735
      End
      Begin VB.Label lblDefaultFont 
         Caption         =   "Default Document Font:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "(Slows down performance on large files)"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Background"
         Height          =   255
         Left            =   -74145
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "line number text"
         Height          =   255
         Left            =   -74145
         TabIndex        =   36
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "line number background"
         Height          =   255
         Left            =   -74145
         TabIndex        =   35
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Preview"
         Height          =   255
         Left            =   -72345
         TabIndex        =   34
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   "Text"
         Height          =   255
         Left            =   -74145
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Active line Background"
         Height          =   255
         Left            =   -74145
         TabIndex        =   32
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Active line text"
         Height          =   255
         Left            =   -74145
         TabIndex        =   31
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblList 
         Caption         =   "The following filetypes can be assosiated with E-Type"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim DefShowLineNumbers As Boolean
'Dim DefShowCursorPossition As Boolean
'Dim DefShowFileData As Boolean
'Dim DefShowFileSize As Boolean
'Dim DefShowIconMenues As Boolean
'Dim DefTrayIcon As Boolean
'Dim DefMaximizeNewDoc As Boolean
'Dim DefNumberOfRecenFiles As Integer
'Dim DefClearUndoBufferWhenSaving As Boolean
'Dim DefAutoSave As Boolean
'Dim DefAutoSaveMinuttes As Integer
'Dim DefBackUpType As String
'Dim DefBackUpOnSave As Boolean
'Dim DefBackUpExtension As String
'Dim DefBackUpFolder As String
'Dim DefOnStartup As String
'Dim DefMainWindow As String
'Dim DefStartIn As String
'Dim DefDragAndDropEditing As Boolean
Dim DefFontName As String
Dim DefFontSize As String
Dim DefFontBold As String
Dim DefFontItalic As Boolean
'Dim DefLineWrap As Boolean
'Dim DefTabSize As Integer
'Dim DefTextColor As String
'Dim DefTextBack As String
'Dim DefLineNumBack As String
'Dim DefLineNumText As String
'Dim DefActiveLineBack As String
'Dim DefActiveLineText As String
'Dim DefActiveLineColoring As Boolean


Private Sub chkEnableLineColoring_Click()
    lblPWHLLine.Visible = chkEnableLineColoring.Value

End Sub


Private Sub cmdBrowseFont_Click()
    ' Set Cancel to True.
    cdOptions.CancelError = True
    On Error GoTo errHandler
    ' Set the Flags property.
    cdOptions.Flags = cdlCFScreenFonts 'cdlCFBoth Or cdlCFEffects
    ' Display the Font dialog box.
    cdOptions.ShowFont
    
    ' the fontdialog defalts to no fontname selected, therefore dont change
    ' fontname if user selects cansel.
    If cdOptions.FontName = "" Then
        ' nothing
    Else
        txtFontName.text = cdOptions.FontName
    End If
    txtFontSize.text = cdOptions.FONTSIZE
    If cdOptions.FontBold = True Then
        chkFontBold.Value = 1
    Else
        chkFontBold.Value = 0
    End If
    If cdOptions.FontItalic = True Then
        chkFontItalic.Value = 1
    Else
        chkFontItalic.Value = 0
    End If

    ' changing the text in fonttype to display the selected font
    lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points "
    If chkFontBold.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Bold"
    End If
    If chkFontItalic.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Italic"
    End If
    If chkFontBold.Value = 1 And chkFontItalic.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Bold, Italic"
    End If

errHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
        
    ' save all settings to registry
    ' Apperance
    SaveSetting App.Title, "Apperance", "ShowLineNumbers", chkLineNr.Value
    SaveSetting App.Title, "Apperance", "ShowToolBar", chkToolBar.Value
    SaveSetting App.Title, "Apperance", "ShowStatusBar", chkStatusBar.Value
    SaveSetting App.Title, "Apperance", "ShowDocumentTabs", chkDocumentTabs.Value
    SaveSetting App.Title, "Apperance", "ShowSideBar", chkSideBar.Value
    ' Doc defaults
    SaveSetting App.Title, "Doc Defaults", "LineWrap", chkLineWrap.Value
    SaveSetting App.Title, "Doc Defaults", "FontName", txtFontName.text
    SaveSetting App.Title, "Doc Defaults", "FontSize", txtFontSize.text
    SaveSetting App.Title, "Doc Defaults", "FontBold", chkFontBold.Value
    SaveSetting App.Title, "Doc Defaults", "FontItalic", chkFontItalic.Value
    SaveSetting App.Title, "Doc Defaults", "TabSize", hssTabSize.Value
    ' Colors
    SaveSetting App.Title, "Colors", "TextColor", picText.BackColor
    SaveSetting App.Title, "Colors", "TextBack", picBackColor.BackColor
    SaveSetting App.Title, "Colors", "LineNumBack", picLineNumBack.BackColor
    SaveSetting App.Title, "Colors", "LineNumText", picLineNumText.BackColor
    SaveSetting App.Title, "Colors", "ActiveLineBack", picActiveLineBack.BackColor
    SaveSetting App.Title, "Colors", "ActiveLineText", picActiveLineText.BackColor
    SaveSetting App.Title, "Colors", "ActiveLineColoring", chkEnableLineColoring.Value
    ' Associations
    SaveSetting App.Title, "Associations", "AssociateTXT", lstFileTypes.Selected(0)
    SaveSetting App.Title, "Associations", "AssociateINI", lstFileTypes.Selected(1)
    SaveSetting App.Title, "Associations", "AssociateINF", lstFileTypes.Selected(2)
    SaveSetting App.Title, "Associations", "AssociateBAT", lstFileTypes.Selected(3)
    
    ' Changing fontattributtes on all open forms.
    ' Since the rtf will belive the text has changed just by canging collor
    ' we must check if the text has changed on every form we apply collors
    ' to before it canges, then change the collors and then sett the
    ' text changed status back to what it was
    ' Then we must cycle through the document array and and do the same thing
    ' on all open windows.
    
    Dim x                    ' document array
    Dim cPosNow              ' the curent cursor possition
    Dim textChangedStatus    ' holds the current fstatus (has text changed)
        On Error Resume Next
        For x = 1 To fIndex  ' number of open documents
            With Document(x)
                textChangedStatus = Document(x).Text1.Tag
                ' Setting all settings to editor
                .Text1.BackColor = picBackColor.BackColor
                ' Use SelStart & SelLength to select the text, then cange the color and
                ' then set the cursor back to its original possition
                cPosNow = .Text1.SelStart
                .Text1.SelStart = 0
                .Text1.SelLength = Len(frmMDI.ActiveForm.Text1.text)
                .Text1.SelColor = picText.BackColor
                ' changing the selected fontattributes on the active forms linenumbers
                .Text1.font.Name = GetSetting(App.Title, "Doc Defaults", "FontName", "Arial")
                .Text1.font.Bold = GetSetting(App.Title, "Doc Defaults", "FontBold", 0)
                .Text1.font.Italic = GetSetting(App.Title, "Doc Defaults", "FontItalic", 0)
                .Text1.font.Size = GetSetting(App.Title, "Doc Defaults", "FontSize", "8")

                .Text1.SelLength = 0
                .Text1.SelStart = cPosNow

                Document(x).Text1.Tag = textChangedStatus
                ' updating the colorchanges to picLines
                .picLines.BackColor = picLineNumBack.BackColor
                .picLines.ForeColor = picLineNumText.BackColor
                If Document(x).picLines.Visible = True Then
                    ' repainting the linenumbers
                    Document(x).DrawLines
                End If
            End With
        Next

    
    ' unloading the form
    Unload Me
End Sub

Private Sub cmdUseDefault_Click()
    picText.BackColor = &H0&                    'black
    picBackColor.BackColor = &HFFFFFF           'white
    picLineNumBack.BackColor = &H8000000A       'menubar
    picLineNumText.BackColor = &H80000010       'button shadow
    picActiveLineBack.BackColor = &H8000000D    'highlight
    picActiveLineText.BackColor = &H8000000E    'highlight text
    ' setting the colors of the preview items the same color as the colorbuttons
    lblPWLineNum.ForeColor = picLineNumText.BackColor
    picPWLinenum.BackColor = picLineNumBack.BackColor
    picPWBackColor.BackColor = picBackColor.BackColor
    lblPWText.ForeColor = picText.BackColor
    lblPWHLLine.BackColor = picActiveLineBack.BackColor
    lblPWHLLine.ForeColor = picActiveLineText.BackColor
    lblPWHLLine.Visible = chkEnableLineColoring.Value

End Sub

Private Sub Command1_Click()
    Dim ProgPath As String
    Dim IconPath As String
    Dim IconPath2 As String
    Dim SystemDIR As String
    
    If Right$(App.Path, 1) <> "\" Then
         ProgPath = App.Path & "\" & "E-Type.exe %1"
    Else
         ProgPath = App.Path & "E-Type.exe %1"
    End If
    
    IconPath = SystemDIR & "\" & "shell32.dll,-152"
    
    If Right$(App.Path, 1) <> "\" Then
        IconPath2 = App.Path & "\" & "E-Type.ico"
    Else
        IconPath2 = App.Path & "E-Type.ico"
    End If

          
          ' lstFileTypes.Selected ()
     If lstFileTypes.Selected(0) = False Then
          SaveString HKEY_CLASSES_ROOT, "txtfile\DefaultIcon", "", SystemDIR & "\shell32.dll,-152"
          SaveString HKEY_CLASSES_ROOT, "txtfile\shell\open\command", "", "NOTEPAD.EXE %1"
          SaveString HKEY_CLASSES_ROOT, "txtfile\shell\print\command", "", "NOTEPAD.EXE /p %1"
          
          SaveSetting App.Title, "Associations", "AssociateTXT", lstFileTypes.Selected(0)
     Else
          SaveString HKEY_CLASSES_ROOT, "txtfile\DefaultIcon", "", IconPath2
          SaveString HKEY_CLASSES_ROOT, "txtfile\shell\open\command", "", ProgPath
          
          SaveSetting App.Title, "Associations", "AssociateTXT", lstFileTypes.Selected(0)
     End If
     
     If lstFileTypes.Selected(1) = False Then
          SaveString HKEY_CLASSES_ROOT, "inifile\shell\open\command", "", "NOTEPAD.EXE %1"
          
          SaveSetting App.Title, "Associations", "AssociateINI", lstFileTypes.Selected(1)
     Else
          SaveString HKEY_CLASSES_ROOT, "inifile\shell\open\command", "", ProgPath
          
          SaveSetting App.Title, "Associations", "AssociateINI", lstFileTypes.Selected(1)
     End If

     If lstFileTypes.Selected(2) = False Then
          SaveString HKEY_CLASSES_ROOT, "inffile\shell\open\command", "", "NOTEPAD.EXE %1"
          
          SaveSetting App.Title, "Associations", "AssociateINF", lstFileTypes.Selected(2)
     Else
          SaveString HKEY_CLASSES_ROOT, "inffile\shell\open\command", "", ProgPath
          
          SaveSetting App.Title, "Associations", "AssociateINF", lstFileTypes.Selected(2)
     End If
     
     If lstFileTypes.Selected(3) = False Then
          SaveString HKEY_CLASSES_ROOT, "batfile\shell\edit\command", "", "NOTEPAD.EXE %1"
          
          SaveSetting App.Title, "Associations", "AssociateBAT", lstFileTypes.Selected(3)
     Else
          SaveString HKEY_CLASSES_ROOT, "batfile\shell\edit\command", "", ProgPath
          
          SaveSetting App.Title, "Associations", "AssociateBAT", lstFileTypes.Selected(3)
     End If
        

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Call rotateText("E-Type Options", picBorder)
    
    ' get all settings from registry
    ' Apperance

    chkLineNr.Value = GetSetting(App.Title, "Apperance", "ShowLineNumbers", 0)
    chkToolBar.Value = GetSetting(App.Title, "Apperance", "ShowToolBar", 0)
    chkStatusBar.Value = GetSetting(App.Title, "Apperance", "ShowStatusBar", 0)
    chkDocumentTabs.Value = GetSetting(App.Title, "Apperance", "ShowDocumentTabs", 0)
    chkSideBar.Value = GetSetting(App.Title, "Apperance", "ShowSideBar", 0)
    ' Doc defaults
    chkLineWrap.Value = GetSetting(App.Title, "Doc Defaults", "LineWrap", 0)
    txtFontName.text = GetSetting(App.Title, "Doc Defaults", "FontName", "Arial")
    txtFontSize.text = GetSetting(App.Title, "Doc Defaults", "FontSize", "8")
    chkFontBold.Value = GetSetting(App.Title, "Doc Defaults", "FontBold", 0)
    chkFontItalic.Value = GetSetting(App.Title, "Doc Defaults", "FontItalic", 0)
    hssTabSize.Value = GetSetting(App.Title, "Doc Defaults", "TabSize", "4")
    ' Colors
    picText.BackColor = GetSetting(App.Title, "Colors", "TextColor", &H0&)           'black
    picBackColor.BackColor = GetSetting(App.Title, "Colors", "TextBack", &HFFFFFF)    'white
    picLineNumBack.BackColor = GetSetting(App.Title, "Colors", "LineNumBack", &H80000004)   'menubar
    picLineNumText.BackColor = GetSetting(App.Title, "Colors", "LineNumText", &H80000010)   'button shadow
    picActiveLineBack.BackColor = GetSetting(App.Title, "Colors", "ActiveLineBack", &H8000000D)  'highlight
    picActiveLineText.BackColor = GetSetting(App.Title, "Colors", "ActiveLineText", &H8000000E)  'highlight text
    chkEnableLineColoring.Value = GetSetting(App.Title, "Colors", "ActiveLineColoring", 1)   ' active line coloring
    ' Associations
    lstFileTypes.Selected(0) = GetSetting(App.Title, "Associations", "AssociateTXT", "0")
    lstFileTypes.Selected(1) = GetSetting(App.Title, "Associations", "AssociateINI", "0")
    lstFileTypes.Selected(2) = GetSetting(App.Title, "Associations", "AssociateINF", "0")
    lstFileTypes.Selected(3) = GetSetting(App.Title, "Associations", "AssociateBAT", "0")

        
    ' changing the text in fonttype to display the selected font
    lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points "
    If chkFontBold.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Bold"
    End If
    If chkFontItalic.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Italic"
    End If
    If chkFontBold.Value = 1 And chkFontItalic.Value = 1 Then
        lblFontAttributes.Caption = txtFontName.text & ", " & txtFontSize.text & " Points " & ", Bold, Italic"
    End If


    ' populating the preview items with some text
    lblPWLineNum.Caption = "1" & Chr(13) & "2" & Chr(13) & "3" & Chr(13) & "4" & Chr(13) & "5" & Chr(13) & "6" & Chr(13) & "7" & Chr(13) & "8" & Chr(13) & "9" & Chr(13) & "10"
    lblPWText.Caption = "mode con codepage prepare=((850)" & Chr(13) & "mode con codepage select=850" & Chr(3) & "keyb no,,C:\WINDOWS\COMMAND\keyboard.sys" & Chr(13) & "c:\maestro.com" & Chr(13) & "doskey"
    ' setting the colors of the preview items the same color as the colorbuttons
    lblPWLineNum.ForeColor = picLineNumText.BackColor
    picPWLinenum.BackColor = picLineNumBack.BackColor
    picPWBackColor.BackColor = picBackColor.BackColor
    lblPWText.ForeColor = picText.BackColor
    lblPWHLLine.BackColor = picActiveLineBack.BackColor
    lblPWHLLine.ForeColor = picActiveLineText.BackColor
    
    lblPWHLLine.Visible = chkEnableLineColoring.Value
        
End Sub


Private Sub picActiveLineBack_Click()
    cdOptions.ShowColor
    picActiveLineBack.BackColor = cdOptions.Color
    lblPWHLLine.BackColor = cdOptions.Color
End Sub

Private Sub picActiveLineText_Click()
    cdOptions.ShowColor
    picActiveLineText.BackColor = cdOptions.Color
    lblPWHLLine.ForeColor = cdOptions.Color
End Sub

Private Sub picBackColor_Click()
    cdOptions.ShowColor
    picBackColor.BackColor = cdOptions.Color
    picPWBackColor.BackColor = cdOptions.Color
End Sub

Private Sub picLineNumBack_Click()
    cdOptions.ShowColor
    picLineNumBack.BackColor = cdOptions.Color
    picPWLinenum.BackColor = cdOptions.Color
End Sub

Private Sub picLineNumText_Click()
    cdOptions.ShowColor
    picLineNumText.BackColor = cdOptions.Color
    lblPWLineNum.ForeColor = cdOptions.Color
End Sub

Private Sub picText_Click()
    cdOptions.ShowColor
    picText.BackColor = cdOptions.Color
    lblPWText.ForeColor = cdOptions.Color
End Sub


