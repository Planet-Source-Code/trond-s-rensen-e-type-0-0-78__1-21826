VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert symbol from "
   ClientHeight    =   5340
   ClientLeft      =   3675
   ClientTop       =   2610
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7545
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
      Height          =   5340
      Left            =   0
      ScaleHeight     =   5340
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   0
      Width           =   330
   End
   Begin VB.TextBox txtInsert 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   160
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1092
   End
   Begin VB.ComboBox cboFonts 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy 
      Cancel          =   -1  'True
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   1092
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3732
      Left            =   6360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   252
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1092
   End
   Begin VB.PictureBox picHolder 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3675
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      Begin VB.Label lblBigDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   720
         Left            =   840
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblsymbols 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Label lblLabel 
      Caption         =   "Insert string:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Caption         =   "All Symbols contained in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   630
      Width           =   1815
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurrentLabel As Integer
Private noperline As Integer
Private linesout As Integer
Private gignore As Boolean
Private minuschars As Integer
Private fntFont As String
Private blnLoadedFonts As Boolean
Private Const BorderWidth As Integer = 100

Private Sub cboFonts_Click()
    lblBigDisplay.Visible = False

    Dim i As Integer ' Declare variable.
    If lblsymbols(0).FontName <> cboFonts.text Then
        For i = 0 To lblsymbols.Count - 1
            lblsymbols(i).FontName = cboFonts.text
        Next
    End If
    If lblBigDisplay.FontName <> cboFonts.text Then
        lblBigDisplay.FontName = cboFonts.text
    End If
    ' updating form caption
    Me.Caption = "Insert symbol from " & lblBigDisplay.FontName
    ' setting fontname in txtInsert
    txtInsert.font = lblBigDisplay.FontName
    lblBigDisplay.Visible = False
End Sub

Private Sub cboFonts_DropDown()
    ' populate combobox with printer fonts
    Dim i As Integer ' Declare variable.
    If Not (blnLoadedFonts) Then
        MousePointer = vbArrowHourglass
        cboFonts.Clear
        For i = 0 To Printer.FontCount - 1 ' Determine number of fonts.
            cboFonts.AddItem Screen.Fonts(i)  ' Put each font into combo box.
        Next i
        MousePointer = vbDefault
        blnLoadedFonts = True
        On Error Resume Next
        cboFonts.text = fntFont
    End If

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.SetText txtInsert.text
    picHolder.SetFocus
End Sub

Private Sub cmdSymbols_Click(index As Integer)
'    On Error Resume Next
'    'frmMDI.ActiveForm.Text1.InsertContents SF_TEXT, cmdSymbols(Index).Caption
'
'    '...paste the Selected item.
'    frmMDI.ActiveForm.ActiveControl.SelText = ""    'This step is crucial!!! for undoing actions
'    ' Place the text from the Clipboard into the active control.
'    frmMDI.ActiveForm.ActiveControl.SelText = cmdSymbols(index).Caption
'    ' Set focus back to the active window
'    'frmMDI.ActiveForm.ActiveControl.SetFocus

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    picHolder.SetFocus
    
    '...paste the Selected item.
    frmMDI.ActiveForm.ActiveControl.SelText = ""    'This step is crucial!!! for undoing actions
    ' Place the text from the Clipboard into the active control.
    frmMDI.ActiveForm.ActiveControl.SelText = txtInsert.text
    ' Set focus back to the active window
    'frmMDI.ActiveForm.ActiveControl.SetFocus
    ' closing the big display
    lblBigDisplay.Visible = False
End Sub

Private Sub cmdInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyUp
        If Shift = 0 Then
            picHolder_KeyDown KeyCode, Shift
        End If
        KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Call rotateText("E-Type Insert Symbol", picBorder)
    
    blnLoadedFonts = False


    ' Set the current font
    fntFont = frmMDI.ActiveForm.Text1.font

    ' sett form caption
    Me.Caption = "Insert symbol from " & fntFont
    lblMessage = "Symbols contained in: "
    ' set the big display to the same font
    lblBigDisplay.font = fntFont
    noperline = 0
    ' set font and size
    lblsymbols(0).font = fntFont
    FillSymbols (0)
    gignore = True
    VScroll1.Max = linesout
    VScroll1.Min = 0
    gignore = False
    ' Set the currently selected label to 0
    CurrentLabel = 0
    
    ' adding one item named the active font name, just to show it
    ' then selecting it. The hole list vil be rebildt the first time
    ' the user click dhe dropdovn button
    cboFonts.AddItem (fntFont), 0
    cboFonts.ListIndex = 0
    
End Sub
Sub FillSymbols(ByVal startnumber As Integer)
    gignore = False
    ' use minus chars to take away left co-or
    minuschars = 1
    ' number of lines
    numberoflines = 1
    ' hide the first symbol
    lblsymbols(0).Left = -5000
    ' number of lines off screen
    linesout = 0
    ' number of symbols per line
    'noperline = 0
    ' Hide the picture box
    picHolder.Visible = False
    For i = 1 To 223
        ' Load the new symbol label
        'On Error Resume Next
        Load lblsymbols(i)
        On Error GoTo 0
        ' change the current char - miss out
        ' the first 32 chars
        currentchar = i + startnumber + 32
        If currentchar > 255 Then Exit For
        ' Set caption to char
        lblsymbols(i).Caption = Chr(currentchar)
        ' New left position
        ' (i - 1) [to allow left to start at 0
        ' - minuschars [to take away the previous
        ' symbols from prev. lines
        ' * (lblsymbols(i).Width - 12)
        ' [To move number from left plus
        ' line width
        NewLeftPos = BorderWidth + ((i) - minuschars) * (lblsymbols(i).Width - 20)
        ' If the new left pos is bigger than
        ' the container width - new symbol
        ' then start a new line
        If NewLeftPos > picHolder.Width - lblsymbols(i).Width Then
            ' Add the number of current symbols
            ' minus the one just created
            minuschars = lblsymbols.Count - 1
            ' Set the number per line (excluding
            ' current symbol, if it is not set
            ' -1 for currentsymbol
            ' -1 for first label which is not shown
            If noperline = 0 Then noperline = lblsymbols.Count - 2
            ' increment the number of lines
            numberoflines = numberoflines + 1
            ' new top position (new line)
            ' lines - 1 [allow for top =0
            ' (lblsymbols(i).Height - 12)
            ' [number of lines - thick line
            newtop = (numberoflines) * (lblsymbols(i).Height - 20)
            ' If the new top pos is greater than
            ' picHolder bottom line then increment
            ' lines out of screen
            If newtop + lblsymbols(i).Height > picHolder.Height Then
                linesout = linesout + 1
            End If
            ' Set the new left to include the new
            ' minuschar value
            'NewLeftPos = ((i) - minuschars) * (lblsymbols(i).Width - 12)
            NewLeftPos = BorderWidth + (i - minuschars) * (lblsymbols(i).Width - 20)
        End If
        ' Refresh pic1
        'picHolder.Refresh
        ' set top pos of symbol
        lblsymbols(i).Top = (numberoflines - 0.7) * (lblsymbols(i).Height - 20)
        ' set new left
        lblsymbols(i).Left = NewLeftPos
        ' make is visible
        lblsymbols(i).Visible = True
    Next
    ' Show the picture again
    picHolder.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblBigDisplay_Click()
'    lblBigDisplay.Visible = False

End Sub

Private Sub lblBigDisplay_DblClick()
    txtInsert.text = txtInsert.text & lblsymbols(CurrentLabel).Caption
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblsymbols_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo errHandler
    lblBigDisplay.Left = lblsymbols(index).Left - ((lblBigDisplay.Width - lblsymbols(index).Width) / 2)
    lblBigDisplay.Top = lblsymbols(index).Top - ((lblBigDisplay.Height - lblsymbols(index).Height) / 2)
    lblBigDisplay.Caption = lblsymbols(index).Caption
    lblBigDisplay.Visible = True
    CurrentLabel = index
    fred = lblsymbols(index).Caption
    lblStatus.Caption = "Special Char " & Asc(fred)
errHandler:

End Sub

Private Sub picHolder_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Not Shift = 0 Then Exit Sub
'    If KeyCode = vbKeyLeft And Not CurrentLabel = 1 Then
'        lblsymbols_Click (CurrentLabel - 1)
'    ElseIf KeyCode = vbKeyRight And Not CurrentLabel = lblsymbols.Count - 2 Then
'        lblsymbols_Click (CurrentLabel + 1)
'    ElseIf KeyCode = vbKeyUp And CurrentLabel > noperline Then
'        lblsymbols_Click (CurrentLabel - noperline)
'    ElseIf KeyCode = vbKeyDown And CurrentLabel < (lblsymbols.Count - 2 + noperline) Then
'        lblsymbols_Click (CurrentLabel + noperline)
'    End If
End Sub

Private Sub txtInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub VScroll1_Change()
    If Not gignore Then
        MousePointer = vbHourglass
        For Each Label In lblsymbols
            If Not Label.index = 0 Then
                Unload Label
            End If
        Next
        charstart = VScroll1.Value * noperline
        FillSymbols (charstart)
        MousePointer = vbDefault
    End If
    lblBigDisplay.Visible = False
    picHolder.SetFocus
End Sub
