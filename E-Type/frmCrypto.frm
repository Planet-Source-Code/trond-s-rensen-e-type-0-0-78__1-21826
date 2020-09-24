VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCrypto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2865
   ClientLeft      =   4560
   ClientTop       =   4530
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrypto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   5745
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
      Height          =   2865
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   330
      TabIndex        =   13
      Top             =   0
      Width           =   330
   End
   Begin VB.TextBox txtLog 
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
      Height          =   1335
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmCrypto.frx":0E42
      Top             =   3120
      Width           =   5175
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4440
      TabIndex        =   8
      Top             =   2400
      Width           =   1200
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1800
      TabIndex        =   7
      Top             =   2400
      Width           =   1200
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1200
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   5115
      TabIndex        =   4
      Top             =   600
      Width           =   5175
      Begin MSComctlLib.ProgressBar pgbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   45
         TabIndex        =   5
         Top             =   50
         Width           =   450
      End
   End
   Begin VB.TextBox txtPassword2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txtPassword1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "File:"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
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
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblPassword2 
      Caption         =   "Retype password:"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblPassword1 
      Caption         =   "Enter password:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEncrypt_Click()
    ' make sure at least 1 character is in passwordsfields
    If Len(txtPassword1) < 1 Then
        MsgBox "Password must be at least one character long!", vbExclamation, "Crypto"
        Exit Sub
    End If
    'Make sure both passwords match exactly
    If txtPassword1.text <> txtPassword2.text Then
        MsgBox "The two passwords do not match!", vbExclamation, "Crypto"
        txtPassword1.SetFocus
        Exit Sub
    End If
    'Encrypt file
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    Refresh
    Encrypt
'    txtFile_Change
    MousePointer = vbDefault
End Sub

Private Sub cmdDecrypt_Click()
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    Refresh
    Decrypt
'    txtFile_Change
    MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Call rotateText("E-Type Crypto", picBorder)
    Dim sHead As String
    Dim sT As String
    
    ' Test to see if there is any text in the document
    If frmMDI.ActiveForm.Text1.text = "" Then
        MsgBox "The file is empty! Nothing to encrypt or decrypt.", vbInformation
        Unload Me
        Exit Sub
    End If
    
    ' save activeform file as e-type.tmp in temp directory
    frmMDI.ActiveForm.Text1.SaveFile GetTmpPath & "\r-type.tmp", rtfText

    'Disable most command buttons
    'Initialize filename field
    lblFileName = StripPath(frmMDI.ActiveForm.Caption)
    lblFileName.ToolTipText = frmMDI.ActiveForm.Caption
    
    ' open active file and check for crypto header.
    ' if header is there enable decrypt button.
    ' else enable encrypt buton
    
    Open GetTmpPath & "\r-type.tmp" For Binary As #1
    
    Line Input #1, sHead
    Close #1
    'Check for header
    sT = Mid(sHead, 1, 8)
    Show
    If sT = "[Secret]" Then
        lblStatus.Caption = "About to decrypt file. Enter password"
        cmdEncrypt.Enabled = False
        cmdDecrypt.Enabled = True
        lblPassword2.Enabled = False
        txtPassword2.Enabled = False
        txtPassword1.SetFocus
    Else
        lblStatus.Caption = "About to encrypt file. Enter password."
        cmdEncrypt.Enabled = True
        cmdDecrypt.Enabled = False
        lblPassword2.Enabled = True
        txtPassword2.Enabled = True
        txtPassword1.SetFocus
    End If
End Sub

Sub Encrypt()
    Dim sHead As String
    Dim sT As String
    Dim sA As String
    Dim cphX As New Cipher
    Dim n As Long
    
    ' save activeform file as e-type.tmp in temp directory
    frmMDI.ActiveForm.Text1.SaveFile GetTmpPath & "\r-type.tmp", rtfText

    Open GetTmpPath & "\r-type.tmp" For Binary As #1
    'Load entire file into sA
    sA = Space$(LOF(1))
    Get #1, , sA
    Close #1
    'Prepare header string with salt characters
    sT = Hash(Date & Str(Timer))
    sHead = "[Secret]" & sT & Hash(sT & txtPassword1.text)
    'Do the encryption
    cphX.KeyString = sHead
    cphX.text = sA
    cphX.DoXor
    cphX.Stretch
    sA = cphX.text
    'Write header
    Open GetTmpPath & "\r-type.tmp" For Output As #1
    Print #1, sHead
    'Write encrypted data
    n = 1
    Do
        Print #1, Mid(sA, n, 70)
        n = n + 70
    Loop Until n > Len(sA)
    Close #1
    
    pgbProgress.Visible = False
    lblStatus.Caption = "File encrypted"
    txtLog.text = txtLog.text & vbCrLf & "File encrypted"
    frmMDI.ActiveForm.Text1.LoadFile GetTmpPath & "\r-type.tmp"
    Kill GetTmpPath & "\r-type.tmp"
    
End Sub

Sub DecryptOld()
    Dim sHead As String
    Dim sA As String
    Dim sT As String
    Dim cphX As New Cipher
    Dim n As Long
    
    ' save activeform file as e-type.tmp in temp directory as textfile
    frmMDI.ActiveForm.Text1.SaveFile GetTmpPath & "r-type.tmp", rtfText

    'Get header (first 18 bytes of encrypted file)
    Open GetTmpPath & "r-type.tmp" For Binary As #1
    Line Input #1, sHead
    Close #1
    'Check for correct password
    sT = Mid(sHead, 9, 8)
    If InStr(sHead, Hash(sT & txtPassword1.text)) <> 17 Then
        MsgBox "Sorry, this is not the correct password!", _
            vbExclamation, "Secret"
            txtPassword1.SetFocus
        Exit Sub
    End If
    'Get file contents
    Open GetTmpPath & "r-type.tmp" For Binary As #1
    'Read past the header
    Line Input #1, sHead
    'Read and build the contents string
    Do Until EOF(1)
        Line Input #1, sT
        sA = sA & sT
    Loop
    Close #1
    'Decrypted file contents
    cphX.KeyString = sHead
    cphX.text = sA
    cphX.Shrink
    cphX.DoXor
    sA = cphX.text
    
    'Replace file with decrypted version
    Kill GetTmpPath & "\r-type.tmp"
    Open GetTmpPath & "\r-type.tmp" For Binary As #1
    Put #1, , sA
    Close #1
    pgbProgress.Visible = False
    lblStatus.Caption = "File decrypted"
    frmMDI.ActiveForm.Text1.LoadFile GetTmpPath & "\r-type.tmp"
    Kill GetTmpPath & "\r-type.tmp"

End Sub

Function Hash(sA As String) As String
    Dim cphHash As New Cipher
    cphHash.KeyString = sA & "123456"
    cphHash.text = sA & "123456"
    cphHash.DoXor
    cphHash.Stretch
    cphHash.KeyString = cphHash.text
    cphHash.text = "123456"
    cphHash.DoXor
    cphHash.Stretch
    Hash = cphHash.text
End Function

Sub Decrypt()
    Dim sHead As String
    Dim sA As String
    Dim sT As String
    Dim cphX As New Cipher
    Dim n As Long
    Dim tmpFile As String
    
    tmpFile = GetTmpPath & "r-type.tmp"
    
    On Error Resume Next
    Dim strContents As String

    ' save a copy of the active file in windows temp directory.
    Open tmpFile For Output As #1
    ' Place the contents of the notepad into a variable.
    strContents = frmMDI.ActiveForm.Text1.text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #1, strContents
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0

    
    'Get header (first 18 bytes of encrypted file)
    Open tmpFile For Input As #1
    Line Input #1, sHead
    Close #1
    'Check for correct password
    sT = Mid(sHead, 9, 8)
    If InStr(sHead, Hash(sT & txtPassword1.text)) <> 17 Then
        MsgBox "Sorry, this is not the correct password!", _
            vbExclamation, "Secret"
        Exit Sub
    End If
    'Get file contents
    Open tmpFile For Input As #1
    'Read past the header
    Line Input #1, sHead
    'Read and build the contents string
    Do Until EOF(1)
        Line Input #1, sT
        sA = sA & sT
    Loop
    Close #1
    'Decrypted file contents
    cphX.KeyString = sHead
    cphX.text = sA
    cphX.Shrink
    cphX.DoXor
    sA = cphX.text
    ' delete the old tmpFile
    Kill tmpFile
    'Replace file with decrypted version
    Open tmpFile For Binary As #1
    Put #1, , sA
    Close #1
    pgbProgress.Visible = False
    lblStatus.Caption = "File decrypted"
    txtLog.text = txtLog.text & vbCrLf & "File encrypted"
    frmMDI.ActiveForm.Text1.LoadFile tmpFile
    Kill tmpFile

End Sub

