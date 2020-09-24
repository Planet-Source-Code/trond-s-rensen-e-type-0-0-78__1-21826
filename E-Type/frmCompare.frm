VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompare 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1920
   ClientLeft      =   4545
   ClientTop       =   5085
   ClientWidth     =   5475
   Icon            =   "frmCompare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
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
      Height          =   1920
      Left            =   0
      ScaleHeight     =   1920
      ScaleWidth      =   330
      TabIndex        =   10
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   375
      Left            =   450
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.CheckBox chkToFile 
      Caption         =   "Output to document"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdFile2 
      Caption         =   "..."
      Height          =   300
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdFile1 
      Caption         =   "..."
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4230
      TabIndex        =   6
      Tag             =   "Cancel"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000000FF&
      Caption         =   "Compare"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Tag             =   "OK"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtFile1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFile2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   5280
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "TXT"
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.Label lblFirstFile 
      Caption         =   "First File:"
      Height          =   255
      Left            =   450
      TabIndex        =   7
      Top             =   375
      Width           =   855
   End
   Begin VB.Label lblSecondFile 
      Caption         =   "Second File:"
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   855
      Width           =   975
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFile1_Click()
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    cdlOpen.DialogTitle = "First File"
    cdlOpen.FileName = ""
    cdlOpen.ShowOpen
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = cdlOpen.FileName
        txtFile1.text = strOpenFileName
    End If

End Sub

Private Sub cmdFile2_Click()
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    cdlOpen.DialogTitle = "Second File"
    cdlOpen.FileName = ""
    cdlOpen.ShowOpen
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = cdlOpen.FileName
        txtFile2.text = strOpenFileName
    End If

End Sub

Private Sub cmdHelp_Click()
    MsgBox "Purpose: Allows a efficent means to manage duplicate records in text files." _
    & Chr(13) & "Logic: Anything found in  First File NOT found in Second File is written to OutputFile." _
    & Chr(13) & "OutputFile will contain only new lines that do not exist in Second File.", vbInformation, "Output To Ducument Help"
End Sub

Private Sub cmdOK_Click()
    If chkToFile.Value = 0 Then
        Dim File1 As String
        Dim File2 As String
        Dim IsSame As Integer
        Dim buffer1 As String
        Dim buffer2 As String
        Dim x As Long
        Dim whole As Long
        Dim part As Long
        Dim Start As Long
        Dim test As Boolean
        
        File1 = txtFile1.text
        File2 = txtFile2.text
        ' Test if there is a filename in both textboxes
        If File1 = "" Then
            MsgBox "Please choose First File!", vbExclamation
            txtFile1.SetFocus
            Exit Sub
        End If
        If File2 = "" Then
            MsgBox "Please choose Second File!", vbExclamation
            txtFile2.SetFocus
            Exit Sub
        End If
        ' test if the first filename are valid
        test = FileExists(File1)
        If test = False Then
            MsgBox "First file not found!", vbCritical
            Close
            txtFile1.SetFocus
            txtFile1.SelStart = 0
            txtFile1.SelLength = Len(txtFile1.text)
            Exit Sub
        Else
            Open File1 For Binary As #1
        End If
        ' test if the second filename are valid
        test = FileExists(File2)
        If test = False Then
            MsgBox "Second file not found!", vbCritical
            Close
            txtFile2.SetFocus
            txtFile2.SelStart = 0
            txtFile2.SelLength = Len(txtFile2.text)
            Exit Sub
        Else
            Open File2 For Binary As #2
        End If

        IsSame = True
        If LOF(1) <> LOF(2) Then
            IsSame = False
        Else
            whole = LOF(1) \ 10000          'number of whole 10,000 byte chunks
            part = LOF(1) Mod 10000         'remaining bytes at end of file
            buffer1 = String$(10000, 0)
            buffer2 = String$(10000, 0)
            Start = 1
            For x = 1 To whole              'this for-next loop will get 10,000
                Get #1, Start, buffer1      'byte chunks at a time.
                Get #2, Start, buffer2
                If buffer1 <> buffer2 Then
                    IsSame = False
                        Exit For
                End If
                Start = Start + 10000
            Next
            buffer1 = String$(part, 0)
            buffer2 = String$(part, 0)
            Get #1, Start, buffer1        'get the remaining bytes at the end
            Get #2, Start, buffer2        'get the remaining bytes at the end
            If buffer1 <> buffer2 Then IsSame = False
            End If
            Close #1
            Close #2
            If IsSame Then
                MsgBox "Files are identical", vbInformation, "Info"
            Else
                     MsgBox "Files are NOT identical", vbInformation, "Info"
        End If
    Else
        MousePointer = vbHourglass
        CompareFiles txtFile2.text, txtFile1.text, GetTmpPath & "CompareResults.txt"

        ' Call the file open procedure, passing a
        ' reference to the selected file name
        OpenFile GetTmpPath & "CompareResults.txt"
'        ' Update the list of recently opened files in the File menu control array.
'        UpdateFileMenu "CompareResults.txt"
        ' If statusbar is visible, update infomation on form
        If frmMDI.SBarMain.Visible = True Then
            GetFileStats    ' filedata, date,size ect.
        End If

        Kill GetTmpPath & "CompareResults.txt"
        MousePointer = vbDefault 'vbHourglass
        frmMDI.SBarMain.Panels(1).text = ""
    End If
End Sub

Private Sub Form_Load()
    Call rotateText("Compare", picBorder)
End Sub
