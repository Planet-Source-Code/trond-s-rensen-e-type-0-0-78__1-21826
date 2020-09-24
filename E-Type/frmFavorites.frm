VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   5100
   ClientTop       =   3060
   ClientWidth     =   6150
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6150
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
      Height          =   3300
      Left            =   0
      ScaleHeight     =   3300
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin MSComctlLib.ListView lwFiles 
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlDriveFileList2"
      SmallIcons      =   "imlDriveFileList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAddOpenFiles 
      Caption         =   "Add &Open Files"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "&Add File"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imlDriveFileList2 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":0E42
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":1166
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":14BA
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFavorites.frx":180E
            Key             =   "TextFile"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddFile_Click()
    Call AddDocToGroup(frmMDI.TVProjects, lwFiles.SelectedItem.SubItems(1))
End Sub

Private Sub cmdBrowse_Click()
    Dim strOpenFileName As String
    Dim strDefaultName As String
    ' Assign a default name to the variable.
    strDefaultName = "NewProject.etp"
    
    ' Display a Save As dialog box and return a filename.
    On Error Resume Next
    frmMDI.CMDialog1.DialogTitle = "Add file"
    frmMDI.CMDialog1.Filter = "Text Files (*.txt)"
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.DefaultExt = "txt"
    frmMDI.CMDialog1.Flags = &H4    'fcdlOFNHideReadOnly
    frmMDI.CMDialog1.ShowOpen
       
    strOpenFileName = frmMDI.CMDialog1.FileName
    ' add selected doc to project
    Call AddDocToGroup(frmMDI.TVProjects, strOpenFileName)

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstFiles_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdOpen_Click()

End Sub

Private Sub Form_Load()
    Dim i As Integer        ' Counter variable
    Dim itmX As ListItem
    
    
    Call RotateText("Add Files(s) to project", picBorder)
    ' Create an object variable for the ColumnHeader object.
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    Dim clmX As ColumnHeader
    
    lwFiles.ColumnHeaders.Add , , "File Name", 2000
    Set clmX = lwFiles.ColumnHeaders.Add(, , "File Path", 8000)
    
    With lwFiles
        ' Cycle through the document array.
        ' Add all open documents to list.
        For i = 1 To UBound(Document)
            Set itmX = .ListItems.Add(, , StripPath(Document(i).Caption), 4)    ' only filename
            If Left((Document(i).Caption), 9) = "Untitled:" Then
                itmX.SubItems(1) = "File not saved"
            Else
                itmX.SubItems(1) = (Document(i).Caption)    ' filename and full path
            End If
            itmX.Icon = 4           ' Set an icon from imlDriveFileList.
            itmX.SmallIcon = 4      ' Set an icon from ImageList2.
        Next i
    End With
End Sub

Private Sub lwFiles_DblClick()
    ' add selected doc to project
    Call AddDocToGroup(frmMDI.TVProjects, lwFiles.SelectedItem.SubItems(1))
End Sub
